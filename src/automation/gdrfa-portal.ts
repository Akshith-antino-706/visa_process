import { WebDriver, By, until, WebElement, Key } from 'selenium-webdriver';
import * as path from 'path';
import * as fs from 'fs';
import sharp from 'sharp';
import {
  VisaApplication,
  PassportDetails,
  ApplicantDetails,
  ContactDetails,
  ApplicationDocuments,
} from '../types/application-data';
import { matchDocumentsToSlots, matchDocumentsToSlotsLocal } from '../utils/doc-matcher';
// S3 helper available if needed in the future:
// import { isS3Configured, getPublicUrl } from '../utils/s3-helper';

// ─── File compression ──────────────────────────────────────────────────────
// GDRFA portal: Allowed .jpg/.pdf/.png, max 1000 KB per file.
// Compress images that exceed this. PDFs are also checked.
const MAX_UPLOAD_BYTES = 1000 * 1024; // 1000 KB — GDRFA portal limit

async function compressImageIfNeeded(
  filePath: string,
): Promise<{ buffer: Buffer; fileName: string; mimeType: string }> {
  const resolved = path.resolve(filePath);
  const originalBuffer = fs.readFileSync(resolved);
  const ext = path.extname(resolved).toLowerCase();
  const originalName = path.basename(resolved);

  const isImage = ['.jpg', '.jpeg', '.png', '.webp'].includes(ext);
  const isPdf = ext === '.pdf';

  // PDFs: pass through if under limit, otherwise convert first page to JPEG
  if (isPdf) {
    if (originalBuffer.length <= MAX_UPLOAD_BYTES) {
      return { buffer: originalBuffer, fileName: originalName, mimeType: 'application/pdf' };
    }
    console.warn(
      `[Compress] ${originalName} is ${(originalBuffer.length / 1024).toFixed(0)} KB (PDF) — exceeds ${(MAX_UPLOAD_BYTES / 1024).toFixed(0)} KB limit. Converting first page to JPEG...`,
    );
    try {
      // sharp can render the first page of a PDF to an image
      let pdfImage = await sharp(originalBuffer, { density: 150 })
        .jpeg({ quality: 60, mozjpeg: true })
        .toBuffer();
      if (pdfImage.length > MAX_UPLOAD_BYTES) {
        pdfImage = await sharp(originalBuffer, { density: 100 })
          .resize({ width: 800, height: 800, fit: 'inside', withoutEnlargement: true })
          .jpeg({ quality: 40, mozjpeg: true })
          .toBuffer();
      }
      const jpegName = originalName.replace(/\.pdf$/i, '.jpg');
      console.log(
        `[Compress] PDF→JPEG: ${(originalBuffer.length / 1024).toFixed(0)} KB → ${(pdfImage.length / 1024).toFixed(0)} KB`,
      );
      return { buffer: pdfImage, fileName: jpegName, mimeType: 'image/jpeg' };
    } catch (pdfErr) {
      console.warn(`[Compress] PDF→JPEG conversion failed: ${pdfErr}. Passing original PDF.`);
      return { buffer: originalBuffer, fileName: originalName, mimeType: 'application/pdf' };
    }
  }

  if (!isImage) {
    return { buffer: originalBuffer, fileName: originalName, mimeType: 'application/octet-stream' };
  }

  // Image is under limit — pass through
  if (originalBuffer.length <= MAX_UPLOAD_BYTES) {
    const mime = ext === '.png' ? 'image/png' : 'image/jpeg';
    return { buffer: originalBuffer, fileName: originalName, mimeType: mime };
  }

  console.log(
    `[Compress] ${originalName}: ${(originalBuffer.length / 1024).toFixed(0)} KB exceeds ${(MAX_UPLOAD_BYTES / 1024).toFixed(0)} KB — compressing...`,
  );

  // Progressive compression: try decreasing quality until under the limit
  let quality = 80;
  let compressed: Buffer = originalBuffer;
  const metadata = await sharp(originalBuffer).metadata();

  // Cap dimensions — large photos are unnecessary for a visa portal
  const maxDim = 1200;
  const needsResize =
    (metadata.width && metadata.width > maxDim) ||
    (metadata.height && metadata.height > maxDim);

  while (quality >= 20) {
    let pipeline = sharp(originalBuffer);

    if (needsResize) {
      pipeline = pipeline.resize({
        width: maxDim,
        height: maxDim,
        fit: 'inside',
        withoutEnlargement: true,
      });
    }

    // Always output as JPEG for best compression
    compressed = await pipeline
      .jpeg({ quality, mozjpeg: true })
      .toBuffer();

    if (compressed.length <= MAX_UPLOAD_BYTES) break;
    quality -= 10;
  }

  // If still too large after quality 20, resize more aggressively in steps
  if (compressed.length > MAX_UPLOAD_BYTES) {
    for (const dim of [800, 600, 400]) {
      compressed = await sharp(originalBuffer)
        .resize({ width: dim, height: dim, fit: 'inside', withoutEnlargement: true })
        .jpeg({ quality: 20, mozjpeg: true })
        .toBuffer();
      if (compressed.length <= MAX_UPLOAD_BYTES) break;
    }
  }

  if (compressed.length > MAX_UPLOAD_BYTES) {
    console.warn(
      `[Compress] WARNING: ${originalName} still ${(compressed.length / 1024).toFixed(0)} KB after max compression — portal may reject it.`,
    );
  }

  const compressedName = originalName.replace(/\.\w+$/, '.jpg');
  console.log(
    `[Compress] ${originalName}: ${(originalBuffer.length / 1024).toFixed(0)} KB → ${(compressed.length / 1024).toFixed(0)} KB (quality=${quality}, JPEG)`,
  );

  return { buffer: compressed, fileName: compressedName, mimeType: 'image/jpeg' };
}

// ─── Page Object ──────────────────────────────────────────────────────────────

export class GdrfaPortalPage {
  private static readonly HOME = 'https://smart.gdrfad.gov.ae/SmartChannels_Th/';
  private static readonly ACTION_TIMEOUT = 20_000;

  constructor(private readonly driver: WebDriver) {}

  // ── Public API ─────────────────────────────────────────────────────────────

  async verifySession(): Promise<void> {
    console.log('[Session] Verifying session...');
    await this.driver.get(GdrfaPortalPage.HOME);
    await this.waitForPageLoad();
    const url = await this.driver.getCurrentUrl();
    if (url.includes('Login.aspx')) {
      throw new Error('[Session] Session expired — run "npm run auth" to log in again.');
    }
    console.log('[Session] Valid. URL:', url);
  }

  async fillApplicationForm(application: VisaApplication): Promise<string> {
    console.log('\n[Flow] ─── Starting navigation ───');
    const stopKeepAlive = this.startSessionKeepAlive();
    try {
      await this.verifySession();
      await this.navigateToNewApplication();
      await this.waitForPageSettle();

      // ── Section 1: Visit Details ──
      console.log('\n[Flow] ── Section: Visit Details ──');
      await this.setVisitReason();
      await this.sleep(500);

      // ── Section 2: Passport Details ──
      console.log('\n[Flow] ── Section: Passport Details ──');
      await this.setPassportType(application.passport.passportType);
      await this.sleep(500);
      await this.enterPassportNumber(application.passport.passportNumber);
      await this.sleep(500);
      await this.setNationality(application.passport.currentNationality);
      await this.sleep(500);
      await this.setPreviousNationality(
        application.passport.previousNationality ?? application.passport.currentNationality
      );
      await this.sleep(500);
      await this.clickSearchDataAndWait();
      await this.sleep(1000);

      // ── Section 3: Passport Names ──
      console.log('\n[Flow] ── Section: Passport Names ──');
      await this.fillPassportNames(application.passport);
      await this.sleep(500);

      // ── Section 4: Passport Dates & Details (dropdowns first, then text) ──
      console.log('\n[Flow] ── Section: Passport Dates ──');
      await this.fillPassportDetails(application.passport);
      await this.sleep(500);

      // ── Section 5: Applicant Details ──
      console.log('\n[Flow] ── Section: Applicant Details ──');
      await this.fillApplicantDetails(application.applicant, application.passport.passportIssueCountry);
      await this.sleep(500);

      // ── Section 6: Contact Details ──
      console.log('\n[Flow] ── Section: Contact Details ──');
      await this.fillContactDetails(application.contact);
      await this.sleep(500);

      // Retry Faith selection before continuing (dropdown can reset after other fields)
      await this.retryFaithSelection(application.applicant.faith || 'Unknown');
      await this.sleep(500);

      // Validate all required fields before clicking Continue — retry any empty ones
      await this.validateAndRetryRequiredFields(application);
      await this.sleep(500);

      await this.clickContinue();

      // Upload documents on the Attachments tab
      console.log('\n[Flow] ── Section: Document Upload ──');
      await this.uploadDocuments(application.documents);

      // Verify the application reached "READY TO PAY" status
      const statusText = await this.verifyReadyToPayStatus();
      if (!statusText) {
        throw new Error('[Verify] Application did not reach "READY TO PAY" status after upload.');
      }
      console.log(`[Flow] Status confirmed: "${statusText}"`);

      // Capture the Application Number from the page
      const appNumber = await this.captureApplicationNumber();
      console.log(`\n[Flow] ─── Application Number: ${appNumber || 'NOT FOUND'} ───\n`);

      return appNumber;
    } finally {
      stopKeepAlive();
    }
  }

  /**
   * Captures the Application Number from the page after form submission.
   * Checks multiple locations: header area, confirmation page, URL params, and page text.
   */
  private async captureApplicationNumber(): Promise<string> {
    console.log('[Capture] Looking for Application Number...');
    await this.sleep(3000);

    const rawAppNo = await this.driver.executeScript<string>(`
      // 1. Check the header area (Application No. field at top of page)
      var headerSpans = document.querySelectorAll('span, td, div');
      for (var i = 0; i < headerSpans.length; i++) {
        var text = (headerSpans[i].textContent || '').trim();
        // Look for "Application No." label followed by a value
        if (text.match(/Application\\s*No\\.?\\s*$/i)) {
          var next = headerSpans[i].nextElementSibling;
          if (next) {
            var val = next.textContent.trim();
            if (val && val !== 'None' && val !== '') return val;
          }
        }
        // Look for combined "Application No. XXXX" text
        var match = text.match(/Application\\s*(?:No\\.?|Number)\\s*[:\\s]*([A-Z0-9\\/-]+)/i);
        if (match && match[1] && match[1] !== 'None') return match[1];
      }

      // 2. Check for Application Number in any bold/highlighted element
      var bolds = document.querySelectorAll('b, strong, .Bold, span[class*="app"], span[class*="App"]');
      for (var i = 0; i < bolds.length; i++) {
        var t = (bolds[i].textContent || '').trim();
        if (/^[A-Z0-9]{2,}[\\/\\-][A-Z0-9]+/i.test(t)) return t;
      }

      // 3. Check URL for application ID
      var urlMatch = window.location.href.match(/[Aa]pplication[Ii]d=([^&]+)/);
      if (urlMatch) return urlMatch[1];

      // 4. Check for any element with "application" in its ID that has a value
      var appEls = document.querySelectorAll('[id*="Application"], [id*="application"]');
      for (var i = 0; i < appEls.length; i++) {
        var val = (appEls[i].textContent || appEls[i].value || '').trim();
        if (val && val !== 'None' && val.length > 3 && val.length < 50) return val;
      }

      // 5. Look for any long numeric string on the page (GDRFA app numbers are 14+ digits)
      var allText = document.body.innerText || '';
      var numMatch = allText.match(/(\\d{10,20})/);
      if (numMatch) return numMatch[1];

      return '';
    `);

    // Clean up the application number — extract just digits
    let appNo = rawAppNo || '';
    if (appNo) {
      // Extract the numeric portion (GDRFA app numbers are purely numeric, 14+ digits)
      const numericMatch = appNo.match(/(\d{10,20})/);
      if (numericMatch) {
        appNo = numericMatch[1];
      } else {
        // Remove common trailing garbage like "DraftUpload", "Draft", etc.
        appNo = appNo.replace(/(Draft|Upload|DraftUpload|Page|View|Edit|New|Status).*$/gi, '').trim();
      }
      console.log(`[Capture] Application Number: ${appNo}`);
    } else {
      console.warn('[Capture] Application Number not found on page.');
      // Save debug screenshot
      try {
        const screenshot = await this.driver.takeScreenshot();
        const ssPath = path.resolve('test-results', 'app-number-debug.png');
        fs.mkdirSync(path.dirname(ssPath), { recursive: true });
        fs.writeFileSync(ssPath, screenshot, 'base64');
        console.log(`[Capture] Debug screenshot: ${ssPath}`);
      } catch {}
    }

    return appNo;
  }

  /**
   * Checks the page for a "READY TO PAY" (or similar) status badge after submission.
   * Waits up to 15s since the portal may take a moment to update the status.
   * Returns the status text if found, or empty string if not.
   */
  private async verifyReadyToPayStatus(): Promise<string> {
    console.log('[Verify] Checking for READY TO PAY status...');

    for (let attempt = 0; attempt < 5; attempt++) {
      const statusText = await this.driver.executeScript<string>(`
        var allEls = document.querySelectorAll('span, div, td, button, a, label, p');
        for (var i = 0; i < allEls.length; i++) {
          var t = (allEls[i].textContent || '').trim().toUpperCase();
          if (t === 'READY TO PAY' || t === 'READYTOPAY' || t === 'READY FOR PAYMENT') {
            return allEls[i].textContent.trim();
          }
        }
        // Also check for status-like badges/buttons with partial match
        var badges = document.querySelectorAll('[class*="badge"], [class*="status"], [class*="Status"], [class*="tag"], [class*="Tag"], [class*="btn"]');
        for (var i = 0; i < badges.length; i++) {
          var t = (badges[i].textContent || '').trim().toUpperCase();
          if (t.indexOf('READY') >= 0 && t.indexOf('PAY') >= 0) {
            return badges[i].textContent.trim();
          }
        }
        return '';
      `);

      if (statusText) return statusText;

      if (attempt < 4) {
        console.log(`[Verify] Status not found yet (${attempt + 1}/5)...`);
        await this.sleep(3000);
      }
    }

    return '';
  }

  /**
   * Uploads documents using OpenAI for intelligent file-to-slot matching,
   * then physically clicks the file input and uses sendKeys to upload each file.
   */
  async uploadDocuments(docs: ApplicationDocuments): Promise<void> {
    console.log('[Upload] Starting document upload...');

    // Ensure we're on the default content (not stuck in an iframe from popup handling)
    try { await this.driver.switchTo().defaultContent(); } catch {}
    await this.waitForPageLoad();
    await this.sleep(1500);

    // Log current URL for debugging
    const uploadUrl = await this.driver.getCurrentUrl();
    console.log(`[Upload] Current URL: ${uploadUrl}`);

    // Check if upload content is inside an iframe — if so, switch into it
    let switchedToIframe = false;
    const fileInputCount = await this.driver.executeScript<number>(
      `return document.querySelectorAll('input[type="file"]').length;`
    );
    if (fileInputCount === 0) {
      console.log('[Upload] No file inputs in main frame — checking iframes...');
      const iframes = await this.driver.findElements(By.css('iframe'));
      for (const iframe of iframes) {
        try {
          await this.driver.switchTo().frame(iframe);
          const count = await this.driver.executeScript<number>(
            `return document.querySelectorAll('input[type="file"]').length;`
          );
          if (count > 0) {
            console.log(`[Upload] Found ${count} file inputs inside iframe — staying in iframe.`);
            switchedToIframe = true;
            break;
          }
          await this.driver.switchTo().defaultContent();
        } catch {
          try { await this.driver.switchTo().defaultContent(); } catch {}
        }
      }
    } else {
      console.log(`[Upload] Found ${fileInputCount} file inputs in main frame.`);
    }

    // Wait for file inputs with data-document-type to render (up to 30s)
    let availableSlots: string[] = [];
    for (let waitAttempt = 0; waitAttempt < 10; waitAttempt++) {
      availableSlots = await this.driver.executeScript<string[]>(`
        return Array.from(document.querySelectorAll('input[type="file"][data-document-type]'))
          .map(function(el) { return el.getAttribute('data-document-type') || ''; })
          .filter(function(s) { return s.length > 0; });
      `);
      if (availableSlots.length > 0) break;
      console.log(`[Upload] Waiting for upload slots to render (${waitAttempt + 1}/10)...`);
      await this.sleep(3000);
    }
    console.log(`[Upload] Available slots (${availableSlots.length}):`, availableSlots);

    // Get the documents folder — prefer the explicit documentsFolder property,
    // fall back to deriving it from the first valid file path.
    let docsFolder = docs.documentsFolder ? path.resolve(docs.documentsFolder) : '';
    if (!docsFolder) {
      const allDocPaths: string[] = [
        docs.sponsoredPassportPage1,
        docs.passportExternalCoverPage,
        docs.personalPhoto,
        docs.hotelReservationPage1,
        docs.returnAirTicketPage1,
        docs.hotelReservationPage2 || '',
        docs.othersPage1 || '',
        docs.returnAirTicketPage2 || '',
        ...(docs.sponsoredPassportPages || []),
      ].filter(Boolean);
      for (const p of allDocPaths) {
        if (p) { docsFolder = path.dirname(path.resolve(p)); break; }
      }
    }
    console.log(`[Upload] Documents folder: ${docsFolder || '(empty)'}`);
    console.log(`[Upload] Folder exists: ${docsFolder ? fs.existsSync(docsFolder) : false}`);

    // List ALL files in the documents folder
    let allFileNames: string[] = [];
    if (docsFolder && fs.existsSync(docsFolder)) {
      allFileNames = fs.readdirSync(docsFolder)
        .filter(f => !f.startsWith('.') && f !== 'desktop.ini');
      console.log(`[Upload] Files in folder (${allFileNames.length}): ${allFileNames.join(', ')}`);
    } else {
      console.warn(`[Upload] Documents folder not found or empty — file matching will fail!`);
    }

    // Use OpenAI to intelligently match files to upload slots
    let docMap: Array<{ label: string; file: string }> = [];
    const applicantName = docsFolder.split(path.sep).pop() || 'Unknown';

    if (allFileNames.length > 0 && availableSlots.length > 0) {
      try {
        console.log('[Upload] Using OpenAI to match documents to slots...');
        const aiMatches = await matchDocumentsToSlots(allFileNames, availableSlots, applicantName);

        if (aiMatches.length > 0) {
          docMap = aiMatches.map(m => ({
            label: m.slotLabel,
            file: path.join(docsFolder, m.fileName),
          }));
        } else {
          throw new Error('No AI matches returned');
        }
      } catch (error) {
        console.warn(`[Upload] OpenAI matching failed: ${error}. Using local fallback.`);
        const localMatches = matchDocumentsToSlotsLocal(allFileNames, availableSlots);
        docMap = localMatches.map(m => ({
          label: m.slotLabel,
          file: path.join(docsFolder, m.fileName),
        }));
      }
    }

    // If no matches from LLM/local, fall back to the hardcoded map
    if (docMap.length === 0) {
      console.log('[Upload] Using hardcoded document map...');
      docMap = [
        { label: 'Sponsored Passport page 1',   file: docs.sponsoredPassportPage1 },
        { label: 'Passport External Cover Page', file: docs.passportExternalCoverPage },
        { label: 'Personal Photo',               file: docs.personalPhoto },
      ];
      if (docs.hotelReservationPage1) {
        docMap.push({ label: 'Hotel reservation/Place of stay - Page 1', file: docs.hotelReservationPage1 });
      }
      if (docs.returnAirTicketPage1) {
        docMap.push({ label: 'Return air ticket - Page 1', file: docs.returnAirTicketPage1 });
      }

      // Check if hotel/flight slots exist — if not, remap to Others
      const hasHotelSlot = availableSlots.some(s => s.toLowerCase().includes('hotel'));
      const hasTicketSlot = availableSlots.some(s => s.toLowerCase().includes('ticket'));

      if (!hasHotelSlot && docs.hotelReservationPage1) {
        const idx = docMap.findIndex(d => d.label.includes('Hotel'));
        if (idx >= 0) docMap[idx].label = 'Others Page 1';
        console.log('[Upload] No hotel slot → using "Others Page 1"');
      }
      if (!hasTicketSlot && docs.returnAirTicketPage1) {
        const idx = docMap.findIndex(d => d.label.includes('Return'));
        if (idx >= 0) docMap[idx].label = 'Others Page 2';
        console.log('[Upload] No ticket slot → using "Others Page 2"');
      }
    }

    // If hotel reservation file is missing, upload passport pic in its slot instead
    const hasHotelDoc = docMap.some(d =>
      d.file && fs.existsSync(path.resolve(d.file)) &&
      (d.label.toLowerCase().includes('hotel') || d.label.toLowerCase().includes('place of stay'))
    );
    if (!hasHotelDoc) {
      // Find the passport pic to reuse
      const passportFile = docs.sponsoredPassportPage1 || docs.passportExternalCoverPage || '';
      if (passportFile && fs.existsSync(path.resolve(passportFile))) {
        // Find the hotel slot label from available slots, or use "Others Page 1"
        const hotelSlot = availableSlots.find(s => s.toLowerCase().includes('hotel') || s.toLowerCase().includes('place of stay'));
        const label = hotelSlot || 'Others Page 1';
        // Remove any existing entry for this label so we don't upload twice
        docMap = docMap.filter(d => d.label !== label);
        docMap.push({ label, file: passportFile });
        console.log(`[Upload] No hotel file → uploading passport pic to "${label}"`);
      }
    }

    // Log the final document map before uploading
    console.log(`[Upload] Document map (${docMap.length} entries):`);
    const validDocs: Array<{ label: string; file: string }> = [];
    for (const doc of docMap) {
      const exists = doc.file ? fs.existsSync(path.resolve(doc.file)) : false;
      console.log(`  "${doc.label}" → ${doc.file || '(empty)'} ${exists ? '✓' : '✗ MISSING'}`);
      if (doc.file && exists) validDocs.push(doc);
      else if (!doc.file) console.warn(`[Upload] Skipping "${doc.label}" — no file path`);
      else console.warn(`[Upload] File not found: ${path.resolve(doc.file)}`);
    }

    // Pre-compress ALL files in parallel (saves ~1-2s per file vs sequential)
    console.log(`[Upload] Pre-compressing ${validDocs.length} files in parallel...`);
    const compressStart = Date.now();
    const preCompressed = await Promise.all(
      validDocs.map(async (doc) => {
        const { buffer, fileName, mimeType } = await compressImageIfNeeded(path.resolve(doc.file));
        return { label: doc.label, buffer, fileName, mimeType, base64: buffer.toString('base64') };
      })
    );
    console.log(`[Upload] Pre-compression done in ${((Date.now() - compressStart) / 1000).toFixed(1)}s`);

    // Upload each document with pre-compressed data
    let uploadCount = 0;
    for (const doc of preCompressed) {
      await this.uploadSingleDocument(doc.label, doc.base64, doc.fileName, doc.mimeType);
      uploadCount++;
    }
    console.log(`[Upload] Uploaded ${uploadCount}/${docMap.length} documents.`);

    console.log('[Upload] All documents sent. Now waiting for Continue button...');
    await this.sleep(1000);

    // Switch back to default content before looking for Continue button
    if (switchedToIframe) {
      try { await this.driver.switchTo().defaultContent(); } catch {}
    }

    // The goal: wait for Continue to appear and click it — that means uploads are done
    await this.clickUploadContinue();

    // Take screenshot after successful Continue click
    try {
      const screenshot = await this.driver.takeScreenshot();
      const ssPath = path.resolve('test-results', 'after-upload-continue.png');
      fs.mkdirSync(path.dirname(ssPath), { recursive: true });
      fs.writeFileSync(ssPath, screenshot, 'base64');
      console.log(`[Upload] Screenshot: ${ssPath}`);
    } catch {}

    console.log('[Upload] Upload section complete — Continue clicked successfully.');
  }

  /**
   * Uploads a single document using a 3-strategy cascade:
   *
   * Strategy 1 (Primary): S3 fetch → DataTransfer injection
   *   - fetch(presignedUrl) inside the browser → Blob → File → DataTransfer → input.files
   *   - Zero filesystem dependency. File goes S3 → browser memory → input.
   *   - Creates a REAL File object indistinguishable from user-selected file.
   *
   * Strategy 2 (Fallback): sendKeys + DataTransfer re-injection
   *   - sendKeys sets .files, then we read the file with FileReader,
   *     re-create it via DataTransfer to ensure proper File objects.
   *
   * Strategy 3 (Last resort): sendKeys only
   *   - Classic sendKeys on the visible file input.
   *
   * After files are set (by any strategy):
   *   → dispatch change event (jQuery + native)
   *   → call readUrldoc_N with correct this context
   *   → wait for OsNotify AJAX postback (NOT SaveButton — it causes page navigation)
   */
  private async uploadSingleDocument(docType: string, base64Data: string, fileName: string, mimeType: string): Promise<void> {
    console.log(`[Upload] "${docType}" → ${fileName}`);

    // Step 1: Find the card by its title span, then get the file input inside.
    //
    // DOM per card:
    //   div[id$="_wtcntDocList"]  ← outer card (dotted border)
    //     span.Bold → "Title"    ← title text
    //     div.CardGray.mt-card   ← upload area
    //       input[type="file"].doc_N[data-document-type="Title"]
    //       readUrldoc_N() → FileReader → OsNotify()
    //
    const inputInfo = await this.driver.executeScript<string>(`
      var target = arguments[0].toLowerCase().trim();

      // 1. Find the title span matching this docType
      var spans = document.querySelectorAll('span.Bold');
      var matched = null;
      // Exact match
      for (var i = 0; i < spans.length; i++) {
        if ((spans[i].textContent || '').trim().toLowerCase() === target) {
          matched = spans[i]; break;
        }
      }
      // Contains match
      if (!matched) {
        for (var i = 0; i < spans.length; i++) {
          var t = (spans[i].textContent || '').trim().toLowerCase();
          if (t.indexOf(target) >= 0 || target.indexOf(t) >= 0) {
            matched = spans[i]; break;
          }
        }
      }
      // Keyword match (handles "Father Passport" → "Parent Passport" etc.)
      if (!matched) {
        var keywords = target.split(/[\\s\\-\\/]+/).filter(function(k) { return k.length > 2; });
        for (var i = 0; i < spans.length; i++) {
          var t = (spans[i].textContent || '').trim().toLowerCase();
          var hits = keywords.filter(function(kw) { return t.indexOf(kw) >= 0; });
          if (hits.length >= 2) { matched = spans[i]; break; }
        }
      }
      // Fallback: match by data-document-type on the file input directly
      if (!matched) {
        var fi = document.querySelector('input[type="file"][data-document-type="' + arguments[0] + '"]');
        if (!fi) {
          // Try case-insensitive
          var all = document.querySelectorAll('input[type="file"][data-document-type]');
          for (var i = 0; i < all.length; i++) {
            if ((all[i].getAttribute('data-document-type') || '').toLowerCase() === target) { fi = all[i]; break; }
          }
        }
        if (fi) {
          // Found via data-document-type — go straight to card
          var outer = fi.closest('[id$="_wtcntDocList"]');
          if (!outer) {
            var p = fi.parentElement;
            for (var d = 0; d < 12 && p; d++) { if (p.id && p.id.indexOf('_wtcntDocList') >= 0) { outer = p; break; } p = p.parentElement; }
          }
          if (outer) {
            outer.scrollIntoView({ block: 'center' });
            var node = fi;
            for (var i = 0; i < 10 && node; i++) { if (node.classList) node.classList.remove('ReadOnly'); node = node.parentElement; }
            fi.removeAttribute('disabled'); fi.removeAttribute('readonly');
            var m = (fi.className || '').match(/doc_(\\d+)/);
            return 'OK|' + fi.id + '|' + (m ? m[1] : '') + '|' + (fi.getAttribute('data-document-type') || '');
          }
        }
        return 'ERR:NO_TITLE';
      }

      // 2. Walk up to the outer card container (div ending with _wtcntDocList)
      var outer = matched.closest('[id$="_wtcntDocList"]');
      if (!outer) {
        // Fallback: walk up until we find a parent containing input[type="file"]
        var p = matched.parentElement;
        for (var d = 0; d < 12 && p; d++) {
          if (p.querySelector('input[type="file"]')) { outer = p; break; }
          p = p.parentElement;
        }
      }
      if (!outer) return 'ERR:NO_CARD';

      // 3. Find the file input inside this specific card
      var fi = outer.querySelector('input[type="file"]');
      if (!fi) return 'ERR:NO_INPUT';

      // 4. Scroll into view + remove ReadOnly
      outer.scrollIntoView({ block: 'center' });
      var node = fi;
      for (var i = 0; i < 10 && node; i++) {
        if (node.classList) node.classList.remove('ReadOnly');
        node = node.parentElement;
      }
      fi.removeAttribute('disabled');
      fi.removeAttribute('readonly');

      // 5. Return id + doc_N number
      var m = (fi.className || '').match(/doc_(\\d+)/);
      return 'OK|' + fi.id + '|' + (m ? m[1] : '') + '|' + (fi.getAttribute('data-document-type') || '');
    `, docType);

    if (inputInfo.startsWith('ERR:')) {
      console.warn(`[Upload] ${inputInfo} for "${docType}"`);
      return;
    }

    const [, fileInputId, docNum, foundType] = inputInfo.split('|');
    console.log(`[Upload] Found card: type="${foundType}", doc_${docNum}, id=...${fileInputId.slice(-40)}`);

    // Step 2: sendKeys on the real file input, then ALWAYS call readUrldoc_N manually.
    //
    //  Why both? sendKeys sets the file and changes the label, but the jQuery
    //  change handler (osjs('.doc_N').bind('change', readUrldoc_N)) often gets
    //  detached after the first upload's AJAX partial-page-refresh. So the native
    //  change event fires but nobody is listening. Calling readUrldoc_N directly
    //  ensures the FileReader → OsNotify → server postback always happens.

    // Wait for any pending OutSystems AJAX to finish before we start this upload.
    // OsNotify from a previous upload may still be in-flight.
    const ajaxIdle = await this.driver.executeScript<boolean>(`
      if (typeof osjs !== 'undefined' && osjs.active > 0) return false;
      if (typeof jQuery !== 'undefined' && jQuery.active > 0) return false;
      return true;
    `);
    if (!ajaxIdle) {
      console.log(`[Upload] Waiting for pending AJAX to settle...`);
      for (let ajaxWait = 0; ajaxWait < 20; ajaxWait++) {
        const ajaxDone = await this.driver.executeScript<boolean>(`
          if (typeof osjs !== 'undefined' && osjs.active > 0) return false;
          if (typeof jQuery !== 'undefined' && jQuery.active > 0) return false;
          return true;
        `);
        if (ajaxDone) break;
        await this.sleep(500);
      }
    }
    await this.sleep(300);

    // ═══════════════════════════════════════════════════════════════════════════
    // STEP 2: Inject file into the input via base64 → DataTransfer
    // ═══════════════════════════════════════════════════════════════════════════
    //
    // How it works:
    //   1. Node reads the local file into a base64 string
    //   2. Browser decodes base64 → Uint8Array → File object
    //   3. DataTransfer API sets input.files = [File]
    //
    // Why this is best:
    //   - Bypasses sendKeys entirely (no event-firing issues)
    //   - Creates a REAL File object identical to user-selected files
    //   - No filesystem path sent to the browser (works with any input state)
    //   - No S3/network dependency
    //   - Works even when the input is hidden/readonly/disabled

    const base64 = base64Data;
    console.log(`[Upload] File: ${fileName} (${(Buffer.from(base64, 'base64').length / 1024).toFixed(1)} KB, ${mimeType})`);

    const injectResult = await this.driver.executeScript<string>(`
      var inputId = arguments[0];
      var b64     = arguments[1];
      var fname   = arguments[2];
      var mime    = arguments[3];

      var fi = document.getElementById(inputId);
      if (!fi) return 'ERR:input gone';

      // Decode base64 → ArrayBuffer → File
      var binary = atob(b64);
      var bytes = new Uint8Array(binary.length);
      for (var i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

      // Create a real File object and inject via DataTransfer
      var file = new File([bytes.buffer], fname, { type: mime, lastModified: Date.now() });
      var dt = new DataTransfer();
      dt.items.add(file);
      fi.files = dt.files;

      if (!fi.files || !fi.files.length) return 'ERR:files not set after DataTransfer';
      return 'OK|files=' + fi.files.length + '|name=' + fi.files[0].name + '|size=' + fi.files[0].size;
    `, fileInputId, base64, fileName, mimeType);

    console.log(`[Upload] DataTransfer inject: ${injectResult}`);
    if (!injectResult.startsWith('OK')) {
      console.warn(`[Upload] File injection failed: ${injectResult} — skipping "${docType}"`);
      return;
    }
    await this.sleep(200);

    // ═══════════════════════════════════════════════════════════════════════════
    // STEP 3: Trigger the upload — change event + readUrldoc_N + OsNotify
    // ═══════════════════════════════════════════════════════════════════════════

    // 3a. Combined: dispatch change event + call readUrldoc_N in a single JS execution.
    //     The change handler may be detached after prior AJAX refreshes, so we always
    //     call readUrldoc_N directly as the reliable upload trigger.
    const callChangeAndReadUrldoc = async (): Promise<string> => {
      return await this.driver.executeScript<string>(`
        var fi = document.getElementById(arguments[0]);
        var docNum = arguments[1];
        if (!fi) return 'ERR:input gone';
        if (!fi.files || !fi.files.length) return 'ERR:no files';
        var log = 'files=' + fi.files.length + ', name=' + fi.files[0].name + '\\n';

        // jQuery/osjs trigger — fires handlers attached via .bind()/.on()
        try {
          if (typeof osjs !== 'undefined' && osjs(fi).trigger) {
            osjs(fi).trigger('change');
            log += 'osjs.trigger(change) OK\\n';
          } else if (typeof jQuery !== 'undefined') {
            jQuery(fi).trigger('change');
            log += 'jQuery.trigger(change) OK\\n';
          }
        } catch(e) { log += 'jq trigger err: ' + e + '\\n'; }

        // Native change event
        try {
          fi.dispatchEvent(new Event('change', { bubbles: true }));
          log += 'native change OK\\n';
        } catch(e) { log += 'native change err: ' + e + '\\n'; }

        // ALWAYS call readUrldoc_N directly — the reliable upload trigger
        var fnName = 'readUrldoc_' + docNum;
        if (typeof window[fnName] === 'function') {
          try { window[fnName].call(fi, fi); log += fnName + '() called OK\\n'; }
          catch(e) { log += fnName + ' error: ' + e + '\\n'; }
        } else {
          log += 'WARN: ' + fnName + ' not found — inline FileReader fallback\\n';
          try {
            var reader = new FileReader();
            reader.onload = function(ev) {
              var base64 = ev.target.result;
              var outer = fi.closest('[id$="_wtcntDocList"]') || fi.parentElement;
              if (outer) {
                var img = outer.querySelector('img');
                if (img) img.src = base64;
                var imgCont = outer.querySelector('[id$="_wtimgContainer"]');
                if (imgCont) imgCont.style.display = '';
              }
              if (typeof OsNotify === 'function') {
                try { OsNotify(fi.id, base64); } catch(e2) {}
              }
            };
            reader.readAsDataURL(fi.files[0]);
            log += 'inline FileReader started\\n';
          } catch(e) { log += 'inline FileReader error: ' + e + '\\n'; }
        }
        return log;
      `, fileInputId, docNum);
    };

    console.log(`[Upload] Triggering change + readUrldoc_${docNum}...`);
    const triggerResult = await callChangeAndReadUrldoc();
    console.log(`[Upload] trigger: ${triggerResult}`);

    // 3c. Wait for FileReader + OsNotify AJAX postback to complete.
    await this.sleep(1000);
    for (let sw = 0; sw < 15; sw++) {
      const settled = await this.driver.executeScript<boolean>(`
        if (typeof osjs !== 'undefined' && osjs.active > 0) return false;
        if (typeof jQuery !== 'undefined' && jQuery.active > 0) return false;
        return true;
      `);
      if (settled) break;
      await this.sleep(500);
    }

    // Step 4: Wait for the upload to process on the server.
    //         Check THIS card's btnContainer visibility.
    console.log(`[Upload] Waiting for "${docType}" to process...`);

    let uploaded = false;
    let retriedReadUrldoc = false;
    for (let attempt = 0; attempt < 15; attempt++) {
      // Check FIRST, then sleep (saves one interval on fast uploads)
      const status = await this.driver.executeScript<string>(`
        // Try finding the card by the original element ID first
        var fi = document.getElementById(arguments[0]);
        var outer = null;
        if (fi) {
          outer = fi.closest('[id$="_wtcntDocList"]');
        }
        // If the old ID is gone (OsNotify re-rendered the DOM), find the card
        // by data-document-type instead — the document type text survives re-renders.
        if (!outer) {
          var target = arguments[1].toLowerCase().trim();
          // Try finding by data-document-type attribute on the new file input
          var allInputs = document.querySelectorAll('input[type="file"][data-document-type]');
          for (var i = 0; i < allInputs.length; i++) {
            var dt = (allInputs[i].getAttribute('data-document-type') || '').toLowerCase().trim();
            if (dt === target || dt.indexOf(target) >= 0 || target.indexOf(dt) >= 0) {
              outer = allInputs[i].closest('[id$="_wtcntDocList"]');
              if (outer) break;
            }
          }
          // Also try by span.Bold text
          if (!outer) {
            var spans = document.querySelectorAll('span.Bold');
            for (var i = 0; i < spans.length; i++) {
              var t = (spans[i].textContent || '').trim().toLowerCase();
              if (t === target || t.indexOf(target) >= 0 || target.indexOf(t) >= 0) {
                outer = spans[i].closest('[id$="_wtcntDocList"]');
                if (outer) break;
              }
            }
          }
        }
        if (!outer) return 'input_gone';
        var btnCont = outer.querySelector('[id$="_wtbtnContainer"]');
        var imgCont = outer.querySelector('[id$="_wtimgContainer"]');
        var label = outer.querySelector('.FileUpload_Label');
        var labelText = label ? label.textContent.trim() : '';
        var btnVisible = btnCont && btnCont.style.display !== 'none';
        var imgVisible = imgCont && imgCont.style.display !== 'none';
        var labelChanged = labelText !== 'Drag here or click to upload a file' && labelText.length > 0;
        if (btnVisible || imgVisible) return 'UPLOADED|btn=' + btnVisible + '|img=' + imgVisible + '|label=' + labelText;
        if (labelChanged) return 'LABEL_CHANGED|' + labelText;
        return 'waiting|btn=' + (btnCont ? btnCont.style.display : 'N/A') + '|img=' + (imgCont ? imgCont.style.display : 'N/A') + '|label=' + labelText;
      `, fileInputId, docType);

      if (status.startsWith('UPLOADED')) {
        console.log(`[Upload] "${docType}" — ${status}`);
        uploaded = true;
        break;
      }

      // If stuck at LABEL_CHANGED after 8s, the OsNotify postback was likely dropped
      // (stale __OSVSTATE). Retry: re-dispatch change event + readUrldoc fallback.
      if (!retriedReadUrldoc && attempt >= 3 && status.startsWith('LABEL_CHANGED')) {
        retriedReadUrldoc = true;
        console.log(`[Upload] "${docType}" — retrying (OsNotify may have been dropped)...`);
        // Wait for any in-flight AJAX to finish first
        for (let w = 0; w < 10; w++) {
          const done = await this.driver.executeScript<boolean>(`
            if (typeof osjs !== 'undefined' && osjs.active > 0) return false;
            if (typeof jQuery !== 'undefined' && jQuery.active > 0) return false;
            return true;
          `);
          if (done) break;
          await this.sleep(500);
        }
        await this.sleep(300);

        // Re-dispatch change event first (handler may have been re-attached after AJAX refresh)
        await this.driver.executeScript(`
          var fi = document.getElementById(arguments[0]);
          if (!fi || !fi.files || !fi.files.length) return;
          try { if (typeof osjs !== 'undefined') osjs(fi).trigger('change'); } catch(e) {}
          try { fi.dispatchEvent(new Event('change', { bubbles: true })); } catch(e) {}
        `, fileInputId);
        await this.sleep(500);

        // Check if the change event retry worked
        const retryWorked = await this.driver.executeScript<boolean>(`
          var fi = document.getElementById(arguments[0]);
          if (!fi) return false;
          var outer = fi.closest('[id$="_wtcntDocList"]');
          if (!outer) return false;
          var btnCont = outer.querySelector('[id$="_wtbtnContainer"]');
          var imgCont = outer.querySelector('[id$="_wtimgContainer"]');
          return (btnCont && btnCont.style.display !== 'none') || (imgCont && imgCont.style.display !== 'none');
        `, fileInputId);

        if (retryWorked) {
          console.log(`[Upload] "${docType}" — change event retry succeeded`);
          continue;
        }

        // Still no preview — call readUrldoc_N manually as last resort
        const retryResult = await callChangeAndReadUrldoc();
        console.log(`[Upload] readUrldoc retry: ${retryResult}`);
        // Wait for FileReader + OsNotify AJAX postback to complete
        await this.sleep(1000);
        for (let sw2 = 0; sw2 < 10; sw2++) {
          const settled = await this.driver.executeScript<boolean>(`
            if (typeof osjs !== 'undefined' && osjs.active > 0) return false;
            if (typeof jQuery !== 'undefined' && jQuery.active > 0) return false;
            return true;
          `);
          if (settled) break;
          await this.sleep(500);
        }
        // Do NOT click SaveButton — it triggers a full page POST that navigates away.
        // OsNotify's AJAX postback handles the upload entirely.
      }

      // Accept LABEL_CHANGED as soft success after 3 checks — file IS on server,
      // UI preview just hasn't rendered yet. Don't waste time waiting.
      if (attempt >= 3 && status.startsWith('LABEL_CHANGED')) {
        console.log(`[Upload] "${docType}" — accepting LABEL_CHANGED as soft success (file on server)`);
        uploaded = true;
        break;
      }

      console.log(`[Upload] "${docType}" — ${status} (${attempt + 1}/15)`);
      await this.sleep(1000);
    }

    if (!uploaded) {
      // Take debug screenshot
      try {
        const screenshot = await this.driver.takeScreenshot();
        const ssPath = path.resolve('test-results', `upload-fail-${docType.replace(/[^a-zA-Z0-9]/g, '_')}.png`);
        fs.mkdirSync(path.dirname(ssPath), { recursive: true });
        fs.writeFileSync(ssPath, screenshot, 'base64');
        console.log(`[Upload] Debug screenshot: ${ssPath}`);
      } catch {}
      console.warn(`[Upload] "${docType}" — upload did NOT complete after 30s.`);
    }

    // Wait for page to fully settle before next upload
    await this.waitForPageLoad();
    await this.sleep(500);
  }

  /**
   * Clicks Continue on the upload page. Retries for up to 60 seconds
   * since the button only appears after all mandatory uploads complete.
   */
  private async clickUploadContinue(): Promise<void> {
    for (let retry = 0; retry < 20; retry++) {
      // Try CSS selectors
      for (const sel of [
        'input[value="Continue"]', 'a[id*="Continue"]', 'button[id*="Continue"]',
        'input[id*="Continue"]', '[data-staticid*="Continue"]',
      ]) {
        try {
          const btn = await this.driver.findElement(By.css(sel));
          if (await btn.isDisplayed().catch(() => false)) {
            await this.driver.executeScript('arguments[0].scrollIntoView({block:"center"});', btn);
            await this.sleep(500);
            await btn.click();
            await this.waitForPageLoad();
            console.log(`[Upload] Continue clicked.`);
            return;
          }
        } catch {}
      }

      // Try text content match
      const found = await this.driver.executeScript<boolean>(`
        var btns = document.querySelectorAll('input[type="submit"], input[type="button"], button, a');
        for (var i = 0; i < btns.length; i++) {
          var t = (btns[i].textContent || btns[i].value || '').trim().toLowerCase();
          if ((t === 'continue' || t === 'next' || t === 'submit')
              && window.getComputedStyle(btns[i]).display !== 'none') {
            btns[i].scrollIntoView({ block: 'center' });
            btns[i].click();
            return true;
          }
        }
        return false;
      `);
      if (found) {
        await this.waitForPageLoad();
        console.log('[Upload] Continue clicked.');
        return;
      }

      console.log(`[Upload] Continue not visible yet (${retry + 1}/20)...`);
      await this.sleep(3000);
    }

    console.warn('[Upload] Continue button not found after 60s.');
  }

  // ── Navigation ─────────────────────────────────────────────────────────────

  private async dismissPromoPopup(): Promise<void> {
    const SKIP_ID = 'WebPatterns_wt2_block_wtMainContent_wt3_EmaratechSG_Patterns_wt8_block_wtMainContent_wt10';
    try {
      // Check for iframe popup
      const iframes = await this.driver.findElements(By.css('iframe'));
      for (const iframe of iframes) {
        try {
          await this.driver.switchTo().frame(iframe);
          const skipBtns = await this.driver.findElements(By.css(`#${SKIP_ID}, input[value="Skip"]`));
          for (const btn of skipBtns) {
            if (await btn.isDisplayed().catch(() => false)) {
              await btn.click();
              console.log('[Nav] Dismissed promotional popup.');
              await this.driver.switchTo().defaultContent();
              return;
            }
          }
          await this.driver.switchTo().defaultContent();
        } catch {
          await this.driver.switchTo().defaultContent();
        }
      }
    } catch { /* non-fatal — popup does not appear on every load */ }
  }

  private async navigateToNewApplication(): Promise<void> {
    console.log('[Nav] Navigating to Existing Applications...');

    // Retry up to 3 times with page refresh — the dashboard sometimes loads
    // without rendering the Existing Applications link (AJAX race condition).
    const existAppSel = '#EmaratechSG_Theme_wtwbLayoutEmaratech_block_wtMainContent_wtwbDashboard_wtCntExistApp';
    let clicked = false;

    for (let attempt = 0; attempt < 3 && !clicked; attempt++) {
      if (attempt > 0) {
        console.log(`[Nav] Retry ${attempt}/2 — refreshing dashboard page...`);
        await this.driver.get(GdrfaPortalPage.HOME);
        await this.waitForPageLoad();
        await this.sleep(3000);
      }

      await this.dismissPromoPopup();

      // Strategy 1: Direct ID selector
      try {
        const el = await this.waitForElement(By.css(existAppSel), 15000);
        if (await el.isDisplayed().catch(() => false)) {
          await this.driver.executeScript('arguments[0].scrollIntoView({block:"center"});', el);
          await this.sleep(300);
          try {
            await el.click();
          } catch {
            await this.driver.executeScript('arguments[0].click();', el);
          }
          clicked = true;
          continue;
        }
      } catch {}

      // Strategy 2: Partial link text
      try {
        const link = await this.driver.findElement(By.partialLinkText('Existing Applications'));
        if (await link.isDisplayed().catch(() => false)) {
          await this.driver.executeScript('arguments[0].scrollIntoView({block:"center"});', link);
          await this.sleep(300);
          await link.click();
          clicked = true;
          continue;
        }
      } catch {}

      // Strategy 3: JS click any element containing the text
      try {
        const jsClicked = await this.driver.executeScript<boolean>(`
          var els = document.querySelectorAll('a, span, div, button');
          for (var i = 0; i < els.length; i++) {
            var t = (els[i].textContent || '').trim();
            if (t === 'Existing Applications' || t === 'existing applications') {
              els[i].scrollIntoView({block:'center'});
              els[i].click();
              return true;
            }
          }
          return false;
        `);
        if (jsClicked) { clicked = true; continue; }
      } catch {}
    }

    if (!clicked) throw new Error('[Nav] Could not find Existing Applications link');

    // Wait for the establishment detail page to fully load
    await this.waitForUrlContains('EstablishmentDetail', 20000).catch(() => {});
    await this.waitForPageLoad();
    await this.dismissPromoPopup();

    const dropdownSel =
      '#EmaratechSG_Theme_wtwbLayoutEmaratechWithoutTitle_block_wtMainContent_EmaratechSG_Patterns_wtwbEstbButtonWithContextInfo_block_wtIcon_wtcntContextActionBtn';
    const firstOptionSel =
      '#EmaratechSG_Theme_wtwbLayoutEmaratechWithoutTitle_block_wtMainContent_EmaratechSG_Patterns_wtwbEstbButtonWithContextInfo_block_wtContent_wtwbEstbTopServices_wtListMyServicesExperiences_ctl00_wtStartTopService';

    let dropdownVisible = false;
    try {
      const dropdown = await this.waitForElement(By.css(dropdownSel), 15000);
      dropdownVisible = await dropdown.isDisplayed();
    } catch {}

    if (dropdownVisible) {
      const dropdown = await this.driver.findElement(By.css(dropdownSel));
      await dropdown.click();
      await this.sleep(150);

      const firstOption = await this.waitForElement(By.css(firstOptionSel), 10000);
      const optionText = await firstOption.getText();
      console.log(`[Nav] Selecting form: "${optionText.trim()}"`);
      await firstOption.click();
    } else {
      // Fallback: dropdown ID not found
      console.log('[Nav] Dropdown not found — using fallback navigation.');
      try {
        const newAppBtn = await this.driver.findElement(By.xpath('//button[contains(text(),"New Application")]'));
        if (await newAppBtn.isDisplayed().catch(() => false)) {
          await newAppBtn.click();
          await this.sleep(300);
        }
      } catch {}

      let firstService: WebElement;
      try {
        firstService = await this.driver.findElement(By.css(firstOptionSel));
      } catch {
        firstService = await this.driver.findElement(By.partialLinkText('New Tourism Entry Permit'));
      }
      await this.driver.wait(until.elementIsVisible(firstService), 10000);
      const serviceText = await firstService.getText();
      console.log(`[Nav] Selecting form (fallback): "${serviceText.trim()}"`);
      await firstService.click();
    }

    // Wait for actual form page (EntryPermitTourism.aspx) to load — not just the EstablishmentDetail page
    console.log('[Nav] Waiting for form page to load...');
    try {
      await this.waitForUrlContains('EntryPermit', 30000);
    } catch {
      // Some portal versions load the form on the same page via AJAX
      console.warn('[Nav] URL did not change to EntryPermit — checking if form loaded on current page...');
    }
    await this.waitForPageLoad();
    await this.sleep(2000);

    // Verify the actual form is present by checking for known form elements
    const formReady = await this.waitForFormElements(30000);
    if (!formReady) {
      throw new Error('[Nav] Form page did not load — no form elements found after 30s.');
    }

    const currentUrl = await this.driver.getCurrentUrl();
    console.log('[Nav] Form page loaded. URL:', currentUrl);
  }

  /**
   * Waits until at least some expected form elements are present (Select2 dropdowns, form sections).
   */
  private async waitForFormElements(timeout: number): Promise<boolean> {
    const deadline = Date.now() + timeout;
    while (Date.now() < deadline) {
      const found = await this.driver.executeScript<boolean>(`
        // Check for Select2 containers (dropdowns for Passport Type, Nationality, etc.)
        var s2 = document.querySelectorAll('.select2-container');
        if (s2.length >= 3) return true;
        // Check for known form sections
        var labels = document.querySelectorAll('label');
        for (var i = 0; i < labels.length; i++) {
          var t = (labels[i].textContent || '').toLowerCase();
          if (t.indexOf('passport type') >= 0 || t.indexOf('passport number') >= 0) return true;
        }
        return false;
      `);
      if (found) return true;
      await this.sleep(1000);
    }
    return false;
  }

  private async waitForPageSettle(): Promise<void> {
    await this.waitForPageLoad();
    await this.sleep(500);
  }

  // ── Physical Select2 click helper ──────────────────────────────────────────

  /**
   * Physically clicks a Select2 dropdown open and selects an option.
   * Works by finding the label text on the page, locating the nearest Select2,
   * clicking it with Selenium, and clicking the matching option.
   *
   * @param labelText - Text near the dropdown (e.g. "Visit Reason", "Passport Type")
   * @param optionText - Option to select (e.g. "Tourism", "Normal")
   * @returns true if successfully selected
   */
  private async clickSelect2ByLabel(labelText: string, optionText: string): Promise<boolean> {
    // Step 1: Find the Select2 container near the label, remove ReadOnly, and return its index.
    // Uses two strategies: (A) find <label> by text → find associated <select> → find its Select2,
    // (B) walk up from each Select2 container checking siblings for label text.
    const s2Index = await this.driver.executeScript<number>(`
      var label = arguments[0];
      var normalLabel = label.replace(/[\\s*]/g, '').toLowerCase();

      // ── Strategy A: Find <label> element → associated <select> → its Select2 ──
      var lbls = Array.from(document.querySelectorAll('label'));
      for (var li = 0; li < lbls.length; li++) {
        var lbl = lbls[li];
        if (lbl.classList.contains('select2-offscreen')) continue;
        var lblText = (lbl.textContent || '').replace(/[\\s*]/g, '').toLowerCase();
        if (lblText !== normalLabel && lblText.indexOf(normalLabel) < 0) continue;

        // Found the label — now find associated select
        var sel = null;
        if (lbl.htmlFor) {
          var el = document.getElementById(lbl.htmlFor);
          if (el instanceof HTMLSelectElement) sel = el;
          else if (el) {
            var row = el.closest('.ThemeGrid_Width6,.ThemeGrid_Width12,.FormTitle');
            if (row && row.parentElement) sel = row.parentElement.querySelector('select');
            if (!sel && el.parentElement) sel = el.parentElement.querySelector('select');
          }
        }
        // Walk up from label to find a nearby select
        if (!sel) {
          var row = lbl.closest('.ThemeGrid_Width6,.ThemeGrid_Width12,.FormTitle');
          if (row && row.parentElement) sel = row.parentElement.querySelector('select');
          if (!sel && row && row.nextElementSibling) sel = row.nextElementSibling.querySelector('select');
        }
        if (!sel) continue;

        // Find the Select2 container for this select
        var s2 = sel.previousElementSibling;
        if (!s2 || !s2.classList.contains('select2-container')) {
          s2 = document.getElementById('s2id_' + sel.id);
        }
        if (!s2 || !s2.classList.contains('select2-container')) continue;

        // Get index among all Select2 containers
        var allS2 = document.querySelectorAll('.select2-container');
        for (var idx = 0; idx < allS2.length; idx++) {
          if (allS2[idx] === s2) {
            // Remove ReadOnly
            s2.classList.remove('ReadOnly');
            s2.querySelectorAll('.ReadOnly').forEach(function(el) { el.classList.remove('ReadOnly'); });
            var p = s2.parentElement;
            for (var k = 0; k < 5 && p; k++) { p.classList.remove('ReadOnly'); p = p.parentElement; }
            return idx;
          }
        }
      }

      // ── Strategy B: Walk up from each Select2 container checking siblings ──
      var s2s = document.querySelectorAll('.select2-container');
      for (var i = 0; i < s2s.length; i++) {
        var s2 = s2s[i];
        var parent = s2.parentElement;
        for (var depth = 0; depth < 8 && parent; depth++) {
          var siblings = parent.children;
          for (var j = 0; j < siblings.length; j++) {
            var sib = siblings[j];
            if (sib === s2 || sib.querySelector('.select2-container')) continue;
            var text = (sib.textContent || '').trim();
            if (text.length < 80 && text.replace(/[\\s*]/g, '').toLowerCase().indexOf(normalLabel) >= 0) {
              s2.classList.remove('ReadOnly');
              s2.querySelectorAll('.ReadOnly').forEach(function(el) { el.classList.remove('ReadOnly'); });
              var p = s2.parentElement;
              for (var k = 0; k < 5 && p; k++) { p.classList.remove('ReadOnly'); p = p.parentElement; }
              return i;
            }
          }
          parent = parent.parentElement;
        }
      }

      return -1;
    `, labelText);

    if (s2Index < 0) {
      console.warn(`[Select2] No dropdown found near "${labelText}".`);
      return false;
    }

    // Get the container with Selenium and check if already set
    const s2Containers = await this.driver.findElements(By.css('.select2-container'));
    if (s2Index >= s2Containers.length) return false;
    const container = s2Containers[s2Index];

    try {
      const chosenSpan = await container.findElement(By.css('.select2-chosen'));
      const currentText = await chosenSpan.getText();
      if (currentText.toUpperCase().includes(optionText.toUpperCase())) {
        console.log(`[Skip] "${labelText}" already set: "${currentText.trim()}".`);
        return true;
      }
    } catch {}

    // ── Step 2: Clear overlays, scroll into view and click the dropdown open ──
    await this.clearBlockingOverlays();
    await this.driver.executeScript('arguments[0].scrollIntoView({block:"center"});', container);
    await this.sleep(300);

    // Use Selenium's real click — Select2 needs real browser events
    try {
      const choiceLink = await container.findElement(By.css('.select2-choice, a.select2-choice'));
      await choiceLink.click();
    } catch {
      await container.click();
    }
    await this.sleep(800);

    // ── Step 3: Unlock the dropdown panel (ReadOnly, pencil icon, search input) ──
    await this.driver.executeScript(`
      var drop = document.querySelector('.select2-drop-active');
      if (!drop) return;
      drop.classList.remove('ReadOnly');
      drop.querySelectorAll('*').forEach(function(el) { el.classList.remove('ReadOnly'); });

      // Click pencil/edit icon if present (OutSystems ReadOnly pattern)
      var pencil = drop.querySelector('.FormEditPencil');
      if (pencil) pencil.click();

      // Unlock the search input
      var input = drop.querySelector('.select2-input');
      if (input) {
        input.removeAttribute('readonly');
        input.removeAttribute('disabled');
        input.classList.remove('ReadOnly');
        input.style.display = '';
        input.style.visibility = 'visible';
      }
    `);
    await this.sleep(300);

    // ── Step 4: Check if dropdown has a search input — type with real Selenium keys ──
    let searchInput: WebElement | null = null;
    try {
      searchInput = await this.driver.findElement(
        By.css('.select2-drop-active .select2-input')
      );
    } catch {}

    if (searchInput) {
      // Focus and clear the input, then type character by character (triggers Select2's AJAX)
      await searchInput.click();
      await this.sleep(200);
      await searchInput.clear();
      for (const char of optionText) {
        await searchInput.sendKeys(char);
        await this.sleep(60);
      }

      // Wait for AJAX results to load
      console.log(`[Select2] "${labelText}" — typed "${optionText}", waiting for results...`);
      await this.waitForCondition(async () => {
        return this.driver.executeScript<boolean>(`
          var results = document.querySelectorAll('.select2-drop-active .select2-results li.select2-result');
          if (results.length === 0) return false;
          // Make sure it's not the "Searching..." or "No results" message
          var text = results[0].textContent.toLowerCase();
          return text.indexOf('searching') < 0 && text.indexOf('no ') !== 0;
        `);
      }, 10000).catch(() => {
        console.warn(`[Select2] "${labelText}" — no AJAX results after 10s.`);
      });
      await this.sleep(500);
    }

    // ── Step 5: Find and click the best matching result ──
    // First identify which result to click via JS (fast), then click it with Selenium (reliable)
    const resultIndex = await this.driver.executeScript<number>(`
      var search = arguments[0].toUpperCase();
      var results = document.querySelectorAll('.select2-drop-active .select2-results li.select2-result');
      if (results.length === 0) results = document.querySelectorAll('.select2-drop-active .select2-results li');
      if (results.length === 0) return -1;

      // Pass 1: exact match (strip number prefix like "349 - ")
      for (var i = 0; i < results.length; i++) {
        var t = results[i].textContent.replace(/^\\d+\\s*-\\s*/, '').trim();
        if (t.toUpperCase() === search) return i;
      }
      // Pass 2: starts-with (stripped)
      for (var i = 0; i < results.length; i++) {
        var t = results[i].textContent.replace(/^\\d+\\s*-\\s*/, '').trim();
        if (t.toUpperCase().indexOf(search) === 0) return i;
      }
      // Pass 3: contains
      for (var i = 0; i < results.length; i++) {
        if (results[i].textContent.toUpperCase().indexOf(search) >= 0) return i;
      }
      // Pass 4: if only one result and it's not "no results", take it
      if (results.length === 1) {
        var t = results[0].textContent.toLowerCase();
        if (t.indexOf('no ') !== 0) return 0;
      }
      return -1;
    `, optionText);

    if (resultIndex >= 0) {
      // Click with Selenium's real click (Select2 needs real browser events)
      const resultElements = await this.driver.findElements(
        By.css('.select2-drop-active .select2-results li.select2-result, .select2-drop-active .select2-results li')
      );
      if (resultIndex < resultElements.length) {
        const matchedText = await resultElements[resultIndex].getText().catch(() => '');
        console.log(`[Select2] "${labelText}" — clicking result: "${matchedText.trim()}"...`);
        await resultElements[resultIndex].click();
        await this.sleep(500);

        // Verify dropdown closed — if not, force selection via mousedown/mouseup/click
        const stillOpen = await this.driver.executeScript<boolean>(
          `return !!document.querySelector('.select2-drop-active');`
        );
        if (stillOpen) {
          console.log(`[Select2] "${labelText}" — dropdown still open, using JS mouse events...`);
          await this.driver.executeScript(`
            var results = document.querySelectorAll('.select2-drop-active .select2-results li.select2-result, .select2-drop-active .select2-results li');
            var item = results[arguments[0]];
            if (item) {
              item.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
              item.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
              item.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
            }
          `, resultIndex);
          await this.sleep(500);
        }

        // Final fallback: programmatic jQuery val + trigger
        const stillOpen2 = await this.driver.executeScript<boolean>(
          `return !!document.querySelector('.select2-drop-active');`
        );
        if (stillOpen2) {
          console.log(`[Select2] "${labelText}" — forcing via jQuery...`);
          await this.driver.executeScript(`
            var s2s = document.querySelectorAll('.select2-container');
            var s2 = s2s[arguments[0]];
            if (!s2) return;
            var sel = s2.nextElementSibling;
            if (!sel || sel.tagName !== 'SELECT') return;
            var text = arguments[1];
            var opt = Array.from(sel.options).find(function(o) {
              return o.text.toUpperCase().indexOf(text.toUpperCase()) >= 0;
            });
            if (opt) {
              sel.value = opt.value;
              var jq = window.jQuery || window.$;
              if (jq) jq(sel).val(opt.value).trigger('change');
              else sel.dispatchEvent(new Event('change', { bubbles: true }));
            }
          `, s2Index, optionText);
        }

        await this.closeOpenSelect2Dropdowns();
        await this.waitForPageLoad();
        await this.sleep(500);
        return true;
      }
    }

    // ── Failed — log options for debugging ──
    const firstOptions = await this.driver.executeScript<string[]>(`
      var results = document.querySelectorAll('.select2-drop-active .select2-results li');
      var arr = [];
      for (var i = 0; i < Math.min(results.length, 10); i++) arr.push(results[i].textContent.trim());
      return arr;
    `);
    console.log(`[Select2] "${labelText}" — ${firstOptions.length} option(s) visible:`);
    for (const opt of firstOptions) console.log(`  "${opt}"`);

    await this.closeOpenSelect2Dropdowns();
    console.warn(`[Select2] "${optionText}" not found in "${labelText}" dropdown.`);
    return false;
  }

  // ── Passport header (Passport Type → Nationality → Search Data) ────────────

  private async setVisitReason(): Promise<void> {
    console.log('[Form] Selecting Visit Reason → Tourism...');
    const ok = await this.clickSelect2ByLabel('Visit Reason', 'Tourism');
    if (!ok) {
      // Fallback: try selectByLabel
      const result = await this.selectByLabel('Visit Reason', 'Tourism');
      if (result.found) {
        console.log(`[Form] Visit Reason set via fallback: "${result.matched}".`);
      } else {
        console.warn('[Form] Visit Reason: FAILED — will retry in validation pass.');
      }
    }
  }

  private async setPassportType(passportType: string): Promise<void> {
    console.log(`[Form] Setting Passport Type → "${passportType}"...`);
    // Try 1: Physical Select2 click
    const ok = await this.clickSelect2ByLabel('Passport Type', passportType);
    if (ok) return;
    // Try 2: selectByLabel (programmatic + jQuery trigger)
    console.log('[Form] Passport Type — clickSelect2 failed, trying selectByLabel...');
    const result = await this.selectByLabel('Passport Type', passportType);
    if (result.found) {
      await this.waitForPageLoad();
      console.log(`[Form] Passport Type set via fallback: "${result.matched}".`);
      return;
    }
    // Try 3: Force via jQuery on any select near the label
    console.log('[Form] Passport Type — selectByLabel failed, trying direct jQuery...');
    await this.forceSelect2ByLabelJS('Passport Type', passportType);
  }

  private async enterPassportNumber(passportNumber: string): Promise<void> {
    console.log(`[Form] Entering Passport Number: "${passportNumber}"...`);
    const input = await this.waitForElement(
      By.css('input[staticid*="PassportNo"], input[id*="inptPassportNo"], input[id*="PassportNo"]'),
      15000
    );
    await this.driver.executeScript('arguments[0].scrollIntoView({block:"center"});', input);
    await input.clear();
    await input.sendKeys(passportNumber);
    console.log('[Form] Passport Number entered.');
  }

  private async setNationality(nationalityCode: string): Promise<void> {
    const name = GdrfaPortalPage.mrzCodeToCountryName(nationalityCode);
    console.log(`[Form] Setting Current Nationality: "${name}"...`);
    // Try 1: Physical Select2 click
    const ok = await this.clickSelect2ByLabel('Current Nationality', name);
    if (ok) {
      await this.waitForPageLoad();
      return;
    }
    // Try 2: selectByLabel (programmatic + jQuery trigger)
    console.log('[Form] Current Nationality — clickSelect2 failed, trying selectByLabel...');
    const result = await this.selectByLabel('Current Nationality', name);
    if (result.found) {
      await this.waitForPageLoad();
      console.log(`[Form] Current Nationality set: "${result.matched}".`);
      return;
    }
    // Try 3: Force via jQuery on nationality-pattern selects
    console.log('[Form] Current Nationality — selectByLabel failed, trying direct jQuery...');
    await this.forceSelect2ByLabelJS('Current Nationality', name);
    await this.waitForPageLoad();
  }

  private async setPreviousNationality(nationalityCode: string): Promise<void> {
    const name = GdrfaPortalPage.mrzCodeToCountryName(nationalityCode);
    console.log(`[Form] Setting Previous Nationality: "${name}"...`);
    // Try 1: Physical Select2 click
    const ok = await this.clickSelect2ByLabel('Previous Nationality', name);
    if (ok) {
      await this.waitForPageLoad();
      return;
    }
    // Try 2: selectByLabel (programmatic + jQuery trigger)
    console.log('[Form] Previous Nationality — clickSelect2 failed, trying selectByLabel...');
    const result = await this.selectByLabel('Previous Nationality', name);
    if (result.found) {
      await this.waitForPageLoad();
      console.log(`[Form] Previous Nationality set: "${result.matched}".`);
      return;
    }
    // Try 3: Force via jQuery on the specific select
    console.log('[Form] Previous Nationality — selectByLabel failed, trying direct jQuery...');
    await this.forceSelect2ByLabelJS('Previous Nationality', name);
    await this.waitForPageLoad();
  }

  /**
   * Last-resort dropdown setter: finds a <select> near the given label text,
   * removes ReadOnly, sets the value via jQuery .val().trigger('change'),
   * and manually updates the Select2 display.
   */
  private async forceSelect2ByLabelJS(labelText: string, searchValue: string): Promise<void> {
    const result = await this.driver.executeScript<{ found: boolean; matched: string }>(`
      var labelTarget = arguments[0];
      var search = arguments[1];

      // Strategy 1: Find by <label> element
      var labels = Array.from(document.querySelectorAll('label'));
      var lbl = labels.find(function(l) {
        var t = (l.textContent || '').replace(/[\\s*]/g, '').toLowerCase();
        return t === labelTarget.replace(/[\\s*]/g, '').toLowerCase();
      });

      var sel = null;
      if (lbl) {
        // Try label.htmlFor
        if (lbl.htmlFor) {
          var el = document.getElementById(lbl.htmlFor);
          if (el instanceof HTMLSelectElement) sel = el;
          else if (el) sel = el.closest('.ThemeGrid_Width6,.ThemeGrid_Width12')
            ? el.closest('.ThemeGrid_Width6,.ThemeGrid_Width12').querySelector('select')
            : (el.parentElement ? el.parentElement.querySelector('select') : null);
        }
        // Walk up to find a nearby select
        if (!sel) {
          var row = lbl.closest('.ThemeGrid_Width6,.FormTitle,.ThemeGrid_Width12');
          if (row && row.parentElement) sel = row.parentElement.querySelector('select');
          if (!sel && row && row.nextElementSibling) sel = row.nextElementSibling.querySelector('select');
        }
      }

      // Strategy 2: Find by nearby text in parent containers (like clickSelect2ByLabel does)
      if (!sel) {
        var selects = document.querySelectorAll('select');
        for (var i = 0; i < selects.length; i++) {
          var s = selects[i];
          var parent = s.parentElement;
          for (var d = 0; d < 6 && parent; d++) {
            var text = (parent.textContent || '').replace(/[\\s*]/g, '').toLowerCase();
            if (text.indexOf(labelTarget.replace(/[\\s*]/g, '').toLowerCase()) >= 0) {
              sel = s;
              break;
            }
            parent = parent.parentElement;
          }
          if (sel) break;
        }
      }

      if (!sel) return { found: false, matched: 'No select found near: ' + labelTarget };

      // Remove ReadOnly from Select2 container
      var s2 = sel.previousElementSibling;
      if (s2 && s2.classList.contains('select2-container')) {
        s2.classList.remove('ReadOnly');
        s2.querySelectorAll('.ReadOnly').forEach(function(el) { el.classList.remove('ReadOnly'); });
      }
      // Also walk up parents
      var p = sel.parentElement;
      for (var k = 0; k < 5 && p; k++) { p.classList.remove('ReadOnly'); p = p.parentElement; }
      sel.removeAttribute('disabled');

      // Find matching option
      var opts = Array.from(sel.options);
      var match = opts.find(function(o) { return o.text.trim().toUpperCase() === search.toUpperCase(); })
        || opts.find(function(o) {
          // Strip leading number prefix (e.g. "349 - SOUTH AFRICA" → "SOUTH AFRICA")
          var clean = o.text.replace(/^\\d+\\s*-\\s*/, '').trim();
          return clean.toUpperCase() === search.toUpperCase();
        })
        || opts.find(function(o) { return o.text.toUpperCase().indexOf(search.toUpperCase()) >= 0; });
      if (!match) return { found: false, matched: 'Option not found: ' + search };

      // Set value with jQuery to trigger Select2 update
      sel.value = match.value;
      var jq = window.jQuery || window.$;
      if (jq) {
        jq(sel).val(match.value).trigger('change');
      } else {
        sel.dispatchEvent(new Event('change', { bubbles: true }));
      }

      // Manually update Select2 display text
      if (s2) {
        var chosen = s2.querySelector('.select2-chosen');
        if (chosen) chosen.textContent = match.text.trim();
      }

      return { found: true, matched: match.text.trim() };
    `, labelText, searchValue);

    if (result.found) {
      console.log(`[Form] ${labelText} force-set: "${result.matched}".`);
    } else {
      console.warn(`[Form] ${labelText} force-set FAILED: ${result.matched}`);
    }
  }

  private async clickSearchDataAndWait(): Promise<void> {
    // Clear overlays and Select2 dropdowns that might intercept the click
    await this.clearBlockingOverlays();
    await this.closeOpenSelect2Dropdowns();
    await this.sleep(300);

    console.log('[Form] Clicking Search Data...');
    const btn = await this.waitForElement(
      By.xpath('//a[contains(text(),"Search Data")] | //button[contains(text(),"Search Data")] | //input[@value="Search Data"]'),
      10000
    );
    await this.driver.executeScript('arguments[0].scrollIntoView({block:"center"});', btn);
    await this.sleep(300);
    // Use JS click as fallback in case something still overlaps
    try {
      await btn.click();
    } catch {
      console.log('[Form] Search Data native click blocked — using JS click...');
      await this.driver.executeScript('arguments[0].click();', btn);
    }
    // Wait for the portal AJAX call to populate SmartInput fields
    await this.waitForPageLoad();

    // Wait for key passport input fields to appear in the DOM
    console.log('[Form] Waiting for passport fields to load...');
    await Promise.all([
      this.waitForElement(By.css('input[data-staticid="inpFirsttNameEn"]'), 20000),
      this.waitForElement(By.css('input[data-staticid="inpLastNameEn"]'), 20000),
      this.waitForElement(By.css('input[data-staticid="inpDateOfBirth"]'), 20000),
      this.waitForElement(By.css('input[data-staticid="inpPassportExpiryDate"]'), 20000),
    ]);
    console.log('[Form] Search Data complete — passport fields populated.');
  }

  // ── Passport name fields (First / Middle / Last) ───────────────────────────

  private async fillPassportNames(passport: PassportDetails): Promise<void> {
    console.log('[Form] Clearing Arabic name fields...');
    await this.clearArField('inpFirstNameAr');
    await this.clearArField('inpMiddleNameAr');
    await this.clearArField('inpLastNameAr');

    console.log(`[Form] Filling First Name: "${passport.firstName}"...`);
    const firstFilled = await this.editAndFill('inpFirsttNameEn', passport.firstName);
    if (firstFilled) {
      await this.driver.executeScript(`if (window.translateInputText) translateInputText('inpFirsttNameEn');`);
      await this.sleep(200);
      console.log('[Form] First Name filled + translated.');
    }

    if (passport.middleName) {
      console.log(`[Form] Filling Middle Name: "${passport.middleName}"...`);
      const midFilled = await this.editAndFill('inpMiddleNameEn', passport.middleName);
      if (midFilled) {
        await this.driver.executeScript(`if (window.translateInputText) translateInputText('inpMiddleNameEn');`);
        await this.sleep(200);
        console.log('[Form] Middle Name filled + translated.');
      }
    } else {
      console.log('[Form] No middle name — field left blank.');
    }

    console.log(`[Form] Filling Last Name: "${passport.lastName}"...`);
    const lastFilled = await this.editAndFill('inpLastNameEn', passport.lastName);
    if (lastFilled) {
      await this.driver.executeScript(`if (window.translateInputText) translateInputText('inpLastNameEn');`);
      await this.sleep(200);
      console.log('[Form] Last Name filled + translated.');
    }
  }

  // ── Passport detail fields (below name fields) ─────────────────────────────

  private async fillPassportDetails(passport: PassportDetails): Promise<void> {
    // ── DROPDOWNS FIRST (they trigger AJAX postbacks that can wipe text/date fields) ──

    const birthCountry = GdrfaPortalPage.mrzCodeToCountryName(passport.birthCountry);
    console.log(`[Form] Setting Birth Country: "${birthCountry}"...`);
    const bcResult = await this.selectByLabel('Birth Country', birthCountry);
    if (bcResult.skipped) {
      console.log(`[Skip] Birth Country already set: "${bcResult.matched}".`);
    } else if (bcResult.found) {
      await this.waitForPageLoad();
      await this.sleep(500);
      console.log(`[Form] Birth Country set: "${bcResult.matched}".`);
    } else {
      console.warn(`[Form] Birth Country not found for: "${birthCountry}".`);
    }

    console.log(`[Form] Setting Gender: "${passport.gender}"...`);
    const gResult = await this.selectByLabel('Gender', passport.gender);
    if (gResult.skipped) {
      console.log(`[Skip] Gender already set: "${gResult.matched}".`);
    } else if (gResult.found) {
      await this.waitForPageLoad();
      await this.sleep(500);
      console.log(`[Form] Gender set: "${gResult.matched}".`);
    } else {
      console.warn(`[Form] Gender not found for: "${passport.gender}".`);
    }

    if (passport.passportIssueCountry) {
      const issueCountry = GdrfaPortalPage.mrzCodeToCountryName(passport.passportIssueCountry);
      console.log(`[Form] Setting Passport Issue Country: "${issueCountry}"...`);
      const icResult = await this.selectByLabel('Passport Issue Country', issueCountry);
      if (icResult.skipped) {
        console.log(`[Skip] Passport Issue Country already set: "${icResult.matched}".`);
      } else if (icResult.found) {
        await this.waitForPageLoad();
        await this.sleep(500);
        console.log(`[Form] Passport Issue Country set: "${icResult.matched}".`);
      } else {
        console.warn(`[Form] Passport Issue Country not found for: "${issueCountry}".`);
      }
    } else {
      console.log('[Form] Passport Issue Country — skipped (empty).');
    }

    // ── DATE FIELDS (after all dropdowns, so AJAX won't wipe them) ──

    // Date of Birth
    const dob = passport.dateOfBirth.replace(/\//g, '-');
    console.log(`[Form] Filling Date of Birth: "${dob}"...`);
    const currentDob = await this.driver.executeScript<string>(`
      var el = document.querySelector('input[data-staticid="inpDateOfBirth"]');
      return el ? (el.value || '').trim() : '';
    `);
    if (currentDob && currentDob.toUpperCase() === dob.toUpperCase()) {
      console.log(`[Skip] Date of Birth already has correct value: "${currentDob}".`);
    } else {
      await this.driver.executeScript(`
        var el = document.querySelector('input[data-staticid="inpDateOfBirth"]');
        if (el) { el.value = ''; el.dispatchEvent(new Event('input', { bubbles: true })); }
      `);
      await this.editAndFill('inpDateOfBirth', dob);
      await this.driver.executeScript(`
        var el = document.querySelector('input[data-staticid="inpDateOfBirth"]');
        if (el) el.dispatchEvent(new Event('change', { bubbles: true }));
      `);
      await this.waitForPageLoad();
      console.log('[Form] Date of Birth filled.');
    }

    if (passport.passportIssueDate) {
      const issueDate = passport.passportIssueDate.replace(/\//g, '-');
      console.log(`[Form] Filling Passport Issue Date: "${issueDate}"...`);
      const currentIssue = await this.driver.executeScript<string>(`
        var el = document.querySelector('input[data-staticid="inpPassportIssueDate"]');
        return el ? (el.value || '').trim() : '';
      `);
      if (currentIssue && currentIssue.toUpperCase() === issueDate.toUpperCase()) {
        console.log(`[Skip] Passport Issue Date already has correct value: "${currentIssue}".`);
      } else {
        await this.driver.executeScript(`
          var el = document.querySelector('input[data-staticid="inpPassportIssueDate"]');
          if (el) { el.value = ''; el.dispatchEvent(new Event('input', { bubbles: true })); }
        `);
        await this.editAndFill('inpPassportIssueDate', issueDate);
        await this.driver.executeScript(`
          var el = document.querySelector('input[data-staticid="inpPassportIssueDate"]');
          if (el) el.dispatchEvent(new Event('change', { bubbles: true }));
        `);
        await this.waitForPageLoad();
        console.log('[Form] Passport Issue Date filled.');
      }
    } else {
      console.log('[Form] Passport Issue Date — skipped (empty).');
    }

    const expiryDate = passport.passportExpiryDate.replace(/\//g, '-');
    console.log(`[Form] Filling Passport Expiry Date: "${expiryDate}"...`);
    const currentExpiry = await this.driver.executeScript<string>(`
      var el = document.querySelector('input[data-staticid="inpPassportExpiryDate"]');
      return el ? (el.value || '').trim() : '';
    `);
    if (currentExpiry && currentExpiry.toUpperCase() === expiryDate.toUpperCase()) {
      console.log(`[Skip] Passport Expiry Date already has correct value: "${currentExpiry}".`);
    } else {
      await this.driver.executeScript(`
        var el = document.querySelector('input[data-staticid="inpPassportExpiryDate"]');
        if (el) { el.value = ''; el.dispatchEvent(new Event('input', { bubbles: true })); }
      `);
      await this.editAndFill('inpPassportExpiryDate', expiryDate);
      await this.driver.executeScript(`
        var el = document.querySelector('input[data-staticid="inpPassportExpiryDate"]');
        if (el) el.dispatchEvent(new Event('change', { bubbles: true }));
      `);
      await this.waitForPageLoad();
      console.log('[Form] Passport Expiry Date filled.');
    }

    // ── TEXT FIELDS (last — least likely to trigger AJAX) ──

    // Birth Place EN
    const birthPlace = /^[A-Z]{3}$/.test(passport.birthPlaceEN.trim())
      ? GdrfaPortalPage.mrzCodeToCountryName(passport.birthPlaceEN.trim())
      : passport.birthPlaceEN;
    console.log(`[Form] Filling Birth Place EN: "${birthPlace}"...`);
    const bpFilled = await this.editAndFill('inpApplicantBirthPlaceEn', birthPlace);
    if (bpFilled) {
      await this.driver.executeScript(`if (window.translateInputText) translateInputText('inpApplicantBirthPlaceEn');`);
      await this.sleep(300);
      console.log('[Form] Birth Place EN filled + translated.');
    }

    if (passport.passportPlaceOfIssueEN) {
      const placeOfIssue = /^[A-Z]{3}$/.test(passport.passportPlaceOfIssueEN.trim())
        ? GdrfaPortalPage.mrzCodeToCountryName(passport.passportPlaceOfIssueEN.trim())
        : passport.passportPlaceOfIssueEN;
      console.log(`[Form] Filling Place of Issue EN: "${placeOfIssue}"...`);
      const poiFilled = await this.editAndFill('inpPassportPlaceIssueEn', placeOfIssue);
      if (poiFilled) {
        await this.driver.executeScript(`if (window.translateInputText) translateInputText('inpPassportPlaceIssueEn');`);
        await this.sleep(300);
        console.log('[Form] Place of Issue EN filled + translated.');
      }
    } else {
      console.log('[Form] Place of Issue EN — skipped (empty).');
    }
  }

  // ── Applicant detail fields ────────────────────────────────────────────────

  private async fillApplicantDetails(applicant: ApplicantDetails, passportIssueCountry?: string): Promise<void> {
    // Is Inside UAE checkbox
    if (applicant.isInsideUAE) {
      console.log('[Form] Checking Is Inside UAE...');
      await this.driver.executeScript(`
        var cb = document.querySelector('input[data-staticid="chkIsInsideUAE"]');
        if (cb && !cb.checked) {
          cb.classList.remove('ReadOnly');
          cb.checked = true;
          cb.dispatchEvent(new MouseEvent('click', { bubbles: true }));
        }
      `);
      await this.waitForPageLoad();
      console.log('[Form] Is Inside UAE checked.');
    }

    // ── DROPDOWNS FIRST (they trigger AJAX that can wipe text fields) ──

    // Marital Status
    if (applicant.maritalStatus) {
      console.log(`[Form] Setting Marital Status: "${applicant.maritalStatus}"...`);
      const msResult = await this.selectByLabel('Marital Status', applicant.maritalStatus);
      if (msResult.skipped) {
        console.log(`[Skip] Marital Status already set: "${msResult.matched}".`);
      } else if (msResult.found) {
        console.log(`[Form] Marital Status set: "${msResult.matched}".`);
      } else {
        console.warn(`[Form] Marital Status not found for: "${applicant.maritalStatus}".`);
      }
    }

    // Religion — always set to "Unknown"
    {
      const religionValue = 'Unknown';
      console.log(`[Form] Setting Religion: "${religionValue}"...`);
      const rResult = await this.selectByLabel('Religion', religionValue);
      if (rResult.skipped) {
        console.log(`[Skip] Religion already set: "${rResult.matched}".`);
      } else if (rResult.found) {
        await this.waitForPageLoad();
        console.log(`[Form] Religion set: "${rResult.matched}".`);
      } else {
        console.warn(`[Form] Religion not found for: "${religionValue}".`);
      }
    }

    // Faith — always set to "Unknown"
    {
      const faithValue = 'Unknown';
      console.log(`[Form] Setting Faith: "${faithValue}"...`);

      const currentFaith = await this.driver.executeScript<string>(`
        var sel = document.querySelector('select[data-staticid="cmbApplicantFaith"]');
        if (!sel) return '';
        return sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].text.trim() : '';
      `);
      if (currentFaith && !currentFaith.includes('Select') &&
          currentFaith.toUpperCase().includes(faithValue.toUpperCase())) {
        console.log(`[Skip] Faith already set: "${currentFaith}".`);
      } else {
        // Wait until the Faith dropdown has been populated
        await this.waitForCondition(async () => {
          return this.driver.executeScript<boolean>(`
            var sel = document.querySelector('select[data-staticid="cmbApplicantFaith"]');
            return sel ? sel.options.length > 1 : false;
          `);
        }, 10000).catch(() => console.warn('[Form] Faith dropdown did not populate in time.'));

        // Try programmatic selection first (most reliable)
        const fResult = await this.selectByStaticId('cmbApplicantFaith', faithValue);
        if (fResult.found) {
          console.log(`[Form] Faith set: "${fResult.matched}".`);
        } else {
          console.warn(`[Form] Faith not found for: "${faithValue}".`);
        }
      }
    }

    // Education
    if (applicant.education) {
      console.log(`[Form] Setting Education: "${applicant.education}"...`);
      const eResult = await this.selectByLabel('Education', applicant.education);
      if (eResult.skipped) {
        console.log(`[Skip] Education already set: "${eResult.matched}".`);
      } else if (eResult.found) {
        console.log(`[Form] Education set: "${eResult.matched}".`);
      } else {
        console.warn(`[Form] Education not found for: "${applicant.education}".`);
      }
    }

    // First Language — always English
    {
      console.log('[Form] Setting First Language → English...');
      const flResult = await this.clickSelect2ByLabel('First Language', 'English');
      if (!flResult) {
        const fallback = await this.selectByLabel('First Language', 'English');
        if (fallback.found) {
          console.log(`[Form] First Language set (fallback): "${fallback.matched}".`);
        } else {
          console.warn('[Form] First Language: could not set — will retry in validation pass.');
        }
      }
    }

    // Profession — hardcoded to SALES EXECUTIVE
    {
      console.log('[Form] Setting Profession → SALES EXECUTIVE...');
      const profSet = await this.driver.executeScript<boolean>(`
        // Check if already set
        var hidden = document.querySelector('input[id*="wtProfession"][type="hidden"], input[id*="Profession"][type="hidden"]');
        if (hidden && hidden.value && hidden.value.trim() !== '') {
          return true; // already set
        }

        // Find the search input for profession (various ID patterns)
        var searchInput = document.querySelector('input[id*="wtProfessionSerch"]')
          || document.querySelector('input[id*="ProfessionSerch"]')
          || document.querySelector('input[id*="Profession"][type="text"]')
          || document.querySelector('input[id*="profession"][type="text"]');

        // Also find the hidden input that stores the actual value
        if (!hidden) {
          hidden = document.querySelector('input[id*="Profession"][type="hidden"]')
            || document.querySelector('input[id*="profession"][type="hidden"]');
        }

        // Strategy 1: Set the hidden value directly
        if (hidden) {
          hidden.value = 'SALES EXECUTIVE';
          hidden.dispatchEvent(new Event('change', { bubbles: true }));
        }

        // Strategy 2: Set the visible search input text
        if (searchInput) {
          // Remove ReadOnly
          searchInput.classList.remove('ReadOnly');
          var p = searchInput.parentElement;
          for (var i = 0; i < 5 && p; i++) { p.classList.remove('ReadOnly'); p = p.parentElement; }
          searchInput.removeAttribute('disabled');
          searchInput.removeAttribute('readonly');

          // Set the value using native setter
          var setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
          setter.call(searchInput, 'SALES EXECUTIVE');
          searchInput.dispatchEvent(new Event('input', { bubbles: true }));
          searchInput.dispatchEvent(new Event('change', { bubbles: true }));
        }

        return !!(hidden || searchInput);
      `);

      if (profSet) {
        await this.sleep(300);
        console.log('[Form] Profession → SALES EXECUTIVE.');
      } else {
        console.warn('[Form] Profession: no input found on page.');
      }
    }

    // Coming From Country (AJAX Select2) — fallback to Passport Issue Country
    const comingFrom = applicant.comingFromCountry || passportIssueCountry || '';
    if (comingFrom) {
      const cfcName = GdrfaPortalPage.mrzCodeToCountryName(comingFrom);
      console.log(`[Form] Setting Coming From Country: "${cfcName}"...`);
      const cfcResult = await this.selectByAjaxSelect2('ComingFromCountry', cfcName);
      if (cfcResult) {
        console.log(`[Form] Coming From Country set: "${cfcResult}".`);
      } else {
        // Fallback: try selectByLabel
        console.warn(`[Form] AJAX Select2 failed — trying label-based fallback...`);
        const lblResult = await this.selectByLabel('Coming From Country', cfcName);
        if (lblResult.found) {
          console.log(`[Form] Coming From Country set (label): "${lblResult.matched}".`);
        } else {
          console.warn(`[Form] Coming From Country not found for: "${cfcName}".`);
        }
      }
    } else {
      console.log('[Form] Coming From Country — skipped (empty).');
    }

    // ── TEXT FIELDS (after all dropdowns, so AJAX won't wipe them) ──

    // Mother Name EN
    if (applicant.motherNameEN) {
      console.log(`[Form] Filling Mother Name EN: "${applicant.motherNameEN}"...`);
      const motherFilled = await this.editAndFill('inpMotherNameEn', applicant.motherNameEN);
      if (motherFilled) {
        await this.driver.executeScript(`if (window.translateInputText) translateInputText('inpMotherNameEn');`);
        console.log('[Form] Mother Name EN filled + translated.');
      }
    }
  }

  // ── Contact detail fields ─────────────────────────────────────────────────

  private async fillContactDetails(contact: ContactDetails): Promise<void> {
    // ── DROPDOWNS FIRST (they trigger AJAX postbacks that can wipe text fields) ──

    if (contact.preferredSMSLanguage) {
      console.log(`[Form] Setting Preferred SMS Language: "${contact.preferredSMSLanguage}"...`);
      const langResult = await this.selectByLabel('Preferred SMS Language', contact.preferredSMSLanguage);
      if (langResult.skipped) {
        console.log(`[Skip] Preferred SMS Language already set: "${langResult.matched}".`);
      } else if (langResult.found) {
        console.log(`[Form] Preferred SMS Language set: "${langResult.matched}".`);
      } else {
        console.warn(`[Form] Preferred SMS Language not found for: "${contact.preferredSMSLanguage}".`);
      }
    }

    // ── Address Inside UAE (dropdowns) ──────────────────────────────────────

    if (contact.uaeEmirate) {
      console.log(`[Form] Setting Emirate: "${contact.uaeEmirate}"...`);
      await this.closeOpenSelect2Dropdowns();
      const emResult = await this.selectByStaticId('cmbAddressInsideEmiratesId', contact.uaeEmirate);
      if (emResult.skipped) {
        console.log(`[Skip] Emirate already set: "${emResult.matched}".`);
      } else if (emResult.found) {
        await this.waitForPageLoad();
        await this.sleep(500);
        console.log(`[Form] Emirate set: "${emResult.matched}".`);
      } else {
        // Fallback: physical Select2 click
        console.log(`[Form] Emirate programmatic failed — trying Select2 click...`);
        const clickResult = await this.clickSelect2ByLabel('Emirates', contact.uaeEmirate);
        if (clickResult) {
          await this.waitForPageLoad();
          console.log(`[Form] Emirate set via Select2 click.`);
        } else {
          console.warn(`[Form] Emirate not found for: "${contact.uaeEmirate}".`);
        }
      }
    }

    if (contact.uaeCity) {
      console.log(`[Form] Setting City: "${contact.uaeCity}"...`);
      await this.closeOpenSelect2Dropdowns();
      await this.waitForCondition(async () => {
        return this.driver.executeScript<boolean>(`
          var sel = document.querySelector('select[data-staticid="cmbAddressInsideCityId"]');
          return sel ? sel.options.length > 1 : false;
        `);
      }, 15000).catch(() => console.warn('[Form] City dropdown did not populate in time.'));
      await this.sleep(500);
      const cityResult = await this.selectByStaticId('cmbAddressInsideCityId', contact.uaeCity, 3);
      if (cityResult.skipped) {
        console.log(`[Skip] City already set: "${cityResult.matched}".`);
      } else if (cityResult.found) {
        await this.waitForPageLoad();
        await this.sleep(500);
        console.log(`[Form] City set: "${cityResult.matched}".`);
      } else {
        // Fallback: physical Select2 click
        console.log(`[Form] City programmatic failed — trying Select2 click...`);
        const clickResult = await this.clickSelect2ByLabel('City', contact.uaeCity);
        if (clickResult) {
          await this.waitForPageLoad();
          console.log(`[Form] City set via Select2 click.`);
        } else {
          console.warn(`[Form] City not found for: "${contact.uaeCity}".`);
        }
      }
    }

    // Area dropdown (depends on City AJAX, so placed after City but before text fields)
    if (contact.uaeArea) {
      console.log(`[Form] Setting Area: "${contact.uaeArea}"...`);
      // Wait longer for Area — it's a cascading dropdown that depends on City AJAX
      await this.waitForCondition(async () => {
        return this.driver.executeScript<boolean>(`
          var sel = document.querySelector('select[data-staticid="cmbAddressInsideAreaId"]');
          return sel ? sel.options.length > 1 : false;
        `);
      }, 15000).catch(() => console.warn('[Form] Area dropdown did not populate in time.'));
      await this.sleep(500);

      // First attempt: programmatic selection (with built-in retries)
      const areaResult = await this.selectByStaticId('cmbAddressInsideAreaId', contact.uaeArea, 3);
      if (areaResult.skipped) {
        console.log(`[Skip] Area already set: "${areaResult.matched}".`);
      } else if (areaResult.found) {
        await this.waitForPageLoad();
        console.log(`[Form] Area set: "${areaResult.matched}".`);
      } else {
        // Fallback: use physical Select2 click interaction
        console.log(`[Form] Area programmatic failed — trying physical Select2 click...`);
        const clickResult = await this.clickSelect2ByLabel('Area', contact.uaeArea);
        if (clickResult) {
          console.log(`[Form] Area set via Select2 click.`);
        } else {
          console.warn(`[Form] Area not found for: "${contact.uaeArea}".`);
        }
      }
    }

    // Outside Country dropdown (before outside text fields)
    if (contact.outsideCountry) {
      const countryName = GdrfaPortalPage.mrzCodeToCountryName(contact.outsideCountry);
      console.log(`[Form] Setting Outside Country: "${countryName}"...`);
      const ocResult = await this.selectByStaticId('cmbApplicantOutsideCountry', countryName);
      if (ocResult.skipped) {
        console.log(`[Skip] Outside Country already set: "${ocResult.matched}".`);
      } else if (ocResult.found) {
        await this.waitForPageLoad();
        console.log(`[Form] Outside Country set: "${ocResult.matched}".`);
      } else {
        console.warn(`[Form] Outside Country not found for: "${countryName}".`);
      }
    }

    // ── TEXT FIELDS (after ALL dropdowns, so AJAX won't wipe them) ──────────

    if (contact.email) {
      console.log(`[Form] Filling Email: "${contact.email}"...`);
      await this.editAndFill('inpEmail', contact.email);
    }

    if (contact.mobileNumber) {
      console.log(`[Form] Filling Mobile Number: "${contact.mobileNumber}"...`);
      await this.editAndFill('inpMobileNumber', contact.mobileNumber);
    }

    if (contact.approvalEmailCopy) {
      console.log(`[Form] Filling Approval Email Copy: "${contact.approvalEmailCopy}"...`);
      await this.editAndFill('inpApprovalEmailCopy', contact.approvalEmailCopy);
    }

    if (contact.uaeStreet) {
      console.log(`[Form] Filling Street: "${contact.uaeStreet}"...`);
      await this.editAndFill('inpAddressInsideStreet2', contact.uaeStreet);
    }

    if (contact.uaeBuilding) {
      console.log(`[Form] Filling Building/Villa: "${contact.uaeBuilding}"...`);
      await this.editAndFill('inpAddressInsideBuilding', contact.uaeBuilding);
    }

    if (contact.uaeFloor) {
      console.log(`[Form] Filling Floor: "${contact.uaeFloor}"...`);
      await this.editAndFill('inpFloorNo', contact.uaeFloor);
    }

    if (contact.uaeFlat) {
      console.log(`[Form] Filling Flat/Villa no.: "${contact.uaeFlat}"...`);
      await this.editAndFill('inpFlatNo', contact.uaeFlat);
    }

    // ── Address Outside UAE (text fields — dropdown already set above) ──────

    if (contact.outsideMobile) {
      console.log(`[Form] Filling Outside Mobile: "${contact.outsideMobile}"...`);
      await this.editAndFill('inpAddressOutsideMobileNumber', contact.outsideMobile);
    }

    if (contact.outsideCity) {
      console.log(`[Form] Filling Outside City: "${contact.outsideCity}"...`);
      await this.editAndFill('inpAddressOutsideCity', contact.outsideCity);
    }

    if (contact.outsideAddress) {
      console.log(`[Form] Filling Outside Address: "${contact.outsideAddress}"...`);
      await this.editAndFill('inpAddressOutsideAddress1', contact.outsideAddress);
    }
  }

  // ── Retry Faith selection (before Continue) ────────────────────────────────

  private async retryFaithSelection(faith: string): Promise<void> {
    const needsRetry = await this.driver.executeScript<boolean>(`
      var sel = document.querySelector('select[data-staticid="cmbApplicantFaith"]');
      return sel ? sel.value === '' || (sel.options[sel.selectedIndex] && sel.options[sel.selectedIndex].text.indexOf('Select') >= 0) : false;
    `);

    if (!needsRetry) {
      console.log('[Form] Faith already set — no retry needed.');
      return;
    }

    console.log(`[Form] Faith still unset — retrying selection: "${faith}"...`);

    // Scroll to Faith area
    await this.driver.executeScript(`
      var sel = document.querySelector('select[data-staticid="cmbApplicantFaith"]');
      if (sel) sel.scrollIntoView({ block: 'center', behavior: 'instant' });
    `);

    // Remove ReadOnly from the Select2 container
    await this.driver.executeScript(`
      var sel = document.querySelector('select[data-staticid="cmbApplicantFaith"]');
      if (!sel) return;
      var container = sel.previousElementSibling;
      if (container && container.classList.contains('select2-container')) {
        container.classList.remove('ReadOnly');
        container.querySelectorAll('.ReadOnly').forEach(function(el) { el.classList.remove('ReadOnly'); });
      }
    `);

    // Click Select2 choice to open dropdown
    await this.driver.executeScript(`
      var sel = document.querySelector('[data-staticid="cmbApplicantFaith"]');
      if (!sel) return;
      var prev = sel.previousElementSibling;
      if (prev && prev.classList.contains('select2-container')) {
        var choice = prev.querySelector('.select2-choice');
        if (choice) choice.click();
      }
    `);
    await this.sleep(500);

    // Remove ReadOnly from the drop panel
    await this.driver.executeScript(`
      var drop = document.querySelector('.select2-drop-active');
      if (drop) {
        drop.classList.remove('ReadOnly');
        drop.querySelectorAll('.ReadOnly').forEach(function(el) { el.classList.remove('ReadOnly'); });
      }
    `);
    await this.sleep(500);

    // Click the matching option
    const matchedText = await this.driver.executeScript<string>(`
      var faith = arguments[0];
      var results = document.querySelectorAll('.select2-drop-active .select2-results li');
      for (var i = 0; i < results.length; i++) {
        if (results[i].textContent.toUpperCase().indexOf(faith.toUpperCase()) >= 0) {
          results[i].click();
          return results[i].textContent.trim();
        }
      }
      return '';
    `, faith);

    if (matchedText) {
      await this.waitForPageLoad();
      console.log(`[Form] Faith retry set: "${matchedText}".`);
    } else {
      // Last resort: programmatic
      const fResult = await this.selectByStaticId('cmbApplicantFaith', faith);
      if (fResult.found) {
        console.log(`[Form] Faith retry set (programmatic): "${fResult.matched}".`);
      } else {
        console.warn(`[Form] Faith retry failed for: "${faith}".`);
      }
    }
  }

  // ── Pre-Continue validation ────────────────────────────────────────────────

  private async validateAndRetryRequiredFields(app: VisaApplication): Promise<void> {
    const MAX_RETRIES = 2;

    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
      console.log(`[Validate] ── Pass ${attempt}/${MAX_RETRIES}: Checking required fields... ──`);

      const emptyFields = await this.getEmptyRequiredFields(app);

      if (emptyFields.length === 0) {
        console.log('[Validate] All required fields are filled.');
        return;
      }

      console.log(`[Validate] ${emptyFields.length} field(s) need retry: ${emptyFields.map(f => f.name).join(', ')}`);

      for (const field of emptyFields) {
        console.log(`[Validate] Retrying "${field.name}"...`);
        try {
          await field.retry();
          // Wait for AJAX to settle between retries (dropdown selections can trigger postbacks)
          await this.waitForPageLoad();
          await this.sleep(500);
          console.log(`[Validate] "${field.name}" — retried.`);
        } catch (e) {
          console.warn(`[Validate] "${field.name}" — retry failed: ${e}`);
        }
      }
    }

    // Final check
    const remaining = await this.getEmptyRequiredFields(app);
    if (remaining.length > 0) {
      console.warn(`[Validate] WARNING: ${remaining.length} field(s) still empty after retries: ${remaining.map(f => f.name).join(', ')}`);
    } else {
      console.log('[Validate] All required fields are filled after retries.');
    }
  }

  private async getEmptyRequiredFields(
    app: VisaApplication,
  ): Promise<Array<{ name: string; retry: () => Promise<unknown> }>> {
    const empty: Array<{ name: string; retry: () => Promise<unknown> }> = [];

    const inputHasValue = async (staticId: string): Promise<boolean> => {
      const val = await this.driver.executeScript<string>(`
        var el = document.querySelector('input[data-staticid="' + arguments[0] + '"]');
        return el ? (el.value || '').trim() : '';
      `, staticId);
      return val.length > 0;
    };

    const selectHasValue = async (staticId: string): Promise<boolean> => {
      const val = await this.driver.executeScript<string>(`
        var sel = document.querySelector('select[data-staticid="' + arguments[0] + '"]');
        if (!sel) return '';
        return sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].text.trim() : '';
      `, staticId);
      return val.length > 0 && !val.includes('Select');
    };

    const selectByLabelHasValue = async (labelText: string): Promise<boolean> => {
      const val = await this.driver.executeScript<string>(`
        var label = arguments[0];
        var labels = Array.from(document.querySelectorAll('label'));
        var lbl = labels.find(function(l) { return l.textContent && l.textContent.trim() === label; });
        if (!lbl) return '';
        var forId = lbl.getAttribute('for') || (lbl.id ? lbl.id.replace(/lbl/i, '') : '');
        var sel = document.querySelector('select[id*="' + forId + '"]')
          || (lbl.closest('.ThemeGrid_Width6') ? lbl.closest('.ThemeGrid_Width6').parentElement.querySelector('select') : null);
        if (!sel) return '';
        return sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].text.trim() : '';
      `, labelText);
      return val.length > 0 && !val.includes('Select');
    };

    // Visit Reason — always Tourism
    if (!await selectByLabelHasValue('Visit Reason')) {
      empty.push({ name: 'Visit Reason', retry: () => this.setVisitReason() });
    }

    // Passport Names
    if (app.passport.firstName && !await inputHasValue('inpFirsttNameEn')) {
      empty.push({ name: 'First Name', retry: () => this.editAndFill('inpFirsttNameEn', app.passport.firstName) });
    }
    if (app.passport.lastName && !await inputHasValue('inpLastNameEn')) {
      empty.push({ name: 'Last Name', retry: () => this.editAndFill('inpLastNameEn', app.passport.lastName) });
    }

    // Passport Details — dates need / → - transformation to match portal format
    if (app.passport.dateOfBirth && !await inputHasValue('inpDateOfBirth')) {
      const dob = app.passport.dateOfBirth.replace(/\//g, '-');
      empty.push({ name: 'Date of Birth', retry: () => this.editAndFill('inpDateOfBirth', dob) });
    }
    if (app.passport.birthPlaceEN && !await inputHasValue('inpApplicantBirthPlaceEn')) {
      const bp = /^[A-Z]{3}$/.test(app.passport.birthPlaceEN.trim())
        ? GdrfaPortalPage.mrzCodeToCountryName(app.passport.birthPlaceEN.trim())
        : app.passport.birthPlaceEN;
      empty.push({ name: 'Birth Place', retry: () => this.editAndFill('inpApplicantBirthPlaceEn', bp) });
    }
    if (app.passport.passportIssueDate && !await inputHasValue('inpPassportIssueDate')) {
      const issueDate = app.passport.passportIssueDate.replace(/\//g, '-');
      empty.push({ name: 'Passport Issue Date', retry: () => this.editAndFill('inpPassportIssueDate', issueDate) });
    }
    if (app.passport.passportExpiryDate && !await inputHasValue('inpPassportExpiryDate')) {
      const expiryDate = app.passport.passportExpiryDate.replace(/\//g, '-');
      empty.push({ name: 'Passport Expiry Date', retry: () => this.editAndFill('inpPassportExpiryDate', expiryDate) });
    }
    if (app.passport.passportPlaceOfIssueEN && !await inputHasValue('inpPassportPlaceIssueEn')) {
      const poi = /^[A-Z]{3}$/.test(app.passport.passportPlaceOfIssueEN.trim())
        ? GdrfaPortalPage.mrzCodeToCountryName(app.passport.passportPlaceOfIssueEN.trim())
        : app.passport.passportPlaceOfIssueEN;
      empty.push({ name: 'Passport Place of Issue', retry: () => this.editAndFill('inpPassportPlaceIssueEn', poi) });
    }

    // Passport Selects — use mrzCodeToCountryName for country fields
    if (app.passport.gender && !await selectByLabelHasValue('Gender')) {
      empty.push({ name: 'Gender', retry: async () => { await this.selectByLabel('Gender', app.passport.gender); } });
    }
    if (app.passport.birthCountry && !await selectByLabelHasValue('Birth Country')) {
      const bcName = GdrfaPortalPage.mrzCodeToCountryName(app.passport.birthCountry);
      empty.push({ name: 'Birth Country', retry: async () => { await this.selectByLabel('Birth Country', bcName); } });
    }
    if (app.passport.passportIssueCountry && !await selectByLabelHasValue('Passport Issue Country')) {
      const icName = GdrfaPortalPage.mrzCodeToCountryName(app.passport.passportIssueCountry);
      empty.push({ name: 'Passport Issue Country', retry: async () => { await this.selectByLabel('Passport Issue Country', icName); } });
    }

    // Applicant Details
    if (app.applicant.motherNameEN && !await inputHasValue('inpMotherNameEn')) {
      empty.push({ name: 'Mother Name', retry: () => this.editAndFill('inpMotherNameEn', app.applicant.motherNameEN) });
    }
    if (app.applicant.maritalStatus && !await selectByLabelHasValue('Marital Status')) {
      empty.push({ name: 'Marital Status', retry: async () => { await this.selectByLabel('Marital Status', app.applicant.maritalStatus); } });
    }
    if (app.applicant.religion && !await selectByLabelHasValue('Religion')) {
      empty.push({ name: 'Religion', retry: async () => { await this.selectByLabel('Religion', app.applicant.religion); } });
    }
    if (app.applicant.education && !await selectByLabelHasValue('Education')) {
      empty.push({ name: 'Education', retry: async () => { await this.selectByLabel('Education', app.applicant.education); } });
    }
    if (app.applicant.faith && !await selectHasValue('cmbApplicantFaith')) {
      empty.push({ name: 'Faith', retry: () => this.retryFaithSelection(app.applicant.faith!) });
    }
    if (!await selectByLabelHasValue('First Language')) {
      empty.push({ name: 'First Language', retry: async () => {
        const ok = await this.clickSelect2ByLabel('First Language', 'English');
        if (!ok) await this.selectByLabel('First Language', 'English');
      }});
    }

    // Contact Details
    if (app.contact.email && !await inputHasValue('inpEmail')) {
      empty.push({ name: 'Email', retry: () => this.editAndFill('inpEmail', app.contact.email) });
    }
    if (app.contact.mobileNumber && !await inputHasValue('inpMobileNumber')) {
      empty.push({ name: 'Mobile Number', retry: () => this.editAndFill('inpMobileNumber', app.contact.mobileNumber) });
    }
    if (app.contact.preferredSMSLanguage && !await selectByLabelHasValue('Preferred SMS Language')) {
      empty.push({ name: 'SMS Language', retry: async () => { await this.selectByLabel('Preferred SMS Language', app.contact.preferredSMSLanguage); } });
    }

    // Address Inside UAE
    if (app.contact.uaeEmirate && !await selectHasValue('cmbAddressInsideEmiratesId')) {
      empty.push({ name: 'UAE Emirate', retry: async () => { await this.selectByStaticId('cmbAddressInsideEmiratesId', app.contact.uaeEmirate!); } });
    }
    if (app.contact.uaeCity && !await selectHasValue('cmbAddressInsideCityId')) {
      empty.push({ name: 'UAE City', retry: async () => { await this.selectByStaticId('cmbAddressInsideCityId', app.contact.uaeCity!); } });
    }
    if (app.contact.uaeStreet && !await inputHasValue('inpAddressInsideStreet2')) {
      empty.push({ name: 'UAE Street', retry: () => this.editAndFill('inpAddressInsideStreet2', app.contact.uaeStreet!) });
    }
    if (app.contact.uaeBuilding && !await inputHasValue('inpAddressInsideBuilding')) {
      empty.push({ name: 'UAE Building', retry: () => this.editAndFill('inpAddressInsideBuilding', app.contact.uaeBuilding!) });
    }

    // Address Outside UAE
    if (app.contact.outsideCountry && !await selectHasValue('cmbApplicantOutsideCountry')) {
      empty.push({ name: 'Outside Country', retry: async () => { await this.selectByStaticId('cmbApplicantOutsideCountry', app.contact.outsideCountry!); } });
    }
    if (app.contact.outsideMobile && !await inputHasValue('inpAddressOutsideMobileNumber')) {
      empty.push({ name: 'Outside Mobile', retry: () => this.editAndFill('inpAddressOutsideMobileNumber', app.contact.outsideMobile!) });
    }
    if (app.contact.outsideCity && !await inputHasValue('inpAddressOutsideCity')) {
      empty.push({ name: 'Outside City', retry: () => this.editAndFill('inpAddressOutsideCity', app.contact.outsideCity!) });
    }
    if (app.contact.outsideAddress && !await inputHasValue('inpAddressOutsideAddress1')) {
      empty.push({ name: 'Outside Address', retry: () => this.editAndFill('inpAddressOutsideAddress1', app.contact.outsideAddress!) });
    }

    return empty;
  }

  // ── Continue button ────────────────────────────────────────────────────────

  private async waitForLoaderToDisappear(timeout = 60000): Promise<void> {
    const start = Date.now();
    const deadline = start + timeout;
    let forceHideAttempted = false;
    console.log('[Loader] Checking for loading overlays...');

    while (Date.now() < deadline) {
      const elapsed = Date.now() - start;

      const status = await this.driver.executeScript<string>(`
        var overlayVisible = false;
        var ajaxActive = false;

        // Check OutSystems AJAX loader
        var loaders = document.querySelectorAll('div.Feedback_AjaxWait, .os-internal-Feedback_AjaxWait');
        for (var i = 0; i < loaders.length; i++) {
          var s = window.getComputedStyle(loaders[i]);
          if (s.display !== 'none' && s.visibility !== 'hidden' && s.opacity !== '0') overlayVisible = true;
        }

        // Check jQuery/osjs AJAX
        if (typeof jQuery !== 'undefined' && jQuery.active > 0) ajaxActive = true;
        if (typeof osjs !== 'undefined' && osjs.active > 0) ajaxActive = true;

        if (!overlayVisible) return 'CLEAR';
        if (ajaxActive) return 'AJAX_ACTIVE';
        return 'OVERLAY_STUCK';
      `);

      if (status === 'CLEAR') {
        console.log(`[Loader] Overlay gone after ${(elapsed / 1000).toFixed(1)}s.`);
        return;
      }

      // After 30s: if AJAX is done but overlay stuck, force-hide it
      if (elapsed > 30000 && status === 'OVERLAY_STUCK' && !forceHideAttempted) {
        forceHideAttempted = true;
        console.log('[Loader] AJAX done but overlay still visible after 30s — force-hiding...');
        await this.driver.executeScript(`
          // Force-hide all OutSystems loading overlays
          var loaders = document.querySelectorAll('div.Feedback_AjaxWait, .os-internal-Feedback_AjaxWait');
          for (var i = 0; i < loaders.length; i++) {
            loaders[i].style.display = 'none';
            loaders[i].style.visibility = 'hidden';
          }
          // Also hide any full-page overlay with the loading message
          document.querySelectorAll('div').forEach(function(el) {
            if ((el.textContent || '').indexOf('Please wait while content is loading') >= 0 && el.children.length < 5) {
              el.style.display = 'none';
              // Also try the parent (sometimes it's wrapped)
              if (el.parentElement) el.parentElement.style.display = 'none';
            }
          });
        `);
        await this.sleep(1000);
        continue;
      }

      // After 45s: regardless of AJAX state, force-hide everything
      if (elapsed > 45000 && !forceHideAttempted) {
        forceHideAttempted = true;
        console.warn('[Loader] Loader stuck for 45s — force-hiding all overlays...');
        await this.driver.executeScript(`
          document.querySelectorAll('div.Feedback_AjaxWait, .os-internal-Feedback_AjaxWait').forEach(function(el) {
            el.style.display = 'none';
          });
          // Kill any pending AJAX
          if (typeof jQuery !== 'undefined') {
            try { jQuery.active = 0; } catch(e) {}
          }
          if (typeof osjs !== 'undefined') {
            try { osjs.active = 0; } catch(e) {}
          }
        `);
        await this.sleep(1000);
        continue;
      }

      await this.sleep(1000);
    }
    console.warn(`[Loader] Timeout after ${timeout / 1000}s — proceeding anyway.`);
  }

  private async clickContinue(): Promise<void> {
    console.log('[Form] Clicking Continue...');
    await this.clearBlockingOverlays();
    await this.closeOpenSelect2Dropdowns();
    await this.sleep(300);
    const btn = await this.driver.wait(
      until.elementLocated(By.css('input[staticid="SmartChannels_EntryPermitNewTourism_btnContinue"]')),
      10000,
      'Continue button not found within 10s',
    );
    await this.driver.executeScript('arguments[0].scrollIntoView({block:"center"});', btn);
    await this.sleep(300);
    try {
      await btn.click();
    } catch {
      console.log('[Form] Continue native click blocked — using JS click...');
      await this.driver.executeScript('arguments[0].click();', btn);
    }
    console.log('[Form] Continue clicked — waiting for popup or next page...');

    // Brief pause so the portal has time to show the popup or the loader
    await this.sleep(2000);

    const popupFrame = await this.findPopupFrame(15000);

    if (popupFrame) {
      console.log('[Form] Existing application popup detected (in iframe).');
      await this.handleExistingApplicationPopup();
    } else {
      console.log('[Form] No popup — waiting for loader to clear...');
    }

    // Wait for the "Please wait while content is loading..." overlay to disappear
    await this.waitForLoaderToDisappear(90000);
    await this.waitForPageLoad();
    await this.sleep(2000);

    // Verify the attachments section actually loaded (document upload cards present)
    const ready = await this.driver.executeScript<boolean>(`
      // Check for file upload inputs (sign of attachments tab)
      var inputs = document.querySelectorAll('input[type="file"]');
      if (inputs.length > 0) return true;
      // Check for document card containers
      var cards = document.querySelectorAll('[id*="cntDocList"], [id*="CardGray"]');
      return cards.length > 0;
    `);

    if (ready) {
      console.log('[Form] Attachments section loaded — document upload cards found.');
    } else {
      console.log('[Form] Attachments section not yet visible — waiting longer...');
      // Wait up to 30 more seconds for file inputs to appear
      for (let i = 0; i < 30; i++) {
        const found = await this.driver.executeScript<boolean>(`
          return document.querySelectorAll('input[type="file"]').length > 0;
        `);
        if (found) {
          console.log('[Form] Document upload cards appeared.');
          break;
        }
        await this.sleep(1000);
      }
    }
  }

  private async findPopupFrame(timeout: number): Promise<boolean> {
    const deadline = Date.now() + timeout;

    while (Date.now() < deadline) {
      // Check all iframes for the popup
      const iframes = await this.driver.findElements(By.css('iframe'));
      for (const iframe of iframes) {
        try {
          await this.driver.switchTo().frame(iframe);
          const popups = await this.driver.findElements(By.css('div.MainPopup'));
          if (popups.length > 0) {
            const frameUrl = await this.driver.getCurrentUrl();
            console.log(`[Form] Found popup in frame: ${frameUrl}`);
            await this.driver.switchTo().defaultContent();
            return true;
          }
          await this.driver.switchTo().defaultContent();
        } catch {
          try { await this.driver.switchTo().defaultContent(); } catch {}
        }
      }
      await this.sleep(250);
    }
    return false;
  }

  private async handleExistingApplicationPopup(): Promise<void> {
    console.log('[Form] On "Existing Application Details" popup!');

    // Find the iframe with the popup and switch to it
    const iframes = await this.driver.findElements(By.css('iframe'));
    for (const iframe of iframes) {
      try {
        await this.driver.switchTo().frame(iframe);
        const popups = await this.driver.findElements(By.css('div.MainPopup'));
        if (popups.length > 0) break;
        await this.driver.switchTo().defaultContent();
      } catch {
        try { await this.driver.switchTo().defaultContent(); } catch {}
      }
    }

    // Extract data
    const values = await this.driver.executeScript<string[]>(`
      var spans = document.querySelectorAll('div.MainPopup span.Bold');
      var vals = [];
      for (var i = 0; i < spans.length; i++) {
        vals.push(spans[i].textContent.replace(/\\u00a0/g, ' ').trim());
      }
      return vals;
    `);

    console.log('[Form] ── Existing Application Details ──');
    console.log(`  Application Number : ${values[0] ?? ''}`);
    console.log(`  Applicant Name     : ${values[1] ?? ''}`);
    console.log(`  Nationality        : ${values[2] ?? ''}`);
    console.log(`  Passport No        : ${values[3] ?? ''}`);
    console.log(`  Sponsor Name       : ${values[4] ?? ''}`);
    console.log(`  Created Date       : ${values[5] ?? ''}`);
    console.log('[Form] ────────────────────────────────────');

    // Click Continue inside the iframe
    try {
      const closeBtn = await this.driver.findElement(
        By.css('input[staticid="CommonTh_ExistingApplicationConfirmationPopUp_btnCancel"]')
      );
      // Find the Continue button (sibling of Cancel)
      const continueBtn = await this.driver.executeScript<WebElement>(`
        var cancel = arguments[0];
        var sibling = cancel.nextElementSibling;
        while (sibling) {
          if (sibling.tagName === 'INPUT' && sibling.value === 'Continue') return sibling;
          sibling = sibling.nextElementSibling;
        }
        return null;
      `, closeBtn);

      if (continueBtn) {
        await continueBtn.click();
        console.log('[Form] Clicked Continue on popup.');
      }
    } catch (e) {
      console.warn('[Form] Could not click Continue on popup:', e);
    }

    // Switch back to main content
    await this.driver.switchTo().defaultContent();

    await this.waitForPageLoad();
    await this.waitForLoaderToDisappear();
    console.log('[Form] Popup dismissed.');
  }

  // ── SmartInput helpers ─────────────────────────────────────────────────────

  private async editAndFill(staticId: string, value: string): Promise<boolean> {
    // Check if the field already has the correct value
    const currentVal = await this.driver.executeScript<string>(`
      var el = document.querySelector('input[data-staticid="' + arguments[0] + '"]');
      return el ? (el.value || '').trim() : '';
    `, staticId);
    if (currentVal !== '' && currentVal.toUpperCase() === value.trim().toUpperCase()) {
      console.log(`[Skip] "${staticId}" already has correct value: "${currentVal}".`);
      return false;
    }

    // Unlock the field: remove ReadOnly, click pencil, enable input — all via JS
    await this.driver.executeScript(`
      var id = arguments[0];
      var input = document.querySelector('input[data-staticid="' + id + '"]');
      if (!input) return;
      input.scrollIntoView({ block: 'center', behavior: 'instant' });

      // Remove ReadOnly from the input and all ancestors
      input.classList.remove('ReadOnly');
      input.removeAttribute('readonly');
      input.removeAttribute('disabled');
      var el = input.parentElement;
      while (el && el !== document.body) {
        el.classList.remove('ReadOnly');
        var pencil = el.querySelector('.FormEditPencil');
        if (pencil) { pencil.click(); break; }
        el = el.parentElement;
      }
    `, staticId);
    await this.sleep(300);

    // Try Selenium sendKeys first (triggers proper key events)
    let filled = false;
    try {
      const field = await this.driver.findElement(By.css(`input[data-staticid="${staticId}"]`));
      const isVisible = await field.isDisplayed().catch(() => false);
      const isReadonly = await field.getAttribute('readonly');
      if (isVisible && !isReadonly) {
        await field.clear();
        await field.sendKeys(value);
        filled = true;
      }
    } catch {}

    if (!filled) {
      // JS fallback with proper event dispatch
      await this.driver.executeScript(`
        var id = arguments[0];
        var val = arguments[1];
        var el = document.querySelector('input[data-staticid="' + id + '"]');
        if (!el) return;
        el.classList.remove('ReadOnly');
        el.removeAttribute('readonly');
        el.removeAttribute('disabled');
        // Clear first
        var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        nativeInputValueSetter.call(el, '');
        el.dispatchEvent(new Event('input', { bubbles: true }));
        // Set value using native setter (triggers React/framework change detection)
        nativeInputValueSetter.call(el, val);
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('blur', { bubbles: true }));
      `, staticId, value);
    }
    // Pause after each field to let the slow website process
    await this.sleep(600);
    return true;
  }

  private async clearArField(staticId: string): Promise<void> {
    await this.driver.executeScript(`
      var el = document.querySelector('input[data-staticid="' + arguments[0] + '"]');
      if (el) { el.value = ''; el.dispatchEvent(new Event('input', { bubbles: true })); }
    `, staticId);
  }

  private async selectByLabel(
    labelText: string,
    searchValue: string
  ): Promise<{ found: boolean; matched: string; skipped?: boolean }> {
    return this.driver.executeScript<{ found: boolean; matched: string; skipped?: boolean }>(`
      var label = arguments[0];
      var search = arguments[1];
      var lbl = Array.from(document.querySelectorAll('label')).find(function(l) {
        return !l.classList.contains('select2-offscreen') &&
               l.textContent && l.textContent.trim().toLowerCase() === label.toLowerCase();
      });
      if (!lbl || !lbl.htmlFor) return { found: false, matched: '' };

      var el = document.getElementById(lbl.htmlFor);
      var sel = (el instanceof HTMLSelectElement) ? el
        : (el && el.parentElement ? el.parentElement.querySelector('select') : null);
      if (!sel) return { found: false, matched: '' };

      var currentText = sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].text.trim() : '';
      if (currentText && currentText.indexOf('Select') < 0 &&
          (currentText.toUpperCase() === search.toUpperCase() ||
           currentText.toUpperCase().indexOf(search.toUpperCase()) >= 0)) {
        return { found: true, matched: currentText, skipped: true };
      }

      var s2 = sel.closest('.select2-container')
        || (sel.previousElementSibling && sel.previousElementSibling.classList.contains('select2-container') ? sel.previousElementSibling : null)
        || (sel.nextElementSibling && sel.nextElementSibling.classList.contains('select2-container') ? sel.nextElementSibling : null)
        || document.getElementById('s2id_' + sel.id);
      if (s2) {
        s2.classList.remove('ReadOnly');
        s2.querySelectorAll('.ReadOnly').forEach(function(el) { el.classList.remove('ReadOnly'); });
      }
      sel.removeAttribute('disabled');

      var opts = Array.from(sel.options);
      var match = opts.find(function(o) { return o.text.trim().toUpperCase() === search.toUpperCase(); })
        || opts.find(function(o) { return o.text.toUpperCase().indexOf(search.toUpperCase()) >= 0; });
      if (!match) return { found: false, matched: '' };

      sel.value = '';
      sel.value = match.value;
      sel.dispatchEvent(new Event('change', { bubbles: true }));

      var jq = window.jQuery || window.$;
      if (jq) { jq(sel).val(match.value).trigger('change'); }
      if (s2) {
        var chosen = s2.querySelector('.select2-chosen');
        if (chosen) chosen.textContent = match.text.trim();
      }

      return { found: true, matched: match.text };
    `, labelText, searchValue);
    // Give the slow website time to process the selection
    // (wait is after the executeScript returns)
  }

  /** Wrapper that adds a pause after selectByLabel for slow pages */
  private async selectByLabelSlow(
    labelText: string,
    searchValue: string
  ): Promise<{ found: boolean; matched: string; skipped?: boolean }> {
    const result = await this.selectByLabel(labelText, searchValue);
    if (result.found && !result.skipped) {
      await this.sleep(800);
    }
    return result;
  }

  private async selectByStaticId(
    staticId: string,
    searchValue: string,
    retries: number = 2
  ): Promise<{ found: boolean; matched: string; skipped?: boolean }> {
    // Close any stale Select2 dropdowns first
    await this.closeOpenSelect2Dropdowns();

    for (let attempt = 0; attempt <= retries; attempt++) {
      const result = await this.driver.executeScript<{ found: boolean; matched: string; skipped?: boolean }>(`
        var id = arguments[0];
        var search = arguments[1];
        var sel = document.querySelector('select[data-staticid="' + id + '"]');
        if (!sel) return { found: false, matched: '' };

        var currentText = sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].text.trim() : '';
        if (currentText && currentText.indexOf('Select') < 0 &&
            (currentText.toUpperCase() === search.toUpperCase() ||
             currentText.toUpperCase().indexOf(search.toUpperCase()) >= 0)) {
          return { found: true, matched: currentText, skipped: true };
        }

        var s2 = sel.closest('.select2-container')
          || (sel.previousElementSibling && sel.previousElementSibling.classList.contains('select2-container') ? sel.previousElementSibling : null)
          || (sel.nextElementSibling && sel.nextElementSibling.classList.contains('select2-container') ? sel.nextElementSibling : null)
          || document.getElementById('s2id_' + sel.id);
        if (s2) {
          s2.classList.remove('ReadOnly');
          s2.querySelectorAll('.ReadOnly').forEach(function(el) { el.classList.remove('ReadOnly'); });
          // Also unlock parent chain
          var p = s2.parentElement;
          for (var k = 0; k < 5 && p; k++) { p.classList.remove('ReadOnly'); p = p.parentElement; }
        }
        sel.removeAttribute('disabled');
        sel.removeAttribute('readonly');

        var opts = Array.from(sel.options);
        var match = opts.find(function(o) { return o.text.trim().toUpperCase() === search.toUpperCase(); })
          || opts.find(function(o) { return o.text.toUpperCase().indexOf(search.toUpperCase()) >= 0; });
        if (!match) return { found: false, matched: '', optionCount: opts.length };

        sel.value = '';
        sel.value = match.value;
        sel.dispatchEvent(new Event('change', { bubbles: true }));

        var jq = window.jQuery || window.$;
        if (jq) {
          jq(sel).val(match.value).trigger('change');
          jq(sel).trigger({ type: 'select2:select', params: { data: { id: match.value, text: match.text } } });
        }
        if (s2) {
          var chosen = s2.querySelector('.select2-chosen');
          if (chosen) chosen.textContent = match.text.trim();
        }

        return { found: true, matched: match.text };
      `, staticId, searchValue);

      if (result.found) return result;

      // If not found and we have retries left, wait for AJAX and try again
      if (attempt < retries) {
        console.log(`[Select2] "${staticId}" — option "${searchValue}" not found (attempt ${attempt + 1}), waiting for AJAX...`);
        await this.sleep(2000);
        await this.waitForPageLoad();
      }
    }

    return { found: false, matched: '' };
  }

  private async selectByAjaxSelect2(
    idFragment: string,
    searchValue: string
  ): Promise<string> {
    // Deep unlock: remove ReadOnly from ALL elements related to this Select2
    const containerId = await this.driver.executeScript<string>(`
      var frag = arguments[0];

      // Find container by ID fragment
      var containers = Array.from(document.querySelectorAll('.select2-container'));
      var match = containers.find(function(el) { return el.id.toLowerCase().indexOf(frag.toLowerCase()) >= 0; });

      // Fallback: find by nearby label text
      if (!match) {
        var labels = Array.from(document.querySelectorAll('label'));
        var lbl = labels.find(function(l) {
          return (l.textContent || '').toLowerCase().indexOf(frag.toLowerCase().replace(/([a-z])([A-Z])/g, '$1 $2').toLowerCase()) >= 0;
        });
        if (lbl) {
          var el = document.getElementById(lbl.htmlFor || '');
          if (el) {
            match = el.previousElementSibling;
            if (!match || !match.classList.contains('select2-container')) {
              match = el.closest('div')?.querySelector('.select2-container') || null;
            }
          }
        }
      }

      if (!match) return '';

      // Deep unlock: remove ReadOnly from container, all children, and parent chain
      match.classList.remove('ReadOnly');
      match.querySelectorAll('*').forEach(function(el) { el.classList.remove('ReadOnly'); });
      var parent = match.parentElement;
      while (parent && parent !== document.body) {
        parent.classList.remove('ReadOnly');
        parent = parent.parentElement;
      }

      // Also unlock the actual <select> element (usually next sibling)
      var selectEl = match.nextElementSibling;
      if (selectEl && selectEl.tagName === 'SELECT') {
        selectEl.classList.remove('ReadOnly');
        selectEl.removeAttribute('readonly');
        selectEl.removeAttribute('disabled');
      }

      return match.id;
    `, idFragment);

    if (!containerId) {
      console.warn(`[Form] AJAX Select2 container not found for fragment: "${idFragment}".`);
      return '';
    }

    // Scroll and click the Select2 choice to open the dropdown
    await this.driver.executeScript(`
      var container = document.getElementById(arguments[0]);
      if (container) {
        container.scrollIntoView({block:'center'});
        var choice = container.querySelector('.select2-choice');
        if (choice) {
          choice.classList.remove('ReadOnly');
          choice.click();
        }
      }
    `, containerId);
    await this.sleep(800);

    // Unlock the dropdown panel, click pencil, enable search input
    await this.driver.executeScript(`
      var drop = document.querySelector('.select2-drop-active');
      if (!drop) return;
      drop.classList.remove('ReadOnly');
      drop.querySelectorAll('*').forEach(function(el) { el.classList.remove('ReadOnly'); });

      var pencil = drop.querySelector('.FormEditPencil');
      if (pencil) pencil.click();

      var input = drop.querySelector('.select2-input');
      if (input) {
        input.removeAttribute('readonly');
        input.removeAttribute('disabled');
        input.classList.remove('ReadOnly');
        input.style.display = '';
        input.style.visibility = 'visible';
        input.focus();
      }
    `);
    await this.sleep(500);

    // Type search value character by character (slow pace for AJAX)
    const activeInput = await this.driver.switchTo().activeElement();
    for (const char of searchValue) {
      await activeInput.sendKeys(char);
      await this.sleep(80);
    }

    // Wait for AJAX results to appear (longer timeout for slow site)
    await this.waitForCondition(async () => {
      return this.driver.executeScript<boolean>(`
        var results = document.querySelectorAll('.select2-drop-active .select2-results li.select2-result');
        return results.length > 0;
      `);
    }, 15000).catch(() => {});
    await this.sleep(500);

    // Click the best matching result — prefer exact match over partial
    const matchedText = await this.driver.executeScript<string>(`
      var search = arguments[0];
      var results = document.querySelectorAll('.select2-drop-active .select2-results li.select2-result');
      if (results.length === 0) return '';

      // Find best match: exact > starts-with > first
      var item = null;
      for (var i = 0; i < results.length; i++) {
        var t = results[i].textContent.trim();
        // Extract text after "NNN - " prefix (e.g. "116 - CHINA" → "CHINA")
        var label = t.replace(/^\\d+\\s*-\\s*/, '').trim();
        if (label.toUpperCase() === search.toUpperCase()) { item = results[i]; break; }
      }
      if (!item) {
        for (var i = 0; i < results.length; i++) {
          var t = results[i].textContent.trim();
          var label = t.replace(/^\\d+\\s*-\\s*/, '').trim();
          if (label.toUpperCase().indexOf(search.toUpperCase()) === 0) { item = results[i]; break; }
        }
      }
      if (!item) item = results[0];

      var text = item.textContent.trim();
      // Select2 listens on mouseup after mousedown — simulate the full sequence
      item.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      item.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
      item.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
      return text;
    `, searchValue);

    if (matchedText) {
      await this.sleep(1500);
      await this.waitForPageLoad();
      console.log(`[Form] AJAX Select2 result clicked: "${matchedText}"`);
    }

    // Verify selection took effect — if dropdown is still open, try Selenium click
    const stillOpen = await this.driver.executeScript<boolean>(`
      return !!document.querySelector('.select2-drop-active');
    `);
    if (stillOpen && matchedText) {
      console.log('[Form] Dropdown still open — using Selenium click...');
      try {
        const resultEl = await this.driver.findElement(
          By.css('.select2-drop-active .select2-results li.select2-result')
        );
        await resultEl.click();
        await this.sleep(1000);
        await this.waitForPageLoad();
      } catch {
        // Force close dropdown and set value programmatically
        console.log('[Form] Selenium click failed — forcing selection via JS...');
        await this.driver.executeScript(`
          var frag = arguments[0];
          var text = arguments[1];
          // Find the underlying <select>
          var containers = Array.from(document.querySelectorAll('.select2-container'));
          var match = containers.find(function(el) { return el.id.toLowerCase().indexOf(frag.toLowerCase()) >= 0; });
          if (!match) return;
          var sel = match.nextElementSibling;
          if (!sel || sel.tagName !== 'SELECT') return;
          var opt = Array.from(sel.options).find(function(o) { return o.text.indexOf(text.split(' - ').pop() || text) >= 0; });
          if (opt) {
            sel.value = opt.value;
            sel.dispatchEvent(new Event('change', { bubbles: true }));
            var jq = window.jQuery || window.$;
            if (jq) jq(sel).val(opt.value).trigger('change');
          }
          // Close any open dropdown
          var drop = document.querySelector('.select2-drop-active');
          if (drop) drop.classList.remove('select2-drop-active');
        `, idFragment, matchedText);
        await this.sleep(1000);
      }
    }

    return matchedText;
  }

  // ── Session Keep-Alive ─────────────────────────────────────────────────────

  private startSessionKeepAlive(intervalMs = 10 * 60 * 1000): () => void {
    console.log(`[KeepAlive] Session keep-alive started (every ${intervalMs / 60000} min).`);
    const timer = setInterval(async () => {
      try {
        await this.driver.executeScript(`
          fetch(window.location.href, { method: 'HEAD', credentials: 'include' });
        `);
        console.log('[KeepAlive] Session ping sent.');
      } catch { /* page may be mid-navigation */ }
    }, intervalMs);
    return () => { clearInterval(timer); console.log('[KeepAlive] Stopped.'); };
  }

  // ── Selenium utility helpers ───────────────────────────────────────────────

  private async waitForElement(locator: By, timeout: number): Promise<WebElement> {
    return this.driver.wait(until.elementLocated(locator), timeout);
  }

  private async waitForPageLoad(timeout = 15000): Promise<void> {
    try {
      await this.driver.wait(async () => {
        const ready = await this.driver.executeScript<boolean>(`
          // Check document.readyState
          if (document.readyState !== 'complete') return false;
          // Check jQuery AJAX
          if (typeof jQuery !== 'undefined' && jQuery.active > 0) return false;
          // Check OutSystems AJAX
          if (typeof osjs !== 'undefined' && osjs.active > 0) return false;
          return true;
        `);
        return ready;
      }, timeout);
      // Brief settle time after all AJAX completes
      await this.sleep(300);
    } catch {
      // Timeout is acceptable — page may have slow-loading resources
    }
  }

  private async waitForUrlContains(fragment: string, timeout: number): Promise<void> {
    await this.driver.wait(until.urlContains(fragment), timeout);
  }

  private async waitForCondition(conditionFn: () => Promise<boolean>, timeout: number): Promise<void> {
    const deadline = Date.now() + timeout;
    while (Date.now() < deadline) {
      if (await conditionFn()) return;
      await this.sleep(250);
    }
    throw new Error(`Condition not met within ${timeout}ms`);
  }

  /** Remove all blocking overlays, modals, popups, and stuck loaders that intercept clicks. */
  private async clearBlockingOverlays(): Promise<void> {
    await this.driver.executeScript(`
      // 1. Hide OutSystems loader overlays
      document.querySelectorAll('.os-loading-overlay, .Feedback_AjaxWait, .loading-overlay').forEach(function(el) {
        el.style.display = 'none';
        el.style.visibility = 'hidden';
      });
      // 2. Remove generic modal backdrops
      document.querySelectorAll('.modal-backdrop, .popup-overlay, .overlay').forEach(function(el) {
        if (el.offsetHeight > 0) {
          el.style.display = 'none';
        }
      });
      // 3. Remove any absolutely positioned element with high z-index blocking the page
      document.querySelectorAll('[style*="z-index"]').forEach(function(el) {
        var z = parseInt(window.getComputedStyle(el).zIndex, 10);
        var pos = window.getComputedStyle(el).position;
        if (z > 1000 && (pos === 'fixed' || pos === 'absolute') && el.offsetWidth > 200 && el.offsetHeight > 200) {
          // Likely a blocking overlay — hide it
          if (!el.querySelector('input') && !el.querySelector('select') && !el.querySelector('table')) {
            el.style.display = 'none';
          }
        }
      });
    `);
  }

  /** Close any open Select2 dropdowns to prevent them from blocking clicks on other elements. */
  private async closeOpenSelect2Dropdowns(): Promise<void> {
    await this.driver.executeScript(`
      // Close active Select2 drop panels
      document.querySelectorAll('.select2-drop-active').forEach(function(drop) {
        drop.style.display = 'none';
        drop.classList.remove('select2-drop-active');
      });
      // Remove active state from containers
      document.querySelectorAll('.select2-container-active').forEach(function(c) {
        c.classList.remove('select2-container-active');
      });
      // Close any select2-drop that's still visible
      document.querySelectorAll('.select2-drop').forEach(function(d) {
        if (window.getComputedStyle(d).display !== 'none') {
          d.style.display = 'none';
        }
      });
    `);
  }

  private sleep(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  private static mrzCodeToCountryName(code: string): string {
    const map: Record<string, string> = {
      AFG: 'AFGHANISTAN',   ALB: 'ALBANIA',       DZA: 'ALGERIA',       AND: 'ANDORRA',
      AGO: 'ANGOLA',        ARG: 'ARGENTINA',      ARM: 'ARMENIA',       AUS: 'AUSTRALIA',
      AUT: 'AUSTRIA',       AZE: 'AZERBAIJAN',     BHS: 'BAHAMAS',       BHR: 'BAHRAIN',
      BGD: 'BANGLADESH',    BRB: 'BARBADOS',       BEL: 'BELGIUM',       BLZ: 'BELIZE',
      BEN: 'BENIN',         BTN: 'BHUTAN',         BOL: 'BOLIVIA',       BIH: 'BOSNIA',
      BWA: 'BOTSWANA',      BRA: 'BRAZIL',         GBR: 'BRITAIN',       BRN: 'BRUNEI',
      BGR: 'BULGARIA',      BFA: 'BURKINA FASO',   MMR: 'BURMA',         BDI: 'BURUNDI',
      CPV: 'CABO VERDE',    KHM: 'CAMBODIA',       CMR: 'CAMEROON',      CAN: 'CANADA',
      CAF: 'CENTRAL AFRICA',TCD: 'CHAD',            CHL: 'CHILE',        CHN: 'CHINA',
      COL: 'COLOMBIA',      COM: 'COMOROS',        COG: 'CONGO',         CRI: 'COSTARICA',
      HRV: 'CROATIA',       CUB: 'CUBA',           CYP: 'CYPRUS',        CZE: 'CZECH',
      DNK: 'DENMARK',       DJI: 'DJIBOUTI',       DOM: 'DOMINICAN',     ECU: 'ECUADOR',
      EGY: 'EGYPT',         SLV: 'EL SALVADOR',    ARE: 'EMIRATES',      ERI: 'ERITREN',
      EST: 'ESTONIA',       ETH: 'ETHIOPIA',       FJI: 'FIJI',          FIN: 'FINLAND',
      FRA: 'FRANCE',        GAB: 'GABON',          GMB: 'GAMBIA',        GEO: 'GEORGIA',
      DEU: 'GERMANY',       GHA: 'GHANA',          GRC: 'GREECE',        GRD: 'GRENADA',
      GTM: 'GUATAMALA',     GUY: 'GUYANA',         HTI: 'HAITI',         NLD: 'HOLLAND',
      HND: 'HONDURAS',      HKG: 'HONG KONG',      HUN: 'HUNGARY',       ISL: 'ICELAND',
      IND: 'INDIA',         IDN: 'INDONESIA',      IRN: 'IRAN',          IRQ: 'IRAQ',
      IRL: 'IRELAND',       ISR: 'ISRAEIL',        ITA: 'ITALY',         CIV: 'IVORY COAST',
      JAM: 'JAMAICA',       JPN: 'JAPAN',          JOR: 'JORDAN',        KAZ: 'KAZAKHESTAN',
      KEN: 'KENYA',         KWT: 'KUWAIT',         KGZ: 'Kyrgyzstani',   LAO: 'LAOS',
      LVA: 'LATVIA',        LBN: 'LEBANON',        LSO: 'LESOTHO',       LBR: 'LIBERIA',
      LBY: 'LIBYA',         LTU: 'LITHUANIA',      LUX: 'LUXEMBOURG',    MAC: 'MACAU',
      MDG: 'MADAGASCAR',    MWI: 'MALAWI',         MYS: 'MALAYSIA',      MDV: 'MALDIVES',
      MLI: 'MALI',          MLT: 'MALTA',          MRT: 'MAURITANIA',    MUS: 'MAURITIUS',
      MEX: 'MEXICO',        MDA: 'MOLDAVIA',       MCO: 'MONACO',        MNG: 'MONGOLIA',
      MNE: 'MONTENEGRO',    MAR: 'MOROCCO',        MOZ: 'MOZAMBIQUE',    NAM: 'NAMEBIA',
      NPL: 'NEPAL',         NZL: 'NEW ZEALAND',    NIC: 'NICARAGUA',     NER: 'NIGER',
      NGA: 'NIGERIA',       PRK: 'NORTH KOREA',    NOR: 'NORWAY',        PAK: 'PAKISTAN',
      PAN: 'PANAMA',        PNG: 'PAPUA NEW GUINEA',PRY: 'PARAGUAY',     PER: 'PERU',
      PHL: 'PHILIPPINES',   POL: 'POLAND',         PRT: 'PORTUGAL',      QAT: 'QATAR',
      ROU: 'ROMANIA',       RWA: 'ROWANDA',        RUS: 'RUSSIA',        SAU: 'SAUDI ARABIA',
      SEN: 'SENEGAL',       SRB: 'SERBIA',         SLE: 'SIERRA LEONE',  SGP: 'SINGAPORE',
      SVK: 'SLOVAKIA',      SVN: 'SLOVENIA',       SOM: 'SOMALIA',       ZAF: 'SOUTH AFRICA',
      KOR: 'SOUTH KOREA',   SSD: 'SOUTH SUDAN',    ESP: 'SPAIN',         LKA: 'SRI LANKA',
      SDN: 'SUDAN',         OMN: 'SULTANATE OF OMAN',SUR: 'SURINAME',    SWZ: 'SWAZILAND',
      SWE: 'SWEDEN',        CHE: 'SWIZERLAND',     SYR: 'SYRIA',         TWN: 'TAIWAN',
      TJK: 'TAJIKSTAN',     TZA: 'TANZANIA',       THA: 'THAILAND',      TLS: 'TIMOR LESTE',
      TGO: 'TOGO',          TON: 'TONGA',          TTO: 'TRINIDAD',      TUN: 'TUNISIA',
      TUR: 'TURKEY',        TKM: 'TURKMENISTAN',   USA: 'U S A',         UGA: 'UGANDA',
      UKR: 'UKRAINE',       URY: 'URGWAY',         UZB: 'UZBAKISTAN',    YEM: 'YEMEN',
      ZMB: 'ZAMBIA',        ZWE: 'ZIMBABWE',
    };
    const upper = code.toUpperCase().trim();

    // Direct MRZ code lookup (e.g. "IND" → "INDIA")
    if (map[upper]) return map[upper];

    // Already a country name in the map? Return as-is (e.g. "INDIA" → "INDIA")
    const allNames = Object.values(map);
    if (allNames.includes(upper)) return upper;

    // Nationality adjective → country name (e.g. "INDIAN" → "INDIA", "CHINESE" → "CHINA")
    const adjectiveMap: Record<string, string> = {
      INDIAN: 'INDIA',           CHINESE: 'CHINA',          AMERICAN: 'U S A',
      BRITISH: 'BRITAIN',        FRENCH: 'FRANCE',          GERMAN: 'GERMANY',
      JAPANESE: 'JAPAN',         KOREAN: 'KOREA',           AUSTRALIAN: 'AUSTRALIA',
      CANADIAN: 'CANADA',        RUSSIAN: 'RUSSIA',         BRAZILIAN: 'BRAZIL',
      MEXICAN: 'MEXICO',         ITALIAN: 'ITALY',          SPANISH: 'SPAIN',
      PORTUGUESE: 'PORTUGAL',    DUTCH: 'HOLLAND',          BELGIAN: 'BELGIUM',
      SWISS: 'SWITZERLAND',      SWEDISH: 'SWEDEN',         NORWEGIAN: 'NORWAY',
      DANISH: 'DENMARK',         FINNISH: 'FINLAND',        POLISH: 'POLAND',
      TURKISH: 'TURKEY',         EGYPTIAN: 'EGYPT',         SOUTH_AFRICAN: 'SOUTH AFRICA',
      NIGERIAN: 'NIGERIA',       KENYAN: 'KENYA',           GHANAIAN: 'GHANA',
      PAKISTANI: 'PAKISTAN',      BANGLADESHI: 'BANGLADESH', SRI_LANKAN: 'SRI LANKA',
      NEPALESE: 'NEPAL',         NEPALI: 'NEPAL',           THAI: 'THAILAND',
      FILIPINO: 'PHILLIPINE',    INDONESIAN: 'INDONESIA',   MALAYSIAN: 'MALAYSIA',
      SINGAPOREAN: 'SINGAPORE',  VIETNAMESE: 'VIETNAM',     IRAQI: 'IRAQ',
      IRANIAN: 'IRAN',           SAUDI: 'SAUDIA',           EMIRATI: 'EMIRATES',
      KUWAITI: 'KUWAIT',         QATARI: 'QATAR',           OMANI: 'OMAN',
      BAHRAINI: 'BAHRAIN',       JORDANIAN: 'JORDAN',       LEBANESE: 'LEBANON',
      SYRIAN: 'SYRIA',           YEMENI: 'YEMEN',           MOROCCAN: 'MOROCCO',
      TUNISIAN: 'TUNISIA',       ALGERIAN: 'ALGERIA',       ETHIOPIAN: 'ETHIOPIA',
      SOMALI: 'SOMALIA',         SUDANESE: 'SUDAN',         AFGHAN: 'AFGHANISTAN',
      UZBEK: 'UZBAKISTAN',       UKRAINIAN: 'UKRAINE',      COLOMBIAN: 'COLOMBIA',
      PERUVIAN: 'PERU',          ARGENTINIAN: 'ARGENTINA',  CHILEAN: 'CHILE',
      ECUADORIAN: 'ECUADOR',     VENEZUELAN: 'VENEZUELA',   CUBAN: 'CUBA',
      JAMAICAN: 'JAMAICA',       TRINIDADIAN: 'TRINIDAD',   IRISH: 'IRELAND',
      SCOTTISH: 'BRITAIN',       WELSH: 'BRITAIN',          GREEK: 'GREECE',
      CZECH: 'CZECH',            HUNGARIAN: 'HUNGARY',      ROMANIAN: 'ROMANIA',
      BULGARIAN: 'BULGARIA',     CROATIAN: 'CROATIA',       SERBIAN: 'SERBIA',
      BOSNIAN: 'BOSNIA',         ALBANIAN: 'ALBANIA',       GEORGIAN: 'GEORGIA',
      ARMENIAN: 'ARMENIA',       AZERBAIJANI: 'AZERBAIJAN', KAZAKH: 'KAZAKHESTAN',
      CAMBODIAN: 'CAMBODIA',     BURMESE: 'BURMA',          LAOTIAN: 'LAOS',
      SOUTH_KOREAN: 'SOUTH KOREA', NORTH_KOREAN: 'NORTH KOREA',
    };
    // Also handle underscores/spaces: "SOUTH AFRICAN" → "SOUTH_AFRICAN"
    const normalised = upper.replace(/\s+/g, '_');
    if (adjectiveMap[upper]) return adjectiveMap[upper];
    if (adjectiveMap[normalised]) return adjectiveMap[normalised];

    // Fuzzy: strip common nationality suffixes and check if that's a country
    const stripped = upper
      .replace(/(IAN|AN|ESE|ISH|I|ER)$/, '')
      .trim();
    const fuzzyMatch = allNames.find(n => n.startsWith(stripped) || n.includes(stripped));
    if (fuzzyMatch) return fuzzyMatch;

    return code;
  }
}

// ─── Convenience exports (test file imports these directly) ───────────────────

/** Thin wrapper so tests can call fillApplicationForm(driver, app) unchanged. Returns Application Number. */
export async function fillApplicationForm(driver: WebDriver, application: VisaApplication): Promise<string> {
  return new GdrfaPortalPage(driver).fillApplicationForm(application);
}

export async function verifySession(driver: WebDriver): Promise<void> {
  await new GdrfaPortalPage(driver).verifySession();
}
