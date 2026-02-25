import { WebDriver, By, until, WebElement, Key } from 'selenium-webdriver';
import * as path from 'path';
import * as fs from 'fs';
import {
  VisaApplication,
  PassportDetails,
  ApplicantDetails,
  ContactDetails,
  ApplicationDocuments,
} from '../types/application-data';

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
      await this.sleep(1000);

      // ── Section 2: Passport Details ──
      console.log('\n[Flow] ── Section: Passport Details ──');
      await this.setPassportType(application.passport.passportType);
      await this.sleep(800);
      await this.enterPassportNumber(application.passport.passportNumber);
      await this.sleep(800);
      await this.setNationality(application.passport.currentNationality);
      await this.sleep(1000);
      await this.setPreviousNationality(
        application.passport.previousNationality ?? application.passport.currentNationality
      );
      await this.sleep(1000);
      await this.clickSearchDataAndWait();
      await this.sleep(2000);

      // ── Section 3: Passport Names ──
      console.log('\n[Flow] ── Section: Passport Names ──');
      await this.fillPassportNames(application.passport);
      await this.sleep(1500);

      // ── Section 4: Passport Dates & Details ──
      console.log('\n[Flow] ── Section: Passport Dates ──');
      await this.fillPassportDetails(application.passport);
      await this.sleep(1500);

      // ── Section 5: Applicant Details ──
      console.log('\n[Flow] ── Section: Applicant Details ──');
      await this.fillApplicantDetails(application.applicant, application.passport.passportIssueCountry);
      await this.sleep(1500);

      // ── Section 6: Contact Details ──
      console.log('\n[Flow] ── Section: Contact Details ──');
      await this.fillContactDetails(application.contact);
      await this.sleep(1500);

      // Retry Faith selection before continuing (dropdown can reset after other fields)
      await this.retryFaithSelection('Unknown');
      await this.sleep(1000);

      // Validate all required fields before clicking Continue — retry any empty ones
      await this.validateAndRetryRequiredFields(application);
      await this.sleep(1000);

      await this.clickContinue();

      // Upload documents on the Attachments tab
      console.log('\n[Flow] ── Section: Document Upload ──');
      await this.uploadDocuments(application.documents);

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

  async uploadDocuments(docs: ApplicationDocuments): Promise<void> {
    console.log('[Upload] Starting document upload...');
    await this.waitForPageLoad();
    await this.sleep(3000);

    // Discover what upload elements exist on the page
    const pageInfo = await this.driver.executeScript<{
      fileInputs: Array<{ id: string; name: string; attrs: Record<string, string> }>;
      uploadZones: string[];
      labels: string[];
    }>(`
      var result = { fileInputs: [], uploadZones: [], labels: [] };
      // All file inputs
      document.querySelectorAll('input[type="file"]').forEach(function(el) {
        var attrs = {};
        for (var i = 0; i < el.attributes.length; i++) {
          attrs[el.attributes[i].name] = el.attributes[i].value;
        }
        result.fileInputs.push({ id: el.id || '', name: el.name || '', attrs: attrs });
      });
      // Upload zones / dropzones
      document.querySelectorAll('[class*="upload"], [class*="dropzone"], [class*="Upload"], [class*="Dropzone"]').forEach(function(el) {
        result.uploadZones.push(el.className + ' | ' + (el.textContent || '').substring(0, 80).trim());
      });
      // Labels that mention document types
      document.querySelectorAll('label, span.upload-label, .document-label, td, th').forEach(function(el) {
        var t = (el.textContent || '').trim();
        if (t && (t.indexOf('Passport') >= 0 || t.indexOf('Photo') >= 0 || t.indexOf('Hotel') >= 0
            || t.indexOf('ticket') >= 0 || t.indexOf('Ticket') >= 0 || t.indexOf('reservation') >= 0
            || t.indexOf('Upload') >= 0 || t.indexOf('Attach') >= 0 || t.indexOf('Document') >= 0)) {
          result.labels.push(t.substring(0, 120));
        }
      });
      return result;
    `);

    console.log(`[Upload] File inputs found: ${pageInfo.fileInputs.length}`);
    for (const fi of pageInfo.fileInputs) {
      console.log(`  - id="${fi.id}" name="${fi.name}" attrs=${JSON.stringify(fi.attrs)}`);
    }
    if (pageInfo.uploadZones.length > 0) {
      console.log(`[Upload] Upload zones: ${pageInfo.uploadZones.length}`);
      for (const uz of pageInfo.uploadZones.slice(0, 10)) {
        console.log(`  - ${uz}`);
      }
    }
    if (pageInfo.labels.length > 0) {
      console.log(`[Upload] Document labels found:`);
      for (const lbl of pageInfo.labels.slice(0, 15)) {
        console.log(`  - ${lbl}`);
      }
    }

    // Map document labels to file paths — primary slots first, then Others fallback
    const slots: Array<{ label: string; keywords: string[]; file: string }> = [
      { label: 'Sponsored Passport page 1',                keywords: ['sponsored', 'passport page', 'passport page 1'], file: docs.sponsoredPassportPage1 },
      { label: 'Passport External Cover Page',             keywords: ['cover', 'external cover'],                       file: docs.passportExternalCoverPage },
      { label: 'Personal Photo',                           keywords: ['photo', 'personal photo'],                       file: docs.personalPhoto },
      { label: 'Hotel reservation/Place of stay - Page 1', keywords: ['hotel', 'reservation', 'place of stay'],         file: docs.hotelReservationPage1 },
      { label: 'Return air ticket - Page 1',               keywords: ['ticket', 'air ticket', 'return'],                file: docs.returnAirTicketPage1 },
    ];

    // Build a list of available data-document-types on the page
    const availableTypes = await this.driver.executeScript<string[]>(`
      return Array.from(document.querySelectorAll('input[type="file"][data-document-type]'))
        .map(function(el) { return el.getAttribute('data-document-type') || ''; });
    `);
    console.log('[Upload] Available upload slot types:', availableTypes);

    // If hotel/flight slots don't exist, try uploading to "Others Page 1" and "Others Page 2"
    const hotelSlotExists = availableTypes.some(t => t.toLowerCase().includes('hotel') || t.toLowerCase().includes('reservation'));
    const ticketSlotExists = availableTypes.some(t => t.toLowerCase().includes('ticket') || t.toLowerCase().includes('air ticket'));

    if (!hotelSlotExists && docs.hotelReservationPage1) {
      console.log('[Upload] No hotel slot found — will try "Others Page 1"');
      // Replace the hotel slot to target "Others Page 1"
      const hotelIdx = slots.findIndex(s => s.label.includes('Hotel'));
      if (hotelIdx >= 0) {
        slots[hotelIdx] = { label: 'Others Page 1', keywords: ['others page 1', 'others'], file: docs.hotelReservationPage1 };
      }
    }
    if (!ticketSlotExists && docs.returnAirTicketPage1) {
      console.log('[Upload] No ticket slot found — will try "Others Page 2"');
      const ticketIdx = slots.findIndex(s => s.label.includes('Return'));
      if (ticketIdx >= 0) {
        slots[ticketIdx] = { label: 'Others Page 2', keywords: ['others page 2', 'others'], file: docs.returnAirTicketPage1 };
      }
    }

    if (pageInfo.fileInputs.length === 0) {
      // No file inputs found — try to find them inside iframes
      console.log('[Upload] No file inputs in main page — checking iframes...');
      const iframes = await this.driver.findElements(By.css('iframe'));
      for (let i = 0; i < iframes.length; i++) {
        try {
          await this.driver.switchTo().frame(iframes[i]);
          const iframeInputs = await this.driver.findElements(By.css('input[type="file"]'));
          if (iframeInputs.length > 0) {
            console.log(`[Upload] Found ${iframeInputs.length} file input(s) in iframe ${i}`);
            // Upload from within this iframe
            await this.uploadInCurrentContext(slots);
            await this.driver.switchTo().defaultContent();
            console.log('[Upload] Iframe uploads complete.');
            return;
          }
          await this.driver.switchTo().defaultContent();
        } catch {
          try { await this.driver.switchTo().defaultContent(); } catch {}
        }
      }

      // Still no inputs — try clicking upload/attach buttons first
      console.log('[Upload] No file inputs in iframes either. Looking for upload buttons...');
      await this.driver.executeScript(`
        var btns = document.querySelectorAll('a, button, input[type="button"]');
        for (var i = 0; i < btns.length; i++) {
          var t = (btns[i].textContent || btns[i].value || '').trim().toLowerCase();
          if (t.indexOf('upload') >= 0 || t.indexOf('attach') >= 0 || t.indexOf('browse') >= 0) {
            console.log('Clicking: ' + t);
          }
        }
      `);
      console.warn('[Upload] Could not find any file upload mechanism on the page.');
      // Take a diagnostic screenshot
      try {
        const screenshot = await this.driver.takeScreenshot();
        const ssPath = path.resolve('test-results', 'upload-page-debug.png');
        fs.mkdirSync(path.dirname(ssPath), { recursive: true });
        fs.writeFileSync(ssPath, screenshot, 'base64');
        console.log(`[Upload] Debug screenshot saved: ${ssPath}`);
      } catch {}
      return;
    }

    // Upload using the discovered file inputs
    await this.uploadInCurrentContext(slots);

    // Take a screenshot to verify uploads
    try {
      const screenshot = await this.driver.takeScreenshot();
      const ssPath = path.resolve('test-results', 'after-upload.png');
      fs.mkdirSync(path.dirname(ssPath), { recursive: true });
      fs.writeFileSync(ssPath, screenshot, 'base64');
      console.log(`[Upload] Post-upload screenshot: ${ssPath}`);
    } catch {}

    console.log('[Upload] All documents uploaded. Waiting before Continue...');
    await this.sleep(5000);

    // Look for Continue button — OutSystems uses various patterns
    await this.clickContinueButton();
  }

  /**
   * Finds and clicks the Continue button on the upload/attachments page.
   * Tries multiple selector strategies for OutSystems portal.
   */
  private async clickContinueButton(): Promise<void> {
    const strategies = [
      // OutSystems common button patterns
      `input[value="Continue"]`,
      `input[value="continue"]`,
      `a[id*="Continue"], a[id*="continue"]`,
      `input[id*="Continue"], input[id*="continue"]`,
      `button[id*="Continue"], button[id*="continue"]`,
      // Static ID patterns
      `[data-staticid*="Continue"], [data-staticid*="continue"]`,
      `[staticid*="Continue"], [staticid*="continue"]`,
      // By text content via XPath is not CSS — handle separately
    ];

    for (const selector of strategies) {
      try {
        const btn = await this.driver.findElement(By.css(selector));
        const displayed = await btn.isDisplayed().catch(() => false);
        if (displayed) {
          await this.driver.executeScript('arguments[0].scrollIntoView({block:"center"});', btn);
          await this.sleep(500);
          await btn.click();
          await this.waitForPageLoad();
          await this.waitForLoaderToDisappear();
          console.log(`[Upload] Continue clicked (${selector}).`);
          return;
        }
      } catch {}
    }

    // Try finding by text content
    try {
      const continueBtn = await this.driver.executeScript<WebElement | null>(`
        var allBtns = document.querySelectorAll('input[type="submit"], input[type="button"], button, a.btn, a.button, a[class*="btn"]');
        for (var i = 0; i < allBtns.length; i++) {
          var text = (allBtns[i].textContent || allBtns[i].value || '').trim().toLowerCase();
          if (text === 'continue' || text === 'next' || text === 'submit') {
            return allBtns[i];
          }
        }
        return null;
      `);
      if (continueBtn) {
        await this.driver.executeScript('arguments[0].scrollIntoView({block:"center"});', continueBtn);
        await this.sleep(500);
        await (continueBtn as WebElement).click();
        await this.waitForPageLoad();
        await this.waitForLoaderToDisappear();
        console.log('[Upload] Continue clicked (text match).');
        return;
      }
    } catch {}

    console.log('[Upload] No Continue button found — page may have auto-advanced.');
    // Save debug screenshot
    try {
      const screenshot = await this.driver.takeScreenshot();
      const ssPath = path.resolve('test-results', 'continue-btn-debug.png');
      fs.mkdirSync(path.dirname(ssPath), { recursive: true });
      fs.writeFileSync(ssPath, screenshot, 'base64');
      console.log(`[Upload] Continue button debug screenshot: ${ssPath}`);
    } catch {}
  }

  /**
   * Uploads a single file to a file input element.
   * Makes the input visible, sends the file path, dispatches change event,
   * and waits for upload to complete.
   */
  private async uploadFileToInput(input: WebElement, filePath: string, label: string): Promise<boolean> {
    try {
      // 1. Make the file input visible and interactable via JS
      await this.driver.executeScript(`
        var el = arguments[0];
        el.style.display = 'block';
        el.style.visibility = 'visible';
        el.style.opacity = '1';
        el.style.position = 'absolute';
        el.style.width = '200px';
        el.style.height = '40px';
        el.style.zIndex = '99999';
        el.style.left = '0px';
        el.style.top = '0px';
        el.removeAttribute('disabled');
        el.removeAttribute('readonly');
        // Also unhide any parent that may be hiding it
        var parent = el.parentElement;
        for (var i = 0; i < 5 && parent; i++) {
          if (window.getComputedStyle(parent).display === 'none' ||
              window.getComputedStyle(parent).visibility === 'hidden' ||
              window.getComputedStyle(parent).overflow === 'hidden') {
            parent.style.display = 'block';
            parent.style.visibility = 'visible';
            parent.style.overflow = 'visible';
            parent.style.height = 'auto';
            parent.style.width = 'auto';
          }
          parent = parent.parentElement;
        }
      `, input);
      await this.sleep(500);

      // 2. Scroll into view
      await this.driver.executeScript('arguments[0].scrollIntoView({block:"center"});', input);
      await this.sleep(300);

      // 3. Send the file path
      await input.sendKeys(filePath);
      console.log(`[Upload] sendKeys done for "${label}": ${path.basename(filePath)}`);

      // 4. Dispatch change and input events to trigger the framework's upload handler
      await this.driver.executeScript(`
        var el = arguments[0];
        var changeEvent = new Event('change', { bubbles: true });
        el.dispatchEvent(changeEvent);
        var inputEvent = new Event('input', { bubbles: true });
        el.dispatchEvent(inputEvent);
      `, input);

      // 5. Wait for upload to process — check for progress indicators
      console.log(`[Upload] Waiting for "${label}" to upload...`);
      await this.sleep(3000);

      // Wait up to 30 seconds for upload completion
      for (let attempt = 0; attempt < 10; attempt++) {
        const uploadState = await this.driver.executeScript<{
          hasProgress: boolean;
          hasSpinner: boolean;
          hasFileName: boolean;
          fileNameText: string;
        }>(`
          var result = { hasProgress: false, hasSpinner: false, hasFileName: false, fileNameText: '' };
          // Check for progress bars / spinners
          var spinners = document.querySelectorAll('.loading, .spinner, .progress, [class*="loading"], [class*="progress"], [class*="uploading"]');
          result.hasSpinner = spinners.length > 0;
          for (var i = 0; i < spinners.length; i++) {
            if (window.getComputedStyle(spinners[i]).display !== 'none') {
              result.hasProgress = true;
              break;
            }
          }
          // Check if file name appears near the input (upload complete indicator)
          var el = arguments[0];
          var container = el.closest('.upload-container, .file-upload, [class*="upload"], [class*="attach"], div') || el.parentElement;
          if (container) {
            var text = container.textContent || '';
            var basename = arguments[1];
            if (text.indexOf(basename) >= 0 || text.indexOf('.jpg') >= 0 || text.indexOf('.pdf') >= 0 || text.indexOf('.png') >= 0) {
              result.hasFileName = true;
              result.fileNameText = text.substring(0, 150).trim();
            }
          }
          return result;
        `, input, path.basename(filePath));

        if (uploadState.hasFileName) {
          console.log(`[Upload] "${label}" upload confirmed — file name visible.`);
          break;
        }
        if (uploadState.hasProgress) {
          console.log(`[Upload] "${label}" still uploading... (attempt ${attempt + 1})`);
          await this.sleep(3000);
        } else {
          // No progress indicator and no file name — upload may have completed quickly
          // or the site uses a different indicator. Wait a bit more and move on.
          await this.sleep(2000);
          break;
        }
      }

      await this.waitForPageLoad();
      await this.sleep(2000);
      return true;
    } catch (err) {
      console.error(`[Upload] Error uploading "${label}":`, err);
      return false;
    }
  }

  /**
   * Uploads documents using file inputs found in the current browsing context.
   * Matches slots to file inputs by: data-document-type, nearby label text, or input order.
   */
  private async uploadInCurrentContext(
    slots: Array<{ label: string; keywords: string[]; file: string }>
  ): Promise<void> {
    const fileInputs = await this.driver.findElements(By.css('input[type="file"]'));
    console.log(`[Upload] ${fileInputs.length} file input(s) in current context.`);

    // Try matching by data-document-type attribute first
    const hasDataAttr = await this.driver.executeScript<boolean>(`
      return document.querySelectorAll('input[type="file"][data-document-type]').length > 0;
    `);

    if (hasDataAttr) {
      // Use data-document-type matching
      for (const slot of slots) {
        if (!slot.file) continue;
        const filePath = path.resolve(slot.file);
        if (!fs.existsSync(filePath)) { console.warn(`[Upload] File missing: ${filePath}`); continue; }
        const input = await this.findFileInputByAttr(slot.label);
        if (input) {
          await this.uploadFileToInput(input, filePath, slot.label);
        } else {
          // Try partial / fuzzy matching on data-document-type
          const fuzzyInput = await this.findFileInputByFuzzyAttr(slot.label, slot.keywords);
          if (fuzzyInput) {
            await this.uploadFileToInput(fuzzyInput, filePath, slot.label);
          } else {
            console.warn(`[Upload] No slot for: "${slot.label}"`);
          }
        }
      }
    } else {
      // Match by nearby label/text or sequential order
      const inputContexts = await this.driver.executeScript<string[]>(`
        return Array.from(document.querySelectorAll('input[type="file"]')).map(function(el) {
          var container = el.closest('div, td, tr, li, section') || el.parentElement;
          var text = container ? container.textContent.trim().substring(0, 200) : '';
          return text;
        });
      `);

      console.log('[Upload] File input contexts:');
      for (let i = 0; i < inputContexts.length; i++) {
        console.log(`  [${i}] ${inputContexts[i].substring(0, 100)}`);
      }

      // Match each slot to an input by keyword matching
      const usedIndices = new Set<number>();
      for (const slot of slots) {
        if (!slot.file) continue;
        const filePath = path.resolve(slot.file);
        if (!fs.existsSync(filePath)) { console.warn(`[Upload] File missing: ${filePath}`); continue; }

        let bestIdx = -1;
        for (let i = 0; i < inputContexts.length; i++) {
          if (usedIndices.has(i)) continue;
          const ctx = inputContexts[i].toLowerCase();
          if (slot.keywords.some(kw => ctx.includes(kw.toLowerCase()))) {
            bestIdx = i;
            break;
          }
        }

        if (bestIdx >= 0) {
          usedIndices.add(bestIdx);
          const input = fileInputs[bestIdx];
          await this.uploadFileToInput(input, filePath, slot.label);
        } else {
          console.warn(`[Upload] No matching input for: "${slot.label}"`);
        }
      }
    }
  }

  private async findFileInputByAttr(label: string): Promise<WebElement | null> {
    try {
      return await this.driver.findElement(By.css(`input[type="file"][data-document-type="${label}"]`));
    } catch {
      const matchIdx = await this.driver.executeScript<number>(`
        var inputs = Array.from(document.querySelectorAll('input[type="file"][data-document-type]'));
        var target = arguments[0].toLowerCase();
        return inputs.findIndex(function(el) { return (el.getAttribute('data-document-type') || '').toLowerCase() === target; });
      `, label);
      if (matchIdx < 0) return null;
      const inputs = await this.driver.findElements(By.css('input[type="file"][data-document-type]'));
      return inputs[matchIdx] || null;
    }
  }

  /**
   * Fuzzy match: finds a file input whose data-document-type contains any of the keywords.
   */
  private async findFileInputByFuzzyAttr(label: string, keywords: string[]): Promise<WebElement | null> {
    const matchIdx = await this.driver.executeScript<number>(`
      var inputs = Array.from(document.querySelectorAll('input[type="file"][data-document-type]'));
      var keywords = arguments[0];
      for (var k = 0; k < keywords.length; k++) {
        var kw = keywords[k].toLowerCase();
        for (var i = 0; i < inputs.length; i++) {
          var dtype = (inputs[i].getAttribute('data-document-type') || '').toLowerCase();
          if (dtype.indexOf(kw) >= 0) return i;
        }
      }
      return -1;
    `, keywords);
    if (matchIdx < 0) return null;
    const inputs = await this.driver.findElements(By.css('input[type="file"][data-document-type]'));
    return inputs[matchIdx] || null;
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
    await this.dismissPromoPopup();

    // Click "Existing Applications"
    const existAppSel = '#EmaratechSG_Theme_wtwbLayoutEmaratech_block_wtMainContent_wtwbDashboard_wtCntExistApp';
    let clicked = false;
    try {
      const el = await this.waitForElement(By.css(existAppSel), 15000);
      await el.click();
      clicked = true;
    } catch {
      // Fallback: click by link text
      try {
        const link = await this.driver.findElement(By.partialLinkText('Existing Applications'));
        await link.click();
        clicked = true;
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

    await this.waitForPageLoad();
    const currentUrl = await this.driver.getCurrentUrl();
    console.log('[Nav] Form page loaded. URL:', currentUrl);
  }

  private async waitForPageSettle(): Promise<void> {
    await this.waitForPageLoad();
  }

  // ── Passport header (Passport Type → Nationality → Search Data) ────────────

  private async setVisitReason(): Promise<void> {
    console.log('[Form] Selecting Visit Reason → Tourism...');
    const set = await this.driver.executeScript<boolean>(`
      // Try multiple selectors for Visit Reason
      var sel = document.querySelector('select[data-staticid="cmbVisitReason"]')
        || document.querySelector('select[id*="VisitReason"]')
        || document.querySelector('select[id*="visitReason"]')
        || document.querySelector('select[name*="VisitReason"]');

      // Fallback: find by label text
      if (!sel) {
        var labels = Array.from(document.querySelectorAll('label'));
        var lbl = labels.find(function(l) {
          var t = (l.textContent || '').trim().toLowerCase();
          return t.indexOf('visit') >= 0 && t.indexOf('reason') >= 0;
        });
        if (lbl && lbl.htmlFor) {
          var el = document.getElementById(lbl.htmlFor);
          sel = (el instanceof HTMLSelectElement) ? el
            : (el && el.parentElement ? el.parentElement.querySelector('select') : null);
        }
      }

      // Last fallback: find the first select in "Visit Details" section
      if (!sel) {
        var sections = document.querySelectorAll('a[href*="Visit"], span, div, td');
        for (var i = 0; i < sections.length; i++) {
          if ((sections[i].textContent || '').trim() === 'Visit Details') {
            var container = sections[i].closest('div, section, table');
            if (container) {
              var selects = container.querySelectorAll('select');
              if (selects.length > 0) { sel = selects[0]; break; }
            }
          }
        }
      }

      if (!sel) {
        // Debug: log all selects on page
        var allSelects = document.querySelectorAll('select');
        console.log('All selects on page: ' + allSelects.length);
        allSelects.forEach(function(s, i) {
          var opts = Array.from(s.options).map(function(o) { return o.text.trim(); }).slice(0, 5).join(', ');
          console.log('  select[' + i + '] id=' + s.id + ' opts: ' + opts);
        });
        return false;
      }

      var currentText = sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].text.trim() : '';
      if (currentText.toUpperCase().indexOf('TOURISM') >= 0) return true;

      // Find the Tourism option
      var tourismOpt = Array.from(sel.options).find(function(o) {
        return o.text.trim().toUpperCase().indexOf('TOURISM') >= 0;
      });
      if (tourismOpt) {
        sel.value = tourismOpt.value;
      } else {
        // Try value '1' as fallback
        sel.value = '1';
      }
      sel.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    `);
    if (set) {
      await this.waitForPageLoad();
      await this.sleep(1000);
      console.log('[Form] Visit Reason → Tourism.');
    } else {
      console.warn('[Form] Visit Reason select not found.');
    }
  }

  private async setPassportType(passportType: string): Promise<void> {
    console.log(`[Form] Setting Passport Type: "${passportType}"...`);
    const set = await this.driver.executeScript<boolean>(`
      var type = arguments[0];

      // Try specific selectors first
      var sel = document.querySelector('select[data-staticid*="PassportType"]')
        || document.querySelector('select[id*="PassportType"]')
        || document.querySelector('select[id*="passportType"]')
        || document.querySelector('select[name*="PassportType"]');

      // Fallback: find by label "Passport Type"
      if (!sel) {
        var labels = Array.from(document.querySelectorAll('label'));
        var lbl = labels.find(function(l) {
          var t = (l.textContent || '').replace(/\\*/, '').trim().toLowerCase();
          return t === 'passport type';
        });
        if (lbl && lbl.htmlFor) {
          var el = document.getElementById(lbl.htmlFor);
          sel = (el instanceof HTMLSelectElement) ? el
            : (el && el.parentElement ? el.parentElement.querySelector('select') : null);
        }
      }

      // Last fallback: find any select that has "Normal" as an option
      if (!sel) {
        var allSels = Array.from(document.querySelectorAll('select'));
        sel = allSels.find(function(s) {
          return Array.from(s.options).some(function(o) {
            return o.text.trim().toLowerCase() === 'normal';
          });
        }) || null;
      }

      if (!sel) return false;

      var match = Array.from(sel.options).find(function(o) {
        return o.text.trim().toLowerCase() === type.toLowerCase();
      });
      if (!match) return false;
      sel.value = match.value;
      sel.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    `, passportType);
    if (set) {
      await this.waitForPageLoad();
      await this.sleep(500);
      console.log('[Form] Passport Type set: ' + passportType);
    } else {
      console.warn('[Form] Passport Type select not found.');
    }
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
    console.log(`[Form] Setting Nationality: "${name}"...`);
    const result = await this.driver.executeScript<{ found: boolean; matched: string }>(`
      var search = arguments[0];
      var sel = Array.from(document.querySelectorAll('select')).find(function(s) {
        return Array.from(s.options).some(function(o) { return /^\\d+ - /.test(o.text.trim()); });
      });
      if (!sel) return { found: false, matched: '' };
      sel.removeAttribute('disabled');
      var match = Array.from(sel.options).find(function(o) { return o.text.toUpperCase().indexOf(search.toUpperCase()) >= 0; });
      if (!match) return { found: false, matched: '' };
      sel.value = match.value;
      sel.dispatchEvent(new Event('change', { bubbles: true }));
      return { found: true, matched: match.text };
    `, name);
    if (result.found) {
      await this.waitForPageLoad();
      console.log(`[Form] Nationality set: "${result.matched}".`);
    } else {
      console.warn(`[Form] Nationality not found for: "${name}".`);
    }
  }

  private async setPreviousNationality(nationalityCode: string): Promise<void> {
    const name = GdrfaPortalPage.mrzCodeToCountryName(nationalityCode);
    console.log(`[Form] Setting Previous Nationality: "${name}"...`);
    const result = await this.driver.executeScript<{ found: boolean; matched: string }>(`
      var search = arguments[0];
      var sel = document.querySelector('select[id*="wtcmbApplicantPreviousNationality"]');
      if (!sel) return { found: false, matched: '' };
      sel.removeAttribute('disabled');
      var match = Array.from(sel.options).find(function(o) { return o.text.toUpperCase().indexOf(search.toUpperCase()) >= 0; });
      if (!match) return { found: false, matched: '' };
      sel.value = match.value;
      sel.dispatchEvent(new Event('change', { bubbles: true }));
      return { found: true, matched: match.text };
    `, name);
    if (result.found) {
      await this.waitForPageLoad();
      console.log(`[Form] Previous Nationality set: "${result.matched}".`);
    } else {
      console.warn(`[Form] Previous Nationality not found for: "${name}".`);
    }
  }

  private async clickSearchDataAndWait(): Promise<void> {
    console.log('[Form] Clicking Search Data...');
    const btn = await this.waitForElement(
      By.xpath('//a[contains(text(),"Search Data")] | //button[contains(text(),"Search Data")] | //input[@value="Search Data"]'),
      10000
    );
    await btn.click();
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

    const birthCountry = GdrfaPortalPage.mrzCodeToCountryName(passport.birthCountry);
    console.log(`[Form] Setting Birth Country: "${birthCountry}"...`);
    const bcResult = await this.selectByLabel('Birth Country', birthCountry);
    if (bcResult.skipped) {
      console.log(`[Skip] Birth Country already set: "${bcResult.matched}".`);
    } else if (bcResult.found) {
      await this.waitForPageLoad();
      console.log(`[Form] Birth Country set: "${bcResult.matched}".`);
    } else {
      console.warn(`[Form] Birth Country not found for: "${birthCountry}".`);
    }

    // Birth Place EN
    const birthPlace = /^[A-Z]{3}$/.test(passport.birthPlaceEN.trim())
      ? GdrfaPortalPage.mrzCodeToCountryName(passport.birthPlaceEN.trim())
      : passport.birthPlaceEN;
    console.log(`[Form] Filling Birth Place EN: "${birthPlace}"...`);
    const bpFilled = await this.editAndFill('inpApplicantBirthPlaceEn', birthPlace);
    if (bpFilled) {
      await this.driver.executeScript(`if (window.translateInputText) translateInputText('inpApplicantBirthPlaceEn');`);
      await this.sleep(200);
      console.log('[Form] Birth Place EN filled + translated.');
    }

    console.log(`[Form] Setting Gender: "${passport.gender}"...`);
    const gResult = await this.selectByLabel('Gender', passport.gender);
    if (gResult.skipped) {
      console.log(`[Skip] Gender already set: "${gResult.matched}".`);
    } else if (gResult.found) {
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
        console.log(`[Form] Passport Issue Country set: "${icResult.matched}".`);
      } else {
        console.warn(`[Form] Passport Issue Country not found for: "${issueCountry}".`);
      }
    } else {
      console.log('[Form] Passport Issue Country — skipped (empty).');
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

    if (passport.passportPlaceOfIssueEN) {
      const placeOfIssue = /^[A-Z]{3}$/.test(passport.passportPlaceOfIssueEN.trim())
        ? GdrfaPortalPage.mrzCodeToCountryName(passport.passportPlaceOfIssueEN.trim())
        : passport.passportPlaceOfIssueEN;
      console.log(`[Form] Filling Place of Issue EN: "${placeOfIssue}"...`);
      const poiFilled = await this.editAndFill('inpPassportPlaceIssueEn', placeOfIssue);
      if (poiFilled) {
        await this.driver.executeScript(`if (window.translateInputText) translateInputText('inpPassportPlaceIssueEn');`);
        await this.sleep(200);
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

    // Mother Name EN
    if (applicant.motherNameEN) {
      console.log(`[Form] Filling Mother Name EN: "${applicant.motherNameEN}"...`);
      const motherFilled = await this.editAndFill('inpMotherNameEn', applicant.motherNameEN);
      if (motherFilled) {
        await this.driver.executeScript(`if (window.translateInputText) translateInputText('inpMotherNameEn');`);
        await this.sleep(200);
        console.log('[Form] Mother Name EN filled + translated.');
      }
    } else {
      console.log('[Form] Mother Name EN — skipped (empty).');
    }

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

    // Profession
    {
      const currentProf = await this.driver.executeScript<string>(`
        var hidden = document.querySelector('input[id*="wtProfession"][type="hidden"]');
        return hidden ? (hidden.value || '').trim() : '';
      `);
      if (currentProf) {
        console.log(`[Skip] Profession already set (value: "${currentProf}").`);
      } else {
        console.log('[Form] Filling Profession: typing "SALES" → selecting "SALES EXECUTIVE"...');
        try {
          const profInput = await this.driver.findElement(By.css('input[id*="wtProfessionSerch"]'));
          if (await profInput.isDisplayed().catch(() => false)) {
            await profInput.click();
            await profInput.clear();

            // Type character by character
            for (const char of 'SALES') {
              await profInput.sendKeys(char);
              await this.sleep(40);
            }

            // Wait for autocomplete dropdown
            await this.sleep(1000);
            const matched = await this.driver.executeScript<boolean>(`
              var items = document.querySelectorAll('ul.os-internal-ui-autocomplete li.os-internal-ui-menu-item');
              for (var i = 0; i < items.length; i++) {
                if (items[i].textContent.trim().toUpperCase() === 'SALES EXECUTIVE') {
                  items[i].click();
                  return true;
                }
              }
              if (items.length > 0) { items[0].click(); return true; }
              return false;
            `);
            if (matched) {
              console.log('[Form] Profession selected.');
            } else {
              console.warn('[Form] Profession autocomplete suggestions not found.');
            }
          } else {
            console.warn('[Form] Profession input not found.');
          }
        } catch {
          console.warn('[Form] Profession input not found.');
        }
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
  }

  // ── Contact detail fields ─────────────────────────────────────────────────

  private async fillContactDetails(contact: ContactDetails): Promise<void> {
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

    // ── Address Inside UAE ──────────────────────────────────────────────────

    if (contact.uaeEmirate) {
      console.log(`[Form] Setting Emirate: "${contact.uaeEmirate}"...`);
      const emResult = await this.selectByStaticId('cmbAddressInsideEmiratesId', contact.uaeEmirate);
      if (emResult.skipped) {
        console.log(`[Skip] Emirate already set: "${emResult.matched}".`);
      } else if (emResult.found) {
        await this.waitForPageLoad();
        console.log(`[Form] Emirate set: "${emResult.matched}".`);
      } else {
        console.warn(`[Form] Emirate not found for: "${contact.uaeEmirate}".`);
      }
    }

    if (contact.uaeCity) {
      console.log(`[Form] Setting City: "${contact.uaeCity}"...`);
      await this.waitForCondition(async () => {
        return this.driver.executeScript<boolean>(`
          var sel = document.querySelector('select[data-staticid="cmbAddressInsideCityId"]');
          return sel ? sel.options.length > 1 : false;
        `);
      }, 10000).catch(() => console.warn('[Form] City dropdown did not populate in time.'));
      const cityResult = await this.selectByStaticId('cmbAddressInsideCityId', contact.uaeCity);
      if (cityResult.skipped) {
        console.log(`[Skip] City already set: "${cityResult.matched}".`);
      } else if (cityResult.found) {
        await this.waitForPageLoad();
        console.log(`[Form] City set: "${cityResult.matched}".`);
      } else {
        console.warn(`[Form] City not found for: "${contact.uaeCity}".`);
      }
    }

    if (contact.uaeArea) {
      console.log(`[Form] Setting Area: "${contact.uaeArea}"...`);
      await this.waitForCondition(async () => {
        return this.driver.executeScript<boolean>(`
          var sel = document.querySelector('select[data-staticid="cmbAddressInsideAreaId"]');
          return sel ? sel.options.length > 1 : false;
        `);
      }, 10000).catch(() => console.warn('[Form] Area dropdown did not populate in time.'));
      const areaResult = await this.selectByStaticId('cmbAddressInsideAreaId', contact.uaeArea);
      if (areaResult.skipped) {
        console.log(`[Skip] Area already set: "${areaResult.matched}".`);
      } else if (areaResult.found) {
        console.log(`[Form] Area set: "${areaResult.matched}".`);
      } else {
        console.warn(`[Form] Area not found for: "${contact.uaeArea}".`);
      }
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

    // ── Address Outside UAE ─────────────────────────────────────────────────

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
    await this.sleep(100);

    // Remove ReadOnly from the drop panel
    await this.driver.executeScript(`
      var drop = document.querySelector('.select2-drop-active');
      if (drop) {
        drop.classList.remove('ReadOnly');
        drop.querySelectorAll('.ReadOnly').forEach(function(el) { el.classList.remove('ReadOnly'); });
      }
    `);
    await this.sleep(100);

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

    // Passport Names
    if (app.passport.firstName && !await inputHasValue('inpFirsttNameEn')) {
      empty.push({ name: 'First Name', retry: () => this.editAndFill('inpFirsttNameEn', app.passport.firstName) });
    }
    if (app.passport.lastName && !await inputHasValue('inpLastNameEn')) {
      empty.push({ name: 'Last Name', retry: () => this.editAndFill('inpLastNameEn', app.passport.lastName) });
    }

    // Passport Details
    if (app.passport.dateOfBirth && !await inputHasValue('inpDateOfBirth')) {
      empty.push({ name: 'Date of Birth', retry: () => this.editAndFill('inpDateOfBirth', app.passport.dateOfBirth) });
    }
    if (app.passport.birthPlaceEN && !await inputHasValue('inpApplicantBirthPlaceEn')) {
      empty.push({ name: 'Birth Place', retry: () => this.editAndFill('inpApplicantBirthPlaceEn', app.passport.birthPlaceEN) });
    }
    if (app.passport.passportIssueDate && !await inputHasValue('inpPassportIssueDate')) {
      empty.push({ name: 'Passport Issue Date', retry: () => this.editAndFill('inpPassportIssueDate', app.passport.passportIssueDate) });
    }
    if (app.passport.passportExpiryDate && !await inputHasValue('inpPassportExpiryDate')) {
      empty.push({ name: 'Passport Expiry Date', retry: () => this.editAndFill('inpPassportExpiryDate', app.passport.passportExpiryDate) });
    }
    if (app.passport.passportPlaceOfIssueEN && !await inputHasValue('inpPassportPlaceIssueEn')) {
      const poi = /^[A-Z]{3}$/.test(app.passport.passportPlaceOfIssueEN.trim())
        ? GdrfaPortalPage.mrzCodeToCountryName(app.passport.passportPlaceOfIssueEN.trim())
        : app.passport.passportPlaceOfIssueEN;
      empty.push({ name: 'Passport Place of Issue', retry: () => this.editAndFill('inpPassportPlaceIssueEn', poi) });
    }

    // Passport Selects
    if (app.passport.gender && !await selectByLabelHasValue('Gender')) {
      empty.push({ name: 'Gender', retry: async () => { await this.selectByLabel('Gender', app.passport.gender); } });
    }
    if (app.passport.birthCountry && !await selectByLabelHasValue('Birth Country')) {
      empty.push({ name: 'Birth Country', retry: async () => { await this.selectByLabel('Birth Country', app.passport.birthCountry); } });
    }
    if (app.passport.passportIssueCountry && !await selectByLabelHasValue('Passport Issue Country')) {
      empty.push({ name: 'Passport Issue Country', retry: async () => { await this.selectByLabel('Passport Issue Country', app.passport.passportIssueCountry); } });
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

  private async waitForLoaderToDisappear(): Promise<void> {
    try {
      const loaders = await this.driver.findElements(By.css('div.Feedback_AjaxWait'));
      for (const loader of loaders) {
        if (await loader.isDisplayed().catch(() => false)) {
          console.log('[Loader] Waiting for loader to disappear...');
          await this.driver.wait(until.stalenessOf(loader), 30000).catch(() => {});
          console.log('[Loader] Loader gone.');
        }
      }
    } catch {
      // Loader may have already disappeared
    }
  }

  private async clickContinue(): Promise<void> {
    console.log('[Form] Clicking Continue...');
    const btn = await this.driver.findElement(By.css('input[staticid="SmartChannels_EntryPermitNewTourism_btnContinue"]'));
    await this.driver.executeScript('arguments[0].scrollIntoView({block:"center"});', btn);
    await btn.click();
    console.log('[Form] Continue clicked — waiting for popup or next page...');

    const popupFrame = await this.findPopupFrame(30000);

    if (popupFrame) {
      console.log('[Form] Existing application popup detected (in iframe).');
      await this.handleExistingApplicationPopup();
    } else {
      console.log('[Form] No popup — proceeding...');
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
    searchValue: string
  ): Promise<{ found: boolean; matched: string; skipped?: boolean }> {
    return this.driver.executeScript<{ found: boolean; matched: string; skipped?: boolean }>(`
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
    `, staticId, searchValue);
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

    // Click the first matching result using mousedown (Select2 uses mousedown, not click)
    const matchedText = await this.driver.executeScript<string>(`
      var results = document.querySelectorAll('.select2-drop-active .select2-results li.select2-result');
      if (results.length === 0) return '';
      var item = results[0];
      var text = item.textContent.trim();
      // Select2 listens on mouseup after mousedown — simulate the full sequence
      item.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      item.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
      item.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
      return text;
    `);

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
        const state = await this.driver.executeScript<string>('return document.readyState');
        return state === 'complete';
      }, timeout);
      // Extra wait for AJAX to settle
      await this.sleep(500);
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
    return map[code.toUpperCase()] ?? code;
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
