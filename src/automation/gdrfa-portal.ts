import { Page } from '@playwright/test';
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

  constructor(private readonly page: Page) {}

  // ── Public API ─────────────────────────────────────────────────────────────

  async verifySession(): Promise<void> {
    console.log('[Session] Verifying session...');
    await this.page.goto(GdrfaPortalPage.HOME, { waitUntil: 'networkidle' });
    if (this.page.url().includes('Login.aspx')) {
      throw new Error('[Session] Session expired — run "npm run auth" to log in again.');
    }
    console.log('[Session] Valid. URL:', this.page.url());
  }

  async fillApplicationForm(application: VisaApplication): Promise<void> {
    console.log('\n[Flow] ─── Starting navigation ───');
    const stopKeepAlive = this.startSessionKeepAlive();
    try {
      await this.verifySession();
      await this.navigateToNewApplication();
      await this.waitForPageSettle();

      await this.setVisitReason();
      await this.setPassportType(application.passport.passportType);
      await this.enterPassportNumber(application.passport.passportNumber);
      await this.setNationality(application.passport.currentNationality);
      await this.setPreviousNationality(
        application.passport.previousNationality ?? application.passport.currentNationality
      );
      await this.clickSearchDataAndWait();

      await this.fillPassportNames(application.passport);
      await this.fillPassportDetails(application.passport);
      await this.fillApplicantDetails(application.applicant);
      await this.fillContactDetails(application.contact);

      // Retry Faith selection before continuing (dropdown can reset after other fields)
      if (application.applicant.faith) {
        await this.retryFaithSelection(application.applicant.faith);
      }

      // Validate all required fields before clicking Continue — retry any empty ones
      await this.validateAndRetryRequiredFields(application);

      await this.clickContinue();

      // Upload documents on the Attachments tab
      await this.uploadDocuments(application.documents);

      console.log('\n[Flow] ─── Steps complete. ───\n');
    } finally {
      stopKeepAlive();
    }
  }

  async uploadDocuments(docs: ApplicationDocuments): Promise<void> {
    console.log('[Upload] Starting document upload...');
    await this.page.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {});

    // Map document type labels to file paths.
    // Labels must match the data-document-type attribute on each input[type="file"].
    const slots: Array<{ label: string; file: string }> = [
      { label: 'Hotel reservation/Place of stay - Page 1', file: docs.hotelReservationPage1 },
      { label: 'Hotel reservation/Place of stay - Page 2', file: docs.hotelReservationPage2 ?? '' },
      { label: 'Passport External Cover Page',             file: docs.passportExternalCoverPage },
      { label: 'Personal Photo',                           file: docs.personalPhoto },
      { label: 'Return air ticket - Page 1',               file: docs.returnAirTicketPage1 },
      { label: 'Sponsored Passport page 1',                file: docs.sponsoredPassportPage1 },
    ];

    // Log all available document types on the page for debugging
    const availableTypes = await this.page.evaluate(() =>
      Array.from(document.querySelectorAll<HTMLInputElement>('input[type="file"][data-document-type]'))
        .map(el => el.getAttribute('data-document-type') ?? '')
    );
    console.log(`[Upload] Available document types on page: ${JSON.stringify(availableTypes)}`);

    for (const slot of slots) {
      if (!slot.file) continue;

      const filePath = path.resolve(slot.file);
      if (!fs.existsSync(filePath)) {
        console.warn(`[Upload] File not found, skipping "${slot.label}": ${filePath}`);
        continue;
      }

      // Match by data-document-type attribute (exact match first, then case-insensitive)
      let fileInput = this.page.locator(`input[type="file"][data-document-type="${slot.label}"]`);
      if (await fileInput.count() === 0) {
        // Fallback: case-insensitive search via evaluate
        const matchIdx = await this.page.evaluate((label: string) => {
          const inputs = Array.from(document.querySelectorAll<HTMLInputElement>('input[type="file"][data-document-type]'));
          const target = label.toLowerCase();
          const idx = inputs.findIndex(el => (el.getAttribute('data-document-type') ?? '').toLowerCase() === target);
          return idx;
        }, slot.label);

        if (matchIdx < 0) {
          console.warn(`[Upload] Slot not found on page: "${slot.label}"`);
          continue;
        }
        fileInput = this.page.locator('input[type="file"][data-document-type]').nth(matchIdx);
      }

      await fileInput.scrollIntoViewIfNeeded().catch(() => {});
      await fileInput.setInputFiles(filePath);
      console.log(`[Upload] "${slot.label}": ${path.basename(filePath)}`);

      // Wait for the upload to process
      await this.page.waitForLoadState('networkidle', { timeout: 15000 }).catch(() => {});
    }

    console.log('[Upload] All documents uploaded. Waiting for Continue button...');

    // Dismiss "Existing Application" popup if it reappears on the upload page
    const popupFrame = await this.findPopupFrame(5000);
    if (popupFrame) {
      console.log('[Upload] Existing application popup detected — dismissing...');
      await this.handleExistingApplicationPopup(popupFrame);
    }

    // Continue button appears only after all mandatory uploads are complete
    const continueBtn = this.page.locator(
      'input[value="Continue"], button:has-text("Continue"), a:has-text("Continue")'
    ).first();
    await continueBtn.waitFor({ state: 'visible', timeout: 60000 });
    await continueBtn.scrollIntoViewIfNeeded();
    await continueBtn.click({ force: true, noWaitAfter: true });
    await this.page.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {});
    await this.waitForLoaderToDisappear();
    console.log('[Upload] Continue clicked — done.');
  }

  // ── Navigation ─────────────────────────────────────────────────────────────

  private async dismissPromoPopup(): Promise<void> {
    const SKIP_ID = 'WebPatterns_wt2_block_wtMainContent_wt3_EmaratechSG_Patterns_wt8_block_wtMainContent_wt10';
    try {
      const skipBtn = this.page.frameLocator('iframe').locator(`#${SKIP_ID}, input[value="Skip"]`).first();
      if (!await skipBtn.isVisible({ timeout: 3000 }).catch(() => false)) return;
      await skipBtn.click();
      console.log('[Nav] Dismissed promotional popup.');
    } catch { /* non-fatal — popup does not appear on every load */ }
  }

  private async navigateToNewApplication(): Promise<void> {
    console.log('[Nav] Navigating to Existing Applications...');
    await this.dismissPromoPopup();

    await this.page.locator(
      '#EmaratechSG_Theme_wtwbLayoutEmaratech_block_wtMainContent_wtwbDashboard_wtCntExistApp, ' +
      'a:has-text("Existing Applications")'
    ).first().click({ timeout: 15000 });

    // Wait for the establishment detail page to fully load (avoids race with secondary redirects)
    await this.page.waitForURL('**/EstablishmentDetail**', { timeout: 20000 }).catch(() => {});
    await this.page.waitForLoadState('domcontentloaded', { timeout: 15000 }).catch(() => {});
    await this.page.waitForLoadState('networkidle', { timeout: 15000 }).catch(() => {});
    await this.dismissPromoPopup();

    const dropdownSel =
      '#EmaratechSG_Theme_wtwbLayoutEmaratechWithoutTitle_block_wtMainContent_EmaratechSG_Patterns_wtwbEstbButtonWithContextInfo_block_wtIcon_wtcntContextActionBtn';
    const firstOptionSel =
      '#EmaratechSG_Theme_wtwbLayoutEmaratechWithoutTitle_block_wtMainContent_EmaratechSG_Patterns_wtwbEstbButtonWithContextInfo_block_wtContent_wtwbEstbTopServices_wtListMyServicesExperiences_ctl00_wtStartTopService';

    const dropdown = this.page.locator(dropdownSel);
    const dropdownVisible = await dropdown.waitFor({ state: 'visible', timeout: 15000 })
      .then(() => true).catch(() => false);

    if (dropdownVisible) {
      await dropdown.click();
      await this.page.waitForTimeout(150);

      const firstOption = this.page.locator(firstOptionSel);
      await firstOption.waitFor({ state: 'visible', timeout: 10000 });
      console.log(`[Nav] Selecting form: "${(await firstOption.textContent())?.trim()}"`);
      await firstOption.click();
    } else {
      // Fallback: dropdown ID not found — click "New Application" button then first service link
      console.log('[Nav] Dropdown not found — using fallback navigation.');
      const newAppBtn = this.page.locator('button:has-text("New Application")');
      if (await newAppBtn.isVisible({ timeout: 5000 }).catch(() => false)) {
        await newAppBtn.click();
        await this.page.waitForTimeout(300);
      }
      const firstService = this.page.locator(
        `${firstOptionSel}, a:has-text("New Tourism Entry Permit")`
      ).first();
      await firstService.waitFor({ state: 'visible', timeout: 10000 });
      console.log(`[Nav] Selecting form (fallback): "${(await firstService.textContent())?.trim()}"`);
      await firstService.click();
    }

    await this.page.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {});
    console.log('[Nav] Form page loaded. URL:', this.page.url());
  }

  private async waitForPageSettle(): Promise<void> {
    await this.page.waitForLoadState('domcontentloaded', { timeout: 15000 }).catch(() => {});
    await this.page.waitForLoadState('networkidle',      { timeout: 15000 }).catch(() => {});
  }

  // ── Passport header (Passport Type → Nationality → Search Data) ────────────

  private async setVisitReason(): Promise<void> {
    console.log('[Form] Selecting Visit Reason → Tourism...');
    // The select may be offscreen (in Visit Details section) — wait for DOM attachment, then scroll
    const sel = this.page.locator('select[data-staticid="cmbVisitReason"]');
    const attached = await sel.waitFor({ state: 'attached', timeout: 15000 }).then(() => true).catch(() => false);
    if (!attached) {
      console.log('[Form] Visit Reason select not found — likely pre-set by form type. Skipping.');
      return;
    }
    await sel.scrollIntoViewIfNeeded().catch(() => {});
    const set = await this.page.evaluate(() => {
      const sel = document.querySelector<HTMLSelectElement>('select[data-staticid="cmbVisitReason"]');
      if (!sel) return false;
      // Skip if already set to Tourism
      const currentText = sel.options[sel.selectedIndex]?.text?.trim() ?? '';
      if (currentText.toUpperCase().includes('TOURISM')) return true;
      sel.value = '1'; // 1 = Tourism
      sel.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    });
    if (set) {
      await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
      console.log('[Form] Visit Reason → Tourism.');
    } else {
      console.warn('[Form] Visit Reason select not found.');
    }
  }

  private async setPassportType(passportType: string): Promise<void> {
    console.log(`[Form] Setting Passport Type: "${passportType}"...`);
    const set = await this.page.evaluate((type: string) => {
      // Passport Type select is identified by having "Normal" as one of its options
      const sel = Array.from(document.querySelectorAll<HTMLSelectElement>('select')).find(s =>
        Array.from(s.options).some(o => o.text.trim() === 'Normal')
      );
      if (!sel) return false;
      const match = Array.from(sel.options).find(o => o.text.trim().toLowerCase() === type.toLowerCase());
      if (!match) return false;
      sel.value = match.value;
      sel.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    }, passportType);
    if (set) {
      await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
      console.log('[Form] Passport Type set.');
    } else {
      console.warn('[Form] Passport Type select not found.');
    }
  }

  private async enterPassportNumber(passportNumber: string): Promise<void> {
    console.log(`[Form] Entering Passport Number: "${passportNumber}"...`);
    const input = this.page.locator(
      'input[staticid*="PassportNo"], input[id*="inptPassportNo"], input[id*="PassportNo"]'
    ).first();
    await input.waitFor({ state: 'attached', timeout: 15000 });
    await input.scrollIntoViewIfNeeded().catch(() => {});
    await input.fill(passportNumber);
    console.log('[Form] Passport Number entered.');
  }

  private async setNationality(nationalityCode: string): Promise<void> {
    const name = GdrfaPortalPage.mrzCodeToCountryName(nationalityCode);
    console.log(`[Form] Setting Nationality: "${name}"...`);
    const result = await this.page.evaluate((search: string) => {
      // Nationality selects have options formatted as "NNN - COUNTRY NAME"
      const sel = Array.from(document.querySelectorAll<HTMLSelectElement>('select')).find(s =>
        Array.from(s.options).some(o => /^\d+ - /.test(o.text.trim()))
      );
      if (!sel) return { found: false, matched: '' };
      sel.removeAttribute('disabled');
      const match = Array.from(sel.options).find(o => o.text.toUpperCase().includes(search.toUpperCase()));
      if (!match) return { found: false, matched: '' };
      sel.value = match.value;
      sel.dispatchEvent(new Event('change', { bubbles: true }));
      return { found: true, matched: match.text };
    }, name);
    if (result.found) {
      await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
      console.log(`[Form] Nationality set: "${result.matched}".`);
    } else {
      console.warn(`[Form] Nationality not found for: "${name}".`);
    }
  }

  private async setPreviousNationality(nationalityCode: string): Promise<void> {
    const name = GdrfaPortalPage.mrzCodeToCountryName(nationalityCode);
    console.log(`[Form] Setting Previous Nationality: "${name}"...`);
    const result = await this.page.evaluate((search: string) => {
      const sel = document.querySelector<HTMLSelectElement>('select[id*="wtcmbApplicantPreviousNationality"]');
      if (!sel) return { found: false, matched: '' };
      sel.removeAttribute('disabled');
      const match = Array.from(sel.options).find(o => o.text.toUpperCase().includes(search.toUpperCase()));
      if (!match) return { found: false, matched: '' };
      sel.value = match.value;
      sel.dispatchEvent(new Event('change', { bubbles: true }));
      return { found: true, matched: match.text };
    }, name);
    if (result.found) {
      await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
      console.log(`[Form] Previous Nationality set: "${result.matched}".`);
    } else {
      console.warn(`[Form] Previous Nationality not found for: "${name}".`);
    }
  }

  private async clickSearchDataAndWait(): Promise<void> {
    console.log('[Form] Clicking Search Data...');
    const btn = this.page.locator(
      'a:has-text("Search Data"), button:has-text("Search Data"), input[value="Search Data"]'
    ).first();
    await btn.waitFor({ state: 'visible', timeout: 10000 });
    await btn.click();
    // Wait for the portal AJAX call to populate SmartInput fields and re-render widgets
    await this.page.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {});

    // Wait for key passport input fields to appear in the DOM
    console.log('[Form] Waiting for passport fields to load...');
    await Promise.all([
      this.page.locator('input[data-staticid="inpFirsttNameEn"]').waitFor({ state: 'attached', timeout: 20000 }),
      this.page.locator('input[data-staticid="inpLastNameEn"]').waitFor({ state: 'attached', timeout: 20000 }),
      this.page.locator('input[data-staticid="inpDateOfBirth"]').waitFor({ state: 'attached', timeout: 20000 }),
      this.page.locator('input[data-staticid="inpPassportExpiryDate"]').waitFor({ state: 'attached', timeout: 20000 }),
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
      await this.page.evaluate(() => (window as any).translateInputText?.('inpFirsttNameEn'));
      await this.page.waitForTimeout(200);
      console.log('[Form] First Name filled + translated.');
    }

    if (passport.middleName) {
      console.log(`[Form] Filling Middle Name: "${passport.middleName}"...`);
      const midFilled = await this.editAndFill('inpMiddleNameEn', passport.middleName);
      if (midFilled) {
        await this.page.evaluate(() => (window as any).translateInputText?.('inpMiddleNameEn'));
        await this.page.waitForTimeout(200);
        console.log('[Form] Middle Name filled + translated.');
      }
    } else {
      console.log('[Form] No middle name — field left blank.');
    }

    console.log(`[Form] Filling Last Name: "${passport.lastName}"...`);
    const lastFilled = await this.editAndFill('inpLastNameEn', passport.lastName);
    if (lastFilled) {
      await this.page.evaluate(() => (window as any).translateInputText?.('inpLastNameEn'));
      await this.page.waitForTimeout(200);
      console.log('[Form] Last Name filled + translated.');
    }
  }

  // ── Passport detail fields (below name fields) ─────────────────────────────

  private async fillPassportDetails(passport: PassportDetails): Promise<void> {
    // Date of Birth
    const dob = passport.dateOfBirth.replace(/\//g, '-');  // DD/MM/YYYY → DD-MM-YYYY
    console.log(`[Form] Filling Date of Birth: "${dob}"...`);
    // Check current value BEFORE clearing — skip if already correct
    const currentDob = await this.page.evaluate(() => {
      const el = document.querySelector<HTMLInputElement>('input[data-staticid="inpDateOfBirth"]');
      return el?.value?.trim() ?? '';
    });
    if (currentDob && currentDob.toUpperCase() === dob.toUpperCase()) {
      console.log(`[Skip] Date of Birth already has correct value: "${currentDob}".`);
    } else {
      await this.page.evaluate(() => {
        const el = document.querySelector<HTMLInputElement>('input[data-staticid="inpDateOfBirth"]');
        if (el) { el.value = ''; el.dispatchEvent(new Event('input', { bubbles: true })); }
      });
      await this.editAndFill('inpDateOfBirth', dob);
      await this.page.evaluate(() => {
        const el = document.querySelector<HTMLInputElement>('input[data-staticid="inpDateOfBirth"]');
        if (el) el.dispatchEvent(new Event('change', { bubbles: true }));
      });
      await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
      console.log('[Form] Date of Birth filled.');
    }

    const birthCountry = GdrfaPortalPage.mrzCodeToCountryName(passport.birthCountry);
    console.log(`[Form] Setting Birth Country: "${birthCountry}"...`);
    const bcResult = await this.selectByLabel('Birth Country', birthCountry);
    if (bcResult.skipped) {
      console.log(`[Skip] Birth Country already set: "${bcResult.matched}".`);
    } else if (bcResult.found) {
      await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
      console.log(`[Form] Birth Country set: "${bcResult.matched}".`);
    } else {
      console.warn(`[Form] Birth Country not found for: "${birthCountry}".`);
    }

    // Birth Place EN: if the JSON holds a raw 3-letter MRZ code, convert it to a country name.
    const birthPlace = /^[A-Z]{3}$/.test(passport.birthPlaceEN.trim())
      ? GdrfaPortalPage.mrzCodeToCountryName(passport.birthPlaceEN.trim())
      : passport.birthPlaceEN;
    console.log(`[Form] Filling Birth Place EN: "${birthPlace}"...`);
    const bpFilled = await this.editAndFill('inpApplicantBirthPlaceEn', birthPlace);
    if (bpFilled) {
      await this.page.evaluate(() => (window as any).translateInputText('inpApplicantBirthPlaceEn'));
      await this.page.waitForTimeout(200);
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
      const currentIssue = await this.page.evaluate(() => {
        const el = document.querySelector<HTMLInputElement>('input[data-staticid="inpPassportIssueDate"]');
        return el?.value?.trim() ?? '';
      });
      if (currentIssue && currentIssue.toUpperCase() === issueDate.toUpperCase()) {
        console.log(`[Skip] Passport Issue Date already has correct value: "${currentIssue}".`);
      } else {
        await this.page.evaluate(() => {
          const el = document.querySelector<HTMLInputElement>('input[data-staticid="inpPassportIssueDate"]');
          if (el) { el.value = ''; el.dispatchEvent(new Event('input', { bubbles: true })); }
        });
        await this.editAndFill('inpPassportIssueDate', issueDate);
        await this.page.evaluate(() => {
          const el = document.querySelector<HTMLInputElement>('input[data-staticid="inpPassportIssueDate"]');
          if (el) el.dispatchEvent(new Event('change', { bubbles: true }));
        });
        await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
        console.log('[Form] Passport Issue Date filled.');
      }
    } else {
      console.log('[Form] Passport Issue Date — skipped (empty).');
    }

    const expiryDate = passport.passportExpiryDate.replace(/\//g, '-');
    console.log(`[Form] Filling Passport Expiry Date: "${expiryDate}"...`);
    const currentExpiry = await this.page.evaluate(() => {
      const el = document.querySelector<HTMLInputElement>('input[data-staticid="inpPassportExpiryDate"]');
      return el?.value?.trim() ?? '';
    });
    if (currentExpiry && currentExpiry.toUpperCase() === expiryDate.toUpperCase()) {
      console.log(`[Skip] Passport Expiry Date already has correct value: "${currentExpiry}".`);
    } else {
      await this.page.evaluate(() => {
        const el = document.querySelector<HTMLInputElement>('input[data-staticid="inpPassportExpiryDate"]');
        if (el) { el.value = ''; el.dispatchEvent(new Event('input', { bubbles: true })); }
      });
      await this.editAndFill('inpPassportExpiryDate', expiryDate);
      await this.page.evaluate(() => {
        const el = document.querySelector<HTMLInputElement>('input[data-staticid="inpPassportExpiryDate"]');
        if (el) el.dispatchEvent(new Event('change', { bubbles: true }));
      });
      await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
      console.log('[Form] Passport Expiry Date filled.');
    }

    if (passport.passportPlaceOfIssueEN) {
      const placeOfIssue = /^[A-Z]{3}$/.test(passport.passportPlaceOfIssueEN.trim())
        ? GdrfaPortalPage.mrzCodeToCountryName(passport.passportPlaceOfIssueEN.trim())
        : passport.passportPlaceOfIssueEN;
      console.log(`[Form] Filling Place of Issue EN: "${placeOfIssue}"...`);
      const poiFilled = await this.editAndFill('inpPassportPlaceIssueEn', placeOfIssue);
      if (poiFilled) {
        await this.page.evaluate(() => (window as any).translateInputText?.('inpPassportPlaceIssueEn'));
        await this.page.waitForTimeout(200);
        console.log('[Form] Place of Issue EN filled + translated.');
      }
    } else {
      console.log('[Form] Place of Issue EN — skipped (empty).');
    }
  }

  // ── Applicant detail fields ────────────────────────────────────────────────

  private async fillApplicantDetails(applicant: ApplicantDetails): Promise<void> {
    // Is Inside UAE checkbox — only interact if applicant IS inside (default is unchecked)
    if (applicant.isInsideUAE) {
      console.log('[Form] Checking Is Inside UAE...');
      await this.page.evaluate(() => {
        const cb = document.querySelector<HTMLInputElement>('input[data-staticid="chkIsInsideUAE"]');
        if (cb && !cb.checked) {
          cb.classList.remove('ReadOnly');
          cb.checked = true;
          cb.dispatchEvent(new MouseEvent('click', { bubbles: true }));
        }
      });
      await this.page.waitForLoadState('networkidle', { timeout: 15000 }).catch(() => {});
      console.log('[Form] Is Inside UAE checked.');
    }

    // Mother Name EN (+ auto-translate to Arabic)
    if (applicant.motherNameEN) {
      console.log(`[Form] Filling Mother Name EN: "${applicant.motherNameEN}"...`);
      const motherFilled = await this.editAndFill('inpMotherNameEn', applicant.motherNameEN);
      if (motherFilled) {
        await this.page.evaluate(() => (window as any).translateInputText?.('inpMotherNameEn'));
        await this.page.waitForTimeout(200);
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

    // Religion (AJAX onChange repopulates Faith dropdown)
    if (applicant.religion) {
      console.log(`[Form] Setting Religion: "${applicant.religion}"...`);
      const rResult = await this.selectByLabel('Religion', applicant.religion);
      if (rResult.skipped) {
        console.log(`[Skip] Religion already set: "${rResult.matched}".`);
      } else if (rResult.found) {
        await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
        console.log(`[Form] Religion set: "${rResult.matched}".`);
      } else {
        console.warn(`[Form] Religion not found for: "${applicant.religion}".`);
      }
    }

    // Faith (options depend on Religion — must be set after Religion AJAX resolves)
    if (applicant.faith) {
      console.log(`[Form] Setting Faith: "${applicant.faith}"...`);

      // Check if Faith already has correct value — skip the entire Select2 interaction
      const currentFaith = await this.page.evaluate(() => {
        const sel = document.querySelector<HTMLSelectElement>('select[data-staticid="cmbApplicantFaith"]');
        if (!sel) return '';
        return sel.options[sel.selectedIndex]?.text?.trim() ?? '';
      });
      if (currentFaith && !currentFaith.includes('Select') &&
          currentFaith.toUpperCase().includes(applicant.faith.toUpperCase())) {
        console.log(`[Skip] Faith already set: "${currentFaith}".`);
      } else {
        // Wait until the Faith dropdown has been populated (more than just "-- Select --")
        await this.page.waitForFunction(() => {
          const sel = document.querySelector<HTMLSelectElement>('select[data-staticid="cmbApplicantFaith"]');
          return sel ? sel.options.length > 1 : false;
        }, { timeout: 10000 }).catch(() => console.warn('[Form] Faith dropdown did not populate in time.'));

        // Remove ReadOnly from the Select2 container (preceding sibling of the native <select>)
        await this.page.evaluate(() => {
          const sel = document.querySelector('select[data-staticid="cmbApplicantFaith"]');
          if (!sel) return;
          const container = sel.previousElementSibling;
          if (container?.classList.contains('select2-container')) {
            container.classList.remove('ReadOnly');
            container.querySelectorAll('.ReadOnly').forEach(el => el.classList.remove('ReadOnly'));
          }
        });

        // Click the Select2 choice link to open the dropdown
        const faithContainer = this.page.locator('[data-staticid="cmbApplicantFaith"]')
          .locator('xpath=preceding-sibling::div[contains(@class,"select2-container")]');
        await faithContainer.locator('.select2-choice').click({ timeout: 5000 });

        // Remove ReadOnly from the drop panel (separate element appended to <body>)
        await this.page.evaluate(() => {
          const drop = document.querySelector('.select2-drop-active');
          if (drop) {
            drop.classList.remove('ReadOnly');
            drop.querySelectorAll('.ReadOnly').forEach(el => el.classList.remove('ReadOnly'));
          }
        });

        // Click the matching option from the results
        const faithOption = this.page.locator('.select2-drop-active .select2-results li').filter({ hasText: applicant.faith }).first();
        if (await faithOption.isVisible({ timeout: 5000 }).catch(() => false)) {
          const matchedText = await faithOption.textContent() ?? '';
          await faithOption.click();
          await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
          console.log(`[Form] Faith set: "${matchedText.trim()}".`);
        } else {
          // Fallback: set value programmatically
          console.warn(`[Form] Faith UI click failed — trying programmatic fallback...`);
          const fResult = await this.selectByStaticId('cmbApplicantFaith', applicant.faith);
          if (fResult.found) {
            console.log(`[Form] Faith set (programmatic): "${fResult.matched}".`);
          } else {
            console.warn(`[Form] Faith not found for: "${applicant.faith}".`);
          }
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

    // Profession (autocomplete widget — type "SALES" and select "SALES EXECUTIVE")
    {
      // Check if profession is already filled (hidden input stores the selected value)
      const currentProf = await this.page.evaluate(() => {
        const hidden = document.querySelector<HTMLInputElement>('input[id*="wtProfession"][type="hidden"]');
        return hidden?.value?.trim() ?? '';
      });
      if (currentProf) {
        console.log(`[Skip] Profession already set (value: "${currentProf}").`);
      } else {
      console.log('[Form] Filling Profession: typing "SALES" → selecting "SALES EXECUTIVE"...');
      const profInput = this.page
        .locator('input[id*="wtProfessionSerch"]')
        .first();
      if (await profInput.isVisible({ timeout: 3000 }).catch(() => false)) {
        await profInput.click();
        await profInput.clear();
        await this.page.keyboard.type('SALES', { delay: 40 });

        // Wait for the autocomplete dropdown
        const suggestionList = this.page.locator('ul.os-internal-ui-autocomplete li.os-internal-ui-menu-item');
        await suggestionList.first().waitFor({ state: 'visible', timeout: 10000 }).catch(() => {});

        // Find and click "SALES EXECUTIVE"
        const allSuggestions = await suggestionList.all();
        let matched = false;
        for (const suggestion of allSuggestions) {
          const text = ((await suggestion.textContent()) ?? '').trim().toUpperCase();
          if (text === 'SALES EXECUTIVE') {
            await suggestion.click();
            console.log('[Form] Profession selected: "SALES EXECUTIVE".');
            matched = true;
            break;
          }
        }
        if (!matched) {
          console.warn('[Form] "SALES EXECUTIVE" not found in suggestions — selecting first match.');
          if (allSuggestions.length > 0 && await allSuggestions[0].isVisible().catch(() => false)) {
            const fallbackText = (await allSuggestions[0].textContent())?.trim() ?? '';
            await allSuggestions[0].click();
            console.log(`[Form] Profession fallback selected: "${fallbackText}".`);
          }
        }
      } else {
        console.warn('[Form] Profession input not found.');
      }
      } // end else (profession not yet set)
    }

    // Coming From Country (AJAX Select2 — options loaded on search)
    if (applicant.comingFromCountry) {
      const cfcName = GdrfaPortalPage.mrzCodeToCountryName(applicant.comingFromCountry);
      console.log(`[Form] Setting Coming From Country: "${cfcName}"...`);
      const cfcResult = await this.selectByAjaxSelect2('ComingFromCountry', cfcName);
      if (cfcResult) {
        console.log(`[Form] Coming From Country set: "${cfcResult}".`);
      } else {
        console.warn(`[Form] Coming From Country not found for: "${cfcName}".`);
      }
    } else {
      console.log('[Form] Coming From Country — skipped (empty).');
    }
  }

  // ── Contact detail fields ─────────────────────────────────────────────────

  private async fillContactDetails(contact: ContactDetails): Promise<void> {
    // Email
    if (contact.email) {
      console.log(`[Form] Filling Email: "${contact.email}"...`);
      await this.editAndFill('inpEmail', contact.email);
    }

    // Mobile Number
    if (contact.mobileNumber) {
      console.log(`[Form] Filling Mobile Number: "${contact.mobileNumber}"...`);
      await this.editAndFill('inpMobileNumber', contact.mobileNumber);
    }

    // Approval Email Copy (optional)
    if (contact.approvalEmailCopy) {
      console.log(`[Form] Filling Approval Email Copy: "${contact.approvalEmailCopy}"...`);
      await this.editAndFill('inpApprovalEmailCopy', contact.approvalEmailCopy);
    }

    // Preferred SMS Language
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
    // Use data-staticid selectors (not label text) to avoid matching
    // identically-named labels in the Host/Submitter section.

    // Emirate (AJAX onChange populates City)
    if (contact.uaeEmirate) {
      console.log(`[Form] Setting Emirate: "${contact.uaeEmirate}"...`);
      const emResult = await this.selectByStaticId('cmbAddressInsideEmiratesId', contact.uaeEmirate);
      if (emResult.skipped) {
        console.log(`[Skip] Emirate already set: "${emResult.matched}".`);
      } else if (emResult.found) {
        await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
        console.log(`[Form] Emirate set: "${emResult.matched}".`);
      } else {
        console.warn(`[Form] Emirate not found for: "${contact.uaeEmirate}".`);
      }
    }

    // City (populated after Emirate AJAX — wait for options)
    if (contact.uaeCity) {
      console.log(`[Form] Setting City: "${contact.uaeCity}"...`);
      await this.page.waitForFunction(() => {
        const sel = document.querySelector<HTMLSelectElement>('select[data-staticid="cmbAddressInsideCityId"]');
        return sel ? sel.options.length > 1 : false;
      }, { timeout: 10000 }).catch(() => console.warn('[Form] City dropdown did not populate in time.'));
      const cityResult = await this.selectByStaticId('cmbAddressInsideCityId', contact.uaeCity);
      if (cityResult.skipped) {
        console.log(`[Skip] City already set: "${cityResult.matched}".`);
      } else if (cityResult.found) {
        await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
        console.log(`[Form] City set: "${cityResult.matched}".`);
      } else {
        console.warn(`[Form] City not found for: "${contact.uaeCity}".`);
      }
    }

    // Area (populated after City AJAX — wait for options)
    if (contact.uaeArea) {
      console.log(`[Form] Setting Area: "${contact.uaeArea}"...`);
      await this.page.waitForFunction(() => {
        const sel = document.querySelector<HTMLSelectElement>('select[data-staticid="cmbAddressInsideAreaId"]');
        return sel ? sel.options.length > 1 : false;
      }, { timeout: 10000 }).catch(() => console.warn('[Form] Area dropdown did not populate in time.'));
      const areaResult = await this.selectByStaticId('cmbAddressInsideAreaId', contact.uaeArea);
      if (areaResult.skipped) {
        console.log(`[Skip] Area already set: "${areaResult.matched}".`);
      } else if (areaResult.found) {
        console.log(`[Form] Area set: "${areaResult.matched}".`);
      } else {
        console.warn(`[Form] Area not found for: "${contact.uaeArea}".`);
      }
    }

    // Street
    if (contact.uaeStreet) {
      console.log(`[Form] Filling Street: "${contact.uaeStreet}"...`);
      await this.editAndFill('inpAddressInsideStreet2', contact.uaeStreet);
    }

    // Building/Villa
    if (contact.uaeBuilding) {
      console.log(`[Form] Filling Building/Villa: "${contact.uaeBuilding}"...`);
      await this.editAndFill('inpAddressInsideBuilding', contact.uaeBuilding);
    }

    // Floor
    if (contact.uaeFloor) {
      console.log(`[Form] Filling Floor: "${contact.uaeFloor}"...`);
      await this.editAndFill('inpFloorNo', contact.uaeFloor);
    }

    // Flat/Villa no.
    if (contact.uaeFlat) {
      console.log(`[Form] Filling Flat/Villa no.: "${contact.uaeFlat}"...`);
      await this.editAndFill('inpFlatNo', contact.uaeFlat);
    }

    // ── Address Outside UAE ─────────────────────────────────────────────────

    // Country
    if (contact.outsideCountry) {
      const countryName = GdrfaPortalPage.mrzCodeToCountryName(contact.outsideCountry);
      console.log(`[Form] Setting Outside Country: "${countryName}"...`);
      const ocResult = await this.selectByStaticId('cmbApplicantOutsideCountry', countryName);
      if (ocResult.skipped) {
        console.log(`[Skip] Outside Country already set: "${ocResult.matched}".`);
      } else if (ocResult.found) {
        await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
        console.log(`[Form] Outside Country set: "${ocResult.matched}".`);
      } else {
        console.warn(`[Form] Outside Country not found for: "${countryName}".`);
      }
    }

    // Mobile Number (outside UAE)
    if (contact.outsideMobile) {
      console.log(`[Form] Filling Outside Mobile: "${contact.outsideMobile}"...`);
      await this.editAndFill('inpAddressOutsideMobileNumber', contact.outsideMobile);
    }

    // City (outside UAE)
    if (contact.outsideCity) {
      console.log(`[Form] Filling Outside City: "${contact.outsideCity}"...`);
      await this.editAndFill('inpAddressOutsideCity', contact.outsideCity);
    }

    // Address (outside UAE)
    if (contact.outsideAddress) {
      console.log(`[Form] Filling Outside Address: "${contact.outsideAddress}"...`);
      await this.editAndFill('inpAddressOutsideAddress1', contact.outsideAddress);
    }
  }

  // ── Retry Faith selection (before Continue) ────────────────────────────────

  private async retryFaithSelection(faith: string): Promise<void> {
    // Check if Faith is still "-- Select --"
    const needsRetry = await this.page.evaluate(() => {
      const sel = document.querySelector<HTMLSelectElement>('select[data-staticid="cmbApplicantFaith"]');
      return sel ? sel.value === '' || sel.options[sel.selectedIndex]?.text.includes('Select') : false;
    });

    if (!needsRetry) {
      console.log('[Form] Faith already set — no retry needed.');
      return;
    }

    console.log(`[Form] Faith still unset — retrying selection: "${faith}"...`);

    // Scroll to Faith area
    await this.page.evaluate(() => {
      const sel = document.querySelector('select[data-staticid="cmbApplicantFaith"]');
      sel?.scrollIntoView({ block: 'center', behavior: 'instant' });
    });

    // Remove ReadOnly from the Select2 container
    await this.page.evaluate(() => {
      const sel = document.querySelector('select[data-staticid="cmbApplicantFaith"]');
      if (!sel) return;
      const container = sel.previousElementSibling;
      if (container?.classList.contains('select2-container')) {
        container.classList.remove('ReadOnly');
        container.querySelectorAll('.ReadOnly').forEach(el => el.classList.remove('ReadOnly'));
      }
    });

    // Click Select2 choice to open dropdown
    const faithContainer = this.page.locator('[data-staticid="cmbApplicantFaith"]')
      .locator('xpath=preceding-sibling::div[contains(@class,"select2-container")]');
    await faithContainer.locator('.select2-choice').click({ timeout: 5000 });
    await this.page.waitForTimeout(100);

    // Remove ReadOnly from the drop panel
    await this.page.evaluate(() => {
      const drop = document.querySelector('.select2-drop-active');
      if (drop) {
        drop.classList.remove('ReadOnly');
        drop.querySelectorAll('.ReadOnly').forEach(el => el.classList.remove('ReadOnly'));
      }
    });
    await this.page.waitForTimeout(100);

    // Click the matching option
    const faithOption = this.page.locator('.select2-drop-active .select2-results li').filter({ hasText: faith }).first();
    if (await faithOption.isVisible({ timeout: 5000 }).catch(() => false)) {
      const matchedText = await faithOption.textContent() ?? '';
      await faithOption.click();
      await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
      console.log(`[Form] Faith retry set: "${matchedText.trim()}".`);
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

  /**
   * Checks all required fields on the form. If any are empty or still "-- Select --",
   * retries filling them from the application data. Runs up to 2 retry passes.
   */
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

    // Final check — log remaining empty fields but don't block
    const remaining = await this.getEmptyRequiredFields(app);
    if (remaining.length > 0) {
      console.warn(`[Validate] WARNING: ${remaining.length} field(s) still empty after retries: ${remaining.map(f => f.name).join(', ')}`);
    } else {
      console.log('[Validate] All required fields are filled after retries.');
    }
  }

  /**
   * Reads the current state of all required form fields and returns
   * the ones that are empty/unset, along with a retry function for each.
   */
  private async getEmptyRequiredFields(
    app: VisaApplication,
  ): Promise<Array<{ name: string; retry: () => Promise<unknown> }>> {
    const empty: Array<{ name: string; retry: () => Promise<unknown> }> = [];

    // Helper: check if a text input (by data-staticid) has a value
    const inputHasValue = async (staticId: string): Promise<boolean> => {
      const val = await this.page.evaluate((id: string) => {
        const el = document.querySelector<HTMLInputElement>(`input[data-staticid="${id}"]`);
        return el?.value?.trim() ?? '';
      }, staticId);
      return val.length > 0;
    };

    // Helper: check if a <select> (by data-staticid) has a real selection (not "-- Select --" or empty)
    const selectHasValue = async (staticId: string): Promise<boolean> => {
      const val = await this.page.evaluate((id: string) => {
        const sel = document.querySelector<HTMLSelectElement>(`select[data-staticid="${id}"]`);
        if (!sel) return '';
        const text = sel.options[sel.selectedIndex]?.text?.trim() ?? '';
        return text;
      }, staticId);
      return val.length > 0 && !val.includes('Select');
    };

    // Helper: check if a select by label has a real selection
    const selectByLabelHasValue = async (labelText: string): Promise<boolean> => {
      const val = await this.page.evaluate((label: string) => {
        const labels = Array.from(document.querySelectorAll('label'));
        const lbl = labels.find(l => l.textContent?.trim() === label);
        if (!lbl) return '';
        const forId = lbl.getAttribute('for') || lbl.id?.replace(/lbl/i, '');
        const sel = document.querySelector<HTMLSelectElement>(`select[id*="${forId}"]`)
          || lbl.closest('.ThemeGrid_Width6')?.parentElement?.querySelector('select');
        if (!sel) return '';
        return sel.options[sel.selectedIndex]?.text?.trim() ?? '';
      }, labelText);
      return val.length > 0 && !val.includes('Select');
    };

    // ── Passport Names ────────────────────────────────────────────────────
    if (app.passport.firstName && !await inputHasValue('inpFirsttNameEn')) {
      empty.push({ name: 'First Name', retry: () => this.editAndFill('inpFirsttNameEn', app.passport.firstName) });
    }
    if (app.passport.lastName && !await inputHasValue('inpLastNameEn')) {
      empty.push({ name: 'Last Name', retry: () => this.editAndFill('inpLastNameEn', app.passport.lastName) });
    }

    // ── Passport Details ──────────────────────────────────────────────────
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

    // ── Passport Selects ──────────────────────────────────────────────────
    if (app.passport.gender && !await selectByLabelHasValue('Gender')) {
      empty.push({ name: 'Gender', retry: async () => { await this.selectByLabel('Gender', app.passport.gender); } });
    }
    if (app.passport.birthCountry && !await selectByLabelHasValue('Birth Country')) {
      empty.push({ name: 'Birth Country', retry: async () => { await this.selectByLabel('Birth Country', app.passport.birthCountry); } });
    }
    if (app.passport.passportIssueCountry && !await selectByLabelHasValue('Passport Issue Country')) {
      empty.push({ name: 'Passport Issue Country', retry: async () => { await this.selectByLabel('Passport Issue Country', app.passport.passportIssueCountry); } });
    }

    // ── Applicant Details ─────────────────────────────────────────────────
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

    // ── Contact Details ───────────────────────────────────────────────────
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

  /**
   * Waits for the OutSystems Feedback_AjaxWait loader to disappear.
   * The portal shows this spinner during AJAX operations and page transitions.
   */
  private async waitForLoaderToDisappear(): Promise<void> {
    const loader = this.page.locator('div.Feedback_AjaxWait:visible');
    try {
      // If a loader is currently visible, wait for it to hide
      if (await loader.isVisible().catch(() => false)) {
        console.log('[Loader] Waiting for loader to disappear...');
        await loader.waitFor({ state: 'hidden', timeout: 30000 });
        console.log('[Loader] Loader gone.');
      }
    } catch {
      // Loader may have already disappeared
    }
  }

  private async clickContinue(): Promise<void> {
    console.log('[Form] Clicking Continue...');
    const btn = this.page.locator('input[staticid="SmartChannels_EntryPermitNewTourism_btnContinue"]');
    await btn.scrollIntoViewIfNeeded();
    await btn.click({ timeout: 10000 });
    console.log('[Form] Continue clicked — waiting for popup or next page...');

    // After clicking, the portal loads ExistingApplicationConfirmation_PopUp.aspx
    // inside an IFRAME (URL stays on EntryPermitTourism.aspx).
    // Wait for an iframe containing the popup to appear.
    const popupFrame = await this.findPopupFrame(30000);

    if (popupFrame) {
      console.log('[Form] Existing application popup detected (in iframe).');
      await this.handleExistingApplicationPopup(popupFrame);
    } else {
      // No popup — might have gone directly to next page
      console.log('[Form] No popup — proceeding...');
    }

  }


  /**
   * Searches all frames on the page for the popup container.
   * The popup loads inside an iframe as ExistingApplicationConfirmation_PopUp.aspx.
   * Returns the Frame if found, or null if not found within timeout.
   */
  private async findPopupFrame(timeout: number): Promise<import('@playwright/test').Frame | null> {
    const deadline = Date.now() + timeout;

    while (Date.now() < deadline) {
      // Check all frames (main + child iframes)
      for (const frame of this.page.frames()) {
        try {
          // Look for the popup's distinctive element inside each frame
          const el = await frame.$('div.MainPopup');
          if (el) {
            console.log(`[Form] Found popup in frame: ${frame.url()}`);
            return frame;
          }
        } catch {
          // Frame may be navigating, skip
        }
      }
      await this.page.waitForTimeout(250);
    }
    return null;
  }

  /**
   * Handles the "Existing Application Details" popup loaded inside an iframe.
   * Extracts data, pauses for review, then clicks Continue inside the iframe.
   */
  private async handleExistingApplicationPopup(frame: import('@playwright/test').Frame): Promise<void> {
    console.log('[Form] On "Existing Application Details" popup!');

    // ── Data extraction from the iframe ──────────────────────────────────────
    const boldSpans = frame.locator('div.MainPopup span.Bold');
    const count     = await boldSpans.count();
    const values: string[] = [];
    for (let i = 0; i < count; i++) {
      const text = await boldSpans.nth(i).textContent() ?? '';
      values.push(text.replace(/\u00a0/g, ' ').trim());
    }

    const popupData = {
      applicationNumber: values[0] ?? '',
      applicantName:     values[1] ?? '',
      nationality:       values[2] ?? '',
      passportNo:        values[3] ?? '',
      sponsorName:       values[4] ?? '',
      createdDate:       values[5] ?? '',
    };

    console.log('[Form] ── Existing Application Details ──');
    console.log(`  Application Number : ${popupData.applicationNumber}`);
    console.log(`  Applicant Name     : ${popupData.applicantName}`);
    console.log(`  Nationality        : ${popupData.nationality}`);
    console.log(`  Passport No        : ${popupData.passportNo}`);
    console.log(`  Sponsor Name       : ${popupData.sponsorName}`);
    console.log(`  Created Date       : ${popupData.createdDate}`);
    console.log('[Form] ────────────────────────────────────');

    // ── Click Continue inside the iframe (once) ────────────────────────────────
    const closeBtn = frame.locator('input[staticid="CommonTh_ExistingApplicationConfirmationPopUp_btnCancel"]');
    const popupContinueBtn = closeBtn.locator('xpath=following-sibling::input[@value="Continue"]');

    // Ensure button is ready before clicking
    await popupContinueBtn.waitFor({ state: 'visible', timeout: 10000 }).catch(() => {});
    await popupContinueBtn.click({ force: true, timeout: 10000 });
    console.log('[Form] Clicked Continue on popup.');

    // Wait for the popup iframe to close and page to settle
    await this.page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await this.waitForLoaderToDisappear();
    console.log('[Form] Popup dismissed.');
  }

  // ── SmartInput helpers ─────────────────────────────────────────────────────

  /**
   * Clicks the pencil icon to switch a SmartInput field from ReadOnly to edit mode,
   * scrolls it into view first, then fills it.  Falls back to JS if the pencil fails.
   */
  /**
   * Returns true if the field was actually filled, false if skipped (already has correct value).
   */
  private async editAndFill(staticId: string, value: string): Promise<boolean> {
    // Check if the field already has the correct value — skip if so
    const currentVal = await this.page.evaluate((id: string) => {
      const el = document.querySelector<HTMLInputElement>(`input[data-staticid="${id}"]`);
      return el?.value?.trim() ?? '';
    }, staticId);
    if (currentVal !== '' && currentVal.toUpperCase() === value.trim().toUpperCase()) {
      console.log(`[Skip] "${staticId}" already has correct value: "${currentVal}".`);
      return false;
    }

    // Scroll the field into view first
    await this.page.evaluate((id: string) => {
      const input = document.querySelector<HTMLInputElement>(`input[data-staticid="${id}"]`);
      if (input) input.scrollIntoView({ block: 'center', behavior: 'instant' });
    }, staticId);

    // Use Playwright's native click on the pencil (sends real pointer events, more reliable
    // than JS element.click()). XPath walks up to the nearest ancestor containing the pencil,
    // then back down to the pencil element itself.
    const pencil = this.page
      .locator(`input[data-staticid="${staticId}"]`)
      .locator('xpath=ancestor::*[.//*[contains(@class,"FormEditPencil")]][1]//*[contains(@class,"FormEditPencil")]')
      .first();

    const pencilClicked = await pencil.click({ timeout: 3000 }).then(() => true).catch(() => false);

    if (!pencilClicked) {
      // Fallback: JS DOM traversal click (in case XPath doesn't resolve)
      await this.page.evaluate((id: string) => {
        const input = document.querySelector<HTMLInputElement>(`input[data-staticid="${id}"]`);
        if (!input) return;
        let el: Element | null = input.parentElement;
        while (el && el !== document.body) {
          const p = el.querySelector<HTMLElement>('.FormEditPencil');
          if (p) { p.click(); return; }
          el = el.parentElement;
        }
      }, staticId);
    }

    const field = this.page.locator(`input[data-staticid="${staticId}"]`);
    const visible = await field.waitFor({ state: 'visible', timeout: 8000 }).then(() => true).catch(() => false);

    if (visible) {
      await field.clear();
      await field.fill(value);
    } else {
      console.warn(`[Form] Pencil mode failed for "${staticId}" — using JS value fallback.`);
      await this.page.evaluate((args: { id: string; val: string }) => {
        const el = document.querySelector<HTMLInputElement>(`input[data-staticid="${args.id}"]`);
        if (!el) return;
        el.classList.remove('ReadOnly');
        el.removeAttribute('readonly');
        el.value = '';
        el.dispatchEvent(new Event('input',  { bubbles: true }));
        el.value = args.val;
        el.dispatchEvent(new Event('input',  { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
      }, { id: staticId, val: value });
    }
    return true;
  }

  /** Clears a ReadOnly Arabic SmartInput field via JS (CSS-hidden, no pencil available). */
  private async clearArField(staticId: string): Promise<void> {
    await this.page.evaluate((id: string) => {
      const el = document.querySelector<HTMLInputElement>(`input[data-staticid="${id}"]`);
      if (el) { el.value = ''; el.dispatchEvent(new Event('input', { bubbles: true })); }
    }, staticId);
  }

  /**
   * Finds a native <select> by its visible label text and sets its value by
   * searching option text.  More reliable than data-staticid because the
   * label → htmlFor → <select> chain is always present regardless of staticid.
   */
  /**
   * Returns { found, matched, skipped } — skipped=true means the dropdown already had the correct value.
   */
  private async selectByLabel(
    labelText: string,
    searchValue: string
  ): Promise<{ found: boolean; matched: string; skipped?: boolean }> {
    return this.page.evaluate(
      ({ label, search }: { label: string; search: string }) => {
        // Skip Select2's auto-generated offscreen labels
        const lbl = Array.from(document.querySelectorAll<HTMLLabelElement>('label')).find(
          l => !l.classList.contains('select2-offscreen') &&
               l.textContent?.trim().toLowerCase() === label.toLowerCase()
        );
        if (!lbl || !lbl.htmlFor) return { found: false, matched: '' };

        const el  = document.getElementById(lbl.htmlFor);
        const sel = el instanceof HTMLSelectElement
          ? el
          : el?.parentElement?.querySelector<HTMLSelectElement>('select') ?? null;
        if (!sel) return { found: false, matched: '' };

        // Check if already has the correct value — skip if so
        const currentText = sel.options[sel.selectedIndex]?.text?.trim() ?? '';
        if (currentText && !currentText.includes('Select') &&
            (currentText.toUpperCase() === search.toUpperCase() ||
             currentText.toUpperCase().includes(search.toUpperCase()))) {
          return { found: true, matched: currentText, skipped: true };
        }

        // Unlock SmartInput ReadOnly — check ancestor, preceding sibling, next sibling, and by Select2 ID convention
        const s2 = sel.closest<HTMLElement>('.select2-container')
          ?? (sel.previousElementSibling?.classList.contains('select2-container') ? sel.previousElementSibling as HTMLElement : null)
          ?? (sel.nextElementSibling?.classList.contains('select2-container') ? sel.nextElementSibling as HTMLElement : null)
          ?? document.getElementById('s2id_' + sel.id) as HTMLElement | null;
        if (s2) {
          s2.classList.remove('ReadOnly');
          s2.querySelectorAll('.ReadOnly').forEach(el => el.classList.remove('ReadOnly'));
        }
        sel.removeAttribute('disabled');

        const opts = Array.from(sel.options);
        // Try exact match first to avoid substring false positives (e.g. "Male" inside "Female")
        const match =
          opts.find(o => o.text.trim().toUpperCase() === search.toUpperCase()) ??
          opts.find(o => o.text.toUpperCase().includes(search.toUpperCase()));
        if (!match) return { found: false, matched: '' };

        sel.value = '';
        sel.value = match.value;
        sel.dispatchEvent(new Event('change', { bubbles: true }));

        // Also trigger jQuery change + update Select2 display text
        const $ = (window as any).jQuery || (window as any).$;
        if ($) {
          $(sel).val(match.value).trigger('change');
        }
        if (s2) {
          const chosen = s2.querySelector('.select2-chosen');
          if (chosen) chosen.textContent = match.text.trim();
        }

        return { found: true, matched: match.text };
      },
      { label: labelText, search: searchValue }
    );
  }

  /**
   * Finds a native <select> by its data-staticid attribute and sets its value.
   * Avoids label-text collisions when multiple sections share the same label
   * (e.g. "Emirate" appears in both Host/Submitter and Contact Details).
   */
  /**
   * Returns { found, matched, skipped } — skipped=true means the dropdown already had the correct value.
   */
  private async selectByStaticId(
    staticId: string,
    searchValue: string
  ): Promise<{ found: boolean; matched: string; skipped?: boolean }> {
    return this.page.evaluate(
      ({ id, search }: { id: string; search: string }) => {
        const sel = document.querySelector<HTMLSelectElement>(`select[data-staticid="${id}"]`);
        if (!sel) return { found: false, matched: '' };

        // Check if already has the correct value — skip if so
        const currentText = sel.options[sel.selectedIndex]?.text?.trim() ?? '';
        if (currentText && !currentText.includes('Select') &&
            (currentText.toUpperCase() === search.toUpperCase() ||
             currentText.toUpperCase().includes(search.toUpperCase()))) {
          return { found: true, matched: currentText, skipped: true };
        }

        // Unlock Select2 ReadOnly — check ancestor, preceding sibling, next sibling, and by Select2 ID convention
        const s2 = sel.closest<HTMLElement>('.select2-container')
          ?? (sel.previousElementSibling?.classList.contains('select2-container') ? sel.previousElementSibling as HTMLElement : null)
          ?? (sel.nextElementSibling?.classList.contains('select2-container') ? sel.nextElementSibling as HTMLElement : null)
          ?? document.getElementById('s2id_' + sel.id) as HTMLElement | null;
        if (s2) {
          s2.classList.remove('ReadOnly');
          s2.querySelectorAll('.ReadOnly').forEach(el => el.classList.remove('ReadOnly'));
        }
        sel.removeAttribute('disabled');

        const opts = Array.from(sel.options);
        const match =
          opts.find(o => o.text.trim().toUpperCase() === search.toUpperCase()) ??
          opts.find(o => o.text.toUpperCase().includes(search.toUpperCase()));
        if (!match) return { found: false, matched: '' };

        sel.value = '';
        sel.value = match.value;
        sel.dispatchEvent(new Event('change', { bubbles: true }));

        // Also trigger jQuery change + update Select2 display text
        const $ = (window as any).jQuery || (window as any).$;
        if ($) {
          $(sel).val(match.value).trigger('change');
        }
        // Update the Select2 display text directly
        if (s2) {
          const chosen = s2.querySelector('.select2-chosen');
          if (chosen) chosen.textContent = match.text.trim();
        }

        return { found: true, matched: match.text };
      },
      { id: staticId, search: searchValue }
    );
  }

  /**
   * Handles AJAX-powered Select2 dropdowns where options are fetched on search.
   * Finds the container by matching the idFragment against element IDs (e.g.
   * "ComingFromCountry" matches "s2id_...wtcmbComingFromCountry").
   * Opens the dropdown, types the search term, waits for results, and clicks
   * the matching option. Returns the matched text or empty string on failure.
   */
  private async selectByAjaxSelect2(
    idFragment: string,
    searchValue: string
  ): Promise<string> {
    // Find the Select2 container by ID fragment and remove ReadOnly
    const containerId = await this.page.evaluate((frag: string) => {
      const containers = Array.from(document.querySelectorAll<HTMLElement>('.select2-container'));
      const match = containers.find(el => el.id.toLowerCase().includes(frag.toLowerCase()));
      if (!match) return '';
      match.classList.remove('ReadOnly');
      match.querySelectorAll('.ReadOnly').forEach(el => el.classList.remove('ReadOnly'));
      return match.id;
    }, idFragment);

    if (!containerId) {
      console.warn(`[Form] AJAX Select2 container not found for fragment: "${idFragment}".`);
      return '';
    }

    // Click the Select2 choice to open the dropdown
    // Use attribute selector to avoid CSS escaping issues with long OutSystems IDs
    const container = this.page.locator(`[id="${containerId}"]`);
    await container.scrollIntoViewIfNeeded();
    await container.locator('.select2-choice').click({ timeout: 5000 });
    await this.page.waitForTimeout(150);

    // Click the pencil icon inside the search box to unlock the SmartInput
    await this.page.evaluate(() => {
      const drop = document.querySelector('.select2-drop-active');
      if (!drop) return;
      // Remove ReadOnly from the drop panel and all children
      drop.classList.remove('ReadOnly');
      drop.querySelectorAll('.ReadOnly').forEach(el => el.classList.remove('ReadOnly'));
      // Click the pencil if present
      const pencil = drop.querySelector<HTMLElement>('.FormEditPencil');
      if (pencil) pencil.click();
      // Also force the search input to be editable
      const input = drop.querySelector<HTMLInputElement>('.select2-input');
      if (input) {
        input.removeAttribute('readonly');
        input.removeAttribute('disabled');
        input.classList.remove('ReadOnly');
      }
    });
    await this.page.waitForTimeout(150);

    // Type using keyboard (bypasses Playwright visibility check on SmartInput fields)
    await this.page.keyboard.type(searchValue, { delay: 30 });

    // Wait for results to appear
    await this.page.waitForFunction(() => {
      const results = document.querySelectorAll('.select2-drop-active .select2-results li.select2-result');
      return results.length > 0;
    }, { timeout: 10000 }).catch(() => {});

    // Click the first matching result
    const firstResult = this.page.locator('.select2-drop-active .select2-results li.select2-result').first();
    if (await firstResult.isVisible({ timeout: 3000 }).catch(() => false)) {
      const matchedText = await firstResult.textContent() ?? '';
      await firstResult.click();
      await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
      return matchedText.trim();
    }

    return '';
  }

  // ── Session Keep-Alive ─────────────────────────────────────────────────────

  /**
   * Fires a lightweight HEAD request every intervalMs to reset the server-side
   * 15-min idle timer.  Returns a stop function to cancel when done.
   */
  private startSessionKeepAlive(intervalMs = 10 * 60 * 1000): () => void {
    console.log(`[KeepAlive] Session keep-alive started (every ${intervalMs / 60000} min).`);
    const timer = setInterval(async () => {
      try {
        await this.page.evaluate(async () => {
          await fetch(window.location.href, { method: 'HEAD', credentials: 'include' });
        });
        console.log('[KeepAlive] Session ping sent.');
      } catch { /* page may be mid-navigation */ }
    }, intervalMs);
    return () => { clearInterval(timer); console.log('[KeepAlive] Stopped.'); };
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

/** Thin wrapper so tests can call fillApplicationForm(page, app) unchanged. */
export async function fillApplicationForm(page: Page, application: VisaApplication): Promise<void> {
  await new GdrfaPortalPage(page).fillApplicationForm(application);
}

export async function verifySession(page: Page): Promise<void> {
  await new GdrfaPortalPage(page).verifySession();
}
