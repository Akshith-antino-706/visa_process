import { Page } from '@playwright/test';
import * as path from 'path';
import * as fs from 'fs';
import {
  VisaApplication,
  PassportDetails,
  ApplicantDetails,
  ApplicationDocuments,
} from '../types/application-data';

// ─── Page Object ──────────────────────────────────────────────────────────────

export class GdrfaPortalPage {
  private static readonly HOME   = 'https://smart.gdrfad.gov.ae/SmartChannels_Th/';
  private static readonly UPLOAD = 'https://smart.gdrfad.gov.ae/SmartChannels/Application_UploadDocuments.aspx';

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

      console.log('[Form] Waiting 10 seconds...');
      await this.page.waitForTimeout(10000);
      console.log('\n[Flow] ─── Steps complete. ───\n');
    } finally {
      stopKeepAlive();
    }
  }

  async uploadDocuments(docs: ApplicationDocuments): Promise<void> {
    console.log('[Upload] Switching to Attachments tab...');
    const attachTab = this.page.locator('a:has-text("Attachments"), li:has-text("Attachments") > a').first();
    if (await attachTab.isVisible({ timeout: 5000 }).catch(() => false)) {
      await attachTab.click();
      await this.page.waitForLoadState('networkidle');
    } else {
      await this.page.goto(GdrfaPortalPage.UPLOAD, { waitUntil: 'networkidle' });
    }

    const slots: Array<{ label: string; file: string }> = [
      { label: 'Hotel reservation',           file: docs.hotelReservationPage1 },
      { label: 'Passport External',            file: docs.passportExternalCoverPage },
      { label: 'Personal Photo',               file: docs.personalPhoto },
      { label: 'Return air ticket',            file: docs.returnAirTicketPage1 },
      { label: 'Sponsored Passport',           file: docs.sponsoredPassportPage1 },
      { label: 'Hotel reservation Page 2',     file: docs.hotelReservationPage2 ?? '' },
      { label: 'Others Page 1',                file: docs.othersPage1 ?? '' },
    ];

    const allFileInputs = await this.page.locator('input[type="file"]').all();
    let inputIdx = 0;
    for (const slot of slots) {
      if (!slot.file) continue;
      const clean = this.ensureCleanFileName(slot.file);
      const input = allFileInputs[inputIdx];
      if (input) {
        await input.setInputFiles(clean);
        console.log(`[Upload] "${slot.label}": ${path.basename(clean)}`);
        await this.page.waitForTimeout(1500);
      }
      inputIdx++;
    }

    try {
      await this.page.waitForSelector('text=ready to pay', { timeout: 10000, state: 'visible' });
      console.log('[Upload] Status: READY TO PAY');
    } catch {
      console.log('[Upload] Documents uploaded (status badge check timed out).');
    }
  }

  // ── Navigation ─────────────────────────────────────────────────────────────

  private async dismissPromoPopup(): Promise<void> {
    const SKIP_ID = 'WebPatterns_wt2_block_wtMainContent_wt3_EmaratechSG_Patterns_wt8_block_wtMainContent_wt10';
    try {
      const skipBtn = this.page.frameLocator('iframe').locator(`#${SKIP_ID}, input[value="Skip"]`).first();
      if (!await skipBtn.isVisible({ timeout: 15000 }).catch(() => false)) return;
      await skipBtn.click();
      console.log('[Nav] Dismissed promotional popup.');
      await this.page.waitForTimeout(600);
    } catch { /* non-fatal — popup does not appear on every load */ }
  }

  private async navigateToNewApplication(): Promise<void> {
    console.log('[Nav] Navigating to Existing Applications...');
    await this.dismissPromoPopup();

    await this.page.locator(
      '#EmaratechSG_Theme_wtwbLayoutEmaratech_block_wtMainContent_wtwbDashboard_wtCntExistApp, ' +
      'a:has-text("Existing Applications")'
    ).first().click({ timeout: 15000 });

    await this.page.waitForLoadState('networkidle', { timeout: 15000 }).catch(() => {});
    await this.dismissPromoPopup();

    const dropdown = this.page.locator(
      '#EmaratechSG_Theme_wtwbLayoutEmaratechWithoutTitle_block_wtMainContent_EmaratechSG_Patterns_wtwbEstbButtonWithContextInfo_block_wtIcon_wtcntContextActionBtn'
    );
    await dropdown.waitFor({ state: 'visible', timeout: 15000 });
    await dropdown.click();
    await this.page.waitForTimeout(1000);

    const firstOption = this.page.locator(
      '#EmaratechSG_Theme_wtwbLayoutEmaratechWithoutTitle_block_wtMainContent_EmaratechSG_Patterns_wtwbEstbButtonWithContextInfo_block_wtContent_wtwbEstbTopServices_wtListMyServicesExperiences_ctl00_wtStartTopService'
    );
    await firstOption.waitFor({ state: 'visible', timeout: 10000 });
    console.log(`[Nav] Selecting form: "${(await firstOption.textContent())?.trim()}"`);
    await firstOption.click();

    await this.page.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {});
    console.log('[Nav] Form page loaded. URL:', this.page.url());
    await this.page.waitForTimeout(10000);
    console.log('[Nav] Wait complete.');
  }

  private async waitForPageSettle(): Promise<void> {
    await this.page.waitForLoadState('domcontentloaded', { timeout: 15000 }).catch(() => {});
    await this.page.waitForLoadState('networkidle',      { timeout: 15000 }).catch(() => {});
  }

  // ── Passport header (Passport Type → Nationality → Search Data) ────────────

  private async setVisitReason(): Promise<void> {
    console.log('[Form] Selecting Visit Reason → Tourism...');
    await this.page.waitForSelector('select[data-staticid="cmbVisitReason"]', { timeout: 15000 });
    const set = await this.page.evaluate(() => {
      const sel = document.querySelector<HTMLSelectElement>('select[data-staticid="cmbVisitReason"]');
      if (!sel) return false;
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
    await input.waitFor({ state: 'visible', timeout: 15000 });
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
    await this.page.waitForTimeout(2000); // Extra buffer for DOM re-render
    await this.page.locator('input[data-staticid="inpFirsttNameEn"]').waitFor({ state: 'attached', timeout: 20000 });
    console.log('[Form] Search Data complete — portal fields populated.');

    // ── Manual review pause ───────────────────────────────────────────────────
    // Execution stops here. Review the portal fields, then click ▶ Resume in the
    // Playwright Inspector (or press F8) to continue filling the form.
    console.log('[Form] ⏸  Paused — review the form and click Resume in the Playwright Inspector.');
    await this.page.pause();
    console.log('[Form] ▶  Resumed — continuing form fill...');
  }

  // ── Passport name fields (First / Middle / Last) ───────────────────────────

  private async fillPassportNames(passport: PassportDetails): Promise<void> {
    console.log('[Form] Clearing Arabic name fields...');
    await this.clearArField('inpFirstNameAr');
    await this.clearArField('inpMiddleNameAr');
    await this.clearArField('inpLastNameAr');

    console.log(`[Form] Filling First Name: "${passport.firstName}"...`);
    await this.editAndFill('inpFirsttNameEn', passport.firstName);
    await this.page.evaluate(() => (window as any).translateInputText?.('inpFirsttNameEn'));
    await this.page.waitForTimeout(1500);
    console.log('[Form] First Name filled + translated.');

    if (passport.middleName) {
      console.log(`[Form] Filling Middle Name: "${passport.middleName}"...`);
      await this.editAndFill('inpMiddleNameEn', passport.middleName);
      await this.page.evaluate(() => (window as any).translateInputText?.('inpMiddleNameEn'));
      await this.page.waitForTimeout(1500);
      console.log('[Form] Middle Name filled + translated.');
    } else {
      console.log('[Form] No middle name — field left blank.');
    }

    console.log(`[Form] Filling Last Name: "${passport.lastName}"...`);
    await this.editAndFill('inpLastNameEn', passport.lastName);
    await this.page.evaluate(() => (window as any).translateInputText?.('inpLastNameEn'));
    await this.page.waitForTimeout(1500);
    console.log('[Form] Last Name filled + translated.');
  }

  // ── Passport detail fields (below name fields) ─────────────────────────────

  private async fillPassportDetails(passport: PassportDetails): Promise<void> {
    // Date of Birth: clear any pre-filled value first, then unlock via pencil and fill.
    const dob = passport.dateOfBirth.replace(/\//g, '-');  // DD/MM/YYYY → DD-MM-YYYY
    console.log(`[Form] Filling Date of Birth: "${dob}"...`);
    // Clear the existing value before unlocking (bypasses ReadOnly directly)
    await this.page.evaluate(() => {
      const el = document.querySelector<HTMLInputElement>('input[data-staticid="inpDateOfBirth"]');
      if (el) { el.value = ''; el.dispatchEvent(new Event('input', { bubbles: true })); }
    });
    await this.editAndFill('inpDateOfBirth', dob);
    // Fire the AJAX change handler explicitly (datepicker fields need this after programmatic fill)
    await this.page.evaluate(() => {
      const el = document.querySelector<HTMLInputElement>('input[data-staticid="inpDateOfBirth"]');
      if (el) el.dispatchEvent(new Event('change', { bubbles: true }));
    });
    await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
    console.log('[Form] Date of Birth filled.');

    const birthCountry = GdrfaPortalPage.mrzCodeToCountryName(passport.birthCountry);
    console.log(`[Form] Setting Birth Country: "${birthCountry}"...`);
    const bcResult = await this.selectByLabel('Birth Country', birthCountry);
    if (bcResult.found) {
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
    await this.editAndFill('inpApplicantBirthPlaceEn', birthPlace);
    // Click the translate button (same as clicking the 🌐 icon next to the field)
    await this.page.evaluate(() => (window as any).translateInputText('inpApplicantBirthPlaceEn'));
    await this.page.waitForTimeout(1500);
    console.log('[Form] Birth Place EN filled + translated.');

    console.log(`[Form] Setting Gender: "${passport.gender}"...`);
    const gResult = await this.selectByLabel('Gender', passport.gender);
    if (gResult.found) {
      await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
      console.log(`[Form] Gender set: "${gResult.matched}".`);
    } else {
      console.warn(`[Form] Gender not found for: "${passport.gender}".`);
    }

    if (passport.passportIssueCountry) {
      const issueCountry = GdrfaPortalPage.mrzCodeToCountryName(passport.passportIssueCountry);
      console.log(`[Form] Setting Passport Issue Country: "${issueCountry}"...`);
      const icResult = await this.selectByLabel('Passport Issue Country', issueCountry);
      if (icResult.found) {
        await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
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
    } else {
      console.log('[Form] Passport Issue Date — skipped (empty).');
    }

    const expiryDate = passport.passportExpiryDate.replace(/\//g, '-');
    console.log(`[Form] Filling Passport Expiry Date: "${expiryDate}"...`);
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

    if (passport.passportPlaceOfIssueEN) {
      console.log(`[Form] Filling Place of Issue EN: "${passport.passportPlaceOfIssueEN}"...`);
      await this.editAndFill('inpPassportPlaceIssueEn', passport.passportPlaceOfIssueEN);
      await this.page.evaluate(() => (window as any).translateInputText?.('inpPassportPlaceIssueEn'));
      await this.page.waitForTimeout(1500);
      console.log('[Form] Place of Issue EN filled + translated.');
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
      await this.editAndFill('inpMotherNameEn', applicant.motherNameEN);
      await this.page.evaluate(() => (window as any).translateInputText?.('inpMotherNameEn'));
      await this.page.waitForTimeout(1500);
      console.log('[Form] Mother Name EN filled + translated.');
    } else {
      console.log('[Form] Mother Name EN — skipped (empty).');
    }

    // Marital Status
    if (applicant.maritalStatus) {
      console.log(`[Form] Setting Marital Status: "${applicant.maritalStatus}"...`);
      const msResult = await this.selectByLabel('Marital Status', applicant.maritalStatus);
      if (msResult.found) {
        console.log(`[Form] Marital Status set: "${msResult.matched}".`);
      } else {
        console.warn(`[Form] Marital Status not found for: "${applicant.maritalStatus}".`);
      }
    }

    // Religion (AJAX onChange repopulates Faith dropdown)
    if (applicant.religion) {
      console.log(`[Form] Setting Religion: "${applicant.religion}"...`);
      const rResult = await this.selectByLabel('Religion', applicant.religion);
      if (rResult.found) {
        await this.page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
        console.log(`[Form] Religion set: "${rResult.matched}".`);
      } else {
        console.warn(`[Form] Religion not found for: "${applicant.religion}".`);
      }
    }

    // Faith (options depend on Religion — must be set after Religion AJAX resolves)
    if (applicant.faith) {
      console.log(`[Form] Setting Faith: "${applicant.faith}"...`);
      // Wait until the Faith dropdown has been populated (more than just "-- Select --")
      await this.page.waitForFunction(() => {
        const lbl = Array.from(document.querySelectorAll<HTMLLabelElement>('label'))
          .find(l => !l.classList.contains('select2-offscreen') && l.textContent?.trim().toLowerCase() === 'faith');
        if (!lbl || !lbl.htmlFor) return false;
        const sel = document.getElementById(lbl.htmlFor) as HTMLSelectElement | null;
        return sel ? sel.options.length > 1 : false;
      }, { timeout: 10000 }).catch(() => console.warn('[Form] Faith dropdown did not populate in time.'));
      const fResult = await this.selectByLabel('Faith', applicant.faith);
      if (fResult.found) {
        console.log(`[Form] Faith set: "${fResult.matched}".`);
      } else {
        // Log available options to help diagnose the mismatch
        await this.page.evaluate(() => {
          const lbl = Array.from(document.querySelectorAll<HTMLLabelElement>('label'))
            .find(l => !l.classList.contains('select2-offscreen') && l.textContent?.trim().toLowerCase() === 'faith');
          if (!lbl || !lbl.htmlFor) return;
          const sel = document.getElementById(lbl.htmlFor) as HTMLSelectElement | null;
          if (sel) console.warn('[Faith options]', Array.from(sel.options).map(o => o.text).join(' | '));
        });
        console.warn(`[Form] Faith not found for: "${applicant.faith}".`);
      }
    }

    // Education
    if (applicant.education) {
      console.log(`[Form] Setting Education: "${applicant.education}"...`);
      const eResult = await this.selectByLabel('Education', applicant.education);
      if (eResult.found) {
        console.log(`[Form] Education set: "${eResult.matched}".`);
      } else {
        console.warn(`[Form] Education not found for: "${applicant.education}".`);
      }
    }

    // Profession (autocomplete widget — skip if empty)
    if (applicant.profession) {
      console.log(`[Form] Filling Profession: "${applicant.profession}"...`);
      const profInput = this.page
        .locator('input[id*="wtProfessionSerch"]')
        .first();
      if (await profInput.isVisible({ timeout: 3000 }).catch(() => false)) {
        await profInput.fill(applicant.profession);
        await this.page.waitForTimeout(1200);
        const firstSuggestion = this.page.locator('ul[id*="profession"] li, .autocomplete-results li').first();
        if (await firstSuggestion.isVisible({ timeout: 3000 }).catch(() => false)) {
          await firstSuggestion.click();
          console.log('[Form] Profession selected from suggestion.');
        } else {
          console.warn('[Form] No profession suggestion appeared — value left as typed.');
        }
      } else {
        console.warn('[Form] Profession input not found.');
      }
    } else {
      console.log('[Form] Profession — skipped (empty).');
    }

    // Coming From Country
    if (applicant.comingFromCountry) {
      console.log(`[Form] Setting Coming From Country: "${applicant.comingFromCountry}"...`);
      const cfcResult = await this.selectByLabel('Coming From Country', applicant.comingFromCountry);
      if (cfcResult.found) {
        console.log(`[Form] Coming From Country set: "${cfcResult.matched}".`);
      } else {
        console.warn(`[Form] Coming From Country not found for: "${applicant.comingFromCountry}".`);
      }
    } else {
      console.log('[Form] Coming From Country — skipped (empty).');
    }
  }

  // ── SmartInput helpers ─────────────────────────────────────────────────────

  /**
   * Clicks the pencil icon to switch a SmartInput field from ReadOnly to edit mode,
   * scrolls it into view first, then fills it.  Falls back to JS if the pencil fails.
   */
  private async editAndFill(staticId: string, value: string): Promise<void> {
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
  private async selectByLabel(
    labelText: string,
    searchValue: string
  ): Promise<{ found: boolean; matched: string }> {
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

        // Unlock SmartInput ReadOnly — remove ReadOnly class from the Select2 container
        // (same effect as clicking the pencil icon for text fields)
        const s2 = sel.closest<HTMLElement>('.select2-container');
        if (s2) s2.classList.remove('ReadOnly');
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
        return { found: true, matched: match.text };
      },
      { label: labelText, search: searchValue }
    );
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

  // ── Upload helper ──────────────────────────────────────────────────────────

  /** Copies file to a clean name if the original contains forbidden characters. */
  private ensureCleanFileName(filePath: string): string {
    const dir   = path.dirname(filePath);
    const ext   = path.extname(filePath);
    const base  = path.basename(filePath, ext);
    const clean = base.replace(/[\\/:*?"<>|]/g, '-');
    if (clean === base) return filePath;
    const cleanPath = path.join(dir, clean + ext);
    if (!fs.existsSync(cleanPath)) {
      fs.copyFileSync(filePath, cleanPath);
      console.log(`[Upload] Renamed: "${base + ext}" → "${clean + ext}"`);
    }
    return cleanPath;
  }

  // ── Country code → portal display name ────────────────────────────────────

  /** Maps a 3-letter MRZ ISO code to the country name used in GDRFA Select2 dropdowns. */
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
