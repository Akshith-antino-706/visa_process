import { test } from '@playwright/test';
import * as fs   from 'fs';
import * as path from 'path';
import { fillApplicationForm } from '../src/automation/gdrfa-portal';
import { VisaApplication } from '../src/types/application-data';

const SESSION_FILE   = path.resolve('auth/session.json');
const PROCESSED_DIR  = path.resolve('data/applications/processed');

// ─── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Collect all processed application JSONs, sorted alphabetically.
 * These are written by `npm run ocr` — run that first if this throws.
 */
function loadProcessedApplications(): string[] {
  if (!fs.existsSync(PROCESSED_DIR)) {
    throw new Error(
      `Processed applications directory not found: ${PROCESSED_DIR}\n` +
      'Run "npm run ocr" first to extract passport data from images.'
    );
  }
  const files = fs.readdirSync(PROCESSED_DIR)
    .filter(f => f.endsWith('.json'))
    .sort()
    .map(f => path.join(PROCESSED_DIR, f));
  if (files.length === 0) {
    throw new Error(
      `No JSON files found in: ${PROCESSED_DIR}\n` +
      'Run "npm run ocr" first to extract passport data from images.'
    );
  }
  return files;
}

function ensureDir(dir: string) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

// ─── Pre-flight checks ────────────────────────────────────────────────────────

test.beforeAll(() => {
  if (!fs.existsSync(SESSION_FILE)) {
    throw new Error(
      'Session not found: auth/session.json\n' +
      'Run "npm run auth" to log in manually and save your session first.'
    );
  }
});

// ─── Tests run in serial order ────────────────────────────────────────────────

test.describe.serial('GDRFA Visa Application — Fill Only (No Submission)', () => {

  const jsonFiles = loadProcessedApplications();
  const total     = jsonFiles.length;

  for (let i = 0; i < jsonFiles.length; i++) {
    const jsonPath = jsonFiles[i];
    const label    = path.basename(jsonPath, '.json');   // e.g. "applicant-01"
    const prefix   = `[Applicant ${i + 1}/${total} — ${label}]`;

    test(`${label} — Fill form (no submission)`, async ({ page }) => {
      console.log(`\n${prefix} Loading application...`);

      const application = JSON.parse(
        fs.readFileSync(jsonPath, 'utf-8')
      ) as VisaApplication;

      console.log(
        `${prefix} Filling form for: ` +
        `${application.passport.fullNameEN} | ${application.passport.passportNumber}`
      );

      await fillApplicationForm(page, application);

      ensureDir('test-results');
      await page.screenshot({
        path: `test-results/${label}-form-filled.png`,
        fullPage: true,
      });
      console.log(`${prefix} Screenshot saved.`);
    });
  }

});
