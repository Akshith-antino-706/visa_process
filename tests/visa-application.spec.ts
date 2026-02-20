import { test } from '@playwright/test';
import * as fs   from 'fs';
import * as path from 'path';
import { fillApplicationForm } from '../src/automation/gdrfa-portal';
import { readApplicationsFromExcel } from '../src/utils/excel-reader';

const SESSION_FILE = path.resolve('auth/session.json');
const EXCEL_FILE   = path.resolve('data/applications/applications.xlsx');

// ─── Helpers ──────────────────────────────────────────────────────────────────

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
  if (!fs.existsSync(EXCEL_FILE)) {
    throw new Error(
      `Excel file not found: ${EXCEL_FILE}\n` +
      'Run "npm run excel" to generate the applications spreadsheet first.'
    );
  }
});

// ─── Tests run in serial order ────────────────────────────────────────────────

test.describe.serial('GDRFA Visa Application — Fill Only (No Submission)', () => {

  const applications = readApplicationsFromExcel(EXCEL_FILE);
  const total        = applications.length;

  if (total === 0) {
    test('no applicants found', () => {
      throw new Error(
        'No rows found in the Excel file.\n' +
        'Add applicant data to: ' + EXCEL_FILE
      );
    });
  }

  for (let i = 0; i < applications.length; i++) {
    const application = applications[i];
    const label       = `applicant-${String(i + 1).padStart(String(total).length, '0')}`;
    const prefix      = `[Applicant ${i + 1}/${total} — ${label}]`;

    test(`${label} — Fill form (no submission)`, async ({ page }) => {
      console.log(`\n${prefix} Loading application...`);
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
