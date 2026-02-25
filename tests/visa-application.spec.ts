import { test } from '@playwright/test';
import * as fs   from 'fs';
import * as path from 'path';
import { fillApplicationForm } from '../src/automation/gdrfa-portal';
import { readApplicationsFromExcel } from '../src/utils/excel-reader';

const SESSION_FILE = path.resolve('auth/session.json');
const EXCEL_FILE   = path.resolve('data/applications/applications.xlsx');

test.use({ storageState: SESSION_FILE });

// ─── Helpers ──────────────────────────────────────────────────────────────────

function ensureDir(dir: string) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

console.log('[Init] Test file loaded');
console.log(`[Init] CWD: ${process.cwd()}`);
console.log(`[Init] Session: ${SESSION_FILE} (exists: ${fs.existsSync(SESSION_FILE)})`);
console.log(`[Init] Excel:   ${EXCEL_FILE} (exists: ${fs.existsSync(EXCEL_FILE)})`);

// ─── Pre-flight checks ────────────────────────────────────────────────────────

test.beforeAll(() => {
  console.log('[beforeAll] Running pre-flight checks...');
  if (!fs.existsSync(SESSION_FILE)) {
    console.error('[beforeAll] FAIL: Session file missing');
    throw new Error(
      'Session not found: auth/session.json\n' +
      'Run "npm run auth" to log in manually and save your session first.'
    );
  }
  console.log('[beforeAll] Session file OK');

  if (!fs.existsSync(EXCEL_FILE)) {
    console.error('[beforeAll] FAIL: Excel file missing');
    throw new Error(
      `Excel file not found: ${EXCEL_FILE}\n` +
      'Run "npm run excel" to generate the applications spreadsheet first.'
    );
  }
  console.log('[beforeAll] Excel file OK');
  console.log('[beforeAll] Pre-flight checks passed');
});

// ─── Tests run in serial order ────────────────────────────────────────────────

test.describe('GDRFA Visa Application — Fill Only (No Submission)', () => {

  console.log('[Init] Reading Excel...');
  const applications = readApplicationsFromExcel(EXCEL_FILE);
  const total        = applications.length;
  console.log(`[Init] Found ${total} applicant(s)`);

  if (total === 0) {
    test('no applicants found', () => {
      throw new Error(
        'No rows found in the Excel file.\n' +
        'Add applicant data to: ' + EXCEL_FILE
      );
    });
  }

  test.beforeEach(({ }, testInfo) => {
    console.log(`\n[beforeEach] Starting: ${testInfo.title}`);
    console.log(`[beforeEach] Timeout: ${testInfo.timeout}ms`);
    console.log(`[beforeEach] Project: ${testInfo.project.name}`);
  });

  test.afterEach(({ }, testInfo) => {
    const status = testInfo.status ?? 'unknown';
    const duration = testInfo.duration;
    console.log(`[afterEach] Finished: ${testInfo.title} — ${status} (${duration}ms)`);
    if (testInfo.error) {
      console.error(`[afterEach] Error: ${testInfo.error.message}`);
    }
  });

  for (let i = 0; i < applications.length; i++) {
    const application = applications[i];
    const label       = `applicant-${String(i + 1).padStart(String(total).length, '0')}`;
    const prefix      = `[Applicant ${i + 1}/${total} — ${label}]`;

    test(`${label} — Fill form (no submission)`, async ({ page }) => {
      console.log(`\n${prefix} Loading application...`);
      console.log(`${prefix} Browser: ${page.context().browser()?.browserType().name() ?? 'unknown'}`);
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
