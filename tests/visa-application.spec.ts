/**
 * GDRFA Visa Application — Selenium with 20 parallel browser workers.
 *
 * Spawns up to 20 Chrome instances simultaneously, each processing a
 * different applicant from the Excel file. Uses a worker pool pattern
 * to distribute applications across available browser slots.
 *
 * Usage: npm test
 */

import 'dotenv/config';
import { WebDriver } from 'selenium-webdriver';
import * as fs from 'fs';
import * as path from 'path';
import { createDriver, loadSession, quitDriver, SESSION_FILE } from '../src/automation/driver-factory';
import { fillApplicationForm } from '../src/automation/gdrfa-portal';
import { readApplicationsFromExcel } from '../src/utils/excel-reader';
import { VisaApplication } from '../src/types/application-data';

const EXCEL_FILE   = path.resolve('data/applications/applications.xlsx');
const MAX_WORKERS  = parseInt(process.env.WORKERS || '20', 10);  // default 20, override with WORKERS=1

// ─── Helpers ──────────────────────────────────────────────────────────────────

function ensureDir(dir: string) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function padIndex(i: number, total: number): string {
  return String(i + 1).padStart(String(total).length, '0');
}

async function takeScreenshot(driver: WebDriver, filePath: string): Promise<void> {
  try {
    const screenshot = await driver.takeScreenshot();
    fs.writeFileSync(filePath, screenshot, 'base64');
  } catch {
    console.warn(`[Screenshot] Failed to save: ${filePath}`);
  }
}

// ─── Worker function ─────────────────────────────────────────────────────────

async function processApplication(
  application: VisaApplication,
  index: number,
  total: number,
): Promise<{ label: string; status: 'passed' | 'failed'; error?: string; duration: number; applicationNumber?: string }> {
  const label  = `applicant-${padIndex(index, total)}`;
  const prefix = `[Applicant ${index + 1}/${total} — ${label}]`;
  const start  = Date.now();

  let driver: WebDriver | null = null;

  try {
    console.log(`\n${prefix} Starting — creating browser...`);
    driver = await createDriver();
    await loadSession(driver);

    console.log(
      `${prefix} Filling form for: ` +
      `${application.passport.fullNameEN} | ${application.passport.passportNumber}`
    );

    const applicationNumber = await fillApplicationForm(driver, application);

    ensureDir('test-results');
    await takeScreenshot(driver, `test-results/${label}-complete.png`);
    console.log(`${prefix} Screenshot saved.`);

    const duration = Date.now() - start;
    console.log(`${prefix} PASSED — Application #: ${applicationNumber || 'N/A'} (${(duration / 1000).toFixed(1)}s)`);
    return { label, status: 'passed', duration, applicationNumber };
  } catch (err: any) {
    const duration = Date.now() - start;
    const errorMsg = err?.message ?? String(err);
    console.error(`${prefix} FAILED: ${errorMsg}`);

    // Save error screenshot
    if (driver) {
      ensureDir('test-results');
      await takeScreenshot(driver, `test-results/${label}-error.png`);
    }

    return { label, status: 'failed', error: errorMsg, duration };
  } finally {
    if (driver) {
      await quitDriver(driver);
    }
  }
}

// ─── Worker Pool ─────────────────────────────────────────────────────────────

async function runWorkerPool(
  applications: VisaApplication[],
  maxWorkers: number
): Promise<Array<{ label: string; status: 'passed' | 'failed'; error?: string; duration: number; applicationNumber?: string }>> {
  const results: Array<{ label: string; status: 'passed' | 'failed'; error?: string; duration: number; applicationNumber?: string }> = [];
  let nextIndex = 0;

  const workers: Promise<void>[] = [];

  async function worker(): Promise<void> {
    while (nextIndex < applications.length) {
      const idx = nextIndex++;
      const result = await processApplication(applications[idx], idx, applications.length);
      results.push(result);
    }
  }

  // Spawn up to maxWorkers concurrent workers
  const workerCount = Math.min(maxWorkers, applications.length);
  console.log(`\n[Pool] Spawning ${workerCount} parallel workers for ${applications.length} application(s)...\n`);

  for (let i = 0; i < workerCount; i++) {
    workers.push(worker());
  }

  await Promise.all(workers);
  return results;
}

// ─── Main ────────────────────────────────────────────────────────────────────

async function main() {
  console.log('[Init] ─── GDRFA Visa Application Automation (Selenium) ───');
  console.log(`[Init] CWD: ${process.cwd()}`);
  console.log(`[Init] Session: ${SESSION_FILE} (exists: ${fs.existsSync(SESSION_FILE)})`);
  console.log(`[Init] Excel:   ${EXCEL_FILE} (exists: ${fs.existsSync(EXCEL_FILE)})`);
  console.log(`[Init] Max parallel workers: ${MAX_WORKERS}`);

  // Pre-flight checks
  if (!fs.existsSync(SESSION_FILE)) {
    console.error('[FAIL] Session file missing');
    throw new Error(
      'Session not found: auth/session.json\n' +
      'Run "npm run auth" to log in manually and save your session first.'
    );
  }

  if (!fs.existsSync(EXCEL_FILE)) {
    console.error('[FAIL] Excel file missing');
    throw new Error(
      `Excel file not found: ${EXCEL_FILE}\n` +
      'Run "npm run excel" to generate the applications spreadsheet first.'
    );
  }

  console.log('[Init] Reading Excel...');
  const applications = readApplicationsFromExcel(EXCEL_FILE);
  console.log(`[Init] Found ${applications.length} applicant(s)`);

  if (applications.length === 0) {
    throw new Error('No rows found in the Excel file.\nAdd applicant data to: ' + EXCEL_FILE);
  }

  // Run all applications through the worker pool
  const startTime = Date.now();
  const results = await runWorkerPool(applications, MAX_WORKERS);
  const totalDuration = Date.now() - startTime;

  // Summary
  const passed = results.filter(r => r.status === 'passed').length;
  const failed = results.filter(r => r.status === 'failed').length;

  console.log('\n[Summary] ═══════════════════════════════════════════════════');
  console.log(`[Summary] Total:  ${results.length} applicant(s)`);
  console.log(`[Summary] Passed: ${passed}`);
  console.log(`[Summary] Failed: ${failed}`);
  console.log(`[Summary] Time:   ${(totalDuration / 1000).toFixed(1)}s`);

  if (passed > 0) {
    console.log('\n[Summary] Successful applications:');
    for (const r of results.filter(r => r.status === 'passed')) {
      console.log(`  ✓ ${r.label}: Application #${r.applicationNumber || 'N/A'}`);
    }
  }

  if (failed > 0) {
    console.log('\n[Summary] Failed applications:');
    for (const r of results.filter(r => r.status === 'failed')) {
      console.log(`  ✗ ${r.label}: ${r.error}`);
    }
  }

  // Save results to JSON for reference
  ensureDir('test-results');
  const resultsSummary = results.map((r, i) => ({
    index: i + 1,
    name: applications[i]?.passport.fullNameEN ?? '',
    passport: applications[i]?.passport.passportNumber ?? '',
    applicationNumber: r.applicationNumber ?? '',
    status: r.status,
    error: r.error ?? '',
    duration: `${(r.duration / 1000).toFixed(1)}s`,
  }));
  fs.writeFileSync('test-results/results.json', JSON.stringify(resultsSummary, null, 2));
  console.log('\n[Summary] Results saved → test-results/results.json');
  console.log('[Summary] ═══════════════════════════════════════════════════\n');

  // Exit with error code if any failed
  if (failed > 0) {
    process.exit(1);
  }
}

main().catch(err => {
  console.error('[Fatal]', err);
  process.exit(1);
});
