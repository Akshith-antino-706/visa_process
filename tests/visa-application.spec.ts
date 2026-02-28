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
import * as XLSX from 'xlsx';
import { createDriver, loadSession, quitDriver, SESSION_FILE } from '../src/automation/driver-factory';
import { fillApplicationForm } from '../src/automation/gdrfa-portal';
import { readApplicationsFromExcel } from '../src/utils/excel-reader';
import { VisaApplication } from '../src/types/application-data';

const EXCEL_FILE   = path.resolve('data/applications/applications.xlsx');
const RESULTS_XLSX = path.resolve('test-results/results.xlsx');
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

// ─── Results Excel Tracker ──────────────────────────────────────────────────

const TRACKER_HEADERS = ['Applicant Name', 'Passport Number', 'Application Number', 'Status', 'Error', 'Duration', 'Timestamp', 'Screenshot'];

/** Read existing tracker rows from results.xlsx (returns empty array if file doesn't exist) */
function readTracker(): string[][] {
  if (!fs.existsSync(RESULTS_XLSX)) return [];
  try {
    const wb = XLSX.readFile(RESULTS_XLSX, { cellDates: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws) return [];
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: '', raw: false });
    return rows.map(r => TRACKER_HEADERS.map(h => String((r as any)[h] ?? '')));
  } catch {
    return [];
  }
}

/** Append a single result row to results.xlsx (creates file if needed) */
function appendToTracker(row: string[]): void {
  ensureDir(path.dirname(RESULTS_XLSX));
  const existing = readTracker();
  const allRows = [TRACKER_HEADERS, ...existing, row];

  const ws = XLSX.utils.aoa_to_sheet(allRows, { cellDates: false });
  // Force all cells to string type
  for (const addr of Object.keys(ws)) {
    if (addr.startsWith('!')) continue;
    const cell = ws[addr];
    if (cell && cell.t !== undefined) { cell.t = 's'; cell.v = String(cell.v ?? ''); }
  }
  ws['!cols'] = TRACKER_HEADERS.map((h, i) => {
    const maxLen = Math.max(h.length, ...allRows.map(r => (r[i] ?? '').length));
    return { wch: Math.min(maxLen + 2, 50) };
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Results');
  XLSX.writeFile(wb, RESULTS_XLSX);
}

/** Get set of passport numbers that already passed (so we skip them) */
function getAlreadyProcessed(): Set<string> {
  const rows = readTracker();
  const done = new Set<string>();
  for (const row of rows) {
    const passport = (row[1] ?? '').trim();
    const status   = (row[3] ?? '').trim().toLowerCase();
    if (passport && status === 'passed') {
      done.add(passport);
    }
  }
  return done;
}

// ─── Worker function ─────────────────────────────────────────────────────────

const MAX_RETRIES = parseInt(process.env.RETRIES || '1', 10); // retry failed applications once by default

async function processApplication(
  application: VisaApplication,
  index: number,
  total: number,
): Promise<{ label: string; status: 'passed' | 'failed'; error?: string; duration: number; applicationNumber?: string }> {
  const label  = `applicant-${padIndex(index, total)}`;
  const prefix = `[Applicant ${index + 1}/${total} — ${label}]`;

  for (let attempt = 1; attempt <= 1 + MAX_RETRIES; attempt++) {
    const start = Date.now();
    let driver: WebDriver | null = null;
    let profileDir: string | undefined;

    try {
      if (attempt > 1) console.log(`\n${prefix} RETRY attempt ${attempt}/${1 + MAX_RETRIES}...`);
      else console.log(`\n${prefix} Starting — creating browser...`);

      const created = await createDriver();
      driver = created.driver;
      profileDir = created.profileDir;
      await loadSession(driver);

      console.log(
        `${prefix} Filling form for: ` +
        `${application.passport.fullNameEN} | ${application.passport.passportNumber}`
      );

      const applicationNumber = await fillApplicationForm(driver, application);

      ensureDir('test-results');
      const screenshotFile = `${label}-complete.png`;
      await takeScreenshot(driver, `test-results/${screenshotFile}`);
      console.log(`${prefix} Screenshot saved.`);

      const duration = Date.now() - start;
      console.log(`${prefix} PASSED — Application #: ${applicationNumber || 'N/A'} (${(duration / 1000).toFixed(1)}s)`);

      // Write to results tracker immediately
      appendToTracker([
        application.passport.fullNameEN,
        application.passport.passportNumber,
        applicationNumber || '',
        'passed',
        '',
        `${(duration / 1000).toFixed(1)}s`,
        new Date().toISOString(),
        screenshotFile,
      ]);
      console.log(`${prefix} Result saved to ${RESULTS_XLSX}`);

      return { label, status: 'passed', duration, applicationNumber };
    } catch (err: any) {
      const duration = Date.now() - start;
      const errorMsg = err?.message ?? String(err);
      const isCrash = errorMsg.includes('session deleted') || errorMsg.includes('disconnected') ||
                       errorMsg.includes('ECONNREFUSED') || errorMsg.includes('no such session') ||
                       errorMsg.includes('target window already closed') || errorMsg.includes('chrome not reachable');

      console.error(`${prefix} FAILED (attempt ${attempt}): ${errorMsg}`);

      // Save error screenshot
      const errorScreenshotFile = `${label}-error.png`;
      if (driver) {
        ensureDir('test-results');
        await takeScreenshot(driver, `test-results/${errorScreenshotFile}`).catch(() => {});
      }

      // If this was a crash and we have retries left, retry with a fresh browser
      if (isCrash && attempt <= MAX_RETRIES) {
        console.log(`${prefix} Browser crashed — will retry with fresh instance...`);
        if (driver) await quitDriver(driver, profileDir);
        await sleep(2000); // Brief cooldown before retry
        continue;
      }

      // Final attempt — write failure to tracker
      appendToTracker([
        application.passport.fullNameEN,
        application.passport.passportNumber,
        '',
        'failed',
        errorMsg.slice(0, 200),
        `${(duration / 1000).toFixed(1)}s`,
        new Date().toISOString(),
        errorScreenshotFile,
      ]);

      return { label, status: 'failed', error: errorMsg, duration };
    } finally {
      if (driver) {
        await quitDriver(driver, profileDir);
      }
    }
  }

  // Should never reach here, but TypeScript needs it
  return { label, status: 'failed', error: 'Unknown error', duration: 0 };
}

function sleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// ─── Worker Pool ─────────────────────────────────────────────────────────────

async function runWorkerPool(
  applications: VisaApplication[],
  maxWorkers: number
): Promise<Array<{ label: string; status: 'passed' | 'failed'; error?: string; duration: number; applicationNumber?: string }>> {
  const results: Array<{ label: string; status: 'passed' | 'failed'; error?: string; duration: number; applicationNumber?: string }> = [];
  let nextIndex = 0;

  const workers: Promise<void>[] = [];

  // Stagger delay between worker launches to prevent resource contention.
  // Each Chrome instance needs ~300-400MB RAM; launching all at once causes crashes.
  const STAGGER_DELAY_MS = parseInt(process.env.STAGGER_MS || '3000', 10);

  async function worker(workerNum: number): Promise<void> {
    // Stagger startup: wait progressively longer for each worker
    if (workerNum > 0) {
      const delay = workerNum * STAGGER_DELAY_MS;
      console.log(`[Pool] Worker ${workerNum + 1} waiting ${(delay / 1000).toFixed(0)}s before starting...`);
      await sleep(delay);
    }

    while (nextIndex < applications.length) {
      const idx = nextIndex++;
      const result = await processApplication(applications[idx], idx, applications.length);
      results.push(result);
    }
  }

  // Spawn up to maxWorkers concurrent workers
  const workerCount = Math.min(maxWorkers, applications.length);
  console.log(`\n[Pool] Spawning ${workerCount} parallel workers for ${applications.length} application(s)...`);
  console.log(`[Pool] Stagger delay: ${STAGGER_DELAY_MS}ms between workers\n`);

  for (let i = 0; i < workerCount; i++) {
    workers.push(worker(i));
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

  // Clean up stale temp profiles from previous runs
  try {
    const profilesDir = '/tmp/selenium-profiles';
    if (fs.existsSync(profilesDir)) {
      const staleProfiles = fs.readdirSync(profilesDir);
      if (staleProfiles.length > 0) {
        console.log(`[Init] Cleaning ${staleProfiles.length} stale Chrome profile(s)...`);
        fs.rmSync(profilesDir, { recursive: true, force: true });
        fs.mkdirSync(profilesDir, { recursive: true });
      }
    }
  } catch { /* best effort */ }

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
  const allApplications = readApplicationsFromExcel(EXCEL_FILE);
  console.log(`[Init] Found ${allApplications.length} applicant(s) in Excel`);

  if (allApplications.length === 0) {
    throw new Error('No rows found in the Excel file.\nAdd applicant data to: ' + EXCEL_FILE);
  }

  // ── Pre-flight: Check mandatory documents for all applicants ──────────────
  console.log('\n[DocCheck] ── Validating mandatory documents ──');
  const MANDATORY_DOCS: Array<{ field: keyof VisaApplication['documents']; label: string }> = [
    { field: 'sponsoredPassportPage1',   label: 'Sponsored Passport page 1' },
    { field: 'passportExternalCoverPage', label: 'Passport External Cover Page' },
    { field: 'personalPhoto',             label: 'Personal Photo' },
    { field: 'returnAirTicketPage1',      label: 'Return Air Ticket' },
  ];

  const applicantsWithMissingDocs: Array<{ name: string; passport: string; missing: string[] }> = [];
  let allDocsOk = true;

  for (const app of allApplications) {
    const name = app.passport.fullNameEN || '(unknown)';
    const pp = app.passport.passportNumber || '(no passport)';
    const missing: string[] = [];

    // Check documents folder exists
    const folder = app.documents.documentsFolder;
    if (!folder || !fs.existsSync(folder)) {
      missing.push('Documents folder missing/not found');
    }

    for (const { field, label } of MANDATORY_DOCS) {
      const filePath = app.documents[field] as string;
      if (!filePath) {
        missing.push(`${label} (no file mapped)`);
      } else if (!fs.existsSync(path.resolve(filePath))) {
        missing.push(`${label} (file not found: ${path.basename(filePath)})`);
      }
    }

    // Hotel reservation is optional — passport pic is used as fallback
    // So we don't flag it as mandatory

    if (missing.length > 0) {
      allDocsOk = false;
      applicantsWithMissingDocs.push({ name, passport: pp, missing });
    }
  }

  if (allDocsOk) {
    console.log(`[DocCheck] All ${allApplications.length} applicant(s) have all mandatory documents.`);
  } else {
    console.warn(`\n[DocCheck] ⚠ ${applicantsWithMissingDocs.length} applicant(s) have MISSING mandatory documents:\n`);
    for (const { name, passport, missing } of applicantsWithMissingDocs) {
      console.warn(`  ${name} (${passport}):`);
      for (const m of missing) {
        console.warn(`    - ${m}`);
      }
    }
    console.warn(`\n[DocCheck] These applicants will proceed but uploads may be incomplete.\n`);
  }
  console.log('[DocCheck] ── Done ──\n');

  // Filter out applicants missing mandatory docs (only process ready ones)
  const ONLY_READY = process.env.ONLY_READY !== 'false'; // default: true — only run ready applicants
  let readyApplications = allApplications;
  if (ONLY_READY && applicantsWithMissingDocs.length > 0) {
    const missingPassports = new Set(applicantsWithMissingDocs.map(a => a.passport));
    readyApplications = allApplications.filter(app => {
      const pp = app.passport.passportNumber || '(no passport)';
      return !missingPassports.has(pp);
    });
    console.log(`[Init] Filtering to ${readyApplications.length} ready applicant(s) (skipping ${allApplications.length - readyApplications.length} with missing docs)`);
    console.log(`[Init] To run ALL applicants (including incomplete): ONLY_READY=false npm test`);
  }

  // Skip applicants that already passed (check results.xlsx tracker)
  const alreadyDone = getAlreadyProcessed();
  if (alreadyDone.size > 0) {
    console.log(`[Init] Already processed (passed): ${alreadyDone.size} applicant(s)`);
    alreadyDone.forEach(pp => console.log(`  ✓ ${pp}`));
  }

  const applications = readyApplications.filter(app => {
    const passport = app.passport.passportNumber?.trim();
    if (passport && alreadyDone.has(passport)) {
      console.log(`[Init] Skipping ${app.passport.fullNameEN} (${passport}) — already passed`);
      return false;
    }
    return true;
  });

  console.log(`[Init] ${applications.length} applicant(s) remaining to process`);

  if (applications.length === 0) {
    console.log('[Init] All applicants already processed! Nothing to do.');
    return;
  }

  // Run remaining applications through the worker pool
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
