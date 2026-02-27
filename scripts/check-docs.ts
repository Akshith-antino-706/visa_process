/**
 * Pre-flight check: lists which applicants have all mandatory docs
 * and which are missing files.
 *
 * Usage: npx ts-node scripts/check-docs.ts
 */
import 'dotenv/config';
import * as fs from 'fs';
import * as path from 'path';
import { readApplicationsFromExcel } from '../src/utils/excel-reader';

const EXCEL_FILE = path.resolve('data/applications/applications.xlsx');

const MANDATORY: Array<{ field: string; label: string }> = [
  { field: 'sponsoredPassportPage1',   label: 'Sponsored Passport page 1' },
  { field: 'passportExternalCoverPage', label: 'Passport External Cover Page' },
  { field: 'personalPhoto',             label: 'Personal Photo' },
  { field: 'returnAirTicketPage1',      label: 'Return Air Ticket' },
  { field: 'hotelReservationPage1',     label: 'Hotel Reservation' },
];

const apps = readApplicationsFromExcel(EXCEL_FILE);

const ready: Array<{ name: string; pp: string; folder: string; files: string[] }> = [];
const notReady: Array<{ name: string; pp: string; missing: string[] }> = [];

for (const app of apps) {
  const name = app.passport.fullNameEN || '(unknown)';
  const pp = app.passport.passportNumber || '(no passport)';
  const folder = app.documents.documentsFolder || '';
  const missing: string[] = [];

  if (!folder || !fs.existsSync(folder)) {
    missing.push('Documents folder missing/not found');
  }

  for (const { field, label } of MANDATORY) {
    const fp = (app.documents as any)[field] as string;
    if (!fp) {
      missing.push(`${label} (no file mapped)`);
    } else if (!fs.existsSync(path.resolve(fp))) {
      missing.push(`${label} (file not found: ${path.basename(fp)})`);
    }
  }

  if (missing.length === 0) {
    const files = folder && fs.existsSync(folder)
      ? fs.readdirSync(folder).filter(f => !f.startsWith('.'))
      : [];
    ready.push({ name, pp, folder, files });
  } else {
    notReady.push({ name, pp, missing });
  }
}

// Print results
console.log(`\n${'='.repeat(70)}`);
console.log(`  READY (${ready.length}) — All mandatory docs present`);
console.log(`${'='.repeat(70)}`);
for (const r of ready) {
  console.log(`  ✓  ${r.name} (${r.pp})`);
  console.log(`     Folder: ${r.folder}`);
  console.log(`     Files:  ${r.files.join(', ')}`);
}

console.log(`\n${'='.repeat(70)}`);
console.log(`  NOT READY (${notReady.length}) — Missing documents`);
console.log(`${'='.repeat(70)}`);
for (const r of notReady) {
  console.log(`  ✗  ${r.name} (${r.pp})`);
  for (const m of r.missing) {
    console.log(`     - ${m}`);
  }
}

console.log(`\n${'─'.repeat(70)}`);
console.log(`  Total: ${apps.length} | Ready: ${ready.length} | Not Ready: ${notReady.length}`);
console.log(`${'─'.repeat(70)}\n`);
