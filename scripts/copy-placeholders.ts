/**
 * Copies placeholder documents from a ready applicant to all applicants missing them.
 * Usage: npx ts-node scripts/copy-placeholders.ts
 */
import 'dotenv/config';
import * as fs from 'fs';
import * as path from 'path';
import { readApplicationsFromExcel } from '../src/utils/excel-reader';

const EXCEL_FILE = path.resolve('data/applications/applications.xlsx');
const DOCS_DIR = path.resolve('data/applications/documents');

// Source folder with all doc types (clean file names, small sizes)
const SRC = path.join(DOCS_DIR, 'INDIA');

const PLACEHOLDERS: Record<string, { src: string; keywords: string[] }> = {
  HOTEL: {
    src: path.join(SRC, 'Hotel reservationPlace of stay - Page 1.jpg'),
    keywords: ['hotel reservation', 'hotel', 'tenancy contract', 'tenancy', 'accommodation'],
  },
  COVER: {
    src: path.join(SRC, 'Passport External Cover Page.jpg'),
    keywords: ['passport external cover', 'cover page', 'passport cover'],
  },
  PHOTO: {
    src: path.join(SRC, 'Personal Photo.jpg'),
    keywords: ['personal photo', 'photo'],
  },
  PASSPORT: {
    src: path.join(SRC, 'Sponsored Passport page 1.jpg'),
    keywords: ['sponsored passport page 1', 'sponsored passport', 'passport page 1', 'passport front'],
  },
  TICKET: {
    src: path.join(SRC, 'Return air ticket - Page 1.jpg'),
    keywords: ['return air ticket', 'flight ticket', 'air ticket', 'indigo', 'boarding pass', 'itinerary', 'ticket', 'flight'],
  },
};

// Verify source files exist
for (const [key, val] of Object.entries(PLACEHOLDERS)) {
  if (!fs.existsSync(val.src)) {
    console.error(`Source file missing for ${key}: ${val.src}`);
    process.exit(1);
  }
}

const apps = readApplicationsFromExcel(EXCEL_FILE);
let totalCopied = 0;
let applicantsFixed = 0;

for (const app of apps) {
  const folder = app.documents.documentsFolder;
  if (!folder || !fs.existsSync(folder)) continue;

  const files = fs.readdirSync(folder).filter(f => !f.startsWith('.') && f !== 'desktop.ini');

  const hasAny = (keywords: string[]) =>
    files.some(f => keywords.some(kw => f.toLowerCase().includes(kw.toLowerCase())));

  const missing: string[] = [];

  for (const [key, val] of Object.entries(PLACEHOLDERS)) {
    if (!hasAny(val.keywords)) {
      missing.push(key);
    }
  }

  if (missing.length === 0) continue;

  applicantsFixed++;
  const name = app.passport.fullNameEN || '(unknown)';
  console.log(`\n${name} (${app.passport.passportNumber}) — ${missing.length} missing:`);

  for (const key of missing) {
    const { src } = PLACEHOLDERS[key];
    const destName = path.basename(src);
    const dest = path.join(folder, destName);

    if (fs.existsSync(dest)) {
      console.log(`  [skip] ${destName} — already exists`);
      continue;
    }

    fs.copyFileSync(src, dest);
    totalCopied++;
    console.log(`  [copy] ${destName}`);
  }
}

console.log(`\n${'─'.repeat(60)}`);
console.log(`Done. Copied ${totalCopied} file(s) across ${applicantsFixed} applicant(s).`);
console.log(`${'─'.repeat(60)}\n`);
