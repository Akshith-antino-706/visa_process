/**
 * OCR Extraction Script
 *
 * Usage:  npm run ocr
 *
 * For each passport image in data/passports/, runs OCR, merges the extracted
 * data into the application template (data/applications/sample.json), and writes
 * the result to data/applications/processed/applicant-01.json, etc.
 *
 * Review (and optionally edit) the output JSON files before running npm test.
 */

import * as fs   from 'fs';
import * as path from 'path';
import * as dotenv from 'dotenv';
import { extractValidateAndParse, PassportData } from '../src/ocr/mrz-parser';
import { VisaApplication } from '../src/types/application-data';

dotenv.config();

const PASSPORTS_DIR = path.resolve('data/passports');
const TEMPLATE_FILE = path.resolve('data/applications/sample.json');
const PROCESSED_DIR = path.resolve('data/applications/processed');
const IMAGE_EXTS    = ['.jpg', '.jpeg', '.png'];

// ─── Helpers ──────────────────────────────────────────────────────────────────

function ensureDir(dir: string): void {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function applicantLabel(index: number, total: number): string {
  const digits = String(total).length;
  return `applicant-${String(index + 1).padStart(digits, '0')}`;
}

function mergeOcrIntoApplication(application: VisaApplication, ocr: PassportData): void {
  const p = application.passport;

  p.passportNumber     = ocr.passportNumber;
  p.fullNameEN         = p.fullNameEN         || ocr.fullName;
  p.dateOfBirth        = p.dateOfBirth        || ocr.dateOfBirth;
  p.passportExpiryDate = p.passportExpiryDate || ocr.expiryDate;
  p.gender             = p.gender             || (ocr.gender === 'Male' ? 'Male' : 'Female');
  p.currentNationality  = p.currentNationality  || ocr.nationality;
  p.previousNationality = p.previousNationality || p.currentNationality;
  p.birthCountry        = p.birthCountry        || p.currentNationality;
  p.birthPlaceEN        = ocr.birthPlace || p.birthPlaceEN || p.currentNationality;
  p.passportIssueCountry    = p.passportIssueCountry    || p.currentNationality;
  p.passportIssueDate       = p.passportIssueDate       || ocr.issueDate;
  p.passportPlaceOfIssueEN  = p.passportPlaceOfIssueEN  || ocr.placeOfIssue;

  application.applicant.comingFromCountry =
    application.applicant.comingFromCountry || p.currentNationality;

  // Split given names into first + middle.
  // Single-character tokens are OCR noise from '<' filler chars misread as 'C'/'L'.
  const givenParts = ocr.givenNames.trim().split(/\s+/).filter(w => w.length > 1);
  p.firstName  = givenParts[0] ?? '';
  p.middleName = givenParts.length > 1 ? givenParts.slice(1).join(' ') : undefined;
  p.lastName   = ocr.surname;
}

// ─── Main ─────────────────────────────────────────────────────────────────────

async function main(): Promise<void> {
  // Validate inputs
  if (!fs.existsSync(PASSPORTS_DIR)) {
    throw new Error(`Passports directory not found: ${PASSPORTS_DIR}`);
  }
  if (!fs.existsSync(TEMPLATE_FILE)) {
    throw new Error(`Application template not found: ${TEMPLATE_FILE}`);
  }

  // Collect passport images
  const images = fs.readdirSync(PASSPORTS_DIR)
    .filter(f => IMAGE_EXTS.includes(path.extname(f).toLowerCase()))
    .sort()
    .map(f => path.join(PASSPORTS_DIR, f));

  if (images.length === 0) {
    throw new Error(`No image files found in: ${PASSPORTS_DIR}`);
  }

  ensureDir(PROCESSED_DIR);

  console.log(`\nFound ${images.length} passport image(s). Starting OCR extraction...\n`);

  let passed = 0;
  let failed = 0;

  for (let i = 0; i < images.length; i++) {
    const imagePath = images[i];
    const label     = applicantLabel(i, images.length);
    const prefix    = `[${label} — ${path.basename(imagePath)}]`;

    console.log(`${prefix} Running OCR...`);

    try {
      const ocrData = await extractValidateAndParse(imagePath);

      // Load a fresh copy of the template for each applicant
      const application = JSON.parse(
        fs.readFileSync(TEMPLATE_FILE, 'utf-8')
      ) as VisaApplication;

      // Point documents to this passport image
      application.documents.sponsoredPassportPage1 = imagePath;

      // Merge OCR data (template values take priority for non-empty fields)
      mergeOcrIntoApplication(application, ocrData);

      // Write to processed directory
      const outPath = path.join(PROCESSED_DIR, `${label}.json`);
      fs.writeFileSync(outPath, JSON.stringify(application, null, 2), 'utf-8');

      console.log(`${prefix} ✓ Stored → ${outPath}`);
      console.log(`${prefix}   Name: ${ocrData.fullName} | Passport: ${ocrData.passportNumber}\n`);
      passed++;
    } catch (err) {
      console.error(`${prefix} ✗ OCR failed: ${(err as Error).message}\n`);
      failed++;
    }
  }

  console.log('─'.repeat(60));
  console.log(`Done: ${passed} succeeded, ${failed} failed.`);

  if (passed > 0) {
    console.log(`\nReview the JSON files in: ${PROCESSED_DIR}`);
    console.log('Edit any fields that need correction, then run:\n');
    console.log('  npm test\n');
  }

  if (failed > 0) process.exit(1);
}

main().catch(err => {
  console.error('\nFatal error:', err.message ?? err);
  process.exit(1);
});
