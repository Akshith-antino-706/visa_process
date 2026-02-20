/**
 * JSON-to-Excel Script
 *
 * Usage:  npx ts-node scripts/json-to-excel.ts
 *
 * Reads all applicant JSON files from data/applications/processed/
 * and writes a single Excel file: data/applications/applications.xlsx
 */

import * as fs   from 'fs';
import * as path from 'path';
import * as XLSX from 'xlsx';

const PROCESSED_DIR = path.resolve('data/applications/processed');
const OUTPUT_FILE   = path.resolve('data/applications/applications.xlsx');

// Column definitions: [header, accessor function]
const COLUMNS: [string, (app: any) => string][] = [
  // Passport / identity
  ['Full Name',            a => a.passport?.fullNameEN ?? ''],
  ['First Name',           a => a.passport?.firstName ?? ''],
  ['Middle Name',          a => a.passport?.middleName ?? ''],
  ['Last Name',            a => a.passport?.lastName ?? ''],
  ['Gender',               a => a.passport?.gender ?? ''],
  ['Date of Birth',        a => a.passport?.dateOfBirth ?? ''],
  ['Passport Number',      a => a.passport?.passportNumber ?? ''],
  ['Passport Type',        a => a.passport?.passportType ?? ''],
  ['Current Nationality',  a => a.passport?.currentNationality ?? ''],
  ['Previous Nationality', a => a.passport?.previousNationality ?? ''],
  ['Birth Country',        a => a.passport?.birthCountry ?? ''],
  ['Birth Place',          a => a.passport?.birthPlaceEN ?? ''],
  ['Passport Issue Country', a => a.passport?.passportIssueCountry ?? ''],
  ['Passport Issue Date',   a => a.passport?.passportIssueDate ?? ''],
  ['Passport Expiry Date',  a => a.passport?.passportExpiryDate ?? ''],
  ['Passport Place of Issue', a => a.passport?.passportPlaceOfIssueEN ?? ''],

  // Applicant details
  ['Mother Name',          a => a.applicant?.motherNameEN ?? ''],
  ['Marital Status',       a => a.applicant?.maritalStatus ?? ''],
  ['Religion',             a => a.applicant?.religion ?? ''],
  ['Faith',                a => a.applicant?.faith ?? ''],
  ['Education',            a => a.applicant?.education ?? ''],
  ['Profession',           a => a.applicant?.profession ?? ''],
  ['First Language',       a => a.applicant?.firstLanguage ?? ''],
  ['Coming From Country',  a => a.applicant?.comingFromCountry ?? ''],
  ['Inside UAE',           a => a.applicant?.isInsideUAE ? 'Yes' : 'No'],
  ['Relationship to Host', a => a.applicant?.relationshipToHost ?? ''],

  // Visit details
  ['Purpose of Visit',     a => a.visit?.purposeOfVisit ?? ''],
  ['Date of Arrival',      a => a.visit?.dateOfArrival ?? ''],
  ['Date of Departure',    a => a.visit?.dateOfDeparture ?? ''],
  ['Port of Entry',        a => a.visit?.portOfEntry ?? ''],
  ['Accommodation Type',   a => a.visit?.accommodationType ?? ''],
  ['Hotel / Place of Stay', a => a.visit?.hotelOrPlaceOfStay ?? ''],

  // Contact
  ['Email',                a => a.contact?.email ?? ''],
  ['Mobile Number',        a => a.contact?.mobileNumber ?? ''],
  ['SMS Language',         a => a.contact?.preferredSMSLanguage ?? ''],
  ['UAE Emirate',          a => a.contact?.uaeEmirate ?? ''],
  ['UAE City',             a => a.contact?.uaeCity ?? ''],
  ['UAE Area',             a => a.contact?.uaeArea ?? ''],
  ['UAE Street',           a => a.contact?.uaeStreet ?? ''],
  ['UAE Building',         a => a.contact?.uaeBuilding ?? ''],
  ['UAE Floor',            a => a.contact?.uaeFloor ?? ''],
  ['UAE Flat',             a => a.contact?.uaeFlat ?? ''],

  // Address Outside UAE
  ['Outside Country',      a => a.contact?.outsideCountry ?? ''],
  ['Outside Mobile',       a => a.contact?.outsideMobile ?? ''],
  ['Outside City',         a => a.contact?.outsideCity ?? ''],
  ['Outside Address',      a => a.contact?.outsideAddress ?? ''],

  // Host / submitter
  ['Establishment',        a => a.hostSubmitter?.establishmentNameEN ?? ''],

  // Documents folder (relative path to folder containing document files)
  ['Documents Folder',     a => a.documentsFolder ?? ''],
];

function main(): void {
  if (!fs.existsSync(PROCESSED_DIR)) {
    console.error(`Processed directory not found: ${PROCESSED_DIR}`);
    process.exit(1);
  }

  const files = fs.readdirSync(PROCESSED_DIR)
    .filter(f => f.endsWith('.json'))
    .sort();

  if (files.length === 0) {
    console.error('No JSON files found in processed directory.');
    process.exit(1);
  }

  console.log(`Found ${files.length} applicant file(s).\n`);

  // Build header row
  const headers = COLUMNS.map(([header]) => header);

  // Build data rows
  const rows: string[][] = [];
  for (const file of files) {
    const filePath = path.join(PROCESSED_DIR, file);
    const app = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    const row = COLUMNS.map(([, accessor]) => accessor(app));
    rows.push(row);
    console.log(`  + ${file} → ${app.passport?.fullNameEN || '(unnamed)'}`);
  }

  // Create workbook — force all cells to text type ('s') so Excel does not
  // auto-convert date-looking values (e.g. "13/10/1973") into serial numbers.
  const wsData = [headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(wsData, { cellDates: false });
  // Mark every data cell as type 's' (string)
  for (const addr of Object.keys(ws)) {
    if (addr.startsWith('!')) continue;
    const cell = ws[addr];
    if (cell && cell.t !== undefined) {
      cell.t = 's';
      cell.v = String(cell.v ?? '');
    }
  }

  // Auto-size columns
  ws['!cols'] = headers.map((h, i) => {
    const maxLen = Math.max(h.length, ...rows.map(r => (r[i] ?? '').length));
    return { wch: Math.min(maxLen + 2, 40) };
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Applications');
  XLSX.writeFile(wb, OUTPUT_FILE);

  console.log(`\nExcel file written to: ${OUTPUT_FILE}`);
}

main();
