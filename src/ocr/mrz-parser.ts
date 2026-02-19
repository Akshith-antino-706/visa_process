import { parse } from 'mrz';
import { extractPassportText, extractVizFields } from './passport-reader';

export interface PassportData {
  // Identity
  surname: string;
  givenNames: string;
  fullName: string;

  // Document
  passportNumber: string;
  nationality: string;
  issuingCountry: string;

  // Dates (formatted as DD/MM/YYYY for form filling)
  dateOfBirth: string;
  expiryDate: string;

  // Other
  gender: string;
  personalNumber: string;
  birthPlace: string;       // Extracted from VIZ (not MRZ); '' if not detected
  placeOfIssue: string;     // Extracted from VIZ (not MRZ); '' if not detected
  issueDate: string;        // Extracted from VIZ (not MRZ); '' if not detected

  // Raw MRZ
  mrzLine1: string;
  mrzLine2: string;
}

/**
 * MRZ check-digit algorithm (weights 7-3-1, A=10…Z=35, 0-9=face value, '<'=0).
 */
function mrzCheckDigit(str: string): number {
  const weights = [7, 3, 1];
  let sum = 0;
  for (let i = 0; i < str.length; i++) {
    const c = str[i];
    let val: number;
    if (c === '<') val = 0;
    else if (c >= '0' && c <= '9') val = parseInt(c, 10);
    else if (c >= 'A' && c <= 'Z') val = c.charCodeAt(0) - 55; // A=10
    else val = 0;
    sum += val * weights[i % 3];
  }
  return sum % 10;
}

/**
 * Corrects common OCR confusions (O↔0, I↔1) in the 9-char MRZ document
 * number by trying all ambiguous combinations until the check digit passes.
 * Returns the corrected number, or the original if no correction is found.
 */
function correctOCRPassportNumber(raw: string, checkDigit: number): string {
  // Positions where OCR may have confused visually similar characters
  const ambiguous: Array<{ idx: number; alts: string[] }> = [];
  for (let i = 0; i < raw.length; i++) {
    if (raw[i] === 'O' || raw[i] === '0') ambiguous.push({ idx: i, alts: ['O', '0'] });
    else if (raw[i] === 'I' || raw[i] === '1') ambiguous.push({ idx: i, alts: ['I', '1'] });
  }
  if (ambiguous.length === 0) return raw;

  const combos = Math.pow(2, ambiguous.length);
  for (let mask = 0; mask < combos; mask++) {
    const chars = raw.split('');
    for (let j = 0; j < ambiguous.length; j++) {
      chars[ambiguous[j].idx] = ambiguous[j].alts[(mask >> j) & 1];
    }
    const candidate = chars.join('');
    if (mrzCheckDigit(candidate) === checkDigit) {
      return candidate;
    }
  }
  return raw; // no substitution found — return as-is
}

/**
 * Fallback name extractor for when OCR reads the << separator as CC/LL/CL/LC.
 *
 * TD3 line 1 name zone (chars 5-43): SURNAME<<GIVENNAME1<GIVENNAME2<<<padding
 * When OCR reads '<' as 'C' or 'L', << becomes CC/LL/CL/LC, and individual
 * filler '<' becomes C or L.  The mrz library then cannot find the separator
 * and returns the full noise string for both surname and givenNames.
 *
 * Strategy: scan for the first run of 2+ chars from {C, L, <} at position ≥ 3
 * (to skip any real double-consonant at the start, e.g. MCCALL).  Split there:
 * surname = everything before, given names = first plausible token after.
 */
function fallbackExtractNames(line1: string): { surname: string; givenNames: string } | null {
  if (!line1 || line1.length < 5) return null;

  // TD3 name zone starts at position 5 (after P<CCC)
  const nameZone = line1.substring(5);

  // Try each run of 2+ potential filler chars in order
  const sepRegex = /[CL<]{2,}/g;
  let match: RegExpExecArray | null;

  while ((match = sepRegex.exec(nameZone)) !== null) {
    const sepIndex = match.index;

    // Surname must be at least 3 chars (skip e.g. MC before MCCALL)
    if (sepIndex < 3) continue;

    const surname   = nameZone.substring(0, sepIndex);
    const afterSep  = nameZone.substring(sepIndex + match[0].length);

    // Split given names on runs of 1+ filler chars; discard short noise tokens
    const givenParts = afterSep.split(/[CL<]+/).filter(p => p.length >= 2);
    if (givenParts.length === 0) continue;

    const givenNames = givenParts.join(' ').trim();
    console.log(`[MRZ] Fallback name extraction: surname="${surname}", givenNames="${givenNames}"`);
    return { surname, givenNames };
  }

  return null;
}

/**
 * Parses two MRZ lines from a TD3 (standard 88-page passport) into structured passport data.
 */
export function parseMRZ(mrzLines: string[]): PassportData {
  if (mrzLines.length < 2) {
    throw new Error(`[MRZ] Expected 2 MRZ lines, got ${mrzLines.length}`);
  }

  const [line1, line2] = mrzLines;

  let parsed: ReturnType<typeof parse>;
  try {
    parsed = parse([line1, line2]);
  } catch (err) {
    throw new Error(`[MRZ] Failed to parse MRZ: ${err}`);
  }

  const fields = parsed.fields;

  // TD3 line 2: positions 0-8 = document number, position 9 = check digit
  let passportNumber = fields.documentNumber ?? '';
  if (line2 && line2.length >= 10) {
    const rawDocField  = line2.substring(0, 9).replace(/</g, '');
    const checkDigitCh = line2[9];
    const expectedCheck = parseInt(checkDigitCh, 10);
    if (!isNaN(expectedCheck) && mrzCheckDigit(rawDocField) !== expectedCheck) {
      const corrected = correctOCRPassportNumber(rawDocField, expectedCheck);
      if (corrected !== rawDocField) {
        console.log(`[MRZ] OCR correction: "${rawDocField}" → "${corrected}" (check digit: ${expectedCheck})`);
        passportNumber = corrected;
      }
    }
  }

  let surname    = fields.lastName  ?? '';
  let givenNames = fields.firstName ?? '';

  // When OCR reads the << separator as CC/LL the mrz library cannot split the name
  // field and returns the full noise string for both fields (they become identical).
  // Apply fallback parser directly on the raw line 1 name zone in that case.
  if (surname && givenNames && surname === givenNames) {
    console.log(`[MRZ] Library returned identical surname/givenNames — applying fallback name parser.`);
    const fallback = fallbackExtractNames(line1);
    if (fallback) {
      surname    = fallback.surname;
      givenNames = fallback.givenNames;
    }
  }

  // Remove filler noise from givenNames. OCR reads individual '<' padding chars as
  // 'C' or 'L', so the library may return tokens like "LLLLL" (pure filler) or
  // "LLLLLLLLLLLLLLLRK" (filler with a stray misread char at the end).
  // Two rules eliminate both cases:
  //   1. Token is entirely C/L characters  → pure filler noise
  //   2. Token contains a run of 3+ C/L   → mostly filler, stray chars at edge
  // Real names never contain 3+ consecutive identical consonants like LLL/CCC.
  if (givenNames) {
    givenNames = givenNames
      .split(/\s+/)
      .filter(w => w.length >= 2 && !/^[CL]+$/.test(w) && !/[CL]{3,}/.test(w))
      .join(' ')
      .trim();
  }

  return {
    surname,
    givenNames,
    fullName: `${givenNames} ${surname}`.trim(),
    passportNumber,
    nationality: fields.nationality ?? '',
    issuingCountry: fields.issuingState ?? '',
    dateOfBirth: formatMRZDate(fields.birthDate ?? ''),
    expiryDate: formatMRZDate(fields.expirationDate ?? ''),
    gender: formatGender(fields.sex ?? ''),
    personalNumber: fields.optional ?? '',
    birthPlace: '',    // populated later by VIZ scan in extractValidateAndParse
    placeOfIssue: '',  // populated later by VIZ scan in extractValidateAndParse
    issueDate: '',     // populated later by VIZ scan in extractValidateAndParse
    mrzLine1: line1,
    mrzLine2: line2,
  };
}

/**
 * Converts MRZ date (YYMMDD) to DD/MM/YYYY.
 * Assumes 2000s for years <= current year, 1900s otherwise.
 */
function formatMRZDate(mrzDate: string): string {
  if (!mrzDate || mrzDate.length !== 6) return mrzDate;

  const yy = parseInt(mrzDate.substring(0, 2), 10);
  const mm = mrzDate.substring(2, 4);
  const dd = mrzDate.substring(4, 6);

  const currentYear = new Date().getFullYear() % 100;
  const fullYear = yy <= currentYear + 10 ? 2000 + yy : 1900 + yy;

  return `${dd}/${mm}/${fullYear}`;
}

function formatGender(sex: string): string {
  if (sex === 'M') return 'Male';
  if (sex === 'F') return 'Female';
  return 'Unspecified';
}

/**
 * Convenience function: run OCR + parse in one call.
 */
export async function extractAndParsePassport(imagePath: string): Promise<PassportData> {
  const { mrzLines, confidence } = await extractPassportText(imagePath);

  console.log(`[MRZ] Detected ${mrzLines.length} MRZ line(s)`);
  if (mrzLines.length < 2) {
    throw new Error(
      `[MRZ] Could not detect MRZ lines from image. OCR confidence: ${confidence.toFixed(1)}%. ` +
        `Ensure the image is clear and the passport bio page is fully visible.`
    );
  }

  const data = parseMRZ(mrzLines);
  console.log(`[MRZ] Parsed passport: ${data.fullName} | ${data.passportNumber} | Expiry: ${data.expiryDate}`);
  return data;
}

// ─── Validation ───────────────────────────────────────────────────────────────

export interface ValidationResult {
  valid: boolean;
  errors: string[];
  warnings: string[];
}

/**
 * Validates extracted passport data before proceeding with portal automation.
 * Throws if any critical field is missing or invalid.
 */
export function validatePassportData(data: PassportData, ocrConfidence?: number): ValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];

  // ── Critical fields ──────────────────────────────────────────────────────
  if (!data.passportNumber || data.passportNumber.trim() === '') {
    errors.push('Passport number is missing.');
  } else if (!/^[A-Z0-9]{6,9}$/.test(data.passportNumber.trim())) {
    errors.push(`Passport number "${data.passportNumber}" looks invalid (expected 6-9 alphanumeric characters).`);
  }

  if (!data.surname || data.surname.trim() === '') {
    errors.push('Surname is missing.');
  }

  if (!data.givenNames || data.givenNames.trim() === '') {
    errors.push('Given name(s) are missing.');
  }

  if (!data.nationality || data.nationality.trim() === '') {
    errors.push('Nationality is missing.');
  }

  // ── Date of Birth ─────────────────────────────────────────────────────────
  if (!data.dateOfBirth || data.dateOfBirth.trim() === '') {
    errors.push('Date of birth is missing.');
  } else {
    const dob = parseDDMMYYYY(data.dateOfBirth);
    if (!dob) {
      errors.push(`Date of birth "${data.dateOfBirth}" could not be parsed (expected DD/MM/YYYY).`);
    } else if (dob >= new Date()) {
      errors.push(`Date of birth "${data.dateOfBirth}" is in the future — OCR likely misread the date.`);
    } else {
      const ageYears = (Date.now() - dob.getTime()) / (1000 * 60 * 60 * 24 * 365.25);
      if (ageYears < 1) {
        warnings.push(`Applicant age appears to be less than 1 year — please verify DOB: ${data.dateOfBirth}`);
      }
      if (ageYears > 120) {
        errors.push(`Date of birth "${data.dateOfBirth}" implies age > 120 years — OCR likely misread the date.`);
      }
    }
  }

  // ── Expiry Date ───────────────────────────────────────────────────────────
  if (!data.expiryDate || data.expiryDate.trim() === '') {
    errors.push('Passport expiry date is missing.');
  } else {
    const expiry = parseDDMMYYYY(data.expiryDate);
    if (!expiry) {
      errors.push(`Expiry date "${data.expiryDate}" could not be parsed (expected DD/MM/YYYY).`);
    } else if (expiry < new Date()) {
      errors.push(`Passport is EXPIRED (expiry: ${data.expiryDate}). Cannot proceed with application.`);
    } else {
      const daysToExpiry = (expiry.getTime() - Date.now()) / (1000 * 60 * 60 * 24);
      if (daysToExpiry < 180) {
        warnings.push(`Passport expires in ${Math.round(daysToExpiry)} days (${data.expiryDate}). Some countries require 6 months validity.`);
      }
    }
  }

  // ── Gender ────────────────────────────────────────────────────────────────
  if (data.gender === 'Unspecified') {
    warnings.push('Gender could not be determined from MRZ. Please verify manually.');
  }

  // ── OCR Confidence ────────────────────────────────────────────────────────
  if (ocrConfidence !== undefined) {
    if (ocrConfidence < 50) {
      errors.push(`OCR confidence is very low (${ocrConfidence.toFixed(1)}%). Image may be too blurry or dark.`);
    } else if (ocrConfidence < 75) {
      warnings.push(`OCR confidence is moderate (${ocrConfidence.toFixed(1)}%). Review extracted data carefully.`);
    }
  }

  return { valid: errors.length === 0, errors, warnings };
}

/**
 * Runs OCR, parses MRZ, validates all fields, and returns the result.
 * Throws with a detailed error report if validation fails.
 */
export async function extractValidateAndParse(imagePath: string): Promise<PassportData> {
  const { mrzLines, confidence } = await extractPassportText(imagePath);

  if (mrzLines.length < 2) {
    throw new Error(
      `[OCR] MRZ not detected. Confidence: ${confidence.toFixed(1)}%.\n` +
      `  Ensure the passport bio page is fully visible, well-lit, and in focus.`
    );
  }

  const data = parseMRZ(mrzLines);

  // Second pass: extract VIZ fields (Place of Birth + Place of Issue)
  console.log('[OCR] Scanning VIZ for Place of Birth and Place of Issue...');
  const viz = await extractVizFields(imagePath);
  data.birthPlace   = viz.birthPlace;
  data.placeOfIssue = viz.placeOfIssue;
  data.issueDate    = viz.issueDate;
  console.log(`[OCR] Birth Place   : ${data.birthPlace   || '(not detected)'}`);
  console.log(`[OCR] Place of Issue: ${data.placeOfIssue || '(not detected)'}`);
  console.log(`[OCR] Issue Date    : ${data.issueDate    || '(not detected)'}`);


  // Don't pass confidence — Tesseract confidence is artificially low when using
  // a character whitelist (all non-MRZ content maps to garbage chars). Rely on
  // structural field validation instead.
  const result = validatePassportData(data);

  // Print summary
  console.log('\n─── OCR Validation Report ────────────────────');
  console.log(`  Name         : ${data.fullName}`);
  console.log(`  Passport No  : ${data.passportNumber}`);
  console.log(`  Nationality  : ${data.nationality}`);
  console.log(`  Date of Birth: ${data.dateOfBirth}`);
  console.log(`  Expiry Date  : ${data.expiryDate}`);
  console.log(`  Gender       : ${data.gender}`);
  console.log(`  Birth Place   : ${data.birthPlace   || '(not detected)'}`);
  console.log(`  Place of Issue: ${data.placeOfIssue || '(not detected)'}`);
  console.log(`  Issue Date    : ${data.issueDate    || '(not detected)'}`);
  console.log(`  OCR Confidence: ${confidence.toFixed(1)}%`);

  if (result.warnings.length > 0) {
    console.log('\n  ⚠ Warnings:');
    result.warnings.forEach(w => console.log(`    - ${w}`));
  }

  if (!result.valid) {
    const report = result.errors.map(e => `    ✗ ${e}`).join('\n');
    throw new Error(
      `[OCR] Passport validation failed — ${result.errors.length} error(s):\n${report}\n` +
      `  Fix the image or correct the data before proceeding.`
    );
  }

  console.log('  ✓ Validation passed — proceeding with portal automation.');
  console.log('──────────────────────────────────────────────\n');

  return data;
}

// ─── Date Helper ──────────────────────────────────────────────────────────────

function parseDDMMYYYY(dateStr: string): Date | null {
  const match = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return null;
  const [, dd, mm, yyyy] = match;
  const d = new Date(Number(yyyy), Number(mm) - 1, Number(dd));
  if (isNaN(d.getTime())) return null;
  return d;
}
