import Tesseract from 'tesseract.js';
import * as path from 'path';

export interface OcrResult {
  rawText: string;
  mrzLines: string[];
  confidence: number;
}

const MRZ_CHARS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<';

/** Extract and score MRZ candidate lines from raw OCR text. */
function extractMrzLines(rawText: string): string[] {
  // Exact 44-char lines first
  const exact = rawText
    .split('\n')
    .map((l) => l.replace(/\s/g, '').toUpperCase())
    .filter((l) => /^[A-Z0-9<]{44}$/.test(l));
  if (exact.length >= 2) return exact.slice(0, 2);

  // Relaxed: 38–46 chars, pad/trim to 44
  const relaxed = rawText
    .split('\n')
    .map((l) => l.replace(/\s/g, '').toUpperCase())
    .filter((l) => /^[A-Z0-9<]{38,46}$/.test(l))
    .map((l) => (l.length < 44 ? l.padEnd(44, '<') : l.substring(0, 44)));
  return relaxed.slice(0, 2);
}

/**
 * Extracts the MRZ (Machine Readable Zone) from a passport bio page image.
 *
 * Strategy:
 *  1. Run OCR restricted to MRZ characters with LSTM engine.
 *  2. Extract exact 44-char TD3 lines; fall back to relaxed (38–46 char) lines.
 *  3. If the first pass yields 0 lines, retry with SPARSE_TEXT and then AUTO
 *     PSM modes — this helps with JPEG images that have complex backgrounds.
 */
export async function extractPassportText(imagePath: string): Promise<OcrResult> {
  const absolutePath = path.resolve(imagePath);
  console.log(`[OCR] Processing image: ${absolutePath}`);

  // PSM modes tried in order — SINGLE_BLOCK works for clean scans; SPARSE_TEXT
  // and AUTO handle cluttered or lower-quality JPEG images better.
  const psmModes = [
    Tesseract.PSM.SINGLE_BLOCK,   // pass 1 — standard
    Tesseract.PSM.SPARSE_TEXT,    // pass 2 — tolerates gaps / backgrounds
    Tesseract.PSM.AUTO,           // pass 3 — let Tesseract decide layout
  ] as const;

  let lastRawText = '';
  let lastConfidence = 0;

  for (let pass = 0; pass < psmModes.length; pass++) {
    const psm = psmModes[pass];
    if (pass > 0) {
      console.log(`[OCR] Retrying with PSM mode ${psm} (pass ${pass + 1})...`);
    }

    const worker = await Tesseract.createWorker({
      logger: (m) => {
        if (m.status === 'recognizing text') {
          process.stdout.write(`\r[OCR] Progress: ${(m.progress * 100).toFixed(0)}%`);
        }
      },
    });

    await worker.loadLanguage('eng');
    await worker.initialize('eng', Tesseract.OEM.LSTM_ONLY);
    await worker.setParameters({
      tessedit_char_whitelist: MRZ_CHARS,
      preserve_interword_spaces: '0',
      tessedit_pageseg_mode:    psm,
      user_defined_dpi:         '300',  // prevent DPI-warning degradation
    });

    const { data } = await worker.recognize(absolutePath);
    await worker.terminate();

    process.stdout.write('\n');
    console.log(`[OCR] Confidence: ${data.confidence.toFixed(1)}%`);

    lastRawText    = data.text;
    lastConfidence = data.confidence;

    const lines = extractMrzLines(data.text);
    if (lines.length >= 2) {
      console.log(`[OCR] MRZ found on pass ${pass + 1} (PSM ${psm}).`);
      return { rawText: data.text, mrzLines: lines, confidence: data.confidence };
    }

    console.warn(`[OCR] Pass ${pass + 1}: found ${lines.length} MRZ line(s) — trying next mode...`);
  }

  // All passes exhausted — return whatever we have (caller will throw)
  console.warn('[OCR] All PSM modes exhausted. MRZ could not be detected.');
  return { rawText: lastRawText, mrzLines: [], confidence: lastConfidence };
}

export interface VizFields {
  birthPlace: string;
  placeOfIssue: string;
  issueDate: string;      // DD/MM/YYYY; '' if not detected
}

/**
 * Single OCR pass on the full passport image (no character whitelist) to extract
 * VIZ (Visual Inspection Zone) fields: Place of Birth and Place of Issue.
 */
export async function extractVizFields(imagePath: string): Promise<VizFields> {
  const absolutePath = path.resolve(imagePath);

  const worker = await Tesseract.createWorker({
    logger: (m) => {
      if (m.status === 'recognizing text') {
        process.stdout.write(`\r[OCR] VIZ scan: ${(m.progress * 100).toFixed(0)}%`);
      }
    },
  });

  await worker.loadLanguage('eng');
  await worker.initialize('eng', Tesseract.OEM.LSTM_ONLY);
  await worker.setParameters({
    tessedit_pageseg_mode: Tesseract.PSM.AUTO,
    user_defined_dpi: '300',
  });

  const { data } = await worker.recognize(absolutePath);
  await worker.terminate();
  process.stdout.write('\n');

  return {
    birthPlace:   parseBirthPlaceFromText(data.text),
    placeOfIssue: parsePlaceOfIssueFromText(data.text),
    issueDate:    parseIssueDateFromText(data.text),
  };
}

/**
 * Extracts the text that follows a matched label pattern within a line.
 * Handles bilingual labels like "जन्म स्थान / Place of Birth TONCA,GOA" correctly
 * by using the match position rather than a plain replace() which would retain
 * any prefix text (e.g. Hindi script) that appears before the English label.
 */
function extractAfterLabel(line: string, pattern: RegExp): string {
  const m = line.match(pattern);
  if (!m || m.index === undefined) return '';
  return line.substring(m.index + m[0].length).replace(/^[\s:/|]+/, '').trim();
}

/** Label patterns for Place of Birth across passport types / languages. */
const BIRTH_LABEL_PATTERNS = [
  /place\s+of\s+birth/i,          // English
  /lieu\s+de\s+naissance/i,        // French
  /geburtsort/i,                   // German
  /lugar\s+de\s+nacimiento/i,      // Spanish
  /luogo\s+di\s+nascita/i,         // Italian
  /born\s+in/i,
];

/** Label patterns for Place of Issue across passport types / languages. */
const ISSUE_LABEL_PATTERNS = [
  /place\s+of\s+issue/i,
  /lieu\s+de\s+d[eé]livrance/i,
  /ausstellungsort/i,
  /lugar\s+de\s+expedici[oó]n/i,
  /luogo\s+di\s+rilascio/i,
];

/** Label patterns for Date of Issue across passport types / languages. */
const ISSUE_DATE_LABEL_PATTERNS = [
  /date\s+of\s+issue/i,
  /date\s+of\s+issuance/i,
  /issued\s+on/i,
  /ausstellungsdatum/i,
  /fecha\s+de\s+expedici[oó]n/i,
];

const MONTH_ABBR: Record<string, string> = {
  JAN: '01', FEB: '02', MAR: '03', APR: '04', MAY: '05', JUN: '06',
  JUL: '07', AUG: '08', SEP: '09', OCT: '10', NOV: '11', DEC: '12',
};

/** Extracts a DD/MM/YYYY date from a string, supporting numeric and alphanumeric formats. */
function extractDateFromString(s: string): string {
  // DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY
  const num = s.match(/\b(\d{2})[\/\-\.](\d{2})[\/\-\.](\d{4})\b/);
  if (num) return `${num[1]}/${num[2]}/${num[3]}`;
  // DD MMM YYYY (e.g. 16 JAN 2026)
  const alpha = s.match(/\b(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})\b/);
  if (alpha) {
    const mm = MONTH_ABBR[alpha[2].toUpperCase()];
    if (mm) return `${alpha[1].padStart(2, '0')}/${mm}/${alpha[3]}`;
  }
  return '';
}

/**
 * Scans OCR text for "Date of Issue" label patterns, returns the date as DD/MM/YYYY.
 */
function parseIssueDateFromText(text: string): string {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    for (const pattern of ISSUE_DATE_LABEL_PATTERNS) {
      if (!pattern.test(line)) continue;
      // Try same line after the label
      const afterLabel = extractAfterLabel(line, pattern);
      const d = extractDateFromString(afterLabel);
      if (d) return d;
      // Try next 1–2 lines
      for (let j = i + 1; j < Math.min(i + 3, lines.length); j++) {
        const d2 = extractDateFromString(lines[j]);
        if (d2) return d2;
      }
    }
  }
  return '';
}

/**
 * Scans OCR text for common "Place of birth" label patterns across passport types,
 * then returns the value on the same line (after the label) or the next non-label line.
 *
 * Handles bilingual labels like "जन्म स्थान / Place of Birth TONCA,GOA" by
 * extracting text after the matched English label using its match position.
 */
function parseBirthPlaceFromText(text: string): string {
  return parseFieldAfterLabel(text, BIRTH_LABEL_PATTERNS, ISSUE_LABEL_PATTERNS);
}

/**
 * Scans OCR text for "Place of issue" label patterns, returns the value.
 */
function parsePlaceOfIssueFromText(text: string): string {
  return parseFieldAfterLabel(text, ISSUE_LABEL_PATTERNS, BIRTH_LABEL_PATTERNS);
}

/**
 * Generic label-value scanner used by both birth place and place of issue parsers.
 * labelPatterns: patterns we are looking for.
 * otherPatterns: patterns that should NOT appear in a candidate value line.
 */
function parseFieldAfterLabel(
  text: string,
  labelPatterns: RegExp[],
  otherPatterns: RegExp[],
): string {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const allLabels = [...labelPatterns, ...otherPatterns];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    for (const pattern of labelPatterns) {
      if (!pattern.test(line)) continue;

      // Value may follow the label on the same line.
      // Use match position so bilingual prefixes (e.g. Hindi script) are ignored.
      const afterLabel = extractAfterLabel(line, pattern);
      if (afterLabel.length > 1) return afterLabel.toUpperCase();

      // Otherwise look on the next 1–2 lines.
      for (let j = i + 1; j < Math.min(i + 3, lines.length); j++) {
        const candidate = lines[j].trim();
        // Skip blank lines, lines with colons (other labels), or any label text
        if (
          candidate.length > 1 &&
          !candidate.includes(':') &&
          !allLabels.some(p => p.test(candidate))
        ) {
          return candidate.toUpperCase();
        }
      }
    }
  }

  return '';
}
