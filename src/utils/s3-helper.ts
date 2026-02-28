import * as path from 'path';

// ─── Configuration ──────────────────────────────────────────────────────────
// Public S3 bucket: https://rayna-tours.s3.eu-north-1.amazonaws.com/Visa-process/
// No AWS credentials needed — the browser fetches files directly via public URL.

const BUCKET = process.env.S3_BUCKET || 'rayna-tours';
const REGION = process.env.AWS_REGION || 'eu-north-1';
const S3_PREFIX = process.env.S3_PREFIX || 'Visa-process';
const BASE_URL = `https://${BUCKET}.s3.${REGION}.amazonaws.com`;

/**
 * Check if S3 is configured (bucket name is set).
 * No credentials needed — we use public URLs.
 */
export function isS3Configured(): boolean {
  return BUCKET.length > 0;
}

/**
 * Given a local file path like:
 *   /Users/.../data/applications/ASMA SAJID KHOT/documents/passport.jpg
 *
 * Derive the S3 key:
 *   Visa-process/applications/ASMA SAJID KHOT/documents/passport.jpg
 *
 * The S3 bucket mirrors the local data/ folder under the Visa-process/ prefix.
 */
export function localPathToS3Key(localFilePath: string): string {
  const resolved = path.resolve(localFilePath);

  // Try to find "applications/" in the path — that's our anchor point
  const marker = '/applications/';
  const idx = resolved.indexOf(marker);
  if (idx >= 0) {
    return `${S3_PREFIX}/applications/${resolved.substring(idx + marker.length)}`;
  }

  // Try "data/" as anchor
  const dataMarker = '/data/';
  const dataIdx = resolved.indexOf(dataMarker);
  if (dataIdx >= 0) {
    return `${S3_PREFIX}/${resolved.substring(dataIdx + dataMarker.length)}`;
  }

  // Fallback: just use the filename under the prefix
  return `${S3_PREFIX}/documents/${path.basename(resolved)}`;
}

/**
 * Build the public S3 URL for a local file path.
 * Each path segment is URI-encoded to handle spaces and special characters.
 *
 * Example:
 *   localPath: /Users/.../applications/ASMA SAJID KHOT/documents/Passport External Cover Page.jpg
 *   S3 URL:    https://rayna-tours.s3.eu-north-1.amazonaws.com/Visa-process/applications/ASMA%20SAJID%20KHOT/documents/Passport%20External%20Cover%20Page.jpg
 */
export function getPublicUrl(localFilePath: string): string {
  const key = localPathToS3Key(localFilePath);
  // Encode each path segment individually (preserving slashes)
  const encodedKey = key
    .split('/')
    .map((segment) => encodeURIComponent(segment))
    .join('/');
  return `${BASE_URL}/${encodedKey}`;
}
