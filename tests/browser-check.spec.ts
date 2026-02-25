/**
 * Debug utility: Basic browser launch test.
 * Usage: ts-node tests/browser-check.spec.ts
 */

import { createDriver, quitDriver } from '../src/automation/driver-factory';

async function main() {
  console.log('[Check] Launching browser...');
  const driver = await createDriver();

  try {
    console.log('[Check] Going to Google...');
    await driver.get('https://www.google.com');
    const url = await driver.getCurrentUrl();
    console.log('[Check] URL:', url);

    if (!url.includes('google')) {
      throw new Error(`Expected Google URL, got: ${url}`);
    }
    console.log('[Check] PASSED');
  } finally {
    await quitDriver(driver);
  }
}

main().catch(err => {
  console.error('[Check] FAILED:', err);
  process.exit(1);
});
