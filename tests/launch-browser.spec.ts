import { test } from '@playwright/test';
import * as path from 'path';

const SESSION_FILE = path.resolve('auth/session.json');

test.use({ storageState: SESSION_FILE });

test('Launch browser with session', async ({ page }) => {
  console.log('[Launch] Browser opened');
  console.log('[Launch] Navigating to portal...');
  await page.goto('https://smart.gdrfad.gov.ae/SmartChannels_Th/', { waitUntil: 'domcontentloaded' });
  console.log('[Launch] URL:', page.url());

  if (page.url().includes('Login.aspx')) {
    console.error('[Launch] Session expired — run "npm run auth" first');
  } else {
    console.log('[Launch] Session is valid — logged in');
  }

  // Keep browser open for inspection
  await page.pause();
});
