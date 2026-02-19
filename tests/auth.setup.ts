import { test as setup, expect } from '@playwright/test';
import * as path from 'path';
import * as fs from 'fs';

const PORTAL_URL  = 'https://smart.gdrfad.gov.ae/SmartChannels_Th/Login.aspx';
const PORTAL_HOME = 'https://smart.gdrfad.gov.ae/SmartChannels_Th/';
export const SESSION_FILE = path.resolve('auth/session.json');

const KEEP_ALIVE_INTERVAL_MS = 4 * 60 * 1000; // 4 min — well within the 15-min idle timeout

setup('Manual login & save session', async ({ page }) => {
  // Allow this test to stay open indefinitely while keep-alive is running
  setup.setTimeout(0);

  // 1. Clear stale cookies so the portal cannot auto-redirect to the dashboard
  await page.context().clearCookies();

  // 2. Navigate to the login page
  await page.goto(PORTAL_URL, { waitUntil: 'domcontentloaded' });
  if (!page.url().includes('Login.aspx')) {
    console.warn('[Auth] Portal redirected away from Login.aspx — retrying...');
    await page.goto(PORTAL_URL, { waitUntil: 'domcontentloaded' });
  }

  // 3. Pause — Playwright Inspector opens.
  //    - Enter username & password
  //    - Solve the CAPTCHA
  //    - Click Login and wait until the dashboard loads
  //    Then click "Resume" in the Inspector toolbar.
  await page.pause();

  // 4. Save the session (cookies + localStorage + sessionStorage)
  console.log('\n[Auth] Saving session...');
  await page.context().storageState({ path: SESSION_FILE });
  console.log(`[Auth] Session saved → ${SESSION_FILE}`);
  expect(fs.existsSync(SESSION_FILE)).toBeTruthy();

  // 5. Start keep-alive immediately — ping the portal every 4 min so the
  //    server-side idle timer never reaches 15 minutes before npm test runs.
  const keepAlive = setInterval(async () => {
    try {
      await page.evaluate(async (url: string) => {
        await fetch(url, { method: 'GET', credentials: 'include' });
      }, PORTAL_HOME);
      console.log('[KeepAlive] Session ping sent — idle timer reset.');
    } catch {
      // Ignore — browser may be mid-navigation
    }
  }, KEEP_ALIVE_INTERVAL_MS);

  // 6. Stay paused so the browser (and session) remain alive while npm test runs
  console.log('\n[Auth] ─────────────────────────────────────────────────────────');
  console.log('[Auth] Session is LIVE — keep-alive pinging every 4 min.');
  console.log('[Auth] ► Open a NEW terminal and run:  npm test');
  console.log('[Auth] ► Once npm test finishes, click Resume here to close.');
  console.log('[Auth] ─────────────────────────────────────────────────────────\n');
  await page.pause();

  // 7. User clicked Resume — clean up and exit
  clearInterval(keepAlive);
  console.log('[Auth] Keep-alive stopped. Auth browser closed.');
});
