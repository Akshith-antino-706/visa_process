/**
 * Manual login & session save for Selenium.
 *
 * Opens a Chrome browser, navigates to the GDRFA login page, and waits
 * for you to log in manually and solve the CAPTCHA.
 *
 * Usage: npm run auth
 *
 * After logging in:
 *   1. Press ENTER in this terminal to save the session.
 *   2. Open a NEW terminal and run: npm test
 *   3. When npm test finishes, press ENTER again here to close the browser.
 */

import { createDriver, saveSession, SESSION_FILE } from '../src/automation/driver-factory';
import * as readline from 'readline';
import * as fs from 'fs';

const PORTAL_URL  = 'https://smart.gdrfad.gov.ae/SmartChannels_Th/Login.aspx';
const PORTAL_HOME = 'https://smart.gdrfad.gov.ae/SmartChannels_Th/';
const KEEP_ALIVE_INTERVAL_MS = 4 * 60 * 1000; // 4 min

function prompt(question: string): Promise<string> {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise(resolve => {
    rl.question(question, answer => {
      rl.close();
      resolve(answer);
    });
  });
}

async function main() {
  console.log('[Auth] Launching Chrome browser...');
  const { driver } = await createDriver();

  try {
    // 1. Clear stale cookies and navigate to login
    await driver.manage().deleteAllCookies();
    await driver.get(PORTAL_URL);

    const url = await driver.getCurrentUrl();
    if (!url.includes('Login.aspx')) {
      console.warn('[Auth] Portal redirected away from Login.aspx — retrying...');
      await driver.get(PORTAL_URL);
    }

    // 2. Wait for manual login
    console.log('\n[Auth] ─────────────────────────────────────────────────────────');
    console.log('[Auth] Browser is open. Please:');
    console.log('[Auth]   1. Enter your username & password');
    console.log('[Auth]   2. Solve the CAPTCHA');
    console.log('[Auth]   3. Click Login and wait until the dashboard loads');
    console.log('[Auth]   4. Then come back here and press ENTER');
    console.log('[Auth] ─────────────────────────────────────────────────────────\n');

    await prompt('[Auth] Press ENTER after you have logged in successfully... ');

    // 3. Save the session
    console.log('\n[Auth] Saving session...');
    await saveSession(driver, SESSION_FILE);
    console.log(`[Auth] Session saved → ${SESSION_FILE}`);

    if (!fs.existsSync(SESSION_FILE)) {
      throw new Error('Session file was not created');
    }

    // 4. Start keep-alive
    const keepAlive = setInterval(async () => {
      try {
        await driver.executeScript(`fetch('${PORTAL_HOME}', { method: 'GET', credentials: 'include' });`);
        console.log('[KeepAlive] Session ping sent — idle timer reset.');
      } catch {
        // Ignore — browser may be mid-navigation
      }
    }, KEEP_ALIVE_INTERVAL_MS);

    // 5. Stay open so session remains alive while npm test runs
    console.log('\n[Auth] ─────────────────────────────────────────────────────────');
    console.log('[Auth] Session is LIVE — keep-alive pinging every 4 min.');
    console.log('[Auth] ► Open a NEW terminal and run:  npm test');
    console.log('[Auth] ► Once npm test finishes, press ENTER here to close.');
    console.log('[Auth] ─────────────────────────────────────────────────────────\n');

    await prompt('[Auth] Press ENTER to close the auth browser... ');

    // 6. Clean up
    clearInterval(keepAlive);
    console.log('[Auth] Keep-alive stopped. Auth browser closing.');
  } finally {
    await driver.quit().catch(() => {});
  }
}

main().catch(err => {
  console.error('[Auth] Fatal error:', err);
  process.exit(1);
});
