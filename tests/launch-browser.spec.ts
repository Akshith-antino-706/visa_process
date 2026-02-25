/**
 * Debug utility: Launch a browser with the saved session and verify it's valid.
 * Usage: ts-node tests/launch-browser.spec.ts
 */

import { createDriver, loadSession, SESSION_FILE } from '../src/automation/driver-factory';
import * as readline from 'readline';

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
  console.log('[Launch] Creating browser with session...');
  const driver = await createDriver();

  try {
    await loadSession(driver, SESSION_FILE);
    console.log('[Launch] Browser opened');
    console.log('[Launch] Navigating to portal...');
    await driver.get('https://smart.gdrfad.gov.ae/SmartChannels_Th/');

    const url = await driver.getCurrentUrl();
    console.log('[Launch] URL:', url);

    if (url.includes('Login.aspx')) {
      console.error('[Launch] Session expired — run "npm run auth" first');
    } else {
      console.log('[Launch] Session is valid — logged in');
    }

    // Keep browser open for inspection
    await prompt('[Launch] Press ENTER to close the browser... ');
  } finally {
    await driver.quit().catch(() => {});
  }
}

main().catch(err => {
  console.error('[Launch] Error:', err);
  process.exit(1);
});
