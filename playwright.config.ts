import { defineConfig } from '@playwright/test';
import * as path from 'path';

const SESSION_FILE = path.resolve('auth/session.json');

export default defineConfig({
  testDir: './tests',
  fullyParallel: false,       // Run sequentially — portal forms are stateful
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  workers: 1,                 // Single worker to avoid overlapping sessions
  reporter: 'html',

  timeout: 120_000,          // 2 minutes per test (form fill can be slow)
  use: {
    trace: 'on-first-retry',
    headless: false,           // Always headed — needed for form visibility
    viewport: null,            // null = let the browser use its natural/maximized size
    launchOptions: {
      args: ['--start-maximized'],  // Start Chrome maximized (full screen)
    },
    screenshot: 'only-on-failure',
    actionTimeout: 20_000,    // 20 s per individual action/click
  },

  projects: [
    // ── Step 1: Manual login (run once with: npm run auth) ──────────────────
    {
      name: 'setup',
      testMatch: '**/auth.setup.ts',
      use: {
        // No storageState — user logs in fresh to solve CAPTCHA
      },
    },

    // ── Step 2: Visa application automation (uses saved session) ─────────────
    {
      name: 'visa-automation',
      testMatch: '**/visa-application.spec.ts',
      use: {
        storageState: SESSION_FILE,   // Inject saved cookies/session automatically
      },
    },
  ],
});
