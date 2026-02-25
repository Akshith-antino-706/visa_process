import { defineConfig } from '@playwright/test';
export default defineConfig({
  testDir: './tests',
  fullyParallel: true,        // Run applicants in parallel across workers
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  workers: 2,                 // 2 parallel browser windows
  reporter: 'html',

  timeout: 600_000,          // 10 minutes per test (form fill + document upload + pauses)
  use: {
    trace: 'off',
    headless: false,           // Always headed — needed for form visibility
    viewport: null,            // null = let the browser use its natural/maximized size
    launchOptions: {
      args: ['--start-maximized'],  // Start Chrome maximized (full screen)
    },
    screenshot: 'only-on-failure',
    video: 'off',
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
    },

    // ── Quick browser launch test ─────────────────────────────────────────────
    {
      name: 'launch',
      testMatch: '**/launch-browser.spec.ts',
    },
  ],
});
