import { test, expect } from '@playwright/test';

test('browser launches and navigates', async ({ page }) => {
  console.log('[Check] Browser launched!');
  console.log('[Check] Going to Google...');
  await page.goto('https://www.google.com');
  console.log('[Check] URL:', page.url());
  expect(page.url()).toContain('google');
  console.log('[Check] PASSED');
});
