/**
 * Creates and manages Selenium WebDriver instances.
 * Supports parallel browser sessions with cookie-based session injection.
 */

import { Builder, WebDriver } from 'selenium-webdriver';
import * as chrome from 'selenium-webdriver/chrome';
import * as fs from 'fs';
import * as path from 'path';

export const SESSION_FILE = path.resolve('auth/session.json');

export interface SessionData {
  cookies: Array<{
    name: string;
    value: string;
    domain?: string;
    path?: string;
    expiry?: number;
    secure?: boolean;
    httpOnly?: boolean;
    sameSite?: string;
  }>;
  localStorage?: Record<string, string>;
  sessionStorage?: Record<string, string>;
}

/**
 * Creates a new Chrome WebDriver instance in headed mode (maximized).
 * Uses a dedicated temp profile in /tmp to stay fully isolated from user's Chrome.
 */
export async function createDriver(): Promise<{ driver: WebDriver; profileDir: string }> {
  const options = new chrome.Options();

  // Core stability flags
  options.addArguments('--no-sandbox');
  options.addArguments('--disable-dev-shm-usage');
  options.addArguments('--disable-gpu');
  options.addArguments('--start-maximized');
  options.addArguments('--lang=en-US');

  // Prevent Chrome from throttling background tabs/windows (critical for parallel workers)
  options.addArguments('--disable-background-timer-throttling');
  options.addArguments('--disable-backgrounding-occluded-windows');
  options.addArguments('--disable-renderer-backgrounding');

  // Prevent Chrome from killing tabs it thinks are flooding IPC channels
  options.addArguments('--disable-ipc-flooding-protection');

  // Reduce memory per instance (important for 20 parallel workers)
  options.addArguments('--renderer-process-limit=1');
  options.addArguments('--disable-extensions');
  options.addArguments('--disable-default-apps');
  options.addArguments('--disable-popup-blocking');

  // Disable unnecessary features that consume memory
  options.addArguments('--disable-features=OptimizationGuideModelDownloading,OptimizationHintsFetching,OptimizationTargetPrediction,OptimizationHints,TranslateUI');

  // Dedicated temp profile — completely isolated from user's Chrome
  const tmpDir = `/tmp/selenium-profiles/profile-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  fs.mkdirSync(tmpDir, { recursive: true });
  options.addArguments(`--user-data-dir=${tmpDir}`);

  const driver = await new Builder()
    .forBrowser('chrome')
    .setChromeOptions(options)
    .build();

  // Set timeouts
  await driver.manage().setTimeouts({
    implicit: 5000,
    pageLoad: 60000,
    script: 30000,
  });

  return { driver, profileDir: tmpDir };
}

/**
 * Loads a saved session (cookies) into a WebDriver instance.
 * The driver must first navigate to the target domain before cookies can be set.
 */
export async function loadSession(driver: WebDriver, sessionFile = SESSION_FILE): Promise<void> {
  if (!fs.existsSync(sessionFile)) {
    throw new Error(
      `Session not found: ${sessionFile}\n` +
      'Run "npm run auth" to log in manually and save your session first.'
    );
  }

  const raw = fs.readFileSync(sessionFile, 'utf-8');
  const session: SessionData = JSON.parse(raw);

  // Navigate to the domain first so we can set cookies
  await driver.get('https://smart.gdrfad.gov.ae/SmartChannels_Th/');
  await driver.manage().deleteAllCookies();

  // Set all cookies
  for (const cookie of session.cookies) {
    try {
      await driver.manage().addCookie({
        name: cookie.name,
        value: cookie.value,
        domain: cookie.domain,
        path: cookie.path || '/',
        expiry: cookie.expiry ? new Date(cookie.expiry * 1000) : undefined,
        secure: cookie.secure,
        httpOnly: cookie.httpOnly,
        sameSite: cookie.sameSite as 'Strict' | 'Lax' | 'None' | undefined,
      });
    } catch {
      // Some cookies may fail (cross-domain, etc.) — not fatal
    }
  }

  // Inject localStorage and sessionStorage if available
  if (session.localStorage) {
    for (const [key, value] of Object.entries(session.localStorage)) {
      await driver.executeScript(`localStorage.setItem(arguments[0], arguments[1]);`, key, value);
    }
  }
  if (session.sessionStorage) {
    for (const [key, value] of Object.entries(session.sessionStorage)) {
      await driver.executeScript(`sessionStorage.setItem(arguments[0], arguments[1]);`, key, value);
    }
  }

  // Refresh to apply cookies
  await driver.navigate().refresh();
}

/**
 * Saves the current session (cookies + storage) from a WebDriver instance.
 */
export async function saveSession(driver: WebDriver, sessionFile = SESSION_FILE): Promise<void> {
  const dir = path.dirname(sessionFile);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

  const cookies = await driver.manage().getCookies();

  const localStorage = await driver.executeScript<Record<string, string>>(`
    var items = {};
    for (var i = 0; i < localStorage.length; i++) {
      var key = localStorage.key(i);
      items[key] = localStorage.getItem(key);
    }
    return items;
  `);

  const sessionStorage = await driver.executeScript<Record<string, string>>(`
    var items = {};
    for (var i = 0; i < sessionStorage.length; i++) {
      var key = sessionStorage.key(i);
      items[key] = sessionStorage.getItem(key);
    }
    return items;
  `);

  const session: SessionData = {
    cookies: cookies.map(c => ({
      name: c.name,
      value: c.value,
      domain: c.domain,
      path: c.path,
      expiry: typeof c.expiry === 'number' ? c.expiry : c.expiry instanceof Date ? Math.floor(c.expiry.getTime() / 1000) : undefined,
      secure: c.secure,
      httpOnly: c.httpOnly,
      sameSite: c.sameSite,
    })),
    localStorage,
    sessionStorage,
  };

  fs.writeFileSync(sessionFile, JSON.stringify(session, null, 2));
}

/**
 * Safely quits a driver and cleans up its temporary profile.
 */
export async function quitDriver(driver: WebDriver, profileDir?: string): Promise<void> {
  try {
    await driver.quit();
  } catch {
    // Driver may already be closed
  }

  // Clean up temp profile to free disk space (important for 20+ workers over many runs)
  if (profileDir && fs.existsSync(profileDir)) {
    try {
      fs.rmSync(profileDir, { recursive: true, force: true });
    } catch {
      // Best effort cleanup
    }
  }
}
