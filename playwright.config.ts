import { defineConfig } from '@playwright/test';

// Playwright is configured to drive the system Chrome install (channel: 'chrome')
// because the Playwright-managed Chromium bundle does not yet publish binaries
// for ubuntu26.04 on this host. The tests themselves are plain DOM/fetch and
// don't depend on anything Chrome-specific.
//
// Port 3002 is used for the dev server so parallel runs in sibling repos
// (docxjs on 8765, pptxjs on 3001) don't collide.
export default defineConfig({
  testDir: './tests',
  testMatch: ['**/*.spec.ts', '**/*.spec.js'],
  fullyParallel: true,
  retries: 0,
  workers: 4,
  reporter: 'list',
  use: {
    baseURL: 'http://localhost:3002',
    trace: 'on-first-retry',
    screenshot: 'only-on-failure',
  },
  projects: [
    {
      name: 'chrome',
      use: {
        browserName: 'chromium',
        channel: 'chrome',
      },
    },
  ],
  webServer: {
    command: 'PORT=3002 node scripts/dev-server.mjs',
    url: 'http://localhost:3002/tests/harness.html',
    reuseExistingServer: !process.env.CI,
    timeout: 30_000,
  },
});
