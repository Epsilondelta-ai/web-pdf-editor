import { defineConfig } from '@playwright/test';
import fs from 'node:fs';
import path from 'node:path';

const sampleDir = process.env.PPT_SAMPLE_DIR ?? path.resolve(process.cwd(), 'sample');
const hasSamples =
  fs.existsSync(sampleDir) &&
  fs.readdirSync(sampleDir, { withFileTypes: true }).some((entry) => entry.isFile() && entry.name.endsWith('.pptx'));

export default defineConfig({
  testDir: './tests/e2e',
  timeout: 120_000,
  expect: {
    timeout: 15_000
  },
  use: {
    baseURL: 'http://127.0.0.1:4173',
    viewport: { width: 1900, height: 1180 },
    deviceScaleFactor: 1,
    trace: 'retain-on-failure'
  },
  webServer: {
    command: 'npm run dev -- --port 4173',
    port: 4173,
    reuseExistingServer: !process.env.CI,
    timeout: 120_000
  },
  reporter: [['list']],
  grepInvert: hasSamples ? undefined : /./
});
