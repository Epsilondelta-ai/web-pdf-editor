import { chromium } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import { spawn } from 'node:child_process';

const rootDir = process.cwd();
const sampleDir = path.resolve(rootDir, 'sample');
const outputDir = path.resolve(rootDir, 'test-results/visual-review');
const port = Number(process.env.VISUAL_PORT ?? 4173);
const baseUrl = process.env.VISUAL_BASE_URL ?? `http://127.0.0.1:${port}`;

async function exists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function waitForServer(url, timeoutMs = 120000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {}
    await new Promise((resolve) => setTimeout(resolve, 1000));
  }
  throw new Error(`Timed out waiting for ${url}`);
}

function spawnDevServer() {
  return spawn('npm', ['run', 'dev', '--', '--port', String(port)], {
    cwd: rootDir,
    stdio: 'inherit',
    shell: true
  });
}

function toDataUrl(mimeType, buffer) {
  return `data:${mimeType};base64,${buffer.toString('base64')}`;
}

async function generateComparison(page, actualBuffer, referenceBuffer) {
  const actualDataUrl = toDataUrl('image/png', actualBuffer);
  const referenceDataUrl = toDataUrl('image/jpeg', referenceBuffer);
  return page.evaluate(async ({ actual, reference }) => {
    async function load(src) {
      const image = new Image();
      image.src = src;
      await image.decode();
      return image;
    }

    const [actualImage, referenceImage] = await Promise.all([load(actual), load(reference)]);
    const width = referenceImage.naturalWidth;
    const height = referenceImage.naturalHeight;

    const actualCanvas = document.createElement('canvas');
    actualCanvas.width = width;
    actualCanvas.height = height;
    const actualContext = actualCanvas.getContext('2d');
    actualContext.drawImage(actualImage, 0, 0, width, height);

    const referenceCanvas = document.createElement('canvas');
    referenceCanvas.width = width;
    referenceCanvas.height = height;
    const referenceContext = referenceCanvas.getContext('2d');
    referenceContext.drawImage(referenceImage, 0, 0, width, height);

    const actualPixels = actualContext.getImageData(0, 0, width, height);
    const referencePixels = referenceContext.getImageData(0, 0, width, height);
    const diffPixels = new ImageData(width, height);

    let totalDelta = 0;
    let mismatchPixels = 0;
    for (let index = 0; index < actualPixels.data.length; index += 4) {
      const dr = Math.abs(actualPixels.data[index] - referencePixels.data[index]);
      const dg = Math.abs(actualPixels.data[index + 1] - referencePixels.data[index + 1]);
      const db = Math.abs(actualPixels.data[index + 2] - referencePixels.data[index + 2]);
      const avg = (dr + dg + db) / 3;
      totalDelta += avg;
      const hot = avg > 20;
      if (hot) mismatchPixels += 1;
      diffPixels.data[index] = hot ? 255 : 0;
      diffPixels.data[index + 1] = hot ? 64 : 0;
      diffPixels.data[index + 2] = hot ? 64 : 0;
      diffPixels.data[index + 3] = hot ? 255 : 30;
    }

    const diffCanvas = document.createElement('canvas');
    diffCanvas.width = width;
    diffCanvas.height = height;
    diffCanvas.getContext('2d').putImageData(diffPixels, 0, 0);

    const composite = document.createElement('canvas');
    composite.width = width * 3;
    composite.height = height + 64;
    const compositeContext = composite.getContext('2d');
    compositeContext.fillStyle = '#101828';
    compositeContext.fillRect(0, 0, composite.width, composite.height);
    compositeContext.fillStyle = '#ffffff';
    compositeContext.font = '28px sans-serif';
    compositeContext.fillText('Actual render', 32, 40);
    compositeContext.fillText('Reference JPEG', width + 32, 40);
    compositeContext.fillText('Hotspot diff', width * 2 + 32, 40);
    compositeContext.drawImage(actualCanvas, 0, 64);
    compositeContext.drawImage(referenceCanvas, width, 64);
    compositeContext.drawImage(diffCanvas, width * 2, 64);

    return {
      width,
      height,
      avgChannelDelta: totalDelta / (width * height),
      mismatchRatio: mismatchPixels / (width * height),
      compositeDataUrl: composite.toDataURL('image/png')
    };
  }, { actual: actualDataUrl, reference: referenceDataUrl });
}

function dataUrlToBuffer(dataUrl) {
  const base64 = dataUrl.split(',')[1] ?? '';
  return Buffer.from(base64, 'base64');
}

async function main() {
  const samplePptx = path.join(sampleDir, 'sample.pptx');
  if (!(await exists(samplePptx))) {
    throw new Error('sample/sample.pptx not found.');
  }

  await fs.mkdir(outputDir, { recursive: true });

  const server = spawnDevServer();
  try {
    await waitForServer(baseUrl);
    const browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 2200, height: 1400 }, deviceScaleFactor: 1 });
    await page.goto(baseUrl);
    await page.waitForFunction(() => (window.__PPT_DEBUG__?.slideCount ?? 0) === 4);

    const summary = [];
    for (const index of [1, 2, 3, 4]) {
      const slide = page.locator(`.ppt-slide[data-slide-index="${index}"]`);
      await slide.scrollIntoViewIfNeeded();
      const actualBuffer = await slide.screenshot();
      const referencePath = path.join(sampleDir, `sample.${String(index).padStart(3, '0')}.jpeg`);
      const referenceBuffer = await fs.readFile(referencePath);
      const comparison = await generateComparison(page, actualBuffer, referenceBuffer);

      await fs.writeFile(path.join(outputDir, `slide-${index}.actual.png`), actualBuffer);
      await fs.writeFile(path.join(outputDir, `slide-${index}.comparison.png`), dataUrlToBuffer(comparison.compositeDataUrl));
      summary.push({
        slide: index,
        avgChannelDelta: comparison.avgChannelDelta,
        mismatchRatio: comparison.mismatchRatio,
        actual: `slide-${index}.actual.png`,
        reference: path.relative(outputDir, referencePath),
        comparison: `slide-${index}.comparison.png`
      });
    }

    await fs.writeFile(path.join(outputDir, 'summary.json'), JSON.stringify(summary, null, 2));
    await browser.close();
  } finally {
    server.kill('SIGTERM');
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
