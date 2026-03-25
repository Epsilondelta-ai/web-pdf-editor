import { expect, test } from '@playwright/test';
import fs from 'node:fs';
import path from 'node:path';

const sampleDir = path.resolve(process.cwd(), 'sample');
const sampleDecks = fs.existsSync(sampleDir)
  ? fs
      .readdirSync(sampleDir)
      .filter((name) => name.endsWith('.pptx'))
      .map((name) => name.replace(/\.pptx$/i, ''))
      .sort()
  : [];

function referenceImagesFor(deck: string): string[] {
  return fs
    .readdirSync(sampleDir)
    .filter((name) => new RegExp(`^${deck}\\.\\d{3}\\.jpeg$`, 'i').test(name))
    .sort();
}

async function compareWithReference(page: import('@playwright/test').Page, pngBuffer: Buffer, referencePath: string) {
  return page.evaluate(
    async ({ actualBase64, referencePath: reference }) => {
      async function loadImage(src: string): Promise<HTMLImageElement> {
        const image = new Image();
        image.src = src;
        await image.decode();
        return image;
      }

      const [actualImage, referenceImage] = await Promise.all([
        loadImage(`data:image/png;base64,${actualBase64}`),
        loadImage(reference)
      ]);

      const width = referenceImage.naturalWidth;
      const height = referenceImage.naturalHeight;
      const actualCanvas = document.createElement('canvas');
      const referenceCanvas = document.createElement('canvas');
      actualCanvas.width = referenceCanvas.width = width;
      actualCanvas.height = referenceCanvas.height = height;
      const actualContext = actualCanvas.getContext('2d');
      const referenceContext = referenceCanvas.getContext('2d');
      if (!actualContext || !referenceContext) {
        throw new Error('Canvas context unavailable');
      }

      actualContext.drawImage(actualImage, 0, 0, width, height);
      referenceContext.drawImage(referenceImage, 0, 0, width, height);
      const actual = actualContext.getImageData(0, 0, width, height).data;
      const expected = referenceContext.getImageData(0, 0, width, height).data;

      let totalDelta = 0;
      let mismatchPixels = 0;
      for (let index = 0; index < actual.length; index += 4) {
        const channelDelta =
          (Math.abs(actual[index] - expected[index]) +
            Math.abs(actual[index + 1] - expected[index + 1]) +
            Math.abs(actual[index + 2] - expected[index + 2])) /
          3;
        totalDelta += channelDelta;
        if (channelDelta > 20) mismatchPixels += 1;
      }

      return {
        width,
        height,
        avgChannelDelta: totalDelta / (width * height),
        mismatchRatio: mismatchPixels / (width * height)
      };
    },
    { actualBase64: pngBuffer.toString('base64'), referencePath }
  );
}

for (const deck of sampleDecks) {
  const references = referenceImagesFor(deck);

  test.describe(`render ${deck}`, () => {
    test.beforeEach(async ({ page }) => {
      await page.goto(`/?deck=${deck}`);
      await expect
        .poll(() =>
          page.evaluate(() =>
            (window as typeof window & { __PPT_DEBUG__?: { deck: string; slideCount: number } }).__PPT_DEBUG__ ?? null
          )
        )
        .toMatchObject({ deck, slideCount: references.length });
    });

    test(`${deck} matches exported references`, async ({ page }) => {
      const debug = await page.evaluate(
        () =>
          (window as typeof window & {
            __PPT_DEBUG__?: { deck: string; slideCount: number; previewCount: number; previewType: string };
          }).__PPT_DEBUG__
      );
      expect(debug?.deck).toBe(deck);
      expect(debug?.slideCount).toBe(references.length);
      expect(debug?.previewCount).toBe(references.length);
      expect(debug?.previewType).toBe('images');

      for (const [index, reference] of references.entries()) {
        const slideNumber = index + 1;
        const slide = page.locator(`.ppt-slide[data-slide-index="${slideNumber}"] .ppt-slide__preview`);
        await expect(slide).toBeVisible();
        const screenshot = await slide.screenshot();
        const comparison = await compareWithReference(page, screenshot, `/sample/${reference}`);
        expect(comparison.avgChannelDelta, `${deck} slide ${slideNumber} average delta`).toBeLessThan(18);
        expect(comparison.mismatchRatio, `${deck} slide ${slideNumber} mismatch ratio`).toBeLessThan(0.18);
      }
    });
  });
}

test('sample1 exposes a viewer-only API', async ({ page }) => {
  test.skip(!sampleDecks.includes('sample1'), 'sample/sample1.pptx fixture not available');
  await page.goto('/?deck=sample1');
  await expect
    .poll(() =>
      page.evaluate(() =>
        (window as typeof window & { __PPT_DEBUG__?: { deck: string; slideCount: number } }).__PPT_DEBUG__ ?? null
      )
    )
    .toMatchObject({ deck: 'sample1', slideCount: 4 });

  await expect(page.locator('#show-overlay-frames')).toHaveCount(0);
  await expect(page.locator('#export-pptx')).toHaveCount(0);

  const apiShape = await page.evaluate(() => {
    const pptViewer = (window as typeof window & { __PPT_VIEWER__?: import('../../src').PptViewer }).__PPT_VIEWER__;
    if (!pptViewer) return null;
    return {
      hasModel: typeof pptViewer.model === 'object',
      hasMount: typeof pptViewer.mount === 'function',
      hasDestroy: typeof pptViewer.destroy === 'function',
      hasUpdateText: 'updateText' in pptViewer,
      hasMoveNode: 'moveNode' in pptViewer,
      hasExportPptx: 'exportPptx' in pptViewer
    };
  });

  expect(apiShape).toEqual({
    hasModel: true,
    hasMount: true,
    hasDestroy: true,
    hasUpdateText: false,
    hasMoveNode: false,
    hasExportPptx: false
  });
});
