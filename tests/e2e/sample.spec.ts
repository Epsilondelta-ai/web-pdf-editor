import { expect, test } from '@playwright/test';

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

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await expect.poll(() => page.evaluate(() => (window as typeof window & { __PPT_DEBUG__?: { slideCount: number } }).__PPT_DEBUG__?.slideCount ?? 0)).toBe(4);
});

test('renders the sample deck against exported references', async ({ page }) => {
  const debug = await page.evaluate(() => (window as typeof window & { __PPT_DEBUG__?: { slideCount: number; previewCount: number } }).__PPT_DEBUG__);
  expect(debug?.slideCount).toBe(4);
  expect(debug?.previewCount).toBe(4);

  for (const index of [1, 2, 3, 4]) {
    const slide = page.locator(`.ppt-slide[data-slide-index="${index}"] .ppt-slide__preview`);
    await expect(slide).toBeVisible();
    const screenshot = await slide.screenshot();
    const comparison = await compareWithReference(page, screenshot, `/sample/sample.${String(index).padStart(3, '0')}.jpeg`);
    expect(comparison.avgChannelDelta, `slide ${index} average delta`).toBeLessThan(18);
    expect(comparison.mismatchRatio, `slide ${index} mismatch ratio`).toBeLessThan(0.18);
  }
});

test('edits text and exports the updated pptx xml', async ({ page }) => {
  const mutation = await page.evaluate(() => {
    const editor = (window as typeof window & { __PPT_EDITOR__: import('../../src').PptEditor }).__PPT_EDITOR__;
    const slide = editor.model.slides[1];
    const node = slide.nodes.find((candidate) => candidate.kind === 'text' && candidate.text.includes('타이틀'));
    if (!node || node.kind !== 'text') {
      throw new Error('Target text node not found');
    }
    editor.updateText(node.id, 'E2E edited title');
    editor.moveNode(node.id, 180000, 120000);
    return { nodeId: node.id, x: node.frame.x, y: node.frame.y };
  });

  await page.getByRole('button', { name: 'Export edited PPTX' }).click();
  await expect.poll(() => page.evaluate(() => (window as typeof window & { __LAST_EXPORTED_XML__?: string }).__LAST_EXPORTED_XML__ ?? '')).toContain('E2E edited title');

  const exportedXml = await page.evaluate(() => (window as typeof window & { __LAST_EXPORTED_XML__?: string }).__LAST_EXPORTED_XML__ ?? '');
  expect(exportedXml).toContain('E2E edited title');
  expect(exportedXml).toContain(`x="${Math.round(mutation.x)}"`);
  expect(exportedXml).toContain(`y="${Math.round(mutation.y)}"`);
});
