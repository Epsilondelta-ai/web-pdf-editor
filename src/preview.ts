import type { PreviewDocument } from './types';

async function imageDataUrlFromArrayBuffer(buffer: ArrayBuffer): Promise<{ width: number; height: number; dataUrl: string }> {
  const blob = new Blob([buffer]);
  const dataUrl = await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error ?? new Error('Failed to read preview image.'));
    reader.onload = () => resolve(String(reader.result));
    reader.readAsDataURL(blob);
  });
  const image = new Image();
  image.src = dataUrl;
  await image.decode();
  return { width: image.width, height: image.height, dataUrl };
}

export async function renderPreviewImages(images: ArrayBuffer[]): Promise<PreviewDocument> {
  const slides = await Promise.all(
    images.map(async (image, index) => {
      const rendered = await imageDataUrlFromArrayBuffer(image);
      return {
        index: index + 1,
        width: rendered.width,
        height: rendered.height,
        dataUrl: rendered.dataUrl
      };
    })
  );
  return { type: 'images', slides };
}
