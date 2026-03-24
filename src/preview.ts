import { GlobalWorkerOptions, getDocument } from 'pdfjs-dist';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.mjs?url';
import type { PreviewDocument } from './types';

GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

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

export async function renderPdfPreview(pdfBuffer: ArrayBuffer, targetWidth = 1750): Promise<PreviewDocument> {
  const pdf = await getDocument({ data: pdfBuffer }).promise;
  const slides = [] as PreviewDocument['slides'];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const baseViewport = page.getViewport({ scale: 1 });
    const scale = targetWidth / baseViewport.width;
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    if (!context) {
      throw new Error('2D canvas context is unavailable.');
    }
    canvas.width = Math.round(viewport.width);
    canvas.height = Math.round(viewport.height);
    await page.render({ canvas, canvasContext: context, viewport }).promise;
    slides.push({
      index: pageNumber,
      width: canvas.width,
      height: canvas.height,
      dataUrl: canvas.toDataURL('image/png')
    });
  }

  return { type: 'pdf', slides };
}
