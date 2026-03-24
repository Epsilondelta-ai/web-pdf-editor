import JSZip from 'jszip';
import { createPptEditor, type PptEditor } from './index';
import './styles.css';

const viewer = document.querySelector<HTMLElement>('#viewer');
const status = document.querySelector<HTMLElement>('#status');
const sampleButton = document.querySelector<HTMLButtonElement>('#load-sample');
const pptxInput = document.querySelector<HTMLInputElement>('#pptx-input');
const pdfInput = document.querySelector<HTMLInputElement>('#pdf-input');
const overlayToggle = document.querySelector<HTMLInputElement>('#show-overlay-frames');
const exportButton = document.querySelector<HTMLButtonElement>('#export-pptx');

let editor: PptEditor | null = null;
let lastExportUrl: string | null = null;

function setStatus(message: string): void {
  if (status) status.textContent = message;
}

async function readFile(file: File): Promise<ArrayBuffer> {
  return file.arrayBuffer();
}

async function mountEditor(pptx: ArrayBuffer, pdf?: ArrayBuffer, previewImages?: ArrayBuffer[]): Promise<void> {
  if (!viewer) return;
  editor?.destroy();
  editor = await createPptEditor(
    {
      pptx,
      previewPdf: previewImages?.length ? undefined : pdf,
      previewImages
    },
    {
      showOverlayFrames: overlayToggle?.checked ?? false,
      slidePixelWidth: 1750
    }
  );
  editor.mount(viewer);
  (window as typeof window & { __PPT_EDITOR__?: PptEditor; __PPT_DEBUG__?: unknown }).__PPT_EDITOR__ = editor;
  (window as typeof window & { __PPT_EDITOR__?: PptEditor; __PPT_DEBUG__?: unknown }).__PPT_DEBUG__ = {
    slideCount: editor.model.slides.length,
    previewCount: editor.model.preview?.slides.length ?? 0,
    previewType: editor.model.preview?.type ?? 'none',
    nodeCounts: editor.model.slides.map((slide) => ({
      index: slide.index,
      textNodes: slide.nodes.filter((node) => node.kind === 'text').length,
      imageNodes: slide.nodes.filter((node) => node.kind === 'image').length
    }))
  };
  setStatus(`Loaded ${editor.model.slides.length} slides.`);
}

async function fetchBinary(path: string): Promise<ArrayBuffer | undefined> {
  const response = await fetch(path);
  if (!response.ok) return undefined;
  return response.arrayBuffer();
}

async function loadBundledSample(): Promise<void> {
  setStatus('Loading sample assets...');
  const [pptx, pdf, ...jpegPreviews] = await Promise.all([
    fetchBinary('/sample/sample.pptx'),
    fetchBinary('/sample/sample.pdf'),
    fetchBinary('/sample/sample.001.jpeg'),
    fetchBinary('/sample/sample.002.jpeg'),
    fetchBinary('/sample/sample.003.jpeg'),
    fetchBinary('/sample/sample.004.jpeg')
  ]);
  if (!pptx) {
    setStatus('sample/sample.pptx not found. Keep sample/ local only; tests will skip when absent.');
    return;
  }
  const previewImages = jpegPreviews.every(Boolean) ? (jpegPreviews as ArrayBuffer[]) : undefined;
  await mountEditor(pptx, pdf, previewImages);
}

async function exportCurrentPptx(): Promise<void> {
  if (!editor) return;
  const blob = await editor.exportPptx();
  if (lastExportUrl) URL.revokeObjectURL(lastExportUrl);
  lastExportUrl = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = lastExportUrl;
  link.download = 'edited-sample.pptx';
  link.click();

  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  const slide2 = await zip.file('ppt/slides/slide2.xml')?.async('string');
  (window as typeof window & { __LAST_EXPORTED_XML__?: string }).__LAST_EXPORTED_XML__ = slide2;
  setStatus('Exported edited-sample.pptx');
}

sampleButton?.addEventListener('click', () => {
  void loadBundledSample();
});

overlayToggle?.addEventListener('change', async () => {
  if (!editor) return;
  const [pptx, pdf, ...jpegPreviews] = await Promise.all([
    fetchBinary('/sample/sample.pptx'),
    fetchBinary('/sample/sample.pdf'),
    fetchBinary('/sample/sample.001.jpeg'),
    fetchBinary('/sample/sample.002.jpeg'),
    fetchBinary('/sample/sample.003.jpeg'),
    fetchBinary('/sample/sample.004.jpeg')
  ]);
  if (pptx) {
    const previewImages = jpegPreviews.every(Boolean) ? (jpegPreviews as ArrayBuffer[]) : undefined;
    await mountEditor(pptx, pdf, previewImages);
  }
});

pptxInput?.addEventListener('change', async () => {
  const pptxFile = pptxInput.files?.[0];
  if (!pptxFile) return;
  const pdfFile = pdfInput?.files?.[0];
  await mountEditor(await readFile(pptxFile), pdfFile ? await readFile(pdfFile) : undefined);
});

pdfInput?.addEventListener('change', async () => {
  const pptxFile = pptxInput?.files?.[0];
  const pdfFile = pdfInput.files?.[0];
  if (!pptxFile) return;
  await mountEditor(await readFile(pptxFile), pdfFile ? await readFile(pdfFile) : undefined);
});

exportButton?.addEventListener('click', () => {
  void exportCurrentPptx();
});

void loadBundledSample();
