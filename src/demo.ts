import { createPptViewer, type PptViewer } from './index';
import './styles.css';

const viewer = document.querySelector<HTMLElement>('#viewer');
const status = document.querySelector<HTMLElement>('#status');
const sampleButton = document.querySelector<HTMLButtonElement>('#load-sample');
const pptxInput = document.querySelector<HTMLInputElement>('#pptx-input');
const pdfInput = document.querySelector<HTMLInputElement>('#pdf-input');
const initialDeck = new URL(window.location.href).searchParams.get('deck') ?? 'sample1';

let pptViewer: PptViewer | null = null;
let currentDeck = initialDeck;

function setStatus(message: string): void {
  if (status) status.textContent = message;
}

async function readFile(file: File): Promise<ArrayBuffer> {
  return file.arrayBuffer();
}

async function mountViewer(pptx: ArrayBuffer, pdf?: ArrayBuffer, previewImages?: ArrayBuffer[]): Promise<void> {
  if (!viewer) return;
  pptViewer?.destroy();
  pptViewer = await createPptViewer(
    {
      pptx,
      previewPdf: previewImages?.length ? undefined : pdf,
      previewImages
    },
    {
      slidePixelWidth: 1750
    }
  );
  pptViewer.mount(viewer);
  (window as typeof window & { __PPT_VIEWER__?: PptViewer; __PPT_DEBUG__?: unknown }).__PPT_VIEWER__ = pptViewer;
  (window as typeof window & { __PPT_VIEWER__?: PptViewer; __PPT_DEBUG__?: unknown }).__PPT_DEBUG__ = {
    deck: currentDeck,
    slideCount: pptViewer.model.slides.length,
    previewCount: pptViewer.model.preview?.slides.length ?? 0,
    previewType: pptViewer.model.preview?.type ?? 'none',
    nodeCounts: pptViewer.model.slides.map((slide) => ({
      index: slide.index,
      textNodes: slide.nodes.filter((node) => node.kind === 'text').length,
      imageNodes: slide.nodes.filter((node) => node.kind === 'image').length
    }))
  };
  setStatus(`Loaded ${pptViewer.model.slides.length} slides in viewer mode.`);
}

async function fetchBinary(path: string): Promise<ArrayBuffer | undefined> {
  const response = await fetch(path);
  if (!response.ok) return undefined;
  const contentType = response.headers.get('content-type') ?? '';
  if (contentType.includes('text/html')) return undefined;
  return response.arrayBuffer();
}

async function fetchDeckPreviewImages(deckName: string, maxImages = 200): Promise<ArrayBuffer[]> {
  const images: ArrayBuffer[] = [];
  for (let index = 1; index <= maxImages; index += 1) {
    const image = await fetchBinary(`/sample/${deckName}.${String(index).padStart(3, '0')}.jpeg`);
    if (!image) {
      if (images.length > 0) break;
      continue;
    }
    images.push(image);
  }
  return images;
}

async function loadBundledSample(deckName = currentDeck): Promise<void> {
  currentDeck = deckName;
  setStatus(`Loading ${deckName} assets...`);
  const [pptx, pdf, jpegPreviews] = await Promise.all([
    fetchBinary(`/sample/${deckName}.pptx`),
    fetchBinary(`/sample/${deckName}.pdf`),
    fetchDeckPreviewImages(deckName)
  ]);
  if (!pptx) {
    setStatus(`sample/${deckName}.pptx not found. Keep sample/ local only; tests will skip when absent.`);
    return;
  }
  await mountViewer(pptx, pdf, jpegPreviews);
}

sampleButton?.addEventListener('click', () => {
  void loadBundledSample();
});

pptxInput?.addEventListener('change', async () => {
  const pptxFile = pptxInput.files?.[0];
  if (!pptxFile) return;
  const pdfFile = pdfInput?.files?.[0];
  await mountViewer(await readFile(pptxFile), pdfFile ? await readFile(pdfFile) : undefined);
});

pdfInput?.addEventListener('change', async () => {
  const pptxFile = pptxInput?.files?.[0];
  const pdfFile = pdfInput.files?.[0];
  if (!pptxFile) return;
  await mountViewer(await readFile(pptxFile), pdfFile ? await readFile(pdfFile) : undefined);
});

void loadBundledSample(initialDeck);
