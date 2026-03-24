import { loadPresentation, type LoadedPresentation } from './pptx';
import type { EditorApi, LoadPresentationOptions, PresentationModel, RenderOptions, SlideNode, TextNode } from './types';

const NS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main'
} as const;

function emuToPx(value: number, slideExtent: number, pixelExtent: number): number {
  return (value / slideExtent) * pixelExtent;
}

function pxToEmu(value: number, pixelExtent: number, slideExtent: number): number {
  return (value / pixelExtent) * slideExtent;
}

function paragraphsFromText(text: string): string[] {
  return text.split(/\r?\n/);
}

function getXmlNodeElement(doc: XMLDocument, elementId: string): Element | null {
  const cNvPrList = Array.from(doc.getElementsByTagNameNS(NS.p, 'cNvPr'));
  const cNvPr = cNvPrList.find((candidate) => candidate.getAttribute('id') === elementId);
  if (!cNvPr) return null;
  let current: Element | null = cNvPr;
  while (current && !['sp', 'pic'].includes(current.localName)) {
    current = current.parentElement;
  }
  return current;
}

function firstChild(parent: Element | null | undefined, ns: string, tag: string): Element | null {
  if (!parent) return null;
  return Array.from(parent.children).find((child) => child.namespaceURI === ns && child.localName === tag) ?? null;
}

function setFrameOnElement(target: Element | null, frame: SlideNode['frame']): void {
  if (!target) return;
  const spPr = firstChild(target, NS.p, 'spPr');
  const xfrm = firstChild(spPr, NS.a, 'xfrm');
  const off = firstChild(xfrm, NS.a, 'off');
  const ext = firstChild(xfrm, NS.a, 'ext');
  off?.setAttribute('x', String(Math.round(frame.x)));
  off?.setAttribute('y', String(Math.round(frame.y)));
  ext?.setAttribute('cx', String(Math.round(frame.cx)));
  ext?.setAttribute('cy', String(Math.round(frame.cy)));
}

function replaceTextBody(target: Element | null, text: string): void {
  if (!target) return;
  const txBody = firstChild(target, NS.p, 'txBody');
  if (!txBody) return;
  const bodyPr = firstChild(txBody, NS.a, 'bodyPr');
  const lstStyle = firstChild(txBody, NS.a, 'lstStyle');
  const templateParagraph = firstChild(txBody, NS.a, 'p');
  Array.from(txBody.children)
    .filter((child) => child.namespaceURI === NS.a && child.localName === 'p')
    .forEach((child) => child.remove());

  for (const paragraphText of paragraphsFromText(text)) {
    const paragraph = templateParagraph?.cloneNode(true) as Element | null;
    const p = paragraph ?? txBody.ownerDocument.createElementNS(NS.a, 'a:p');
    Array.from(p.children).forEach((child) => child.remove());
    const run = txBody.ownerDocument.createElementNS(NS.a, 'a:r');
    const runProps = templateParagraph?.getElementsByTagNameNS(NS.a, 'rPr')[0]?.cloneNode(true) as Element | undefined;
    run.append(runProps ?? txBody.ownerDocument.createElementNS(NS.a, 'a:rPr'));
    const textNode = txBody.ownerDocument.createElementNS(NS.a, 'a:t');
    textNode.textContent = paragraphText;
    run.append(textNode);
    p.append(run);
    const endPara = templateParagraph?.getElementsByTagNameNS(NS.a, 'endParaRPr')[0]?.cloneNode(true) as Element | undefined;
    p.append(endPara ?? txBody.ownerDocument.createElementNS(NS.a, 'a:endParaRPr'));
    txBody.append(p);
  }

  if (!bodyPr) {
    txBody.prepend(txBody.ownerDocument.createElementNS(NS.a, 'a:bodyPr'));
  }
  if (!lstStyle) {
    const body = firstChild(txBody, NS.a, 'bodyPr');
    body?.after(txBody.ownerDocument.createElementNS(NS.a, 'a:lstStyle'));
  }
}

export class PptEditor implements EditorApi {
  readonly model: PresentationModel;

  private readonly loaded: LoadedPresentation;
  private readonly options: Required<RenderOptions>;
  private container: HTMLElement | null = null;
  private selectedNodeId: string | null = null;
  private editingNodeId: string | null = null;
  private nodeElementMap = new Map<string, HTMLElement>();
  private dragState: { nodeId: string; pointerId: number; originX: number; originY: number } | null = null;

  constructor(loaded: LoadedPresentation, options: RenderOptions = {}) {
    this.loaded = loaded;
    this.model = loaded.model;
    this.options = {
      slidePixelWidth: options.slidePixelWidth ?? loaded.model.preview?.slides[0]?.width ?? 1280,
      showOverlayFrames: options.showOverlayFrames ?? false
    };
  }

  mount(container: HTMLElement): void {
    this.container = container;
    container.classList.add('ppt-editor');
    this.render();
  }

  destroy(): void {
    if (this.container) {
      this.container.innerHTML = '';
    }
    this.nodeElementMap.clear();
    this.container = null;
  }

  setSelectedNode(nodeId: string | null): void {
    this.selectedNodeId = nodeId;
    this.syncSelectionState();
  }

  updateText(nodeId: string, text: string): void {
    const node = this.findNode(nodeId);
    if (!node || node.kind !== 'text') return;
    node.text = text;
    node.paragraphs = paragraphsFromText(text).map((paragraph) => ({ text: paragraph }));
    const element = this.nodeElementMap.get(nodeId);
    if (element) {
      this.renderTextNode(element, node, element.dataset.previewBacked === 'true');
    }
  }

  moveNode(nodeId: string, deltaXEmu: number, deltaYEmu: number): void {
    const node = this.findNode(nodeId);
    if (!node) return;
    node.frame.x += deltaXEmu;
    node.frame.y += deltaYEmu;
    this.positionNodeElement(nodeId, node);
  }

  async exportPptx(): Promise<Blob> {
    const serializer = new XMLSerializer();
    for (const slide of this.model.slides) {
      const doc = this.loaded.slideXmlDocs.get(slide.path);
      if (!doc) continue;
      for (const node of slide.nodes) {
        const xmlNode = getXmlNodeElement(doc, node.source.elementId);
        setFrameOnElement(xmlNode, node.frame);
        if (node.kind === 'text') {
          replaceTextBody(xmlNode, node.text);
        }
      }
      this.loaded.zip.file(slide.path, serializer.serializeToString(doc));
    }
    const bytes = await this.loaded.zip.generateAsync({ type: 'uint8array' });
    const copy = new Uint8Array(bytes.byteLength);
    copy.set(bytes);
    return new Blob([copy.buffer], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
  }

  private findNode(nodeId: string): SlideNode | undefined {
    for (const slide of this.model.slides) {
      const node = slide.nodes.find((candidate) => candidate.id === nodeId);
      if (node) return node;
    }
    return undefined;
  }

  private render(): void {
    if (!this.container) return;
    this.container.innerHTML = '';
    this.nodeElementMap.clear();

    for (const slide of this.model.slides) {
      const previewSlide = this.model.preview?.slides.find((candidate) => candidate.index === slide.index);
      const width = previewSlide?.width ?? this.options.slidePixelWidth;
      const height = previewSlide?.height ?? Math.round((width * this.model.size.cy) / this.model.size.cx);
      const slideElement = document.createElement('section');
      slideElement.className = 'ppt-slide';
      slideElement.dataset.slideIndex = String(slide.index);
      slideElement.style.width = `${width}px`;
      slideElement.style.height = `${height}px`;

      if (previewSlide) {
        const previewImage = document.createElement('img');
        previewImage.className = 'ppt-slide__preview';
        previewImage.src = previewSlide.dataUrl;
        previewImage.alt = `Slide ${slide.index} preview`;
        previewImage.draggable = false;
        slideElement.append(previewImage);
      }

      const overlay = document.createElement('div');
      overlay.className = 'ppt-slide__overlay';
      slideElement.append(overlay);

      for (const node of slide.nodes) {
        if (node.kind === 'shape' && previewSlide) {
          continue;
        }
        const nodeElement = this.createNodeElement(node, width, height, Boolean(previewSlide));
        overlay.append(nodeElement);
        this.nodeElementMap.set(node.id, nodeElement);
      }

      this.container.append(slideElement);
    }

    this.syncSelectionState();
  }

  private createNodeElement(node: SlideNode, slidePixelWidth: number, slidePixelHeight: number, previewBacked: boolean): HTMLElement {
    const element = document.createElement(node.kind === 'text' ? 'div' : 'div');
    element.className = `ppt-node ppt-node--${node.kind}`;
    element.dataset.nodeId = node.id;
    element.dataset.previewBacked = String(previewBacked);
    this.positionNodeElement(node.id, node, element, slidePixelWidth, slidePixelHeight);

    if (node.kind === 'text') {
      this.renderTextNode(element, node, previewBacked);
    } else if (node.kind === 'image') {
      this.renderImageNode(element, node, previewBacked);
    } else {
      this.renderShapeNode(element, node, previewBacked);
    }

    element.addEventListener('click', (event) => {
      event.stopPropagation();
      this.selectedNodeId = node.id;
      this.syncSelectionState();
    });

    element.addEventListener('dblclick', (event) => {
      event.stopPropagation();
      if (node.kind !== 'text') return;
      this.editingNodeId = node.id;
      element.contentEditable = 'true';
      element.classList.add('is-editing');
      element.focus();
      document.getSelection()?.selectAllChildren(element);
    });

    element.addEventListener('input', () => {
      if (node.kind !== 'text') return;
      this.updateText(node.id, element.innerText);
    });

    element.addEventListener('blur', () => {
      if (node.kind !== 'text') return;
      if (this.editingNodeId !== node.id) return;
      this.editingNodeId = null;
      element.contentEditable = 'false';
      element.classList.remove('is-editing');
      if (previewBacked) {
        element.textContent = node.text;
      }
      this.syncSelectionState();
    });

    element.addEventListener('pointerdown', (event) => {
      if (this.editingNodeId === node.id) return;
      element.setPointerCapture(event.pointerId);
      this.dragState = {
        nodeId: node.id,
        pointerId: event.pointerId,
        originX: event.clientX,
        originY: event.clientY
      };
      this.selectedNodeId = node.id;
      this.syncSelectionState();
    });

    element.addEventListener('pointermove', (event) => {
      if (!this.dragState || this.dragState.pointerId !== event.pointerId || this.dragState.nodeId !== node.id) return;
      const deltaPxX = event.clientX - this.dragState.originX;
      const deltaPxY = event.clientY - this.dragState.originY;
      this.dragState.originX = event.clientX;
      this.dragState.originY = event.clientY;
      this.moveNode(
        node.id,
        pxToEmu(deltaPxX, slidePixelWidth, this.model.size.cx),
        pxToEmu(deltaPxY, slidePixelHeight, this.model.size.cy)
      );
    });

    element.addEventListener('pointerup', (event) => {
      if (this.dragState?.pointerId === event.pointerId) {
        this.dragState = null;
      }
    });

    return element;
  }

  private renderTextNode(element: HTMLElement, node: TextNode, previewBacked: boolean): void {
    element.classList.toggle('ppt-node--ghost', previewBacked && this.editingNodeId !== node.id);
    element.replaceChildren();
    for (const paragraph of node.paragraphs) {
      const paragraphElement = document.createElement('div');
      paragraphElement.className = 'ppt-node__text-paragraph';
      paragraphElement.textContent = paragraph.text;
      element.append(paragraphElement);
    }
    if (!node.paragraphs.length) {
      element.textContent = node.text;
    }
    element.contentEditable = 'false';
    element.style.whiteSpace = 'pre-wrap';
    element.style.color = node.style.color;
    element.style.fontFamily = node.style.fontFamily;
    element.style.fontSize = `${node.style.fontSizePx}px`;
    element.style.fontWeight = String(node.style.fontWeight);
    element.style.fontStyle = node.style.fontStyle;
    element.style.textAlign = node.style.textAlign;
    element.style.justifyContent = node.style.verticalAlign === 'center' ? 'center' : node.style.verticalAlign === 'end' ? 'flex-end' : 'flex-start';
    element.style.alignItems = node.style.textAlign === 'center' ? 'center' : node.style.textAlign === 'right' ? 'flex-end' : 'flex-start';
  }

  private renderImageNode(element: HTMLElement, node: Extract<SlideNode, { kind: 'image' }>, previewBacked: boolean): void {
    element.classList.toggle('ppt-node--ghost', previewBacked);
    const asset = this.model.imageAssets.get(node.assetId);
    if (!previewBacked && asset) {
      const image = document.createElement('img');
      image.src = asset.dataUrl;
      image.alt = node.name;
      image.draggable = false;
      image.className = 'ppt-node__image';
      element.append(image);
    }
  }

  private renderShapeNode(element: HTMLElement, node: Extract<SlideNode, { kind: 'shape' }>, previewBacked: boolean): void {
    element.classList.toggle('ppt-node--ghost', previewBacked);
    element.classList.add(`ppt-node--shape-${node.geometry}`);
  }

  private positionNodeElement(nodeId: string, node: SlideNode, explicitElement?: HTMLElement, explicitWidth?: number, explicitHeight?: number): void {
    const element = explicitElement ?? this.nodeElementMap.get(nodeId);
    if (!element) return;
    const slide = element.closest('.ppt-slide') as HTMLElement | null;
    const slideWidth = explicitWidth ?? slide?.clientWidth ?? this.options.slidePixelWidth;
    const slideHeight = explicitHeight ?? slide?.clientHeight ?? Math.round((slideWidth * this.model.size.cy) / this.model.size.cx);
    element.style.left = `${emuToPx(node.frame.x, this.model.size.cx, slideWidth)}px`;
    element.style.top = `${emuToPx(node.frame.y, this.model.size.cy, slideHeight)}px`;
    element.style.width = `${emuToPx(node.frame.cx, this.model.size.cx, slideWidth)}px`;
    element.style.height = `${emuToPx(node.frame.cy, this.model.size.cy, slideHeight)}px`;
  }

  private syncSelectionState(): void {
    for (const [nodeId, element] of this.nodeElementMap) {
      element.classList.toggle('is-selected', nodeId === this.selectedNodeId);
      element.classList.toggle('show-frame', this.options.showOverlayFrames);
      if (element.dataset.previewBacked === 'true' && this.editingNodeId !== nodeId && !element.classList.contains('is-selected')) {
        element.classList.add('ppt-node--ghost');
      } else if (element.dataset.previewBacked === 'true') {
        element.classList.remove('ppt-node--ghost');
      }
    }
  }
}

export async function createPptEditor(options: LoadPresentationOptions, renderOptions?: RenderOptions): Promise<PptEditor> {
  const loaded = await loadPresentation(options);
  return new PptEditor(loaded, renderOptions);
}
