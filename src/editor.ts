import { loadPresentation, type LoadedPresentation } from './pptx';
import type { LoadPresentationOptions, PresentationModel, RenderOptions, SlideNode, TextNode, ViewerApi } from './types';

function emuToPx(value: number, slideExtent: number, pixelExtent: number): number {
  return (value / slideExtent) * pixelExtent;
}

export class PptViewer implements ViewerApi {
  readonly model: PresentationModel;

  private readonly options: Required<RenderOptions>;
  private container: HTMLElement | null = null;

  constructor(loaded: LoadedPresentation, options: RenderOptions = {}) {
    this.model = loaded.model;
    this.options = {
      slidePixelWidth: options.slidePixelWidth ?? 1280
    };
  }

  mount(container: HTMLElement): void {
    this.container = container;
    container.classList.add('ppt-viewer');
    this.render();
  }

  destroy(): void {
    if (this.container) {
      this.container.classList.remove('ppt-viewer');
      this.container.innerHTML = '';
    }
    this.container = null;
  }

  private render(): void {
    if (!this.container) return;
    this.container.innerHTML = '';

    for (const slide of this.model.slides) {
      const width = this.options.slidePixelWidth;
      const height = Math.round((width * this.model.size.cy) / this.model.size.cx);
      const slideElement = document.createElement('section');
      slideElement.className = 'ppt-slide';
      slideElement.dataset.slideIndex = String(slide.index);
      slideElement.style.width = `${width}px`;
      slideElement.style.height = `${height}px`;

      const content = document.createElement('div');
      content.className = 'ppt-slide__content';
      slideElement.append(content);

      for (const node of slide.nodes) {
        if (node.kind === 'shape' && node.geometry === 'unsupported') {
          continue;
        }
        const nodeElement = this.createNodeElement(node, width, height);
        content.append(nodeElement);
      }

      this.container.append(slideElement);
    }
  }

  private createNodeElement(node: SlideNode, slidePixelWidth: number, slidePixelHeight: number): HTMLElement {
    const element = document.createElement('div');
    element.className = `ppt-node ppt-node--${node.kind}`;
    this.positionNodeElement(node, element, slidePixelWidth, slidePixelHeight);

    if (node.kind === 'text') {
      this.renderTextNode(element, node);
    } else if (node.kind === 'image') {
      this.renderImageNode(element, node);
    } else {
      this.renderShapeNode(element, node);
    }

    return element;
  }

  private renderTextNode(element: HTMLElement, node: TextNode): void {
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

  private renderImageNode(element: HTMLElement, node: Extract<SlideNode, { kind: 'image' }>): void {
    const asset = this.model.imageAssets.get(node.assetId);
    if (asset) {
      const image = document.createElement('img');
      image.src = asset.dataUrl;
      image.alt = node.name;
      image.draggable = false;
      image.className = 'ppt-node__image';
      element.append(image);
    }
  }

  private renderShapeNode(element: HTMLElement, node: Extract<SlideNode, { kind: 'shape' }>): void {
    element.classList.add(`ppt-node--shape-${node.geometry}`);
  }

  private positionNodeElement(node: SlideNode, element: HTMLElement, slideWidth: number, slideHeight: number): void {
    element.style.left = `${emuToPx(node.frame.x, this.model.size.cx, slideWidth)}px`;
    element.style.top = `${emuToPx(node.frame.y, this.model.size.cy, slideHeight)}px`;
    element.style.width = `${emuToPx(node.frame.cx, this.model.size.cx, slideWidth)}px`;
    element.style.height = `${emuToPx(node.frame.cy, this.model.size.cy, slideHeight)}px`;
  }
}

export async function createPptViewer(options: LoadPresentationOptions, renderOptions?: RenderOptions): Promise<PptViewer> {
  const loaded = await loadPresentation(options);
  return new PptViewer(loaded, renderOptions);
}
