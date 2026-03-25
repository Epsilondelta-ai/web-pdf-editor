export type Emu = number;

export interface SlideFrame {
  x: Emu;
  y: Emu;
  cx: Emu;
  cy: Emu;
}

export interface SourceRef {
  slidePath: string;
  elementId: string;
}

export interface TextParagraph {
  text: string;
}

export interface TextStyle {
  color: string;
  fontFamily: string;
  fontSizePx: number;
  fontWeight: number;
  fontStyle: 'normal' | 'italic';
  textAlign: 'left' | 'center' | 'right';
  verticalAlign: 'start' | 'center' | 'end';
}

export interface BaseNode {
  id: string;
  name: string;
  kind: 'text' | 'image' | 'shape';
  frame: SlideFrame;
  source: SourceRef;
}

export interface TextNode extends BaseNode {
  kind: 'text';
  paragraphs: TextParagraph[];
  text: string;
  style: TextStyle;
}

export interface ImageNode extends BaseNode {
  kind: 'image';
  assetId: string;
  contentType: string;
}

export interface ShapeNode extends BaseNode {
  kind: 'shape';
  geometry: 'rect' | 'line' | 'unsupported';
}

export type SlideNode = TextNode | ImageNode | ShapeNode;

export interface SlideModel {
  id: string;
  index: number;
  path: string;
  nodes: SlideNode[];
}

export interface PreviewSlide {
  index: number;
  width: number;
  height: number;
  dataUrl: string;
}

export interface PreviewDocument {
  type: 'images';
  slides: PreviewSlide[];
}

export interface PresentationModel {
  size: { cx: Emu; cy: Emu };
  slides: SlideModel[];
  imageAssets: Map<string, { dataUrl: string; contentType: string }>;
  preview?: PreviewDocument;
}

export interface LoadPresentationOptions {
  pptx: ArrayBuffer;
  previewImages?: ArrayBuffer[];
}

export interface RenderOptions {
  slidePixelWidth?: number;
}

export interface ViewerApi {
  model: PresentationModel;
  mount(container: HTMLElement): void;
  destroy(): void;
}
