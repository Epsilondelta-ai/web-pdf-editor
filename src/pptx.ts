import JSZip from 'jszip';
import type { ImageNode, LoadPresentationOptions, PresentationModel, PreviewDocument, ShapeNode, SlideFrame, SlideModel, SlideNode, TextNode, TextStyle } from './types';
import { renderPreviewImages } from './preview';

const NS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships'
} as const;

const EMUS_PER_INCH = 914400;
const PX_PER_POINT = 96 / 72;

interface FlatTransform {
  sx: number;
  sy: number;
  tx: number;
  ty: number;
}

export interface LoadedPresentation {
  model: PresentationModel;
  zip: JSZip;
  slideXmlDocs: Map<string, XMLDocument>;
}

function parseXml(source: string): XMLDocument {
  return new DOMParser().parseFromString(source, 'application/xml');
}

function childrenByTag(parent: Element | Document, ns: string, tag: string): Element[] {
  return Array.from(parent.childNodes).filter(
    (node): node is Element => node instanceof Element && node.namespaceURI === ns && node.localName === tag
  );
}

function firstChild(parent: Element | Document | null | undefined, ns: string, tag: string): Element | null {
  if (!parent) return null;
  return childrenByTag(parent, ns, tag)[0] ?? null;
}

function attrNumber(node: Element | null, name: string, fallback = 0): number {
  const value = node?.getAttribute(name);
  return value ? Number(value) : fallback;
}

function resolveTarget(basePath: string, target: string): string {
  const baseParts = basePath.split('/');
  baseParts.pop();
  const stack = [...baseParts];
  for (const piece of target.split('/')) {
    if (!piece || piece === '.') continue;
    if (piece === '..') {
      stack.pop();
      continue;
    }
    stack.push(piece);
  }
  return stack.join('/');
}

function relsPathFor(partPath: string): string {
  const pieces = partPath.split('/');
  const fileName = pieces.pop();
  return `${pieces.join('/')}/_rels/${fileName}.rels`;
}

function mapRelationships(doc: XMLDocument, basePath: string): Map<string, string> {
  const result = new Map<string, string>();
  for (const rel of Array.from(doc.getElementsByTagNameNS(NS.rel, 'Relationship'))) {
    const id = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    if (id && target) {
      result.set(id, resolveTarget(basePath, target));
    }
  }
  return result;
}

function defaultTransform(): FlatTransform {
  return { sx: 1, sy: 1, tx: 0, ty: 0 };
}

function composeTransform(parent: FlatTransform, local: FlatTransform): FlatTransform {
  return {
    sx: parent.sx * local.sx,
    sy: parent.sy * local.sy,
    tx: parent.tx + parent.sx * local.tx,
    ty: parent.ty + parent.sy * local.ty
  };
}

function frameFromXfrm(xfrm: Element | null, transform: FlatTransform): SlideFrame {
  const off = firstChild(xfrm, NS.a, 'off');
  const ext = firstChild(xfrm, NS.a, 'ext');
  return {
    x: transform.tx + attrNumber(off, 'x') * transform.sx,
    y: transform.ty + attrNumber(off, 'y') * transform.sy,
    cx: attrNumber(ext, 'cx') * transform.sx,
    cy: attrNumber(ext, 'cy') * transform.sy
  };
}

function groupTransform(group: Element, parent: FlatTransform): FlatTransform {
  const grpPr = firstChild(group, NS.p, 'grpSpPr');
  const xfrm = firstChild(grpPr, NS.a, 'xfrm');
  const extCx = attrNumber(firstChild(xfrm, NS.a, 'ext'), 'cx', 1);
  const extCy = attrNumber(firstChild(xfrm, NS.a, 'ext'), 'cy', 1);
  const chExtCx = Math.max(attrNumber(firstChild(xfrm, NS.a, 'chExt'), 'cx', 1), 1);
  const chExtCy = Math.max(attrNumber(firstChild(xfrm, NS.a, 'chExt'), 'cy', 1), 1);
  const sx = extCx / chExtCx;
  const sy = extCy / chExtCy;
  const local: FlatTransform = {
    sx,
    sy,
    tx: attrNumber(firstChild(xfrm, NS.a, 'off'), 'x') - attrNumber(firstChild(xfrm, NS.a, 'chOff'), 'x') * sx,
    ty: attrNumber(firstChild(xfrm, NS.a, 'off'), 'y') - attrNumber(firstChild(xfrm, NS.a, 'chOff'), 'y') * sy
  };
  return composeTransform(parent, local);
}

function encodeDataUrl(contentType: string, payload: Uint8Array): string {
  let binary = '';
  for (let index = 0; index < payload.length; index += 0x8000) {
    binary += String.fromCharCode(...payload.subarray(index, index + 0x8000));
  }
  return `data:${contentType};base64,${btoa(binary)}`;
}

function contentTypeFromPath(path: string): string {
  const lower = path.toLowerCase();
  if (lower.endsWith('.png')) return 'image/png';
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
  if (lower.endsWith('.svg')) return 'image/svg+xml';
  return 'application/octet-stream';
}

function getNodeName(element: Element, childTag: string): string {
  const nvTag = childTag === 'pic' ? 'nvPicPr' : 'nvSpPr';
  const nv = firstChild(element, NS.p, nvTag);
  const cNvPr = firstChild(nv, NS.p, 'cNvPr');
  return cNvPr?.getAttribute('name') ?? `${childTag}-${cNvPr?.getAttribute('id') ?? 'unknown'}`;
}

function getNodeId(element: Element, childTag: string, slideIndex: number): { globalId: string; localId: string } {
  const nvTag = childTag === 'pic' ? 'nvPicPr' : 'nvSpPr';
  const nv = firstChild(element, NS.p, nvTag);
  const cNvPr = firstChild(nv, NS.p, 'cNvPr');
  const localId = cNvPr?.getAttribute('id') ?? `${childTag}-${slideIndex}`;
  return { globalId: `slide-${slideIndex}-${localId}`, localId };
}

function firstDescendant(parent: Element | null | undefined, ns: string, tag: string): Element | null {
  if (!parent) return null;
  return parent.getElementsByTagNameNS(ns, tag)[0] ?? null;
}

function parseColor(fill: Element | null | undefined): string | null {
  const colorNode = fill ? Array.from(fill.children)[0] : null;
  if (!colorNode) return null;
  if (colorNode.localName === 'srgbClr') {
    const value = colorNode.getAttribute('val');
    return value ? `#${value}` : null;
  }
  if (colorNode.localName === 'sysClr') {
    const fallback = colorNode.getAttribute('lastClr');
    return fallback ? `#${fallback}` : null;
  }
  return null;
}

function cssTextAlign(alignment: string | null | undefined): TextStyle['textAlign'] {
  if (alignment === 'ctr') return 'center';
  if (alignment === 'r') return 'right';
  return 'left';
}

function cssVerticalAlign(anchor: string | null | undefined): TextStyle['verticalAlign'] {
  if (anchor === 'ctr' || anchor === 'mid') return 'center';
  if (anchor === 'b') return 'end';
  return 'start';
}

function fontFamilyWithFallback(typeface: string | null | undefined): string {
  const base = typeface?.trim() ? `"${typeface.trim()}"` : '';
  const fallbacks = ['"Malgun Gothic"', '"Apple SD Gothic Neo"', '"Noto Sans KR"', 'Arial', 'sans-serif'];
  return [base, ...fallbacks].filter(Boolean).join(', ');
}

function parseTextStyle(txBody: Element, fallbackColor = '#111827'): TextStyle {
  const paragraph = firstDescendant(txBody, NS.a, 'p');
  const paragraphProps = firstDescendant(paragraph, NS.a, 'pPr');
  const bodyProps = firstChild(txBody, NS.a, 'bodyPr');
  const runProps = firstDescendant(txBody, NS.a, 'rPr') ?? firstDescendant(txBody, NS.a, 'defRPr');
  const solidFill = firstDescendant(runProps, NS.a, 'solidFill') ?? firstDescendant(txBody, NS.a, 'solidFill');
  const fontSize = attrNumber(runProps, 'sz', 2400) / 100;
  const typeface =
    firstDescendant(runProps, NS.a, 'ea')?.getAttribute('typeface') ??
    firstDescendant(runProps, NS.a, 'latin')?.getAttribute('typeface') ??
    firstDescendant(txBody, NS.a, 'ea')?.getAttribute('typeface') ??
    firstDescendant(txBody, NS.a, 'latin')?.getAttribute('typeface') ??
    '';

  return {
    color: parseColor(solidFill) ?? fallbackColor,
    fontFamily: fontFamilyWithFallback(typeface),
    fontSizePx: Math.max(fontSize * PX_PER_POINT, 12),
    fontWeight: attrNumber(runProps, 'b', 0) ? 700 : 400,
    fontStyle: attrNumber(runProps, 'i', 0) ? 'italic' : 'normal',
    textAlign: cssTextAlign(paragraphProps?.getAttribute('algn')),
    verticalAlign: cssVerticalAlign(bodyProps?.getAttribute('anchor'))
  };
}

function parseTextNode(element: Element, slidePath: string, slideIndex: number, transform: FlatTransform): TextNode | null {
  const spPr = firstChild(element, NS.p, 'spPr');
  const txBody = firstChild(element, NS.p, 'txBody');
  if (!txBody || !spPr) return null;
  const xfrm = firstChild(spPr, NS.a, 'xfrm');
  const paragraphs = Array.from(txBody.getElementsByTagNameNS(NS.a, 'p')).map((paragraph) => ({
    text: Array.from(paragraph.getElementsByTagNameNS(NS.a, 't'))
      .map((textNode) => textNode.textContent ?? '')
      .join('')
  }));
  const { globalId, localId } = getNodeId(element, 'sp', slideIndex);
  return {
    id: globalId,
    kind: 'text',
    name: getNodeName(element, 'sp'),
    frame: frameFromXfrm(xfrm, transform),
    paragraphs,
    text: paragraphs.map((paragraph) => paragraph.text).join('\n'),
    style: parseTextStyle(txBody),
    source: { slidePath, elementId: localId }
  };
}

function parseShapeNode(element: Element, slidePath: string, slideIndex: number, transform: FlatTransform): ShapeNode | null {
  const spPr = firstChild(element, NS.p, 'spPr');
  if (!spPr) return null;
  const xfrm = firstChild(spPr, NS.a, 'xfrm');
  const prst = firstChild(spPr, NS.a, 'prstGeom')?.getAttribute('prst');
  const geometry = prst === 'rect' ? 'rect' : prst === 'line' ? 'line' : 'unsupported';
  const { globalId, localId } = getNodeId(element, 'sp', slideIndex);
  return {
    id: globalId,
    kind: 'shape',
    name: getNodeName(element, 'sp'),
    frame: frameFromXfrm(xfrm, transform),
    geometry,
    source: { slidePath, elementId: localId }
  };
}

function parseImageNode(
  element: Element,
  rels: Map<string, string>,
  assets: Map<string, { dataUrl: string; contentType: string }>,
  slidePath: string,
  slideIndex: number,
  transform: FlatTransform
): ImageNode | null {
  const blip = element.getElementsByTagNameNS(NS.a, 'blip')[0] ?? null;
  const rId = blip?.getAttributeNS(NS.r, 'embed') ?? blip?.getAttribute('r:embed');
  if (!rId) return null;
  const assetId = rels.get(rId);
  if (!assetId) return null;
  const spPr = firstChild(element, NS.p, 'spPr');
  const xfrm = firstChild(spPr, NS.a, 'xfrm');
  const asset = assets.get(assetId);
  const { globalId, localId } = getNodeId(element, 'pic', slideIndex);
  return {
    id: globalId,
    kind: 'image',
    name: getNodeName(element, 'pic'),
    frame: frameFromXfrm(xfrm, transform),
    assetId,
    contentType: asset?.contentType ?? contentTypeFromPath(assetId),
    source: { slidePath, elementId: localId }
  };
}

function parseNodes(
  container: Element,
  slidePath: string,
  slideIndex: number,
  rels: Map<string, string>,
  assets: Map<string, { dataUrl: string; contentType: string }>,
  transform: FlatTransform,
  nodes: SlideNode[]
): void {
  for (const child of Array.from(container.children)) {
    if (child.namespaceURI !== NS.p) continue;
    if (child.localName === 'sp') {
      const textNode = parseTextNode(child, slidePath, slideIndex, transform);
      if (textNode) {
        nodes.push(textNode);
        continue;
      }
      const shapeNode = parseShapeNode(child, slidePath, slideIndex, transform);
      if (shapeNode) {
        nodes.push(shapeNode);
      }
      continue;
    }
    if (child.localName === 'pic') {
      const imageNode = parseImageNode(child, rels, assets, slidePath, slideIndex, transform);
      if (imageNode) {
        nodes.push(imageNode);
      }
      continue;
    }
    if (child.localName === 'grpSp') {
      parseNodes(child, slidePath, slideIndex, rels, assets, groupTransform(child, transform), nodes);
    }
  }
}

async function readImageAssets(zip: JSZip): Promise<Map<string, { dataUrl: string; contentType: string }>> {
  const assets = new Map<string, { dataUrl: string; contentType: string }>();
  const mediaEntries = Object.keys(zip.files).filter((path) => path.startsWith('ppt/media/') && !zip.files[path].dir);
  for (const path of mediaEntries) {
    const payload = await zip.file(path)?.async('uint8array');
    if (!payload) continue;
    const contentType = contentTypeFromPath(path);
    assets.set(path, { contentType, dataUrl: encodeDataUrl(contentType, payload) });
  }
  return assets;
}

async function parsePresentation(pptx: ArrayBuffer, preview?: PreviewDocument): Promise<LoadedPresentation> {
  const zip = await JSZip.loadAsync(pptx);
  const presentationXml = await zip.file('ppt/presentation.xml')?.async('string');
  const presentationRelsXml = await zip.file('ppt/_rels/presentation.xml.rels')?.async('string');
  if (!presentationXml || !presentationRelsXml) {
    throw new Error('Missing required presentation XML parts.');
  }

  const presentationDoc = parseXml(presentationXml);
  const presentationRels = mapRelationships(parseXml(presentationRelsXml), 'ppt/presentation.xml');
  const slideSize = firstChild(presentationDoc.documentElement, NS.p, 'sldSz');
  const sldIdList = firstChild(presentationDoc.documentElement, NS.p, 'sldIdLst');
  const slideRefs = childrenByTag(sldIdList ?? presentationDoc.documentElement, NS.p, 'sldId');
  const imageAssets = await readImageAssets(zip);

  const slides: SlideModel[] = [];
  const slideXmlDocs = new Map<string, XMLDocument>();

  for (const [index, slideRef] of slideRefs.entries()) {
    const rId = slideRef.getAttributeNS(NS.r, 'id') ?? slideRef.getAttribute('r:id');
    if (!rId) continue;
    const slidePath = presentationRels.get(rId);
    if (!slidePath) continue;
    const slideXml = await zip.file(slidePath)?.async('string');
    if (!slideXml) continue;
    const slideDoc = parseXml(slideXml);
    slideXmlDocs.set(slidePath, slideDoc);

    const slideRelsXml = await zip.file(relsPathFor(slidePath))?.async('string');
    const slideRels = slideRelsXml ? mapRelationships(parseXml(slideRelsXml), slidePath) : new Map<string, string>();
    const cSld = firstChild(slideDoc.documentElement, NS.p, 'cSld');
    const spTree = firstChild(cSld, NS.p, 'spTree');
    const nodes: SlideNode[] = [];
    if (spTree) {
      parseNodes(spTree, slidePath, index + 1, slideRels, imageAssets, defaultTransform(), nodes);
    }
    slides.push({
      id: slideRef.getAttribute('id') ?? `slide-${index + 1}`,
      index: index + 1,
      path: slidePath,
      nodes
    });
  }

  return {
    model: {
      size: {
        cx: attrNumber(slideSize, 'cx', 10 * EMUS_PER_INCH),
        cy: attrNumber(slideSize, 'cy', 5.625 * EMUS_PER_INCH)
      },
      slides,
      imageAssets,
      preview
    },
    zip,
    slideXmlDocs
  };
}

export async function loadPresentation(options: LoadPresentationOptions): Promise<LoadedPresentation> {
  const preview: PreviewDocument | undefined = options.previewImages?.length ? await renderPreviewImages(options.previewImages) : undefined;
  return parsePresentation(options.pptx, preview);
}
