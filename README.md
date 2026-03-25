# web-ppt-editor

Browser PPTX viewer library with optional PDF or image-based slide previews.

## Install

```bash
npm install web-ppt-editor
```

## Quick start

```ts
import 'web-ppt-editor/styles.css';
import { createPptViewer } from 'web-ppt-editor';

const container = document.querySelector('#viewer');
const pptx = await fetch('/slides/deck.pptx').then((response) => response.arrayBuffer());
const previewPdf = await fetch('/slides/deck.pdf').then((response) => response.arrayBuffer());

if (!container) {
  throw new Error('Missing #viewer container');
}

const viewer = await createPptViewer(
  {
    pptx,
    previewPdf
  },
  {
    slidePixelWidth: 1600
  }
);

viewer.mount(container);
```

## API

### `createPptViewer(options, renderOptions?)`

Creates a browser viewer instance and loads the presentation model.

`options`:
- `pptx: ArrayBuffer` - required `.pptx` file bytes.
- `previewPdf?: ArrayBuffer` - optional exported PDF used as the slide preview layer.
- `previewImages?: ArrayBuffer[]` - optional per-slide preview images. When present, they take precedence over `previewPdf`.

`renderOptions`:
- `slidePixelWidth?: number` - target slide width when no preview asset defines the width.

### `viewer.mount(container)`

Renders slides into the provided HTML element.

### `viewer.destroy()`

Clears the mounted DOM.

### `viewer.model`

Parsed presentation model containing slide metadata, text nodes, image nodes, shapes, and preview information.

### `loadPresentation(options)`

Lower-level parser that loads the same data model without mounting DOM immediately.

## Runtime notes

- Browser-only library. It depends on DOM APIs such as `DOMParser`, `Image`, `FileReader`, and `canvas`.
- SSR and pure Node.js execution are not supported.
- PDF preview rendering runs on the main thread.
- Without preview assets, the library falls back to semantic PPTX rendering of text, images, and basic shapes.

## Scripts

```bash
npm install
npm run dev
npm run build
npm run lint
npm test
```
