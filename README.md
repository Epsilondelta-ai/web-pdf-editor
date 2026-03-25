# web-ppt-editor

Browser PPTX viewer library.

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

if (!container) {
  throw new Error('Missing #viewer container');
}

const viewer = await createPptViewer(
  {
    pptx
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

`renderOptions`:
- `slidePixelWidth?: number` - target slide width for rendering.

### `viewer.mount(container)`

Renders slides into the provided HTML element.

### `viewer.destroy()`

Clears the mounted DOM.

### `viewer.model`

Parsed presentation model containing slide metadata, text nodes, image nodes, and shapes.

## Runtime notes

- Browser-only library. It depends on DOM APIs such as `DOMParser` and standard HTML rendering primitives.
- SSR and pure Node.js execution are not supported.
- The library renders text, images, and basic shapes directly from PPTX OOXML data.

## Scripts

```bash
npm install
npm run build
npm run typecheck
npm pack --dry-run
```
