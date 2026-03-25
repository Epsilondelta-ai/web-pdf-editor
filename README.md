# web-ppt-editor

A minimal browser PPT viewer prototype.

## What it does

- Loads `.pptx` files directly in the browser.
- Parses slide text boxes, images, and basic shape frames from OOXML.
- Uses an optional exported PDF as a high-fidelity preview layer.
- Renders semantic slide content directly when no preview asset is available.
- Provides Playwright e2e coverage that compares rendered slides against exported JPEG references when `sample/` exists locally.

## Scripts

```bash
npm install
npm run dev
npm run build
npm run lint
npm test
```

## Sample assets

Keep `sample/` local only.

Expected local files for the e2e suite:

- `sample/sample.pptx`
- `sample/sample.pdf`
- `sample/sample.001.jpeg`
- `sample/sample.002.jpeg`
- `sample/sample.003.jpeg`
- `sample/sample.004.jpeg`

The repository ignores `sample/`, so tests automatically skip when those fixtures are absent.

## Visual review workflow

For visual artifact generation plus LLM review against the local sample exports:

```bash
npm run test:visual
```

This creates:

- `test-results/visual-review/slide-*.actual.png`
- `test-results/visual-review/slide-*.comparison.png`
- `test-results/visual-review/summary.json`
- `test-results/visual-review/llm-verdicts.json`
- `.omx/artifacts/claude-visual-review-slide-*.md`

Notes:
- `test:visual:artifacts` captures the real rendered slides and builds side-by-side comparison images.
- `test:visual:llm` asks the local Claude CLI to review those comparison images.
- This requires a working local Claude CLI auth/session.
