# web-ppt-editor

A minimal ONLYOFFICE-inspired browser PPT editor prototype.

## What it does

- Loads `.pptx` files directly in the browser.
- Parses slide text boxes, images, and basic shape frames from OOXML.
- Uses an optional exported PDF as a high-fidelity preview layer.
- Lets you edit text and drag editable overlays on top of slides.
- Re-exports the edited PPTX with updated text and element positions.
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
