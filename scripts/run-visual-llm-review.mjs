import fs from 'node:fs/promises';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const rootDir = process.cwd();
const visualDir = path.resolve(rootDir, 'test-results/visual-review');
const artifactDir = path.resolve(rootDir, '.omx/artifacts');
const verdictStatePath = path.resolve(rootDir, '.omx/state/visual-review/ralph-progress.json');
const timestamp = new Date().toISOString().replace(/[:.]/g, '-');

function extractJsonBlock(raw) {
  const fenced = raw.match(/```json\s*([\s\S]*?)```/i);
  if (fenced) return fenced[1].trim();
  const plain = raw.match(/\{[\s\S]*\}/);
  if (plain) return plain[0].trim();
  throw new Error('No JSON block found in Claude output.');
}

function summarizeVerdict(verdict) {
  return `${verdict.verdict.toUpperCase()} (${verdict.score}) - ${verdict.reasoning}`;
}

function normalizeVerdictLabel(value) {
  const normalized = String(value ?? '').trim().toLowerCase();
  if (['pass', 'acceptable', 'accept', 'accepted', 'ok'].includes(normalized)) return 'pass';
  if (['fail', 'failed'].includes(normalized)) return 'fail';
  return 'revise';
}

async function runClaude(prompt) {
  const { stdout } = await execFileAsync('claude', ['-p', prompt], { cwd: rootDir, maxBuffer: 1024 * 1024 * 4 });
  return stdout.trim();
}

async function deckSummaries() {
  const entries = await fs.readdir(visualDir, { withFileTypes: true });
  const decks = [];
  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    const summaryPath = path.join(visualDir, entry.name, 'summary.json');
    try {
      const parsed = JSON.parse(await fs.readFile(summaryPath, 'utf8'));
      decks.push({ deck: entry.name, slides: parsed });
    } catch {}
  }
  return decks.sort((a, b) => a.deck.localeCompare(b.deck));
}

async function reviewSlide(deck, slide) {
  const comparisonPath = path.resolve(visualDir, deck, slide.comparison);
  const prompt = [
    'You are doing visual QA on a rendered PPT slide.',
    `Deck: ${deck}.`,
    `Analyze the image file at ${comparisonPath}.`,
    'The image contains exactly three side-by-side panels:',
    '1) actual render, 2) reference JPEG, 3) hotspot diff.',
    `Supporting metrics: avgChannelDelta=${slide.avgChannelDelta.toFixed(4)}, mismatchRatio=${slide.mismatchRatio.toFixed(4)}.`,
    'Focus on concrete visual mismatches in text visibility, alignment, spacing, cropping, and layout hierarchy.',
    'Use only one of these values for "verdict": "pass", "revise", or "fail".',
    'Return JSON only with this exact shape:',
    '{',
    '  "score": 0,',
    '  "verdict": "revise",',
    '  "category_match": false,',
    '  "differences": ["..."],',
    '  "suggestions": ["..."],',
    '  "reasoning": "short explanation"',
    '}'
  ].join('\n');

  const raw = await runClaude(prompt);
  const verdict = JSON.parse(extractJsonBlock(raw));
  verdict.verdict = normalizeVerdictLabel(verdict.verdict);

  const artifactPath = path.join(artifactDir, `claude-visual-review-${deck}-slide-${slide.slide}-${timestamp}.md`);
  await fs.writeFile(
    artifactPath,
    [
      '# Original user task',
      '사용자가 실제 렌더 결과와 sample 기준 이미지를 LLM으로 시각 비교해 달라고 요청했다.',
      '',
      '# Final prompt sent to Claude CLI',
      '```text',
      prompt,
      '```',
      '',
      '# Claude output (raw)',
      '```text',
      raw,
      '```',
      '',
      '# Concise summary',
      summarizeVerdict(verdict),
      '',
      '# Action items / next steps',
      ...(verdict.suggestions ?? []).map((item) => `- ${item}`)
    ].join('\n'),
    'utf8'
  );

  return {
    deck,
    slide: slide.slide,
    comparison: slide.comparison,
    metrics: {
      avgChannelDelta: slide.avgChannelDelta,
      mismatchRatio: slide.mismatchRatio
    },
    verdict
  };
}

async function main() {
  await fs.mkdir(artifactDir, { recursive: true });
  await fs.mkdir(path.dirname(verdictStatePath), { recursive: true });

  const summaries = await deckSummaries();
  if (!summaries.length) {
    throw new Error('No visual-review summaries found. Run test:visual:artifacts first.');
  }

  const results = [];
  for (const summary of summaries) {
    const slides = [];
    for (const slide of summary.slides) {
      slides.push(await reviewSlide(summary.deck, slide));
    }
    results.push({ deck: summary.deck, slides });
  }

  const allSlides = results.flatMap((deck) => deck.slides);
  const aggregate = {
    generatedAt: new Date().toISOString(),
    overall: {
      averageScore: allSlides.reduce((sum, item) => sum + Number(item.verdict.score || 0), 0) / Math.max(allSlides.length, 1),
      failingSlides: allSlides.filter((item) => item.verdict.verdict !== 'pass').map((item) => `${item.deck}:${item.slide}`)
    },
    decks: results
  };

  await fs.writeFile(path.join(visualDir, 'llm-verdicts.json'), JSON.stringify(aggregate, null, 2));
  await fs.writeFile(verdictStatePath, JSON.stringify(aggregate, null, 2));
  console.log(JSON.stringify(aggregate, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
