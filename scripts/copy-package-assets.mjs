import fs from 'node:fs/promises';
import path from 'node:path';

const rootDir = path.resolve(import.meta.dirname, '..');
const sourceCss = path.join(rootDir, 'src', 'styles.css');
const targetCss = path.join(rootDir, 'dist', 'styles.css');

await fs.mkdir(path.dirname(targetCss), { recursive: true });
await fs.copyFile(sourceCss, targetCss);
