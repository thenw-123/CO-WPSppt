/**
 * Copy runtime assets into extension/bundled for vsce packaging.
 * Run from repo root: node scripts/copy-bundled.mjs
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, '..');
const extRoot = path.join(repoRoot, 'extension');
const dest = path.join(extRoot, 'bundled');

const copies = [
  'tools',
  'wps-driver',
  'specs',
  'themes',
  'tests',
  'docs',
  'manifest.json',
  'requirements-charts.txt',
  'skills',
];

function rmrf(p) {
  if (fs.existsSync(p)) fs.rmSync(p, { recursive: true, force: true });
}

rmrf(dest);
fs.mkdirSync(dest, { recursive: true });

for (const name of copies) {
  const src = path.join(repoRoot, name);
  const out = path.join(dest, name);
  if (!fs.existsSync(src)) {
    console.warn('skip missing:', src);
    continue;
  }
  const st = fs.statSync(src);
  if (st.isDirectory()) {
    fs.cpSync(src, out, { recursive: true });
  } else {
    fs.mkdirSync(path.dirname(out), { recursive: true });
    fs.copyFileSync(src, out);
  }
}

console.log('bundled ->', dest);
