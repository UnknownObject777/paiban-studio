// 一次性迁移脚本：把 src/onlyoffice-comp（vendored from onlyoffice-web-comp，Next.js bundler 风格）
// 的相对 import 改写为 NodeNext/浏览器 ESM 可加载的显式 .js 形式。
// 规则：相对路径无扩展名时，若目标是目录则补 "/index.js"，否则补 ".js"。
import { existsSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';

const TARGET = resolve('src/onlyoffice-comp');

function* walk(dir) {
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) yield* walk(p);
    else if (p.endsWith('.ts')) yield p;
  }
}

function fixSpecifier(file, spec) {
  if (!spec.startsWith('.')) return spec;
  if (/\.(js|mjs|json)$/.test(spec)) return spec;
  const abs = resolve(dirname(file), spec);
  if (existsSync(abs) && statSync(abs).isDirectory()) return spec + '/index.js';
  return spec + '.js';
}

let changed = 0;
for (const file of walk(TARGET)) {
  const src = readFileSync(file, 'utf8');
  const out = src
    .replace(/(from\s*["'])(\.[^"']+)(["'])/g, (_m, a, spec, b) => a + fixSpecifier(file, spec) + b)
    .replace(/(import\s*\(\s*["'])(\.[^"']+)(["']\s*\))/g, (_m, a, spec, b) => a + fixSpecifier(file, spec) + b)
    // bundler 风格的 worker URL：tsc 产物是同目录 .js
    .replace('new URL("./x2t.worker.ts", import.meta.url)', 'new URL("./x2t.worker.js", import.meta.url)');
  if (out !== src) {
    writeFileSync(file, out);
    changed++;
  }
}
console.log(`rewritten files: ${changed}`);
