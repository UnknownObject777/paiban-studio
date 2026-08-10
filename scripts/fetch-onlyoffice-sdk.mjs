// fetch-onlyoffice-sdk.mjs — 拉取 OnlyOffice 静态 SDK 资产到 public/packages/onlyoffice/。
//
// 背景：预览渲染层（D4）使用 onlyoffice-web-comp 的纯浏览器端 OnlyOffice 方案，
// SDK 静态资产体积约 600MB（含 x2t wasm），不入 git，由本脚本按需拉取。
//
// 来源：https://github.com/electroluxcode/onlyoffice-web-comp （AGPL-3.0，见仓库 LICENSE）
// 方式：git sparse-checkout（--depth 1 --filter=blob:none），只拉 public/packages/onlyoffice。
// 可重复执行：已存在时先清空再重新拉取，保证与上游快照一致。

import { execFileSync } from 'node:child_process';
import { cpSync, existsSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const REPO = 'https://github.com/electroluxcode/onlyoffice-web-comp';
const SPARSE_PATH = 'public/packages/onlyoffice';
const DEST = join(ROOT, 'public/packages/onlyoffice');

const git = (args, cwd) =>
  execFileSync('git', args, { cwd, stdio: ['ignore', 'inherit', 'inherit'] });

const staging = mkdtempSync(join(tmpdir(), 'paiban-oo-sdk-'));
try {
  console.log('[fetch:onlyoffice] sparse 克隆上游仓库…');
  git(['clone', '--depth', '1', '--filter=blob:none', '--sparse', REPO, 'repo'], staging);
  const repoDir = join(staging, 'repo');
  git(['sparse-checkout', 'set', SPARSE_PATH], repoDir);

  if (existsSync(DEST)) rmSync(DEST, { recursive: true, force: true });
  cpSync(join(repoDir, SPARSE_PATH), DEST, { recursive: true });

  const head = execFileSync('git', ['rev-parse', 'HEAD'], { cwd: repoDir, encoding: 'utf8' }).trim();
  console.log(`[fetch:onlyoffice] 完成：${DEST}（上游快照 ${head}）`);
} finally {
  rmSync(staging, { recursive: true, force: true });
}
