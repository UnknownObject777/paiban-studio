// 极简静态文件服务器：为规则集预览页提供本地服务（fetch 不支持 file:// 协议）。
// 用法：npm run preview  →  http://localhost:4173/preview/ruleset-gallery.html

import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

// 项目根：源码在 preview/（上级）；编译产物在 dist/preview/（上两级）。
const MODULE_DIR = path.dirname(fileURLToPath(import.meta.url));
const ROOT = [
  path.resolve(MODULE_DIR, '..'),
  path.resolve(MODULE_DIR, '../..'),
].find((p) => fs.existsSync(path.join(p, 'package.json')))!;
const PORT = Number(process.env.PORT ?? 4173);

const MIME = new Map<string, string>([
  ['.html', 'text/html; charset=utf-8'],
  ['.css', 'text/css; charset=utf-8'],
  ['.js', 'text/javascript; charset=utf-8'],
  ['.mjs', 'text/javascript; charset=utf-8'],
  ['.json', 'application/json; charset=utf-8'],
  ['.md', 'text/plain; charset=utf-8'],
]);

const server = http.createServer((req: http.IncomingMessage, res: http.ServerResponse) => {
  const urlPath = decodeURIComponent(new URL(req.url ?? '/', 'http://x').pathname);
  const filePath = path.normalize(path.join(ROOT, urlPath));
  if (!filePath.startsWith(ROOT)) {
    res.writeHead(403).end('forbidden');
    return;
  }
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404).end('not found');
      return;
    }
    res.writeHead(200, { 'Content-Type': MIME.get(path.extname(filePath)) ?? 'application/octet-stream' });
    res.end(data);
  });
});

server.listen(PORT, () => {
  console.log(`规则集预览：http://localhost:${PORT}/preview/ruleset-gallery.html`);
});
