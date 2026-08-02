// 规则集加载器（Node 侧）：读取目录下的 recognizers.json + styles.json 并校验。
// 校验失败时抛出聚合错误；调用方拿到的一定是合法规则集。

import fs from 'node:fs';
import path from 'node:path';

import { validateRuleset } from './schema.js';
import { COMPONENTS } from './components.js';

function readJson(filePath) {
  let raw;
  try {
    raw = fs.readFileSync(filePath, 'utf8');
  } catch {
    throw new Error(`规则集文件不存在或不可读：${path.basename(filePath)}（${filePath}）`);
  }
  try {
    return JSON.parse(raw);
  } catch (err) {
    throw new Error(`${path.basename(filePath)} 不是合法 JSON：${err.message}`);
  }
}

/**
 * 加载并校验一个规则集目录。
 * @param {string} dir 含 recognizers.json 与 styles.json 的目录
 * @returns {{ recognizers: object, styles: object, components: typeof COMPONENTS }}
 */
export function loadRuleset(dir) {
  const recognizers = readJson(path.join(dir, 'recognizers.json'));
  const styles = readJson(path.join(dir, 'styles.json'));
  const errors = validateRuleset(recognizers, styles);
  if (errors.length > 0) {
    throw new Error(`规则集校验失败（${dir}）：\n- ${errors.join('\n- ')}`);
  }
  return { recognizers, styles, components: COMPONENTS };
}
