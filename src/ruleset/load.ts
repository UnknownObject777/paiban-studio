// 规则集加载器（Node 侧）：读取目录下的 recognizers.json + styles.json 并校验。
// 校验失败时抛出聚合错误；调用方拿到的一定是合法规则集。

import fs from 'node:fs';
import path from 'node:path';

import { validateRuleset } from './schema.js';
import { COMPONENTS } from './components.js';
import type { ComponentDef } from './components.js';

function readJson(filePath: string): unknown {
  let raw: string;
  try {
    raw = fs.readFileSync(filePath, 'utf8');
  } catch {
    throw new Error(`规则集文件不存在或不可读：${path.basename(filePath)}（${filePath}）`);
  }
  try {
    return JSON.parse(raw);
  } catch (err) {
    throw new Error(`${path.basename(filePath)} 不是合法 JSON：${(err as Error).message}`);
  }
}

export interface LoadedRuleset {
  /** 规则集 JSON 数据（运行时对象，宽松类型便于组件键/样式访问）。 */
  recognizers: Record<string, any>;
  styles: Record<string, any>;
  components: readonly ComponentDef[];
}

/**
 * 加载并校验一个规则集目录。
 * @param dir 含 recognizers.json 与 styles.json 的目录
 * @returns 校验通过的两文件对象 + 组件清单常量
 */
export function loadRuleset(dir: string): LoadedRuleset {
  const recognizers = readJson(path.join(dir, 'recognizers.json'));
  const styles = readJson(path.join(dir, 'styles.json'));
  const errors = validateRuleset(recognizers, styles);
  if (errors.length > 0) {
    throw new Error(`规则集校验失败（${dir}）：\n- ${errors.join('\n- ')}`);
  }
  return { recognizers: recognizers as Record<string, any>, styles: styles as Record<string, any>, components: COMPONENTS };
}
