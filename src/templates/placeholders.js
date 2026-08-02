// templates/placeholders.js — 模板 {{占位符}} 提取与合并（spec 模块 4）。
//
// 占位符语法：{{名称}}（中英文/数字/下划线/短横线）。
// 提取：扫描全部段落拼接文本，汇总占位符清单（名称 + 出现次数 + 首现路径）。
// 合并：生成编辑内核 findReplace 命令序列（{{名称}} → 值），跨 run 由内核处理。

import { walkParagraphs } from '../docx-core/model.js';
import { openDocx } from '../docx-core/docx.js';
import { tagOf } from '../docx-core/ooxml.js';

export const PLACEHOLDER_RE = /\{\{([\w一-龥][\w一-龥-]{0,30})\}\}/g;

function plainText(pNode) {
  let out = '';
  const visit = (n) => {
    for (const c of (n[tagOf(n)] || [])) {
      if (c['#text'] !== undefined) out += c['#text'];
      else visit(c);
    }
  };
  visit(pNode);
  return out;
}

/** 提取模板中的占位符清单。返回 [{ name, count, firstPath }]。 */
export function extractPlaceholders(docxBuffer) {
  const docx = openDocx(docxBuffer);
  const found = new Map();
  walkParagraphs(docx, (p, path) => {
    const text = plainText(p);
    for (const m of text.matchAll(PLACEHOLDER_RE)) {
      const name = m[1];
      if (!found.has(name)) found.set(name, { name, count: 0, firstPath: path });
      found.get(name).count++;
    }
  });
  return [...found.values()];
}

/** 占位符值表 → findReplace 命令序列（供 applyEdits 执行）。 */
export function placeholderCommands(values) {
  return Object.entries(values)
    .filter(([name]) => name.trim().length > 0)
    .map(([name, value]) => ({
      command: 'findReplace',
      find: `{{${name}}}`,
      replace: String(value ?? ''),
    }));
}
