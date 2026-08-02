// docx-core/outline.ts — 文档结构 dump（agent "dump → batch" 往返的 dump 端，D6）。
//
// 输出紧凑的段落清单（路径 / 样式 / 大纲级 / 文本预览）与节信息，
// 供 agent 系统提示词注入或 doc_outline 工具返回，让 LLM 能据路径精确寻址。

import { openDocx } from './docx.js';
import { walkParagraphs, walkSections } from './model.js';
import { findChild, getAttr, childrenOf, isElement, tagOf } from './ooxml.js';
import type { XmlNode } from './xml.js';

function plainText(pNode: XmlNode): string {
  let out = '';
  const visit = (n: XmlNode): void => {
    for (const c of (n[tagOf(n) as string] || []) as XmlNode[]) {
      if (c['#text'] !== undefined) out += c['#text'];
      else visit(c);
    }
  };
  visit(pNode);
  return out;
}

export interface OutlineEntry {
  path: string;
  text: string;
  length: number;
  style?: string;
  outlineLevel?: number;
  numbering?: boolean;
}

export interface OutlineSection {
  path: string;
  pageSize: { w: number; h: number; orient: string } | null;
  margins: Record<string, number> | null;
}

export interface OutlineResult {
  paragraphs: OutlineEntry[];
  sections: OutlineSection[];
  paragraphCount: number;
}

export interface OutlineOptions {
  textPreview?: number;
  includeEmpty?: boolean;
}

/**
 * 文档结构大纲。
 * @param docxBuffer 原始 docx 字节
 * @param opts 选项（textPreview 预览长度 / includeEmpty 是否含空段）
 */
export function dumpOutline(docxBuffer: Buffer | ArrayBuffer | Uint8Array, opts: OutlineOptions = {}): OutlineResult {
  const preview = opts.textPreview ?? 60;
  const docx = openDocx(docxBuffer);
  const paragraphs: OutlineEntry[] = [];
  walkParagraphs(docx, (p, path) => {
    const text = plainText(p);
    if (!text && opts.includeEmpty === false) return;
    const pPr = findChild(p, 'w:pPr');
    const style = pPr ? getAttr(findChild(pPr, 'w:pStyle'), 'w:val') : undefined;
    const outline = pPr ? getAttr(findChild(pPr, 'w:outlineLvl'), 'w:val') : undefined;
    const numPr = pPr ? findChild(pPr, 'w:numPr') : undefined;
    const entry: OutlineEntry = {
      path,
      text: text.length > preview ? text.slice(0, preview) + '…' : text,
      length: text.length,
    };
    if (style) entry.style = style;
    if (outline !== undefined) entry.outlineLevel = Number(outline);
    if (numPr) entry.numbering = true;
    paragraphs.push(entry);
  });
  const sections: OutlineSection[] = [];
  walkSections(docx, (sect, path) => {
    const pgSz = findChild(sect, 'w:pgSz');
    const pgMar = findChild(sect, 'w:pgMar');
    sections.push({
      path,
      pageSize: pgSz ? { w: Number(getAttr(pgSz, 'w:w')), h: Number(getAttr(pgSz, 'w:h')), orient: getAttr(pgSz, 'w:orient') || 'portrait' } : null,
      margins: pgMar ? Object.fromEntries(Object.entries(pgMar[':@'] || {}).map(([k, v]) => [k.replace(/^@_w:/, ''), Number(v)])) : null,
    });
  });
  return { paragraphs, sections, paragraphCount: paragraphs.length };
}
