// docx-core/model.ts — 文档模型访问层：路径寻址 + 结构遍历。
//
// 路径语法（借鉴 OfficeCLI 路径寻址，D6）：
//   /body/p[1]            document.xml 中 w:body 下第 1 个 w:p（1 起，按同 tag 兄弟计数）
//   /body/p[1]/r[2]       该段第 2 个 run
//   /body/tbl[1]/tr[2]/tc[1]/p[1]
//   /body/p[last]         最后一个段落
//   /body/sectPr          节属性
// 段落标签可写 p / w:p（自动补 w: 前缀）。

import { tagOf, isElement, childrenOf, findChild, textOf } from './ooxml.js';
import { getXmlTree } from './docx.js';
import type { XmlNode } from './xml.js';
import type { Docx } from './docx.js';

const W = (t: string): string => (t.includes(':') ? t : 'w:' + t);

export class PathError extends Error {
  code: string;
  suggestion?: string;
  constructor(message: string, suggestion?: string) {
    super(message);
    this.name = 'PathError';
    this.code = 'PATH_NOT_FOUND';
    this.suggestion = suggestion;
  }
}

/** 路径段：tag 为规范后的 w:xxx；index 为 1 起序号或 'last'。 */
export interface PathSegment {
  tag: string;
  index: number | 'last';
}

export interface ResolvedPath {
  node: XmlNode;
  parent: XmlNode | null;
  segments: PathSegment[];
}

// 解析路径字符串 → [{ tag, index }]（index 1 起；省略等同 [1]；'last' 表示末尾）。
export function parsePath(path: string): PathSegment[] {
  if (typeof path !== 'string' || !path.startsWith('/')) {
    throw new PathError(`路径必须以 / 开头: ${JSON.stringify(path)}`, '示例: /body/p[1]/r[2]');
  }
  const segs: PathSegment[] = [];
  for (const raw of path.split('/').filter(Boolean)) {
    const m = /^([\w:.-]+?)(?:\[(\d+|last)\])?$/.exec(raw);
    if (!m) throw new PathError(`无法解析路径段: ${raw}`, '形如 p[1] / r[2] / sectPr');
    segs.push({ tag: W(m[1]), index: m[2] === undefined ? 1 : m[2] === 'last' ? 'last' : Number(m[2]) });
  }
  if (!segs.length) throw new PathError('空路径', '示例: /body/p[1]');
  return segs;
}

// 在 children 数组中取第 index 个 tag 元素（1 起；'last' 取末尾）。
function nthOfTag(children: XmlNode[], tag: string, index: number | 'last'): XmlNode | undefined {
  const matches = children.filter((c) => isElement(c, tag));
  if (index === 'last') return matches[matches.length - 1];
  return matches[index - 1];
}

/**
 * 解析路径 → { node, parent, segments }
 * partName 默认在 document.xml 下寻址。
 */
export function resolvePath(docx: Docx, path: string, partName = 'word/document.xml'): ResolvedPath {
  const tree = getXmlTree(docx, partName);
  if (!tree) throw new PathError(`部件不存在: ${partName}`);
  const segments = parsePath(path);
  // 路径从文档根元素（w:document 等）之下开始寻址；若首段即根元素本身则从树层开始
  let children = tree;
  const rootEl = tree.find((n) => {
    const t = tagOf(n);
    return t !== undefined && t !== '?xml'; // 跳过 XML 声明处理指令
  });
  if (rootEl && tagOf(rootEl) !== segments[0].tag) {
    children = childrenOf(rootEl);
  }
  let node: XmlNode | null = null;
  let parent: XmlNode | null = null;
  for (let i = 0; i < segments.length; i++) {
    const { tag, index } = segments[i];
    const found = nthOfTag(children, tag, index);
    if (!found) {
      const available = [...new Set(children.map(tagOf).filter((x): x is string => Boolean(x)))].join(', ');
      throw new PathError(
        `路径 ${path} 在第 ${i + 1} 段未找到 ${tag}[${index}]`,
        `该层可用元素: ${available || '(无)'}`,
      );
    }
    parent = node;
    node = found;
    children = childrenOf(node);
  }
  return { node: node!, parent, segments };
}

// ---- 结构遍历 ----

export function getBodyNode(docx: Docx): XmlNode {
  return resolvePath(docx, '/body').node;
}

// 深度遍历 body，枚举所有段落（含表格单元格内段落），回调 (pNode, path)。
export function walkParagraphs(docx: Docx, fn: (p: XmlNode, path: string) => void): void {
  const body = getBodyNode(docx);
  const walk = (node: XmlNode, path: string): void => {
    const counters = new Map<string, number>(); // tag -> 序号
    for (const child of childrenOf(node)) {
      const tag = tagOf(child);
      if (!tag) continue;
      const n = (counters.get(tag) || 0) + 1;
      counters.set(tag, n);
      const childPath = `${path}/${tag.replace(/^w:/, '')}[${n}]`;
      if (tag === 'w:p') fn(child, childPath);
      else if (tag === 'w:tbl' || tag === 'w:tr' || tag === 'w:tc' || tag === 'w:sdt') walk(child, childPath);
    }
  };
  walk(body, '/body');
}

// 枚举所有 sectPr（body 末尾的 + 段落内嵌的分节符），回调 (sectPr, path)。
export function walkSections(docx: Docx, fn: (sectPr: XmlNode, path: string) => void): number {
  const body = getBodyNode(docx);
  const direct = findChild(body, 'w:sectPr');
  let paraCount = 0;
  walkParagraphs(docx, (p, path) => {
    paraCount++;
    const sect = findChild(findChild(p, 'w:pPr') || { 'w:pPr': [] }, 'w:sectPr');
    if (sect) fn(sect, `${path}/pPr/sectPr`);
  });
  if (direct) fn(direct, '/body/sectPr');
  return paraCount;
}

// 段落拼接纯文本（所有 w:t）。
export function paragraphText(pNode: XmlNode): string {
  return textOf(pNode);
}
