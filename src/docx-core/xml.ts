// docx-core/xml.js — OOXML 部件的 XML 解析 / 序列化封装。
//
// round-trip 保真的核心：fast-xml-parser@5 (preserveOrder) 解析 → 序列化 → 重新解析。
// fxp 默认输出与 Word 原生 XML 存在几处无害的字节差异，本模块在 build 后做归一化还原：
//   1. UTF-8 BOM（Word 部件常带 BOM，fxp 解析时丢弃）
//   2. XML 声明后根元素前的 `\r\n` 分隔（Word 原生换行，fxp build 丢弃）
//   3. self-closing 空元素（`<w:jc/>` → fxp build 成 `<w:jc></w:jc>`）
//   4. 文本节点内 `"` 被过度转义为 `&quot;`（Word 文本节点不转义 `"`，仅属性值需要）
//
// 这些归一化保证：build(parse(xml)) === xml（逐字节稳定），从而编辑原语在 AST 上的
// 改动不会产生意外的字节漂移，Word/WPS 双端打开版式一致。

import { XMLParser, XMLBuilder } from 'fast-xml-parser';

// ---- fxp preserveOrder 树类型 ----

/** 属性对象（键形如 `@_w:val`；OOXML 属性值恒为字符串）。 */
export interface XmlAttrs {
  [attr: string]: string;
}

/**
 * fast-xml-parser preserveOrder 节点：
 *   元素节点: { "w:p": [ ...children ], ":@": { "@_w:val": "center" } }
 *   文本节点: { "#text": "..." }
 * 通过索引签名表达动态 tag（元素节点）与特殊键（:@ 属性 / #text 文本）。
 */
export interface XmlNode {
  [tag: string]: XmlNode[] | XmlAttrs | string | undefined;
  ':@'?: XmlAttrs;
  '#text'?: string;
}

/** 部件 XML 的 byte 级细节（BOM、XML 声明后分隔符），build 时按此还原。 */
export interface ParseMeta {
  bom: boolean;
  declSep: '' | '\r\n' | '\n';
}

export interface ParsedXml {
  tree: XmlNode[];
  meta: ParseMeta;
}

export const DECL_RE = /^(﻿)?<\?xml[^>]*\?>/;

export function createParser(): XMLParser {
  return new XMLParser({
    preserveOrder: true,
    ignoreAttributes: false,
    // 关键：保留文本节点前后空白。OOXML 大量使用 xml:space="preserve" 的文本
    // （如 w:instrText 字段指令、w:t 空 run），trimValues 默认 true 会破坏它们。
    trimValues: false,
    processEntities: true,
  });
}

export function createBuilder(): XMLBuilder {
  return new XMLBuilder({
    preserveOrder: true,
    ignoreAttributes: false,
    format: false,
  });
}

// 把 fxp 展开的空元素闭合标签还原为 self-closing（无空格：`<w:jc/>`）。
// 匹配形态：`<tag attrs></tag>`；attrs 内不出现裸 `>`，因此 `[^>]*` 安全。
function restoreSelfClosing(s: string): string {
  return s.replace(/<([\w:.-]+)((?:[^>]*)?)><\/\1>/g, (_, tag, attrs) => {
    return '<' + tag + attrs + '/>';
  });
}

// 文本节点内 `&quot;`/`&apos;` 还原为 `"`/`'`（XML 文本节点不需要转义引号，Word 也不转义）。
// 匹配 `>文本<` 区间（区间内无 `<`，故为纯文本节点；属性区位于 `<...>` 内不受影响），
// 把该区间内所有引号实体还原——覆盖同一文本节点内的多处引号（如 TOC 字段指令）。
// 仅作用于 build 输出（dirty 部件），实体转义还原后 XML 语义不变。
function restoreTextQuotes(s: string): string {
  return s.replace(/>([^<]*?)</g, (m, text) => {
    if (!text.includes('&quot;') && !text.includes('&apos;')) return m;
    return '>' + text.replace(/&quot;/g, '"').replace(/&apos;/g, "'") + '<';
  });
}

// 解析 OOXML 部件 XML 字符串 → fxp preserveOrder 结构数组。
// 返回 { tree, meta }，meta 记录 byte 级细节（BOM、声明后分隔符）供 build 还原。
export function parseXml(xml: string): ParsedXml {
  const meta: ParseMeta = { bom: xml.charCodeAt(0) === 0xfeff, declSep: '' };
  const body = meta.bom ? xml.slice(1) : xml;
  const m = DECL_RE.exec(body);
  if (m) {
    const after = body.slice(m[0].length);
    if (after.startsWith('\r\n')) meta.declSep = '\r\n';
    else if (after.startsWith('\n')) meta.declSep = '\n';
  }
  const parser = createParser();
  let tree = parser.parse(xml) as XmlNode[];
  // fxp 把 BOM 保留为根元素前的文本节点；剥离它，由 buildXml 依据 meta.bom 统一加回，
  // 避免"fxp 输出自带 BOM + 手动加回"产生双重 BOM。
  if (meta.bom && tree.length && tree[0] && tree[0]['#text']) {
    const head = tree[0]['#text'];
    if (head.charCodeAt(0) === 0xfeff || head === '﻿') {
      tree = tree.slice(1);
    }
  }
  return { tree, meta };
}

// 序列化 fxp 结构数组 → 还原 byte 级细节 → OOXML 部件 XML 字符串。
export function buildXml(tree: XmlNode[], meta: ParseMeta = { bom: false, declSep: '' }): string {
  let out = createBuilder().build(tree) as string;
  out = restoreTextQuotes(restoreSelfClosing(out));
  if (meta.declSep === '\r\n' || meta.declSep === '\n') {
    out = out.replace(DECL_RE, (m) => m + meta.declSep);
  }
  if (meta.bom) out = '﻿' + out;
  return out;
}

// 便捷：字符串 → 字符串 round-trip（供测试与无 meta 场景）。
export function roundTripXml(xml: string): string {
  const { tree, meta } = parseXml(xml);
  return buildXml(tree, meta);
}
