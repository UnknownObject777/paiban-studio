// templates/rulesetFromSample.ts — 从样例 docx 反推规则集（spec 模块 4：规则集抽取）。
//
// 反推策略（保守、可解释）：
//   title     ：首个非空段落 → 段落/run 属性
//   body      ：文本最长段落 → 段落/run 属性
//   heading1-4 / caption / attachment：首个命中默认规则集 regex 识别规则的表外段落 → 段落/run 属性
//   table     ：首个表格数据行（≥2 行时取第二行首格，避开表头；否则首行）首个非空单元格 → 字体/字号/对齐
//   page      ：第一节 sectPr → 纸张 + 页边距（twips → cm）
//
// 样式值来源（issue #28）：真实模板的格式常定义在 styles.xml 命名样式里，段落只用
// <w:pStyle w:val="..."/> 引用、本身无 rPr/pPr 直接格式。因此先解析 word/styles.xml
// 建立 styleId → 段落/字符属性映射（含 basedOn 链，子样式覆盖父样式）；段落提取时按
// Word 层叠合并：命名样式属性垫底，段落直接格式（pPr / 首个含文本 run 的 rPr）覆盖之。
// 组件识别规则本身（哪个段落是 title/heading1/body/…）沿用默认规则集 recognizers，不动。
// 未实测到的组件继承内置公文默认规则集的对应节（本函数以默认集为底合并），
// 保证两文件组件键集完整、通过 schema 校验。

import { openDocx, getXmlTree } from '../docx-core/docx.js';
import type { Docx } from '../docx-core/docx.js';
import { walkParagraphs, walkSections } from '../docx-core/model.js';
import { findChild, findChildren, getAttr, childrenOf, tagOf } from '../docx-core/ooxml.js';
import { loadRuleset } from '../ruleset/load.js';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';
import { existsSync } from 'node:fs';
import type { XmlNode } from '../docx-core/xml.js';

// 默认规则集目录：相对本模块定位项目根 templates/。
// 源码位于 src/templates/（上两级）；编译产物位于 dist/src/templates/（上三级）。
const HERE = fileURLToPath(new URL('.', import.meta.url));
const DEFAULT_RULESET_DIR = [
  join(HERE, '../../templates/rulesets/gongwen-default'),
  join(HERE, '../../../templates/rulesets/gongwen-default'),
].find((p) => existsSync(p)) || join(HERE, '../../templates/rulesets/gongwen-default');

function plainText(pNode: XmlNode): string {
  let out = '';
  const visit = (n: XmlNode): void => {
    for (const c of childrenOf(n)) {
      if (c['#text'] !== undefined) out += c['#text'];
      else visit(c);
    }
  };
  visit(pNode);
  return out;
}

const num = (v: string | number | undefined): number | undefined =>
  v === undefined || v === null ? undefined : Number(v);

/** 反推出的组件样式（仅含从样例实测到的键）。 */
interface ExtractedStyle {
  align?: string;
  lineSpacingPt?: number;
  lineSpacingMultiple?: number;
  firstLineIndentChars?: number;
  fontEastAsia?: string;
  fontAscii?: string;
  sizePt?: number;
  bold?: boolean;
  [key: string]: unknown;
}

// 从 pPr 读出段落属性（对齐/行距/首行缩进），写入 out
function propsFromPPr(pPr: XmlNode | undefined, out: ExtractedStyle): void {
  if (!pPr) return;
  const jc = findChild(pPr, 'w:jc');
  if (jc) out.align = getAttr(jc, 'w:val') === 'both' ? 'justify' : getAttr(jc, 'w:val');
  const spacing = findChild(pPr, 'w:spacing');
  if (spacing && getAttr(spacing, 'w:line')) {
    const rule = getAttr(spacing, 'w:lineRule');
    const line = num(getAttr(spacing, 'w:line'));
    if (rule === 'exact' || rule === 'atLeast') out.lineSpacingPt = line! / 20;
    else if (line) out.lineSpacingMultiple = line / 240;
  }
  const ind = findChild(pPr, 'w:ind');
  if (ind && getAttr(ind, 'w:firstLineChars') !== undefined) {
    out.firstLineIndentChars = num(getAttr(ind, 'w:firstLineChars'))! / 100;
  }
}

// 从 rPr 读出字符属性（字体/字号/加粗），写入 out
function propsFromRPr(rPr: XmlNode | undefined, out: ExtractedStyle): void {
  if (!rPr) return;
  const fonts = findChild(rPr, 'w:rFonts');
  if (fonts) {
    if (getAttr(fonts, 'w:eastAsia')) out.fontEastAsia = getAttr(fonts, 'w:eastAsia');
    if (getAttr(fonts, 'w:ascii')) out.fontAscii = getAttr(fonts, 'w:ascii');
  }
  const sz = findChild(rPr, 'w:sz');
  if (sz) out.sizePt = num(getAttr(sz, 'w:val'))! / 2;
  const b = findChild(rPr, 'w:b');
  if (b) out.bold = getAttr(b, 'w:val') !== 'false';
}

/**
 * 解析 word/styles.xml → styleId → 段落/字符属性（仅 paragraph/character 样式）。
 * basedOn 链按 Word 层叠解析：子样式覆盖父样式（带环保护）。
 */
function parseNamedStyles(docx: Docx): Map<string, ExtractedStyle> {
  const defs = new Map<string, { props: ExtractedStyle; basedOn?: string }>();
  const tree = getXmlTree(docx, 'word/styles.xml');
  const stylesEl = tree?.find((n) => tagOf(n) === 'w:styles');
  if (stylesEl) {
    for (const st of findChildren(stylesEl, 'w:style')) {
      const type = getAttr(st, 'w:type');
      const styleId = getAttr(st, 'w:styleId');
      if (!styleId || (type !== 'paragraph' && type !== 'character')) continue;
      const props: ExtractedStyle = {};
      propsFromPPr(findChild(st, 'w:pPr'), props);
      propsFromRPr(findChild(st, 'w:rPr'), props);
      const basedOnEl = findChild(st, 'w:basedOn');
      defs.set(styleId, { props, basedOn: basedOnEl ? getAttr(basedOnEl, 'w:val') : undefined });
    }
  }
  const resolved = new Map<string, ExtractedStyle>();
  const resolve = (id: string, seen: Set<string>): ExtractedStyle => {
    const hit = resolved.get(id);
    if (hit) return hit;
    const def = defs.get(id);
    if (!def) return {};
    if (seen.has(id)) return def.props; // basedOn 成环：断在自身属性
    seen.add(id);
    const merged: ExtractedStyle = { ...(def.basedOn ? resolve(def.basedOn, seen) : {}), ...def.props };
    resolved.set(id, merged);
    return merged;
  };
  for (const id of defs.keys()) resolve(id, new Set());
  return resolved;
}

// 从段落读出组件样式：pStyle 命名样式属性垫底，直接格式（pPr + 首个含文本 run 的 rPr）覆盖之
function styleFromParagraph(pNode: XmlNode, namedStyles: Map<string, ExtractedStyle>): ExtractedStyle {
  const direct: ExtractedStyle = {};
  const pPr = findChild(pNode, 'w:pPr');
  propsFromPPr(pPr, direct);
  for (const r of findChildren(pNode, 'w:r')) {
    const t = findChild(r, 'w:t');
    if (!t || !childrenOf(t).length) continue;
    propsFromRPr(findChild(r, 'w:rPr'), direct);
    break; // 只取首个含文本 run
  }
  const pStyle = pPr ? findChild(pPr, 'w:pStyle') : undefined;
  const styleId = pStyle ? getAttr(pStyle, 'w:val') : undefined;
  const inherited = styleId ? namedStyles.get(styleId) : undefined;
  return inherited ? { ...inherited, ...direct } : direct;
}

export interface ExtractOptions {
  name?: string;
  defaultRulesetDir?: string;
}

export interface ExtractResult {
  recognizers: Record<string, unknown>;
  styles: Record<string, any>;
  extracted: string[];
}

/**
 * 从样例 docx 反推两文件规则集（title/body/heading/caption/table/page 实测 + 其余组件继承默认集）。
 * @returns 反推出的规则集两文件对象 + 实测到的组件清单
 */
export function extractRulesetFromSample(docxBuffer: Buffer | ArrayBuffer | Uint8Array, opts: ExtractOptions = {}): ExtractResult {
  const name = opts.name || 'extracted';
  const defaults = loadRuleset(opts.defaultRulesetDir || DEFAULT_RULESET_DIR);
  const docx = openDocx(docxBuffer);
  const namedStyles = parseNamedStyles(docx);

  // 收集段落快照（path 用于区分表内/表外段落）
  const paras: Array<{ node: XmlNode; text: string; path: string }> = [];
  walkParagraphs(docx, (p, path) => paras.push({ node: p, text: plainText(p), path }));
  const nonEmpty = paras.filter((x) => x.text.trim().length > 0);
  const extracted: string[] = [];

  const recognizers: Record<string, unknown> = JSON.parse(JSON.stringify(defaults.recognizers));
  const styles: Record<string, any> = JSON.parse(JSON.stringify(defaults.styles));
  recognizers.ruleset = name;
  styles.ruleset = name;

  if (nonEmpty.length) {
    const titleStyle = styleFromParagraph(nonEmpty[0].node, namedStyles);
    if (Object.keys(titleStyle).length) {
      styles.components.title = { ...styles.components.title, ...titleStyle, notes: '从样例首个非空段落反推' };
      extracted.push('title');
    }
    const longest = nonEmpty.reduce((a, b) => (b.text.length > a.text.length ? b : a));
    const bodyStyle = styleFromParagraph(longest.node, namedStyles);
    if (Object.keys(bodyStyle).length) {
      styles.components.body = { ...styles.components.body, ...bodyStyle, notes: '从样例最长段落反推' };
      extracted.push('body');
    }
  }

  // heading1-4 / caption / attachment：用默认规则集的 regex 识别规则在样例中取样
  // （识别规则本身不动，只补样式值来源；表内段落不参与，防止单元格文本误中标题正则）
  const bodyParas = nonEmpty.filter((x) => !x.path.includes('/tbl['));
  for (const id of ['heading1', 'heading2', 'heading3', 'heading4', 'caption', 'attachment']) {
    const entries = (defaults.recognizers.components?.[id]?.match || []) as Array<{ type?: string; pattern?: string }>;
    const entry = entries.find((e) => e.type === 'regex' && e.pattern);
    if (!entry) continue;
    const re = new RegExp(entry.pattern!);
    const hitPara = bodyParas.find((x) => re.test(x.text.trim()));
    if (!hitPara) continue;
    const st = styleFromParagraph(hitPara.node, namedStyles);
    if (Object.keys(st).length) {
      styles.components[id] = {
        ...styles.components[id],
        ...st,
        notes: `从样例「${hitPara.text.trim().slice(0, 20)}」反推`,
      };
      extracted.push(id);
    }
  }

  // table：首个表格数据行单元格取样（表内段落路径形如 /body/tbl[1]/tr[2]/tc[1]/p[1]）
  const cellOf = (path: string): { tbl: number; tr: number } | null => {
    const m = /\/tbl\[(\d+)\]\/tr\[(\d+)\]/.exec(path);
    return m ? { tbl: Number(m[1]), tr: Number(m[2]) } : null;
  };
  const cells = paras.filter((x) => cellOf(x.path) !== null);
  if (cells.length) {
    const firstTbl = cellOf(cells[0].path)!.tbl;
    const inFirst = cells.filter((x) => cellOf(x.path)!.tbl === firstTbl);
    const rows = [...new Set(inFirst.map((x) => cellOf(x.path)!.tr))].sort((a, b) => a - b);
    const dataRow = rows.length >= 2 ? rows[1] : rows[0];
    const sample =
      inFirst.find((x) => cellOf(x.path)!.tr === dataRow && x.text.trim().length > 0) ||
      inFirst.find((x) => cellOf(x.path)!.tr === dataRow);
    if (sample) {
      const full = styleFromParagraph(sample.node, namedStyles);
      // 表格组件样式键受 schema 限制：只取字体/字号/对齐（表头加粗等维度由 headerBold/smartAlign 承担）
      const tblStyle: ExtractedStyle = {};
      if (full.fontEastAsia !== undefined) tblStyle.fontEastAsia = full.fontEastAsia;
      if (full.fontAscii !== undefined) tblStyle.fontAscii = full.fontAscii;
      if (full.sizePt !== undefined) tblStyle.sizePt = full.sizePt;
      if (full.align !== undefined) tblStyle.align = full.align;
      if (Object.keys(tblStyle).length) {
        styles.components.table = { ...styles.components.table, ...tblStyle, notes: '从样例首个表格数据行单元格反推' };
        extracted.push('table');
      }
    }
  }

  walkSections(docx, (sect) => {
    const pgSz = findChild(sect, 'w:pgSz');
    const pgMar = findChild(sect, 'w:pgMar');
    const toCm = (tw: number | undefined): number => Math.round((tw! / 566.929) * 100) / 100;
    const page: Record<string, any> = { ...styles.page };
    if (pgSz && num(getAttr(pgSz, 'w:w')) === 11906) page.paper = 'A4';
    else page.paper = 'preserve';
    if (pgMar) {
      page.margins = {
        topCm: toCm(num(getAttr(pgMar, 'w:top'))),
        bottomCm: toCm(num(getAttr(pgMar, 'w:bottom'))),
        leftCm: toCm(num(getAttr(pgMar, 'w:left'))),
        rightCm: toCm(num(getAttr(pgMar, 'w:right'))),
      };
    }
    if (pgMar && getAttr(pgMar, 'w:footer')) page.footerDistanceCm = toCm(num(getAttr(pgMar, 'w:footer')));
    page.notes = '从样例 sectPr 反推';
    styles.page = page;
    extracted.push('page');
    return true;
  });

  return { recognizers, styles, extracted };
}
