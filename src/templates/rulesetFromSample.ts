// templates/rulesetFromSample.ts — 从样例 docx 反推规则集（spec 模块 4：规则集抽取，
// MVP 先做"标题 / 正文 / 页边距"最常用组件，完整组件抽取随 #5 原型迭代）。
//
// 反推策略（保守、可解释）：
//   title：首个非空段落 → 取其 run 字体/字号/粗体与段落对齐
//   body ：文本最长段落 → 取其 run 字体/字号与段落行距/首行缩进/对齐
//   page ：第一节 sectPr → 纸张 + 页边距（twips → cm）
// 其余组件（heading1..4/caption/table/attachment）继承内置公文默认规则集的对应节
// （由 templateStore 合并），保证两文件组件键集完整、通过 schema 校验。

import { openDocx } from '../docx-core/docx.js';
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

// 从段落读出组件样式（取首个含文本 run 的 rPr + pPr）
function styleFromParagraph(pNode: XmlNode): ExtractedStyle {
  const style: ExtractedStyle = {};
  const pPr = findChild(pNode, 'w:pPr');
  if (pPr) {
    const jc = findChild(pPr, 'w:jc');
    if (jc) style.align = getAttr(jc, 'w:val') === 'both' ? 'justify' : getAttr(jc, 'w:val');
    const spacing = findChild(pPr, 'w:spacing');
    if (spacing && getAttr(spacing, 'w:line')) {
      const rule = getAttr(spacing, 'w:lineRule');
      const line = num(getAttr(spacing, 'w:line'));
      if (rule === 'exact' || rule === 'atLeast') style.lineSpacingPt = line! / 20;
      else if (line) style.lineSpacingMultiple = line / 240;
    }
    const ind = findChild(pPr, 'w:ind');
    if (ind && getAttr(ind, 'w:firstLineChars') !== undefined) {
      style.firstLineIndentChars = num(getAttr(ind, 'w:firstLineChars'))! / 100;
    }
  }
  for (const r of findChildren(pNode, 'w:r')) {
    const t = findChild(r, 'w:t');
    if (!t || !childrenOf(t).length) continue;
    const rPr = findChild(r, 'w:rPr');
    if (!rPr) break;
    const fonts = findChild(rPr, 'w:rFonts');
    if (fonts) {
      if (getAttr(fonts, 'w:eastAsia')) style.fontEastAsia = getAttr(fonts, 'w:eastAsia');
      if (getAttr(fonts, 'w:ascii')) style.fontAscii = getAttr(fonts, 'w:ascii');
    }
    const sz = findChild(rPr, 'w:sz');
    if (sz) style.sizePt = num(getAttr(sz, 'w:val'))! / 2;
    const b = findChild(rPr, 'w:b');
    if (b) style.bold = getAttr(b, 'w:val') !== 'false';
    break; // 只取首个含文本 run
  }
  return style;
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
 * 从样例 docx 反推两文件规则集（title/body/page 实测 + 其余组件继承默认集）。
 * @returns 反推出的规则集两文件对象 + 实测到的组件清单
 */
export function extractRulesetFromSample(docxBuffer: Buffer | ArrayBuffer | Uint8Array, opts: ExtractOptions = {}): ExtractResult {
  const name = opts.name || 'extracted';
  const defaults = loadRuleset(opts.defaultRulesetDir || DEFAULT_RULESET_DIR);
  const docx = openDocx(docxBuffer);

  // 收集段落快照
  const paras: Array<{ node: XmlNode; text: string }> = [];
  walkParagraphs(docx, (p) => paras.push({ node: p, text: plainText(p) }));
  const nonEmpty = paras.filter((x) => x.text.trim().length > 0);
  const extracted: string[] = [];

  const recognizers: Record<string, unknown> = JSON.parse(JSON.stringify(defaults.recognizers));
  const styles: Record<string, any> = JSON.parse(JSON.stringify(defaults.styles));
  recognizers.ruleset = name;
  styles.ruleset = name;

  if (nonEmpty.length) {
    const titleStyle = styleFromParagraph(nonEmpty[0].node);
    if (Object.keys(titleStyle).length) {
      styles.components.title = { ...styles.components.title, ...titleStyle, notes: '从样例首个非空段落反推' };
      extracted.push('title');
    }
    const longest = nonEmpty.reduce((a, b) => (b.text.length > a.text.length ? b : a));
    const bodyStyle = styleFromParagraph(longest.node);
    if (Object.keys(bodyStyle).length) {
      styles.components.body = { ...styles.components.body, ...bodyStyle, notes: '从样例最长段落反推' };
      extracted.push('body');
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
