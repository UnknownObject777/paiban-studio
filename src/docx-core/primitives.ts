// docx-core/primitives.ts — 编辑原语实现（MVP 子集，spec 模块 1）。
//
// 覆盖：
//   段落：对齐 / 首行缩进（字符或磅）/ 行距（磅值或倍数）/ 段前段后 / 分页控制 / 大纲级别 / 样式
//   run ：中文字体(eastAsia) / 西文字体 / 字号 / 粗斜下划线 / 颜色 / 突出显示
//   节  ：页边距 / 纸张大小方向 / 页码格式（pgNumType + 页脚 PAGE 字段）
//   文本：findReplace（跨 run 自动拆分）
//   结构：addParagraph / remove / move
//
// 所有原语直接操作 fxp 树；单位换算集中在顶部常量（公文用户习惯 cm / 磅 / 字符）。

import {
  el, textEl, tagOf, isElement, attrsOf, getAttr, setAttr, removeAttr,
  childrenOf, findChild, findChildren, ensureChild, insertOrdered,
  removeChildren, ensurePropContainer, setPPrLeaf, setRPrLeaf,
  PPR_ORDER, RPR_ORDER, SECTPR_ORDER,
} from './ooxml.js';
import { getXmlTree, markDirty } from './docx.js';
import { walkSections, walkParagraphs, parsePath, resolvePath } from './model.js';
import { parseXml } from './xml.js';
import type { XmlNode } from './xml.js';
import type { Docx } from './docx.js';

export const CM_TO_TWIPS = 566.929; // 1cm = 566.929 twips
export const PT_TO_TWIPS = 20;

const twips = (cm: number): number => Math.round(cm * CM_TO_TWIPS);
const lineTwips = (pt: number): number => Math.round(pt * PT_TO_TWIPS);

// ---- 段落属性 ----

export interface ParagraphProps {
  align?: string;
  style?: string;
  firstLineChars?: number;
  firstLinePt?: number;
  indentLeftPt?: number;
  indentRightPt?: number;
  hangingChars?: number;
  lineSpacingPt?: number;
  lineSpacingMinPt?: number;
  lineSpacingMultiple?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  spacingBeforeLines?: number;
  spacingAfterLines?: number;
  pageBreakBefore?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  outlineLevel?: number;
  [key: string]: unknown;
}

const ALIGN_MAP: Record<string, string> = {
  left: 'left', center: 'center', right: 'right',
  justify: 'both', both: 'both', distribute: 'distribute',
  左: 'left', 居中: 'center', 右: 'right', 两端: 'both', 分散: 'distribute',
};

/**
 * 设置段落属性。props 支持：
 *   align                对齐（left/center/right/justify/both/distribute 或中文别名）
 *   style                段落样式 id（pStyle）
 *   firstLineChars       首行缩进字符数 ×100（公文标准：200 = 2 字符）
 *   firstLinePt          首行缩进磅值
 *   indentLeftPt/indentRightPt  左右缩进磅值
 *   lineSpacingPt        固定行距磅值（lineRule=exact）
 *   lineSpacingMinPt     最小行距磅值（lineRule=atLeast）
 *   lineSpacingMultiple  倍数行距（1.5 → line=360 auto）
 *   spacingBeforePt/spacingAfterPt  段前段后磅值
 *   pageBreakBefore      段前分页（bool）
 *   keepNext/keepLines   与下段同页 / 段内不分页（bool）
 *   outlineLevel         大纲级别 0-8（9 = 正文）
 */
export function setParagraphProps(pNode: XmlNode, props: ParagraphProps): string[] {
  const pPr = ensurePropContainer(pNode, 'w:pPr');
  const applied: string[] = [];

  if (props.align !== undefined) {
    const val = ALIGN_MAP[props.align];
    if (!val) throw new Error(`未知对齐方式: ${props.align}`);
    setPPrLeaf(pPr, 'w:jc', { 'w:val': val });
    applied.push(`align=${val}`);
  }
  if (props.style !== undefined) {
    setPPrLeaf(pPr, 'w:pStyle', { 'w:val': props.style });
    applied.push(`style=${props.style}`);
  }
  // 缩进（firstLineChars 与 firstLine 可同时写，Word 以 chars 优先）
  const ind: Record<string, string | number> = {};
  if (props.firstLineChars !== undefined) ind['w:firstLineChars'] = Math.round(props.firstLineChars);
  if (props.firstLinePt !== undefined) ind['w:firstLine'] = lineTwips(props.firstLinePt);
  if (props.indentLeftPt !== undefined) ind['w:left'] = lineTwips(props.indentLeftPt);
  if (props.indentRightPt !== undefined) ind['w:right'] = lineTwips(props.indentRightPt);
  if (props.hangingChars !== undefined) ind['w:hangingChars'] = Math.round(props.hangingChars);
  if (Object.keys(ind).length) {
    setPPrLeaf(pPr, 'w:ind', ind);
    applied.push(`ind=${JSON.stringify(ind)}`);
  }
  // 行距与段距
  const spacing: Record<string, string | number> = {};
  if (props.lineSpacingPt !== undefined) {
    spacing['w:line'] = lineTwips(props.lineSpacingPt);
    spacing['w:lineRule'] = 'exact';
  } else if (props.lineSpacingMinPt !== undefined) {
    spacing['w:line'] = lineTwips(props.lineSpacingMinPt);
    spacing['w:lineRule'] = 'atLeast';
  } else if (props.lineSpacingMultiple !== undefined) {
    spacing['w:line'] = Math.round(props.lineSpacingMultiple * 240);
    spacing['w:lineRule'] = 'auto';
  }
  if (props.spacingBeforePt !== undefined) spacing['w:before'] = lineTwips(props.spacingBeforePt);
  if (props.spacingAfterPt !== undefined) spacing['w:after'] = lineTwips(props.spacingAfterPt);
  if (props.spacingBeforeLines !== undefined) spacing['w:beforeLines'] = Math.round(props.spacingBeforeLines * 100);
  if (props.spacingAfterLines !== undefined) spacing['w:afterLines'] = Math.round(props.spacingAfterLines * 100);
  if (Object.keys(spacing).length) {
    setPPrLeaf(pPr, 'w:spacing', spacing);
    applied.push(`spacing=${JSON.stringify(spacing)}`);
  }
  // 分页控制
  for (const [key, tag] of [['pageBreakBefore', 'w:pageBreakBefore'], ['keepNext', 'w:keepNext'], ['keepLines', 'w:keepLines']] as const) {
    if (props[key] !== undefined) {
      if (props[key]) setPPrLeaf(pPr, tag, {});
      else removeChildren(pPr, tag);
      applied.push(`${key}=${!!props[key]}`);
    }
  }
  if (props.outlineLevel !== undefined) {
    setPPrLeaf(pPr, 'w:outlineLvl', { 'w:val': props.outlineLevel });
    applied.push(`outlineLvl=${props.outlineLevel}`);
  }
  return applied;
}

// ---- run 属性 ----

export interface RunProps {
  eastAsia?: string;
  ascii?: string;
  hAnsi?: string;
  cs?: string;
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: string | boolean;
  color?: string;
  highlight?: string;
  [key: string]: unknown;
}

const UNDERLINE_MAP: Record<string, string> = {
  true: 'single', single: 'single', double: 'double', dotted: 'dotted',
  dashed: 'dash', wave: 'wave', none: 'none', false: 'none',
};

/**
 * 设置 run 属性。props 支持：
 *   eastAsia   中文字体（w:rFonts w:eastAsia，如 仿宋_GB2312 / 黑体）
 *   ascii/hAnsi  西文字体
 *   sizePt     字号（磅，自动 ×2 写半磅；三号=16 四号=14 小四=12 五号=10.5）
 *   bold/italic  粗体/斜体（bool）
 *   underline  下划线（true=single 或 single/double/dotted/dashed/wave/none）
 *   color      颜色 hex（不带 #）
 *   highlight  突出显示色（yellow 等）
 */
export function setRunProps(rNode: XmlNode, props: RunProps): string[] {
  const rPr = ensurePropContainer(rNode, 'w:rPr');
  const applied: string[] = [];

  const fonts: Record<string, string> = {};
  if (props.eastAsia !== undefined) fonts['w:eastAsia'] = props.eastAsia;
  if (props.ascii !== undefined) fonts['w:ascii'] = props.ascii;
  if (props.hAnsi !== undefined) fonts['w:hAnsi'] = props.hAnsi;
  if (props.cs !== undefined) fonts['w:cs'] = props.cs;
  if (Object.keys(fonts).length) {
    setRPrLeaf(rPr, 'w:rFonts', fonts);
    applied.push(`rFonts=${JSON.stringify(fonts)}`);
  }
  if (props.sizePt !== undefined) {
    const half = Math.round(props.sizePt * 2);
    setRPrLeaf(rPr, 'w:sz', { 'w:val': half });
    setRPrLeaf(rPr, 'w:szCs', { 'w:val': half });
    applied.push(`sz=${half}`);
  }
  for (const [key, tag] of [['bold', 'w:b'], ['italic', 'w:i']] as const) {
    if (props[key] !== undefined) {
      const leaf = setRPrLeaf(rPr, tag, {});
      // OOXML 开关：显式写 val 以覆盖样式继承
      setAttr(leaf, 'w:val', props[key] ? 'true' : 'false');
      const csTag = tag + 'Cs';
      const cs = setRPrLeaf(rPr, csTag, {});
      setAttr(cs, 'w:val', props[key] ? 'true' : 'false');
      applied.push(`${key}=${!!props[key]}`);
    }
  }
  if (props.underline !== undefined) {
    const val = UNDERLINE_MAP[String(props.underline)];
    if (!val) throw new Error(`未知下划线: ${props.underline}`);
    if (val === 'none') removeChildren(rPr, 'w:u');
    else setRPrLeaf(rPr, 'w:u', { 'w:val': val });
    applied.push(`u=${val}`);
  }
  if (props.color !== undefined) {
    setRPrLeaf(rPr, 'w:color', { 'w:val': String(props.color).replace(/^#/, '').toUpperCase() });
    applied.push(`color=${props.color}`);
  }
  if (props.highlight !== undefined) {
    if (props.highlight === 'none') removeChildren(rPr, 'w:highlight');
    else setRPrLeaf(rPr, 'w:highlight', { 'w:val': props.highlight });
    applied.push(`highlight=${props.highlight}`);
  }
  return applied;
}

// 对段落内全部 run 应用 run 属性（空段落无 run 时可选创建空 run 承载格式）。
export function setParagraphRunProps(pNode: XmlNode, props: RunProps): { runCount: number; applied: string[] } {
  const runs = findChildren(pNode, 'w:r');
  const applied: string[] = [];
  for (const r of runs) applied.push(...setRunProps(r, props));
  return { runCount: runs.length, applied };
}

// ---- 节属性（页面设置） ----

const PAGE_SIZES: Record<string, { w: number; h: number }> = {
  a4: { w: 11906, h: 16838 }, a3: { w: 16838, h: 23811 },
  letter: { w: 12240, h: 15840 }, legal: { w: 12240, h: 20160 },
  '16k': { w: 11040, h: 15600 },
};

export interface SectionProps {
  pageSize?: string | { widthCm: number; heightCm: number };
  orientation?: string;
  marginsCm?: Record<string, number>;
  pageNumFmt?: string;
  pageNumStart?: number;
  [key: string]: unknown;
}

/**
 * 设置节属性（作用于指定 sectPr）。props 支持：
 *   marginsCm: { top, right, bottom, left, header, footer, gutter }  单位 cm
 *   pageSize: 'a4'|'a3'|'letter'|'legal'|'16k' 或 { widthCm, heightCm }
 *   orientation: 'portrait'|'landscape'
 *   pageNumFmt: 'decimal'|'chineseCounting'|'lowerRoman'|...
 *   pageNumStart: 起始页码
 */
export function setSectionProps(sectPr: XmlNode, props: SectionProps): string[] {
  const applied: string[] = [];

  if (props.pageSize !== undefined || props.orientation !== undefined) {
    let size: { w: number; h: number } | null = null;
    if (typeof props.pageSize === 'string') {
      size = PAGE_SIZES[props.pageSize.toLowerCase()];
      if (!size) throw new Error(`未知纸张: ${props.pageSize}`);
    } else if (props.pageSize && typeof props.pageSize === 'object') {
      size = { w: twips(props.pageSize.widthCm), h: twips(props.pageSize.heightCm) };
    }
    const pgSz = ensureChild(sectPr, 'w:pgSz', SECTPR_ORDER);
    if (size) {
      setAttr(pgSz, 'w:w', size.w);
      setAttr(pgSz, 'w:h', size.h);
    }
    if (props.orientation !== undefined) {
      if (props.orientation === 'landscape') {
        setAttr(pgSz, 'w:orient', 'landscape');
        // Word 语义：landscape 时 w > h；若当前是纵向尺寸则交换
        const w = Number(getAttr(pgSz, 'w:w')), h = Number(getAttr(pgSz, 'w:h'));
        if (w && h && w < h) { setAttr(pgSz, 'w:w', h); setAttr(pgSz, 'w:h', w); }
      } else {
        removeAttr(pgSz, 'w:orient');
        const w = Number(getAttr(pgSz, 'w:w')), h = Number(getAttr(pgSz, 'w:h'));
        if (w && h && w > h) { setAttr(pgSz, 'w:w', h); setAttr(pgSz, 'w:h', w); }
      }
    }
    applied.push('pgSz');
  }

  if (props.marginsCm !== undefined) {
    const pgMar = ensureChild(sectPr, 'w:pgMar', SECTPR_ORDER);
    for (const [k, v] of Object.entries(props.marginsCm)) {
      if (v === undefined) continue;
      setAttr(pgMar, 'w:' + k, twips(v));
    }
    applied.push(`pgMar=${JSON.stringify(props.marginsCm)}`);
  }

  if (props.pageNumFmt !== undefined || props.pageNumStart !== undefined) {
    const pgNum = ensureChild(sectPr, 'w:pgNumType', SECTPR_ORDER);
    if (props.pageNumFmt !== undefined) setAttr(pgNum, 'w:fmt', props.pageNumFmt);
    if (props.pageNumStart !== undefined) setAttr(pgNum, 'w:start', props.pageNumStart);
    applied.push('pgNumType');
  }
  return applied;
}

// 对文档所有节应用属性；sectionIndex 指定时只应用该节（0 起，按文档顺序）。
export function setAllSectionsProps(docx: Docx, props: SectionProps, sectionIndex?: number): { sections: number; applied: string[] } {
  const sects: XmlNode[] = [];
  walkSections(docx, (sect) => sects.push(sect));
  if (!sects.length) throw new Error('文档无 sectPr（非合法 Word 文档？）');
  const targets = sectionIndex === undefined ? sects : [sects[sectionIndex]].filter(Boolean);
  if (!targets.length) throw new Error(`节索引越界: ${sectionIndex}（共 ${sects.length} 节）`);
  const applied: string[] = [];
  for (const s of targets) applied.push(...setSectionProps(s, props));
  markDirty(docx, 'word/document.xml');
  return { sections: targets.length, applied };
}

// ---- 页脚页码字段 ----

// 在 document.xml.rels 中分配新 rId。
function nextRelId(relsTree: XmlNode[]): { root: XmlNode; id: string } {
  const root = relsTree.find((n) => isElement(n, 'Relationships'));
  if (!root) throw new Error('非法 rels 部件');
  let max = 0;
  for (const r of findChildren(root, 'Relationship')) {
    const m = /^rId(\d+)$/.exec(getAttr(r, 'Id') || '');
    if (m) max = Math.max(max, Number(m[1]));
  }
  return { root, id: 'rId' + (max + 1) };
}

// 确保 [Content_Types].xml 含 footer 的 Override。
function ensureContentTypeOverride(ctTree: XmlNode[], partName: string, contentType: string): void {
  const root = ctTree.find((n) => isElement(n, 'Types'));
  if (!root) throw new Error('非法 [Content_Types].xml');
  const exists = findChildren(root, 'Override').some((o) => getAttr(o, 'PartName') === partName);
  if (!exists) {
    childrenOf(root).push(el('Override', { PartName: partName, ContentType: contentType }));
  }
}

export interface FooterOptions {
  align?: string;
  sectionIndex?: number;
}

/**
 * 为文档（指定节，默认最后一节）插入居中页码页脚（PAGE 字段）。
 * 若节已有 footerReference 则复用既有 footer 部件并追加页码段；否则新建 footerN.xml。
 */
export function ensurePageNumberFooter(docx: Docx, { align = 'center', sectionIndex }: FooterOptions = {}): { footer: string; align: string } {
  const docPart = 'word/document.xml';
  // 1. 找目标 sectPr
  const sects: XmlNode[] = [];
  walkSections(docx, (sect) => sects.push(sect));
  if (!sects.length) throw new Error('文档无 sectPr');
  const sectPr = sectionIndex === undefined ? sects[sects.length - 1] : sects[sectionIndex];
  if (!sectPr) throw new Error(`节索引越界: ${sectionIndex}`);

  // 2. 已有 footerReference → 复用部件
  let footerName: string | null = null;
  const existing = findChildren(sectPr, 'w:footerReference').find((f) => getAttr(f, 'w:type') === 'default');
  const relsTree = getXmlTree(docx, 'word/_rels/document.xml.rels')!;
  if (existing) {
    const rId = getAttr(existing, 'r:id');
    const root = relsTree.find((n) => isElement(n, 'Relationships'));
    const rel = root ? findChildren(root, 'Relationship').find((r) => getAttr(r, 'Id') === rId) : undefined;
    if (rel) footerName = 'word/' + getAttr(rel, 'Target');
  }

  if (!footerName) {
    // 3. 新建 footer 部件
    let n = 1;
    while (docx.parts.has(`word/footer${n}.xml`)) n++;
    footerName = `word/footer${n}.xml`;
    const footerXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"></w:ftr>';
    const { tree, meta } = parseXml(footerXml);
    docx.parts.set(footerName, { kind: 'xml', tree, meta, text: footerXml, dirty: true });

    const { root, id } = nextRelId(relsTree);
    childrenOf(root).push(el('Relationship', {
      Id: id,
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
      Target: `footer${n}.xml`,
    }));
    markDirty(docx, 'word/_rels/document.xml.rels');

    const ctTree = getXmlTree(docx, '[Content_Types].xml')!;
    ensureContentTypeOverride(ctTree, '/' + footerName,
      'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');
    markDirty(docx, '[Content_Types].xml');

    const ref = el('w:footerReference', { 'w:type': 'default', 'r:id': id });
    insertOrdered(sectPr, ref, SECTPR_ORDER);
  }

  // 4. 向 footer 写入 PAGE 字段段落（先清掉旧页码段：含 PAGE 指令的段落）
  const fPart = docx.parts.get(footerName);
  if (!fPart || fPart.kind !== 'xml') throw new Error(`footer 部件非法: ${footerName}`);
  const fRoot = fPart.tree.find((node) => isElement(node, 'w:ftr'))!;
  if (!fRoot) throw new Error(`footer 部件非法: ${footerName}`);
  const jcVal = align === 'left' ? 'left' : align === 'right' ? 'right' : 'center';
  const p = el('w:p', {}, [
    el('w:pPr', {}, [el('w:jc', { 'w:val': jcVal })]),
    el('w:r', {}, [el('w:fldChar', { 'w:fldCharType': 'begin' })]),
    el('w:r', {}, [textEl('w:instrText', ' PAGE ', { 'xml:space': 'preserve' })]),
    el('w:r', {}, [el('w:fldChar', { 'w:fldCharType': 'separate' })]),
    el('w:r', {}, [textEl('w:t', '1')]),
    el('w:r', {}, [el('w:fldChar', { 'w:fldCharType': 'end' })]),
  ]);
  // 移除已有 PAGE 字段段，避免重复
  const hasPageField = (pNode: XmlNode): boolean => {
    let found = false;
    const visit = (n: XmlNode): void => {
      for (const c of childrenOf(n)) {
        if (isElement(c, 'w:instrText') && /PAGE/.test(childrenOf(c).map((x) => x['#text'] || '').join(''))) found = true;
        else if (tagOf(c)) visit(c);
      }
    };
    visit(pNode);
    return found;
  };
  fRoot['w:ftr'] = childrenOf(fRoot).filter((c) => !(isElement(c, 'w:p') && hasPageField(c)));
  childrenOf(fRoot).push(p);
  markDirty(docx, footerName);
  markDirty(docx, docPart);
  return { footer: footerName, align: jcVal };
}

// ---- findReplace（跨 run） ----

// 段落的文本节点清单：[{ tNode, text, start, end }]（start/end 为拼接文本中的偏移）
interface TextNodeRef {
  tNode: XmlNode;
  text: string;
  start: number;
  end: number;
}

function paragraphTextNodes(pNode: XmlNode): TextNodeRef[] {
  const nodes: TextNodeRef[] = [];
  let offset = 0;
  const visit = (node: XmlNode): void => {
    for (const c of childrenOf(node)) {
      if (isElement(c, 'w:t')) {
        const text = childrenOf(c).map((x) => x['#text'] || '').join('');
        nodes.push({ tNode: c, text, start: offset, end: offset + text.length });
        offset += text.length;
      } else if (isElement(c, 'w:tab')) {
        offset += 1; // tab 计 1 字符（与拼接文本约定一致）
      } else if (tagOf(c)) {
        visit(c);
      }
    }
  };
  visit(pNode);
  return nodes;
}

function setTextNodeText(tNode: XmlNode, text: string): void {
  tNode['w:t'] = text === '' ? [] : [{ '#text': text }];
  if (text.startsWith(' ') || text.endsWith(' ')) {
    setAttr(tNode, 'xml:space', 'preserve');
  }
}

// findReplace 的替换串暂存（replaceOnceInParagraph 需要，避免长参数链）。
let pendingReplacement = '';

// 在段落中执行一次替换（替换第一个匹配），返回是否命中。
// 跨 run 匹配：替换文本写入首个重叠文本节点，其余重叠节点截去重叠部分；
// 每次替换后由调用方重算节点偏移（段落级开销可忽略，逻辑零漂移）。
function replaceOnceInParagraph(pNode: XmlNode, find: string, caseSensitive: boolean): boolean {
  const nodes = paragraphTextNodes(pNode);
  const full = nodes.map((n) => n.text).join('');
  const haystack = caseSensitive ? full : full.toLowerCase();
  const needle = caseSensitive ? find : find.toLowerCase();
  const ms = haystack.indexOf(needle);
  if (ms === -1) return false;
  const me = ms + needle.length;
  let first = true;
  for (const n of nodes) {
    if (n.end <= ms || n.start >= me) continue; // 不重叠
    const localS = Math.max(0, ms - n.start);
    const localE = Math.min(n.text.length, me - n.start);
    if (first) {
      setTextNodeText(n.tNode, n.text.slice(0, localS) + pendingReplacement + n.text.slice(localE));
      first = false;
    } else {
      setTextNodeText(n.tNode, n.text.slice(0, localS) + n.text.slice(localE));
    }
  }
  return true;
}

export interface FindReplaceOptions {
  caseSensitive?: boolean;
  maxCount?: number;
}

/**
 * 全文查找替换。自动处理跨 run 匹配。options: { caseSensitive, maxCount }
 * 返回 { replaced, paragraphs }。
 */
export function findReplace(docx: Docx, find: string, replace: string, options: FindReplaceOptions = {}): { replaced: number; paragraphs: string[] } {
  if (!find) throw new Error('findReplace 需要非空 find');
  let replaced = 0;
  const touched = new Set<string>();
  pendingReplacement = replace;

  walkParagraphs(docx, (pNode, path) => {
    while (!options.maxCount || replaced < options.maxCount) {
      if (!replaceOnceInParagraph(pNode, find, options.caseSensitive ?? false)) break;
      replaced++;
      touched.add(path);
    }
  });

  pendingReplacement = '';
  if (replaced) markDirty(docx, 'word/document.xml');
  return { replaced, paragraphs: [...touched] };
}

// ---- 结构原语：add / remove / move ----

export interface ParagraphSpec {
  kind?: 'paragraph';
  text?: string;
  props?: ParagraphProps;
  runs?: Array<{ text?: string; props?: RunProps }>;
  runProps?: RunProps;
}

/**
 * 构造段落节点。spec: { text, props(段落属性), runs:[{ text, props(run属性) }] }
 * runs 缺省时以单 run 承载 text。
 */
export function buildParagraph({ text = '', props = {}, runs = undefined, runProps = {} }: ParagraphSpec = {}): XmlNode {
  const p = el('w:p');
  if (Object.keys(props).length) setParagraphProps(p, props);
  const runSpecs = runs !== undefined ? runs : [{ text, props: runProps }];
  for (const spec of runSpecs) {
    const r = el('w:r');
    if (spec.props && Object.keys(spec.props).length) setRunProps(r, spec.props);
    childrenOf(r).push(textEl('w:t', spec.text ?? '', { 'xml:space': 'preserve' }));
    childrenOf(p).push(r);
  }
  return p;
}

export type AddPosition = 'end' | 'start' | { before: string } | { after: string };

/**
 * 在 parentPath 下插入节点（paragraph spec 或原始 fxp 节点）。
 * position: 'end'（默认）| 'start' | { before: path } | { after: path }
 */
export function addNode(docx: Docx, parentPath: string, nodeSpec: unknown, position: AddPosition = 'end'): { inserted: true } {
  const { node: parent } = resolveParent(docx, parentPath);
  const node = nodeSpec && typeof nodeSpec === 'object' && (nodeSpec as ParagraphSpec).kind === 'paragraph'
    ? buildParagraph(nodeSpec as ParagraphSpec)
    : (nodeSpec as XmlNode); // 原始 fxp 节点
  const siblings = childrenOf(parent);
  if (position === 'end') {
    // body 末尾的 sectPr 必须保持最后
    const last = siblings[siblings.length - 1];
    if (isElement(last, 'w:sectPr')) siblings.splice(siblings.length - 1, 0, node);
    else siblings.push(node);
  } else if (position === 'start') {
    siblings.unshift(node);
  } else if (position && typeof position === 'object' && 'before' in position) {
    siblings.splice(indexOfPath(docx, parent, position.before), 0, node);
  } else if (position && typeof position === 'object' && 'after' in position) {
    siblings.splice(indexOfPath(docx, parent, position.after) + 1, 0, node);
  }
  markDirty(docx, 'word/document.xml');
  return { inserted: true };
}

function resolveParent(docx: Docx, parentPath: string): { node: XmlNode } {
  return resolvePath(docx, parentPath);
}

function indexOfPath(docx: Docx, parent: XmlNode, path: string): number {
  const segs = parsePath(path);
  const lastSeg = segs[segs.length - 1];
  const matches = childrenOf(parent).filter((c) => isElement(c, lastSeg.tag));
  const target = lastSeg.index === 'last' ? matches[matches.length - 1] : matches[(lastSeg.index as number) - 1];
  if (!target) throw new Error(`定位插入点失败: ${path}`);
  return childrenOf(parent).indexOf(target);
}

// 删除路径指向的节点。禁止删除 body 末尾 sectPr。
export function removeNode(docx: Docx, path: string): { removed: true } {
  const { node, parent } = resolvePath(docx, path);
  if (!parent) throw new Error('不能删除根元素');
  if (isElement(node, 'w:sectPr') && isElement(parent, 'w:body')) {
    throw new Error('不能删除 body 级 sectPr（会破坏文档结构）；请用 set 修改节属性');
  }
  const arr = childrenOf(parent);
  arr.splice(arr.indexOf(node), 1);
  markDirty(docx, 'word/document.xml');
  return { removed: true };
}

// 移动节点到 parentPath 下（默认末尾；保持 body 末尾 sectPr 在最后）。
export function moveNode(docx: Docx, path: string, parentPath: string, position: AddPosition = 'end'): { moved: true } {
  const { node, parent } = resolvePath(docx, path);
  if (!parent) throw new Error('不能移动根元素');
  const srcArr = childrenOf(parent);
  srcArr.splice(srcArr.indexOf(node), 1);
  const { node: newParent } = resolvePath(docx, parentPath);
  const dstArr = childrenOf(newParent);
  if (position === 'start') dstArr.unshift(node);
  else {
    const last = dstArr[dstArr.length - 1];
    if (isElement(last, 'w:sectPr')) dstArr.splice(dstArr.length - 1, 0, node);
    else dstArr.push(node);
  }
  markDirty(docx, 'word/document.xml');
  return { moved: true };
}
