// docx-core/numbering.ts — numbering.xml 多级编号封装（spec 标记的首要风险点之一，单独成模块重点测试）。
//
// 职责：
//   1. 在 numbering.xml 中定义多级编号（abstractNum + num），幂等（重复定义复用既有 numId）
//   2. 给段落挂载/卸载编号引用（w:numPr）
//   3. 公文常用预设：多级标题 "一、 / （一） / 1. / （1）"（四级中文编号）
//
// numbering.xml 结构：<w:numbering><w:abstractNum w:abstractNumId="N">...<w:lvl w:ilvl="0">...
//   </w:abstractNum>...<w:num w:numId="M"><w:abstractNumId w:val="N"/></w:num>...</w:numbering>

import {
  el, isElement, getAttr, setAttr,
  childrenOf, findChild, findChildren, ensurePropContainer, setPPrLeaf,
} from './ooxml.js';
import { getXmlTree, markDirty } from './docx.js';
import { parseXml } from './xml.js';
import type { XmlNode } from './xml.js';
import type { Docx } from './docx.js';

const NUMBERING_PART = 'word/numbering.xml';
const NUMBERING_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
const NUMBERING_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';

/** 编号级别定义（numbering define 命令的 levels 元素）。 */
export interface NumberingLevelDef {
  ilvl: number;
  numFmt: string;
  lvlText: string;
  start?: number;
  indentPt?: number;
  hangingChars?: number;
  runProps?: { eastAsia?: string; ascii?: string; sizePt?: number };
}

// 公文多级标题预设：一、 / （一） / 1. / （1）
export const GONGWEN_HEADING_LEVELS: NumberingLevelDef[] = [
  { ilvl: 0, numFmt: 'chineseCounting', lvlText: '%1、', indentPt: 0, hangingChars: 0 },
  { ilvl: 1, numFmt: 'chineseCounting', lvlText: '（%2）', indentPt: 0, hangingChars: 0 },
  { ilvl: 2, numFmt: 'decimal', lvlText: '%3.', indentPt: 0, hangingChars: 0 },
  { ilvl: 3, numFmt: 'decimalEnclosedParen', lvlText: '%4', indentPt: 0, hangingChars: 0 },
];

function numberingRoot(docx: Docx): XmlNode | undefined {
  const tree = getXmlTree(docx, NUMBERING_PART);
  if (!tree) return undefined;
  return tree.find((n) => isElement(n, 'w:numbering'));
}

// 确保 numbering.xml 部件存在（含 rels / content-types 注册）。
function ensureNumberingPart(docx: Docx): XmlNode {
  let root = numberingRoot(docx);
  if (root) return root;

  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"></w:numbering>';
  const { tree, meta } = parseXml(xml);
  docx.parts.set(NUMBERING_PART, { kind: 'xml', tree, meta, text: xml, dirty: true });
  root = tree.find((n) => isElement(n, 'w:numbering'))!;

  // rels 注册
  const relsTree = getXmlTree(docx, 'word/_rels/document.xml.rels')!;
  const relsRoot = relsTree.find((n) => isElement(n, 'Relationships'));
  const ids = findChildren(relsRoot!, 'Relationship').map((r) => getAttr(r, 'Id'));
  let max = 0;
  for (const id of ids) {
    const m = /^rId(\d+)$/.exec(id || '');
    if (m) max = Math.max(max, Number(m[1]));
  }
  childrenOf(relsRoot!).push(el('Relationship', {
    Id: 'rId' + (max + 1), Type: NUMBERING_REL_TYPE, Target: 'numbering.xml',
  }));
  markDirty(docx, 'word/_rels/document.xml.rels');

  // content-types 注册
  const ctTree = getXmlTree(docx, '[Content_Types].xml')!;
  const ctRoot = ctTree.find((n) => isElement(n, 'Types'));
  const has = findChildren(ctRoot!, 'Override').some((o) => getAttr(o, 'PartName') === '/' + NUMBERING_PART);
  if (!has) {
    childrenOf(ctRoot!).push(el('Override', { PartName: '/' + NUMBERING_PART, ContentType: NUMBERING_CT }));
    markDirty(docx, '[Content_Types].xml');
  }
  return root;
}

function nextIds(root: XmlNode): { absId: number; numId: number } {
  let maxAbs = 0, maxNum = 0;
  for (const a of findChildren(root, 'w:abstractNum')) {
    maxAbs = Math.max(maxAbs, Number(getAttr(a, 'w:abstractNumId') || 0) + 1);
  }
  for (const n of findChildren(root, 'w:num')) {
    maxNum = Math.max(maxNum, Number(getAttr(n, 'w:numId') || 0) + 1);
  }
  return { absId: maxAbs, numId: Math.max(maxNum, 1) };
}

// 构造一级 lvl 定义。
function buildLvl(def: NumberingLevelDef): XmlNode {
  const { ilvl, numFmt, lvlText, start = 1, indentPt = 0, hangingChars = 0, runProps } = def;
  const lvlChildren = [
    el('w:start', { 'w:val': start }),
    el('w:numFmt', { 'w:val': numFmt }),
    el('w:lvlText', { 'w:val': lvlText }),
    el('w:lvlJc', { 'w:val': 'left' }),
  ];
  const ind: Record<string, string | number> = {};
  if (indentPt) ind['w:left'] = Math.round(indentPt * 20);
  if (hangingChars) ind['w:hangingChars'] = Math.round(hangingChars);
  if (Object.keys(ind).length) lvlChildren.push(el('w:pPr', {}, [el('w:ind', ind)]));
  if (runProps) {
    const rPr = el('w:rPr');
    if (runProps.eastAsia || runProps.ascii) {
      const fonts: Record<string, string> = {};
      if (runProps.eastAsia) fonts['w:eastAsia'] = runProps.eastAsia;
      if (runProps.ascii) fonts['w:ascii'] = runProps.ascii;
      childrenOf(rPr).push(el('w:rFonts', fonts));
    }
    if (runProps.sizePt) {
      childrenOf(rPr).push(el('w:sz', { 'w:val': Math.round(runProps.sizePt * 2) }));
    }
    lvlChildren.push(rPr);
  }
  return el('w:lvl', { 'w:ilvl': ilvl }, lvlChildren);
}

// 两级定义的签名（用于幂等判重）。
function lvlSignature(defs: NumberingLevelDef[]): string {
  return defs.map((d) => `${d.ilvl}:${d.numFmt}:${d.lvlText}`).join('|');
}

/**
 * 定义多级编号，返回 numId。幂等：相同签名（lvl 序列）复用既有 num。
 * levels: [{ ilvl, numFmt, lvlText, start?, indentPt?, hangingChars?, runProps? }]
 */
export function defineNumbering(docx: Docx, levels: NumberingLevelDef[]): number {
  const root = ensureNumberingPart(docx);
  const sig = lvlSignature(levels);

  // 幂等：检查既有 abstractNum 是否同签名
  for (const abs of findChildren(root, 'w:abstractNum')) {
    const lvls = findChildren(abs, 'w:lvl');
    const existingSig = lvlSignature(lvls.map((l) => ({
      ilvl: Number(getAttr(l, 'w:ilvl')),
      numFmt: getAttr(findChild(l, 'w:numFmt') || el('w:numFmt'), 'w:val') || '',
      lvlText: getAttr(findChild(l, 'w:lvlText') || el('w:lvlText'), 'w:val') || '',
    })));
    if (existingSig === sig) {
      const absId = getAttr(abs, 'w:abstractNumId');
      const num = findChildren(root, 'w:num').find(
        (n) => getAttr(findChild(n, 'w:abstractNumId') || el('w:abstractNumId'), 'w:val') === absId);
      if (num) return Number(getAttr(num, 'w:numId'));
    }
  }

  const { absId, numId } = nextIds(root);
  const abs = el('w:abstractNum', { 'w:abstractNumId': absId }, [
    el('w:multiLevelType', { 'w:val': 'multilevel' }),
    ...levels.map(buildLvl),
  ]);
  // schema：abstractNum 全部在 num 之前
  const firstNum = findChild(root, 'w:num');
  const arr = childrenOf(root);
  if (firstNum) arr.splice(arr.indexOf(firstNum), 0, abs);
  else arr.push(abs);
  arr.push(el('w:num', { 'w:numId': numId }, [el('w:abstractNumId', { 'w:val': absId })]));
  markDirty(docx, NUMBERING_PART);
  return numId;
}

/** 给段落挂载编号引用。ilvl 默认 0。 */
export function setParagraphNumbering(pNode: XmlNode, numId: number, ilvl = 0): { numId: number; ilvl: number } {
  const pPr = ensurePropContainer(pNode, 'w:pPr');
  const numPr = setPPrLeaf(pPr, 'w:numPr', {});
  // numPr 子元素顺序：ilvl 在前 numId 在后
  let ilvlEl = findChild(numPr, 'w:ilvl');
  if (!ilvlEl) {
    ilvlEl = el('w:ilvl', { 'w:val': ilvl });
    childrenOf(numPr).unshift(ilvlEl);
  } else setAttr(ilvlEl, 'w:val', ilvl);
  let numIdEl = findChild(numPr, 'w:numId');
  if (!numIdEl) {
    numIdEl = el('w:numId', { 'w:val': numId });
    childrenOf(numPr).push(numIdEl);
  } else setAttr(numIdEl, 'w:val', numId);
  return { numId, ilvl };
}

/** 卸载段落编号。 */
export function clearParagraphNumbering(pNode: XmlNode): boolean {
  const pPr = findChild(pNode, 'w:pPr');
  if (!pPr) return false;
  const before = childrenOf(pPr).length;
  pPr['w:pPr'] = childrenOf(pPr).filter((c) => !isElement(c, 'w:numPr'));
  return childrenOf(pPr).length !== before;
}
