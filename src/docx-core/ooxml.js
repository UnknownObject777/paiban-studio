// docx-core/ooxml.js — fast-xml-parser preserveOrder 树的通用操作助手。
//
// fxp preserveOrder 节点形态：
//   元素节点: { "w:p": [ ...children ], ":@": { "@_w:val": "center" } }
//   文本节点: { "#text": "..." }
// 本模块提供节点构造、属性读写、子元素查找，以及**按 OOXML schema 顺序插入子元素**
// （Word 对 pPr/rPr/sectPr 的子元素顺序敏感，乱序插入会导致 Word 打不开或忽略属性）。

// ---- 基础访问 ----

export function el(tag, attrs = {}, children = []) {
  const node = { [tag]: children };
  const keys = Object.keys(attrs);
  if (keys.length) {
    node[':@'] = Object.fromEntries(keys.map((k) => ['@_' + k, String(attrs[k])]));
  }
  return node;
}

export function textEl(tag, text, attrs = {}) {
  return el(tag, attrs, text === '' ? [] : [{ '#text': String(text) }]);
}

// 元素节点的标签名（跳过 :@ 属性键与 #text 文本节点）
export function tagOf(node) {
  if (!node || typeof node !== 'object') return undefined;
  for (const k of Object.keys(node)) if (k !== ':@' && !k.startsWith('#')) return k;
  return undefined;
}

export function isElement(node, tag = undefined) {
  const t = tagOf(node);
  return t !== undefined && (tag === undefined || t === tag);
}

export function attrsOf(node) {
  return node[':@'] || {};
}

export function getAttr(node, name) {
  return attrsOf(node)['@_' + name];
}

export function setAttr(node, name, value) {
  if (!node[':@']) node[':@'] = {};
  node[':@']['@_' + name] = String(value);
}

export function removeAttr(node, name) {
  if (node[':@']) delete node[':@']['@_' + name];
}

// 元素节点的子节点数组（不含文本节点的便捷访问请用 findChildren）
export function childrenOf(node) {
  const t = tagOf(node);
  return t ? node[t] : [];
}

export function findChild(node, tag) {
  return childrenOf(node).find((c) => isElement(c, tag));
}

export function findChildren(node, tag) {
  return childrenOf(node).filter((c) => isElement(c, tag));
}

// 节点拼接文本（递归收集 #text）
export function textOf(node) {
  let out = '';
  for (const c of childrenOf(node)) {
    if (c['#text'] !== undefined) out += c['#text'];
    else out += textOf(c);
  }
  return out;
}

// ---- schema 顺序插入 ----

// w:pPr 子元素顺序（CT_PPr 序列，取 MVP 用到的子集，保持相对顺序即可）
export const PPR_ORDER = [
  'w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr',
  'w:widowControl', 'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd',
  'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap', 'w:overflowPunct',
  'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd',
  'w:snapToGrid', 'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents',
  'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment',
  'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr',
  'w:sectPr', 'w:pPrChange',
];

// w:rPr 子元素顺序（CT_RPr 序列子集）
export const RPR_ORDER = [
  'w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps',
  'w:strike', 'w:dstrike', 'w:outline', 'w:shadow', 'w:emboss', 'w:imprint',
  'w:noProof', 'w:snapToGrid', 'w:vanish', 'w:webHidden', 'w:color', 'w:spacing',
  'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs', 'w:highlight', 'w:u',
  'w:effect', 'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs',
  'w:em', 'w:lang', 'w:eastAsianLayout', 'w:specVanish', 'w:oMath',
];

// w:sectPr 子元素顺序（CT_SectPr 序列子集）
export const SECTPR_ORDER = [
  'w:headerReference', 'w:footerReference', 'w:footnotePr', 'w:endnotePr',
  'w:type', 'w:pgSz', 'w:pgMar', 'w:paperSrc', 'w:pgBorders', 'w:lnNumType',
  'w:pgNumType', 'w:cols', 'w:formProt', 'w:vAlign', 'w:noEndnote', 'w:titlePg',
  'w:textDirection', 'w:bidi', 'w:rtlGutter', 'w:docGrid', 'w:printerSettings',
  'w:sectPrChange',
];

// w:p / w:tbl 级：属性元素（pPr/tblPr）必须排在内容之前
const PARA_ORDER = ['w:pPr', 'w:rPr'];

function orderIndex(order, tag) {
  const i = order.indexOf(tag);
  return i === -1 ? order.length : i;
}

// 在 parent 的子节点中按给定顺序表插入 child（同 tag 已存在时不插入，返回既有节点）。
export function insertOrdered(parent, child, order) {
  const tag = tagOf(child);
  const children = childrenOf(parent);
  const existing = children.find((c) => isElement(c, tag));
  if (existing) return existing;
  const want = orderIndex(order, tag);
  let pos = children.length;
  for (let i = 0; i < children.length; i++) {
    const t = tagOf(children[i]);
    if (t === undefined) continue; // 文本节点视为最前
    if (orderIndex(order, t) > want) {
      pos = i;
      break;
    }
  }
  children.splice(pos, 0, child);
  return child;
}

// 确保 parent 下存在某个子元素（不存在则按顺序创建插入）。
export function ensureChild(parent, tag, order) {
  const found = findChild(parent, tag);
  if (found) return found;
  return insertOrdered(parent, el(tag), order);
}

// 在 pPr/rPr 中确保一个**叶子设置元素**（如 w:jc / w:sz），并按 attrs 覆盖属性。
export function ensureLeaf(parent, tag, order, attrs) {
  const leaf = ensureChild(parent, tag, order);
  for (const [k, v] of Object.entries(attrs)) setAttr(leaf, k, v);
  return leaf;
}

// 移除 parent 下所有 tag 子元素。
export function removeChildren(parent, tag) {
  const t = tagOf(parent);
  parent[t] = childrenOf(parent).filter((c) => !isElement(c, tag));
}

// 确保段落/表格属性元素（w:pPr 等）位于内容之前。
export function ensurePropContainer(node, propTag) {
  const found = findChild(node, propTag);
  if (found) return found;
  const prop = el(propTag);
  childrenOf(node).unshift(prop);
  return prop;
}

// 在段落中按 schema 设置 pPr 子元素。
export function setPPrLeaf(pPr, tag, attrs) {
  return ensureLeaf(pPr, tag, PPR_ORDER, attrs);
}

export function setRPrLeaf(rPr, tag, attrs) {
  return ensureLeaf(rPr, tag, RPR_ORDER, attrs);
}

export { PARA_ORDER };
