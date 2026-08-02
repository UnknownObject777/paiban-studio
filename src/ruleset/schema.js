// 模板规则集两文件 schema 校验器
// schema 草案见 issue #1；recognizers.json（component id → match 规则）
// 与 styles.json（component id → style + page）组件 id 绑定，
// 组件键集必须与 components.js 的 COMPONENT_IDS 完全一致。
//
// 纯 ESM、零依赖，浏览器（预览页）与 Node（加载器/测试）共用。

import { COMPONENT_IDS, PAGE_ID, getComponent } from './components.js';

export const MATCH_TYPES = Object.freeze(['regex', 'position', 'heuristic']);

export const ALIGN_VALUES = Object.freeze(['left', 'center', 'right', 'justify']);
export const PAGE_ALIGN_VALUES = Object.freeze(['left', 'center', 'right']);
export const PAPER_VALUES = Object.freeze(['A4', 'A3', 'A5', 'B5', 'preserve', 'custom']);

// 段落组件允许的样式属性
const PARAGRAPH_STYLE_KEYS = new Set([
  'fontEastAsia',
  'fontAscii',
  'sizePt',
  'bold',
  'italic',
  'align',
  'firstLineIndentChars',
  'lineSpacingPt',
  'lineSpacingMultiple',
  'spaceBeforePt',
  'spaceAfterPt',
  'pageBreakBefore',
  'outlineLevel',
  'notes',
]);

// 表格组件额外允许的样式属性
const TABLE_STYLE_KEYS = new Set([
  'fontEastAsia',
  'fontAscii',
  'sizePt',
  'headerFontEastAsia',
  'headerBold',
  'lineSpacingPt',
  'smartAlign',
  'borders',
  'widthPct',
  'notes',
]);

const NUMERIC_STYLE_KEYS = new Set([
  'sizePt',
  'firstLineIndentChars',
  'lineSpacingPt',
  'lineSpacingMultiple',
  'spaceBeforePt',
  'spaceAfterPt',
  'outlineLevel',
  'widthPct',
]);

const BOOLEAN_STYLE_KEYS = new Set([
  'bold',
  'italic',
  'pageBreakBefore',
  'smartAlign',
  'headerBold',
]);

function isPlainObject(value) {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function isNonNegativeNumber(value) {
  return typeof value === 'number' && Number.isFinite(value) && value >= 0;
}

function validateMatchEntry(componentId, entry, errors) {
  if (!isPlainObject(entry) || typeof entry.type !== 'string') {
    errors.push(`recognizers.${componentId}: match 条目必须是含 type 的对象`);
    return;
  }
  if (!MATCH_TYPES.includes(entry.type)) {
    errors.push(
      `recognizers.${componentId}: 未知 match 类型 "${entry.type}"（允许：${MATCH_TYPES.join('/')}）`,
    );
    return;
  }
  if (entry.type === 'regex') {
    if (typeof entry.pattern !== 'string' || entry.pattern.length === 0) {
      errors.push(`recognizers.${componentId}: regex 规则缺 pattern`);
      return;
    }
    try {
      new RegExp(entry.pattern, typeof entry.flags === 'string' ? entry.flags : 'u');
    } catch (err) {
      errors.push(`recognizers.${componentId}: 正则不可编译 "${entry.pattern}"（${err.message}）`);
    }
  }
}

function validateRecognizerComponent(componentId, recognizer, errors) {
  const prefix = `recognizers.${componentId}`;
  if (!isPlainObject(recognizer)) {
    errors.push(`${prefix}: 必须是对象`);
    return;
  }
  const hasMatch = Array.isArray(recognizer.match) && recognizer.match.length > 0;
  const hasFallback = recognizer.fallback === true;
  if (!hasMatch && !hasFallback) {
    errors.push(`${prefix}: 既无 match 规则也无 fallback: true`);
    return;
  }
  if (hasFallback && hasMatch) {
    errors.push(`${prefix}: fallback 组件不得再声明 match 规则`);
  }
  if (hasMatch) {
    for (const entry of recognizer.match) {
      validateMatchEntry(componentId, entry, errors);
    }
  }
}

function validateStyleComponent(componentId, style, errors) {
  const prefix = `styles.${componentId}`;
  if (!isPlainObject(style)) {
    errors.push(`${prefix}: 必须是对象`);
    return;
  }
  const component = getComponent(componentId);
  const allowed = component.kind === 'table' ? TABLE_STYLE_KEYS : PARAGRAPH_STYLE_KEYS;
  for (const key of Object.keys(style)) {
    if (!allowed.has(key)) {
      errors.push(`${prefix}: 未知或不适用的样式属性 "${key}"`);
      continue;
    }
    const value = style[key];
    if (NUMERIC_STYLE_KEYS.has(key) && !isNonNegativeNumber(value)) {
      errors.push(`${prefix}.${key}: 必须是非负有限数`);
    }
    if (key === 'align' && !ALIGN_VALUES.includes(value)) {
      errors.push(`${prefix}.align: 非法值 "${value}"（允许：${ALIGN_VALUES.join('/')}）`);
    }
    if (key === 'fontEastAsia' || key === 'fontAscii' || key === 'headerFontEastAsia') {
      if (typeof value !== 'string' || value.length === 0) {
        errors.push(`${prefix}.${key}: 必须是非空字符串`);
      }
    }
    for (const boolKey of BOOLEAN_STYLE_KEYS) {
      if (key === boolKey && typeof value !== 'boolean') {
        errors.push(`${prefix}.${key}: 必须是布尔值`);
      }
    }
    if (key === 'outlineLevel' && (!Number.isInteger(value) || value < 0 || value > 9)) {
      errors.push(`${prefix}.outlineLevel: 必须是 0–9 的整数`);
    }
  }
}

function validatePage(page, errors) {
  if (!isPlainObject(page)) {
    errors.push('styles.page: 必须是对象');
    return;
  }
  if (!PAPER_VALUES.includes(page.paper)) {
    errors.push(`styles.page.paper: 非法值（允许：${PAPER_VALUES.join('/')}）`);
  }
  if (page.paper === 'custom') {
    for (const key of ['widthCm', 'heightCm']) {
      if (!isNonNegativeNumber(page[key]) || page[key] === 0) {
        errors.push(`styles.page.${key}: custom 纸张必须提供正数 ${key}`);
      }
    }
  }
  if (!isPlainObject(page.margins)) {
    errors.push('styles.page.margins: 必须是对象');
  } else {
    for (const key of ['topCm', 'bottomCm', 'leftCm', 'rightCm']) {
      if (!isNonNegativeNumber(page.margins[key])) {
        errors.push(`styles.page.margins.${key}: 必须是非负有限数`);
      }
    }
  }
  if (!isNonNegativeNumber(page.footerDistanceCm)) {
    errors.push('styles.page.footerDistanceCm: 必须是非负有限数');
  }
  const pn = page.pageNumber;
  if (!isPlainObject(pn)) {
    errors.push('styles.page.pageNumber: 必须是对象');
    return;
  }
  if (typeof pn.oddEven !== 'boolean') {
    errors.push('styles.page.pageNumber.oddEven: 必须是布尔值');
  }
  for (const key of ['oddAlign', 'evenAlign']) {
    if (!PAGE_ALIGN_VALUES.includes(pn[key])) {
      errors.push(`styles.page.pageNumber.${key}: 非法值（允许：${PAGE_ALIGN_VALUES.join('/')}）`);
    }
  }
  if (!isNonNegativeNumber(pn.sizePt) || pn.sizePt === 0) {
    errors.push('styles.page.pageNumber.sizePt: 必须是正数');
  }
}

/**
 * 校验一份规则集（recognizers + styles 两文件已解析为对象）。
 * 返回错误信息数组，空数组表示通过。
 */
export function validateRuleset(recognizers, styles) {
  const errors = [];
  const files = [
    ['recognizers', recognizers],
    ['styles', styles],
  ];

  for (const [label, data] of files) {
    if (!isPlainObject(data)) {
      errors.push(`${label}: 必须是对象`);
    }
  }
  if (errors.length > 0) return errors;

  if (typeof recognizers.ruleset !== 'string' || recognizers.ruleset.length === 0) {
    errors.push('recognizers.ruleset: 必须是非空字符串');
  }
  if (recognizers.ruleset !== styles.ruleset) {
    errors.push(
      `两文件 ruleset id 不一致："${recognizers.ruleset}" vs "${styles.ruleset}"`,
    );
  }
  if (!Number.isInteger(recognizers.version) || recognizers.version < 1) {
    errors.push('recognizers.version: 必须是 ≥1 的整数');
  }
  if (recognizers.version !== styles.version) {
    errors.push(`两文件 version 不一致：${recognizers.version} vs ${styles.version}`);
  }

  for (const [label, data] of files) {
    if (!isPlainObject(data.components)) {
      errors.push(`${label}.components: 必须是对象`);
    }
  }
  if (errors.length > 0) return errors;

  for (const [label, data] of files) {
    const ids = new Set(Object.keys(data.components));
    for (const id of COMPONENT_IDS) {
      if (!ids.has(id)) errors.push(`${label}.components: 缺少组件 "${id}"`);
    }
    for (const id of ids) {
      if (!COMPONENT_IDS.includes(id)) errors.push(`${label}.components: 未知组件 "${id}"`);
    }
  }
  if (recognizers[PAGE_ID] !== undefined) {
    errors.push('recognizers 不得含 page 节（页面设置只属 styles.json）');
  }

  // 识别规则
  let fallbackCount = 0;
  for (const id of COMPONENT_IDS) {
    const recognizer = recognizers.components[id];
    if (isPlainObject(recognizer) && recognizer.fallback === true) fallbackCount += 1;
    validateRecognizerComponent(id, recognizer, errors);
  }
  if (fallbackCount !== 1) {
    errors.push(`fallback 组件必须恰好 1 个，当前 ${fallbackCount} 个`);
  }

  // 样式
  for (const id of COMPONENT_IDS) {
    validateStyleComponent(id, styles.components[id], errors);
  }

  validatePage(styles[PAGE_ID], errors);

  return errors;
}
