import test from 'node:test';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

import { validateRuleset } from '../src/ruleset/schema.js';
import { loadRuleset } from '../src/ruleset/load.js';
import { COMPONENT_IDS, PAGE_ID } from '../src/ruleset/components.js';

const REPO_ROOT = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const BUILTIN_DIR = path.join(REPO_ROOT, 'templates', 'rulesets', 'gongwen-default');

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function validPair() {
  const { recognizers, styles } = loadRuleset(BUILTIN_DIR);
  return { recognizers: clone(recognizers), styles: clone(styles) };
}

// ---------- 内置规则集（正例）----------

test('内置公文默认规则集通过校验', () => {
  const { recognizers, styles } = loadRuleset(BUILTIN_DIR);
  assert.deepEqual(validateRuleset(recognizers, styles), []);
});

test('内置规则集两文件组件键集与组件清单常量一致', () => {
  const { recognizers, styles } = loadRuleset(BUILTIN_DIR);
  assert.deepEqual(Object.keys(recognizers.components).sort(), [...COMPONENT_IDS].sort());
  assert.deepEqual(Object.keys(styles.components).sort(), [...COMPONENT_IDS].sort());
  assert.ok(styles[PAGE_ID], 'styles.json 必须含 page 节');
  assert.equal(recognizers[PAGE_ID], undefined, 'recognizers.json 不得含 page 节');
});

test('loadRuleset 对缺失文件给出可读错误', () => {
  assert.throws(
    () => loadRuleset(path.join(REPO_ROOT, 'templates', 'rulesets', 'no-such-set')),
    /recognizers\.json/,
  );
});

// ---------- 一致性校验（负例） ----------

test('组件缺失：recognizers 少一个组件时报错并点名', () => {
  const { recognizers, styles } = validPair();
  delete recognizers.components.heading3;
  const errors = validateRuleset(recognizers, styles);
  assert.ok(errors.some((e) => e.includes('heading3')), errors.join('\n'));
});

test('组件多余：styles 多一个未知组件时报错', () => {
  const { recognizers, styles } = validPair();
  styles.components.heading5 = { sizePt: 16 };
  const errors = validateRuleset(recognizers, styles);
  assert.ok(errors.some((e) => e.includes('heading5')), errors.join('\n'));
});

test('两文件 ruleset id 不一致时报错', () => {
  const { recognizers, styles } = validPair();
  styles.ruleset = 'other-set';
  assert.ok(validateRuleset(recognizers, styles).some((e) => e.includes('ruleset')));
});

test('两文件 version 不一致时报错', () => {
  const { recognizers, styles } = validPair();
  styles.version = 2;
  assert.ok(validateRuleset(recognizers, styles).some((e) => e.includes('version')));
});

// ---------- 识别规则校验（负例） ----------

test('正则不可编译时报错并点名组件', () => {
  const { recognizers, styles } = validPair();
  recognizers.components.heading1.match = [{ type: 'regex', pattern: '([' }];
  const errors = validateRuleset(recognizers, styles);
  assert.ok(errors.some((e) => e.includes('heading1')), errors.join('\n'));
});

test('未知 match 类型时报错', () => {
  const { recognizers, styles } = validPair();
  recognizers.components.title.match = [{ type: 'magic' }];
  assert.ok(validateRuleset(recognizers, styles).some((e) => e.includes('magic')));
});

test('组件既无 match 也无 fallback 时报错', () => {
  const { recognizers, styles } = validPair();
  delete recognizers.components.body.fallback;
  assert.ok(validateRuleset(recognizers, styles).some((e) => e.includes('body')));
});

test('fallback 必须恰好一个：零个或两个都报错', () => {
  {
    const { recognizers, styles } = validPair();
    recognizers.components.caption.fallback = true;
    assert.ok(validateRuleset(recognizers, styles).some((e) => e.includes('fallback')));
  }
  {
    const { recognizers, styles } = validPair();
    delete recognizers.components.body.fallback;
    recognizers.components.body.match = [{ type: 'regex', pattern: '^$' }];
    assert.ok(validateRuleset(recognizers, styles).some((e) => e.includes('fallback')));
  }
});

// ---------- 样式校验（负例） ----------

test('样式未知属性报错', () => {
  const { recognizers, styles } = validPair();
  styles.components.body.fontWeight = 700;
  assert.ok(validateRuleset(recognizers, styles).some((e) => e.includes('fontWeight')));
});

test('段落组件不允许表格专属属性 smartAlign', () => {
  const { recognizers, styles } = validPair();
  styles.components.body.smartAlign = true;
  assert.ok(validateRuleset(recognizers, styles).some((e) => e.includes('smartAlign')));
});

test('align 非法值报错', () => {
  const { recognizers, styles } = validPair();
  styles.components.title.align = 'middle';
  assert.ok(validateRuleset(recognizers, styles).some((e) => e.includes('align')));
});

test('数值字段必须为非负有限数', () => {
  const { recognizers, styles } = validPair();
  styles.components.body.sizePt = '三号';
  assert.ok(validateRuleset(recognizers, styles).some((e) => e.includes('sizePt')));
});

// ---------- 页面节校验（负例） ----------

test('page 缺 margins 字段报错', () => {
  const { recognizers, styles } = validPair();
  delete styles.page.margins.topCm;
  assert.ok(validateRuleset(recognizers, styles).some((e) => e.includes('topCm')));
});

test('pageNumber 对齐值非法时报错', () => {
  const { recognizers, styles } = validPair();
  styles.page.pageNumber.oddAlign = 'justify';
  assert.ok(validateRuleset(recognizers, styles).some((e) => e.includes('oddAlign')));
});
