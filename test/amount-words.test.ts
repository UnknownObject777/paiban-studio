// amount-words.test.ts — 人民币金额大写工具（amountWords.ts + amount_words 工具接线）单测。
//
// 覆盖：
//   1. 参考断言（财务规范核心：零的衔接、万/亿级间零、末尾零、整/角/分）
//   2. 边界：0.x / x.0x / 20 万整零衔接 / 千亿级上限 / 前导零 / 千分位逗号（含全角）
//   3. 四舍五入到分：字符串按精确十进制（"1.005" → 壹元零壹分）；
//      number 按 double 精确值（1.005 实际是 1.00499… → 壹元整）；进位到整数部分
//   4. 错误路径：负数（number/string/-0）、非有限数、格式非法、超出千亿级
//   5. 工具路径：经 createTools 调 amount_words 的 execute，断言返回 words 与错误语义

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { amountToWordsCn } from '../src/docgen/amountWords.js';
import { createTools } from '../src/agent-core/tools.js';

test('参考断言：财务规范核心用例全对', () => {
  assert.equal(amountToWordsCn(0), '零元整');
  assert.equal(amountToWordsCn(5), '伍元整');
  assert.equal(amountToWordsCn(10), '壹拾元整');
  assert.equal(amountToWordsCn(100.07), '壹佰元零柒分');
  assert.equal(amountToWordsCn(1001), '壹仟零壹元整');
  assert.equal(amountToWordsCn(10010), '壹万零壹拾元整');
  assert.equal(amountToWordsCn(100000), '壹拾万元整');
  assert.equal(amountToWordsCn(100000000), '壹亿元整');
  assert.equal(amountToWordsCn(100200300.4), '壹亿零贰拾万零叁佰元肆角');
  assert.equal(amountToWordsCn(12345.67), '壹万贰仟叁佰肆拾伍元陆角柒分');
  assert.equal(amountToWordsCn(1000001), '壹佰万零壹元整');
  assert.equal(amountToWordsCn('1,234.5'), '壹仟贰佰叁拾肆元伍角');
});

test('边界：分位 / 角位 / 20 万级零衔接 / 亿级零衔接', () => {
  assert.equal(amountToWordsCn(10.02), '壹拾元零贰分');          // 角位 0 分位非 0 → 元后补零
  assert.equal(amountToWordsCn(0.5), '零元伍角');                // 整数 0 有小数 → 零元前缀
  assert.equal(amountToWordsCn(0.07), '零元零柒分');
  assert.equal(amountToWordsCn(200000), '贰拾万元整');            // 20 万整零衔接
  assert.equal(amountToWordsCn(200000.5), '贰拾万元伍角');
  assert.equal(amountToWordsCn(100000000.01), '壹亿元零壹分');     // 亿 + 角位 0 分位非 0
  assert.equal(amountToWordsCn(101000000), '壹亿零壹佰万元整');    // 亿级间零衔接
  assert.equal(amountToWordsCn(100010000), '壹亿零壹万元整');
  assert.equal(amountToWordsCn(100000010), '壹亿零壹拾元整');
  assert.equal(amountToWordsCn(0.004), '零元整');                 // 不足半分 → 舍去
  assert.equal(amountToWordsCn(0.005), '零元零壹分');             // 半分 → 进位
});

test('千亿级上限：999,999,999,999.99 全大写正确', () => {
  assert.equal(
    amountToWordsCn('999999999999.99'),
    '玖仟玖佰玖拾玖亿玖仟玖佰玖拾玖万玖仟玖佰玖拾玖元玖角玖分',
  );
});

test('输入归一化：前导零 / 全角逗号 / 首尾空白 / number 大值', () => {
  assert.equal(amountToWordsCn('0005'), '伍元整');
  assert.equal(amountToWordsCn(' 12，345.67 '), '壹万贰仟叁佰肆拾伍元陆角柒分'); // 全角逗号 + 空白
  assert.equal(amountToWordsCn(100000000), '壹亿元整'); // number 整数大值走 toFixed(16) 精确串
  assert.equal(amountToWordsCn(1234.5), '壹仟贰佰叁拾肆元伍角');
});

test('四舍五入到分：字符串按精确十进制，number 按 double 精确值', () => {
  // 字符串 "1.005" 第 3 位小数 5 → 半分进位 → 壹元零壹分
  assert.equal(amountToWordsCn('1.005'), '壹元零壹分');
  assert.equal(amountToWordsCn('1.235'), '壹元贰角肆分'); // 第 3 位 5 → 进位
  assert.equal(amountToWordsCn('1.234'), '壹元贰角叁分'); // 第 3 位 4 → 舍去
  // number 1.005 实际是 1.00499…（double 精确值）→ 舍到 1.00 → 壹元整
  assert.equal(amountToWordsCn(1.005), '壹元整');
  assert.equal(amountToWordsCn(2.675), '贰元陆角柒分'); // double 精确值 2.67499… → 不进位
  // 进位波及整数部分：9.999 → 10.00
  assert.equal(amountToWordsCn('9.999'), '壹拾元整');
  assert.equal(amountToWordsCn(9.999), '壹拾元整'); // number 9.999 = 9.998999… → 进位到 10.00
});

test('错误路径：负数 / 非有限数 / 格式非法 / 超出千亿级', () => {
  // 负数（中文 Error 信息）
  assert.throws(() => amountToWordsCn(-5), /金额不能为负数/);
  assert.throws(() => amountToWordsCn('-5'), /金额不能为负数/);
  assert.throws(() => amountToWordsCn('-0'), /金额不能为负数/);
  // 非有限数
  assert.throws(() => amountToWordsCn(NaN), /必须是有限数字/);
  assert.throws(() => amountToWordsCn(Infinity), /必须是有限数字/);
  // 格式非法
  assert.throws(() => amountToWordsCn(''), /格式非法/);
  assert.throws(() => amountToWordsCn('abc'), /格式非法/);
  assert.throws(() => amountToWordsCn('12a4'), /格式非法/);
  assert.throws(() => amountToWordsCn('1.2.3'), /格式非法/);
  // 超出千亿级（13 位整数即万亿）
  assert.throws(() => amountToWordsCn(1000000000000), /超出支持范围/);
  assert.throws(() => amountToWordsCn('1000000000000'), /超出支持范围/);
  // 千亿级封顶值四舍五入后越界 → 同样报错
  assert.throws(() => amountToWordsCn('999999999999.995'), /超出支持范围/);
  // 文档化语义：number -0 数值上不小于 0 → 视为 0
  assert.equal(amountToWordsCn(-0), '零元整');
});

test('amount_words 工具：经 createTools execute 返回 words；错误走 isError', async () => {
  const byName = Object.fromEntries(createTools({} as never).map((t) => [t.name, t]));
  const tool = byName.amount_words;
  assert.ok(tool, 'createTools 含 amount_words');
  assert.equal(tool.executionMode, 'sequential');

  const res = await tool.execute('t1', { amount: '12345.67' });
  assert.equal(res.isError, false);
  assert.deepEqual(JSON.parse(res.content[0].text), {
    amount: '12345.67',
    words: '壹万贰仟叁佰肆拾伍元陆角柒分',
  });

  // 千分位逗号输入
  const res2 = await tool.execute('t2', { amount: '1,234.5' });
  assert.equal(JSON.parse(res2.content[0].text).words, '壹仟贰佰叁拾肆元伍角');

  // 错误路径：负数 → isError + 中文错误
  const neg = await tool.execute('t3', { amount: '-5' });
  assert.equal(neg.isError, true);
  assert.ok(JSON.parse(neg.content[0].text).error.includes('金额不能为负数'));
  // 空字符串 → isError
  const empty = await tool.execute('t4', { amount: '   ' });
  assert.equal(empty.isError, true);
  assert.ok(JSON.parse(empty.content[0].text).error.includes('不能为空'));
  // 金额 "0" 是合法输入（falsy 但不能当空）
  const zero = await tool.execute('t5', { amount: '0' });
  assert.equal(JSON.parse(zero.content[0].text).words, '零元整');
});
