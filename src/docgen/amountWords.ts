// docgen/amountWords.ts — 人民币金额大写（零依赖纯函数）。
//
// 确定性工具：金额 → 中文大写（财务规范），供 agent 填报价表 / 外汇申报单的
// 「金额（大写）」栏，替代 LLM 心算（心算大写经常出错）。
//
//   amountToWordsCn(input: number | string): string
//
// 规则：
//   - 接受 number 或字符串（可含 ASCII/全角千分位逗号，如 "1,234.5"）。
//   - 最多两位小数，超出按「四舍五入到分」（第 3 位小数 ≥5 进 1）。
//   - 范围：0 ~ 999,999,999,999.99（千亿级）；负数与超出范围抛中文 Error。
//   - 大写用 零壹贰叁肆伍陆柒捌玖 + 拾佰仟万亿 + 元角分整。
//   - 零的处理：中间连续零只读一个「零」；万/亿级间零照常读出；末尾零不读；
//     无小数部分（角分均为 0）结尾加「整」；恰为 0 → 「零元整」。
//   - 角位为 0 而分位非 0 时，元后写「零」（如 100.07 → 壹佰元零柒分）；
//     整数部分为 0 且有小数部分时写「零元」（如 0.5 → 零元伍角）；
//     角后不写「整」（与 100200300.4 → …元肆角 的惯例一致）。
//
// 精度语义（避免浮点坑）：
//   - number 输入用 toFixed(16) 拿到 double 的精确十进制再解析——1.005 实际是
//     1.00499…，四舍五入到分 → 壹元整；不会因 toString 最近表示丢尾数。
//   - 字符串输入按精确十进制四舍五入——"1.005" 第 3 位小数 5 → 壹元零壹分。
//   - 负数判定：number 用 < 0（-0 视为 0）；字符串以 '-' 开头的任何输入（含 "-0"）一律抛错。

const DIGITS = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];
const GROUP_UNITS = ['', '拾', '佰', '仟']; // 组内单位（个级：拾佰仟）
const LEVEL_UNITS = ['', '万', '亿'];       // 组级单位（4 位一组；千亿级封顶，亿为最高级）

const AMOUNT_RE = /^[+-]?\d+(\.\d+)?$/;
const MAX_INT_DIGITS = 12; // 999,999,999,999（千亿级封顶）
const MAX_AMOUNT = 999_999_999_999.99; // 千亿级封顶值（number 路径预检用）

/** 金额 → 中文大写。规则与精度语义见文件头注释。 */
export function amountToWordsCn(input: number | string): string {
  const { intDigits, fen } = parseAmount(input);
  const intWords = intToWords(intDigits);
  return composeWords(intWords, fen);
}

// ---- 解析：输入 → 整数部分数字串 + 分（0..99） ----

function parseAmount(input: number | string): { intDigits: string; fen: number } {
  let text: string;
  if (typeof input === 'number') {
    if (!Number.isFinite(input)) throw new Error(`金额必须是有限数字：${input}`);
    if (input < 0) throw new Error(`金额不能为负数：${input}`);
    if (input > MAX_AMOUNT) throw new Error(`金额超出支持范围（最多到千亿级 999,999,999,999.99）：${input}`);
    // toFixed(16)：double 的精确十进制（17 位有效数字往返），避免 toString 丢尾数；
    // 该函数永不产生指数形式，随后按字符串路径统一做四舍五入。
    text = trimDecimalZeros(input.toFixed(16));
  } else if (typeof input === 'string') {
    text = input.trim();
  } else {
    throw new Error(`金额类型不支持（应为 number 或 string）：${String(input)}`);
  }

  const cleaned = text.replace(/[,，]/g, '');
  if (!AMOUNT_RE.test(cleaned)) {
    throw new Error(`金额格式非法：${JSON.stringify(input)}（应为数字，可含千分位逗号，最多两位小数）`);
  }
  if (cleaned.startsWith('-')) {
    throw new Error(`金额不能为负数：${JSON.stringify(input)}`);
  }

  let [intPart, fracPart = ''] = cleaned.replace(/^[+]/, '').split('.');
  intPart = intPart.replace(/^0+/, '') || '0';

  // 四舍五入到分：第 3 位小数 ≥5 则分进 1（进位可波及整数部分）。
  let fen = 0;
  if (fracPart.length > 2) {
    const roundUp = fracPart[2] >= '5';
    fracPart = fracPart.slice(0, 2);
    if (roundUp) fen = 1;
  }
  fen += fracPart ? Number(fracPart.padEnd(2, '0')) : 0; // 一位小数是「角」：5 → 50 分
  if (fen === 100) {
    fen = 0;
    intPart = incDecimal(intPart);
  }
  if (intPart.length > MAX_INT_DIGITS) {
    throw new Error(`金额超出支持范围（最多到千亿级 999,999,999,999.99）：${JSON.stringify(input)}`);
  }
  return { intDigits: intPart, fen };
}

// 去小数末尾零（"16.5000000000000000" → "16.5"）。
function trimDecimalZeros(s: string): string {
  if (!s.includes('.')) return s;
  s = s.replace(/0+$/, '');
  return s.endsWith('.') ? s.slice(0, -1) : s;
}

// 十进制数字串 +1（"999" → "1000"）。
function incDecimal(digits: string): string {
  const arr = digits.split('');
  let i = arr.length - 1;
  while (i >= 0 && arr[i] === '9') {
    arr[i] = '0';
    i--;
  }
  if (i < 0) arr.unshift('1');
  else arr[i] = String(Number(arr[i]) + 1);
  return arr.join('');
}

// ---- 整数部分转大写（无前导零数字串，'0' → '零'） ----
//
// 逐位读法 + 待补零标记：遇到非零位时若此前有未读的零（且已读出内容）则补一个「零」。
// 这天然覆盖「中间连续零只读一个零」「万/亿级间零衔接」「末尾零不读」三条规范：
//   - 中间连续零：pendingZero 只置位一次，输出时只补一个零；
//   - 万/亿级间零：如 100200300 → 壹亿零贰拾万零叁佰（亿位与十万位间、十万位与百位间各有零）；
//   - 末尾零：数字串读完后 pendingZero 不输出，如 100000 → 壹拾万。
// 组级单位（万/亿）只在组尾位（pos%4==0）出现一次；若组尾位本身是 0 但组内有非零位，
// 仍要补该级单位（如 100200300 的万位是 0 → 贰拾万）。
function intToWords(digits: string): string {
  if (digits === '0') return '零';
  const len = digits.length;
  const groupCount = Math.ceil(len / 4);
  const groupNonZero = new Array<boolean>(groupCount).fill(false);
  for (let i = 0; i < len; i++) {
    if (digits.charCodeAt(i) !== 48) groupNonZero[Math.floor((len - 1 - i) / 4)] = true;
  }
  let res = '';
  let pendingZero = false;
  for (let i = 0; i < len; i++) {
    const pos = len - 1 - i; // 个位 = 0
    const d = digits.charCodeAt(i) - 48;
    if (d === 0) {
      pendingZero = true;
      if (pos % 4 === 0 && groupNonZero[pos / 4]) res += LEVEL_UNITS[pos / 4];
    } else {
      if (pendingZero && res) res += '零';
      res += DIGITS[d];
      if (pos % 4 === 0) res += LEVEL_UNITS[pos / 4];
      else res += GROUP_UNITS[pos % 4];
      pendingZero = false;
    }
  }
  return res;
}

// ---- 组装：整数大写 + 元角分 ----

function composeWords(intWords: string, fen: number): string {
  const jiao = Math.floor(fen / 10);
  const f = fen % 10;
  if (jiao === 0 && f === 0) return `${intWords}元整`;
  let s = `${intWords}元`;
  if (jiao > 0) s += `${DIGITS[jiao]}角`;
  if (f > 0) {
    if (jiao === 0) s += '零'; // 角位为 0 而分位非 0：元后补「零」
    s += `${DIGITS[f]}分`;
  }
  return s;
}
