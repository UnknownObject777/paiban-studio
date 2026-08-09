// docx-core/applyEdits.ts — 全系统唯一编辑 seam（spec 测试策略核心）。
//
//   applyEdits(docxBuffer, commands) → { buffer, result }
//
// commands 协议（借鉴 OfficeCLI batch，D6）：
//   { command: "set",    path, props }                       段落/run/节属性（按路径目标类型分派）
//   { command: "set",    match: { text: "正则" }, props }    全文段落批量匹配设置（"正文统一…"场景）
//   { command: "add",    parent, node: { kind:"paragraph", text, props, runs }, position }
//   { command: "remove", path }
//   { command: "move",   path, parent, position }
//   { command: "findReplace", find, replace, caseSensitive?, maxCount? }
//   { command: "normalize", ruleset: { rules: [...] } }      规则集驱动的全文重排（模板层供给）
//   { command: "numbering", action: "define"|"attach"|"clear", ... }
//   { command: "pageNumber", action: "footer", align?, evenAlign?, sectionIndex? }  页脚页码字段（evenAlign 时奇偶页分脚）
//
// 约定：
//   - 全部命令按数组顺序应用；单条失败不中断（结构化错误收集，含自愈建议，D6）。
//   - 产出前**生成后自检**：重新打开产物（重解析全部件）+ document.xml 结构冒烟。
//   - result.applied 逐条记录实际生效摘要，供对话层展示"AI 改了什么"。

import { openDocx, toBuffer, markDirty, getXmlTree } from './docx.js';
import { resolvePath, walkParagraphs, getBodyNode, PathError } from './model.js';
import { isElement, tagOf, childrenOf, findChildren, textOf } from './ooxml.js';
import {
  setParagraphProps, setRunProps, setParagraphRunProps,
  setAllSectionsProps, ensurePageNumberFooter,
  findReplace, addNode, removeNode, moveNode,
} from './primitives.js';
import { defineNumbering, setParagraphNumbering, clearParagraphNumbering } from './numbering.js';
import type { XmlNode } from './xml.js';
import type { Docx } from './docx.js';
import type { ParagraphProps, RunProps, AddPosition } from './primitives.js';

/** 编辑命令（命令对象，按 command 分派；其余字段按协议取值）。 */
export interface EditCommand {
  command: string;
  [key: string]: any;
}

export class CommandError extends Error {
  code: string;
  suggestion?: string | null;
  constructor(message: string, suggestion?: string | null, code = 'COMMAND_FAILED') {
    super(message);
    this.name = 'CommandError';
    this.code = code;
    this.suggestion = suggestion;
  }
}

// ---- set 命令的目标类型分派 ----

function applySet(docx: Docx, cmd: EditCommand): Record<string, any> {
  const { props = {} } = cmd;

  // 全文匹配批量设置：{ command:"set", match:{ text:"^…$", style?:"正文" }, props }
  if (cmd.match) {
    const re = new RegExp(cmd.match.text || '.*');
    const hits: string[] = [];
    walkParagraphs(docx, (p, path) => {
      const text = paragraphPlainText(p);
      if (!re.test(text)) return;
      if (cmd.scope === 'runs' || props.run) setParagraphRunProps(p, props.run || props);
      else setParagraphProps(p, props.paragraph || props);
      hits.push(path);
    });
    if (hits.length) markDirty(docx, 'word/document.xml');
    return { matched: hits.length, paths: hits.slice(0, 10), truncated: hits.length > 10 };
  }

  if (!cmd.path) throw new CommandError('set 命令缺少 path 或 match', '例: { command:"set", path:"/body/p[1]", props:{ align:"center" } }');
  const { node } = resolvePath(docx, cmd.path);
  const tag = tagOf(node);

  if (tag === 'w:p') {
    const applied: string[] = [];
    if (props.run) {
      const r = setParagraphRunProps(node, props.run);
      applied.push(`run×${r.runCount}`);
    }
    if (props.numbering !== undefined) {
      if (props.numbering === null || props.numbering.clear) {
        clearParagraphNumbering(node);
        applied.push('numbering cleared');
      } else {
        setParagraphNumbering(node, props.numbering.numId, props.numbering.ilvl || 0);
        applied.push(`numbering numId=${props.numbering.numId} ilvl=${props.numbering.ilvl || 0}`);
      }
    }
    const paraProps: ParagraphProps = props.paragraph || Object.fromEntries(
      Object.entries(props).filter(([k]) => !['run', 'numbering', 'paragraph'].includes(k)));
    if (Object.keys(paraProps).length) applied.push(...setParagraphProps(node, paraProps));
    markDirty(docx, 'word/document.xml');
    return { target: 'paragraph', applied };
  }

  if (tag === 'w:r') {
    const applied = setRunProps(node, props.run || props);
    markDirty(docx, 'word/document.xml');
    return { target: 'run', applied };
  }

  if (tag === 'w:sectPr') {
    const applied = setAllSectionsProps(docx, props, cmd.sectionIndex);
    return { target: 'section', applied };
  }

  throw new CommandError(
    `set 不支持的目标类型: ${tag}（路径 ${cmd.path}）`,
    '支持目标: 段落 /body/p[N]、run /body/p[N]/r[M]、节 /body/sectPr；批量用 match');
}

// 段落纯文本（供 match 过滤）
function paragraphPlainText(pNode: XmlNode): string {
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

// ---- normalize：规则集驱动全文重排 ----
// ruleset: { rules: [{ name, match?: { text?, position?, fallback?, element? }, set: { paragraph?, run? } }] }
// match.text      段落拼接文本正则
// match.position  'documentStart'（首个非空段落）| 'after:<规则名>'（该规则命中段之后的下一个非空段落）
// match.fallback  兜底（未被前面规则命中的段落）
// match.element   'table'：段落位于 w:tbl 内即命中（表格组件样式；空单元格段落也套用）
// rule.smartAlign 表格智能对齐（table 规则携带）：主循环后按表格结构追加一遍——
//   首行（表头）单元格段落套 smartAlign.header；正文行中"所有非空单元格均为数值"的列
//   套 smartAlign.numericColumn。对齐属性由模板层供给，套用走 setParagraphProps 原语。
// 每条段落按规则顺序首个命中者生效；rules 为空时报错（防止误清空文档）。

interface NormalizeMatch {
  text?: string;
  notText?: string;
  position?: string;
  fallback?: boolean;
  element?: string;
}

interface SmartAlignProps {
  header?: ParagraphProps;
  numericColumn?: ParagraphProps;
}

interface NormalizeRule {
  name: string;
  match?: NormalizeMatch;
  set?: { paragraph?: ParagraphProps; run?: RunProps };
  smartAlign?: SmartAlignProps;
  _re?: RegExp | null;
  _notRe?: RegExp | null;
}

// 数值单元格判定：整数/小数/千分位/百分数，可带正负号
const NUMERIC_CELL_RE = /^[+-]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?%?$/;

// 表格智能对齐：遍历 body 下所有表格（含嵌套），表头行居中、数值列右对齐。
// 只处理表格的直属行/单元格/段落（嵌套表格作为独立表格自行处理）。
function applySmartAlign(docx: Docx, smart: SmartAlignProps): number {
  const tables: XmlNode[] = [];
  const collect = (node: XmlNode): void => {
    for (const c of childrenOf(node)) {
      if (isElement(c, 'w:tbl')) { tables.push(c); collect(c); } // 嵌套表格也收集
      else collect(c);
    }
  };
  collect(getBodyNode(docx));

  let count = 0;
  const setOnCellParas = (tc: XmlNode, props: ParagraphProps): void => {
    for (const p of findChildren(tc, 'w:p')) {
      setParagraphProps(p, props);
      count++;
    }
  };
  for (const tbl of tables) {
    const rows = findChildren(tbl, 'w:tr').map((tr) => findChildren(tr, 'w:tc'));
    if (!rows.length) continue;
    // 表头行（第一行）居中
    if (smart.header) {
      for (const tc of rows[0]) setOnCellParas(tc, smart.header);
    }
    // 数值列右对齐：正文行里该列所有非空单元格均为数值才认定
    if (smart.numericColumn && rows.length > 1) {
      const bodyRows = rows.slice(1);
      const colCount = Math.max(...bodyRows.map((r) => r.length));
      for (let col = 0; col < colCount; col++) {
        const texts = bodyRows
          .map((r) => (r[col] ? textOf(r[col]).trim() : ''))
          .filter((t) => t.length > 0);
        if (!texts.length || !texts.every((t) => NUMERIC_CELL_RE.test(t))) continue;
        for (const r of bodyRows) if (r[col]) setOnCellParas(r[col], smart.numericColumn);
      }
    }
  }
  return count;
}

function applyNormalize(docx: Docx, cmd: EditCommand): { normalized: Record<string, number> } {
  const rules = cmd.ruleset?.rules as NormalizeRule[] | undefined;
  if (!Array.isArray(rules) || !rules.length) {
    throw new CommandError('normalize 需要非空 ruleset.rules', '规则集由模板层供给；例: rules:[{ name:"body", match:{}, set:{...} }]');
  }
  const compiled: NormalizeRule[] = rules.map((r) => ({
    ...r,
    _re: r.match?.text ? new RegExp(r.match.text) : null,
    _notRe: r.match?.notText ? new RegExp(r.match.notText) : null,
  }));
  const byName = new Map(compiled.map((r) => [r.name, r]));
  const stats: Record<string, number> = {};
  // 段落快照（含文本与是否在表格内），保证 position/element 谓词按文档顺序求值
  const paras: Array<{ node: XmlNode; text: string; inTable: boolean }> = [];
  walkParagraphs(docx, (p, path) => paras.push({ node: p, text: paragraphPlainText(p), inTable: path.includes('/tbl[') }));

  // position 谓词预解析
  const documentStartIdx = paras.findIndex((x) => x.text.trim().length > 0);
  const matchedIdxByRule = new Map<string, Set<number>>(); // ruleName -> Set(idx)

  paras.forEach(({ node, text, inTable }, idx) => {
    const nonEmpty = text.trim().length > 0;
    for (const rule of compiled) {
      if (rule._notRe && nonEmpty && rule._notRe.test(text)) continue; // 负向守卫
      let hit = false;
      if (rule._re) {
        hit = nonEmpty && rule._re.test(text);
      } else if (rule.match?.element === 'table') {
        hit = inTable; // 表格内段落（含空单元格），靠规则顺序先于文本类规则认领
      } else if (rule.match?.position === 'documentStart') {
        hit = idx === documentStartIdx;
      } else if (rule.match?.position?.startsWith('after:')) {
        const prevName = rule.match.position.slice(6);
        const prevSet = matchedIdxByRule.get(prevName);
        // 前一规则命中的最后一段之后的第一个非空段落
        if (prevSet && prevSet.size) {
          const lastPrev = Math.max(...prevSet);
          hit = nonEmpty && idx > lastPrev &&
            !paras.slice(lastPrev + 1, idx).some((x) => x.text.trim().length > 0);
        }
      } else if (rule.match?.fallback) {
        hit = nonEmpty; // 前面的规则都没命中才走到这里（首个命中者生效）
      } else {
        hit = nonEmpty && !rule._re && !rule.match; // 无 match = catch-all
      }
      if (!hit) continue;
      if (rule.set?.paragraph) setParagraphProps(node, rule.set.paragraph);
      if (rule.set?.run) setParagraphRunProps(node, rule.set.run);
      stats[rule.name] = (stats[rule.name] || 0) + 1;
      if (!matchedIdxByRule.has(rule.name)) matchedIdxByRule.set(rule.name, new Set());
      matchedIdxByRule.get(rule.name)!.add(idx);
      break; // 首个命中者生效
    }
  });
  // 表格智能对齐追加 pass（结构感知，主循环只按段落认领、无法表达行列语义）
  for (const rule of compiled) {
    if (!rule.smartAlign) continue;
    const n = applySmartAlign(docx, rule.smartAlign);
    if (n) stats[`${rule.name}.smartAlign`] = n;
  }
  markDirty(docx, 'word/document.xml');
  return { normalized: stats };
}

// ---- 命令分派表 ----

interface Handler {
  (docx: Docx, cmd: EditCommand): Record<string, any>;
}

const HANDLERS: Record<string, Handler> = {
  set: applySet,
  add: (docx, cmd) => {
    if (!cmd.parent || !cmd.node) throw new CommandError('add 需要 parent 与 node', 'node: { kind:"paragraph", text, props, runs }');
    return addNode(docx, cmd.parent, cmd.node, cmd.position as AddPosition);
  },
  remove: (docx, cmd) => {
    if (!cmd.path) throw new CommandError('remove 需要 path');
    return removeNode(docx, cmd.path);
  },
  move: (docx, cmd) => {
    if (!cmd.path || !cmd.parent) throw new CommandError('move 需要 path 与 parent');
    return moveNode(docx, cmd.path, cmd.parent, cmd.position as AddPosition);
  },
  findReplace: (docx, cmd) => findReplace(docx, cmd.find, cmd.replace ?? '', { caseSensitive: cmd.caseSensitive, maxCount: cmd.maxCount }),
  normalize: applyNormalize,
  numbering: (docx, cmd) => {
    if (cmd.action === 'define') {
      if (!Array.isArray(cmd.levels) || !cmd.levels.length) throw new CommandError('numbering define 需要 levels');
      return { numId: defineNumbering(docx, cmd.levels) };
    }
    if (cmd.action === 'attach') {
      const { node } = resolvePath(docx, cmd.path);
      if (!isElement(node, 'w:p')) throw new CommandError('numbering attach 目标必须是段落');
      const r = setParagraphNumbering(node, cmd.numId, cmd.ilvl || 0);
      markDirty(docx, 'word/document.xml');
      return r;
    }
    if (cmd.action === 'clear') {
      const { node } = resolvePath(docx, cmd.path);
      const cleared = clearParagraphNumbering(node);
      markDirty(docx, 'word/document.xml');
      return { cleared };
    }
    throw new CommandError(`未知 numbering action: ${cmd.action}`, '支持 define / attach / clear');
  },
  pageNumber: (docx, cmd) => ensurePageNumberFooter(docx, { align: cmd.align, evenAlign: cmd.evenAlign, sectionIndex: cmd.sectionIndex }),
};

// ---- 生成后自检 ----

function selfCheck(buffer: Buffer): { ok: true; parts: number } {
  const docx = openDocx(buffer); // 重解析全部件（解析失败即抛）
  const tree = getXmlTree(docx, 'word/document.xml');
  const docEl = tree?.find((n) => isElement(n, 'w:document'));
  if (!docEl) throw new CommandError('自检失败：document.xml 缺少 w:document 根', null, 'SELF_CHECK_FAILED');
  const body = (docEl['w:document'] as XmlNode[] | undefined || []).find((n) => isElement(n, 'w:body'));
  if (!body) throw new CommandError('自检失败：document.xml 缺少 w:body', null, 'SELF_CHECK_FAILED');
  return { ok: true, parts: docx.parts.size };
}

export interface ApplyDetail {
  index: number;
  command: string;
  path: string | null;
  detail: Record<string, any>;
}

export interface ApplyError {
  index: number;
  command: string | null;
  path: string | null;
  code: string;
  message: string;
  suggestion: string | null | undefined;
}

export interface ApplyResult {
  buffer: Buffer;
  result: {
    applied: ApplyDetail[];
    errors: ApplyError[];
    selfCheck: { ok: true; parts: number };
  };
}

/**
 * 唯一编辑 seam。
 * @param docxBuffer 原文档
 * @param commands 命令数组（空数组 = 纯 round-trip）
 * @returns 编辑产物 + 逐条 applied/errors + 生成后自检结果
 */
export function applyEdits(docxBuffer: Buffer | ArrayBuffer | Uint8Array, commands: EditCommand[] = []): ApplyResult {
  if (!Array.isArray(commands)) throw new CommandError('commands 必须是数组');
  const docx = openDocx(docxBuffer);
  const applied: ApplyDetail[] = [];
  const errors: ApplyError[] = [];

  commands.forEach((cmd, i) => {
    const handler = HANDLERS[cmd?.command];
    try {
      if (!handler) {
        throw new CommandError(`未知命令: ${cmd?.command}`, `支持: ${Object.keys(HANDLERS).join(' / ')}`, 'UNKNOWN_COMMAND');
      }
      const detail = handler(docx, cmd);
      applied.push({ index: i, command: cmd.command, path: cmd.path || cmd.parent || null, detail });
    } catch (err) {
      errors.push({
        index: i, command: cmd?.command, path: cmd?.path || null,
        code: (err as CommandError).code || 'COMMAND_FAILED', message: (err as Error).message, suggestion: (err as CommandError).suggestion || null,
      });
    }
  });

  const buffer = toBuffer(docx);
  const check = selfCheck(buffer);
  return { buffer, result: { applied, errors, selfCheck: check } };
}
