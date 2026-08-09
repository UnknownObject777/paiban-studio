// templates/rulesetToCommands.ts — 两文件规则集（recognizers.json + styles.json）
// 转编辑内核命令（normalize 规则 + 页面设置），模板层与编辑内核之间的翻译器。
//
// 组件顺序（首个命中者生效，特殊→兜底）：
//   table(element:table，表格内段落优先认领，防止单元格文本误中标题正则)
//   → title(documentStart) → subtitle(after:title) → heading1..4(regex) → caption(regex)
//   → attachment(regex) → body(fallback)
// page 节 → set /body/sectPr（margins/pageSize）+ pageNumber footer（oddEven 时带 evenAlign）。

import type { ParagraphProps, RunProps } from '../docx-core/primitives.js';
import type { EditCommand } from '../docx-core/applyEdits.js';

/** normalize 规则（与 applyEdits 中 applyNormalize 的结构对齐）。 */
export interface NormalizeRule {
  name: string;
  match?: { text?: string; notText?: string; position?: string; fallback?: boolean; element?: string };
  set?: { paragraph?: ParagraphProps; run?: RunProps };
  /** 表格智能对齐（smartAlign: true 时由 table 组件规则携带）：表头行 / 数值列分别套用的段落属性。 */
  smartAlign?: { header?: ParagraphProps; numericColumn?: ParagraphProps };
  _re?: RegExp | null;
  _notRe?: RegExp | null;
}

export interface RulesetStyle {
  align?: string;
  firstLineIndentChars?: number;
  lineSpacingPt?: number;
  lineSpacingMultiple?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  pageBreakBefore?: boolean;
  outlineLevel?: number;
  fontEastAsia?: string;
  fontAscii?: string;
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  [key: string]: unknown;
}

export interface KernelProps {
  paragraph: ParagraphProps;
  run: RunProps;
}

// styles.json 组件样式 → 内核 paragraph/run props
export function styleToKernelProps(style: RulesetStyle, { forHeading = false }: { forHeading?: boolean } = {}): KernelProps {
  const paragraph: ParagraphProps = {};
  const run: RunProps = {};
  if (style.align !== undefined) paragraph.align = style.align; // 内核 ALIGN_MAP 处理 justify→both
  if (style.firstLineIndentChars !== undefined) paragraph.firstLineChars = style.firstLineIndentChars * 100;
  if (style.lineSpacingPt !== undefined) paragraph.lineSpacingPt = style.lineSpacingPt;
  if (style.lineSpacingMultiple !== undefined) paragraph.lineSpacingMultiple = style.lineSpacingMultiple;
  if (style.spaceBeforePt !== undefined) paragraph.spacingBeforePt = style.spaceBeforePt;
  if (style.spaceAfterPt !== undefined) paragraph.spacingAfterPt = style.spaceAfterPt;
  if (style.pageBreakBefore !== undefined) paragraph.pageBreakBefore = style.pageBreakBefore;
  if (style.outlineLevel !== undefined) paragraph.outlineLevel = Math.max(0, style.outlineLevel - 1); // 规则集 1 起 → OOXML 0 起
  if (style.fontEastAsia !== undefined) run.eastAsia = style.fontEastAsia;
  if (style.fontAscii !== undefined) run.ascii = style.fontAscii;
  if (style.fontAscii !== undefined) run.hAnsi = style.fontAscii;
  if (style.sizePt !== undefined) run.sizePt = style.sizePt;
  if (style.bold !== undefined) run.bold = style.bold;
  if (style.italic !== undefined) run.italic = style.italic;
  return { paragraph, run };
}

interface RecognizerEntry {
  type?: string;
  pattern?: string;
  where?: string;
  params?: { component?: string };
  [key: string]: unknown;
}

interface Recognizer {
  fallback?: boolean;
  match?: RecognizerEntry[];
  [key: string]: unknown;
}

interface NormalizeMatch {
  text?: string;
  notText?: string;
  position?: string;
  fallback?: boolean;
  element?: string;
}

// recognizers 中取某组件的首个可用内核 match
// （regex → text；position → 位置谓词；heuristic isTableElement → element:table；其余 heuristic → null）
function recognizerToMatch(componentId: string, recognizer: Recognizer | undefined): NormalizeMatch | null {
  if (!recognizer) return null;
  if (recognizer.fallback) return { fallback: true };
  for (const entry of recognizer.match || []) {
    if (entry.type === 'regex') return { text: entry.pattern };
    if (entry.type === 'position' && entry.where === 'documentStart') return { position: 'documentStart' };
    if (entry.type === 'position' && entry.where === 'afterComponent' && entry.params?.component) {
      return { position: 'after:' + entry.params.component };
    }
    if (entry.type === 'heuristic' && (entry as { kind?: string }).kind === 'isTableElement') {
      return { element: 'table' }; // 表格元素自身即识别依据，由 normalize 按“段落位于 w:tbl 内”判定
    }
    // 其余 heuristic 无可执行内核映射 → 跳过（notes 保留在规则集里供人读）
  }
  return null;
}

// 组件应用顺序：table 最先（认领表格内段落），特殊规则在前，兜底在后
const RULE_ORDER: readonly string[] = ['table', 'title', 'subtitle', 'heading1', 'heading2', 'heading3', 'heading4', 'caption', 'attachment', 'body'];

/**
 * 规则集 → 内核命令数组。
 * @returns [normalize 命令, （可选）页面设置命令, （可选）页码命令]
 */
export function rulesetToCommands(
  recognizers: Record<string, any>,
  styles: Record<string, any>,
): EditCommand[] {
  // 副标题守卫：position 规则（after:title）不能吞噬形似标题的段落
  const headingPatterns = ['heading1', 'heading2', 'heading3', 'heading4']
    .map((id) => (recognizers.components?.[id]?.match || []).find((e: RecognizerEntry) => e.type === 'regex')?.pattern)
    .filter(Boolean)
    .map((p: string) => `(?:${p.replace(/^\^/, '')})`);
  const headingGuard = headingPatterns.length ? `^(?:${headingPatterns.join('|')})` : null;

  const rules: NormalizeRule[] = [];
  for (const id of RULE_ORDER) {
    const style = styles.components?.[id] as RulesetStyle | undefined;
    if (!style) continue;
    const match = recognizerToMatch(id, recognizers.components?.[id] as Recognizer | undefined);
    if (!match) continue; // 该组件无可执行识别规则 → 本轮 normalize 不处理
    if (id === 'subtitle' && headingGuard && match.position) {
      match.notText = headingGuard;
    }
    const { paragraph, run } = styleToKernelProps(style);
    const rule: NormalizeRule = { name: id, match: match as NormalizeRule['match'], set: {} };
    if (Object.keys(paragraph).length) rule.set!.paragraph = paragraph;
    if (Object.keys(run).length) rule.set!.run = run;
    // 表格智能对齐：表头行居中、数值列右对齐（内核按单元格内容分桶，走 set.paragraph 原语）
    if (id === 'table' && style.smartAlign === true) {
      rule.smartAlign = { header: { align: 'center' }, numericColumn: { align: 'right' } };
    }
    rules.push(rule);
  }

  const commands: EditCommand[] = [];
  if (rules.length) commands.push({ command: 'normalize', ruleset: { rules } });

  // 页面设置
  const page = styles.page as Record<string, any> | undefined;
  if (page) {
    const props: Record<string, any> = {};
    if (page.paper && page.paper !== 'preserve' && page.paper !== 'custom') {
      props.pageSize = page.paper.toLowerCase();
    }
    if (page.margins) {
      props.marginsCm = {};
      if (page.margins.topCm !== undefined) props.marginsCm.top = page.margins.topCm;
      if (page.margins.bottomCm !== undefined) props.marginsCm.bottom = page.margins.bottomCm;
      if (page.margins.leftCm !== undefined) props.marginsCm.left = page.margins.leftCm;
      if (page.margins.rightCm !== undefined) props.marginsCm.right = page.margins.rightCm;
    }
    if (page.footerDistanceCm !== undefined) {
      props.marginsCm = props.marginsCm || {};
      props.marginsCm.footer = page.footerDistanceCm;
    }
    if (Object.keys(props).length) {
      commands.push({ command: 'set', path: '/body/sectPr', props });
    }
    if (page.pageNumber) {
      const pn = page.pageNumber;
      const cmd: EditCommand = {
        command: 'pageNumber', action: 'footer',
        align: pn.oddAlign || 'center',
      };
      // 奇偶页不同对齐：补发偶数页页脚（内核负责 evenAndOddHeaders + even footer 部件）
      if (pn.oddEven && pn.evenAlign) cmd.evenAlign = pn.evenAlign;
      commands.push(cmd);
    }
  }
  return commands;
}
