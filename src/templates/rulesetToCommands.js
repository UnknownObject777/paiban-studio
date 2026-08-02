// templates/rulesetToCommands.js — 两文件规则集（recognizers.json + styles.json）
// 转编辑内核命令（normalize 规则 + 页面设置），模板层与编辑内核之间的翻译器。
//
// 组件顺序（首个命中者生效，特殊→兜底）：
//   title(documentStart) → subtitle(after:title) → heading1..4(regex) → caption(regex)
//   → attachment(regex) → body(fallback)
// page 节 → set /body/sectPr（margins/pageSize）+ pageNumber footer。

// styles.json 组件样式 → 内核 paragraph/run props
export function styleToKernelProps(style, { forHeading = false } = {}) {
  const paragraph = {};
  const run = {};
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

// recognizers 中取某组件的首个可用内核 match（regex → text；position/heuristic → 位置谓词或 null）
function recognizerToMatch(componentId, recognizer) {
  if (!recognizer) return null;
  if (recognizer.fallback) return { fallback: true };
  for (const entry of recognizer.match || []) {
    if (entry.type === 'regex') return { text: entry.pattern };
    if (entry.type === 'position' && entry.where === 'documentStart') return { position: 'documentStart' };
    if (entry.type === 'position' && entry.where === 'afterComponent' && entry.params?.component) {
      return { position: 'after:' + entry.params.component };
    }
    // heuristic 无可执行内核映射 → 跳过（notes 保留在规则集里供人读）
  }
  return null;
}

// 组件应用顺序：特殊规则在前，兜底在后
const RULE_ORDER = ['title', 'subtitle', 'heading1', 'heading2', 'heading3', 'heading4', 'caption', 'attachment', 'body'];

/**
 * 规则集 → 内核命令数组。
 * @returns {Array} [normalize 命令, （可选）页面设置命令, （可选）页码命令]
 */
export function rulesetToCommands(recognizers, styles) {
  // 副标题守卫：position 规则（after:title）不能吞噬形似标题的段落
  const headingPatterns = ['heading1', 'heading2', 'heading3', 'heading4']
    .map((id) => (recognizers.components?.[id]?.match || []).find((e) => e.type === 'regex')?.pattern)
    .filter(Boolean)
    .map((p) => `(?:${p.replace(/^\^/, '')})`);
  const headingGuard = headingPatterns.length ? `^(?:${headingPatterns.join('|')})` : null;

  const rules = [];
  for (const id of RULE_ORDER) {
    const style = styles.components?.[id];
    if (!style) continue;
    const match = recognizerToMatch(id, recognizers.components?.[id]);
    if (!match) continue; // 该组件无可执行识别规则 → 本轮 normalize 不处理
    if (id === 'subtitle' && headingGuard && match.position) {
      match.notText = headingGuard;
    }
    const { paragraph, run } = styleToKernelProps(style);
    const rule = { name: id, match, set: {} };
    if (Object.keys(paragraph).length) rule.set.paragraph = paragraph;
    if (Object.keys(run).length) rule.set.run = run;
    rules.push(rule);
  }

  const commands = [];
  if (rules.length) commands.push({ command: 'normalize', ruleset: { rules } });

  // 页面设置
  const page = styles.page;
  if (page) {
    const props = {};
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
      commands.push({
        command: 'pageNumber', action: 'footer',
        align: page.pageNumber.oddAlign || 'center',
      });
    }
  }
  return commands;
}
