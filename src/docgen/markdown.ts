// docgen/markdown.ts — 轻量 markdown 解析器（零依赖，纯 ESM）。
//
// 覆盖 docgen 生成链路所需子集：
//   ATX 标题 #..######（带行内解析）
//   段落（连续非空行合并为一段）
//   GFM 表格（表头行 + 分隔行；单元格仅纯文本）
//   无序列表 - / * 与有序列表 1.（每项一个块，保留原编号）
//   代码块 ```（fenced，内容按行保留）
//   行内 **bold** / *italic* / `code`
//
// 输出块数组（type 判别）：heading / paragraph / table / list-item / code。
// 本模块不依赖 docx-core —— 行内格式只表达为最小 { text, props } 形态，
// 由 generate.ts 合并进规则集组件样式。

export interface InlineRun {
  text: string;
  /** 行内标记（浅层，不做嵌套）：粗体 / 斜体 / 代码等宽字体。 */
  props?: { bold?: boolean; italic?: boolean; ascii?: string; hAnsi?: string };
}

export type MdBlock =
  | { type: 'heading'; level: number; text: string; runs: InlineRun[] }
  | { type: 'paragraph'; runs: InlineRun[] }
  | { type: 'table'; header: string[]; rows: string[][] }
  | { type: 'list-item'; ordered: boolean; index: number; text: string; runs: InlineRun[] }
  | { type: 'code'; text: string };

// 行内标记匹配：**粗体** / `代码` / *斜体*（同一位置按此顺序取首个命中）。
const INLINE_RE = /(\*\*[^*]+?\*\*|`[^`\n]+?`|\*[^*\n]+?\*)/;

/** 行内解析：**bold** / *italic* / `code` → runs（普通文本为无 props 的 run）。 */
export function parseInline(text: string): InlineRun[] {
  const runs: InlineRun[] = [];
  let rest = text;
  for (;;) {
    const m = INLINE_RE.exec(rest);
    if (!m) {
      if (rest) runs.push({ text: rest });
      break;
    }
    const before = rest.slice(0, m.index);
    if (before) runs.push({ text: before });
    const tok = m[1];
    if (tok.startsWith('**')) {
      runs.push({ text: tok.slice(2, -2), props: { bold: true } });
    } else if (tok.startsWith('`')) {
      runs.push({ text: tok.slice(1, -1), props: { ascii: 'Courier New', hAnsi: 'Courier New' } });
    } else {
      runs.push({ text: tok.slice(1, -1), props: { italic: true } });
    }
    rest = rest.slice(m.index + tok.length);
  }
  return runs;
}

const ATX_RE = /^(#{1,6})\s+(.*)$/;
const CODE_FENCE_RE = /^```/;
const UNORDERED_RE = /^[-*]\s+(.*)$/;
const ORDERED_RE = /^(\d+)\.\s+(.*)$/;
const TABLE_ROW_RE = /^\s*\|.*\|\s*$/;

// 表格分隔行：每个单元格是 - / : 组合（:---: 等），至少一个 -。
function isTableSeparator(line: string): boolean {
  const cells = splitTableRow(line);
  return cells.length > 0 && cells.every((c) => /^:?-+:?$/.test(c));
}

// 行是否可作表格行（以 | 开头且含 |；分隔行本身也匹配，由调用方排除）。
function isTableRow(line: string): boolean {
  return TABLE_ROW_RE.test(line) && splitTableRow(line).length > 0;
}

// 拆表格行：去首尾 | 后按 | 切分并 trim 每格。
function splitTableRow(line: string): string[] {
  const trimmed = line.trim();
  const inner = trimmed.replace(/^\|/, '').replace(/\|$/, '');
  return inner.split('|').map((c) => c.trim());
}

/**
 * 解析 markdown 文本 → 块数组（行序一致；空行不产出块）。
 * 表格：仅当表头行下一行为分隔行时识别，后续连续表格行作为数据行。
 */
export function parseMarkdown(markdown: string): MdBlock[] {
  const blocks: MdBlock[] = [];
  const lines = markdown.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');

  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    const trimmed = line.trim();
    if (!trimmed) {
      i++;
      continue;
    }

    // ATX 标题
    const h = ATX_RE.exec(line);
    if (h) {
      const text = h[2].trim();
      blocks.push({ type: 'heading', level: h[1].length, text, runs: parseInline(text) });
      i++;
      continue;
    }

    // 代码块（fenced ```；无闭合围栏时取到文末）
    if (CODE_FENCE_RE.test(trimmed)) {
      const codeLines: string[] = [];
      i++;
      while (i < lines.length && !CODE_FENCE_RE.test(lines[i].trim())) {
        codeLines.push(lines[i]);
        i++;
      }
      i++; // 跳过闭合围栏（可能已越界）
      blocks.push({ type: 'code', text: codeLines.join('\n') });
      continue;
    }

    // GFM 表格：表头行 + 下一行分隔行
    if (isTableRow(line) && i + 1 < lines.length && isTableSeparator(lines[i + 1])) {
      const header = splitTableRow(line);
      const rows: string[][] = [];
      let j = i + 2;
      while (j < lines.length && isTableRow(lines[j]) && !isTableSeparator(lines[j])) {
        rows.push(splitTableRow(lines[j]));
        j++;
      }
      blocks.push({ type: 'table', header, rows });
      i = j;
      continue;
    }

    // 列表项（每项一个块）
    const um = UNORDERED_RE.exec(line);
    if (um) {
      const text = um[1];
      blocks.push({ type: 'list-item', ordered: false, index: 1, text, runs: parseInline(text) });
      i++;
      continue;
    }
    const om = ORDERED_RE.exec(line);
    if (om) {
      const text = om[2];
      blocks.push({ type: 'list-item', ordered: true, index: Number(om[1]), text, runs: parseInline(text) });
      i++;
      continue;
    }

    // 段落：连续非空行合并（软换行按空格并入），遇标题/列表/代码/表格停止
    const paraLines: string[] = [line];
    i++;
    while (i < lines.length) {
      const t = lines[i].trim();
      if (!t) break;
      if (ATX_RE.test(lines[i]) || CODE_FENCE_RE.test(t)) break;
      if (UNORDERED_RE.test(lines[i]) || ORDERED_RE.test(lines[i])) break;
      if (isTableRow(lines[i]) && i + 1 < lines.length && isTableSeparator(lines[i + 1])) break;
      paraLines.push(lines[i]);
      i++;
    }
    const text = paraLines.join(' ');
    blocks.push({ type: 'paragraph', runs: parseInline(text) });
  }

  return blocks;
}
