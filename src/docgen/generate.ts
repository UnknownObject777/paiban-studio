// docgen/generate.ts — markdown + 内置规则集 → 规范排版 docx 的生成器。
//
// 链路：parseMarkdown → 块 → 组件映射（title/heading1..4/body/table）
//     → add 命令（段落类走 buildParagraph；表格手工构造 w:tbl fxp 节点）
//     → 追加规则集页面/页码命令（过滤 normalize：空白文档无需重排）
//     → 一次 applyEdits（唯一 seam，产出前自检）→ 产物 Buffer。
//
// 空白模板：PizZip 内存打包最小 OOXML 集（document/styles/content-types/rels），
// body 内含空 w:sectPr —— add 命令插在其前，页面设置/页码命令直接作用其上，
// 保证 applyEdits 自检、docx-preview 与 Word 均可打开。

import PizZip from 'pizzip';
import { applyEdits } from '../docx-core/applyEdits.js';
import { el, childrenOf } from '../docx-core/ooxml.js';
import { buildParagraph } from '../docx-core/primitives.js';
import { styleToKernelProps, rulesetToCommands } from '../templates/rulesetToCommands.js';
import { parseMarkdown } from './markdown.js';
import type { InlineRun } from './markdown.js';
import type { EditCommand } from '../docx-core/applyEdits.js';
import type { ParagraphProps, RunProps } from '../docx-core/primitives.js';
import type { XmlNode } from '../docx-core/xml.js';

// ---- 组件映射 ----

// 块层级 → 组件 id：##→heading1、###→heading2、####→heading3、#####/######→heading4。
const HEADING_MAP: Record<number, string> = {
  2: 'heading1', 3: 'heading2', 4: 'heading3', 5: 'heading4', 6: 'heading4',
};

// 代码块 run 字体：等宽 Courier New（西文；中文仍走组件 eastAsia）。
const CODE_FONT: Partial<RunProps> = { ascii: 'Courier New', hAnsi: 'Courier New' };

/** 段落类块 → add 命令（组件段落 props + 组件 run props 合并行内 props，行内优先）。 */
function addParagraphCommand(
  componentId: string,
  runs: InlineRun[],
  styles: Record<string, any>,
  prefix?: string,
  forceRun: Partial<RunProps> = {},
): EditCommand {
  const { paragraph, run } = styleToKernelProps((styles.components && styles.components[componentId]) || {});
  const runSpecs = runs.map((r) => ({
    text: r.text,
    props: { ...run, ...(r.props as Partial<RunProps> | undefined), ...forceRun },
  }));
  if (prefix && runSpecs.length) {
    runSpecs[0] = { ...runSpecs[0], text: prefix + runSpecs[0].text };
  }
  return {
    command: 'add', parent: '/body', position: 'end',
    node: { kind: 'paragraph', props: paragraph, runs: runSpecs },
  };
}

// ---- 表格（原始 fxp 节点） ----

const TBL_BORDER_SIDE = { 'w:val': 'single', 'w:sz': 4, 'w:space': 0, 'w:color': 'auto' };

/**
 * 表格块 → w:tbl fxp 节点（原始节点，直接作 add 命令的 node）。
 * tblPr：tblW 100% + 单线边框（styles.components.table.borders === false 时去边框）；
 * 表头行 run 套 headerFontEastAsia/headerBold，数据行套组件 fontEastAsia/ascii/sizePt；
 * 单元格段落 align 用 table.align（缺省 left）。
 */
function buildTableNode(table: { header: string[]; rows: string[][] }, styles: Record<string, any>): XmlNode {
  const tStyle = (styles.components && styles.components.table) || {};
  const { run } = styleToKernelProps(tStyle);
  const cellAlign: ParagraphProps['align'] = tStyle.align || 'left';

  const colCount = Math.max(table.header.length, ...table.rows.map((r) => r.length));
  const pad = (cells: string[]): string[] => {
    const out = cells.slice();
    while (out.length < colCount) out.push('');
    return out;
  };
  const header = pad(table.header);
  const rows = table.rows.map(pad);

  const tbl = el('w:tbl');
  const tblPr = el('w:tblPr');
  const tblPrKids: XmlNode[] = [el('w:tblW', { 'w:w': 5000, 'w:type': 'pct' })];
  if (tStyle.borders !== false) {
    const borders = el('w:tblBorders', {}, ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
      .map((s) => el('w:' + s, TBL_BORDER_SIDE)));
    tblPrKids.push(borders);
  }
  tblPr['w:tblPr'] = tblPrKids;
  childrenOf(tbl).push(tblPr);

  // 表格网格：列数等分（A4 内容宽约 9026 twips，页面命令随后可再改页边距）
  const colW = Math.floor(9026 / colCount);
  const grid = el('w:tblGrid');
  grid['w:tblGrid'] = Array.from({ length: colCount }, () => el('w:gridCol', { 'w:w': colW }));
  childrenOf(tbl).push(grid);

  const cellPara = (text: string, isHeader: boolean): XmlNode => {
    const runProps: RunProps = { ...run };
    if (isHeader) {
      if (tStyle.headerFontEastAsia !== undefined) runProps.eastAsia = tStyle.headerFontEastAsia;
      if (tStyle.headerBold !== undefined) runProps.bold = tStyle.headerBold;
    }
    return buildParagraph({ props: { align: cellAlign }, runs: [{ text, props: runProps }] });
  };

  const trNodes = [header, ...rows].map((cells, ri) => {
    const isHeader = ri === 0;
    return el('w:tr', {}, cells.map((cell) => {
      const tc = el('w:tc');
      tc['w:tc'] = [
        el('w:tcPr', {}, [el('w:tcW', { 'w:w': 0, 'w:type': 'auto' })]),
        cellPara(cell, isHeader),
      ];
      return tc;
    }));
  });
  for (const tr of trNodes) childrenOf(tbl).push(tr);
  return tbl;
}

// ---- 空白 docx 模板（最小 OOXML 集） ----

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function buildBlankDocx(): Buffer {
  const zip = new PizZip();
  zip.file('[Content_Types].xml', DECL +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  zip.file('_rels/.rels', DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', DECL +
    `<w:document ${W_NS} ${R_NS}><w:body><w:sectPr/></w:body></w:document>`);
  zip.file('word/_rels/document.xml.rels', DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('word/styles.xml', DECL +
    `<w:styles ${W_NS}><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`);
  return zip.generate({ type: 'nodebuffer' });
}

// ---- 入口 ----

export interface GenerateOptions {
  recognizers: Record<string, any>;
  styles: Record<string, any>;
}

/**
 * markdown + 规则集 → 规范排版 docx Buffer。
 * 全部命令一次 applyEdits（含生成后自检）；result.errors 非空时 throw 并附摘要。
 */
export function generateFromMarkdown(markdown: string, opts: GenerateOptions): Buffer {
  if (typeof markdown !== 'string' || !markdown.trim()) {
    throw new Error('generateFromMarkdown 需要非空 markdown');
  }
  const { recognizers, styles } = opts;
  const commands: EditCommand[] = [];
  let titleUsed = false;

  for (const block of parseMarkdown(markdown)) {
    switch (block.type) {
      case 'heading': {
        const component = block.level === 1
          ? (titleUsed ? 'heading1' : 'title')
          : HEADING_MAP[block.level] || 'body';
        if (component === 'title') titleUsed = true;
        commands.push(addParagraphCommand(component, block.runs, styles));
        break;
      }
      case 'paragraph':
        commands.push(addParagraphCommand('body', block.runs, styles));
        break;
      case 'list-item': {
        const prefix = block.ordered ? `${block.index}. ` : '• ';
        commands.push(addParagraphCommand('body', block.runs, styles, prefix));
        break;
      }
      case 'code': {
        for (const line of block.text.split('\n')) {
          commands.push(addParagraphCommand('body', [{ text: line }], styles, undefined, CODE_FONT));
        }
        break;
      }
      case 'table':
        commands.push({ command: 'add', parent: '/body', position: 'end', node: buildTableNode(block, styles) });
        break;
    }
  }

  // 页面 / 页码命令：空白文档生成无需 normalize 重排，过滤之
  for (const cmd of rulesetToCommands(recognizers, styles)) {
    if (cmd.command !== 'normalize') commands.push(cmd);
  }

  const { buffer, result } = applyEdits(buildBlankDocx(), commands);
  if (result.errors.length) {
    throw new Error(`生成 docx 失败：${result.errors.map((e) => e.message).join('；')}`);
  }
  return buffer;
}
