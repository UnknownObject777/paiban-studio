// generate.ts — issue #15 实验报告 fixtures 生成脚本（可复现）。
//
// 产物（写入本目录，与脚本同源提交，保证可复现）：
//   template.docx    —— 规范排版的实验报告模板：命名段落样式（报告标题 / 一级·二级·三级标题 /
//                        正文 / 题注）、统一字体字号行距、首行缩进 2 字符、题注居中、
//                        A4 规范页边距（参照 templates/rulesets/gongwen-default 风格）、页脚页码域、
//                        带边框数据表格、单摆示意图。
//   messy-draft.docx —— 与 template 完全相同的文字内容与结构（逐段文本一致），但格式混乱：
//                        标题全部为普通正文样式（无 pStyle、无大纲级别）、字体字号五花八门、
//                        行距忽大忽小、题注不居中无样式、页边距随意、插入多余空行。
//   pendulum.png     —— 单摆示意图（手写 PNG 编码，确定性生成，字节稳定）。
//
// 运行方式（必须先用项目构建链编译，保证走项目 docx-core 解析链路）：
//   npm run build && node dist/test/fixtures/lab-report/generate.js
//
// 所有 XML 部件都经过项目 docx-core 的 parseXml → buildXml 归一化（并断言幂等），
// 因此产物天然满足 test/lab-report.test.ts 的 build 不动点断言。

import PizZip from 'pizzip';
import { deflateSync } from 'node:zlib';
import { writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseXml, buildXml } from '../../../src/docx-core/xml.js';
import { openDocx } from '../../../src/docx-core/docx.js';

// ---------------------------------------------------------------------------
// 公共 XML 工具
// ---------------------------------------------------------------------------

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** 经项目 docx-core 解析→序列化归一化，并断言 build∘parse 是幂等不动点。 */
function canonical(xml: string): string {
  const { tree, meta } = parseXml(xml);
  const once = buildXml(tree, meta);
  const { tree: t2, meta: m2 } = parseXml(once);
  const twice = buildXml(t2, m2);
  if (once !== twice) {
    throw new Error('XML 非幂等（build(parse(x)) !== x）:\n--- once ---\n' + once + '\n--- twice ---\n' + twice);
  }
  return once;
}

function run(text: string, rPr: string): string {
  return `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}

/** 在 xml（完整部件文本）上做一次幂等归一化。 */
function partXml(xml: string): string {
  return canonical(XML_DECL + xml);
}

// ---------------------------------------------------------------------------
// 内容模型：两份文件共享同一份内容（同样的文字内容与结构）
// ---------------------------------------------------------------------------

interface Block {
  kind: 'title' | 'h1' | 'h2' | 'h3' | 'body' | 'caption' | 'figure' | 'table';
  text?: string;
  header?: string[];
  rows?: string[][];
}

const CONTENT: Block[] = [
  { kind: 'title', text: '单摆测重力加速度实验报告' },

  { kind: 'h1', text: '一、实验目的' },
  { kind: 'body', text: '通过测定不同摆长下单摆的周期，验证单摆周期公式 T＝2π√(L/g)，并测定本地重力加速度 g 的数值，学习用图解法与最小二乘法处理实验数据。' },

  { kind: 'h1', text: '二、实验原理' },
  { kind: 'body', text: '单摆在摆角小于 5° 时可视作简谐运动，其周期 T 与摆长 L 满足 T＝2π√(L/g)。由该式可得 g＝4π²L/T²，因此只需精确测量摆长 L 与周期 T，即可求得重力加速度 g。摆长取悬点到球心的距离，周期取连续摆动 50 次的平均时间。' },

  { kind: 'h1', text: '三、实验装置与器材' },
  { kind: 'body', text: '铁架台、细线、金属小球、米尺、游标卡尺、秒表。' },
  { kind: 'figure' },
  { kind: 'caption', text: '图 1 单摆实验装置示意图' },

  { kind: 'h1', text: '四、实验步骤' },
  { kind: 'h2', text: '（一）仪器调整' },
  { kind: 'body', text: '将铁架台置于水平桌面，调节底座旋钮使摆线悬点位于支架顶端，确保摆动平面与桌面平行。' },
  { kind: 'h2', text: '（二）数据测量' },
  { kind: 'body', text: '用游标卡尺测量小球直径三次，取平均值计算半径；用米尺测量悬点到球顶的距离，加上小球半径得到摆长 L。使小球从摆角小于 5° 的位置释放，用秒表记录摆动 50 次的时间，重复三次取平均得到周期 T。' },

  { kind: 'h1', text: '五、实验数据记录' },
  { kind: 'h2', text: '（一）原始数据' },
  { kind: 'caption', text: '表 1 单摆摆长与周期测量数据' },
  {
    kind: 'table',
    header: ['序号', '摆长 L/cm', '周期 T/s', '重力加速度 g/(m·s⁻²)'],
    rows: [
      // 注：单元格为纯数字文本，经项目 docx-core 解析会被 fxp 转为 number 再序列化，
      // 尾零（如 50.0 / 1.90）会被丢掉，故测量值一律不带尾零，保证 build 幂等。
      ['1', '50.5', '1.43', '9.75'],
      ['2', '60.5', '1.56', '9.81'],
      ['3', '70.5', '1.68', '9.86'],
      ['4', '80.5', '1.79', '9.92'],
      ['5', '90.5', '1.91', '9.79'],
    ],
  },

  { kind: 'h2', text: '（二）数据处理' },
  { kind: 'h3', text: '1. 周期平均值计算' },
  { kind: 'body', text: '对每组摆长，取三次测量时间的平均值作为该组的周期，记录于表中。' },
  { kind: 'h3', text: '2. 重力加速度计算' },
  { kind: 'body', text: '按公式 g＝4π²L/T² 计算各组重力加速度，取其平均值作为最终结果。' },

  { kind: 'h1', text: '六、实验结论' },
  { kind: 'body', text: '实验测得本地重力加速度的平均值为 9.83 m/s²，与标准值 9.80 m/s² 的相对误差约为 0.3%，验证了单摆周期公式的正确性。' },
];

// ---------------------------------------------------------------------------
// 单摆示意图（手写 PNG：RGB8 无压缩方案，zlib deflate，确定性字节）
// ---------------------------------------------------------------------------

const IMG_W = 360;
const IMG_H = 260;

function crc32(buf: Buffer): number {
  let c = ~0;
  for (let i = 0; i < buf.length; i++) {
    c ^= buf[i];
    for (let k = 0; k < 8; k++) c = (c >>> 1) ^ (0xedb88320 & -(c & 1));
  }
  return (~c) >>> 0;
}

function pngChunk(type: string, data: Buffer): Buffer {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body), 0);
  return Buffer.concat([len, body, crc]);
}

function drawPendulum(): Buffer {
  const px = Buffer.alloc(IMG_W * IMG_H * 3);
  // 白底
  for (let i = 0; i < px.length; i += 3) {
    px[i] = 255;
    px[i + 1] = 255;
    px[i + 2] = 255;
  }
  const set = (x: number, y: number, r: number, g: number, b: number) => {
    if (x < 0 || y < 0 || x >= IMG_W || y >= IMG_H) return;
    const i = (y * IMG_W + x) * 3;
    px[i] = r;
    px[i + 1] = g;
    px[i + 2] = b;
  };
  const rect = (x0: number, y0: number, x1: number, y1: number, r: number, g: number, b: number) => {
    for (let y = y0; y <= y1; y++) for (let x = x0; x <= x1; x++) set(x, y, r, g, b);
  };
  const line = (x0: number, y0: number, x1: number, y1: number, r: number, g: number, b: number, thick = 1) => {
    const dx = Math.abs(x1 - x0);
    const dy = -Math.abs(y1 - y0);
    const sx = x0 < x1 ? 1 : -1;
    const sy = y0 < y1 ? 1 : -1;
    let err = dx + dy;
    let x = x0;
    let y = y0;
    for (;;) {
      for (let t = 0; t < thick; t++) for (let u = 0; u < thick; u++) set(x + t - Math.floor(thick / 2), y + u - Math.floor(thick / 2), r, g, b);
      if (x === x1 && y === y1) break;
      const e2 = 2 * err;
      if (e2 >= dy) {
        err += dy;
        x += sx;
      }
      if (e2 <= dx) {
        err += dx;
        y += sy;
      }
    }
  };
  const circle = (cx: number, cy: number, rad: number, r: number, g: number, b: number, fill = true) => {
    for (let dy = -rad; dy <= rad; dy++) {
      for (let dx = -rad; dx <= rad; dx++) {
        const d = Math.sqrt(dx * dx + dy * dy);
        if (fill ? d <= rad : Math.abs(d - rad) <= 1.2) set(cx + dx, cy + dy, r, g, b);
      }
    }
  };

  // 铁架台：底座 + 立杆 + 横梁（深灰）
  rect(50, 230, 310, 248, 90, 90, 90);
  rect(110, 40, 120, 230, 90, 90, 90);
  rect(110, 40, 320, 52, 90, 90, 90);
  // 平衡位置虚线（浅灰）
  for (let y = 60; y <= 200; y += 14) rect(236, y, 244, y + 7, 190, 190, 190);
  // 悬点
  circle(240, 52, 6, 60, 60, 60);
  // 摆线（斜向右下）+ 摆球
  line(240, 52, 300, 170, 40, 40, 40, 2);
  circle(300, 170, 22, 200, 200, 200);
  circle(300, 170, 22, 40, 40, 40, false);
  // 地面
  rect(40, 252, 320, 256, 140, 140, 140);

  // 组装 PNG：签名 + IHDR + IDAT + IEND
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(IMG_W, 0);
  ihdr.writeUInt32BE(IMG_H, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // color type: truecolor
  const raw = Buffer.alloc(IMG_H * (1 + IMG_W * 3));
  for (let y = 0; y < IMG_H; y++) {
    raw[y * (1 + IMG_W * 3)] = 0; // filter: none
    px.copy(raw, y * (1 + IMG_W * 3) + 1, y * IMG_W * 3, (y + 1) * IMG_W * 3);
  }
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  return Buffer.concat([sig, pngChunk('IHDR', ihdr), pngChunk('IDAT', deflateSync(raw)), pngChunk('IEND', Buffer.alloc(0))]);
}

// ---------------------------------------------------------------------------
// 模板渲染（规范排版：样式驱动）
// ---------------------------------------------------------------------------

const TPL_RPR = {
  body: '<w:rFonts w:ascii="Times New Roman" w:eastAsia="仿宋_GB2312" w:hAnsi="Times New Roman"/><w:sz w:val="32"/><w:szCs w:val="32"/>',
  caption: '<w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/>',
  tableHead: '<w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/>',
  tableCell: '<w:rFonts w:ascii="Times New Roman" w:eastAsia="仿宋_GB2312" w:hAnsi="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/>',
};

const TPL_COLS = [1248, 2496, 2496, 2608]; // 合计 8848 twips ≈ 15.6cm

function tplDrawing(): string {
  return (
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>' +
    '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
    '<wp:extent cx="3429000" cy="2476500"/>' +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    '<wp:docPr id="1" name="单摆示意图"/>' +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="pendulum.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3429000" cy="2476500"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
  );
}

function tplTable(b: Block): string {
  const headCells = b.header!.map((h, i) => {
    const w = TPL_COLS[i];
    return (
      '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>' +
      '<w:p><w:pPr><w:jc w:val="center"/><w:rPr>' + TPL_RPR.tableHead + '</w:rPr></w:pPr>' +
      run(h, TPL_RPR.tableHead) + '</w:p></w:tc>'
    );
  }).join('');
  const bodyRows = b.rows!.map((row) => {
    const cells = row
      .map((c, i) => {
        const w = TPL_COLS[i];
        return (
          '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>' +
          '<w:p><w:pPr><w:jc w:val="center"/><w:rPr>' + TPL_RPR.tableCell + '</w:rPr></w:pPr>' +
          run(c, TPL_RPR.tableCell) + '</w:p></w:tc>'
        );
      })
      .join('');
    return '<w:tr>' + cells + '</w:tr>';
  }).join('');
  return (
    '<w:tbl><w:tblPr>' +
    '<w:tblW w:w="8848" w:type="dxa"/><w:jc w:val="center"/>' +
    '<w:tblBorders>' +
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '</w:tblBorders>' +
    '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar>' +
    '</w:tblPr>' +
    '<w:tblGrid>' + TPL_COLS.map((c) => '<w:gridCol w:w="' + c + '"/>').join('') + '</w:tblGrid>' +
    '<w:tr>' + headCells + '</w:tr>' + bodyRows +
    '</w:tbl>'
  );
}

function renderTemplateDoc(): string {
  const STYLE: Record<string, string> = {
    title: 'ReportTitle',
    h1: 'Heading1',
    h2: 'Heading2',
    h3: 'Heading3',
    body: 'Body',
    caption: 'Caption',
  };
  const blocks = CONTENT.map((b) => {
    if (b.kind === 'figure') return tplDrawing();
    if (b.kind === 'table') return tplTable(b);
    return '<w:p><w:pPr><w:pStyle w:val="' + STYLE[b.kind] + '"/></w:pPr>' + run(b.text!, '') + '</w:p>';
  }).join('');
  return (
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<w:body>' + blocks +
    '<w:sectPr>' +
    '<w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="2098" w:right="1474" w:bottom="1984" w:left="1588" w:header="851" w:footer="1418" w:gutter="0"/>' +
    '<w:footerReference w:type="default" r:id="rIdFooter"/>' +
    '<w:cols w:space="425"/>' +
    '<w:docGrid w:linePitch="312"/>' +
    '</w:sectPr>' +
    '</w:body></w:document>'
  );
}

function renderTemplateStyles(): string {
  return (
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:docDefaults><w:rPrDefault><w:rPr>' +
    '<w:rFonts w:ascii="Times New Roman" w:eastAsia="仿宋_GB2312" w:hAnsi="Times New Roman"/>' +
    '<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:rPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    // 报告标题：方正小标宋 22pt，居中，33pt 行距
    '<w:style w:type="paragraph" w:styleId="ReportTitle"><w:name w:val="报告标题"/><w:basedOn w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="240" w:line="660" w:lineRule="exact"/><w:jc w:val="center"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="方正小标宋简体" w:hAnsi="Times New Roman"/>' +
    '<w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style>' +
    // 一级标题：黑体 16pt，28pt 行距，首行缩进 2 字符，大纲 1
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:keepNext/><w:keepLines/><w:outlineLvl w:val="0"/>' +
    '<w:spacing w:before="240" w:after="120" w:line="560" w:lineRule="exact"/>' +
    '<w:ind w:firstLineChars="200" w:firstLine="640"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman"/>' +
    '<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>' +
    // 二级标题：楷体 16pt
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:keepNext/><w:keepLines/><w:outlineLvl w:val="1"/>' +
    '<w:spacing w:before="120" w:after="60" w:line="560" w:lineRule="exact"/>' +
    '<w:ind w:firstLineChars="200" w:firstLine="640"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="楷体_GB2312" w:hAnsi="Times New Roman"/>' +
    '<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>' +
    // 三级标题：仿宋 16pt 加粗
    '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:keepNext/><w:keepLines/><w:outlineLvl w:val="2"/>' +
    '<w:spacing w:before="120" w:after="60" w:line="560" w:lineRule="exact"/>' +
    '<w:ind w:firstLineChars="200" w:firstLine="640"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="仿宋_GB2312" w:hAnsi="Times New Roman"/><w:b/>' +
    '<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>' +
    // 正文：仿宋 16pt，28pt 行距，两端对齐，首行缩进 2 字符
    '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="正文"/><w:basedOn w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:spacing w:line="560" w:lineRule="exact"/><w:jc w:val="justify"/>' +
    '<w:ind w:firstLineChars="200" w:firstLine="640"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="仿宋_GB2312" w:hAnsi="Times New Roman"/>' +
    '<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>' +
    // 题注：黑体 14pt 居中（比正文小一号）
    '<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="题注"/><w:basedOn w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:keepNext/><w:spacing w:line="560" w:lineRule="exact"/><w:jc w:val="center"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman"/>' +
    '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>' +
    '</w:styles>'
  );
}

function renderTemplateFooter(): string {
  const rpr = '<w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/>';
  return (
    '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:p><w:pPr><w:jc w:val="right"/><w:rPr>' + rpr + '</w:rPr></w:pPr>' +
    '<w:r><w:rPr>' + rpr + '</w:rPr><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:rPr>' + rpr + '</w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
    '<w:r><w:rPr>' + rpr + '</w:rPr><w:fldChar w:fldCharType="end"/></w:r>' +
    '</w:p></w:ftr>'
  );
}

// ---------------------------------------------------------------------------
// 乱排版渲染（同样内容，格式混乱）
// ---------------------------------------------------------------------------

interface Chaos {
  font: string;
  ascii: string;
  sz: string;
  spacing: string;
  indent: string;
  jc: string;
}

const CHAOS: Chaos[] = [
  { font: '宋体', ascii: 'Times New Roman', sz: '24', spacing: '', indent: '', jc: '' },
  { font: '黑体', ascii: 'Arial', sz: '28', spacing: '<w:spacing w:line="360" w:lineRule="auto"/>', indent: '<w:ind w:firstLineChars="200" w:firstLine="560"/>', jc: '' },
  { font: '楷体', ascii: '宋体', sz: '21', spacing: '<w:spacing w:line="480" w:lineRule="auto"/>', indent: '', jc: '<w:jc w:val="center"/>' },
  { font: '仿宋', ascii: 'Calibri', sz: '32', spacing: '<w:spacing w:line="440" w:lineRule="exact"/>', indent: '<w:ind w:left="720"/>', jc: '' },
  { font: '微软雅黑', ascii: 'Times New Roman', sz: '36', spacing: '<w:spacing w:line="720" w:lineRule="exact"/>', indent: '<w:ind w:left="240"/>', jc: '<w:jc w:val="right"/>' },
];

function chaosRPr(c: Chaos, szDelta = 0): string {
  const sz = String(Number(c.sz) + szDelta);
  return (
    '<w:rFonts w:ascii="' + c.ascii + '" w:eastAsia="' + c.font + '" w:hAnsi="' + c.ascii + '"/>' +
    '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/>'
  );
}

let chaosIdx = 0;
function nextChaos(): Chaos {
  return CHAOS[chaosIdx++ % CHAOS.length];
}

function messyPPr(c: Chaos, extra = ''): string {
  return '<w:pPr>' + c.spacing + c.indent + c.jc + extra + '</w:pPr>';
}

/** 正文段：乱格式 + 段内双 run 字号混排。 */
function messyBodyRuns(text: string, c: Chaos): string {
  if (text.length < 14) return run(text, chaosRPr(c));
  const cut = Math.floor(text.length * 0.55);
  const next = CHAOS[(chaosIdx + 1) % CHAOS.length];
  const delta = chaosIdx % 2 === 0 ? 4 : -4;
  return run(text.slice(0, cut), chaosRPr(c)) + run(text.slice(cut), chaosRPr(next, delta));
}

function messyDrawing(): string {
  return (
    '<w:p><w:pPr><w:spacing w:line="720" w:lineRule="exact"/></w:pPr><w:r><w:drawing>' +
    '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
    '<wp:extent cx="3429000" cy="2476500"/>' +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    '<wp:docPr id="1" name="单摆示意图"/>' +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="pendulum.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3429000" cy="2476500"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
  );
}

function messyTable(b: Block): string {
  const M_CELL = [
    { font: '楷体', ascii: 'Times New Roman', sz: '21', jc: '' },
    { font: '宋体', ascii: 'Arial', sz: '24', jc: '<w:jc w:val="right"/>' },
  ];
  const headCells = b.header!.map((h, i) => {
    const w = TPL_COLS[i];
    const rpr = '<w:rFonts w:ascii="Arial" w:eastAsia="宋体" w:hAnsi="Arial"/><w:sz w:val="24"/><w:szCs w:val="24"/>';
    return (
      '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/></w:tcPr>' +
      '<w:p><w:pPr><w:rPr>' + rpr + '</w:rPr></w:pPr>' + run(h, rpr) + '</w:p></w:tc>'
    );
  }).join('');
  let m = 0;
  const bodyRows = b.rows!.map((row) => {
    const cells = row
      .map((c, i) => {
        const w = TPL_COLS[i];
        const cell = M_CELL[m % M_CELL.length];
        m++;
        const rpr =
          '<w:rFonts w:ascii="' + cell.ascii + '" w:eastAsia="' + cell.font + '" w:hAnsi="' + cell.ascii + '"/>' +
          '<w:sz w:val="' + cell.sz + '"/><w:szCs w:val="' + cell.sz + '"/>';
        return (
          '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/></w:tcPr>' +
          '<w:p><w:pPr>' + cell.jc + '<w:rPr>' + rpr + '</w:rPr></w:pPr>' + run(c, rpr) + '</w:p></w:tc>'
        );
      })
      .join('');
    return '<w:tr>' + cells + '</w:tr>';
  }).join('');
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="8848" w:type="dxa"/>' +
    '<w:tblBorders>' +
    '<w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>' +
    '<w:left w:val="single" w:sz="8" w:space="0" w:color="000000"/>' +
    '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>' +
    '<w:right w:val="single" w:sz="8" w:space="0" w:color="000000"/>' +
    '<w:insideH w:val="single" w:sz="2" w:space="0" w:color="000000"/>' +
    '<w:insideV w:val="single" w:sz="8" w:space="0" w:color="000000"/>' +
    '</w:tblBorders></w:tblPr>' +
    '<w:tblGrid>' + TPL_COLS.map((c) => '<w:gridCol w:w="' + c + '"/>').join('') + '</w:tblGrid>' +
    '<w:tr>' + headCells + '</w:tr>' + bodyRows +
    '</w:tbl>'
  );
}

/** 多余空行：CONTENT 块下标 → 其后追加空段落数。 */
const MESSY_BLANKS: Record<number, number> = {
  0: 1, // 标题后
  3: 1, // 二、实验原理后
  7: 1, // 装置图后
  8: 1, // 图题后
  15: 1, // （一）原始数据后
  17: 2, // 表格后
};

function renderMessyDoc(): string {
  const blocks: string[] = [];
  CONTENT.forEach((b, i) => {
    if (b.kind === 'title') {
      const r1 = run('单摆测重力', '<w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="44"/><w:szCs w:val="44"/>');
      const r2 = run('加速度实验报告', '<w:rFonts w:ascii="Arial" w:eastAsia="宋体" w:hAnsi="Arial"/><w:sz w:val="28"/><w:szCs w:val="28"/>');
      blocks.push('<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="480" w:lineRule="auto"/></w:pPr>' + r1 + r2 + '</w:p>');
    } else if (b.kind === 'h1') {
      const c = chaosIdx % 2 === 0 ? CHAOS[0] : CHAOS[1];
      chaosIdx++;
      const bold = chaosIdx % 2 === 0 ? '<w:b/>' : '';
      const rpr = chaosRPr(c) + bold;
      blocks.push('<w:p>' + messyPPr(c) + run(b.text!, rpr) + '</w:p>');
    } else if (b.kind === 'h2') {
      const c = chaosIdx % 2 === 0 ? CHAOS[1] : CHAOS[2];
      chaosIdx++;
      blocks.push('<w:p>' + messyPPr(c) + run(b.text!, chaosRPr(c)) + '</w:p>');
    } else if (b.kind === 'h3') {
      const c = CHAOS[2];
      blocks.push('<w:p>' + messyPPr(c) + run(b.text!, chaosRPr(c) + '<w:i/>') + '</w:p>');
    } else if (b.kind === 'body') {
      const c = nextChaos();
      blocks.push('<w:p>' + messyPPr(c) + messyBodyRuns(b.text!, c) + '</w:p>');
    } else if (b.kind === 'caption') {
      // 题注无样式：宋体五号，左对齐，与正文同形
      const c = CHAOS[2];
      blocks.push('<w:p>' + messyPPr(c) + run(b.text!, chaosRPr(c)) + '</w:p>');
    } else if (b.kind === 'figure') {
      blocks.push(messyDrawing());
    } else if (b.kind === 'table') {
      blocks.push(messyTable(b));
    }
    const blanks = MESSY_BLANKS[i] ?? 0;
    for (let k = 0; k < blanks; k++) blocks.push('<w:p/>');
  });
  return (
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<w:body>' + blocks.join('') +
    // 页边距随意：上 1.8cm 右 1.2cm 下 2.5cm 左 3.2cm（不对称），无页脚页码
    '<w:sectPr>' +
    '<w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1020" w:right="680" w:bottom="1417" w:left="1815" w:header="708" w:footer="992" w:gutter="0"/>' +
    '<w:cols w:space="425"/>' +
    '</w:sectPr>' +
    '</w:body></w:document>'
  );
}

function renderMessyStyles(): string {
  return (
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:docDefaults><w:rPrDefault><w:rPr>' +
    '<w:rFonts w:ascii="Calibri" w:eastAsia="宋体" w:hAnsi="Calibri"/>' +
    '<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '</w:styles>'
  );
}

// ---------------------------------------------------------------------------
// 打包与写出
// ---------------------------------------------------------------------------

const NS_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_OFFDOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function contentTypes(withFooter: boolean): string {
  let s =
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>';
  if (withFooter) {
    s += '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>';
  }
  return s + '</Types>';
}

function rootRels(): string {
  return (
    '<Relationships xmlns="' + NS_RELS + '">' +
    '<Relationship Id="rId1" Type="' + NS_OFFDOC + '/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>'
  );
}

function docRels(withFooter: boolean): string {
  let s =
    '<Relationships xmlns="' + NS_RELS + '">' +
    '<Relationship Id="rIdStyles" Type="' + NS_OFFDOC + '/styles" Target="styles.xml"/>' +
    '<Relationship Id="rIdSettings" Type="' + NS_OFFDOC + '/settings" Target="settings.xml"/>';
  if (withFooter) {
    s += '<Relationship Id="rIdFooter" Type="' + NS_OFFDOC + '/footer" Target="footer1.xml"/>';
  }
  s += '<Relationship Id="rIdImage" Type="' + NS_OFFDOC + '/image" Target="media/pendulum.png"/>';
  return s + '</Relationships>';
}

function settingsXml(): string {
  return (
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:zoom w:percent="100"/><w:defaultTabStop w:val="425"/>' +
    '</w:settings>'
  );
}

function buildDocx(parts: Array<{ name: string; data: string | Buffer }>): Buffer {
  const zip = new PizZip();
  for (const p of parts) zip.file(p.name, p.data);
  return zip.generate({
    type: 'nodebuffer',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
  });
}

function main(): void {
  const outDir = fileURLToPath(new URL('../../../../test/fixtures/lab-report/', import.meta.url));
  mkdirSync(outDir, { recursive: true });

  const png = drawPendulum();
  writeFileSync(join(outDir, 'pendulum.png'), png);

  // 模板（规范排版）
  const templateParts: Array<{ name: string; data: string | Buffer }> = [
    { name: '[Content_Types].xml', data: partXml(contentTypes(true)) },
    { name: '_rels/.rels', data: partXml(rootRels()) },
    { name: 'word/document.xml', data: partXml(renderTemplateDoc()) },
    { name: 'word/_rels/document.xml.rels', data: partXml(docRels(true)) },
    { name: 'word/styles.xml', data: partXml(renderTemplateStyles()) },
    { name: 'word/settings.xml', data: partXml(settingsXml()) },
    { name: 'word/footer1.xml', data: partXml(renderTemplateFooter()) },
    { name: 'word/media/pendulum.png', data: png },
  ];

  // 乱排版原稿
  const messyParts: Array<{ name: string; data: string | Buffer }> = [
    { name: '[Content_Types].xml', data: partXml(contentTypes(false)) },
    { name: '_rels/.rels', data: partXml(rootRels()) },
    { name: 'word/document.xml', data: partXml(renderMessyDoc()) },
    { name: 'word/_rels/document.xml.rels', data: partXml(docRels(false)) },
    { name: 'word/styles.xml', data: partXml(renderMessyStyles()) },
    { name: 'word/settings.xml', data: partXml(settingsXml()) },
    { name: 'word/media/pendulum.png', data: png },
  ];

  const templateBuf = buildDocx(templateParts);
  const messyBuf = buildDocx(messyParts);
  writeFileSync(join(outDir, 'template.docx'), templateBuf);
  writeFileSync(join(outDir, 'messy-draft.docx'), messyBuf);

  // 自检：产物经项目 openDocx 链路可打开
  for (const [name, buf] of [
    ['template.docx', templateBuf],
    ['messy-draft.docx', messyBuf],
  ] as const) {
    const doc = openDocx(buf);
    if (!doc.parts.has('word/document.xml')) throw new Error(name + ' 缺 document.xml');
    console.log(`[generate] ${name}: 部件数=${doc.parts.size}, PNG=${png.length}B`);
  }
  console.log('[generate] 产物写入 ' + outDir);
}

main();
