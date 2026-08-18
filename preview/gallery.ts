// 组件画廊 + 页面线框渲染器
// 读取两文件规则集（recognizers.json + styles.json），复用 schema.ts 校验，
// 然后以真实字体/字号/行距渲染每个组件的样例文字，并绘制 A4 页面线框。

import { COMPONENTS, PAGE_ID } from '../src/ruleset/components.js';
import { validateRuleset } from '../src/ruleset/schema.js';

// 内置规则集清单（手写资产，与 templates/rulesets/ 目录一一对应；新增规则集在此登记）
const RULESET_IDS = ['gongwen-default', 'lab-report-default', 'bid-default', 'fx-form-default'] as const;

const PT_TO_PX = 96 / 72; // 1pt = 1.333px
const CM_TO_PX = 96 / 2.54;

// 纸张尺寸（cm）；paper=preserve 时预览按 A4 渲染
const PAPER_SIZES: Record<string, { widthCm: number; heightCm: number }> = {
  A4: { widthCm: 21, heightCm: 29.7 },
  A3: { widthCm: 29.7, heightCm: 42 },
  A5: { widthCm: 14.8, heightCm: 21 },
  B5: { widthCm: 17.6, heightCm: 25 },
};

const STYLE_LABELS: Record<string, string> = {
  fontEastAsia: '中文字体',
  fontAscii: '西文字体',
  sizePt: '字号(pt)',
  bold: '加粗',
  align: '对齐',
  firstLineIndentChars: '首行缩进(字符)',
  lineSpacingPt: '行距(pt)',
  lineSpacingMultiple: '行距(倍数)',
  spaceBeforePt: '段前(pt)',
  spaceAfterPt: '段后(pt)',
  pageBreakBefore: '段前分页',
  outlineLevel: '大纲级别',
  smartAlign: '智能对齐',
  headerBold: '表头加粗',
  headerFontEastAsia: '表头中文字体',
  borders: '表格边框',
  widthPct: '表格宽度(%)',
};

function ptToPx(ptValue: number): string {
  return `${(ptValue * PT_TO_PX).toFixed(2)}px`;
}

async function fetchJson(url: string): Promise<any> {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`${url} → HTTP ${res.status}`);
  return res.json();
}

function summarizeMatch(match: any): string {
  if (!match) return 'fallback（未命中任何规则的段落）';
  return match
    .map((rule: any) => {
      if (rule.type === 'regex') return `regex ${rule.pattern}`;
      if (rule.type === 'position') return `position ${rule.where}`;
      return `heuristic ${rule.kind}`;
    })
    .join(' + ');
}

function paragraphSampleStyle(style: any): Record<string, string> {
  const css: Record<string, string> = {};
  if (style.fontEastAsia || style.fontAscii) {
    const families = [style.fontEastAsia, style.fontAscii].filter(Boolean);
    css.fontFamily = families.map((f: string) => `"${f}"`).join(', ') + ', sans-serif';
  }
  if (style.sizePt) css.fontSize = ptToPx(style.sizePt);
  if (style.lineSpacingPt) css.lineHeight = ptToPx(style.lineSpacingPt);
  else if (style.lineSpacingMultiple) css.lineHeight = String(style.lineSpacingMultiple);
  if (style.bold) css.fontWeight = '700';
  if (style.italic) css.fontStyle = 'italic';
  css.textAlign = style.align ?? 'left';
  if (style.firstLineIndentChars) css.textIndent = `${style.firstLineIndentChars}em`;
  return css;
}

function applyStyle(el: HTMLElement, css: Record<string, string>): void {
  for (const [key, value] of Object.entries(css)) (el.style as any)[key] = value;
}

function metaEntries(style: any): Array<[string, unknown]> {
  return Object.entries(style).filter(([key]) => key !== 'notes');
}

function renderCard(component: any, recognizer: any, style: any): HTMLElement {
  const card = document.createElement('article');
  card.className = 'card';

  const badges: string[] = [];
  if (recognizer.fallback) badges.push('<span class="card__badge">fallback</span>');
  if (style.pageBreakBefore) badges.push('<span class="card__badge">段前分页</span>');

  const head = document.createElement('header');
  head.className = 'card__head';
  head.innerHTML = `
    <span class="card__name">${component.name}</span>
    <span class="card__id mono-label">${component.id}</span>
    ${badges.join('')}
    <span class="card__rules mono-label">${summarizeMatch(recognizer.match)}</span>
  `;
  card.appendChild(head);

  const body = document.createElement('div');
  body.className = 'card__body';

  const sampleBox = document.createElement('div');
  sampleBox.className = 'card__sample';
  if (component.kind === 'table') {
    sampleBox.appendChild(renderSampleTable(style));
  } else {
    const p = document.createElement('p');
    p.className = 'sample-line';
    p.textContent = component.sample;
    applyStyle(p, paragraphSampleStyle(style));
    sampleBox.appendChild(p);
  }
  body.appendChild(sampleBox);

  const meta = document.createElement('aside');
  meta.className = 'card__meta';
  const dl = document.createElement('dl');
  for (const [key, value] of metaEntries(style)) {
    const dt = document.createElement('dt');
    dt.textContent = STYLE_LABELS[key] ?? key;
    const dd = document.createElement('dd');
    dd.textContent = String(value);
    dl.append(dt, dd);
  }
  meta.appendChild(dl);
  const notes = style.notes ?? recognizer.notes;
  if (notes) {
    const p = document.createElement('p');
    p.className = 'card__notes';
    p.textContent = `备注：${notes}`;
    meta.appendChild(p);
  }
  body.appendChild(meta);

  card.appendChild(body);
  return card;
}

function renderSampleTable(style: any): HTMLTableElement {
  const table = document.createElement('table');
  table.className = 'sample-table';
  const rows = [
    ['要素', '字体', '字号'],
    ['题目', '方正小标宋简体', '二号'],
    ['正文', '仿宋_GB2312', '三号(16pt)'],
  ];
  rows.forEach((cells, rowIndex) => {
    const tr = document.createElement('tr');
    for (const text of cells) {
      const cell = document.createElement(rowIndex === 0 ? 'th' : 'td');
      cell.textContent = text;
      applyStyle(cell, paragraphSampleStyle(style));
      cell.style.lineHeight = '1.6';
      if (rowIndex === 0 && style.headerBold) cell.style.fontWeight = '700';
      tr.appendChild(cell);
    }
    table.appendChild(tr);
  });
  return table;
}

function renderGallery(recognizers: any, styles: any): void {
  const gallery = document.getElementById('gallery')!;
  gallery.textContent = '';
  for (const component of COMPONENTS) {
    gallery.appendChild(
      renderCard(component, recognizers.components[component.id], styles.components[component.id]),
    );
  }
}

// ---------- 页面线框 ----------

function pagePaper(styles: any): { widthCm: number; heightCm: number; effective: string } {
  const page = styles[PAGE_ID];
  const key = page.paper === 'preserve' || page.paper === 'custom' ? 'A4' : page.paper;
  return { ...PAPER_SIZES[key], effective: key };
}

function renderWireframe(styles: any, parity: 'odd' | 'even'): void {
  const page = styles[PAGE_ID];
  const { widthCm, heightCm, effective } = pagePaper(styles);

  const heightPx = 640;
  const scale = heightPx / (heightCm * CM_TO_PX);
  const widthPx = widthCm * CM_TO_PX * scale;
  const cm = (v: number): number => v * CM_TO_PX * scale; // 厘米 → 线框像素

  const wf = document.getElementById('wireframe')!;
  wf.textContent = '';
  wf.style.width = `${widthPx}px`;
  wf.style.height = `${heightPx}px`;

  const m = page.margins;
  const margins = document.createElement('div');
  margins.className = 'wf-margins';
  margins.style.top = `${cm(m.topCm)}px`;
  margins.style.bottom = `${cm(m.bottomCm)}px`;
  margins.style.left = `${cm(m.leftCm)}px`;
  margins.style.right = `${cm(m.rightCm)}px`;
  wf.appendChild(margins);

  // 版心示意
  const zone = document.createElement('div');
  zone.className = 'wf-textzone';
  zone.style.top = `${cm(m.topCm + 0.4)}px`;
  zone.style.bottom = `${cm(m.bottomCm + 0.4)}px`;
  zone.style.left = `${cm(m.leftCm + 0.3)}px`;
  zone.style.right = `${cm(m.rightCm + 0.3)}px`;
  wf.appendChild(zone);

  // 页脚距线
  const footerLine = document.createElement('div');
  footerLine.className = 'wf-footer-line';
  footerLine.style.bottom = `${cm(page.footerDistanceCm)}px`;
  wf.appendChild(footerLine);

  // 页码圆钮（真实字号随线框比例缩放）
  const pn = page.pageNumber;
  const align = parity === 'odd' ? pn.oddAlign : pn.evenAlign;
  const pageno = document.createElement('div');
  pageno.className = 'wf-pageno';
  pageno.textContent = '— 1 —';
  pageno.style.fontSize = `${(pn.sizePt * PT_TO_PX * scale).toFixed(2)}px`;
  pageno.style.fontFamily = `"${pn.fontEastAsia}", serif`;
  pageno.style.bottom = `${cm(page.footerDistanceCm) - 16}px`;
  const xByAlign: Record<string, string> = { left: '24px', right: `${widthPx - 56}px`, center: `${widthPx / 2 - 16}px` };
  pageno.style.left = xByAlign[align] ?? xByAlign.center;
  wf.appendChild(pageno);

  // 尺寸标注
  const dims = [
    { text: `上 ${m.topCm}cm`, x: 8, y: cm(m.topCm) / 2 },
    { text: `左 ${m.leftCm}cm`, x: 8, y: heightPx / 2 },
    { text: `右 ${m.rightCm}cm`, x: widthPx - 88, y: heightPx / 2 },
    { text: `下 ${m.bottomCm}cm`, x: 8, y: heightPx - cm(m.bottomCm) / 2 },
    { text: `页脚距 ${page.footerDistanceCm}cm`, x: widthPx - 130, y: heightPx - cm(page.footerDistanceCm) - 20 },
  ];
  for (const dim of dims) {
    const el = document.createElement('span');
    el.className = 'wf-dim';
    el.textContent = dim.text;
    el.style.left = `${dim.x}px`;
    el.style.top = `${dim.y}px`;
    wf.appendChild(el);
  }

  // 事实清单
  const facts: Array<[string, string]> = [
    ['纸张', page.paper === 'preserve' ? `保留原纸张（预览按 ${effective}）` : page.paper],
    ['页面尺寸', `${widthCm} × ${heightCm} cm`],
    ['页边距', `上${m.topCm} / 下${m.bottomCm} / 左${m.leftCm} / 右${m.rightCm} cm`],
    ['页脚距', `${page.footerDistanceCm} cm`],
    ['页码', pn.oddEven ? '奇偶分页' : '连续'],
    ['奇数页页码', pn.oddAlign],
    ['偶数页页码', pn.evenAlign],
    ['页码字体', `${pn.fontEastAsia} ${pn.sizePt}pt`],
  ];
  const ul = document.getElementById('page-facts')!;
  ul.textContent = '';
  for (const [label, value] of facts) {
    const li = document.createElement('li');
    const b = document.createElement('b');
    b.textContent = label;
    li.append(b, document.createTextNode(String(value)));
    ul.appendChild(li);
  }

  document.getElementById('paper-note')!.textContent =
    page.paper === 'preserve' ? `paper=preserve（force_a4=false），线框按 ${effective} 渲染` : '';
}

// ---------- 装配 ----------

function switchView(view: 'gallery' | 'page'): void {
  for (const name of ['gallery', 'page'] as const) {
    document.getElementById(`view-${name}`)!.classList.toggle('is-hidden', name !== view);
    document.querySelector(`.tab[data-view="${name}"]`)!.classList.toggle('is-active', name === view);
  }
}

let currentStyles: any = null;

async function loadRuleset(rulesetId: string): Promise<void> {
  const status = document.getElementById('status')!;
  try {
    const base = `../templates/rulesets/${rulesetId}`;
    const [recognizers, styles] = await Promise.all([
      fetchJson(`${base}/recognizers.json`),
      fetchJson(`${base}/styles.json`),
    ]);
    const errors = validateRuleset(recognizers, styles);
    if (errors.length > 0) {
      throw new Error(`规则集校验未通过：\n- ${errors.join('\n- ')}`);
    }
    currentStyles = styles;

    document.getElementById('ruleset-name')!.textContent =
      `${recognizers.ruleset} · v${recognizers.version}`;
    status.textContent = `${COMPONENTS.length} 个组件 · 两文件一致性校验通过`;
    setTimeout(() => status.classList.add('is-hidden'), 2400);

    renderGallery(recognizers, styles);
    const parity = (document.querySelector('.page-toolbar .pill.is-active') as HTMLElement | null)?.dataset.parity ?? 'odd';
    renderWireframe(styles, parity as 'odd' | 'even');

    document.querySelectorAll('#ruleset-pills .pill').forEach((pill) =>
      pill.classList.toggle('is-active', (pill as HTMLElement).dataset.ruleset === rulesetId),
    );
  } catch (err) {
    status.classList.add('status--error');
    status.textContent = `加载失败：${(err as Error).message}\n\n请先用 npm run preview 启动本地服务（fetch 不支持 file:// 直开）。`;
  }
}

async function main(): Promise<void> {
  const pills = document.getElementById('ruleset-pills')!;
  for (const id of RULESET_IDS) {
    const btn = document.createElement('button');
    btn.className = 'pill';
    btn.type = 'button';
    btn.dataset.ruleset = id;
    btn.textContent = id;
    btn.addEventListener('click', () => void loadRuleset(id));
    pills.appendChild(btn);
  }

  document.querySelectorAll('.tab').forEach((tab) =>
    tab.addEventListener('click', () => switchView((tab as HTMLElement).dataset.view as 'gallery' | 'page')),
  );
  document.querySelectorAll('.page-toolbar .pill').forEach((pill) =>
    pill.addEventListener('click', () => {
      document.querySelectorAll('.page-toolbar .pill').forEach((p) => p.classList.toggle('is-active', p === pill));
      renderWireframe(currentStyles, (pill as HTMLElement).dataset.parity as 'odd' | 'even');
    }),
  );

  await loadRuleset(RULESET_IDS[0]);
}

main();
