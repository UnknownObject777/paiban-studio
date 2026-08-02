// 排版组件清单常量（兜底校验两文件一致性）
// recognizers.json 与 styles.json 的 components 键集必须与此完全一致；
// styles.json 另含 PAGE_ID 节。
//
// 本文件同时被 Node（校验器、测试）与浏览器（预览页 ES module）直接 import，
// 不得引入任何 Node 内置模块或依赖。

export const PAGE_ID = 'page';

export const COMPONENT_KINDS = Object.freeze({
  PARAGRAPH: 'paragraph',
  TABLE: 'table',
});

export interface ComponentDef {
  id: string;
  name: string;
  kind: string;
  sample: string | null;
}

// MVP 最小组件集（docs/mvp-spec.md R2 / issue #1 已拍板）
export const COMPONENTS: readonly ComponentDef[] = Object.freeze([
  {
    id: 'title',
    name: '题目',
    kind: COMPONENT_KINDS.PARAGRAPH,
    sample: '关于加强公文规范化管理工作的通知',
  },
  {
    id: 'subtitle',
    name: '副标题',
    kind: COMPONENT_KINDS.PARAGRAPH,
    sample: '（2026年修订版）',
  },
  {
    id: 'heading1',
    name: '一级标题',
    kind: COMPONENT_KINDS.PARAGRAPH,
    sample: '一、总体要求',
  },
  {
    id: 'heading2',
    name: '二级标题',
    kind: COMPONENT_KINDS.PARAGRAPH,
    sample: '（一）指导思想',
  },
  {
    id: 'heading3',
    name: '三级标题',
    kind: COMPONENT_KINDS.PARAGRAPH,
    sample: '1. 工作目标',
  },
  {
    id: 'heading4',
    name: '四级标题',
    kind: COMPONENT_KINDS.PARAGRAPH,
    sample: '（1）具体指标',
  },
  {
    id: 'body',
    name: '正文',
    kind: COMPONENT_KINDS.PARAGRAPH,
    sample:
      '各科室、直属各单位：为进一步规范机关公文格式，提高公文处理质量和效率，根据有关规定，现将公文规范化管理有关事项通知如下。请各单位结合实际，认真贯彻执行。',
  },
  {
    id: 'caption',
    name: '图表标题',
    kind: COMPONENT_KINDS.PARAGRAPH,
    sample: '表1 公文字号对照表',
  },
  {
    id: 'table',
    name: '表格',
    kind: COMPONENT_KINDS.TABLE,
    sample: null,
  },
  {
    id: 'attachment',
    name: '附件标识',
    kind: COMPONENT_KINDS.PARAGRAPH,
    sample: '附件：公文格式要素对照表',
  },
]);

export const COMPONENT_IDS: readonly string[] = Object.freeze(COMPONENTS.map((c) => c.id));

export function getComponent(id: string): ComponentDef | null {
  return COMPONENTS.find((c) => c.id === id) ?? null;
}
