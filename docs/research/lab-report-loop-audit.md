# 链路现状盘点：模板→规则集→重排→预览→导出（issue #16）

> 日期：2026-08-09
> 方法：逐环读代码 + 实际跑测试（Node v24.16.0 / Windows）。
> 测试命令：`npm run build` 后 `node --test dist/test/**/*.test.js`（54 例：**53 通过 / 0 失败 / 1 跳过**，跳过项为需真实 LLM 凭证的 e2e-agent）。
> 注意：`npm test` 原脚本 `node --test dist/test/`（目录参数）在本机失败——目录被当作模块解析报 `MODULE_NOT_FOUND`；在纯 ASCII 路径下同样复现，与中文路径无关，疑为 Node 24/Windows 行为。改用 glob 即正常。

---

## 环 1 · 模板上传与规则集 —— 已通

- `src/templates/templateStore.ts:uploadTemplate`：buffer 入对象存储 → `extractPlaceholders`（`placeholders.ts`，`{{名称}}` 正则）→ `extractRulesetFromSample`（`rulesetFromSample.ts`）→ `validateRuleset` 两文件一致性校验 → 落盘 `recognizers.json / styles.json / meta.json`。
- 反推策略（`rulesetFromSample.ts`）：**title**（首个非空段落）、**body**（文本最长段落）、**page**（首个 sectPr 纸张/页边距）三个组件从样例实测；其余 7 个组件继承 `templates/rulesets/gongwen-default` 的对应节，保证组件键集完整、过校验。
- `src/ruleset/`：schema 校验器严格（组件键集与 `components.ts` 常量双向一致、fallback 恰好 1 个、正则可编译、样式属性白名单、page 节字段齐全）；`load.ts` 加载即校验。内置公文默认规则集 `gongwen-default`（recognizers + styles 两文件）存在且通过校验。
- 测试佐证：`ruleset.test.ts` 17 例、`templates.test.ts`「规则集反推」「占位符提取」「实例化」全过。

**缺口**：

- 反推只实测 title/body/page（`rulesetFromSample.ts` 注释注明 MVP 有意为之），heading1..4/caption/table/attachment 的 recognizers **原样继承公文正则**（`一、`、`（一）`、`^\d+[.．]`、`（1）`）。实验报告标题形态若不在此列（如「实验一」「1 实验原理」「第 1 章」），normalize 时这些段落会全部落入 body 兜底。
- 无「实验报告」专用规则集资产；全库 grep「实验报告」零命中。

## 环 2 · 规则集 → 编辑命令 —— 已通（有一处静默丢弃）

- `src/templates/rulesetToCommands.ts`：9 个段落组件按 `RULE_ORDER`（title→subtitle→heading1..4→caption→attachment→body）翻译为 `normalize` 规则（regex→`match.text`、position→`documentStart`/`after:<组件>`、fallback）；`page` 节 → `set /body/sectPr`（纸张/页边距/页脚距）+ `pageNumber footer` 命令。subtitle 另有 headingGuard 负向守卫防吞噬标题。
- spec 的 `normalize` 全文重排原语**已实现**：`src/docx-core/applyEdits.ts:applyNormalize`（首个命中者生效、`notText` 负向守卫、position 谓词、空规则集报错防误清空）。
- 测试佐证：`templates.test.ts`「规则集 → 内核命令」逐字段断言 +「端到端：内置公文规则集 normalize 一篇乱排版文档」全过；`workspace.test.ts`「模板链路：上传模板 → 规则集命令 → 应用到乱排版文档」过。

**缺口**：

- `heuristic` 类识别规则（连续居中判 title、`isTableElement`、图表相邻判 caption）在 `recognizerToMatch`（`rulesetToCommands.ts:86`）**静默跳过**——`table` 组件因此完全没有执行路径，`smartAlign` 等表格样式无法落地。
- `pageNumber` 只翻译单一 `align`（`rulesetToCommands.ts:158` 只取 `oddAlign`），gongwen 规则集的 odd/even 双对齐丢失。

## 环 3 · 编辑内核 —— 已通

- `applyEdits.ts` 命令表：`set`（路径 + 全文 match 批量）/ `add` / `remove` / `move` / `findReplace` / `normalize` / `numbering`（define/attach/clear）/ `pageNumber`；单条失败不中断，结构化错误含自愈建议；产出前生成后自检（重解析全部件 + document 结构冒烟）。
- `primitives.ts` 原语：段落（对齐/首行缩进字符或磅/行距磅值或倍数/段前段后/分页控制/大纲级/pStyle）、run（eastAsia/ascii 字体、半磅字号、粗斜下划线、颜色、highlight）、节（页边距 cm/纸张/方向/pgNumType）、findReplace 跨 run 拆分、页脚 PAGE 字段（含 footer 部件注册）。
- **round-trip 保真**（`test/roundtrip.test.ts`，19 个真实 docx 样本回归集）：未编辑 round-trip 部件级无损、XML build 幂等（不动点）、dirty 部件语义——3 例全过。`edits.test.ts` 15 例全过。

**缺口**：表格原语缺失（无 w:tbl 级操作）；其余 spec MVP 原语子集均已落地。

## 环 4 · agent 工具链 —— 已通

- `src/agent-core/tools.ts`：四个自研工具全部就位——`doc_outline`（结构 dump）、`doc_edit`（走唯一 seam + 自动快照）、`template_read`（规则集/占位符/大纲 + 直接可用的 `rulesetCommands`）、`version_store`（list/rollback）。`executionMode: 'sequential'`。
- `bridge.ts`：`createAgentSession` 白名单只放四工具（裁剪内置文件工具，防绕过内核）；`SYSTEM_PRIMER` 首次发送注入排版规范；docId 首次发送注入文档结构摘要（前 40 段路径+文本预览）。规则集**不预注入**系统提示词，由 agent 经 `template_read` 按需取（提示词第 5 条明确教了这个动作）。
- 测试佐证：`e2e-mock-agent.test.ts`（mock OpenAI 端点：bridge→pi session→doc_edit→自动快照新版本→文档实际变化→事件流 tool_start/tool_end/done）**通过**（2.4s）；`workspace.test.ts`「agent 工具：doc_outline → doc_edit → version_store 全链路（无 LLM）」过。

## 环 5 · 预览与导出 —— 已通（代码接通；未做 UI 人工验证）

- 预览：`public/preview.html` iframe 内 `docx-preview.renderAsync`（vendor 自带 `docx-preview.min.js` + `jszip.min.js`），父页面 `app.js` 经 `workbench:getBuffer` 取 ArrayBuffer、postMessage 传入、250ms 防抖合并连续编辑，iframe 回传 `preview-ready` 冒烟信号。
- 导出：`main.ts:workbench:download` IPC → `showSaveDialog` → 写盘当前或指定历史版本；左栏版本时间线每条有「预览」「下载」。
- **缺口**：UI 无「按模板一键重排」按钮——重排只能走对话路径（agent 调 template_read→doc_edit）；`templateRulesetCommands` 未在 `preload.mts` 白名单暴露，渲染层拿不到规则集命令。目的地流程里「按模板规则集一键重排」目前实际形态是「对 agent 说一句话，agent 转 normalize」，没有脱开 LLM 的直接按钮。

## 环 6 · 端到端测试 —— 部分

| 测试 | 覆盖段 | 本次实测 |
| --- | --- | --- |
| `e2e-mock-agent.test.ts` | 对话→工具→编辑→快照→事件流（mock LLM） | ✅ 通过 |
| `e2e-agent.test.ts` | 真实 LLM 排版任务 + 文档自查 | ⏭ 跳过（无 `PAIBAN_E2E`/凭证，设计如此） |
| `workspace.test.ts` 冒烟 | 上传→编辑→快照→大纲→版本→回滚→下载（service 级） | ✅ 通过 |
| `roundtrip.test.ts` | 19 样本保真回归 | ✅ 通过 |

**缺口**：无 UI 级 e2e（`main.ts:runSmoke` 需 `PAIBAN_SMOKE=1` 手动跑，不在 npm test 内）；spec 回归集要求的「docx-preview 渲染冒烟」未自动化（roundtrip 只测 buffer 层）；无实验报告场景 fixture。

---

## 总评：跑通最小闭环还缺什么

**链路主体已通，目的地可行。** 上传 docx → 选模板 → 对 agent 说「按 X 模板排版」→（agent 经 template_read 取 rulesetCommands 原样喂 doc_edit）→ normalize 全文重排 → 预览刷新 → 下载 docx——每一环都有代码且有测试佐证（mock e2e 已覆盖 agent 全链路的结构行为）。剩余缺口按优先级：

1. **真实 LLM 验证（最后一公里）**：默认 `deepseek/deepseek-v4-flash`，需 `DEEPSEEK_API_KEY` 后跑 `PAIBAN_E2E=1 npm run test:e2e-agent` + 人工开应用过一遍目的地流程。目前真实凭证链路零验证。
2. **实验报告规则集资产**：要么手工写一份 `templates/rulesets/lab-report-default`（识别正则贴合实验报告标题形态），要么接受「上传实验报告模板时 recognizers 继承公文正则」的局限——后者对编号形态不符的标题会全部落入 body。
3. **一键重排 UI 入口**：目的地文案是「一键重排」，现状只有对话路径；最小补法是把 `templateRulesetCommands` 透到 preload + 前端加一个「按模板重排」按钮（不经 LLM，直接 `applyCommands`）。
4. **Word/WPS 双端人工打开验证**：spec 首要风险的最后一步，round-trip 保真已过但双端实测无记录。
5. **heuristic 识别与表格原语**：影响乱排版文档中「无编号标题 / 图表标题 / 表格样式」的识别率，不阻塞主链路但决定重排质量上限。
6. **测试脚本兼容性（附带发现）**：`npm test` 的目录参数形式在 Node 24/Windows 下跑不起来，建议改为 `node --test dist/test/**/*.test.js`。
