# 旧库资产盘点 — Word-Formatter-Pro → paiban-studio

> 日期：2026-08-02
> 旧库位置：`D:\全自动科研\aicoding\Word-Formatter-Pro`（GitHub: UnknownObject777/Word-Formatter-Pro，已归档为只读决策档案）
> 用途：实现开工前查这张表——什么东西在旧库哪里、以什么形式进新库。与 `docs/mvp-spec.md` 的 R2（规则迁移）/「workbench 处置」一节互为补充，本文是**逐文件级**的清单。

## 一、代码级直接复用（拷过来改，已经过验证）

| 旧库路径 | 是什么 | 验证依据 | 进新库后的去向 |
|---|---|---|---|
| `workbench/server/storage.js` | sha256 内容寻址对象存储 + 版本链 + meta JSON 原子写入（tmp+rename） | `test/smoke.js` 第 5 节版本管理断言 | `storage` 模块原样复用，仅改 `paths.js` 的数据目录约定（D8 决策的本地文件实现） |
| `workbench/public/`（index.html 3.3KB / app.js 12KB / style.css 5.8KB） | 三栏布局前端：模板库/对话流/预览，SSE 事件渲染 | 原型期人工走通 | 前端骨架复用，样式按 `AGENTS.md` 的 Figma design system 重写 |
| `workbench/test/smoke.js`（5.1KB） | 端到端冒烟骨架：起服务→打 API→断言，覆盖上传/解析/实例化/版本 | 原型期跑通 | 测试骨架复用，REST 端点改为 IPC/UI 级调用 |
| `workbench/test/e2e-agent.js`（4.9KB） | 真实 LLM 链路测试：agent 用工具完成排版任务并自查 | 原型期跑通 | 同上当骨架复用 |
| `workbench/data/` 下的 docx 样例 | 真实文档样本（docs/ 2 篇 + templates/ 1 篇，含版本链实例） | — | **round-trip 保真回归集的种子样本**（spec 首要风险 spike 的输入） |

## 二、借鉴范式（不直接拷，重写时对照）

| 旧库路径 | 借鉴什么 | 注意 |
|---|---|---|
| `workbench/server/agent.js`（19KB） | 模型配置优先级（界面配置 > 环境变量 > 默认值，对应 spec R4）、多 provider 端点适配、工具注册与事件桥接的形状 | 包名是旧的 `@mariozechner/pi-agent-core`，新库必须用 `@earendil-works/*`（spec D5）；`LEGACY_PROJECT_ROOT` 调用 Python 的路径整体废弃 |
| `workbench/server/templates.js`（4.2KB） | 模板上传→解析→持久化→实例化的流程划分、meta JSON 管理模式 | 解析逻辑要从 office.js 换成自研编辑内核；加入 #5 拍板的规则集抽取环节 |
| `workbench/server/routes.js` + `index.js` | API 端点划分（上传/对话/版本/下载）作为 IPC 通道设计的对照清单 | 新库不用本地 HTTP（R1 决策），逐端点映射为 IPC handler |
| `workbench/server/paths.js` | 数据目录集中管理 + `WORKBENCH_DATA_DIR` 环境变量覆盖的约定 | 新库沿用此模式（Electron 下指向 userData 目录） |
| 对话编辑「工具串行执行」 | `toolExecution: "sequential"` 防并发写同一文件 | 已在 spec R1 记录，实现时勿丢 |

## 三、知识资产（文档化提取，代码不带过来）

| 旧库路径 | 内容 | 提取去向（spec R2 四路分流） |
|---|---|---|
| `wfp_core.py`（1895 行，85KB） | 多年积累的排版识别引擎：主副标题识别、四级标题、图表标题定位、附件格式化、表格智能对齐、符号标准化 | 版式规范类→模板规则集数据；文本清理类→预处理器；结构识别类→agent 工具（后置）；python-docx 细节→丢弃 |
| `docs/knowledge/wfp-formatting-rules.md` | ✅ **已迁入新库**——上述规则的人类可读提炼 | 内置公文默认规则集（#5 混合起步）的直接翻译源 |
| `skills/doc-format/SKILL.md` | 用户视角的能力描述与边界（什么能排、什么跳过） | 规则集字段语义与 agent 系统提示词的措辞参考 |
| `skills/doc-format/references/cli-reference.md`、`config-reference.md` | 全部可配置项清单（字体/行距/空行模式/表格调整开关…） | `styles.json` 可配置字段的命名与取值蓝本 |
| `wfp_config.py`（45 行） | 配置 schema 与默认值 | 同上，配置项的最小权威定义 |
| `wfp_tests.py`（311 行） | 规则回归测试（标题识别/空行模式/附件等场景断言） | 把测试**场景**改写为 `normalize` 原语的 buffer→buffer 断言，测试方法不迁 |

## 四、工程与产品素材

| 旧库路径 | 内容 | 用途 |
|---|---|---|
| `packaging/build_release.py` | 三平台发布流水线（PyInstaller 不跨平台编译，须在各 OS 本机构建） | electron-builder 发布流水线的经验参照；信创 Kylin 构建注意事项 |
| `packaging/release-notes-v2.7.5.md`、`v2.7.6.md` | 面向中文办公用户的发版说明写法 | 后续发版沟通格式参考 |
| `wfp_version.py` | 版本号集中单点定义模式 | 沿用此模式（新库为 package.json version） |
| `screenshot.png`、`demo_word_before_after.png`、`demo_txt_before_after.png` | 排版前后对比演示图 | 产品页/README/模板画廊素材；新产品出 demo 后替换 |

## 五、明确废弃（决策已定，不要再翻）

- `workbench/server/office.js` —— OfficeCLI 二进制依赖（spec：原型已否决，仅借鉴协议范式）
- `wfp_gui.py`、`wfp_cli.py` —— Tkinter GUI / CLI 运行时路径（R2 丢弃类）
- `requirements.txt` —— Python 依赖（python-docx 等，全量 Node.js 重写后无对应物）
- `conversations/` —— 空目录，无内容
- 旧格式 .doc/.wps 转换路径 —— 依赖 COM/LibreOffice，MVP 明确不做

## 迁移执行提示

1. 一类资产（代码复用）拷进新库时**不带旧 git history**，逐个文件拷入并在 commit message 注明出处。
2. 拷 `storage.js` 时连同 `test/smoke.js` 第 5 节断言一起拷——版本链语义由测试锁定，防拷丢行为。
3. `workbench/data` 样例文档先拷入 `tests/fixtures/`（或等价目录）供 spike 使用，再逐步扩充真实 Word/WPS 产出样本。
