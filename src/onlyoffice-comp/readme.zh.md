# OnlyOffice Web Comp

> 📖 [English](readme.md) | 中文

基于 OnlyOffice 静态 SDK 的 **Web 端文档编辑组件库**，支持 Word / Excel / PowerPoint 的在线编辑、只读预览、导出与 x2t 格式转换。**无需自建 Document Server**，只需托管 SDK 静态资源。

> 本文档为**入口页**。完整说明见 [`docs/`](./docs/概述.md)。演示站点会直接渲染本目录下的 Markdown，对应路由 `/docs`。

## 文档导航

| 编号 | 文档 | 说明 |
|------|------|------|
| - | [概述](./docs/概述.md) | 文档索引与阅读路径 |
| - | [快速开始](./docs/快速开始.md) | 初始化、容器挂载、创建编辑器 |
| - | [核心 API](./docs/核心API.md) | `OnlyOfficeManager`、`EditorManager`、多实例 |
| - | [事件系统](./docs/事件系统.md) | EventBus、事件类型与监听 |
| - | [完整示例](./docs/完整示例.md) | React 集成模式 |
| - | [API 参考](./docs/API参考.md) | 常量、类型定义 |
| - | [注意事项与支持格式](./docs/注意事项与支持格式.md) | 前置条件、文件格式、常见坑 |
| - | [批注修订与 Word API](./docs/批注修订与-Word-API.md) | 批注、修订、SDK 回调 |
| - | [字体配置](./docs/字体配置.md) | `__custom_font_registry__`、catalog 转换 |
| - | [单实例示例](./docs/单实例示例.md) | 单编辑器 Demo 与源码说明 |
| - | [多实例示例](./docs/多实例示例.md) | Tab 多实例完整源码 |

**推荐阅读路径**

| 场景 | 路径 |
|------|------|
| 首次接入 | [快速开始](./docs/快速开始.md) → [核心 API](./docs/核心API.md) |
| 在线体验 | [单实例示例](./docs/单实例示例.md) · [多实例示例](./docs/多实例示例.md) |
| React 页面集成 | [完整示例](./docs/完整示例.md) |
| 多实例 / 导出 | [核心 API](./docs/核心API.md) · [事件系统](./docs/事件系统.md) |

## 目录结构

```
onlyoffice-web-comp/
├── const/       常量、静态资源路径、文件类型、主题
├── store/       文档 / 语言等跨页面状态
├── util/        SDK 初始化、x2t 转换、下载
├── core/        EditorManager、OnlyOfficeManager、EventBus
├── feature/     批注、修订
├── docs/        完整使用文档（Markdown 源文件）
└── internal/    mock server / x2t worker（不对外导出）
```

## 最小示例

```typescript
import {
  OnlyOfficeManager,
  ONLYOFFICE_ID,
  FILE_TYPE,
} from "@/components/onlyoffice-web-comp";

// 新建空白文档
const manager = await OnlyOfficeManager.create({
  containerId: ONLYOFFICE_ID,
  fileType: FILE_TYPE.DOCX,
  defaultFileName: "New_Document.docx",
});

// 打开已有 File（先取文件，再挂载）
const file = await fetch("/test.xlsx").then((r) => r.blob())
  .then((blob) => new File([blob], "test.xlsx", { type: blob.type }));
await OnlyOfficeManager.createWithFile({
  containerId: ONLYOFFICE_ID,
  fileType: FILE_TYPE.XLSX,
  defaultFileName: "test.xlsx",
}, file);
```

更多用法（事件、导出、多实例、主题、语言、只读切换）见 [docs/核心API.md](./docs/核心API.md)。

## 本仓库中的演示

`docs/` 下的 Markdown 由演示站点渲染；示例页内嵌可交互编辑器。

| 路由 | 说明 |
|------|------|
| `/docs` | 渲染本目录全部文档 |
| `/docs/demos?tab=single` | 单实例在线示例（[单实例示例](./docs/单实例示例.md)） |
| `/docs/demos?tab=multi` | 多实例 Tab 示例（[多实例示例](./docs/多实例示例.md)） |

演示组件：`src/features/demo/`（`office-preview-page.tsx`、`tabs-multi-page.tsx`）

本地启动：`pnpm dev` → http://localhost:3001

## 相关链接

- [项目根 README（仓库总览）](../../../README.zh.md)
- [OnlyOffice 官方 API](https://api.onlyoffice.com/zh-CN/docs/docs-api/usage-api/config/document/)
