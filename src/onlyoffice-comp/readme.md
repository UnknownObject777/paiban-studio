# OnlyOffice Web Comp

> 📖 English | [中文](readme.zh.md)

A **browser-side document editor component library** built on the OnlyOffice static SDK. Supports online editing, read-only preview, export, and x2t conversion for Word, Excel, and PowerPoint. **No self-hosted Document Server required**—only static SDK assets on your site.

> This file is an **entry point**. Full documentation lives in [`docs/`](./docs/概述.md). On the demo site, these Markdown files are rendered at `/docs`.

## Documentation

| # | Doc | Description |
|---|-----|-------------|
| - | [Overview](./docs/概述.md) | Index and reading path |
| - | [Quick Start](./docs/快速开始.md) | Init, container mount, create editor |
| - | [Core API](./docs/核心API.md) | `OnlyOfficeManager`, `EditorManager`, multi-instance |
| - | [Event System](./docs/事件系统.md) | EventBus, event types, listeners |
| - | [Full Examples](./docs/完整示例.md) | React integration patterns |
| - | [API Reference](./docs/API参考.md) | Constants and types |
| - | [Notes & Formats](./docs/注意事项与支持格式.md) | Prerequisites, formats, pitfalls |
| - | [Comments, Revisions & Word API](./docs/批注修订与-Word-API.md) | Comments, revisions, SDK callbacks |
| - | [Fonts](./docs/字体配置.md) | `__custom_font_registry__`, catalog conversion |
| - | [Single-instance Demo](./docs/单实例示例.md) | Single editor demo + source walkthrough |
| - | [Multi-instance Demo](./docs/多实例示例.md) | Full Tab demo source |

**Suggested paths**

| Scenario | Path |
|----------|------|
| First integration | [快速开始](./docs/快速开始.md) → [核心 API](./docs/核心API.md) |
| Try live demos | [单实例示例](./docs/单实例示例.md) · [多实例示例](./docs/多实例示例.md) |
| React page integration | [完整示例](./docs/完整示例.md) |
| Multi-instance / export | [核心 API](./docs/核心API.md) · [事件系统](./docs/事件系统.md) |

## Package Layout

```
onlyoffice-web-comp/
├── const/       Constants, static paths, file types, themes
├── store/       Document / language state
├── util/        SDK init, x2t conversion, download
├── core/        EditorManager, OnlyOfficeManager, EventBus
├── feature/     Comments, revisions
├── docs/        Full documentation (Markdown source of truth)
└── internal/    Mock server / x2t worker (not exported)
```

## Minimal Example

```typescript
import {
  OnlyOfficeManager,
  ONLYOFFICE_ID,
  FILE_TYPE,
} from "@/components/onlyoffice-web-comp";

// Create a blank document
const manager = await OnlyOfficeManager.create({
  containerId: ONLYOFFICE_ID,
  fileType: FILE_TYPE.DOCX,
  defaultFileName: "New_Document.docx",
});

// Open an existing File (fetch first, then mount)
const file = await fetch("/test.xlsx").then((r) => r.blob())
  .then((blob) => new File([blob], "test.xlsx", { type: blob.type }));
await OnlyOfficeManager.createWithFile({
  containerId: ONLYOFFICE_ID,
  fileType: FILE_TYPE.XLSX,
  defaultFileName: "test.xlsx",
}, file);
```

See [docs/核心API.md](./docs/核心API.md) for events, export, multi-instance, theme, language, and read-only toggling.

## Demos in This Repo

Documentation Markdown under `docs/` is rendered by the demo site. Live editors are embedded on the demo pages below.

| Route | Description |
|-------|-------------|
| `/docs` | All docs from this `docs/` folder |
| `/docs/demos?tab=single` | Single-instance demo ([单实例示例](./docs/单实例示例.md)) |
| `/docs/demos?tab=multi` | Multi-instance Tab demo ([多实例示例](./docs/多实例示例.md)) |

Demo components: `src/features/demo/` (`office-preview-page.tsx`, `tabs-multi-page.tsx`)

Run locally: `pnpm dev` → http://localhost:3001

## Links

- [Repository README (project overview)](../../../README.md)
- [OnlyOffice official API](https://api.onlyoffice.com/docs/docs-api/usage-api/config/document/)
