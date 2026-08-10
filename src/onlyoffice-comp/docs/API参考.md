# API 参考

[← 完整示例](./完整示例.md) | [注意事项 →](./注意事项与支持格式.md)

## 常量

### `ONLYOFFICE_ID`

编辑器容器的 DOM ID，默认为 `'iframe-office-id'`。

### `ONLYOFFICE_CONTAINER_CONFIG`

| 字段 | 说明 |
|------|------|
| `PARENT_SELECTOR` | 父元素选择器 `.onlyoffice-container` |
| `PARENT_CLASS_NAME` | 父元素类名 `onlyoffice-container` |
| `STYLE` | 容器绝对定位样式 `{ position, inset }` |

### `ONLYOFFICE_EVENT_KEYS`

| 常量 | 值 | 说明 |
|------|-----|------|
| `SAVE_DOCUMENT` | `saveDocument` | 保存完成，含 `binData` |
| `DOCUMENT_READY` | `documentReady` | 文档就绪 |
| `LOADING_CHANGE` | `loadingChange` | Loading 状态 |
| `ONSAVE` | `onSave` | 保存流程结束（轻量） |

### `FILE_TYPE`

- `FILE_TYPE.DOCX` — Word
- `FILE_TYPE.XLSX` — Excel
- `FILE_TYPE.PPTX` — PowerPoint

### `ONLYOFFICE_LANG_KEY`

- `ONLYOFFICE_LANG_KEY.ZH` — `zh`（**默认语言**）
- `ONLYOFFICE_LANG_KEY.EN` — `en`

### `OFFICE_THEME`

OnlyOffice 编辑器界面主题（`customization.uiTheme`）。

| 常量 | 值 | 说明 |
|------|-----|------|
| `OFFICE_THEME.WHITE` | `theme-white` | 浅色（**默认**） |
| `OFFICE_THEME.CLASSIC_LIGHT` | `theme-classic-light` | 经典浅色 |
| `OFFICE_THEME.LIGHT` | `theme-light` | Light |
| `OFFICE_THEME.DARK` | `theme-dark` | 深色 |
| `OFFICE_THEME.NIGHT` | `theme-night` | 夜间 |
| `OFFICE_THEME.CONTRAST_DARK` | `theme-contrast-dark` | 高对比深色 |

相关导出：

- `DEFAULT_OFFICE_THEME` — 默认 `OFFICE_THEME.WHITE`
- `OFFICE_THEME_OPTIONS` — `{ id, label }[]`，供示例页 / UI 下拉使用

### `READONLY_SWITCH_MIN_DELAY_MS`

只读 ↔ 编辑切换时 loading 最短展示时长，值为 `200`（ms）。

### `STATIC_RESOURCE`

OnlyOffice SDK 与 x2t 静态资源路径总入口。

```typescript
import { STATIC_RESOURCE } from "@/components/onlyoffice-web-comp";

STATIC_RESOURCE.onlyoffice.root     // 默认 /packages/onlyoffice/9.4.0-develop
STATIC_RESOURCE.onlyoffice.apiUrl     // api.js 绝对 URL
STATIC_RESOURCE.x2t.script            // x2t.js 路径
STATIC_RESOURCE.x2t.wasm              // x2t.wasm 路径
```

可通过环境变量 `NEXT_PUBLIC_APP_ROOT` 覆盖 SDK 根路径。

### `__custom_font_registry__`

SDK 侧字体注册表，定义于 `public/packages/onlyoffice/9.4.0-develop/sdkjs/common/AllFonts.js`。键为 catalog 文件 id（如 `"1001"`），值为文档内字体别名数组。完整配置流程见 [字体配置](./字体配置.md)。

## 类型定义

### `OnlyOfficeManagerOptions`

```typescript
type OnlyOfficeManagerOptions = {
  containerId?: string;
  fileType: FileType;
  defaultFileName: string;
  readOnly?: boolean;
  lang?: OnlyOfficeLang;
  theme?: OfficeTheme;
  officeXmlEvent?: OfficeXmlEventConfig;
};
```

### `OpenDocumentInput`

```typescript
type OpenDocumentInput = {
  fileName: string;
  file?: File;
  isNew?: boolean;
  readOnly?: boolean;
  officeXmlEvent?: OfficeXmlEventConfig;
};
```

### `OfficeXmlEventConfig`

```typescript
type OfficeXmlEventConfig = {
  isEnable?: boolean;   // 默认 false
  limitBytes?: number;  // 默认 2GB
};
```

`OFFICE_XML_EVENT_CONFIG.default` 为默认配置。开启后，打开 Office ZIP 文件前会统计 `.xml` / `.rels` 的解压后总大小，超过阈值时触发 `OFFICE_XML_SIZE_LIMIT_EXCEEDED` 事件并显示默认错误层。

### `DocumentReadyData`

```typescript
type DocumentReadyData = {
  fileName: string;
  fileType: string;
  instanceId?: string;
};
```

### `SaveDocumentData`

```typescript
type SaveDocumentData = {
  fileName: string;
  fileType: string;
  binData: Uint8Array;
  instanceId: string;
  media?: Record<string, Uint8Array>;
};
```

### `OnSaveData`

```typescript
type OnSaveData = {
  fileName: string;
  instanceId: string;
};
```

### `LoadingChangeData`

```typescript
type LoadingChangeData = {
  loading: boolean;
};
```

### `OfficeTheme`

编辑器界面主题 ID，与 `OFFICE_THEME` 常量值一致，例如 `"theme-white"`、`"theme-dark"`。类型自 `const/index.ts` 导出为 `OfficeThemeId`，在 `EditorManager` / `OnlyOfficeManager` 中写作 `OfficeTheme`。

### `AscWordApiMethod`

Word 编辑器 iframe 内 SDK 方法名联合类型，定义于 `type/word-api.ts`。用于 `EditorManager.subscribe({ type, fn })` 的 `type` 参数。

常用条目（节选）：

```typescript
// 批注
| 'asc_onAddComment'
| 'asc_onChangeCommentData'
| 'asc_onRemoveComment'
// 修订
| 'asc_onShowRevisionsChange'
// 文档状态
| 'asc_onDocumentModifiedChanged'
| 'asc_onSaveCallback'
```

完整列表见源码 `type/word-api.ts`。

## 导出清单

统一从包入口导入：

```typescript
import {
  // 门面
  OnlyOfficeManager,
  onlyOfficeManagerFactory,
  // 底层
  EditorManager,
  EditorLogger,
  editorManager,
  editorManagerFactory,
  createEditorView,
  initializeOnlyOffice,
  convertBinToDocument,
  // 事件
  onlyofficeEventbus,
  ONLYOFFICE_EVENT_KEYS,
  // 常量
  FILE_TYPE,
  ONLYOFFICE_ID,
  ONLYOFFICE_CONTAINER_CONFIG,
  ONLYOFFICE_LANG_KEY,
  OFFICE_THEME,
  DEFAULT_OFFICE_THEME,
  OFFICE_THEME_OPTIONS,
  STATIC_RESOURCE,
  // store
  setDocumentObj,
  getDocumentObj,
} from "@/components/onlyoffice-web-comp";
```

`EditorLogger` 记录单个编辑器实例的 socket、downloadAs、x2t worker 与操作日志。通过 `manager.getLogger()` / `manager.printLogs()` 读取或打印当前实例日志。

类型通过 `export type *` 从 `type/word-api.ts`、`type/sdk-internal.ts` 导出。
