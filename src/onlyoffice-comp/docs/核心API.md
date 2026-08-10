# 核心 API

[← 快速开始](./快速开始.md) | [事件系统 →](./事件系统.md)

## 业务门面

`OnlyOfficeManager` 面向业务页面，收敛初始化、开文档、导出、只读、语言、主题等调用。

### 创建空白文档

调用 `OnlyOfficeManager.create(options)` 加载 DocsAPI，并以 `defaultFileName` 新建空白文档。

```typescript
import {
  OnlyOfficeManager,
  ONLYOFFICE_ID,
  FILE_TYPE,
  DEFAULT_OFFICE_THEME,
} from "@/components/onlyoffice-web-comp";

const manager = await OnlyOfficeManager.create({
  containerId: ONLYOFFICE_ID,   // 可选，默认 ONLYOFFICE_ID
  fileType: FILE_TYPE.DOCX,
  defaultFileName: "New_Document.docx",
  readOnly: false,
  lang: "zh",                   // 可选，默认 zh
  theme: DEFAULT_OFFICE_THEME,  // 可选，默认 theme-white
});
```

### 打开本地文件

调用 `OnlyOfficeManager.createWithFile(options, file)`，先拿到 `File`，再一次性挂载编辑器（不会先打开空白文档）。

```typescript
const manager = await OnlyOfficeManager.createWithFile(
  {
    containerId: ONLYOFFICE_ID,
    fileType: FILE_TYPE.XLSX,
    defaultFileName: "test.xlsx",
  },
  file,
);
```

### 静态资源地址运行时注册

OnlyOffice SDK、x2t、PDF 字体等静态资源默认从 `public/packages` 加载。需要在运行时切到 CDN 或独立静态服务器时，可在创建编辑器前注册资源地址。

```typescript
import {
  FILE_TYPE,
  OnlyOfficeManager,
  getStaticResource,
  isOnlyOfficeCdnMode,
} from "@/components/onlyoffice-web-comp";

// packages 根地址，目录下应包含 onlyoffice/{version}/...
OnlyOfficeManager.registerStaticResource({
  cdnOrigin: "https://ca0eac5f.onlyoffice-packages.pages.dev",
});

const resource = getStaticResource();
console.log(resource.onlyoffice.apiUrl);
console.log(resource.x2t.script);
console.log(isOnlyOfficeCdnMode());

const manager = await OnlyOfficeManager.create({
  fileType: FILE_TYPE.DOCX,
  defaultFileName: "New_Document.docx",
});

// 恢复默认 public/packages 配置
OnlyOfficeManager.resetStaticResource();
```

`registerStaticResource` 支持以下参数：

```typescript
type OnlyOfficeStaticResourceOptions = {
  /** CDN packages 根地址，例如 https://ca0eac5f.onlyoffice-packages.pages.dev */
  cdnOrigin?: string | null;
  /** CDN 的 SDK 版本；默认当前 9.4 SDK */
  onlyofficeVersion?: string | null;
};
```

运行时切换已有实例时，需要先销毁旧实例，再注册新地址并重新创建编辑器。DocsAPI script、preload iframe、x2t worker 都会按新的静态资源地址重新初始化。

### 实例方法

| 方法 | 说明 |
|------|------|
| `openDocument(input)` | 打开/切换文档（上传、新建、重开） |
| `openFile(file, readOnly?)` | 打开本地文件 |
| `openNew(fileName, readOnly?)` | 新建文档 |
| `isReady()` | 是否已打开文档 |
| `getReadOnly()` / `setReadOnly()` / `toggleReadOnly()` | 只读切换（同步，底层 `asc_setRestriction`） |
| `getLanguage()` / `setLanguage()` / `toggleLanguage()` | 语言切换 |
| `getTheme()` / `setTheme(theme)` / `toggleTheme()` | 界面主题切换（见下文） |
| `exportDocument()` | 导出 bin 数据 |
| `exportAsBlob()` | 导出为 Blob |
| `downloadExport()` | 导出并触发浏览器下载 |
| `createConnector(options?)` | 获取当前编辑器唯一的 Developer Edition Connector；销毁或重开编辑器时自动断开 |
| `onLoadingChange(handler)` | 监听 loading，返回取消函数 |
| `getEditor()` | 获取底层 `EditorManager` |
| `getLogger()` | 获取当前实例的 `EditorLogger` |
| `printLogs()` | 将当前实例日志历史打印到控制台 |
| `destroy()` | 销毁实例 |

### Developer Edition Connector

`createConnector()` 返回当前编辑器唯一的 OnlyOffice Connector，用于从父页面调用编辑器的 Automation API。重复调用会返回同一实例（与 `getLogger()` 的生命周期一致），连接器会使用当前 iframe 的真实 `frameEditorId`，本地和 CDN 跨域模式均可用；编辑器销毁或重开文档时组件会自动断开并释放它。

```typescript
const connector = manager.createConnector();

connector.executeMethod("GetEditorType", [], () => {
  console.log("Connector request completed");
});

// 可提前释放；不调用也会在 manager.destroy() 时自动断开。
connector.disconnect();
```

### 实例日志

每个 `EditorManager` 都持有一个 `EditorLogger`，用于记录当前实例的 socket、downloadAs、x2t worker 与关键操作日志。日志仍会按原有 console 参数输出，便于 e2e/CDP 继续监听；同时可通过 `manager.printLogs()` 打印当前实例的历史记录。

### 主题切换

主题对应 OnlyOffice `customization.uiTheme`。切换时会短暂 remount iframe（先保存未提交编辑），可通过 `onLoadingChange` 监听 loading。

```typescript
import {
  OnlyOfficeManager,
  OFFICE_THEME,
  OFFICE_THEME_OPTIONS,
  DEFAULT_OFFICE_THEME,
  type OfficeTheme,
} from "@/components/onlyoffice-web-comp";

// 创建时指定初始主题
const manager = await OnlyOfficeManager.create({
  fileType: FILE_TYPE.DOCX,
  defaultFileName: "New_Document.docx",
  theme: OFFICE_THEME.DARK,
});

// 运行时切换
await manager.setTheme(OFFICE_THEME.NIGHT);
const current = manager.getTheme();

// 在浅色 / 深色之间快捷切换
await manager.toggleTheme();

// UI 下拉框可遍历 OFFICE_THEME_OPTIONS
OFFICE_THEME_OPTIONS.map(({ id, label }) => (
  <option key={id} value={id}>{label}</option>
));
```

可用主题常量见 [API 参考 · OFFICE_THEME](./API参考.md#office_theme)。

### 打开文档参数

```typescript
type OpenDocumentInput = {
  fileName: string;
  file?: File;
  isNew?: boolean;
  readOnly?: boolean;
};
```

## 多实例管理

`OnlyOfficeManagerFactory` 用于多容器场景，按 `containerId` 缓存 `OnlyOfficeManager` 门面。组件导出的单例为 `onlyOfficeManagerFactory`。

```typescript
import {
  onlyOfficeManagerFactory,
  FILE_TYPE,
} from "@/components/onlyoffice-web-comp";

const manager = await onlyOfficeManagerFactory.open(
  {
    containerId: "editor-1",
    fileType: FILE_TYPE.DOCX,
    defaultFileName: "New_Document.docx",
    readOnly: false,
  },
  {
    fileName: "New_Document.docx",
    isNew: true,
  },
);

onlyOfficeManagerFactory.get("editor-1");
onlyOfficeManagerFactory.destroy("editor-1");
onlyOfficeManagerFactory.destroyAll();
```

## 底层能力

### 初始化资源

调用 `initializeOnlyOffice()` 手动初始化 OnlyOffice 静态资源。

```typescript
import { initializeOnlyOffice } from "@/components/onlyoffice-web-comp";

await initializeOnlyOffice();
```

- 单例模式，多次调用只初始化一次
- `OnlyOfficeManager.create` / `EditorManager.create` 内部会自动调用
- 仅在手动 `fromEditor` 绑定等高级场景需要显式调用

### 创建编辑器视图

调用 `createEditorView(options)` 直接创建底层编辑器视图。

```typescript
import { createEditorView } from "@/components/onlyoffice-web-comp";

await createEditorView({
  isNew: boolean;
  fileName: string;
  file?: File;
  url?: string;
  loader?: (url: string) => Promise<ArrayBuffer>;
  fileType?: string;
  readOnly?: boolean;
  lang?: string;              // 默认跟随 store，初始为 zh
  containerId?: string;
  editorManager?: EditorManager;
  theme?: OfficeTheme;
});
```

**返回值：** `Promise<EditorManager>`

**支持的文件类型：**

- Word: `.docx`, `.doc`, `.odt`, `.rtf`, `.txt`
- Excel: `.xlsx`, `.xls`, `.ods`, `.csv`
- PowerPoint: `.pptx`, `.ppt`, `.odp`

### 编辑器管理器

`editorManagerFactory` 和 `EditorManager` 提供更底层的编辑器控制能力。

#### 单实例

```typescript
import { editorManagerFactory } from "@/components/onlyoffice-web-comp";

const editorManager = editorManagerFactory.getDefault();

if (editorManager.exists()) {
  // 编辑器已创建
}

const binData = await editorManager.export();

editorManager.setReadOnly(true);   // 同步方法
editorManager.setReadOnly(false);

const isReadOnly = editorManager.getReadOnly();

editorManager.destroy();
```

#### 多实例

```typescript
const manager1 = editorManagerFactory.create("editor-1");
const manager2 = editorManagerFactory.get("editor-2"); // 不存在时自动 create

const allManagers = editorManagerFactory.getAll();

editorManagerFactory.destroy("editor-1");
editorManagerFactory.destroyAll();
```

#### `EditorManager` 实例方法

| 方法 | 说明 |
|------|------|
| `exists()` | 检查编辑器是否存在 |
| `export()` | 导出文档二进制数据 |
| `setReadOnly(readOnly)` | 切换只读/可编辑（同步） |
| `getReadOnly()` | 获取当前只读状态 |
| `getInstanceId()` | 获取实例唯一 ID |
| `getContainerId()` | 获取容器 ID |
| `getFileName()` | 获取当前文件名 |
| `getTheme()` / `setTheme(theme)` | 获取 / 切换界面主题 |
| `updateMedia(key, url)` | 更新媒体文件映射 |
| `getMedia()` | 获取媒体文件映射 |
| `destroy()` | 销毁编辑器实例 |
| `subscribe({ type, fn })` | 订阅 Word SDK 回调，见 [批注修订与 Word API](./批注修订与-Word-API.md) |

**`export()` 返回值：**

```typescript
{
  fileName: string;
  fileType: string;
  binData: Uint8Array;
  instanceId?: string;
  media?: Record<string, Uint8Array>;
}
```

多实例下 `export()` 通过 `instanceId`（固定等于当前 `containerId`）过滤 `SAVE_DOCUMENT` 事件。只读模式下直接返回已缓存的 `binData`。

### 文档格式转换

调用 `convertBinToDocument()` 将 `Editor.bin` 转回目标 Office 文档格式。

```typescript
import { convertBinToDocument, FILE_TYPE } from "@/components/onlyoffice-web-comp";

const result = await convertBinToDocument(
  binData.binData,
  binData.fileName,
  FILE_TYPE.DOCX,
  binData.media,
);

// result: { fileName: string, data: ArrayBuffer }
```

业务页面导出优先使用 `OnlyOfficeManager.downloadExport()`，无需手动转换。
