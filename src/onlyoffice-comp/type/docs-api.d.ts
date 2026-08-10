// OnlyOffice DocsAPI 由 SDK 的 api.js 在运行时注入 window（见 util/initialize.ts），
// 编译期不可见，这里补全局声明。DocEditor 的完整配置类型见 type/word-api.ts。
interface Window {
  DocsAPI?: {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    DocEditor: any;
  };
}
