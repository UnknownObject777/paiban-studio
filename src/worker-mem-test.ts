// 临时诊断 worker（冒烟后删除）：验证 module worker 内的 Wasm 内存分配能力。
self.postMessage({
  isSecureContext: self.isSecureContext,
  hasDS: 'DecompressionStream' in self,
});
try {
  const m = new WebAssembly.Memory({ initial: 4536, maximum: 32768 });
  self.postMessage({ memOK: true, bytes: m.buffer.byteLength });
} catch (e) {
  self.postMessage({ memOK: false, error: (e as Error).message });
}
export {};
