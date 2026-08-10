// @ts-nocheck — vendored brotli 解码器（JS 移植，无类型标注），不参与 strict 检查
function base64ToBytes(base64: string) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/**
 * Browser-friendly dictionary bootstrap: dictionary.bin is a compressed
 * copy of the static dictionary; decompress it on first use.
 */
import { BrotliDecompressBuffer } from "./decode.js";
import dictionaryBin from "./dictionary.bin.js";

export function init() {
  const compressed = base64ToBytes(dictionaryBin);
  return BrotliDecompressBuffer(compressed);
}
