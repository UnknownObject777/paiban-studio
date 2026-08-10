// @ts-nocheck — vendored brotli 解码器（JS 移植，无类型标注），不参与 strict 检查
/** Vendored Brotli decompressor (MIT, brotli.js / Google). See dec/decode.ts header. */
import { BrotliDecompressBuffer } from "./dec/decode.js";

export function brotliDecompress(input: Uint8Array): Uint8Array {
  return BrotliDecompressBuffer(input) as Uint8Array;
}
