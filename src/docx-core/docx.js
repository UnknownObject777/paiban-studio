// docx-core/docx.js — OOXML 文档的加载 / 保存封装。
//
// round-trip 保真架构（dirty-tracking）：
//   - openDocx 为每个部件保留原始字节（XML 部件的原始文本 text、二进制部件的原始 bytes）。
//   - 编辑内核只修改需要变更的部件，并把该部件标记为 dirty。
//   - toBuffer 时，未 dirty 的 XML 部件**原样写回原始文本**，二进制部件原样写回原始字节；
//     只有 dirty 部件才走 buildXml（fast-xml-parser preserveOrder 树 → 序列化）。
//
// 这保证"打开 → 保存（未编辑）→ 打开"100% 无损（含 XML 声明风格、self-closing 空格、
// 实体转义等字节细节）；编辑只影响被修改的部件，不引入意外漂移。Word/WPS 双端打开一致。

import PizZip from 'pizzip';
import { parseXml, buildXml } from './xml.js';

export const XML_EXT = /\.(xml|rels)$/i;

/**
 * 打开 docx buffer。
 * @param {Buffer|ArrayBuffer|Uint8Array} buffer
 * @returns {{ parts: Map<string, Part> }}
 *  Part（xml）: { kind:'xml', tree, meta, text, dirty:boolean }
 *  Part（binary）: { kind:'binary', bytes:Uint8Array }
 */
export function openDocx(buffer) {
  const zip = new PizZip(buffer);
  const parts = new Map();
  for (const name of Object.keys(zip.files)) {
    const entry = zip.files[name];
    if (entry.dir) continue;
    if (XML_EXT.test(name)) {
      const text = entry.asText();
      const { tree, meta } = parseXml(text);
      parts.set(name, { kind: 'xml', tree, meta, text, dirty: false });
    } else {
      parts.set(name, { kind: 'binary', bytes: entry.asUint8Array() });
    }
  }
  return { parts };
}

/** 标记某 XML 部件已被修改（编辑内核修改 tree 后调用）。 */
export function markDirty(docx, name) {
  const part = docx.parts.get(name);
  if (!part || part.kind !== 'xml') return;
  part.dirty = true;
}

/** 读取某 XML 部件的 fxp 树（编辑入口）。 */
export function getXmlTree(docx, name) {
  const part = docx.parts.get(name);
  if (!part || part.kind !== 'xml') return undefined;
  return part.tree;
}

/**
 * 把打开后的文档对象重新打包为 .docx buffer。
 * @param {{ parts: Map }} docx 由 openDocx 返回（或编辑内核返回）的对象
 * @returns {Buffer}
 */
export function toBuffer(docx) {
  const zip = new PizZip();
  for (const [name, part] of docx.parts) {
    if (part.kind === 'xml') {
      // 未修改部件原样写回，已修改部件走 buildXml
      const payload = part.dirty ? buildXml(part.tree, part.meta) : part.text;
      zip.file(name, payload);
    } else {
      zip.file(name, part.bytes);
    }
  }
  return zip.generate({
    type: 'nodebuffer',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
  });
}

// 便捷：buffer → buffer 的无编辑 round-trip（供测试与无编辑场景）。
export function roundTripDocx(buffer) {
  return toBuffer(openDocx(buffer));
}
