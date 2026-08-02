// storage/versionStore.ts — 文档版本链（spec 模块 3）。
//
// 语义（对应测试策略"快照幂等 / 回滚语义 / 内容寻址去重"）：
//   - createDocument(buffer, meta)  新建工作文档，初始版本 v1
//   - snapshot(docId, buffer, info) 内容变化才产生新版本（幂等：与 head 同 hash 不产生空版本）
//   - rollback(docId, versionId)    回滚**记录为新版本**（hash 指向历史内容，parent 为当前 head）
//   - list(docId) / getBuffer(docId, versionId?) / head(docId)
//
// 布局：<baseDir>/docs/<docId>/versions.json + 内容存 ObjectStore（跨文档去重）。

import { mkdirSync, readFileSync, writeFileSync, existsSync, readdirSync, renameSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { randomUUID } from 'node:crypto';
import type { LocalFsObjectStore, DocxInput } from './objectStore.js';

export interface VersionEntry {
  id: string;
  hash: string;
  parent: string | null;
  at: string;
  source: string;
  note?: string;
}

export interface DocChain {
  docId: string;
  name: string;
  meta: Record<string, unknown>;
  createdAt: string;
  versions: VersionEntry[];
}

export interface SnapshotResult {
  version: VersionEntry;
  created: boolean;
  rolledBackTo?: string;
}

export class VersionStore {
  baseDir: string;
  store: LocalFsObjectStore;
  docsDir: string;

  constructor(baseDir: string, objectStore: LocalFsObjectStore) {
    this.baseDir = baseDir;
    this.store = objectStore;
    this.docsDir = join(baseDir, 'docs');
    mkdirSync(this.docsDir, { recursive: true });
  }

  _chainPath(docId: string): string {
    return join(this.docsDir, docId, 'versions.json');
  }

  _load(docId: string): DocChain {
    const p = this._chainPath(docId);
    if (!existsSync(p)) {
      const err = new Error(`文档不存在: ${docId}`) as Error & { code?: string };
      err.code = 'DOC_NOT_FOUND';
      throw err;
    }
    return JSON.parse(readFileSync(p, 'utf8')) as DocChain;
  }

  _save(docId: string, chain: DocChain): void {
    const dir = join(this.docsDir, docId);
    mkdirSync(dir, { recursive: true });
    const p = this._chainPath(docId);
    const tmp = p + '.tmp-' + process.pid;
    writeFileSync(tmp, JSON.stringify(chain, null, 2));
    renameSyncSafe(tmp, p);
  }

  /** 新建工作文档。meta: { name, origin: 'upload'|'template', ... } */
  createDocument(buffer: DocxInput, meta: Record<string, unknown> = {}): { docId: string; version: VersionEntry } {
    const docId = randomUUID().slice(0, 8);
    const hash = this.store.put(buffer);
    const chain: DocChain = {
      docId,
      name: (meta.name as string) || docId,
      meta,
      createdAt: new Date().toISOString(),
      versions: [{
        id: 'v1', hash, parent: null,
        at: new Date().toISOString(),
        source: (meta.origin as string) || 'upload',
        note: (meta.note as string) || '初始版本',
      }],
    };
    this._save(docId, chain);
    return { docId, version: chain.versions[0] };
  }

  /**
   * 快照。幂等：内容与 head 相同（hash 一致）→ 不产生新版本，返回 { version: head, created: false }。
   */
  snapshot(docId: string, buffer: DocxInput, { source = 'edit', note = '' }: { source?: string; note?: string } = {}): SnapshotResult {
    const chain = this._load(docId);
    const headV = chain.versions[chain.versions.length - 1];
    const hash = this.store.put(buffer);
    if (hash === headV.hash) {
      return { version: headV, created: false };
    }
    const version: VersionEntry = {
      id: 'v' + (chain.versions.length + 1),
      hash,
      parent: headV.id,
      at: new Date().toISOString(),
      source,
      note,
    };
    chain.versions.push(version);
    this._save(docId, chain);
    return { version, created: true };
  }

  /** 回滚到历史版本：记录为新版本（hash 指向历史内容）。 */
  rollback(docId: string, versionId: string, { note }: { note?: string } = {}): SnapshotResult {
    const chain = this._load(docId);
    const target = chain.versions.find((v) => v.id === versionId);
    if (!target) {
      const err = new Error(`版本不存在: ${docId}@${versionId}`) as Error & { code?: string };
      err.code = 'VERSION_NOT_FOUND';
      throw err;
    }
    const headV = chain.versions[chain.versions.length - 1];
    if (target.hash === headV.hash) {
      return { version: headV, created: false, rolledBackTo: versionId };
    }
    const version: VersionEntry = {
      id: 'v' + (chain.versions.length + 1),
      hash: target.hash,
      parent: headV.id,
      at: new Date().toISOString(),
      source: 'rollback',
      note: note || `回滚到 ${versionId}`,
    };
    chain.versions.push(version);
    this._save(docId, chain);
    return { version, created: true, rolledBackTo: versionId };
  }

  list(docId: string): VersionEntry[] {
    return this._load(docId).versions;
  }

  head(docId: string): VersionEntry {
    const vs = this.list(docId);
    return vs[vs.length - 1];
  }

  /** 取文档内容（默认 head）。 */
  getBuffer(docId: string, versionId?: string): Buffer {
    const chain = this._load(docId);
    const v = versionId
      ? chain.versions.find((x) => x.id === versionId)
      : chain.versions[chain.versions.length - 1];
    if (!v) {
      const err = new Error(`版本不存在: ${docId}@${versionId}`) as Error & { code?: string };
      err.code = 'VERSION_NOT_FOUND';
      throw err;
    }
    return this.store.get(v.hash);
  }

  /** 列出全部工作文档（摘要）。 */
  listDocuments(): Array<{ docId: string; name: string; meta: Record<string, unknown>; createdAt: string; head: string; versionCount: number }> {
    if (!existsSync(this.docsDir)) return [];
    return readdirSync(this.docsDir, { withFileTypes: true })
      .filter((d) => d.isDirectory())
      .map((d) => {
        try {
          const chain = this._load(d.name);
          return {
            docId: chain.docId, name: chain.name, meta: chain.meta,
            createdAt: chain.createdAt,
            head: chain.versions[chain.versions.length - 1].id,
            versionCount: chain.versions.length,
          };
        } catch {
          return null;
        }
      })
      .filter((x): x is NonNullable<typeof x> => x !== null);
  }
}

// renameSync 跨平台小封装（Windows 上目标存在时 rename 会失败，先删）
function renameSyncSafe(tmp: string, target: string): void {
  try {
    renameSync(tmp, target);
  } catch {
    rmSync(target, { force: true });
    renameSync(tmp, target);
  }
}
