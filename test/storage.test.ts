// storage.test.js — 存储版本链行为测试（spec 测试策略：快照幂等 / 回滚语义 / 内容寻址去重）。

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, readdirSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { LocalFsObjectStore, sha256 } from '../src/storage/objectStore.js';
import { VersionStore } from '../src/storage/versionStore.js';

function freshStore() {
  const dir = mkdtempSync(join(tmpdir(), 'paiban-storage-'));
  return {
    dir,
    objects: new LocalFsObjectStore(dir),
    versions: new VersionStore(dir, new LocalFsObjectStore(dir)),
    cleanup: () => rmSync(dir, { recursive: true, force: true }),
  };
}

test('对象存储：put 幂等（同内容同 key，只写一次盘），get 还原', () => {
  const { dir, objects, cleanup } = freshStore();
  const a = Buffer.from('hello docx bytes');
  const k1 = objects.put(a);
  const k2 = objects.put(Buffer.from('hello docx bytes'));
  assert.equal(k1, k2);
  assert.equal(k1, sha256(a));
  assert.deepEqual(objects.get(k1), a);
  assert.ok(objects.has(k1));
  // 分桶目录下只有一个对象文件
  const bucket = readdirSync(join(dir, 'objects', k1.slice(0, 2)));
  assert.equal(bucket.length, 1);
  cleanup();
});

test('版本链：创建文档产生 v1，快照随内容变化递增', () => {
  const { versions, cleanup } = freshStore();
  const { docId, version } = versions.createDocument(Buffer.from('v1 content'), { name: '公文.docx', origin: 'upload' });
  assert.equal(version.id, 'v1');
  const s1 = versions.snapshot(docId, Buffer.from('v2 content'), { source: 'edit', note: '标题居中' });
  assert.equal(s1.created, true);
  assert.equal(s1.version.id, 'v2');
  assert.equal(s1.version.parent, 'v1');
  assert.equal(s1.version.note, '标题居中');
  assert.equal(versions.list(docId).length, 2);
  assert.deepEqual(versions.getBuffer(docId), Buffer.from('v2 content'));
  assert.deepEqual(versions.getBuffer(docId, 'v1'), Buffer.from('v1 content'));
  cleanup();
});

test('快照幂等：内容无变化不产生空版本', () => {
  const { versions, cleanup } = freshStore();
  const { docId } = versions.createDocument(Buffer.from('same'));
  const s = versions.snapshot(docId, Buffer.from('same'), { source: 'edit' });
  assert.equal(s.created, false);
  assert.equal(s.version.id, 'v1');
  assert.equal(versions.list(docId).length, 1);
  cleanup();
});

test('回滚记录为新版本：hash 指向历史内容，parent 为当前 head', () => {
  const { versions, cleanup } = freshStore();
  const { docId } = versions.createDocument(Buffer.from('original'));
  versions.snapshot(docId, Buffer.from('edited-1'));
  versions.snapshot(docId, Buffer.from('edited-2'));
  const r = versions.rollback(docId, 'v1');
  assert.equal(r.created, true);
  assert.equal(r.version.id, 'v4');
  assert.equal(r.version.parent, 'v3');
  assert.equal(r.version.source, 'rollback');
  assert.deepEqual(versions.getBuffer(docId), Buffer.from('original'));
  // 再回滚到 v3（edited-2）
  const r2 = versions.rollback(docId, 'v3');
  assert.deepEqual(versions.getBuffer(docId), Buffer.from('edited-2'));
  // 回滚到与 head 同内容的版本 → 幂等不产生新版本
  const r3 = versions.rollback(docId, 'v3');
  assert.equal(r3.created, false);
  cleanup();
});

test('内容寻址去重：两个文档相同内容共享对象，只存一份', () => {
  const { dir, versions, cleanup } = freshStore();
  const content = Buffer.from('shared content');
  const a = versions.createDocument(content, { name: 'a.docx' });
  const b = versions.createDocument(Buffer.from('shared content'), { name: 'b.docx' });
  assert.equal(a.version.hash, b.version.hash);
  const bucket = readdirSync(join(dir, 'objects', a.version.hash.slice(0, 2)));
  assert.equal(bucket.length, 1);
  cleanup();
});

test('错误：不存在文档 / 版本报结构化错误', () => {
  const { versions, cleanup } = freshStore();
  assert.throws(() => versions.list('nope'), /文档不存在/);
  const { docId } = versions.createDocument(Buffer.from('x'));
  assert.throws(() => versions.getBuffer(docId, 'v99'), /版本不存在/);
  assert.throws(() => versions.rollback(docId, 'v99'), /版本不存在/);
  cleanup();
});

test('listDocuments 汇总工作文档', () => {
  const { versions, cleanup } = freshStore();
  versions.createDocument(Buffer.from('a'), { name: '甲.docx' });
  versions.createDocument(Buffer.from('b'), { name: '乙.docx' });
  const docs = versions.listDocuments();
  assert.equal(docs.length, 2);
  assert.deepEqual(docs.map((d) => d.name).sort(), ['乙.docx', '甲.docx'].sort());
  cleanup();
});
