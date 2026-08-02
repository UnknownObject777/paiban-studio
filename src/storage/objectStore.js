// storage/objectStore.js — 内容寻址对象存储（D8：S3 兼容抽象接口 + 本地文件实现）。
//
// 抽象接口（未来可换 MinIO / 云 OSS 实现）：
//   put(bytes) → key（sha256 hex）     幂等：同内容同 key，重复 put 不写盘
//   get(key)   → Buffer
//   has(key)   → boolean
//
// 本地实现布局：<baseDir>/objects/<ab>/<abcdef...>（sha256 前 2 位分桶）。

import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync, existsSync, renameSync } from 'node:fs';
import { join, dirname } from 'node:path';

export function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
}

export class LocalFsObjectStore {
  constructor(baseDir) {
    this.baseDir = baseDir;
    this.objectsDir = join(baseDir, 'objects');
    mkdirSync(this.objectsDir, { recursive: true });
  }

  _path(key) {
    return join(this.objectsDir, key.slice(0, 2), key);
  }

  put(bytes) {
    const key = sha256(bytes);
    const p = this._path(key);
    if (!existsSync(p)) {
      mkdirSync(dirname(p), { recursive: true });
      // 临时文件 + rename，避免半写状态
      const tmp = p + '.tmp-' + process.pid;
      writeFileSync(tmp, bytes);
      renameSync(tmp, p);
    }
    return key;
  }

  get(key) {
    const p = this._path(key);
    if (!existsSync(p)) {
      const err = new Error(`对象不存在: ${key}`);
      err.code = 'OBJECT_NOT_FOUND';
      throw err;
    }
    return readFileSync(p);
  }

  has(key) {
    return existsSync(this._path(key));
  }
}
