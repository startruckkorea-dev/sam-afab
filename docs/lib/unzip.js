// 의존성 없는 최소 ZIP 리더 — .docx(=zip) 안의 word/document.xml 을 꺼내기 위한 것.
// 브라우저: DecompressionStream('deflate-raw') 사용 (Chrome 103+/Edge/FF113+/Safari16.4+)
// Node(테스트): zlib.inflateRawSync 사용.
// app.js 에 이미 있는 zip WRITER 와 짝을 이루는 READER 다.
(function (root) {
  'use strict';

  function u16(b, o) { return b[o] | (b[o + 1] << 8); }
  function u32(b, o) { return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0; }

  async function inflateRaw(bytes) {
    // Node 경로 (테스트 하네스)
    if (typeof process !== 'undefined' && process.versions && process.versions.node) {
      const zlib = require('zlib');
      return new Uint8Array(zlib.inflateRawSync(Buffer.from(bytes)));
    }
    const ds = new DecompressionStream('deflate-raw');
    const stream = new Blob([bytes]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  // 중앙 디렉터리(EOCD)를 찾아 엔트리 목록을 만든다.
  function readCentralDirectory(b) {
    // EOCD 시그니처 0x06054b50 를 뒤에서부터 탐색 (주석 최대 64KB)
    let eocd = -1;
    const min = Math.max(0, b.length - 66000);
    for (let i = b.length - 22; i >= min; i--) {
      if (u32(b, i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('ZIP: EOCD 를 찾을 수 없습니다 (손상된 파일?)');
    const count = u16(b, eocd + 10);
    let off = u32(b, eocd + 16);
    const entries = [];
    for (let n = 0; n < count; n++) {
      if (u32(b, off) !== 0x02014b50) break;
      const method = u16(b, off + 10);
      const compSize = u32(b, off + 20);
      const fnLen = u16(b, off + 28);
      const exLen = u16(b, off + 30);
      const cmLen = u16(b, off + 32);
      const localOff = u32(b, off + 42);
      let name = '';
      for (let i = 0; i < fnLen; i++) name += String.fromCharCode(b[off + 46 + i]);
      entries.push({ name: name, method: method, compSize: compSize, localOff: localOff });
      off += 46 + fnLen + exLen + cmLen;
    }
    return entries;
  }

  // 엔트리 하나의 원본 바이트를 돌려준다.
  async function readEntry(b, e) {
    const lo = e.localOff;
    if (u32(b, lo) !== 0x04034b50) throw new Error('ZIP: 로컬 헤더 불일치 ' + e.name);
    const fnLen = u16(b, lo + 26);
    const exLen = u16(b, lo + 28);
    const start = lo + 30 + fnLen + exLen;
    // 로컬 헤더의 크기 필드는 0 일 수 있어(스트리밍 저장) 중앙 디렉터리 값을 쓴다.
    const data = b.subarray(start, start + e.compSize);
    if (e.method === 0) return data;            // stored
    if (e.method === 8) return inflateRaw(data); // deflate
    throw new Error('ZIP: 지원하지 않는 압축 방식 ' + e.method);
  }

  /** zip(ArrayBuffer|Uint8Array) 안에서 파일 하나를 UTF-8 문자열로 읽는다. */
  async function readTextFile(buf, path) {
    const b = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    const entries = readCentralDirectory(b);
    const e = entries.find(function (x) { return x.name === path; });
    if (!e) throw new Error('ZIP: 항목 없음 ' + path);
    const bytes = await readEntry(b, e);
    return new TextDecoder('utf-8').decode(bytes);
  }

  const api = { readTextFile: readTextFile, readCentralDirectory: readCentralDirectory };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Unzip = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
