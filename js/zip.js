/**
 * Minimal offline ZIP reader/writer (STORED + DEFLATE raw) using Web Streams.
 * Zero external dependencies; works in iOS Safari 16.4+, Chrome, and Node 20+.
 */

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder('utf-8');

const CRC_TABLE = (() => {
  const tbl = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    tbl[i] = c >>> 0;
  }
  return tbl;
})();

export function crc32(bytes) {
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) {
    crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

async function runRawStream(data, Ctor) {
  const ds = new Ctor('deflate-raw');
  const writer = ds.writable.getWriter();
  const writePromise = writer.write(data).then(() => writer.close());
  const reader = ds.readable.getReader();
  const chunks = [];
  let total = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    total += value.length;
  }
  await writePromise;
  const out = new Uint8Array(total);
  let off = 0;
  for (const ch of chunks) {
    out.set(ch, off);
    off += ch.length;
  }
  return out;
}

async function inflateRaw(bytes) {
  if (typeof DecompressionStream === 'undefined') {
    throw new Error('Trình duyệt không hỗ trợ giải nén DEFLATE ngoại tuyến.');
  }
  return runRawStream(bytes, DecompressionStream);
}

async function deflateRaw(bytes) {
  if (typeof CompressionStream === 'undefined') {
    return null;
  }
  return runRawStream(bytes, CompressionStream);
}

function findEocdOffset(view) {
  const len = view.byteLength;
  const minOffset = Math.max(0, len - 65557);
  for (let i = len - 22; i >= minOffset; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      return i;
    }
  }
  return -1;
}

/**
 * Unzips a Buffer/Uint8Array into a Map<string, Uint8Array>.
 */
export async function unzipBytes(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocdOff = findEocdOffset(view);
  if (eocdOff < 0) {
    throw new Error('File không phải định dạng ZIP/DOCX hợp lệ.');
  }

  const totalEntries = view.getUint16(eocdOff + 10, true);
  if (totalEntries > 2000) {
    throw new Error('Tệp chứa quá nhiều thành phần (vượt mức an toàn).');
  }
  const cdOffset = view.getUint32(eocdOff + 16, true);
  let pos = cdOffset;
  const entries = new Map();
  let totalUncompressed = 0;
  const MAX_TOTAL_BYTES = 120 * 1024 * 1024;

  for (let idx = 0; idx < totalEntries; idx++) {
    if (pos + 46 > bytes.length || view.getUint32(pos, true) !== 0x02014b50) {
      throw new Error('Cấu trúc thư mục trung tâm ZIP bị lỗi.');
    }
    const method = view.getUint16(pos + 10, true);
    const compSize = view.getUint32(pos + 20, true);
    const uncompSize = view.getUint32(pos + 24, true);
    const nameLen = view.getUint16(pos + 28, true);
    const extraLen = view.getUint16(pos + 30, true);
    const commentLen = view.getUint16(pos + 32, true);
    const localHeaderOff = view.getUint32(pos + 42, true);

    const rawName = textDecoder.decode(bytes.subarray(pos + 46, pos + 46 + nameLen));
    const entryName = rawName.replace(/\\/g, '/');
    pos += 46 + nameLen + extraLen + commentLen;

    if (entryName.includes('..') || entryName.startsWith('/')) {
      throw new Error(`Đường dẫn không an toàn trong gói nén: ${entryName}`);
    }
    if (entryName.endsWith('/')) continue;

    if (localHeaderOff + 30 > bytes.length || view.getUint32(localHeaderOff, true) !== 0x04034b50) {
      throw new Error(`Header nội bộ bị lỗi cho mục ${entryName}`);
    }
    const lhNameLen = view.getUint16(localHeaderOff + 26, true);
    const lhExtraLen = view.getUint16(localHeaderOff + 28, true);
    const dataStart = localHeaderOff + 30 + lhNameLen + lhExtraLen;
    const dataEnd = dataStart + compSize;
    if (dataEnd > bytes.length) {
      throw new Error(`Dữ liệu bị cắt xén ở mục ${entryName}`);
    }

    totalUncompressed += uncompSize;
    if (totalUncompressed > MAX_TOTAL_BYTES) {
      throw new Error('Dung lượng giải nén vượt giới hạn an toàn 120MB.');
    }

    const slice = bytes.subarray(dataStart, dataEnd);
    let data;
    if (method === 0) {
      data = new Uint8Array(slice);
    } else if (method === 8) {
      data = await inflateRaw(slice);
    } else {
      throw new Error(`Không hỗ trợ phương thức nén ZIP #${method}`);
    }
    entries.set(entryName, data);
  }

  return entries;
}

/**
 * Packs Array<[name, Uint8Array | string]> into a ZIP Uint8Array.
 */
export async function zipEntries(entryList) {
  const localParts = [];
  const cdParts = [];
  let offset = 0;

  const now = new Date();
  const dosTime = ((now.getHours() & 0x1f) << 11) | ((now.getMinutes() & 0x3f) << 5) | ((now.getSeconds() >> 1) & 0x1f);
  const dosDate = (((Math.max(1980, now.getFullYear()) - 1980) & 0x7f) << 9) | (((now.getMonth() + 1) & 0x0f) << 5) | (now.getDate() & 0x1f);

  for (const [name, rawVal] of entryList) {
    const cleanName = name.replace(/^\/+/, '');
    const nameBytes = textEncoder.encode(cleanName);
    const rawBytes = typeof rawVal === 'string' ? textEncoder.encode(rawVal) : rawVal;
    const crc = crc32(rawBytes);

    let compressed = rawBytes;
    let method = 0;
    if (rawBytes.length >= 64) {
      const def = await deflateRaw(rawBytes);
      if (def && def.length < rawBytes.length) {
        compressed = def;
        method = 8;
      }
    }

    const lh = new Uint8Array(30 + nameBytes.length);
    const lhv = new DataView(lh.buffer);
    lhv.setUint32(0, 0x04034b50, true);
    lhv.setUint16(4, 20, true);
    lhv.setUint16(6, 0x0800, true); // UTF-8 filename flag
    lhv.setUint16(8, method, true);
    lhv.setUint16(10, dosTime, true);
    lhv.setUint16(12, dosDate, true);
    lhv.setUint32(14, crc, true);
    lhv.setUint32(18, compressed.length, true);
    lhv.setUint32(22, rawBytes.length, true);
    lhv.setUint16(26, nameBytes.length, true);
    lhv.setUint16(28, 0, true);
    lh.set(nameBytes, 30);

    const cd = new Uint8Array(46 + nameBytes.length);
    const cdv = new DataView(cd.buffer);
    cdv.setUint32(0, 0x02014b50, true);
    cdv.setUint16(4, 20, true);
    cdv.setUint16(6, 20, true);
    cdv.setUint16(8, 0x0800, true);
    cdv.setUint16(10, method, true);
    cdv.setUint16(12, dosTime, true);
    cdv.setUint16(14, dosDate, true);
    cdv.setUint32(16, crc, true);
    cdv.setUint32(20, compressed.length, true);
    cdv.setUint32(24, rawBytes.length, true);
    cdv.setUint16(28, nameBytes.length, true);
    cdv.setUint16(30, 0, true);
    cdv.setUint16(32, 0, true);
    cdv.setUint16(34, 0, true);
    cdv.setUint16(36, 0, true);
    cdv.setUint32(38, 0, true);
    cdv.setUint32(42, offset, true);
    cd.set(nameBytes, 46);

    localParts.push(lh, compressed);
    cdParts.push(cd);
    offset += lh.length + compressed.length;
  }

  const cdSize = cdParts.reduce((s, p) => s + p.length, 0);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, entryList.length, true);
  ev.setUint16(10, entryList.length, true);
  ev.setUint32(12, cdSize, true);
  ev.setUint32(16, offset, true);

  const totalLen = offset + cdSize + eocd.length;
  const out = new Uint8Array(totalLen);
  let cursor = 0;
  for (const part of localParts) {
    out.set(part, cursor);
    cursor += part.length;
  }
  for (const part of cdParts) {
    out.set(part, cursor);
    cursor += part.length;
  }
  out.set(eocd, cursor);
  return out;
}
