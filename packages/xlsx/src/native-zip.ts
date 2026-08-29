const LOCAL_FILE_SIGNATURE = 0x04034b50;
const CENTRAL_FILE_SIGNATURE = 0x02014b50;
const END_SIGNATURE = 0x06054b50;
const UTF8_FLAG = 0x0800;
const MAX_ENTRIES = 2048;
const MAX_ENTRY_SIZE = 32 * 1024 * 1024;
const MAX_TOTAL_SIZE = 128 * 1024 * 1024;
const MAX_COMPRESSION_RATIO = 200;

const crcTable = new Uint32Array(256);
for (let index = 0; index < crcTable.length; index += 1) {
  let value = index;
  for (let bit = 0; bit < 8; bit += 1) {
    value = (value & 1) === 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  }
  crcTable[index] = value >>> 0;
}

const crc32 = (data: Uint8Array): number => {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = (crc >>> 8) ^ (crcTable[(crc ^ byte) & 0xff] as number);
  }
  return (crc ^ 0xffffffff) >>> 0;
};

const assertRange = (data: Uint8Array, offset: number, length: number): void => {
  if (!Number.isSafeInteger(offset) || !Number.isSafeInteger(length) || offset < 0 || length < 0 || offset + length > data.length) {
    throw new Error("Invalid ZIP record bounds.");
  }
};

const readUint16 = (data: Uint8Array, offset: number): number => {
  assertRange(data, offset, 2);
  return data[offset]! | (data[offset + 1]! << 8);
};

const readUint32 = (data: Uint8Array, offset: number): number => {
  assertRange(data, offset, 4);
  return (data[offset]! | (data[offset + 1]! << 8) | (data[offset + 2]! << 16) | (data[offset + 3]! << 24)) >>> 0;
};

const writeUint16 = (data: Uint8Array, offset: number, value: number): void => {
  data[offset] = value & 0xff;
  data[offset + 1] = (value >>> 8) & 0xff;
};

const writeUint32 = (data: Uint8Array, offset: number, value: number): void => {
  data[offset] = value & 0xff;
  data[offset + 1] = (value >>> 8) & 0xff;
  data[offset + 2] = (value >>> 16) & 0xff;
  data[offset + 3] = (value >>> 24) & 0xff;
};

const normalizeEntryName = (name: string): string => {
  const normalized = name.replaceAll("\\", "/");
  if (!normalized || normalized.startsWith("/") || /^[A-Za-z]:/.test(normalized) || normalized.includes("\0")) {
    throw new Error("Unsafe ZIP entry path.");
  }
  const segments = normalized.split("/");
  if (segments.some((segment) => !segment || segment === "." || segment === "..")) {
    throw new Error("Unsafe ZIP entry path.");
  }
  return segments.join("/");
};

const findEndRecord = (data: Uint8Array): number => {
  const minimumOffset = Math.max(0, data.length - 65_557);
  for (let offset = data.length - 22; offset >= minimumOffset; offset -= 1) {
    if (readUint32(data, offset) === END_SIGNATURE) {
      return offset;
    }
  }
  throw new Error("ZIP end record not found.");
};

const inflateRaw = async (input: Uint8Array): Promise<Uint8Array> => {
  const stream = new Blob([input.slice().buffer]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
};

export const unpackZip = async (input: Uint8Array): Promise<Record<string, Uint8Array>> => {
  const data = new Uint8Array(input);
  const endOffset = findEndRecord(data);
  const diskNumber = readUint16(data, endOffset + 4);
  const centralDisk = readUint16(data, endOffset + 6);
  const entryCount = readUint16(data, endOffset + 10);
  const centralSize = readUint32(data, endOffset + 12);
  const centralOffset = readUint32(data, endOffset + 16);
  if (diskNumber !== 0 || centralDisk !== 0 || entryCount > MAX_ENTRIES) {
    throw new Error("Unsupported ZIP archive layout.");
  }
  assertRange(data, centralOffset, centralSize);

  const files: Record<string, Uint8Array> = {};
  let offset = centralOffset;
  let totalSize = 0;
  for (let index = 0; index < entryCount; index += 1) {
    if (readUint32(data, offset) !== CENTRAL_FILE_SIGNATURE) {
      throw new Error("Invalid ZIP central directory.");
    }
    const flags = readUint16(data, offset + 8);
    const method = readUint16(data, offset + 10);
    const expectedCrc = readUint32(data, offset + 16);
    const compressedSize = readUint32(data, offset + 20);
    const uncompressedSize = readUint32(data, offset + 24);
    const nameLength = readUint16(data, offset + 28);
    const extraLength = readUint16(data, offset + 30);
    const commentLength = readUint16(data, offset + 32);
    const localOffset = readUint32(data, offset + 42);
    assertRange(data, offset + 46, nameLength + extraLength + commentLength);
    if ((flags & 1) !== 0 || (method !== 0 && method !== 8) || uncompressedSize > MAX_ENTRY_SIZE) {
      throw new Error("Unsupported or oversized ZIP entry.");
    }
    if (compressedSize === 0 ? uncompressedSize !== 0 : uncompressedSize / compressedSize > MAX_COMPRESSION_RATIO) {
      throw new Error("Suspicious ZIP compression ratio.");
    }
    totalSize += uncompressedSize;
    if (totalSize > MAX_TOTAL_SIZE) {
      throw new Error("ZIP archive exceeds the uncompressed size limit.");
    }
    const nameBytes = data.subarray(offset + 46, offset + 46 + nameLength);
    const name = normalizeEntryName(new TextDecoder((flags & UTF8_FLAG) !== 0 ? "utf-8" : "windows-1252").decode(nameBytes));
    if (name in files) {
      throw new Error("Duplicate ZIP entry.");
    }
    if (readUint32(data, localOffset) !== LOCAL_FILE_SIGNATURE) {
      throw new Error("Invalid ZIP local header.");
    }
    const localNameLength = readUint16(data, localOffset + 26);
    const localExtraLength = readUint16(data, localOffset + 28);
    const payloadOffset = localOffset + 30 + localNameLength + localExtraLength;
    assertRange(data, payloadOffset, compressedSize);
    const compressed = data.subarray(payloadOffset, payloadOffset + compressedSize);
    const output = method === 0
      ? new Uint8Array(compressed)
      : await inflateRaw(compressed);
    if (output.length !== uncompressedSize || crc32(output) !== expectedCrc) {
      throw new Error("Corrupt ZIP entry.");
    }
    files[name] = output;
    offset += 46 + nameLength + extraLength + commentLength;
  }
  if (offset > centralOffset + centralSize) {
    throw new Error("Invalid ZIP central directory size.");
  }
  return files;
};

export const packZip = (files: Record<string, Uint8Array | string>): Uint8Array => {
  const encoder = new TextEncoder();
  const entries = Object.entries(files).map(([rawName, content]) => {
    const name = normalizeEntryName(rawName);
    const nameBytes = encoder.encode(name);
    const data = typeof content === "string" ? encoder.encode(content) : new Uint8Array(content);
    if (data.length > MAX_ENTRY_SIZE || nameBytes.length > 0xffff) {
      throw new Error("ZIP entry exceeds writer limits.");
    }
    return { nameBytes, data, crc: crc32(data), offset: 0 };
  });
  if (entries.length > MAX_ENTRIES || entries.reduce((sum, entry) => sum + entry.data.length, 0) > MAX_TOTAL_SIZE) {
    throw new Error("ZIP archive exceeds writer limits.");
  }

  const localSize = entries.reduce((sum, entry) => sum + 30 + entry.nameBytes.length + entry.data.length, 0);
  const centralSize = entries.reduce((sum, entry) => sum + 46 + entry.nameBytes.length, 0);
  const output = new Uint8Array(localSize + centralSize + 22);
  let offset = 0;
  for (const entry of entries) {
    entry.offset = offset;
    writeUint32(output, offset, LOCAL_FILE_SIGNATURE);
    writeUint16(output, offset + 4, 20);
    writeUint16(output, offset + 6, UTF8_FLAG);
    writeUint16(output, offset + 8, 0);
    writeUint32(output, offset + 14, entry.crc);
    writeUint32(output, offset + 18, entry.data.length);
    writeUint32(output, offset + 22, entry.data.length);
    writeUint16(output, offset + 26, entry.nameBytes.length);
    output.set(entry.nameBytes, offset + 30);
    output.set(entry.data, offset + 30 + entry.nameBytes.length);
    offset += 30 + entry.nameBytes.length + entry.data.length;
  }
  const centralOffset = offset;
  for (const entry of entries) {
    writeUint32(output, offset, CENTRAL_FILE_SIGNATURE);
    writeUint16(output, offset + 4, 20);
    writeUint16(output, offset + 6, 20);
    writeUint16(output, offset + 8, UTF8_FLAG);
    writeUint16(output, offset + 10, 0);
    writeUint32(output, offset + 16, entry.crc);
    writeUint32(output, offset + 20, entry.data.length);
    writeUint32(output, offset + 24, entry.data.length);
    writeUint16(output, offset + 28, entry.nameBytes.length);
    writeUint32(output, offset + 42, entry.offset);
    output.set(entry.nameBytes, offset + 46);
    offset += 46 + entry.nameBytes.length;
  }
  writeUint32(output, offset, END_SIGNATURE);
  writeUint16(output, offset + 8, entries.length);
  writeUint16(output, offset + 10, entries.length);
  writeUint32(output, offset + 12, centralSize);
  writeUint32(output, offset + 16, centralOffset);
  return output;
};
