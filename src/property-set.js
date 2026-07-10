const VT_I2 = 0x0002;
const VT_I4 = 0x0003;
const VT_R8 = 0x0005;
const VT_BOOL = 0x000b;
const VT_UI4 = 0x0013;
const VT_LPSTR = 0x001e;
const VT_LPWSTR = 0x001f;
const VT_FILETIME = 0x0040;

const SUMMARY_PROPERTY_NAMES = new Map([
  [2, "title"], [3, "subject"], [4, "creator"], [5, "keywords"],
  [6, "description"], [7, "template"], [8, "lastModifiedBy"], [9, "revision"],
  [10, "totalEditingTime"], [11, "lastPrinted"], [12, "created"], [13, "modified"],
  [14, "pages"], [15, "words"], [16, "characters"], [18, "application"], [19, "security"],
]);

const DOCUMENT_SUMMARY_PROPERTY_NAMES = new Map([
  [2, "category"], [3, "presentationFormat"], [4, "bytes"], [5, "lines"],
  [6, "paragraphs"], [7, "slides"], [8, "notes"], [9, "hiddenSlides"],
  [10, "multimediaClips"], [14, "manager"], [15, "company"],
  [16, "linksDirty"], [17, "charactersWithSpaces"], [26, "contentType"],
  [27, "contentStatus"], [28, "language"], [29, "documentVersion"],
]);

export function parseOlePropertySet(stream, { kind = "generic" } = {}) {
  if (!Buffer.isBuffer(stream)) stream = Buffer.from(stream ?? []);
  if (stream.length < 28 || stream.readUInt16LE(0) !== 0xfffe) {
    throw new Error("Invalid OLE property-set stream header");
  }
  const sectionCount = stream.readUInt32LE(24);
  if (sectionCount < 1 || 28 + sectionCount * 20 > stream.length) {
    throw new Error(`Invalid OLE property-set section count ${sectionCount}`);
  }
  const sections = [];
  for (let index = 0; index < sectionCount; index += 1) {
    const descriptor = 28 + index * 20;
    const offset = stream.readUInt32LE(descriptor + 16);
    sections.push(parsePropertySection(stream, offset, kind === "documentSummary" && index === 1));
  }
  const names = kind === "summary" ? SUMMARY_PROPERTY_NAMES
    : kind === "documentSummary" ? DOCUMENT_SUMMARY_PROPERTY_NAMES
      : new Map();
  const properties = {};
  for (const [id, value] of sections[0].values) {
    const name = names.get(id);
    if (name && value !== undefined) properties[name] = normalizeNamedPropertyValue(kind, id, value);
  }
  const customProperties = kind === "documentSummary" && sections[1]
    ? parseCustomProperties(sections[1])
    : [];
  return { properties, customProperties, sections };
}

function parsePropertySection(stream, offset, hasDictionary = false) {
  if (offset < 0 || offset + 8 > stream.length) throw new Error("Invalid OLE property-set section offset");
  const size = stream.readUInt32LE(offset);
  const end = offset + size;
  if (size < 8 || end > stream.length) throw new Error("Invalid OLE property-set section size");
  const count = stream.readUInt32LE(offset + 4);
  if (offset + 8 + count * 8 > end) throw new Error("Invalid OLE property table size");
  const entries = [];
  for (let index = 0; index < count; index += 1) {
    const entry = offset + 8 + index * 8;
    const id = stream.readUInt32LE(entry);
    const relativeOffset = stream.readUInt32LE(entry + 4);
    if (relativeOffset >= size) throw new Error(`Invalid OLE property ${id} offset`);
    entries.push({ id, offset: offset + relativeOffset });
  }
  entries.sort((a, b) => a.offset - b.offset);
  const codePageEntry = entries.find((entry) => entry.id === 1);
  const codePage = codePageEntry ? readTypedProperty(stream, codePageEntry.offset, end, 1252).value : 1252;
  const values = new Map();
  for (let index = 0; index < entries.length; index += 1) {
    const entry = entries[index];
    const propertyEnd = entries[index + 1]?.offset ?? end;
    if (entry.id === 0 && hasDictionary) {
      values.set(0, parsePropertyDictionary(stream, entry.offset, propertyEnd, codePage));
      continue;
    }
    values.set(entry.id, readTypedProperty(stream, entry.offset, propertyEnd, codePage).value);
  }
  return { offset, size, codePage, values };
}

function readTypedProperty(stream, offset, end, codePage) {
  if (offset + 4 > end) throw new Error("Invalid truncated OLE typed property");
  const type = stream.readUInt32LE(offset);
  const valueOffset = offset + 4;
  switch (type) {
    case VT_I2:
      requireBytes(valueOffset, 2, end, type);
      return { type, value: stream.readInt16LE(valueOffset) };
    case VT_I4:
      requireBytes(valueOffset, 4, end, type);
      return { type, value: stream.readInt32LE(valueOffset) };
    case VT_UI4:
      requireBytes(valueOffset, 4, end, type);
      return { type, value: stream.readUInt32LE(valueOffset) };
    case VT_R8:
      requireBytes(valueOffset, 8, end, type);
      return { type, value: stream.readDoubleLE(valueOffset) };
    case VT_BOOL:
      requireBytes(valueOffset, 2, end, type);
      return { type, value: stream.readUInt16LE(valueOffset) !== 0 };
    case VT_LPSTR: {
      requireBytes(valueOffset, 4, end, type);
      const byteLength = stream.readUInt32LE(valueOffset);
      requireBytes(valueOffset + 4, byteLength, end, type);
      const bytes = stream.subarray(valueOffset + 4, valueOffset + 4 + byteLength);
      return { type, value: decodeCodePage(stripTrailingNullBytes(bytes), codePage) };
    }
    case VT_LPWSTR: {
      requireBytes(valueOffset, 4, end, type);
      const charCount = stream.readUInt32LE(valueOffset);
      requireBytes(valueOffset + 4, charCount * 2, end, type);
      return { type, value: stripNulls(stream.subarray(valueOffset + 4, valueOffset + 4 + charCount * 2).toString("utf16le")) };
    }
    case VT_FILETIME: {
      requireBytes(valueOffset, 8, end, type);
      const ticks = stream.readBigUInt64LE(valueOffset);
      if (ticks === 0n) return { type, value: null };
      return { type, value: { filetimeTicks: ticks } };
    }
    default:
      return { type, value: undefined };
  }
}

function parsePropertyDictionary(stream, offset, end, codePage) {
  if (offset + 4 > end) throw new Error("Invalid OLE property dictionary");
  const count = stream.readUInt32LE(offset);
  const dictionary = new Map();
  let cursor = offset + 4;
  for (let index = 0; index < count; index += 1) {
    requireBytes(cursor, 8, end, "dictionary");
    const id = stream.readUInt32LE(cursor);
    const storedLength = stream.readUInt32LE(cursor + 4);
    const byteLength = codePage === 1200 ? storedLength * 2 : storedLength;
    cursor += 8;
    requireBytes(cursor, byteLength, end, "dictionary name");
    dictionary.set(id, decodeCodePage(stripTrailingNullBytes(stream.subarray(cursor, cursor + byteLength)), codePage));
    cursor += byteLength;
    cursor = (cursor + 3) & ~3;
  }
  return dictionary;
}

function parseCustomProperties(section) {
  const dictionary = section.values.get(0);
  if (!(dictionary instanceof Map)) return [];
  const properties = [];
  for (const [id, name] of dictionary) {
    if (id === 0 || id === 1) continue;
    const value = normalizeFiletimeValue(section.values.get(id));
    if (value !== undefined) properties.push({ id, name, value });
  }
  return properties;
}

function normalizeNamedPropertyValue(kind, id, value) {
  if (kind === "summary" && id === 10 && value?.filetimeTicks != null) {
    return Number(value.filetimeTicks / 600000000n);
  }
  return normalizeFiletimeValue(value);
}

function normalizeFiletimeValue(value) {
  if (value?.filetimeTicks == null) return value;
  const unixMilliseconds = Number((value.filetimeTicks - 116444736000000000n) / 10000n);
  return new Date(unixMilliseconds).toISOString();
}

function decodeCodePage(bytes, codePage) {
  if (codePage === 1200) return stripNulls(bytes.toString("utf16le"));
  const labels = new Map([
    [65001, "utf-8"], [1250, "windows-1250"], [1251, "windows-1251"],
    [1252, "windows-1252"], [1253, "windows-1253"], [1254, "windows-1254"],
    [1255, "windows-1255"], [1256, "windows-1256"], [1257, "windows-1257"],
    [1258, "windows-1258"], [932, "shift_jis"], [936, "gbk"], [949, "euc-kr"],
    [950, "big5"], [874, "windows-874"], [10000, "macintosh"],
  ]);
  const label = labels.get(codePage);
  if (!label) throw new Error(`Unimplemented OLE property-set code page ${codePage}`);
  return stripNulls(new TextDecoder(label, { fatal: false }).decode(bytes));
}

function stripTrailingNullBytes(bytes) {
  let end = bytes.length;
  while (end > 0 && bytes[end - 1] === 0) end -= 1;
  return bytes.subarray(0, end);
}

function stripNulls(value) {
  return value.replace(/\0+$/g, "");
}

function requireBytes(offset, length, end, type) {
  if (offset < 0 || length < 0 || offset + length > end) {
    throw new Error(`Invalid OLE property value for type ${type}`);
  }
}
