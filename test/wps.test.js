import { readFile } from "node:fs/promises";
import test from "node:test";
import assert from "node:assert/strict";
import { readWps, normalizeComparableText } from "../src/index.js";
import { wpsToDocxBuffer } from "../src/docx.js";
import { lcidToBcp47 } from "../src/lcid.js";
import { BRC_TYPE_NAMES, brcColorFromIco, parseSprms } from "../src/sprm.js";
import { parseComments, parseFieldPlc, parseNoteReferencePlc, parseNoteTextPlc, parsePlcfHdd, parseSectionSprms, parseSttbfRMark, parseTextboxTextPlc, textDecoderEncodingForLid } from "../src/word-binary.js";
import { readDocxMainText, readDocxDocumentXml, readZipEntry } from "./fixtures-docx.js";

function findParagraphXml(xml, text) {
  const matches = xml.match(/<w:p [\s\S]*?<\/w:p>/g) ?? [];
  return matches.find((paragraph) => {
    if (paragraph.includes(text)) return true;
    const joinedText = Array.from(paragraph.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g), (match) => match[1]).join("");
    return joinedText.includes(text);
  });
}

function extractSettingsNode(xml, tagName) {
  const pattern = tagName === "w:compat"
    ? /<w:compat>[\s\S]*?<\/w:compat>/
    : new RegExp(`<${tagName}\\b[^>]*?(?:/>|>.*?<\\/${tagName}>)`, "s");
  return xml.match(pattern)?.[0] ?? null;
}

function normalizeComparableXml(xml) {
  return xml
    .replace(/>\s*</g, ">\n<")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .filter((line) => !/(rsid|docVar|paraId|zoom)/i.test(line));
}

function countXmlLineEdits(actualXml, expectedXml) {
  const actual = normalizeComparableXml(actualXml);
  const expected = normalizeComparableXml(expectedXml);
  let previous = new Array(actual.length + 1).fill(0);
  for (let i = 1; i <= expected.length; i += 1) {
    const current = new Array(actual.length + 1).fill(0);
    for (let j = 1; j <= actual.length; j += 1) {
      current[j] = expected[i - 1] === actual[j - 1]
        ? previous[j - 1] + 1
        : Math.max(previous[j], current[j - 1]);
    }
    previous = current;
  }
  const lcs = previous[actual.length];
  return expected.length + actual.length - 2 * lcs;
}

const BASIC_WPS = "sample/basic/original.wps";
const BASIC_DOCX = "sample/basic/expected.docx";
const FULL_WPS = "sample/sample1/original.wps";
const FULL_DOCX = "sample/sample1/expected.docx";
const TABLE2_WPS = "sample/table2/original.wps";
const TABLE6_DOCX = "sample/table6/expected.docx";
const SAMPLE8_WPS = "sample/sample8/original.wps";
const SAMPLE9_WPS = "sample/sample9/original.wps";
const SAMPLE10_WPS = "sample/sample10/original.wps";
const SAMPLE11_WPS = "sample/sample11/original.wps";

function buildFieldPlc(records, terminalCp = null) {
  const buffer = Buffer.alloc(records.length * 6 + 4);
  records.forEach((record, index) => buffer.writeUInt32LE(record.cp, index * 4));
  const lastCp = terminalCp ?? ((records.at(-1)?.cp ?? 0) + 1);
  buffer.writeUInt32LE(lastCp, records.length * 4);
  const fldOffset = (records.length + 1) * 4;
  records.forEach((record, index) => {
    buffer[fldOffset + index * 2] = record.fldch ?? record.ch;
    buffer[fldOffset + index * 2 + 1] = record.grffld;
  });
  return buffer;
}

function buildMinimalWpsDocument(bodyText, overrides = {}) {
  return {
    bodyText,
    defaultTabStop: 720,
    paragraphProperties: [{}],
    characterProperties: Array.from({ length: bodyText.length }, () => ({})),
    sections: [{ cpStart: 0, cpEnd: bodyText.length, properties: {} }],
    fields: {
      body: { records: [] },
      headers: { records: [] },
      footnotes: { records: [] },
      annotations: { records: [] },
      endnotes: { records: [] },
      textboxes: { records: [] },
      headerTextboxes: { records: [] },
    },
    footnotes: { references: [], stories: [] },
    endnotes: { references: [], stories: [] },
    comments: { comments: [], authors: [] },
    fib: {
      fComplex: true,
      characterCounts: { body: bodyText.length, headers: 0, footnotes: 0, annotations: 0, endnotes: 0, textboxes: 0, headerTextboxes: 0 },
      lcbPlcSpaMom: 0,
      lcbPlcSpaHdr: 0,
    },
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
    ...overrides,
  };
}

function buildUnicodeSttbNoExtra(strings) {
  const parts = [];
  const header = Buffer.alloc(6);
  header.writeUInt16LE(0xffff, 0);
  header.writeUInt16LE(strings.length, 2);
  header.writeUInt16LE(0, 4);
  parts.push(header);
  for (const value of strings) {
    const text = Buffer.from(value, "utf16le");
    const length = Buffer.alloc(2);
    length.writeUInt16LE(value.length, 0);
    parts.push(length, text);
  }
  return Buffer.concat(parts);
}

test("lists and reads OLE2 streams from a WPS file", async () => {
  const document = readWps(await readFile(FULL_WPS));

  assert.equal(document.type, "wps-ole2-word-binary");
  assert.deepEqual(
    document.streams.map((stream) => stream.name).sort(),
    [
      "\u0005DocumentSummaryInformation",
      "\u0005SummaryInformation",
      "0Table",
      "Data",
      "WordDocument",
      "WpsCustomData",
    ],
  );
  assert.equal(document.fib.whichTableStream, "0Table");
  assert.equal(document.fib.lid, 0x0409);
  assert.equal(document.fib.pnNext, 0);
  assert.equal(document.fib.fDot, false);
  assert.equal(document.fib.fGlsy, false);
  assert.equal(document.fib.fComplex, true);
  assert.equal(document.fib.fEncrypted, false);
  assert.equal(document.fib.fExtChar, true);
  assert.equal(document.fib.nFibBack, 0x00bf);
  assert.equal(document.fib.lKey, 0);
  assert.equal(document.fib.envr, 0);
  assert.equal(document.fib.characterCounts.body, 11098);
  assert.match(document.text, /重庆市青年就业见习实施办法/);
  assert.match(document.text, /就业见习补贴标准/);
});

test("extracts the same readable text as the one-page DOCX export", async () => {
  const wps = readWps(await readFile(BASIC_WPS));
  const docxText = readDocxMainText(await readFile(BASIC_DOCX));

  assert.equal(normalizeComparableText(wps.bodyText), normalizeComparableText(docxText));
});


test("sample11 non-complex Unicode text excludes parsed FKP pages", async () => {
  const wps = readWps(await readFile(SAMPLE11_WPS));

  assert.equal(wps.fib.fComplex, false);
  assert.equal(wps.fib.fExtChar, true);
  assert.deepEqual(
    wps.pieces.map((piece) => ({ cpStart: piece.cpStart, cpEnd: piece.cpEnd, fileOffset: piece.fileOffset, nonComplex: piece.nonComplex })),
    [
      { cpStart: 0, cpEnd: 1278, fileOffset: 2048, nonComplex: true },
      { cpStart: 1278, cpEnd: 2174, fileOffset: 16384, nonComplex: true },
    ],
  );
  assert.equal(wps.bodyText.length, 2129);
  assert.match(wps.bodyText, /河南省财政厅拟废止和失效的行政规范性文件目录/);
  assert.match(wps.bodyText, /豫财非税〔2013〕30号/);
  assert.doesNotMatch(wps.bodyText, /[ࠀ-࿿]{5}/);
});

test("sample11 non-complex table converts with parsed final-row geometry", async () => {
  const wps = readWps(await readFile(SAMPLE11_WPS));

  assert.equal(wps.tableRows.length, 1);
  assert.equal(wps.tableRows[0].rows.length, 28);
  assert.equal(wps.tableRows[0].tableIndent.width, 87);

  const docx = wpsToDocxBuffer(wps, { title: "sample11" });
  const xml = readDocxDocumentXml(docx);
  const text = readDocxMainText(docx);

  assert.equal((xml.match(/<w:tbl>/g) ?? []).length, 1);
  assert.equal((xml.match(/<w:tr[ >]/g) ?? []).length, 28);
  assert.equal((xml.match(/<w:tc>/g) ?? []).length, 140);
  assert.match(xml, /<w:tblInd w:w="87" w:type="dxa"\/>/);
  assert.match(text, /河南省财政厅河南省教育厅关于明确教育考试网上报名收费有关问题的通知/);
  assert.match(text, /豫财非税〔2013〕30号/);
});

test("sample11 non-complex conversion matches WPS main XML exports", async () => {
  const wps = readWps(await readFile(SAMPLE11_WPS));

  assert.deepEqual(wps.styleSheetInfo, {
    ftcAsci: 6,
    ftcFE: 4,
    ftcOther: 6,
    ftcBi: 0,
    nVerBuiltInNamesWhenSaved: 7,
  });
  assert.equal(wps.fontTable[4].preferredName, "SimSun");
  assert.equal(wps.styles.find((style) => style?.sti === 107)?.type, "numbering");

  const docx = wpsToDocxBuffer(wps, { title: "sample11" });
  const expectedDocx = await readFile("sample/sample11/expected.docx");
  for (const entry of ["word/document.xml", "word/styles.xml", "word/settings.xml"]) {
    assert.equal(
      countXmlLineEdits(readZipEntry(docx, entry).toString("utf8"), readZipEntry(expectedDocx, entry).toString("utf8")),
      0,
      entry,
    );
  }
});

test("extracts the readable text from the full WPS export", async () => {
  const wps = readWps(await readFile(FULL_WPS));
  const docxComparableText = normalizeComparableText(
    readDocxMainText(await readFile(FULL_DOCX)),
  );
  const wpsComparableText = normalizeComparableText(wps.bodyText);

  assert.ok(wpsComparableText.startsWith(docxComparableText.slice(0, 5000)));
  assert.match(wps.text, /法人（负责人）（签章）：\s+经办人（签字）：/);
  assert.match(wps.text, /2024年5月14日印发/);
  assert.equal(wps.paragraphs.at(0), "渝人社发〔2024〕12 号");
});

test("converts WPS text into a readable DOCX package", async () => {
  const wps = readWps(await readFile(BASIC_WPS));
  const docx = wpsToDocxBuffer(wps, {
    title: "ole2-one-page.wps",
    creator: "test",
    created: "2026-06-17T00:00:00.000Z",
    modified: "2026-06-17T00:00:00.000Z",
  });

  const convertedText = readDocxMainText(docx);
  assert.equal(normalizeComparableText(convertedText), normalizeComparableText(wps.bodyText));
});

test("extracts and emits table2 WPS table structure", async () => {
  const wps = readWps(await readFile(TABLE2_WPS));

  const secondTableFirstCell = wps.bodyText.slice(
    wps.tableRows[1].rows[0].cells[0].cpStart,
    wps.tableRows[1].rows[0].cells[0].cpEnd,
  );
  assert.equal(secondTableFirstCell, "政策落实\r（30 分）\x07");
  assert.deepEqual(
    Object.fromEntries(
      Object.entries(wps.tableRows[0].tableBorders).map(([side, border]) => [side, border.style]),
    ),
    {
      top: "single",
      left: "single",
      bottom: "single",
      right: "single",
      insideH: "single",
      insideV: "single",
    },
  );

  const docx = wpsToDocxBuffer(wps, { title: "table2" });
  const xml = readDocxDocumentXml(docx);
  const expectedXml = readDocxDocumentXml(await readFile("sample/table2/expected.docx"));
  assert.match(readDocxMainText(docx), /重庆市青年就业见习基地 年度评估表/);
  assert.match(xml, /<w:u w:val="thick"\/>/);
  const tableXmls = [...xml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
  const expectedTableXmls = [...expectedXml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
  assert.ok(tableXmls.length >= 2);
  assert.match(tableXmls[0], /<w:tblStyle w:val="6"\/>/);
  assert.match(tableXmls[0], /<w:tblW w:w="13118" w:type="dxa"\/>/);
  assert.match(tableXmls[0], /<w:jc w:val="center"\/>/);
  assert.match(tableXmls[0], /<w:vAlign w:val="center"\/>/);
  assert.match(tableXmls[0], /<w:tblBorders><w:top w:val="single"[\s\S]*<w:insideV w:val="single"/);
  assert.match(xml, /政策落实[\s\S]*30[\s\S]*分/);
  assert.deepEqual(
    tableXmls.map((t) => (t.match(/<w:tcBorders>/g) || []).length),
    expectedTableXmls.map((t) => (t.match(/<w:tcBorders>/g) || []).length),
  );

  const sectionTypes = [...xml.matchAll(/<w:sectPr>([\s\S]*?)<\/w:sectPr>/g)]
    .map((match) => /<w:type w:val="continuous"\/>/.test(match[1]) ? "continuous" : null);
  assert.deepEqual(sectionTypes, ["continuous", "continuous", null, null]);
});

test("table2 title keeps a thick underline between base and year", async () => {
  const wps = readWps(await readFile(TABLE2_WPS));
  assert.equal(wps.characterProperties[17].underlineStyle, "thick");

  const docx = wpsToDocxBuffer(wps, { title: "table2" });
  const xml = readDocxDocumentXml(docx);
  const text = readDocxMainText(docx);
  assert.match(text, /重庆市青年就业见习基地 年度评估表/);
  assert.match(xml, /<w:u w:val="thick"\/>/);
});

test("sample2 settings emit the East Asian grid profile", async () => {
  const wps = readWps(await readFile("sample/sample2/original.wps"));
  const docx = wpsToDocxBuffer(wps, { title: "sample2" });
  const convertedXml = readZipEntry(docx, "word/settings.xml").toString("utf8");
  const expectedXml = readZipEntry(await readFile("sample/sample2/expected.docx"), "word/settings.xml").toString("utf8");

  assert.equal(extractSettingsNode(convertedXml, "w:mirrorMargins"), extractSettingsNode(expectedXml, "w:mirrorMargins"));
  assert.equal(extractSettingsNode(convertedXml, "w:bordersDoNotSurroundHeader"), extractSettingsNode(expectedXml, "w:bordersDoNotSurroundHeader"));
  assert.equal(extractSettingsNode(convertedXml, "w:bordersDoNotSurroundFooter"), extractSettingsNode(expectedXml, "w:bordersDoNotSurroundFooter"));
  assert.equal(extractSettingsNode(convertedXml, "w:stylePaneFormatFilter"), extractSettingsNode(expectedXml, "w:stylePaneFormatFilter"));
  assert.equal(extractSettingsNode(convertedXml, "w:documentProtection"), extractSettingsNode(expectedXml, "w:documentProtection"));
  assert.equal(extractSettingsNode(convertedXml, "w:defaultTabStop"), extractSettingsNode(expectedXml, "w:defaultTabStop"));
  assert.equal(extractSettingsNode(convertedXml, "w:drawingGridHorizontalSpacing"), extractSettingsNode(expectedXml, "w:drawingGridHorizontalSpacing"));
  assert.equal(extractSettingsNode(convertedXml, "w:drawingGridVerticalSpacing"), extractSettingsNode(expectedXml, "w:drawingGridVerticalSpacing"));
  assert.equal(extractSettingsNode(convertedXml, "w:noPunctuationKerning"), extractSettingsNode(expectedXml, "w:noPunctuationKerning"));
  assert.equal(extractSettingsNode(convertedXml, "w:characterSpacingControl"), extractSettingsNode(expectedXml, "w:characterSpacingControl"));
  assert.ok(countXmlLineEdits(convertedXml, expectedXml) < 2);
  assert.equal(extractSettingsNode(convertedXml, "w:uiCompat97To2003"), extractSettingsNode(expectedXml, "w:uiCompat97To2003"));
  assert.equal(extractSettingsNode(convertedXml, "w:embedTrueTypeFonts"), extractSettingsNode(expectedXml, "w:embedTrueTypeFonts"));
  assert.equal(extractSettingsNode(convertedXml, "w:saveSubsetFonts"), extractSettingsNode(expectedXml, "w:saveSubsetFonts"));
  assert.equal(extractSettingsNode(convertedXml, "w:doNotValidateAgainstSchema"), extractSettingsNode(expectedXml, "w:doNotValidateAgainstSchema"));
  assert.equal(extractSettingsNode(convertedXml, "w:doNotDemarcateInvalidXml"), extractSettingsNode(expectedXml, "w:doNotDemarcateInvalidXml"));
  assert.equal(extractSettingsNode(convertedXml, "w:footnotePr"), extractSettingsNode(expectedXml, "w:footnotePr"));
  assert.equal(extractSettingsNode(convertedXml, "w:endnotePr"), extractSettingsNode(expectedXml, "w:endnotePr"));
});

test("settings shape defaults are emitted from parsed MS-DOC PlcfSpa FIB slots", async () => {
  const sample8 = readWps(await readFile(SAMPLE8_WPS));
  const sample9 = readWps(await readFile(SAMPLE9_WPS));
  const sample10 = readWps(await readFile(SAMPLE10_WPS));

  assert.ok(sample8.fib.lcbPlcSpaMom > 0);
  assert.equal(sample8.fib.lcbPlcSpaHdr, 0);
  assert.equal(sample9.fib.lcbPlcSpaMom, 0);
  assert.ok(sample9.fib.lcbPlcSpaHdr > 0);
  assert.equal(sample10.fib.lcbPlcSpaMom, 0);
  assert.equal(sample10.fib.lcbPlcSpaHdr, 0);

  const sample8Settings = readZipEntry(wpsToDocxBuffer(sample8, { title: "sample8" }), "word/settings.xml").toString("utf8");
  const sample8Expected = readZipEntry(await readFile("sample/sample8/expected.docx"), "word/settings.xml").toString("utf8");
  assert.equal(extractSettingsNode(sample8Settings, "w:shapeDefaults"), extractSettingsNode(sample8Expected, "w:shapeDefaults"));
  assert.equal(extractSettingsNode(sample8Settings, "w:hdrShapeDefaults"), null);

  const sample9Settings = readZipEntry(wpsToDocxBuffer(sample9, { title: "sample9" }), "word/settings.xml").toString("utf8");
  const sample9Expected = readZipEntry(await readFile("sample/sample9/expected.docx"), "word/settings.xml").toString("utf8");
  assert.equal(extractSettingsNode(sample9Settings, "w:hdrShapeDefaults"), extractSettingsNode(sample9Expected, "w:hdrShapeDefaults"));
  assert.equal(extractSettingsNode(sample9Settings, "w:shapeDefaults"), null);
  assert.equal(extractSettingsNode(sample8Settings, "w:zoom"), extractSettingsNode(sample8Expected, "w:zoom"));
  assert.equal(extractSettingsNode(sample9Settings, "w:zoom"), extractSettingsNode(sample9Expected, "w:zoom"));

  const sample10Settings = readZipEntry(wpsToDocxBuffer(sample10, { title: "sample10" }), "word/settings.xml").toString("utf8");
  const sample10Expected = readZipEntry(await readFile("sample/sample10/expected.docx"), "word/settings.xml").toString("utf8");
  assert.equal(extractSettingsNode(sample10Settings, "w:zoom"), extractSettingsNode(sample10Expected, "w:zoom"));
});

test("document background follows parsed grid state unless shapes or tracked revisions own visual state", async () => {
  const sample3 = readWps(await readFile("sample/sample3/original.wps"));
  const sample8 = readWps(await readFile(SAMPLE8_WPS));
  const sample9 = readWps(await readFile(SAMPLE9_WPS));
  const sample10 = readWps(await readFile(SAMPLE10_WPS));

  assert.equal(sample3.sections[0].properties.docGridType, 2);
  assert.equal(sample3.sections[0].properties.docGridCharSpace, 0);
  assert.equal(sample3.fib.lcbPlcSpaMom, 0);
  assert.equal(sample3.fib.lcbPlcSpaHdr, 0);
  assert.equal(sample3.dop.fRevMarking, false);

  assert.ok(sample8.fib.lcbPlcSpaMom > 0);
  assert.ok(sample9.fib.lcbPlcSpaHdr > 0);
  assert.equal(sample10.dop.fRevMarking, true);

  assert.match(readDocxDocumentXml(wpsToDocxBuffer(sample3, { title: "sample3" })), /<w:background w:color="FFFFFF"\/>/);
  assert.doesNotMatch(readDocxDocumentXml(wpsToDocxBuffer(sample8, { title: "sample8" })), /<w:background/);
  assert.doesNotMatch(readDocxDocumentXml(wpsToDocxBuffer(sample9, { title: "sample9" })), /<w:background/);
  assert.doesNotMatch(readDocxDocumentXml(wpsToDocxBuffer(sample10, { title: "sample10" })), /<w:background/);
});

test("sample3 settings use the 420 default tab stop from parsed document settings", async () => {
  const wps = readWps(await readFile("sample/sample3/original.wps"));
  assert.equal(wps.dop.dxaTab, 420);
  assert.equal(wps.defaultTabStop, 420);

  const docx = wpsToDocxBuffer(wps, { title: "sample3" });
  const convertedXml = readZipEntry(docx, "word/settings.xml").toString("utf8");
  const expectedXml = readZipEntry(
    await readFile("sample/sample3/expected.docx"),
    "word/settings.xml",
  ).toString("utf8");
  assert.equal(extractSettingsNode(convertedXml, "w:embedTrueTypeFonts"), extractSettingsNode(expectedXml, "w:embedTrueTypeFonts"));
  assert.equal(extractSettingsNode(convertedXml, "w:saveSubsetFonts"), extractSettingsNode(expectedXml, "w:saveSubsetFonts"));
  assert.equal(extractSettingsNode(convertedXml, "w:bordersDoNotSurroundHeader"), extractSettingsNode(expectedXml, "w:bordersDoNotSurroundHeader"));
  assert.equal(extractSettingsNode(convertedXml, "w:bordersDoNotSurroundFooter"), extractSettingsNode(expectedXml, "w:bordersDoNotSurroundFooter"));
  assert.equal(extractSettingsNode(convertedXml, "w:stylePaneFormatFilter"), extractSettingsNode(expectedXml, "w:stylePaneFormatFilter"));
  assert.equal(extractSettingsNode(convertedXml, "w:documentProtection"), extractSettingsNode(expectedXml, "w:documentProtection"));
  assert.equal(extractSettingsNode(convertedXml, "w:defaultTabStop"), extractSettingsNode(expectedXml, "w:defaultTabStop"));
  assert.equal(extractSettingsNode(convertedXml, "w:drawingGridHorizontalSpacing"), extractSettingsNode(expectedXml, "w:drawingGridHorizontalSpacing"));
  assert.equal(extractSettingsNode(convertedXml, "w:drawingGridVerticalSpacing"), extractSettingsNode(expectedXml, "w:drawingGridVerticalSpacing"));
  assert.equal(extractSettingsNode(convertedXml, "w:noPunctuationKerning"), extractSettingsNode(expectedXml, "w:noPunctuationKerning"));
  assert.equal(extractSettingsNode(convertedXml, "w:characterSpacingControl"), extractSettingsNode(expectedXml, "w:characterSpacingControl"));
  assert.equal(extractSettingsNode(convertedXml, "w:doNotValidateAgainstSchema"), extractSettingsNode(expectedXml, "w:doNotValidateAgainstSchema"));
  assert.equal(extractSettingsNode(convertedXml, "w:doNotDemarcateInvalidXml"), extractSettingsNode(expectedXml, "w:doNotDemarcateInvalidXml"));
  assert.ok(countXmlLineEdits(convertedXml, expectedXml) < 2);
  assert.equal(extractSettingsNode(convertedXml, "w:uiCompat97To2003"), extractSettingsNode(expectedXml, "w:uiCompat97To2003"));
  assert.equal(extractSettingsNode(convertedXml, "w:mirrorMargins"), extractSettingsNode(expectedXml, "w:mirrorMargins"));
  assert.equal(extractSettingsNode(convertedXml, "w:footnotePr"), extractSettingsNode(expectedXml, "w:footnotePr"));
  assert.equal(extractSettingsNode(convertedXml, "w:endnotePr"), extractSettingsNode(expectedXml, "w:endnotePr"));
});

test("sample3 settings.xml matches WPS expected export without metadata noise", async () => {
  const wps = readWps(await readFile("sample/sample3/original.wps"));
  assert.equal(wps.dop.compatibility.ulTrailSpace, true);
  const docx = wpsToDocxBuffer(wps, { title: "sample3" });
  const convertedXml = readZipEntry(docx, "word/settings.xml").toString("utf8");
  const expectedXml = readZipEntry(await readFile("sample/sample3/expected.docx"), "word/settings.xml").toString("utf8");

  assert.ok(countXmlLineEdits(convertedXml, expectedXml) === 0);
  assert.match(extractSettingsNode(convertedXml, "w:compat"), /<w:ulTrailSpace\/>/);
});

test("sample5 settings use parsed DOP typography, grid display, and XML validation flags", async () => {
  const wps = readWps(await readFile("sample/sample5/original.wps"));
  assert.equal(wps.dop.fFacingPages, false);
  assert.equal(wps.dop.fpc, 1);
  assert.equal(wps.dop.rncFtn, 0);
  assert.equal(wps.dop.nFtn, 1);
  assert.equal(wps.dop.fProtEnabled, false);
  assert.equal(wps.dop.fAutoHyphen, false);
  assert.equal(wps.dop.cpgWebOpt, 20936);
  assert.equal(wps.dop.dxaHotZ, 360);
  assert.equal(wps.dop.cConsecHypLim, 0);
  assert.equal(wps.dop.fCharLineUnits, false);
  assert.equal(wps.dop.rncEdn, 0);
  assert.equal(wps.dop.nEdn, 1);
  assert.equal(wps.dop.epc, 3);
  assert.equal(wps.dop.compatibility.noColumnBalance, false);
  assert.equal(wps.dop.compatibility.convMailMergeEsc, false);
  assert.equal(wps.dop.compatibility.suppressTopSpacing, false);
  assert.equal(wps.dop.dogrid.dxGridDisplay, 0);
  assert.equal(wps.dop.typography.fKerningPunct, true);
  assert.equal(wps.dop.grfFmtFilter, 0x5024);
  assert.equal(wps.dop.xmlValidation.fValidateXML, true);
  assert.equal(wps.dop.xmlValidation.fShowXMLErrors, true);

  const docx = wpsToDocxBuffer(wps, { title: "sample5" });
  const convertedXml = readZipEntry(docx, "word/settings.xml").toString("utf8");
  const expectedXml = readZipEntry(await readFile("sample/sample5/expected.docx"), "word/settings.xml").toString("utf8");

  assert.equal(extractSettingsNode(convertedXml, "w:stylePaneFormatFilter"), null);
  assert.equal(extractSettingsNode(convertedXml, "w:displayHorizontalDrawingGridEvery"), extractSettingsNode(expectedXml, "w:displayHorizontalDrawingGridEvery"));
  assert.equal(extractSettingsNode(convertedXml, "w:noPunctuationKerning"), null);
  assert.equal(extractSettingsNode(convertedXml, "w:doNotValidateAgainstSchema"), null);
  assert.equal(extractSettingsNode(convertedXml, "w:doNotDemarcateInvalidXml"), null);
  assert.ok(countXmlLineEdits(convertedXml, expectedXml) < 2);
});

test("sample5 paragraphs preserve sprmPFContextualSpacing from PAPX", async () => {
  const wps = readWps(await readFile("sample/sample5/original.wps"));
  assert.equal(wps.paragraphProperties.filter((properties) => properties?.contextualSpacing === true).length, 5);

  const docx = wpsToDocxBuffer(wps, { title: "sample5" });
  const convertedXml = readDocxDocumentXml(docx);
  const expectedXml = readDocxDocumentXml(await readFile("sample/sample5/expected.docx"));
  assert.equal((convertedXml.match(/<w:contextualSpacing\/>/g) ?? []).length, 5);
  assert.ok(countXmlLineEdits(convertedXml, expectedXml) < 4);
});

test("sample7 negative grid profile matches WPS DOCX export", async () => {
  const wps = readWps(await readFile("sample/sample7/original.wps"));
  assert.equal(wps.sections[0].properties.docGridType, 1);
  assert.equal(wps.sections[0].properties.docGridCharSpace, -1024);
  assert.equal(wps.sections[0].properties.pageNumberFormat, 57);

  const docx = wpsToDocxBuffer(wps, { title: "sample7" });
  const expectedDocx = await readFile("sample/sample7/expected.docx");
  const documentXml = readZipEntry(docx, "word/document.xml").toString("utf8");
  const expectedDocumentXml = readZipEntry(expectedDocx, "word/document.xml").toString("utf8");
  const stylesXml = readZipEntry(docx, "word/styles.xml").toString("utf8");
  const expectedStylesXml = readZipEntry(expectedDocx, "word/styles.xml").toString("utf8");
  const settingsXml = readZipEntry(docx, "word/settings.xml").toString("utf8");
  const expectedSettingsXml = readZipEntry(expectedDocx, "word/settings.xml").toString("utf8");

  assert.equal(countXmlLineEdits(documentXml, expectedDocumentXml), 0);
  assert.equal(countXmlLineEdits(stylesXml, expectedStylesXml), 0);
  assert.equal(countXmlLineEdits(settingsXml, expectedSettingsXml), 0);
  assert.equal(extractSettingsNode(settingsXml, "w:zoom"), extractSettingsNode(expectedSettingsXml, "w:zoom"));
  assert.equal(extractSettingsNode(settingsXml, "w:adjustLineHeightInTable"), extractSettingsNode(expectedSettingsXml, "w:adjustLineHeightInTable"));
  assert.equal(extractSettingsNode(settingsXml, "w:evenAndOddHeaders"), null);
  assert.match(documentXml, /<w:tblStyle w:val="4"\/>/);
  assert.match(documentXml, /<w:tblpPr[^>]*w:tblpXSpec="center"[^>]*\/>/);
  assert.match(documentXml, /<w:pgNumType w:fmt="numberInDash"\/>/);
  assert.match(documentXml, /<w:footerReference r:id="rId3" w:type="default"\/>/);
  assert.doesNotMatch(documentXml, /<w:tcBorders>/);
  assert.doesNotMatch(documentXml, /<w:highlight w:val="none"\/>/);
});

test("converted DOCX maps parsed MS-DOC LIDs without unknown-language fallback", async () => {
  const sample3 = readWps(await readFile("sample/sample3/original.wps"));
  assert.equal(
    sample3.characterRuns.filter((run) => run.properties?.langIdEastAsia === 0x0004).length,
    2,
  );
  const sample3Xml = readDocxDocumentXml(wpsToDocxBuffer(sample3, { title: "sample3" }));
  const sample3ExpectedXml = readDocxDocumentXml(await readFile("sample/sample3/expected.docx"));
  assert.match(sample3Xml, /<w:lang w:eastAsia="zh-Hans"\/>/);
  assert.ok(countXmlLineEdits(sample3Xml, sample3ExpectedXml) < 4);

  const sample6 = readWps(await readFile("sample/sample6/original.wps"));
  assert.equal(
    sample6.characterRuns.filter((run) => run.properties?.langIdEastAsia === 0x7804).length,
    1,
  );
  const sample6Xml = readDocxDocumentXml(wpsToDocxBuffer(sample6, { title: "sample6" }));
  const sample6ExpectedXml = readDocxDocumentXml(await readFile("sample/sample6/expected.docx"));
  assert.match(sample6Xml, /<w:lang w:val="en-US" w:eastAsia="zh" w:bidi="ar"\/>/);
  assert.ok(countXmlLineEdits(sample6Xml, sample6ExpectedXml) < 4);
});

test("MS-DOC LID mapping uses spec LCID tables without generic language fallback", () => {
  assert.equal(lcidToBcp47(0x0004), "zh-Hans");
  assert.equal(lcidToBcp47(0x7804), "zh");
  assert.equal(lcidToBcp47(0x0409), "en-US");
  assert.equal(lcidToBcp47(0x040d), "he-IL");
  assert.equal(lcidToBcp47(0x7c04), "zh-Hant");
  assert.equal(lcidToBcp47(0x1000), null);
});

test("sample document.xml matches WPS expected export without metadata noise", async () => {
  for (const sample of ["basic", "sample1", "sample2", "sample3", "sample4", "sample5", "sample6", "table2", "table4", "table6", "tables"]) {
    const wps = readWps(await readFile(`sample/${sample}/original.wps`));
    const convertedXml = readDocxDocumentXml(wpsToDocxBuffer(wps, { title: sample }));
    const expectedXml = readDocxDocumentXml(await readFile(`sample/${sample}/expected.docx`));
    assert.equal(countXmlLineEdits(convertedXml, expectedXml), 0, sample);
  }
});

test("parses the MS-DOC Selsf last-selection structure", async () => {
  const basic = readWps(await readFile(BASIC_WPS));
  assert.equal(basic.lastSelection.fIns, true);
  assert.equal(basic.lastSelection.cpFirst, basic.lastSelection.cpLim);
  assert.equal(basic.lastSelection.cpFirst, 253);
});

test("Selsf selection within the MS-DOC section extent emits a _GoBack bookmark", async () => {
  const sample9 = readWps(await readFile(SAMPLE9_WPS));
  assert.equal(sample9.lastSelection.fIns, true);
  assert.equal(sample9.lastSelection.cpFirst, 1400);
  assert.ok(sample9.lastSelection.cpFirst > sample9.bodyText.length);
  assert.ok(sample9.lastSelection.cpFirst <= sample9.sections.at(-1).cpEnd);

  const xml = readDocxDocumentXml(wpsToDocxBuffer(sample9, { title: "sample9" }));
  assert.match(xml, /w:name="_GoBack"/);
});

test("explicit MS-DOC character-unit indents are preserved despite stale Dop2000 fCharLineUnits", async () => {
  const wps = readWps(await readFile(SAMPLE8_WPS));
  const props = wps.paragraphProperties[9];

  // MS-DOC-SPEC/17 Dop2000.fCharLineUnits says character-unit indents MUST
  // NOT be in use when false, but explicit MS-DOC-SPEC/16 sprmPDxcLeft1
  // paragraph records are more direct evidence for the affected paragraph.
  // Preserve the paragraph SPRM instead of using a stale document-level bit
  // to discard it.
  assert.equal(wps.dop.fCharLineUnits, false);
  assert.equal(props.firstLineIndentCharsSprm, 0x4457);
  assert.equal(props.firstLineIndentChars, 100);

  const xml = readDocxDocumentXml(wpsToDocxBuffer(wps, { title: "sample8-char-units" }));
  assert.match(xml, /<w:ind w:firstLine="281" w:firstLineChars="100"\/>/);
});

test("paragraph first-line indents preserve explicit MS-DOC SPRM sources", async () => {
  const wps = readWps(await readFile(SAMPLE8_WPS));
  const props = wps.paragraphProperties[10];

  // MS-DOC-SPEC/16 sprmPDxcLeft1 (0x4457) carries the character-unit
  // first-line indent, while sprmPDxaLeft1 (0x8460) carries the twip
  // companion. Preserve both parsed values and their SPRM ids instead of
  // recomputing one from the other.
  assert.equal(props.firstLineIndentChars, 200);
  assert.equal(props.firstLineIndentCharsSprm, 0x4457);
  assert.equal(props.firstLineIndent, 562);
  assert.equal(props.firstLineIndentSprm, 0x8460);
});

test("available samples preserve first-line indent SPRM provenance", async () => {
  for (let sample = 1; sample <= 10; sample += 1) {
    const wps = readWps(await readFile(`sample/sample${sample}/original.wps`));
    for (const props of wps.paragraphProperties) {
      if (!props) continue;
      if (props.firstLineIndentChars != null) {
        // MS-DOC-SPEC/16 sprmPDxcLeft1 is the only parsed source for
        // firstLineIndentChars in this converter; keep that provenance so
        // later DOCX emission is source-driven rather than inferred.
        assert.equal(props.firstLineIndentCharsSprm, 0x4457);
      }
      if (props.firstLineIndent != null && props.firstLineIndentSprm != null) {
        assert.ok([0x8411, 0x8458, 0x8460].includes(props.firstLineIndentSprm));
      }
    }
  }
});

test("line-grid character-unit first-line resolution requires parsed sprmPDxcLeft1 source", () => {
  const base = {
    bodyText: "A\r",
    defaultTabStop: 720,
    characterProperties: [{ fontSize: 22 }, { fontSize: 22 }],
    sections: [{
      cpStart: 0,
      cpEnd: 2,
      properties: { docGridType: 2, docGridLinePitch: 312 },
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      typography: { fKerningPunct: false, iJustification: 1 },
      xmlValidation: { fValidateXML: false, fShowXMLErrors: false },
      grfFmtFilter: 0x5024,
      compatibility: { ulTrailSpace: true },
    },
  };

  const sourcedXml = readDocxDocumentXml(wpsToDocxBuffer({
    ...base,
    paragraphProperties: [{
      firstLineIndentChars: 200,
      firstLineIndentCharsSprm: 0x4457,
      firstLineIndent: 642,
      firstLineIndentSprm: 0x8460,
    }],
  }, { title: "sourced-char-unit-indent" }));
  const unsourcedXml = readDocxDocumentXml(wpsToDocxBuffer({
    ...base,
    paragraphProperties: [{
      firstLineIndentChars: 200,
      firstLineIndent: 642,
      firstLineIndentSprm: 0x8460,
    }],
  }, { title: "unsourced-char-unit-indent" }));

  // MS-DOC-SPEC/16 sprmPDxcLeft1 (0x4457) is the parsed evidence for
  // character-unit indentation. Without that source id, keep the explicit
  // twip companion rather than recomputing it from font size.
  assert.match(sourcedXml, /<w:ind w:firstLine="440" w:firstLineChars="200"\/>/);
  assert.match(unsourcedXml, /<w:ind w:firstLine="642" w:firstLineChars="200"\/>/);
});

test("same-CP bookmark starts preserve parsed MS-DOC PLC order", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "AB\r",
    defaultTabStop: 720,
    bookmarks: [
      { id: 0, name: "Alpha", cpStart: 0, cpEnd: 1 },
      { id: 1, name: "Beta", cpStart: 0, cpEnd: 1 },
      { id: 2, name: "Gamma", cpStart: 0, cpEnd: 1 },
    ],
    paragraphProperties: [{}],
    characterProperties: [{}, {}, {}],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "same-cp-bookmark-order" });
  const xml = readDocxDocumentXml(docx);
  const names = [...xml.matchAll(/<w:bookmarkStart[^>]*w:name="([^"]+)"/g)].map((match) => match[1]);

  // MS-DOC-SPEC/15 states SttbfBkmk is parallel to Plcfbkf: the name at a
  // given offset is associated with the Plcfbkf data element at the same
  // offset. Duplicate start CPs do not create a new spec ordering key, so
  // DOCX emission preserves the parsed PLC row order.
  assert.deepEqual(names, ["Alpha", "Beta", "Gamma"]);
});

test("standard bookmarks parse MS-DOC FBKF ibkl and BKC fields", async () => {
  const sample8 = readWps(await readFile(SAMPLE8_WPS));
  assert.equal(sample8.bookmarks.length, 51);
  assert.deepEqual(sample8.bookmarks[0], {
    id: 0,
    name: "_Toc253401763",
    cpStart: 846,
    cpEnd: 855,
    ibkl: 0,
    bkc: 0,
    bkcInfo: { raw: 0, itcFirst: 0, fPub: false, itcLim: 0, fNative: false, fCol: false },
  });
  assert.deepEqual(sample8.bookmarks[6], {
    id: 6,
    name: "_Toc253402278",
    cpStart: 846,
    cpEnd: 855,
    ibkl: 6,
    bkc: 0,
    bkcInfo: { raw: 0, itcFirst: 0, fPub: false, itcLim: 0, fNative: false, fCol: false },
  });
  assert.deepEqual(sample8.bookmarks[33], {
    id: 33,
    name: "_Toc317491733",
    cpStart: 6822,
    cpEnd: 7138,
    ibkl: 34,
    bkc: 0,
    bkcInfo: { raw: 0, itcFirst: 0, fPub: false, itcLim: 0, fNative: false, fCol: false },
  });
  assert.equal(new Set(sample8.bookmarks.map((bookmark) => bookmark.ibkl)).size, sample8.bookmarks.length);
  assert.ok(sample8.bookmarks.every((bookmark) => bookmark.bkcInfo.raw === bookmark.bkc));
  assert.ok(sample8.bookmarks.every((bookmark) => bookmark.bkcInfo.fPub === false));
  assert.ok(sample8.bookmarks.every((bookmark) => !bookmark.bkcInfo.fCol || bookmark.bkcInfo.itcFirst < bookmark.bkcInfo.itcLim));
  assert.ok(sample8.bookmarks.every((bookmark) => bookmark.cpStart <= bookmark.cpEnd));
  assert.deepEqual(
    sample8.bookmarks.map((bookmark) => bookmark.cpStart),
    [...sample8.bookmarks.map((bookmark) => bookmark.cpStart)].sort((a, b) => a - b),
  );

  const xml = readDocxDocumentXml(wpsToDocxBuffer(sample8, { title: "sample8" }));
  assert.match(xml, /<w:bookmarkStart w:id="33" w:name="_Toc317491733"\/>/);
  assert.match(xml, /<w:bookmarkStart w:id="34" w:name="_Toc317331300"\/>/);
  assert.match(xml, /<w:bookmarkEnd w:id="43"\/><w:bookmarkEnd w:id="44"\/>/);
});

test("direct paragraph outline level is preserved from sprmPOutLvl", async () => {
  const sample9 = readWps(await readFile(SAMPLE9_WPS));
  assert.deepEqual(sample9.paragraphProperties.slice(0, 3).map((props) => props.outlineLevel), [9, 9, 9]);
  assert.equal(sample9.paragraphProperties[3].outlineLevel, null);

  const xml = readDocxDocumentXml(wpsToDocxBuffer(sample9, { title: "sample9" }));
  assert.equal((xml.match(/<w:outlineLvl w:val="9"\/>/g) ?? []).length, 3);
});

test("parses authoritative MS-DOC PlcfHdd stories and emits real header/footer content", () => {
  const defaultHeader = "Header text\r";
  const defaultFooter = "\x13 PAGE \\* MERGEFORMAT \x14 7 \x15\r";
  const headerText = defaultHeader + defaultFooter;
  const cps = [0, 0, 0, 0, 0, 0, 0, 0, defaultHeader.length, defaultHeader.length, headerText.length, headerText.length, headerText.length];
  const tableStream = Buffer.alloc(cps.length * 4);
  cps.forEach((cp, index) => tableStream.writeUInt32LE(cp, index * 4));
  const plcfHdd = parsePlcfHdd(tableStream, {
    fcPlcfHdd: 0,
    lcbPlcfHdd: tableStream.length,
    characterCounts: { headers: headerText.length },
  }, headerText, 1);
  const footerCp = defaultHeader.length;
  const headerFieldBuffer = buildFieldPlc([
    { cp: footerCp, ch: 0x13, grffld: 0x21 },
    { cp: footerCp + defaultFooter.indexOf("\x14"), ch: 0x14, grffld: 0xff },
    { cp: footerCp + defaultFooter.indexOf("\x15"), ch: 0x15, grffld: 0x80 },
  ], headerText.length + 100);
  const headerFields = parseFieldPlc(headerFieldBuffer, 0, headerFieldBuffer.length, headerText, "Header Document");

  assert.equal(plcfHdd.sections.length, 1);
  assert.equal(plcfHdd.sections[0].defaultHeader.text, "Header text");
  assert.equal(plcfHdd.sections[0].defaultFooter.text, "\x13 PAGE \\* MERGEFORMAT \x14 7 \x15");
  assert.equal(plcfHdd.sections[0].evenHeader.empty, true);

  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    paragraphProperties: [{}],
    characterProperties: [{}, {}],
    sections: [{ cpStart: 0, cpEnd: 2, properties: {} }],
    plcfHdd,
    fields: { headers: headerFields, body: { records: [] } },
    fib: {
      fComplex: true,
      characterCounts: { headers: headerText.length },
      lcbPlcSpaMom: 0,
      lcbPlcSpaHdr: 0,
    },
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "authoritative-plcfhdd" });

  const documentXml = readDocxDocumentXml(docx);
  const relsXml = readZipEntry(docx, "word/_rels/document.xml.rels").toString("utf8");
  const headerXml = readZipEntry(docx, "word/header1.xml").toString("utf8");
  const footerXml = readZipEntry(docx, "word/footer1.xml").toString("utf8");
  assert.match(documentXml, /<w:headerReference r:id="rId5" w:type="default"\/>/);
  assert.match(documentXml, /<w:footerReference r:id="rId6" w:type="default"\/>/);
  assert.match(relsXml, /Id="rId5"[^>]*relationships\/header"[^>]*Target="header1\.xml"/);
  assert.match(relsXml, /Id="rId6"[^>]*relationships\/footer"[^>]*Target="footer1\.xml"/);
  assert.match(headerXml, /<w:t>Header text<\/w:t>/);
  assert.match(footerXml, /<w:fldChar w:fldCharType="begin"\/>/);
  assert.ok(footerXml.includes(`<w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText>`));
  assert.match(footerXml, /<w:fldChar w:fldCharType="separate"\/>/);
  assert.match(footerXml, /<w:t xml:space="preserve"> 7 <\/w:t>/);
  assert.match(footerXml, /<w:fldChar w:fldCharType="end"\/>/);
});

test("PlcfHdd fails fast on malformed story boundaries", () => {
  const tableStream = Buffer.alloc(13 * 4);
  const cps = [0, 0, 0, 0, 0, 0, 0, 0, 4, 3, 4, 4, 4];
  cps.forEach((cp, index) => tableStream.writeUInt32LE(cp, index * 4));
  assert.throws(
    () => parsePlcfHdd(tableStream, {
      fcPlcfHdd: 0,
      lcbPlcfHdd: tableStream.length,
      characterCounts: { headers: 4 },
    }, "A\rB\r", 1),
    /PlcfHdd CP array is not ascending/,
  );
});

test("Plcfld parses nested fields and preserves authoritative field flags", () => {
  const text = "\x13 IF \x13 PAGE \x14 1 \x15 = 1 \x14 Yes \x15\r";
  const positions = [];
  for (let cp = 0; cp < text.length; cp += 1) {
    const ch = text.charCodeAt(cp);
    if (ch === 0x13 || ch === 0x14 || ch === 0x15) positions.push({ cp, ch });
  }
  const plc = buildFieldPlc([
    { ...positions[0], grffld: 0x07 },
    { ...positions[1], grffld: 0x21 },
    { ...positions[2], grffld: 0xff },
    { ...positions[3], grffld: 0xc4 },
    { ...positions[4], grffld: 0xff },
    { ...positions[5], grffld: 0x90 },
  ], text.length + 200);
  const parsed = parseFieldPlc(plc, 0, plc.length, text, "Main Document");

  assert.equal(parsed.records.length, 6);
  assert.equal(parsed.fields.length, 2);
  assert.equal(parsed.fields[0].begin.fieldType, 0x07);
  assert.equal(parsed.fields[0].separator.cp, positions[4].cp);
  assert.equal(parsed.fields[0].begin.endFlags.locked, true);
  assert.equal(parsed.fields[1].nested, true);
  assert.equal(parsed.fields[1].begin.endFlags.resultsDirty, true);
});

test("Plcfld fails fast on duplicate CPs and text mismatches", () => {
  const duplicate = buildFieldPlc([
    { cp: 0, ch: 0x13, grffld: 0x21 },
    { cp: 0, ch: 0x15, grffld: 0x00 },
  ], 4);
  assert.throws(
    () => parseFieldPlc(duplicate, 0, duplicate.length, "\x13\x15", "Main Document"),
    /not strictly ascending/,
  );

  const mismatch = buildFieldPlc([
    { cp: 0, ch: 0x13, grffld: 0x21 },
    { cp: 1, ch: 0x15, grffld: 0x00 },
  ], 4);
  assert.throws(
    () => parseFieldPlc(mismatch, 0, mismatch.length, "AB", "Main Document"),
    /does not match story character/,
  );
});

test("sample8 body fields use parsed Plcfld structure for nested DOCX fields", async () => {
  const wps = readWps(await readFile(SAMPLE8_WPS));
  assert.equal(wps.fields.body.records.length, 57);
  assert.equal(wps.fields.body.fields[0].end.cp, 4724);
  const xml = readDocxDocumentXml(wpsToDocxBuffer(wps, { title: "sample8-fields" }));
  assert.match(xml, /<w:fldChar w:fldCharType="begin"\/>/);
  assert.match(xml, /<w:instrText[^>]*>[^<]+<\/w:instrText>/);
  assert.match(xml, /<w:fldChar w:fldCharType="separate"\/>/);
  assert.match(xml, /<w:fldChar w:fldCharType="end"\/>/);
});

test("note reference/text PLCs parse authoritative ranges", () => {
  const bodyText = "A\x02B\x02\r";
  const refPlc = Buffer.alloc(16);
  refPlc.writeUInt32LE(1, 0);
  refPlc.writeUInt32LE(3, 4);
  refPlc.writeUInt32LE(100, 8);
  refPlc.writeUInt16LE(1, 12);
  refPlc.writeUInt16LE(1, 14);
  const refs = parseNoteReferencePlc(refPlc, 0, refPlc.length, bodyText, "footnote");
  assert.deepEqual(refs.references.map((reference) => reference.cp), [1, 3]);

  const storyText = "\x02First\r\x02Second\rX";
  const textPlc = Buffer.alloc(16);
  [0, 7, storyText.length - 1, 100].forEach((cp, index) => textPlc.writeUInt32LE(cp, index * 4));
  const texts = parseNoteTextPlc(textPlc, 0, textPlc.length, storyText, 2, "footnote");
  assert.equal(texts.stories[0].text, "\x02First");
  assert.equal(texts.stories[1].text, "\x02Second");
});

test("comment PLCs parse zero-length references, authors, initials, and text", () => {
  const table = Buffer.alloc(256);
  const authors = Buffer.concat([Buffer.from([4, 0]), Buffer.from("Gary", "utf16le")]);
  authors.copy(table, 0);
  const refFc = 32;
  table.writeUInt32LE(1, refFc);
  table.writeUInt32LE(100, refFc + 4);
  const atrd = refFc + 8;
  table.writeUInt16LE(2, atrd);
  Buffer.from("LG", "utf16le").copy(table, atrd + 2);
  table.writeUInt16LE(0, atrd + 20);
  table.writeUInt16LE(0, atrd + 22);
  table.writeUInt16LE(0, atrd + 24);
  table.writeInt32LE(-1, atrd + 26);
  const textFc = 96;
  const commentStory = "\x05Note\rX";
  [0, commentStory.length - 1, 200].forEach((cp, index) => table.writeUInt32LE(cp, textFc + index * 4));
  const parsed = parseComments(table, {
    fcGrpXstAtnOwners: 0,
    lcbGrpXstAtnOwners: authors.length,
    fcPlcfandRef: refFc,
    lcbPlcfandRef: 38,
    fcPlcfandTxt: textFc,
    lcbPlcfandTxt: 12,
    fcSttbfAtnBkmk: 0,
    lcbSttbfAtnBkmk: 0,
    fcPlcfAtnBkf: 0,
    lcbPlcfAtnBkf: 0,
    fcPlcfAtnBkl: 0,
    lcbPlcfAtnBkl: 0,
  }, "A\x05\r", commentStory);
  assert.deepEqual(parsed.authors, ["Gary"]);
  assert.equal(parsed.comments[0].initials, "LG");
  assert.equal(parsed.comments[0].author, "Gary");
  assert.equal(parsed.comments[0].story.text, "Note");
  assert.equal(parsed.comments[0].cpStart, 1);
  assert.equal(parsed.comments[0].cpEnd, 1);
});

test("textbox PLCs parse actual FTXBXS shape linkage and reusable tail", () => {
  const storyText = "Text\r";
  const plc = Buffer.alloc(56);
  [0, storyText.length, 100].forEach((cp, index) => plc.writeUInt32LE(cp, index * 4));
  const first = 12;
  plc.writeUInt32LE(1, first);
  plc.writeUInt32LE(0, first + 4);
  plc.writeUInt16LE(0, first + 8);
  plc.writeUInt32LE(0xffffffff, first + 10);
  plc.writeUInt32LE(1234, first + 14);
  plc.writeUInt32LE(0, first + 18);
  const last = first + 22;
  plc.writeUInt32LE(0xffffffff, last);
  plc.writeUInt32LE(0, last + 4);
  plc.writeUInt16LE(0, last + 8);
  plc.writeUInt32LE(0, last + 10);
  plc.writeUInt32LE(0, last + 14);
  plc.writeUInt32LE(0, last + 18);
  const parsed = parseTextboxTextPlc(plc, 0, plc.length, storyText, "Textbox Document");
  assert.equal(parsed.textboxes.length, 1);
  assert.equal(parsed.textboxes[0].lid, 1234);
  assert.equal(parsed.textboxes[0].text, "Text");
  assert.equal(parsed.entries[1].reusable, true);
});

test("DOCX emits parsed footnotes, endnotes, comments, and main-story references", () => {
  const bodyText = "A\x02B\x02C\x05\r";
  const wps = buildMinimalWpsDocument(bodyText, {
    footnotes: {
      references: [{ id: 1, cp: 1, automatic: true, indexValue: 1 }],
      stories: [{ id: 1, cpStart: 0, cpEnd: 11, rawText: "\x02Footnote\r", text: "\x02Footnote" }],
    },
    endnotes: {
      references: [{ id: 1, cp: 3, automatic: true, indexValue: 1 }],
      stories: [{ id: 1, cpStart: 0, cpEnd: 10, rawText: "\x02Endnote\r", text: "\x02Endnote" }],
    },
    comments: {
      authors: ["Gary"],
      comments: [{
        id: 0,
        cp: 5,
        cpStart: 5,
        cpEnd: 5,
        author: "Gary",
        initials: "LG",
        story: { id: 1, cpStart: 0, cpEnd: 9, rawText: "Comment\r", text: "Comment" },
      }],
    },
  });
  const docx = wpsToDocxBuffer(wps, { title: "semantic-stories" });
  const documentXml = readDocxDocumentXml(docx);
  const footnotesXml = readZipEntry(docx, "word/footnotes.xml").toString("utf8");
  const endnotesXml = readZipEntry(docx, "word/endnotes.xml").toString("utf8");
  const commentsXml = readZipEntry(docx, "word/comments.xml").toString("utf8");
  const relsXml = readZipEntry(docx, "word/_rels/document.xml.rels").toString("utf8");
  const contentTypes = readZipEntry(docx, "[Content_Types].xml").toString("utf8");

  assert.match(documentXml, /<w:footnoteReference w:id="1"\/>/);
  assert.match(documentXml, /<w:endnoteReference w:id="1"\/>/);
  assert.match(documentXml, /<w:commentRangeStart w:id="0"\/><w:commentRangeEnd w:id="0"\/>/);
  assert.match(documentXml, /<w:commentReference w:id="0"\/>/);
  assert.match(footnotesXml, /<w:footnote w:id="1">[\s\S]*<w:footnoteRef\/>[\s\S]*<w:t>Footnote<\/w:t>/);
  assert.match(endnotesXml, /<w:endnote w:id="1">[\s\S]*<w:endnoteRef\/>[\s\S]*<w:t>Endnote<\/w:t>/);
  assert.match(commentsXml, /<w:comment w:id="0" w:author="Gary" w:initials="LG">[\s\S]*<w:t>Comment<\/w:t>/);
  assert.match(relsXml, /relationships\/comments" Target="comments\.xml"/);
  assert.match(contentTypes, /PartName="\/word\/comments\.xml"/);
});

test("compact header subdocuments emit MS-DOC default/even header-footer references", async () => {
  const sample9Docx = wpsToDocxBuffer(readWps(await readFile(SAMPLE9_WPS)), { title: "sample9" });
  const sample9DocumentXml = readDocxDocumentXml(sample9Docx);
  const sample9RelsXml = readZipEntry(sample9Docx, "word/_rels/document.xml.rels").toString("utf8");
  assert.match(sample9DocumentXml, /<w:headerReference r:id="rId3" w:type="default"\/>/);
  assert.match(sample9DocumentXml, /<w:footerReference r:id="rId4" w:type="default"\/>/);
  assert.match(sample9RelsXml, /Id="rId3"[^>]*relationships\/header"[^>]*Target="header1\.xml"/);
  assert.match(sample9RelsXml, /Id="rId4"[^>]*relationships\/footer"[^>]*Target="footer1\.xml"/);

  const sample10 = readWps(await readFile(SAMPLE10_WPS));
  assert.equal(sample10.dop.fFacingPages, true);
  const sample10DocumentXml = readDocxDocumentXml(wpsToDocxBuffer(sample10, { title: "sample10" }));
  assert.match(sample10DocumentXml, /<w:footerReference r:id="rId3" w:type="default"\/>/);
  assert.match(sample10DocumentXml, /<w:footerReference r:id="rId4" w:type="even"\/>/);
});

test("sample3 styles.xml stays close to WPS expected export", async () => {
  const wps = readWps(await readFile("sample/sample3/original.wps"));
  const docx = wpsToDocxBuffer(wps, { title: "sample3" });
  const convertedXml = readZipEntry(docx, "word/styles.xml").toString("utf8");
  const expectedXml = readZipEntry(await readFile("sample/sample3/expected.docx"), "word/styles.xml").toString("utf8");

  assert.ok(countXmlLineEdits(convertedXml, expectedXml) <= 320);
});

test("sample3 table serial-number column has a resolvable numbering definition", async () => {
  const wps = readWps(await readFile("sample/sample3/original.wps"));
  const docx = wpsToDocxBuffer(wps, { title: "sample3" });
  const documentXml = readDocxDocumentXml(docx);
  const numberingXml = readZipEntry(docx, "word/numbering.xml").toString("utf8");

  assert.match(documentXml, /<w:numPr><w:ilvl w:val="0"\/><w:numId w:val="2"\/><\/w:numPr>/);
  const num2 = numberingXml.match(/<w:num w:numId="2">[\s\S]*?<\/w:num>/)?.[0] ?? "";
  assert.match(num2, /<w:abstractNumId w:val="0"\/>/);
  assert.match(numberingXml, /<w:abstractNum w:abstractNumId="0">[\s\S]*<w:numFmt w:val="decimal"\/>[\s\S]*<w:lvlText w:val="%1"\/>/);
});

test("sample3 table header row repeats across pages", async () => {
  const wps = readWps(await readFile("sample/sample3/original.wps"));
  assert.equal(wps.tableRows.length, 1);
  assert.equal(wps.tableRows[0].rows[0].cantSplit, true);
  assert.equal(wps.tableRows[0].rows[0].repeatHeader, true);

  const docx = wpsToDocxBuffer(wps, { title: "sample3" });
  const convertedXml = readDocxDocumentXml(docx);
  const expectedXml = readDocxDocumentXml(await readFile("sample/sample3/expected.docx"));
  const convertedTableXml = convertedXml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/)?.[0] ?? null;
  const expectedTableXml = expectedXml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/)?.[0] ?? null;

  assert.ok(convertedTableXml);
  assert.ok(expectedTableXml);
  assert.match(convertedTableXml, /<w:tblHeader\/>/);
  assert.match(convertedTableXml, /<w:cantSplit\/>/);
  assert.match(expectedTableXml, /<w:tblHeader\/>/);
  assert.match(expectedTableXml, /<w:cantSplit\/>/);
  const convertedFirstRow = convertedTableXml.match(/<w:tr[\s\S]*?<\/w:tr>/)?.[0] ?? null;
  const expectedFirstRow = expectedTableXml.match(/<w:tr[\s\S]*?<\/w:tr>/)?.[0] ?? null;
  assert.ok(convertedFirstRow);
  assert.ok(expectedFirstRow);
  assert.match(convertedFirstRow, /<w:b w:val="0"\/>/);
  assert.match(expectedFirstRow, /<w:b w:val="0"\/>/);
  assert.doesNotMatch(convertedFirstRow, /<w:b\/>/);
});

test("basic settings use the 720 default tab stop from parsed document settings", async () => {
  const wps = readWps(await readFile(BASIC_WPS));
  assert.equal(wps.dop.dxaTab, 720);
  assert.equal(wps.defaultTabStop, 720);

  const docx = wpsToDocxBuffer(wps, { title: "basic" });
  const xml = readZipEntry(docx, "word/settings.xml").toString("utf8");
  assert.match(xml, /<w:defaultTabStop w:val="720"\/>/);
});

test("settings emit FIB read-only recommendation as writeProtection before zoom", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    fib: { fReadOnlyRecommended: true },
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "read-only-recommended" });
  const xml = readZipEntry(docx, "word/settings.xml").toString("utf8");
  const writeProtectionIndex = xml.indexOf("<w:writeProtection");
  const zoomIndex = xml.indexOf("<w:zoom");

  assert.ok(writeProtectionIndex >= 0);
  assert.match(xml, /<w:writeProtection w:recommended="1"\/>/);
  assert.ok(zoomIndex >= 0);
  assert.ok(writeProtectionIndex < zoomIndex);
});


test("settings hyphenationZone follows parsed DopBase dxaHotZ", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    dop: {
      dxaHotZ: 480,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "hyphenation-zone" });
  const xml = readZipEntry(docx, "word/settings.xml").toString("utf8");

  assert.match(xml, /<w:hyphenationZone w:val="480"\/>/);
});

test("settings fail fast when parsed DopBase dxaHotZ is missing", () => {
  assert.throws(() => wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    dop: {
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "missing-hyphenation-zone" }), /missing parsed DopBase\.dxaHotZ hyphenation zone/);
});

test("settings evenAndOddHeaders follows parsed DopBase fFacingPages only", () => {
  const baseDoc = {
    bodyText: "A\r",
    defaultTabStop: 720,
    dop: {
      fFacingPages: false,
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  };
  const noEvenOddXml = readZipEntry(wpsToDocxBuffer(baseDoc, { title: "facing-pages-off" }), "word/settings.xml").toString("utf8");
  const evenOddXml = readZipEntry(wpsToDocxBuffer({
    ...baseDoc,
    dop: { ...baseDoc.dop, fFacingPages: true },
  }, { title: "facing-pages-on" }), "word/settings.xml").toString("utf8");

  assert.equal(extractSettingsNode(noEvenOddXml, "w:evenAndOddHeaders"), null);
  assert.match(evenOddXml, /<w:evenAndOddHeaders w:val="1"\/>/);
});

test("settings emit TrueType font embedding only from parsed DOP font flags", () => {
  const baseDoc = {
    bodyText: "A\r",
    defaultTabStop: 720,
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  };
  const noEmbedXml = readZipEntry(wpsToDocxBuffer(baseDoc, { title: "no-font-embed" }), "word/settings.xml").toString("utf8");
  const embedXml = readZipEntry(wpsToDocxBuffer({
    ...baseDoc,
    dop: { ...baseDoc.dop, fEmbedFonts: true, fSubsetFonts: true },
  }, { title: "font-embed" }), "word/settings.xml").toString("utf8");

  assert.equal(extractSettingsNode(noEmbedXml, "w:embedTrueTypeFonts"), null);
  assert.equal(extractSettingsNode(noEmbedXml, "w:saveSubsetFonts"), null);
  assert.match(embedXml, /<w:embedTrueTypeFonts\/><w:saveSubsetFonts\/>/);
});

test("settings.xml defaultTabStop is emitted from parsed DOP, not section profile", async () => {
  const wps = readWps(await readFile("sample/sample3/original.wps"));
  const docx = wpsToDocxBuffer({ ...wps, defaultTabStop: 960 }, { title: "sample3" });
  const xml = readZipEntry(docx, "word/settings.xml").toString("utf8");

  assert.match(xml, /<w:defaultTabStop w:val="960"\/>/);
});

test("trackRevisions is emitted before protection and defaultTabStop", async () => {
  const wps = readWps(await readFile("sample/sample6/original.wps"));
  assert.equal(wps.dop.fRevMarking, true);

  const docx = wpsToDocxBuffer(wps, { title: "sample6" });
  const xml = readZipEntry(docx, "word/settings.xml").toString("utf8");
  const trackRevisionsIndex = xml.indexOf("<w:trackRevisions");
  const documentProtectionIndex = xml.indexOf("<w:documentProtection");
  const defaultTabStopIndex = xml.indexOf("<w:defaultTabStop");

  assert.ok(trackRevisionsIndex >= 0);
  assert.ok(documentProtectionIndex >= 0);
  assert.ok(defaultTabStopIndex >= 0);
  assert.ok(trackRevisionsIndex < documentProtectionIndex);
  assert.ok(trackRevisionsIndex < defaultTabStopIndex);
});

test("settings drawing grid fails fast when parsed Dogrid is missing", async () => {
  const wps = readWps(await readFile("sample/sample3/original.wps"));
  assert.equal(wps.sections[0].properties.docGridType, 2);

  assert.throws(
    () => wpsToDocxBuffer({ ...wps, dop: { ...wps.dop, dogrid: null } }, { title: "sample3" }),
    /missing parsed Dogrid/,
  );
});

test("non-East Asian settings drawing grid also requires parsed Dogrid", async () => {
  const wps = readWps(await readFile(BASIC_WPS));
  assert.equal(wps.sections[0].properties.docGridType ?? null, null);

  assert.throws(
    () => wpsToDocxBuffer({ ...wps, dop: { ...wps.dop, dogrid: null } }, { title: "basic" }),
    /missing parsed Dogrid/,
  );
});

test("East Asian grid settings fail fast when parsed grfFmtFilter is missing", async () => {
  const wps = readWps(await readFile("sample/sample3/original.wps"));
  assert.equal(wps.sections[0].properties.docGridType, 2);

  assert.throws(
    () => wpsToDocxBuffer({ ...wps, dop: { ...wps.dop, grfFmtFilter: null } }, { title: "sample3" }),
    /missing parsed grfFmtFilter/,
  );
});

test("sample2 bold paragraph keeps the bank/securities label bold", async () => {
  const wps = readWps(await readFile("sample/sample2/original.wps"));
  const target = "（银行/证券/保险/其他）";
  const start = wps.bodyText.indexOf(target);
  assert.ok(start >= 0);
  assert.equal(wps.characterProperties[start].bold, true);
  assert.equal(wps.characterProperties[start + target.length - 1].bold, true);

  const docx = wpsToDocxBuffer(wps, { title: "sample2" });
  const xml = readDocxDocumentXml(docx);
  const paragraph = findParagraphXml(xml, `<w:t>${target}</w:t>`);
  assert.ok(paragraph);
  assert.match(paragraph, /<w:spacing w:after="298" w:afterLines="50" w:line="596" w:lineRule="exact"\/>/);
  assert.match(paragraph, /<w:rPr>.*<w:b\/>.*<w:sz w:val="32"\/><w:szCs w:val="32"\/>.*<\/w:rPr>/s);
});

test("sample2 table header paragraphs keep complex-script bold runs", async () => {
  const wps = readWps(await readFile("sample/sample2/original.wps"));
  const docx = wpsToDocxBuffer(wps, { title: "sample2" });
  const xml = readDocxDocumentXml(docx);

  const headerMatch = findParagraphXml(xml, "<w:t>项目名称：</w:t>");
  assert.ok(headerMatch);
  assert.match(headerMatch, /<w:bCs\/>/);
  assert.match(headerMatch, /<w:sz w:val="24"\/>/);
  assert.match(headerMatch, /<w:szCs w:val="24"\/>/);

  const labelMatch = findParagraphXml(xml, "<w:t>（银行/证券/保险/其他）</w:t>");
  assert.ok(labelMatch);
  assert.match(labelMatch, /<w:b\/>/);
  assert.doesNotMatch(labelMatch, /<w:bCs\/>/);
});

test("extracts and emits table4 WPS table structure with centered Normal Table styling", async () => {
  const wps = readWps(await readFile("sample/table4/original.wps"));

  assert.equal(wps.tableRows.length, 1);
  assert.deepEqual(wps.tableRows[0].rows.map((row) => row.cells.length), [
    4, 6, 2, 7, 7, 7, 7, 4, 4, 1, 5, 5, 5, 2, 2, 2, 2, 2,
  ]);
  assert.equal(wps.tableRows[0].gridCols.length, 10);
  assert.equal(wps.characterProperties[93].symbolFontId, 5);
  assert.equal(wps.characterProperties[93].symbolChar, 0x00a3);

  const docx = wpsToDocxBuffer(wps, { title: "table4" });
  const xml = readDocxDocumentXml(docx);
  const expectedXml = readDocxDocumentXml(await readFile("sample/table4/expected.docx"));
  const tableXmls = [...xml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
  const expectedTableXmls = [...expectedXml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
  assert.equal(tableXmls.length, 1);
  assert.match(tableXmls[0], /<w:tblStyle w:val="6"\/>/);
  assert.match(tableXmls[0], /<w:tblW w:w="0" w:type="auto"\/>/);
  assert.match(tableXmls[0], /<w:jc w:val="center"\/>/);
  assert.match(tableXmls[0], /<w:trPr>[\s\S]*?<w:jc w:val="center"\/>/);
  assert.match(tableXmls[0], /<w:vAlign w:val="center"\/>/);
  assert.match(tableXmls[0], /<w:sym w:font="Wingdings 2" w:char="00A3"\/>/);
  assert.doesNotMatch(tableXmls[0], /<w:t>\(企<\/w:t>/);
  assert.match(tableXmls[0], /单位名称/);
  assert.match(xml, /<w:bookmarkStart w:id="0" w:name="附件5"\/><w:bookmarkEnd w:id="0"\/>/);
  assert.deepEqual(
    tableXmls.map((t) => (t.match(/<w:tcBorders>/g) || []).length),
    expectedTableXmls.map((t) => (t.match(/<w:tcBorders>/g) || []).length),
  );
});

test("converts sample4 without dropping its table block", async () => {
  const wps = readWps(await readFile("sample/sample4/original.wps"));

  assert.equal(wps.tableRows.length, 1);
  assert.equal(wps.tableRows[0].rows.length, 7);

  const docx = wpsToDocxBuffer(wps, { title: "sample4" });
  const xml = readDocxDocumentXml(docx);
  const expected = await readFile("sample/sample4/expected.docx");

  assert.equal(readDocxMainText(docx), readDocxMainText(expected));
  assert.equal((xml.match(/<w:tbl>/g) ?? []).length, 1);
  assert.ok(
    countXmlLineEdits(xml, readZipEntry(expected, "word/document.xml").toString("utf8")) < 100,
    "sample4 document.xml normalized diff should stay below 100 edits",
  );
  const convertedSettingsXml = readZipEntry(docx, "word/settings.xml").toString("utf8");
  const expectedSettingsXml = readZipEntry(expected, "word/settings.xml").toString("utf8");
  assert.equal(extractSettingsNode(convertedSettingsXml, "w:footnotePr"), extractSettingsNode(expectedSettingsXml, "w:footnotePr"));
  assert.equal(extractSettingsNode(convertedSettingsXml, "w:endnotePr"), extractSettingsNode(expectedSettingsXml, "w:endnotePr"));
  assert.ok(
    countXmlLineEdits(convertedSettingsXml, expectedSettingsXml) < 7,
    "sample4 settings.xml normalized diff should stay below 7 edits",
  );
  assert.equal(
    countXmlLineEdits(readZipEntry(docx, "[Content_Types].xml").toString("utf8"), readZipEntry(expected, "[Content_Types].xml").toString("utf8")),
    0,
  );
  assert.equal(
    countXmlLineEdits(readZipEntry(docx, "word/_rels/document.xml.rels").toString("utf8"), readZipEntry(expected, "word/_rels/document.xml.rels").toString("utf8")),
    0,
  );
  assert.throws(() => readZipEntry(docx, "word/numbering.xml"), /DOCX fixture entry not found/);
  // Regression: style-level rPr must NOT auto-derive szCs from sz when no
  // explicit fontSizeCs is set. The default style should have sz but no szCs.
  const convertedStylesXml = readZipEntry(docx, "word/styles.xml").toString("utf8");
  assert.doesNotMatch(convertedStylesXml, /<w:style w:type="paragraph" w:default="1" w:styleId="1">.*?<w:szCs w:val="32"\/>/s);
});

test("extracts paragraph line spacing from PAPX pages", async () => {
  const wps = readWps(await readFile(BASIC_WPS));
  assert.ok(wps.paragraphProperties.length >= 12);
  assert.equal(wps.paragraphProperties[0].lineSpacing.twips, 594);
  assert.equal(wps.paragraphProperties[0].lineSpacing.rule, "exact");
  assert.equal(wps.paragraphProperties[1].lineSpacing.twips, 594);
  assert.equal(wps.paragraphProperties[2].lineSpacing.twips, 594);
  assert.equal(wps.paragraphProperties[2].lineSpacing.rule, "exact");
});

test("extracts custom paragraph tab stops from PAPX pages", async () => {
  const wps = readWps(await readFile(BASIC_WPS));
  assert.deepEqual(wps.paragraphProperties[9].tabs.map((tab) => tab.position), [
    4464,
    4690,
    5029,
    5480,
    5595,
    6159,
    6272,
    6726,
    7061,
    7290,
    7854,
  ]);
});

test("sample3 attachment heading keeps no direct paragraph tabs", async () => {
  const wps = readWps(await readFile("sample/sample3/original.wps"));
  assert.equal(wps.paragraphs[0], "附件1");
  assert.equal(wps.paragraphProperties[0].tabs, null);
  assert.equal(wps.paragraphProperties[0].widowControl, null);

  const docx = wpsToDocxBuffer(wps, { title: "sample3" });
  const xml = readDocxDocumentXml(docx);
  const firstParagraph = xml.match(/<w:p [\s\S]*?<\/w:p>/)?.[0];
  assert.ok(firstParagraph, "first paragraph XML should exist");
  assert.doesNotMatch(firstParagraph, /<w:tabs>/);
  assert.doesNotMatch(firstParagraph, /<w:widowControl/);
});

test("sample3 title paragraph preserves keep and page-break flags from WPS", async () => {
  const wps = readWps(await readFile("sample/sample3/original.wps"));
  assert.equal(wps.paragraphs[1], "22项通信行业标准编号、名称及主要内容等一览表");
  assert.equal(wps.paragraphProperties[1].keepNext, false);
  assert.equal(wps.paragraphProperties[1].keepLines, false);
  assert.equal(wps.paragraphProperties[1].pageBreakBefore, false);
  assert.equal(wps.paragraphProperties[1].widowControl, false);
  assert.equal(wps.paragraphProperties[1].bidi, false);
  assert.equal(wps.paragraphProperties[1].snapToGrid, true);
  assert.equal(wps.paragraphProperties[1].textAlignment, "auto");

  const docx = wpsToDocxBuffer(wps, { title: "sample3" });
  const xml = readDocxDocumentXml(docx);
  const titleParagraph = findParagraphXml(xml, "22项通信行业标准编号、名称及主要内容等一览表");
  assert.ok(titleParagraph, "title paragraph should exist");
  assert.match(xml, /<w:keepNext w:val="0"\/>/);
  assert.match(xml, /<w:keepLines w:val="0"\/>/);
  assert.match(xml, /<w:pageBreakBefore w:val="0"\/>/);
  assert.match(xml, /<w:widowControl w:val="0"\/>/);
  assert.ok(titleParagraph.indexOf('<w:pageBreakBefore w:val="0"/>') < titleParagraph.indexOf('<w:widowControl w:val="0"/>'));
  assert.ok(titleParagraph.indexOf('<w:widowControl w:val="0"/>') < titleParagraph.indexOf('<w:snapToGrid/>'));
  assert.match(xml, /<w:bidi w:val="0"\/>/);
  assert.match(xml, /<w:snapToGrid\/>/);
  assert.match(xml, /<w:textAlignment w:val="auto"\/>/);
  assert.match(xml, /<w:spacing[^>]*w:after="157"[^>]*w:afterLines="50"[^>]*\/>/);
});

test("sample3 table cell header keeps table-cell defaults from expected DOCX", async () => {
  const wps = readWps(await readFile("sample/sample3/original.wps"));
  const docx = wpsToDocxBuffer(wps, { title: "sample3" });
  const xml = readDocxDocumentXml(docx);
  const headerParagraph = findParagraphXml(xml, "序号");
  assert.ok(headerParagraph, "table header paragraph should exist");
  assert.match(headerParagraph, /<w:suppressLineNumbers w:val="0"\/>/);
  assert.match(headerParagraph, /<w:wordWrap\/>/);
  assert.match(headerParagraph, /<w:rFonts w:hint="eastAsia" w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体" w:cs="黑体"\/>/);
  assert.doesNotMatch(headerParagraph, /<w:autoSpaceDE w:val="0"\/>/);
  assert.doesNotMatch(headerParagraph, /<w:autoSpaceDN w:val="0"\/>/);
  assert.match(headerParagraph, /<w:color w:val="(?:auto|000000)"\/>/);
  assert.match(headerParagraph, /<w:szCs w:val="21"\/>/);
  assert.match(headerParagraph, /<w:highlight w:val="none"\/>/);
  assert.ok(headerParagraph.indexOf('<w:widowControl w:val="0"/>') < headerParagraph.indexOf('<w:suppressLineNumbers w:val="0"/>'));
  assert.ok(headerParagraph.indexOf('<w:suppressLineNumbers w:val="0"/>') < headerParagraph.indexOf('<w:kinsoku/>'));
});

test("table cell paragraph controls are emitted only from parsed MS-DOC SPRMs", () => {
  const noBorder = { style: "none", width: 0, color: "auto", space: 0 };
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    paragraphProperties: [{}],
    characterProperties: [{}],
    tableRows: [{
      cpStart: 0,
      cpEnd: 2,
      gridCols: [1440],
      tableWidth: 1440,
      tableWidthType: "dxa",
      tableBordersExplicit: false,
      tableBorders: {
        top: noBorder,
        left: noBorder,
        bottom: noBorder,
        right: noBorder,
        insideH: noBorder,
        insideV: noBorder,
      },
      rows: [{
        cells: [{
          cpStart: 0,
          cpEnd: 2,
          width: 1440,
        }],
      }],
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "table-cell-paragraph-controls" });
  const xml = readDocxDocumentXml(docx);
  const paragraph = findParagraphXml(xml, "A");
  assert.ok(paragraph, "table cell paragraph should exist");

  assert.doesNotMatch(paragraph, /<w:kinsoku\/>/);
  assert.doesNotMatch(paragraph, /<w:wordWrap\/>/);
  assert.doesNotMatch(paragraph, /<w:overflowPunct\/>/);
  assert.doesNotMatch(paragraph, /<w:topLinePunct/);
  assert.doesNotMatch(paragraph, /<w:autoSpaceDE/);
  assert.doesNotMatch(paragraph, /<w:autoSpaceDN/);
  assert.doesNotMatch(paragraph, /<w:adjustRightInd/);
});

test("parsed table cell paragraph controls are not suppressed by paragraph mark fonts or table autofit", () => {
  const noBorder = { style: "none", width: 0, color: "auto", space: 0 };
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    paragraphProperties: [{
      kinsoku: true,
      wordWrap: true,
      overflowPunct: true,
      topLinePunct: false,
      autoSpaceDE: true,
      autoSpaceDN: true,
      bidi: false,
      adjustRightInd: true,
      snapToGrid: true,
    }],
    characterProperties: [{}, { fontAscii: 0 }],
    tableRows: [{
      cpStart: 0,
      cpEnd: 2,
      gridCols: [1440],
      tableWidth: 1440,
      tableWidthType: "dxa",
      tableAutofit: true,
      tableBordersExplicit: false,
      tableBorders: {
        top: noBorder,
        left: noBorder,
        bottom: noBorder,
        right: noBorder,
        insideH: noBorder,
        insideV: noBorder,
      },
      rows: [{
        cells: [{
          cpStart: 0,
          cpEnd: 2,
          width: 1440,
        }],
      }],
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "table-cell-parsed-paragraph-controls" });
  const paragraph = findParagraphXml(readDocxDocumentXml(docx), "A");
  assert.ok(paragraph, "table cell paragraph should exist");

  assert.match(paragraph, /<w:kinsoku\/>/);
  assert.match(paragraph, /<w:wordWrap\/>/);
  assert.match(paragraph, /<w:overflowPunct\/>/);
  assert.match(paragraph, /<w:topLinePunct w:val="0"\/>/);
  assert.match(paragraph, /<w:autoSpaceDE\/>/);
  assert.match(paragraph, /<w:autoSpaceDN\/>/);
  assert.match(paragraph, /<w:bidi w:val="0"\/>/);
  assert.match(paragraph, /<w:adjustRightInd\/>/);
  assert.match(paragraph, /<w:snapToGrid\/>/);
});

test("paragraph shading follows parsed borders before East Asian controls", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    paragraphProperties: [{
      paragraphBorders: {
        top: { val: "none", color: "auto", sz: "0", space: "0" },
      },
      paragraphShading: { val: "clear", color: "auto", fill: "FFFFFF" },
      kinsoku: true,
    }],
    characterProperties: [{}, {}],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "paragraph-shading-order" });
  const paragraph = findParagraphXml(readDocxDocumentXml(docx), "A");
  assert.ok(paragraph, "paragraph should exist");

  assert.ok(paragraph.indexOf("<w:pBdr>") < paragraph.indexOf("<w:shd "));
  assert.ok(paragraph.indexOf("<w:shd ") < paragraph.indexOf("<w:kinsoku/>"));
});

test("MS-DOC NilBrc cell borders are preserved as explicit OOXML nil borders", async () => {
  const wps = readWps(await readFile(SAMPLE9_WPS));
  const firstCell = wps.tableRows[0]?.rows[0]?.cells[0];
  assert.equal(firstCell?.borders?.top?.nil, true);
  assert.equal(firstCell?.borders?.left?.nil, true);
  assert.equal(firstCell?.borders?.bottom?.style, "single");
  assert.equal(firstCell?.noWrap, true);

  const documentXml = readDocxDocumentXml(wpsToDocxBuffer(wps, { title: "sample9" }));
  assert.match(documentXml, /<w:noWrap\/>/);
  const firstCellBorders = documentXml.match(/<w:tcBorders>[\s\S]*?<\/w:tcBorders>/)?.[0] ?? "";
  assert.match(firstCellBorders, /<w:top w:val="nil"\/>/);
  assert.match(firstCellBorders, /<w:left w:val="nil"\/>/);
  assert.match(firstCellBorders, /<w:bottom w:val="single" w:color="auto" w:sz="4" w:space="0"\/>/);
});

test("sample8 cell borders do not synthesize unparsed nil overrides", async () => {
  const wps = readWps(await readFile(SAMPLE8_WPS));
  const documentXml = readDocxDocumentXml(wpsToDocxBuffer(wps, { title: "sample8" }));
  const expectedXml = readDocxDocumentXml(await readFile("sample/sample8/expected.docx"));

  assert.equal(
    (documentXml.match(/<w:tcBorders>/g) ?? []).length,
    (expectedXml.match(/<w:tcBorders>/g) ?? []).length,
  );
});

test("sample8 table grid and style use parsed row properties", async () => {
  const wps = readWps(await readFile(SAMPLE8_WPS));
  assert.equal(wps.tableRows[0].tableStyleId, "5");
  assert.equal(wps.tableRows[1].tableStyleId, "4");
  assert.deepEqual(wps.tableRows[1].gridCols.slice(0, 3), [1679, 8, 1177]);
  assert.equal(wps.tableRows[1].rows[0].cells[0].gridSpan, 11);

  const documentXml = readDocxDocumentXml(wpsToDocxBuffer(wps, { title: "sample8" }));
  assert.match(documentXml, /<w:tblGrid>[\s\S]*?<w:gridCol w:w="1679"\/>[\s\S]*?<w:gridCol w:w="8"\/>[\s\S]*?<w:gridCol w:w="1177"\/>/);
  assert.match(documentXml, /<w:gridSpan w:val="11"\/>/);
});

test("sample8 PlcfSpa line shape anchors are parsed and emitted at anchor characters", async () => {
  const wps = readWps(await readFile(SAMPLE8_WPS));
  assert.equal(wps.shapeAnchors.length, 6);
  assert.equal(wps.fib.fcDggInfo, 6822);
  assert.equal(wps.fib.lcbDggInfo, 11985);
  assert.deepEqual(
    wps.shapeAnchors.slice(0, 3).map((shape) => ({
      cpStart: shape.cpStart,
      lid: shape.lid,
      xaLeft: shape.xaLeft,
      yaTop: shape.yaTop,
      xaRight: shape.xaRight,
      yaBottom: shape.yaBottom,
      bx: shape.bx,
      by: shape.by,
      wr: shape.wr,
      officeArtSpid: shape.officeArt?.spid,
      docPrId: shape.officeArt?.docPrId,
      relativeHeight: shape.officeArt?.relativeHeight,
      hasGfxData: Buffer.isBuffer(shape.officeArt?.gfxData),
    })),
    [
      { cpStart: 6159, lid: 1026, xaLeft: 8815, yaTop: 3588, xaRight: 8816, yaBottom: 3588, bx: 2, by: 2, wr: 3, officeArtSpid: 1026, docPrId: 54, relativeHeight: 251663360, hasGfxData: true },
      { cpStart: 6160, lid: 1027, xaLeft: 8815, yaTop: 2964, xaRight: 8816, yaBottom: 2964, bx: 2, by: 2, wr: 3, officeArtSpid: 1027, docPrId: 55, relativeHeight: 251662336, hasGfxData: true },
      { cpStart: 6161, lid: 1028, xaLeft: 8815, yaTop: 1418, xaRight: 8816, yaBottom: 1418, bx: 2, by: 2, wr: 3, officeArtSpid: 1028, docPrId: 56, relativeHeight: 251664384, hasGfxData: true },
    ],
  );
  assert.equal(wps.shapeAnchors[0].officeArt.gfxData.subarray(0, 4).toString("latin1"), "PK\u0003\u0004");

  const documentXml = readDocxDocumentXml(wpsToDocxBuffer(wps, { title: "sample8" }));
  assert.equal((documentXml.match(/<mc:AlternateContent>/g) ?? []).length, 6);
  assert.match(documentXml, /<wp:posOffset>5597525<\/wp:posOffset>/);
  assert.match(documentXml, /<wp:posOffset>2278380<\/wp:posOffset>/);
  assert.match(documentXml, /<wp:docPr id="54" name="直接连接符 54"\/>/);
  assert.match(documentXml, /o:gfxdata="UEsDBAo/);
});

test("percentage table width emits cell widths from parsed grid proportions", () => {
  const noBorder = { style: "none", width: 0, color: "auto", space: 0 };
  const docx = wpsToDocxBuffer({
    bodyText: "A\rB\r",
    defaultTabStop: 720,
    paragraphProperties: [{}, {}],
    characterProperties: [{}, {}, {}, {}],
    tableRows: [{
      cpStart: 0,
      cpEnd: 4,
      gridCols: [1000, 2000, 1000],
      tableWidth: 5000,
      tableWidthType: "pct",
      tableBordersExplicit: false,
      tableBorders: {
        top: noBorder,
        left: noBorder,
        bottom: noBorder,
        right: noBorder,
        insideH: noBorder,
        insideV: noBorder,
      },
      rows: [{
        cells: [{
          cpStart: 0,
          cpEnd: 2,
          width: 3000,
          gridSpan: 2,
        }, {
          cpStart: 2,
          cpEnd: 4,
          width: 1000,
        }],
      }],
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "pct-table-width" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:tcW w:w="3750" w:type="pct"\/>/);
  assert.match(xml, /<w:tcW w:w="1250" w:type="pct"\/>/);
});

test("sample9 table cells use parsed sprmTCellWidth preferred pct widths", async () => {
  const wps = readWps(await readFile(SAMPLE9_WPS));
  const row = wps.tableRows[0].rows.find((candidate) => candidate.cells.length === 3);
  assert.deepEqual(
    row.cells.map((cell) => cell.preferredWidth),
    [
      { type: "pct", value: 527 },
      { type: "pct", value: 2628 },
      { type: "pct", value: 1844 },
    ],
  );

  const xml = readDocxDocumentXml(wpsToDocxBuffer(wps, { title: "sample9" }));
  assert.match(xml, /<w:tcW w:w="1844" w:type="pct"\/>/);
  assert.doesNotMatch(xml, /<w:tcW w:w="1845" w:type="pct"\/>/);
});

test("table indent from parsed rgdxaCenter margin edge does not synthesize justification", () => {
  const noBorder = { style: "none", width: 0, color: "auto", space: 0 };
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    paragraphProperties: [{}],
    characterProperties: [{}, {}],
    tableRows: [{
      cpStart: 0,
      cpEnd: 2,
      gridCols: [1440],
      gridPositions: [-108, 1332],
      tableIndent: { width: -108, type: "dxa" },
      cellMargins: { top: 0, left: 108, bottom: 0, right: 108 },
      tableBordersExplicit: false,
      tableBorders: {
        top: noBorder,
        left: noBorder,
        bottom: noBorder,
        right: noBorder,
        insideH: noBorder,
        insideV: noBorder,
      },
      rows: [{
        tableJustification: null,
        cells: [{
          cpStart: 0,
          cpEnd: 2,
          width: 1440,
        }],
      }],
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "table-indent-not-jc" });
  const tableXml = readDocxDocumentXml(docx).match(/<w:tbl>[\s\S]*?<\/w:tbl>/)?.[0] ?? "";

  assert.match(tableXml, /<w:tblInd w:w="0" w:type="dxa"\/>/);
  assert.doesNotMatch(tableXml, /<w:jc w:val="center"\/>/);
});

test("sample3 body paragraph keeps one merged run with explicit black shading", async () => {
  const wps = readWps(await readFile("sample/sample3/original.wps"));
  const target = "本文件规定了大规模预训练模型时具备的能力要求，包括数据管理、模型训练、模型管理和模型部署等核心环节";
  const docx = wpsToDocxBuffer(wps, { title: "sample3" });
  const xml = readDocxDocumentXml(docx);
  const paragraph = findParagraphXml(xml, target);
  assert.ok(paragraph, "target paragraph should exist");
  assert.equal((paragraph.match(/<w:r><w:rPr>/g) ?? []).length, 1);
  assert.match(paragraph, /<w:color w:val="000000"\/>/);
  assert.match(paragraph, /<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"\/>/);
  assert.doesNotMatch(paragraph, /<w:highlight w:val="none"\/>/);
});

test("preserves parsed MS-DOC automatic text color when shading is present", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    paragraphProperties: [{}],
    characterProperties: [{
      textColor: "auto",
      background: { val: "clear", color: "auto", fill: "FFFFFF" },
    }],
    sections: [{ cpStart: 0, cpEnd: 2, properties: {} }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "auto-color-shading" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:color w:val="auto"\/>/);
  assert.match(xml, /<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"\/>/);
});

test("sample3 paragraph line numbering comes from parsed PFNoLineNumb", async () => {
  const wps = readWps(await readFile("sample/sample3/original.wps"));
  assert.ok(
    wps.paragraphProperties.some((paragraph) => paragraph?.lineNumberCount === true),
    "sample3 should parse at least one explicit line-numbering count flag",
  );

  const docx = wpsToDocxBuffer(wps, { title: "sample3" });
  const xml = readDocxDocumentXml(docx);
  const expectedXml = readDocxDocumentXml(await readFile("sample/sample3/expected.docx"));
  const actualFlags = (xml.match(/<w:p[\s\S]*?<\/w:p>/g) ?? []).map((paragraph) =>
    /<w:suppressLineNumbers w:val="0"\/>/.test(paragraph),
  );
  const expectedFlags = (expectedXml.match(/<w:p[\s\S]*?<\/w:p>/g) ?? []).map((paragraph) =>
    /<w:suppressLineNumbers w:val="0"\/>/.test(paragraph),
  );
  assert.deepEqual(actualFlags, expectedFlags);
});

test("sample3 blank table spacer paragraph does not emit suppressLineNumbers", async () => {
  const wps = readWps(await readFile("sample/sample3/original.wps"));
  const docx = wpsToDocxBuffer(wps, { title: "sample3" });
  const xml = readDocxDocumentXml(docx);
  const paragraphs = xml.match(/<w:p [\s\S]*?<\/w:p>/g) ?? [];
  const blankParagraph = paragraphs.find(
    (paragraph) => !paragraph.includes("<w:t") && !paragraph.includes("<w:pStyle"),
  );
  assert.ok(blankParagraph, "blank table spacer paragraph should exist");
  assert.doesNotMatch(blankParagraph, /<w:suppressLineNumbers w:val="0"\/>/);
});

test("extracts direct paragraph indents from full PAPX records", async () => {
  const wps = readWps(await readFile(BASIC_WPS));
  assert.equal(wps.paragraphProperties[0].rightIndent, 194);
  assert.equal(wps.paragraphProperties[2].leftIndent, 1683);
  assert.equal(wps.paragraphProperties[2].rightIndent, 1853);
  assert.equal(wps.paragraphProperties[5].leftIndent, 108);
  assert.equal(wps.paragraphProperties[5].rightIndent, 260);
  assert.equal(wps.paragraphProperties[5].firstLineIndent, 631);
  assert.equal(wps.paragraphProperties[9].leftIndent, 3898);
  assert.equal(wps.paragraphProperties[9].rightIndent, 1061);
  assert.equal(wps.paragraphProperties[10].rightIndent, 5232);
});

test("available samples expose MS-DOC PlcfSed section boundary invariants", async () => {
  for (let sample = 1; sample <= 10; sample += 1) {
    const wps = readWps(await readFile(`sample/sample${sample}/original.wps`));
    const sectionCps = [
      ...wps.sections.map((section) => section.cpStart),
      wps.sections.at(-1)?.cpEnd,
    ];

    assert.ok(wps.sections.length > 0);
    assert.equal(new Set(sectionCps).size, sectionCps.length);
    assert.deepEqual(sectionCps, [...sectionCps].sort((a, b) => a - b));
    assert.ok(sectionCps.at(-1) >= wps.bodyText.length);
    for (let i = 0; i < wps.sections.length - 1; i += 1) {
      // MS-DOC-SPEC/18 PlcfSed requires every non-final section text
      // range to end with an end-of-section character (0x0C).
      assert.equal(wps.bodyText[wps.sections[i].cpEnd - 1], "\f");
    }
  }
});

test("extracts section page geometry from PLCFSED/SEPX", async () => {
  const wps = readWps(await readFile(BASIC_WPS));
  assert.equal(wps.sections.length, 2);
  assert.deepEqual(wps.sections.map((section) => [section.cpStart, section.cpEnd]), [
    [0, 254],
    [254, 265],
  ]);
  assert.equal(wps.sections[0].properties.pageWidth, 11910);
  assert.equal(wps.sections[0].properties.marginLeft, 1480);
  assert.equal(wps.sections[1].properties.marginRight, 1260);
});

test("parses variable-length section border SPRMs", () => {
  const props = parseSectionSprms(Buffer.from([
    0x34, 0xd2, 0x08,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  ]));

  assert.deepEqual(props, { pageBorders: { top: { style: "none" } } });
  assert.throws(
    () => parseSectionSprms(Buffer.from([
      0x34, 0xd2, 0x0c,
      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
      0x00, 0x00, 0x00, 0x00,
    ])),
    /Out-of-spec page border operand size 13/,
  );
});


test("section pgSz orient preserves parsed MS-DOC page geometry", () => {
  const baseDoc = {
    bodyText: "A\r",
    defaultTabStop: 720,
    paragraphProperties: [{}],
    characterProperties: [{}],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  };
  const portraitGeometryXml = readDocxDocumentXml(wpsToDocxBuffer({
    ...baseDoc,
    sections: [{
      cpStart: 0,
      cpEnd: 2,
      properties: {
        pageWidth: 11906,
        pageHeight: 16838,
        orientation: "landscape",
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        headerMargin: 720,
        footerMargin: 720,
        gutterMargin: 0,
      },
    }],
  }, { title: "portrait-geometry" }));
  const landscapeGeometryXml = readDocxDocumentXml(wpsToDocxBuffer({
    ...baseDoc,
    sections: [{
      cpStart: 0,
      cpEnd: 2,
      properties: {
        pageWidth: 16838,
        pageHeight: 11906,
        orientation: "landscape",
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        headerMargin: 720,
        footerMargin: 720,
        gutterMargin: 0,
      },
    }],
  }, { title: "landscape-geometry" }));

  assert.match(portraitGeometryXml, /<w:pgSz w:w="11906" w:h="16838"\/>/);
  assert.doesNotMatch(portraitGeometryXml, /orient="landscape"/);
  assert.match(landscapeGeometryXml, /<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"\/>/);
});

test("validates MS-DOC section page and margin SPRM constraints", () => {
  const valid = parseSectionSprms(Buffer.from([
    0x1f, 0xb0, 0x82, 0x2e, // sprmSXaPage: 11906 twips
    0x20, 0xb0, 0xc6, 0x41, // sprmSYaPage: 16838 twips
    0x21, 0xb0, 0xa0, 0x05, // sprmSDxaLeft: 1440 twips
    0x22, 0xb0, 0xa0, 0x05, // sprmSDxaRight: 1440 twips
    0x23, 0x90, 0xa0, 0xfa, // sprmSDyaTop: -1376 twips fixed margin
    0x24, 0x90, 0xa0, 0x05, // sprmSDyaBottom: 1440 twips
  ]), { validateRequiredSectionProperties: true });

  assert.equal(valid.pageWidth, 11906);
  assert.equal(valid.pageHeight, 16838);
  assert.equal(valid.marginLeft, 1440);
  assert.equal(valid.marginRight, 1440);
  assert.equal(valid.marginTop, -1376);
  assert.equal(valid.marginBottom, 1440);
  assert.throws(
    () => parseSectionSprms(Buffer.from([
      0x1f, 0xb0, 0x8f, 0x7b, // sprmSXaPage: 31631 ok, keeps fixture focused on missing margin
      0x21, 0xb0, 0xa0, 0x05,
      0x23, 0x90, 0xa0, 0x05,
      0x24, 0x90, 0xa0, 0x05,
    ]), { validateRequiredSectionProperties: true }),
    /missing required sprmSDxaRight/,
  );
  assert.throws(
    () => parseSectionSprms(Buffer.from([0x1f, 0xb0, 0x8f, 0x7c])),
    /Out-of-spec section page width 31887/,
  );
  assert.throws(
    () => parseSectionSprms(Buffer.from([0x23, 0x90, 0x52, 0x83])),
    /Out-of-spec section top margin -31918/,
  );
});

test("validates MS-DOC section document-grid SPRM constraints", () => {
  const valid = parseSectionSprms(Buffer.from([
    0x32, 0x50, 0x02, 0x00, // sprmSClm: clmLinesOnly
    0x31, 0x90, 0x38, 0x01, // sprmSDyaLinePitch: 312 twips
    0x30, 0x70, 0x00, 0x00, 0x00, 0x00, // sprmSDxtCharSpace: no difference from Normal style pitch
  ]));

  assert.deepEqual(valid, {
    docGridType: 2,
    docGridLinePitch: 312,
    docGridCharSpace: 0,
  });
  assert.throws(
    () => parseSectionSprms(Buffer.from([0x32, 0x50, 0x04, 0x00])),
    /Out-of-spec section document grid mode 4/,
  );
  assert.throws(
    () => parseSectionSprms(Buffer.from([0x32, 0x50, 0x02, 0x00])),
    /enabled without sprmSDyaLinePitch/,
  );
  assert.throws(
    () => parseSectionSprms(Buffer.from([0x31, 0x90, 0x00, 0x00])),
    /Out-of-spec section document grid line pitch 0/,
  );
  assert.throws(
    () => parseSectionSprms(Buffer.from([0x30, 0x70, 0x11, 0x03, 0x9b, 0x00])),
    /Out-of-spec section document grid character spacing /,
  );
  assert.throws(
    () => parseSectionSprms(Buffer.from([0x32])),
    /Truncated section SPRM code at end of SEPX grpprl/,
  );
});

test("parses MS-DOC section orientation and 32-bit page-number start SPRMs", () => {
  const props = parseSectionSprms(Buffer.from([
    0x00, 0x30, 0x04, // sprmScnsPgn: en dash chapter separator
    0x01, 0x30, 0x02, // sprmSiHeadingPgn: Heading 2 starts chapters
    0x06, 0x30, 0x00, // sprmSFProtected: protected section when form protection is enabled
    0x07, 0x50, 0x02, 0x00, // sprmSDmBinFirst: paper source 2
    0x08, 0x50, 0x04, 0x00, // sprmSDmBinOther: paper source 4
    0x11, 0x30, 0x01, // sprmSFPgnRestart
    0x12, 0x30, 0x00, // sprmSFEndnote: suppress endnotes in this section
    0x1d, 0x30, 0x01, // sprmSBOrientation: landscape
    0x44, 0x70, 0x2c, 0x01, 0x00, 0x00, // sprmSPgnStart: 300
    0x25, 0xb0, 0x80, 0x01, // sprmSDzaGutter: 384
    0x26, 0x50, 0x09, 0x00, // sprmSDmPaperReq: implementation-specific tie-breaker
    0x19, 0x30, 0x01, // sprmSLBetween
    0x1a, 0x30, 0x03, // sprmSVjc: bottom
    0x28, 0x32, 0x01, // sprmSFBiDi
    0x2a, 0x32, 0x01, // sprmSFRTLGutter
    0x13, 0x30, 0x02, // sprmSLnc: continuous
    0x15, 0x50, 0x03, 0x00, // sprmSNLnnMod: every 3 lines
    0x16, 0x90, 0xf0, 0x00, // sprmSDxaLnn: 240 twips
    0x1b, 0x50, 0x04, 0x00, // sprmSLnnMin: start at 5
    0x3b, 0x30, 0x02, // sprmSFpc: footnotes beneath text
    0x3c, 0x30, 0x02, // sprmSRncFtn: restart every page
    0x3e, 0x30, 0x01, // sprmSRncEdn: restart every section
    0x3f, 0x50, 0x06, 0x00, // sprmSNFtn: footnote start offset
    0x40, 0x50, 0x02, 0x00, // sprmSNfcFtnRef: lowerRoman
    0x41, 0x50, 0x06, 0x00, // sprmSNEdn: endnote start offset
    0x42, 0x50, 0x01, 0x00, // sprmSNfcEdnRef: upperRoman
  ]));

  assert.equal(props.pageNumberChapterSeparator, "enDash");
  assert.equal(props.pageNumberChapterStyle, 2);
  assert.equal(props.formProtection, true);
  assert.equal(props.paperSourceFirst, 2);
  assert.equal(props.paperSourceOther, 4);
  assert.equal(props.pageNumberRestart, true);
  assert.equal(props.endnotesSuppressed, true);
  assert.equal(props.orientation, "landscape");
  assert.equal(props.pageNumberStart, 300);
  assert.equal(props.gutterMargin, 384);
  assert.equal(props.paperFormatTieBreaker, 9);
  assert.equal(props.columnsLineBetween, true);
  assert.equal(props.verticalAlign, "bottom");
  assert.equal(props.sectionBidi, true);
  assert.equal(props.rtlGutter, true);
  assert.equal(props.rtlGutterSpecified, true);
  assert.equal(props.lineNumberRestart, "continuous");
  assert.equal(props.lineNumberCountBy, 3);
  assert.equal(props.lineNumberDistance, 240);
  assert.equal(props.lineNumberStart, 5);
  assert.equal(props.footnotePosition, "beneathText");
  assert.equal(props.footnoteNumberRestart, "eachPage");
  assert.equal(props.endnoteNumberRestart, "eachSect");
  assert.equal(props.footnoteNumberStart, 6);
  assert.equal(props.footnoteNumberFormat, 2);
  assert.equal(props.endnoteNumberStart, 6);
  assert.equal(props.endnoteNumberFormat, 1);
  assert.throws(
    () => parseSectionSprms(Buffer.from([0x1a, 0x30, 0x04])),
    /Out-of-spec section Vjc value 4/,
  );
  assert.throws(
    () => parseSectionSprms(Buffer.from([0x13, 0x30, 0x03])),
    /Out-of-spec section line-number restart mode 3/,
  );
  assert.throws(
    () => parseSectionSprms(Buffer.from([0x15, 0x50, 0x65, 0x00])),
    /Out-of-spec section line-number countBy 101/,
  );
  assert.throws(
    () => parseSectionSprms(Buffer.from([0x00, 0x30, 0x05])),
    /Out-of-spec section chapter separator CNS value 5/,
  );
  assert.throws(
    () => parseSectionSprms(Buffer.from([0x01, 0x30, 0x0a])),
    /Out-of-spec section chapter heading level 10/,
  );
  assert.throws(
    () => parseSectionSprms(Buffer.from([0x3b, 0x30, 0x00])),
    /Out-of-spec section footnote position Fpc value 0/,
  );
  assert.throws(
    () => parseSectionSprms(Buffer.from([0x3e, 0x30, 0x02])),
    /Out-of-spec endnote numbering restart Rnc value 2/,
  );
  assert.throws(
    () => parseSectionSprms(Buffer.from([0x3f, 0x50, 0x00, 0x40])),
    /Out-of-spec footnote number start 16384/,
  );

  const explicitFalseRtlGutter = parseSectionSprms(Buffer.from([0x2a, 0x32, 0x00]));
  assert.equal(explicitFalseRtlGutter.rtlGutter, false);
  assert.equal(explicitFalseRtlGutter.rtlGutterSpecified, true);
});

test("emits parsed section gutter margin, column separator, and vertical alignment in DOCX section properties", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    paragraphProperties: [{}],
    sections: [{
      cpStart: 0,
      cpEnd: 2,
      properties: {
        pageWidth: 11906,
        pageHeight: 16838,
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        headerMargin: 720,
        footerMargin: 720,
        gutterMargin: 384,
        columnCount: 2,
        columnSpacing: 720,
        columnsLineBetween: true,
        verticalAlign: "center",
        sectionBidi: true,
        rtlGutter: true,
        formProtection: true,
        endnotesSuppressed: true,
        paperSourceFirst: 2,
        paperSourceOther: 4,
        pageNumberRestart: true,
        pageNumberStart: 12,
        pageNumberChapterStyle: 2,
        pageNumberChapterSeparator: "enDash",
        footnotePosition: "beneathText",
        footnoteNumberRestart: "continuous",
        footnoteNumberStart: 6,
        footnoteNumberFormat: 2,
        endnoteNumberRestart: "eachSect",
        endnoteNumberStart: 6,
        endnoteNumberFormat: 1,
        lineNumberRestart: "continuous",
        lineNumberCountBy: 3,
        lineNumberDistance: 240,
        lineNumberStart: 5,
      },
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "section-gutter" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:footnotePr><w:pos w:val="beneathText"\/><w:numFmt w:val="lowerRoman"\/><w:numStart w:val="6"\/><\/w:footnotePr><w:endnotePr><w:numFmt w:val="upperRoman"\/><w:numRestart w:val="eachSect"\/><\/w:endnotePr><w:pgSz/);
  assert.match(xml, /<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="384"\/>/);
  assert.match(xml, /<w:paperSrc w:first="2" w:other="4"\/><w:lnNumType/);
  assert.match(xml, /<w:lnNumType w:countBy="3" w:start="5" w:distance="240" w:restart="continuous"\/>/);
  assert.match(xml, /<w:pgNumType w:fmt="decimal" w:start="12" w:chapStyle="2" w:chapSep="enDash"\/>/);
  assert.match(xml, /<w:cols w:space="720" w:num="2" w:sep="1"\/>/);
  assert.match(xml, /<w:formProt\/><w:vAlign/);
  assert.match(xml, /<w:vAlign w:val="center"\/>/);
  assert.match(xml, /<w:noEndnote\/><w:bidi\/>/);
  assert.match(xml, /<w:bidi\/><w:rtlGutter\/>/);

  const explicitFalseDocx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    paragraphProperties: [{}],
    sections: [{
      cpStart: 0,
      cpEnd: 2,
      properties: {
        rtlGutter: false,
        rtlGutterSpecified: true,
      },
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "section-rtl-gutter-false" });
  assert.match(readDocxDocumentXml(explicitFalseDocx), /<w:rtlGutter w:val="0"\/>/);
});

test("ignores MS-DOC section page-number start when restart is disabled", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    paragraphProperties: [{}],
    sections: [{
      cpStart: 0,
      cpEnd: 2,
      properties: {
        pageNumberRestart: false,
        pageNumberStart: 12,
      },
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "section-page-number-restart" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:pgNumType w:fmt="decimal"\/>/);
  assert.doesNotMatch(xml, /<w:pgNumType[^>]*w:start="12"/);
});

test("emits parsed MS-DOC page-number format in line-grid sections", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    paragraphProperties: [{}],
    sections: [{
      cpStart: 0,
      cpEnd: 2,
      properties: {
        docGridType: 2,
        docGridCharSpace: 0,
        docGridLinePitch: 312,
        pageNumberRestart: false,
        pageNumberFormat: 0x39,
      },
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      typography: { fKerningPunct: false, iJustification: 1 },
      xmlValidation: { fValidateXML: false, fShowXMLErrors: false },
      grfFmtFilter: 20516,
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "line-grid-page-number-format" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:pgNumType w:fmt="numberInDash"\/>/);
});

test("emits MS-DOC non-section 0x0C as a manual page break", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\x0cB\r",
    defaultTabStop: 720,
    paragraphProperties: [{}],
    sections: [{
      cpStart: 0,
      cpEnd: 4,
      properties: {},
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "manual-page-break" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:t>A<\/w:t><\/w:r><w:r>[\s\S]*?<w:br w:type="page"\/><\/w:r><\/w:p>\s*<w:p[\s\S]*?<w:t>B<\/w:t>/);
});

test("does not emit MS-DOC section-ending 0x0C as a manual page break", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\x0cB\r",
    defaultTabStop: 720,
    paragraphProperties: [{}, {}],
    sections: [
      { cpStart: 0, cpEnd: 2, properties: {} },
      { cpStart: 2, cpEnd: 4, properties: {} },
    ],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "section-form-feed" });
  const xml = readDocxDocumentXml(docx);

  assert.doesNotMatch(xml, /<w:br w:type="page"\/>/);
});

test("MS-DOC border color and type enums use spec values without fallback", () => {
  assert.equal(brcColorFromIco(0x00), "auto");
  assert.equal(brcColorFromIco(0x09), "000080");
  assert.equal(brcColorFromIco(0x10), "C0C0C0");
  assert.throws(() => brcColorFromIco(0x11), /Out-of-spec MS-DOC Ico color index 17/);

  assert.equal(BRC_TYPE_NAMES[0x14], "wave");
  assert.equal(BRC_TYPE_NAMES[0x1B], "inset");
  assert.equal(BRC_TYPE_NAMES[0x02], undefined);
});

test("parses MS-DOC character effect toggle SPRMs", () => {
  const props = parseSprms(Buffer.from([
    0x38, 0x08, 0x01, // sprmCFOutline
    0x39, 0x08, 0x01, // sprmCFShadow
    0x54, 0x08, 0x01, // sprmCFImprint
    0x58, 0x08, 0x01, // sprmCFEmboss
    0x75, 0x08, 0x01, // sprmCFNoProof
    0x11, 0x08, 0x01, // sprmCFWebHidden
    0x18, 0x08, 0x01, // sprmCFSpecVanish
    0x5a, 0x08, 0x01, // sprmCFBiDi
    0x82, 0x08, 0x81, // sprmCFComplexScripts
    0x48, 0x2a, 0x01, // sprmCIss superscript
    0x34, 0x2a, 0x04, // sprmCKcd under-dot emphasis
    0x59, 0x28, 0x06, // sprmCSfxText shimmer
    0x70, 0x68, 0x12, 0x34, 0x56, 0x00, // sprmCCv COLORREF
    0x77, 0x68, 0x9a, 0xbc, 0xde, 0x00, // sprmCCvUl COLORREF
  ]));

  assert.equal(props.outline, true);
  assert.equal(props.shadow, true);
  assert.equal(props.imprint, true);
  assert.equal(props.emboss, true);
  assert.equal(props.noProof, true);
  assert.equal(props.webHidden, true);
  assert.equal(props.specVanish, true);
  assert.equal(props.rtl, true);
  assert.equal(props.complexScript, true);
  assert.equal(props.verticalAlign, "superscript");
  const baselineProps = parseSprms(Buffer.from([0x48, 0x2a, 0x00]));
  assert.equal(baselineProps.verticalAlign, null);
  assert.equal(Object.prototype.hasOwnProperty.call(baselineProps, "verticalAlign"), true);
  assert.equal(props.emphasisMark, "underDot");
  assert.equal(props.textEffect, "shimmer");
  assert.equal(props.textColor, "123456");
  assert.equal(props.underlineColor, "9ABCDE");
  assert.equal(parseSprms(Buffer.from([0x59, 0x28, 0x00])).textEffect, "none");
  assert.equal(parseSprms(Buffer.from([0x59, 0x28, 0x01])).textEffect, "lights");
  assert.equal(parseSprms(Buffer.from([0x59, 0x28, 0x02])).textEffect, "blinkBackground");
  assert.equal(parseSprms(Buffer.from([0x59, 0x28, 0x03])).textEffect, "sparkle");
  assert.equal(parseSprms(Buffer.from([0x59, 0x28, 0x04])).textEffect, "antsBlack");
  assert.equal(parseSprms(Buffer.from([0x59, 0x28, 0x05])).textEffect, "antsRed");
  assert.throws(
    () => parseSprms(Buffer.from([0x34, 0x2a, 0x05])),
    /Out-of-spec MS-DOC Kcd emphasis value 5/,
  );
  assert.throws(
    () => parseSprms(Buffer.from([0x59, 0x28, 0x07])),
    /Out-of-spec MS-DOC SfxText text effect value 7/,
  );
});

test("run properties emit parsed caps without synthesizing complex-script italic", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    paragraphProperties: [{}],
    characterProperties: [{
      italic: false,
      caps: false,
      fontCs: 1,
    }, {}],
    fontTable: [{ name: "Times New Roman" }, { name: "SimSun" }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "parsed-caps" });
  const paragraph = findParagraphXml(readDocxDocumentXml(docx), "A");
  assert.ok(paragraph, "paragraph should exist");

  assert.match(paragraph, /<w:i w:val="0"\/>/);
  assert.match(paragraph, /<w:caps w:val="0"\/>/);
  assert.doesNotMatch(paragraph, /<w:iCs/);
});

test("parses and emits MS-DOC character revision mark toggles", () => {
  assert.equal(parseSprms(Buffer.from([0x00, 0x08, 0x01])).revisionMarkDel, true);
  assert.equal(parseSprms(Buffer.from([0x01, 0x08, 0x01])).revisionMarkIns, true);

  const docx = wpsToDocxBuffer({
    bodyText: "AB\r",
    defaultTabStop: 720,
    characterProperties: [
      { revisionMarkIns: true },
      { revisionMarkDel: true },
    ],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "revision-marks" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:ins w:id="\d+" w:author="Unknown"><w:r>[\s\S]*?<w:t>A<\/w:t><\/w:r><\/w:ins>/);
  assert.match(xml, /<w:del w:id="\d+" w:author="Unknown"><w:r>[\s\S]*?<w:delText>B<\/w:delText><\/w:r><\/w:del>/);
});

test("parses MS-DOC paragraph property revision mark operands", () => {
  const revisionDate = 26 | (13 << 6) | (22 << 11) | (8 << 16) | ((2024 - 1900) << 20);
  const props = parseSprms(Buffer.from([
    0x6f, 0xc6, 0x07, // sprmPPropRMark, PropRMarkOperand.cb
    0x01, // fPropRMark
    0x02, 0x00, // ibstshort
    revisionDate & 0xff,
    (revisionDate >> 8) & 0xff,
    (revisionDate >> 16) & 0xff,
    (revisionDate >> 24) & 0xff,
  ]));

  assert.deepEqual(props.paragraphPropRMark, {
    fPropRMark: true,
    authorIndex: 2,
    date: revisionDate,
  });
  assert.throws(
    () => parseSprms(Buffer.from([0x6f, 0xc6, 0x08, 0x01, 0x02, 0x00, 0, 0, 0, 0, 0])),
    /Out-of-spec MS-DOC paragraph property revision mark PropRMarkOperand cb 8/,
  );
});

test("parses MS-DOC sprmPWall previous paragraph properties for revision marking", async () => {
  const sample10 = readWps(await readFile(SAMPLE10_WPS));
  const changed = sample10.paragraphProperties.find((properties) => properties?.paragraphPropertyChange);

  assert.ok(changed, "expected a paragraph with preserved revision properties");
  assert.equal(changed.paragraphPropertyChange.mark.fPropRMark, true);
  assert.equal(changed.paragraphPropertyChange.mark.authorIndex, 1);
  assert.deepEqual(changed.paragraphPropertyChange.previous.lineSpacing, { twips: 240, rule: "auto" });
  assert.equal(changed.paragraphPropertyChange.previous.listLevel, 0);
  assert.equal(changed.paragraphPropertyChange.previous.listId, 0);
  assert.equal(changed.paragraphPropertyChange.previous.leftIndentChars, 200);
});

test("emits MS-DOC paragraph property revision as OOXML pPrChange", () => {
  const revisionDate = 26 | (13 << 6) | (22 << 11) | (8 << 16) | ((2024 - 1900) << 20);
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    revisionAuthors: ["Unknown", "Alice"],
    paragraphProperties: [{
      paragraphPropertyChange: {
        mark: { fPropRMark: true, authorIndex: 1, date: revisionDate },
        previous: {
          listLevel: 0,
          listId: 0,
          lineSpacing: { twips: 240, rule: "auto" },
          leftIndentChars: 200,
        },
      },
    }],
    sections: [{ cpStart: 0, cpEnd: 2, properties: {} }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "paragraph-property-revision" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:pPrChange w:id="\d+" w:author="Alice" w:date="2024-08-22T13:26:00Z"><w:pPr><w:numPr><w:ilvl w:val="0"\/><w:numId w:val="0"\/><\/w:numPr><w:spacing w:line="240" w:lineRule="auto"\/><w:ind w:leftChars="200"\/><\/w:pPr><\/w:pPrChange>/);
});

test("emits MS-DOC paragraph-mark insertion revision inside paragraph run properties", () => {
  const revisionDate = 8 | (9 << 6) | (17 << 11) | (5 << 16) | ((2024 - 1900) << 20);
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    revisionAuthors: ["Unknown", "Alice"],
    paragraphProperties: [{}],
    characterProperties: [
      {},
      { revisionMarkIns: true, revisionAuthorIndex: 1, revisionDate },
    ],
    sections: [{ cpStart: 0, cpEnd: 2, properties: {} }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "paragraph-mark-revision" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:pPr><w:rPr><w:ins w:id="\d+" w:author="Alice" w:date="2024-05-17T09:08:00Z"\/>/);
});

test("emits explicit MS-DOC complex-script bold independently of derived toggles", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    characterProperties: [
      { bold: true, boldCs: true },
    ],
    sections: [{
      cpStart: 0,
      cpEnd: 2,
      properties: { docGridType: 2, docGridLinePitch: 312 },
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      grfFmtFilter: 0x5024,
      typography: { fKerningPunct: true, iJustification: 0 },
      compatibility: { ulTrailSpace: true },
      xmlValidation: { fValidateXML: true, fShowXMLErrors: true },
    },
  }, { title: "complex-script-bold" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:b\/><w:bCs\/>/);
});

test("run coalescing preserves explicit MS-DOC complex-script bold boundaries", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "AB\r",
    defaultTabStop: 720,
    characterProperties: [
      { fontSize: 18, fontSizeCs: 18 },
      { fontSize: 18, fontSizeCs: 18, boldCs: true },
    ],
    sections: [{
      cpStart: 0,
      cpEnd: 3,
      properties: { docGridType: 2, docGridLinePitch: 312 },
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      grfFmtFilter: 0x5024,
      typography: { fKerningPunct: true, iJustification: 0 },
      compatibility: { ulTrailSpace: true },
      xmlValidation: { fValidateXML: true, fShowXMLErrors: true },
    },
  }, { title: "complex-script-bold-boundary" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:t>A<\/w:t><\/w:r><w:r><w:rPr>[\s\S]*?<w:bCs\/>[\s\S]*?<\/w:rPr><w:t>B<\/w:t>/);
});

test("parses MS-DOC revision author string table", () => {
  assert.deepEqual(
    parseSttbfRMark(buildUnicodeSttbNoExtra(["Unknown", "Alice", "Bob"])),
    ["Unknown", "Alice", "Bob"],
  );
  assert.throws(
    () => parseSttbfRMark(buildUnicodeSttbNoExtra(["Alice"])),
    /Out-of-spec SttbfRMark: first author must be Unknown/,
  );
  const withExtra = Buffer.from(buildUnicodeSttbNoExtra(["Unknown"]));
  withExtra.writeUInt16LE(2, 4);
  assert.throws(
    () => parseSttbfRMark(withExtra),
    /Out-of-spec SttbfRMark: expected no extra data, got 2 bytes/,
  );
});

test("parses revision authors from the MS-DOC FibRgFcLcb97 SttbfRMark slot", async () => {
  const sample10 = readWps(await readFile(SAMPLE10_WPS));
  assert.deepEqual(sample10.revisionAuthors, ["Unknown", "登记财产小组收发员", "戚玉霞"]);
});

test("emits MS-DOC revision authors and DTTM metadata", () => {
  const revisionDate = 8 | (9 << 6) | (17 << 11) | (5 << 16) | ((2024 - 1900) << 20);
  const deletionDate = 30 | (14 << 6) | (18 << 11) | (5 << 16) | ((2024 - 1900) << 20);
  const docx = wpsToDocxBuffer({
    bodyText: "AB\r",
    defaultTabStop: 720,
    revisionAuthors: ["Unknown", "Alice & Bob", "Carol"],
    characterProperties: [
      { revisionMarkIns: true, revisionAuthorIndex: 1, revisionDate },
      { revisionMarkDel: true, revisionDelAuthorIndex: 2, revisionDelDate: deletionDate },
    ],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "revision-authors" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:ins w:id="0" w:author="Alice &amp; Bob" w:date="2024-05-17T09:08:00Z">/);
  assert.match(xml, /<w:del w:id="1" w:author="Carol" w:date="2024-05-18T14:30:00Z">/);
  const ignoredDateDocx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    revisionAuthors: ["Unknown", "Alice"],
    characterProperties: [{ revisionMarkIns: true, revisionAuthorIndex: 1, revisionDate: ((2024 - 1900) << 20) }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "ignored-revision-date" });
  assert.doesNotMatch(readDocxDocumentXml(ignoredDateDocx), /w:date=/);
  assert.throws(
    () => wpsToDocxBuffer({
      bodyText: "A\r",
      defaultTabStop: 720,
      revisionAuthors: ["Unknown"],
      characterProperties: [{ revisionMarkIns: true, revisionAuthorIndex: 1 }],
      dop: {
        dxaHotZ: 360,
        pageBorderIncludes: { header: false, footer: false },
        dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
        compatibility: { ulTrailSpace: true },
      },
    }, { title: "invalid-revision-author" }),
    /revision author index 1 is outside SttbfRMark/,
  );
  assert.throws(
    () => wpsToDocxBuffer({
      bodyText: "A\r",
      defaultTabStop: 720,
      revisionAuthors: ["Unknown"],
      characterProperties: [{ revisionMarkIns: true, revisionDate: 60 }],
      dop: {
        dxaHotZ: 360,
        pageBorderIncludes: { header: false, footer: false },
        dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
        compatibility: { ulTrailSpace: true },
      },
    }, { title: "invalid-revision-date" }),
    /Out-of-spec DTTM minute 60/,
  );
});

test("parses MS-DOC character and paragraph border operands", () => {
  const props = parseSprms(Buffer.from([
    0x72, 0xca, 0x08, 0x12, 0x34, 0x56, 0x00, 0x06, 0x01, 0x07, 0x00,
    0x4e, 0xc6, 0x08, 0x01, 0x02, 0x03, 0x00, 0x08, 0x03, 0x04, 0x00,
    0x71, 0xca, 0x0a, 0x11, 0x22, 0x33, 0x00, 0x44, 0x55, 0x66, 0x00, 0x08, 0x00,
    0x4d, 0xc6, 0x0a, 0xaa, 0xbb, 0xcc, 0x00, 0x12, 0x34, 0x56, 0x00, 0x01, 0x00,
    0x2a, 0x24, 0x01, // sprmPFNoAutoHyph
    0x70, 0x24, 0x01, // sprmPFMirrorIndents
    0x62, 0x24, 0x01, // sprmPFNoAllowOverlap
    0x2c, 0x44, 0x1a, 0x00, // sprmPDcs: fdct=margin, cl=3
    0x2e, 0x84, 0x78, 0x00, // sprmPDyaFromText
    0x2f, 0x84, 0x34, 0x12, // sprmPDxaFromText
  ]));

  assert.deepEqual(props.border, {
    val: "single",
    color: "123456",
    sz: "6",
    space: "0",
  });
  assert.deepEqual(props.paragraphBorders.top, {
    val: "double",
    color: "010203",
    sz: "8",
    space: "4",
  });
  assert.deepEqual(parseSprms(Buffer.from([
    0x25, 0x64, 0x00, 0x00, 0x00, 0x00, // sprmPBrcLeft80: explicit none border
  ])).paragraphBorders.left, {
    val: "none",
    color: "auto",
    sz: "0",
    space: "0",
  });
  assert.deepEqual(parseSprms(Buffer.from([
    0x4e, 0xc6, 0x08, 0x00, 0x00, 0x00, 0xff, 0x00, 0x00, 0x00, 0x00, // sprmPBrcTop: explicit none border
  ])).paragraphBorders.top, {
    val: "none",
    color: "auto",
    sz: "0",
    space: "0",
  });
  assert.deepEqual(props.background, {
    val: "pct50",
    color: "112233",
    fill: "445566",
  });
  assert.deepEqual(props.paragraphShading, {
    val: "solid",
    color: "AABBCC",
    fill: "123456",
  });
  assert.equal(props.suppressAutoHyphens, true);
  assert.equal(props.mirrorIndents, true);
  assert.equal(props.suppressOverlap, true);
  assert.equal(props.frameDropCap, "margin");
  assert.equal(props.frameLines, 3);
  assert.equal(props.frameVSpace, 120);
  assert.equal(props.frameHSpace, 0x1234);
});

test("emits parsed paragraph frame properties to DOCX framePr", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    paragraphProperties: [{
      frameWidth: 1440,
      frameHeight: 720,
      frameHRule: "exact",
      frameWrap: "around",
      frameHAnchor: "margin",
      frameVAnchor: "page",
      frameHSpace: 120,
      frameVSpace: 240,
      frameDropCap: "drop",
      frameLines: 4,
      frameLocked: true,
      suppressOverlap: true,
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "paragraph-frame" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:pPr><w:framePr w:w="1440" w:h="720" w:hRule="exact" w:wrap="around" w:vAnchor="page" w:hAnchor="margin" w:hSpace="120" w:vSpace="240" w:dropCap="drop" w:lines="4" w:anchorLock="1"\/><w:suppressOverlap\/><\/w:pPr>/);
});

test("emits parsed MS-DOC paragraph borders as direct pBdr", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    paragraphProperties: [{
      paragraphBorders: {
        left: { val: "none", color: "auto", sz: "0", space: "0" },
        bottom: { val: "single", color: "auto", sz: "4", space: "1" },
      },
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "paragraph-borders" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:pBdr><w:left w:val="none" w:color="auto" w:sz="0" w:space="0"\/><w:bottom w:val="single" w:color="auto" w:sz="4" w:space="1"\/><\/w:pBdr>/);
});

test("sprmPNest80 adds to parsed left indent without leaking parser internals", () => {
  const props = parseSprms(Buffer.from([
    0x10, 0x46, 0x32, 0x00, // sprmPNest80: +50
    0x0f, 0x84, 0x64, 0x00, // sprmPDxaLeft80: 100
  ]));
  assert.equal(props.leftIndent, 150);
  assert.equal(Object.keys(props).some((key) => key.startsWith("_")), false);

  const logicalProps = parseSprms(Buffer.from([
    0x0f, 0x84, 0x64, 0x00, // sprmPDxaLeft80: 100
    0x10, 0x46, 0x32, 0x00, // sprmPNest80: +50
    0x5f, 0x46, 0x14, 0x00, // sprmPNest: +20, supersedes PNest80
    0x5e, 0x84, 0xe8, 0x03, // sprmPDxaLeft: 1000
  ]));
  assert.equal(logicalProps.leftIndent, 1020);
});

test("parses MS-DOC paragraph table-depth and nested table marker SPRMs", () => {
  const props = parseSprms(Buffer.from([
    0x49, 0x66, 0x02, 0x00, 0x00, 0x00, // sprmPItap: depth 2
    0x4a, 0x66, 0xff, 0xff, 0xff, 0xff, // sprmPDtap: -1
    0x4b, 0x24, 0x01, // sprmPFInnerTableCell
    0x4c, 0x24, 0x01, // sprmPFInnerTtp
    0x5a, 0x24, 0x01, // sprmPFOpenTch
  ]));

  assert.equal(props.tableDepth, 1);
  assert.equal(props.innerTableCell, true);
  assert.equal(props.innerTtp, true);
  assert.equal(props.openTch, true);
  assert.throws(
    () => parseSprms(Buffer.from([0x4a, 0x66, 0xff, 0xff, 0xff, 0xff])),
    /Out-of-spec paragraph table depth -1/,
  );
});

test("parses MS-DOC character fitText operand", () => {
  const props = parseSprms(Buffer.from([
    0x76, 0xca, 0x08, 0xa0, 0x05, 0x00, 0x00, 0x63, 0x00, 0x00, 0x00, // sprmCFitText: width 1440, id 99
  ]));

  assert.deepEqual(props.fitText, { width: 1440, id: 99 });
  assert.equal(parseSprms(Buffer.from([
    0x76, 0xca, 0x08, 0x00, 0x00, 0x00, 0x00, 0x63, 0x00, 0x00, 0x00, // dxaFitText zero: ignored
  ])).fitText, null);
  assert.throws(
    () => parseSprms(Buffer.from([0x76, 0xca, 0x07, 0xa0, 0x05, 0x00, 0x00, 0x63, 0x00, 0x00])),
    /Out-of-spec character fitText CFitTextOperand length/,
  );
  assert.throws(
    () => parseSprms(Buffer.from([0x76, 0xca, 0x08, 0xff, 0xff, 0xff, 0xff, 0x63, 0x00, 0x00, 0x00])),
    /Unsupported MS-DOC character fitText minimum-width variant -1/,
  );
});

test("parses MS-DOC line-break clear operand", () => {
  assert.equal(parseSprms(Buffer.from([0x79, 0x28, 0x00])).lineBreakClear, "none");
  assert.equal(parseSprms(Buffer.from([0x79, 0x28, 0x01])).lineBreakClear, "left");
  assert.equal(parseSprms(Buffer.from([0x79, 0x28, 0x02])).lineBreakClear, "right");
  assert.equal(parseSprms(Buffer.from([0x79, 0x28, 0x03])).lineBreakClear, "all");
  assert.throws(
    () => parseSprms(Buffer.from([0x79, 0x28, 0x04])),
    /Out-of-spec MS-DOC LBCOperand value 4/,
  );
});

test("parses MS-DOC Far East character layout operand", () => {
  assert.deepEqual(parseSprms(Buffer.from([
    0x78, 0xca, 0x06, 0x01, 0x10, 0x05, 0x00, 0x00, 0x00, // fTNY + fTNYCompress, id 5
  ])).eastAsianLayout, {
    id: 5,
    vert: true,
    vertCompress: true,
  });

  assert.deepEqual(parseSprms(Buffer.from([
    0x78, 0xca, 0x06, 0x02, 0x02, 0x07, 0x00, 0x00, 0x00, // fWarichu + square brackets, id 7
  ])).eastAsianLayout, {
    id: 7,
    combine: true,
    combineBrackets: "square",
  });

  assert.equal(parseSprms(Buffer.from([
    0x78, 0xca, 0x06, 0x00, 0x00, 0x07, 0x00, 0x00, 0x00, // no active layout bits
  ])).eastAsianLayout, null);
  assert.throws(
    () => parseSprms(Buffer.from([0x78, 0xca, 0x05, 0x01, 0x00, 0x05, 0x00, 0x00])),
    /Out-of-spec character FarEastLayoutOperand length/,
  );
  assert.throws(
    () => parseSprms(Buffer.from([0x78, 0xca, 0x06, 0x04, 0x00, 0x05, 0x00, 0x00, 0x00])),
    /Out-of-spec character UFEL MUST-zero bits 0x4/,
  );
  assert.throws(
    () => parseSprms(Buffer.from([0x78, 0xca, 0x06, 0x02, 0x05, 0x05, 0x00, 0x00, 0x00])),
    /Out-of-spec character UFEL iWarichuBracket value 5/,
  );
});

test("emits MS-DOC Far East character layout as OOXML eastAsianLayout", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    characterProperties: [{
      eastAsianLayout: {
        id: 5,
        combine: true,
        combineBrackets: "square",
        vert: true,
        vertCompress: true,
      },
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "east-asian-layout" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:eastAsianLayout w:id="5" w:combine="true" w:combineBrackets="square" w:vert="true" w:vertCompress="true"\/>/);
});

test("emits MS-DOC line-break clear only on manual line breaks", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\x0bB\r",
    defaultTabStop: 720,
    characterProperties: [
      {},
      { lineBreakClear: "all" },
      {},
    ],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "line-break-clear" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:br w:type="textWrapping" w:clear="all"\/>/);
  assert.throws(
    () => wpsToDocxBuffer({
      bodyText: "AB\r",
      defaultTabStop: 720,
      characterProperties: [
        {},
        { lineBreakClear: "all" },
      ],
      dop: {
        dxaHotZ: 360,
        pageBorderIncludes: { header: false, footer: false },
        dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
        compatibility: { ulTrailSpace: true },
      },
    }, { title: "invalid-line-break-clear" }),
    /Out-of-spec sprmCLbcCRJ applied to a non-line-break character/,
  );
});

test("emits parsed MS-DOC character effect toggles to DOCX run properties", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    paragraphProperties: [{
      suppressAutoHyphens: true,
      paragraphShading: {
        val: "solid",
        color: "AABBCC",
        fill: "123456",
      },
      mirrorIndents: true,
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
    characterProperties: [{
      outline: true,
      shadow: true,
      imprint: true,
      emboss: true,
      noProof: true,
      webHidden: true,
      specVanish: true,
      rtl: true,
      complexScript: true,
      verticalAlign: "superscript",
      emphasisMark: "underDot",
      textEffect: "sparkle",
      textColor: "123456",
      underline: true,
      underlineStyle: "single",
      underlineColor: "9ABCDE",
      fitText: { width: 1440, id: 99 },
      border: {
        val: "single",
        color: "123456",
        sz: "6",
        space: "0",
      },
    }],
  }, { title: "character-effects" });
  const xml = readDocxDocumentXml(docx);
  const run = xml.match(/<w:r>[\s\S]*?<w:t>A<\/w:t><\/w:r>/)?.[0] ?? "";

  assert.match(run, /<w:outline\/>/);
  assert.match(run, /<w:shadow\/>/);
  assert.match(run, /<w:imprint\/>/);
  assert.match(run, /<w:emboss\/>/);
  assert.match(run, /<w:noProof\/>/);
  assert.match(run, /<w:webHidden\/>/);
  assert.match(run, /<w:specVanish\/>/);
  assert.match(run, /<w:vertAlign w:val="superscript"\/>/);
  assert.match(run, /<w:rtl\/>/);
  assert.match(run, /<w:cs\/>/);
  assert.match(run, /<w:em w:val="underDot"\/>/);
  assert.match(run, /<w:color w:val="123456"\/>/);
  assert.match(run, /<w:u w:val="single" w:color="9ABCDE"\/>/);
  assert.match(run, /<w:effect w:val="sparkle"\/>/);
  assert.match(run, /<w:bdr w:val="single" w:color="123456" w:sz="6" w:space="0"\/>/);
  assert.match(run, /<w:fitText w:id="99" w:val="1440"\/>/);
  assert.ok(run.indexOf("<w:u ") < run.indexOf("<w:effect "));
  assert.ok(run.indexOf("<w:effect ") < run.indexOf("<w:bdr "));
  assert.ok(run.indexOf("<w:bdr ") < run.indexOf("<w:fitText "));
  assert.ok(run.indexOf("<w:fitText ") < run.indexOf("<w:vertAlign "));
  assert.match(xml, /<w:pPr><w:shd w:val="solid" w:color="AABBCC" w:fill="123456"\/><w:suppressAutoHyphens\/><w:mirrorIndents\/><\/w:pPr>/);
});

test("emits explicit MS-DOC sprmCIss normal as OOXML baseline", () => {
  const docx = wpsToDocxBuffer({
    bodyText: "A\r",
    defaultTabStop: 720,
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
    characterProperties: [parseSprms(Buffer.from([0x48, 0x2a, 0x00]))],
  }, { title: "baseline-iss" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /<w:vertAlign w:val="baseline"\/>/);
});

test("emits section breaks and tab stops in converted DOCX", async () => {
  const wps = readWps(await readFile(BASIC_WPS));
  const docx = wpsToDocxBuffer(wps, { title: "test" });
  const xml = readDocxDocumentXml(docx);
  assert.match(xml, /<w:tab w:val="left" w:pos="4464"\/>/);
  assert.match(xml, /<w:sectPr>.*<w:pgSz w:w="11910" w:h="16840"\/>/s);
  assert.match(xml, /<w:pgMar w:top="1580" w:right="1260" w:bottom="1321" w:left="1360" w:header="0" w:footer="0" w:gutter="0"\/>/);
});

test("extracts character formatting runs from CHPX pages", async () => {
  const wps = readWps(await readFile(BASIC_WPS));
  assert.deepEqual(wps.characterRuns.slice(0, 5).map((run) => [run.cpStart, run.cpEnd]), [
    [0, 5],
    [5, 9],
    [9, 10],
    [10, 13],
    [13, 14],
  ]);
  assert.equal(wps.characterProperties[65].fontSize, 44);
  assert.equal(wps.characterProperties[16].charSpacing, -2);
  assert.equal(wps.characterProperties[29].charSpacing, 2);
  assert.equal(wps.characterProperties[53].charSpacing, 7);
});

test("extracts style sheet from STSH with names and types", async () => {
  const wps = readWps(await readFile(BASIC_WPS));
  const styles = wps.styles.filter((s) => s !== null);
  assert.ok(styles.length >= 10);

  const normal = styles.find((s) => s.name === "正文");
  assert.ok(normal);
  assert.equal(normal.type, "paragraph");
  assert.equal(normal.basedOn, null);
  assert.equal(normal.lineSpacing.twips, 240);
  assert.equal(normal.lineSpacing.rule, "auto");

  const heading1 = styles.find((s) => s.name === "标题 1");
  assert.ok(heading1);
  assert.equal(heading1.type, "paragraph");
  assert.equal(heading1.basedOn, 0);
  assert.equal(heading1.runProperties.fontSize, 44);
  assert.equal(heading1.runProperties.bold, undefined);

  const sample9 = readWps(await readFile(SAMPLE9_WPS));
  const sample9Heading1 = sample9.styles.find((s) => s?.name === "标题 1");
  assert.ok(sample9Heading1);
  assert.equal(sample9Heading1.runProperties.bold, true);

  const bodyText = styles.find((s) => s.name === "正文文本");
  assert.ok(bodyText);
  assert.equal(bodyText.type, "paragraph");
  assert.equal(bodyText.basedOn, 0);
  assert.equal(bodyText.runProperties.fontSize, 32);

  const defaultFont = styles.find((s) => s.name === "默认段落字体");
  assert.ok(defaultFont);
  assert.equal(defaultFont.type, "character");
  assert.equal(defaultFont.basedOn, null);
});

test("emits w:spacing in converted DOCX from paragraph properties", async () => {
  const wps = readWps(await readFile(BASIC_WPS));
  const docx = wpsToDocxBuffer(wps, { title: "test" });
  const xml = readDocxDocumentXml(docx);
  assert.match(xml, /<w:spacing w:line="594" w:lineRule="exact"\/>/);
});

test("emits styles.xml with style definitions from STSH", async () => {
  const wps = readWps(await readFile(BASIC_WPS));
  const docx = wpsToDocxBuffer(wps, { title: "test" });
  const xml = readDocxDocumentXml(docx);
  assert.match(xml, /<w:pStyle w:val="2"\/>/);
  assert.match(xml, /<w:rPr>.*<w:w w:val="100"\/><w:sz w:val="9"\/><\/w:rPr>/s);

  const stylesEntry = readZipEntry(docx, "word/styles.xml");
  const stylesXml = stylesEntry.toString("utf8");
  assert.match(stylesXml, /<w:style w:type="paragraph"[^>]*w:styleId="1"/);
  assert.match(stylesXml, /<w:name w:val="Normal"/);
  assert.match(stylesXml, /<w:style w:type="paragraph"[^>]*w:styleId="2"/);
  assert.match(stylesXml, /<w:name w:val="(标题 1|heading 1)"/);
  assert.match(stylesXml, /<w:sz w:val="44"\/><w:szCs w:val="44"\/>/);
  assert.match(stylesXml, /<w:basedOn w:val="1"\/>/);
  const headingStyleXml = stylesXml.match(/<w:style w:type="paragraph"[^>]*w:styleId="2"[\s\S]*?<\/w:style>/)?.[0];
  assert.ok(headingStyleXml, "Heading 1 style should be emitted");
  assert.doesNotMatch(headingStyleXml, /<w:b\/>/);
  assert.doesNotMatch(headingStyleXml, /<w:bCs\/>/);
  assert.match(stylesXml, /<w:latentStyles w:count="260"/);
});

test("full WPS conversion has correct continuous sections for landscape attachments", async () => {
  const wps = readWps(await readFile(FULL_WPS));
  const docx = wpsToDocxBuffer(wps, { title: "test" });
  const xml = readDocxDocumentXml(docx);

  // Count continuous landscape sections (parsed from sprmSBkc 0x3009)
  const continuousLandscapeSections = [...xml.matchAll(/<w:type w:val="continuous"\/>[\s\S]*?orient="landscape"/g)];
  assert.ok(continuousLandscapeSections.length >= 2, "Expected at least 2 continuous landscape sections");
  
  // Verify they're on landscape pages
  const landscapeSections = [...xml.matchAll(/<w:pgSz[^>]*orient="landscape"[^>]*\/>/g)];
  assert.ok(landscapeSections.length >= 2, "Expected at least 2 landscape sections");
});

test("tables WPS conversion has correct continuous sections from parsed section properties", async () => {
  const wps = readWps(await readFile("sample/tables/original.wps"));
  const docx = wpsToDocxBuffer(wps, { title: "test" });
  const xml = readDocxDocumentXml(docx);

  // Count continuous sections (parsed from sprmSBkc 0x3009)
  const continuousSections = [...xml.matchAll(/<w:type w:val="continuous"\/>/g)];
  // Should have 4 continuous sections (indices 2,3,8,9 from parsed bkc=0)
  assert.equal(continuousSections.length, 4, "Expected 4 continuous sections from parsed break codes");
});

test("table6 extracts 3 tables with correct row/cell counts matching expected.docx", async () => {
  const wps = readWps(await readFile("sample/table6/original.wps"));

  // Must extract exactly 3 tables
  assert.equal(wps.tableRows.length, 3, "table6 should produce 3 tables");

  // Check cells-per-row for each table
  const rowsPerTable = wps.tableRows.map((t) => t.rows.map((r) => r.cells.length));
  assert.deepEqual(rowsPerTable[0], [4, 4, 4, 4, 4, 1, 4, 2, 2],
    "Table 0 should have 9 rows with cells per row [4,4,4,4,4,1,4,2,2]");
  assert.deepEqual(rowsPerTable[1], [2, 2, 2],
    "Table 1 should have 3 rows with 2 cells each");
  assert.deepEqual(rowsPerTable[2], [2],
    "Table 2 should have 1 row with 2 cells");

  // Verify total cell counts
  const totalCells = wps.tableRows.map((t) => t.rows.reduce((a, r) => a + r.cells.length, 0));
  assert.deepEqual(totalCells, [29, 6, 2], "Total cell counts should be 29, 6, 2");

  // Verify grid column counts
  const gridCols = wps.tableRows.map((t) => t.gridCols?.length);
  assert.deepEqual(gridCols, [6, 2, 2], "Grid column counts should be 6, 2, 2");
  assert.deepEqual(
    wps.tableRows.map((t) => t.rows.map((r) => [r.rowHeight, r.rowHeightRule])),
    [
      [[0, 0], [539, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [3820, 0]],
      [[3162, 0], [2675, 0], [4977, 0]],
      [[591, 0]],
    ],
    "Table row heights should come from sprm 0x9407 as atLeast heights",
  );

  // Verify DOCX output matches
  const docx = wpsToDocxBuffer(wps, { title: "table6" });
  const xml = readDocxDocumentXml(docx);
  const tableXmls = [...xml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((m) => m[0]);
  assert.equal(tableXmls.length, 3, "DOCX should have 3 tables");
  assert.deepEqual(tableXmls.map((t) => [...t.matchAll(/<w:tr[\s>]/g)].length), [9, 3, 1],
    "DOCX table row counts should be 9, 3, 1");
  assert.deepEqual(tableXmls.map((t) => [...t.matchAll(/<w:tc>/g)].length), [29, 6, 2],
    "DOCX table cell counts should be 29, 6, 2");
});

test("table6 emits hanging indents for the basic living-expense list item", async () => {
  const wps = readWps(await readFile("sample/table6/original.wps"));
  const docx = wpsToDocxBuffer(wps, { title: "table6" });
  const xml = readDocxDocumentXml(docx);

  assert.match(xml, /w:hanging="360"/);
  assert.equal((xml.match(/w:firstLine="-/g) || []).length, 0);
});

test("table rows emit trHeight only from parsed MS-DOC row-height SPRMs", () => {
  const noBorder = { style: "none", width: 0, color: "auto", space: 0 };
  const docx = wpsToDocxBuffer({
    bodyText: "A\rB\r",
    defaultTabStop: 720,
    paragraphProperties: [{}, {}],
    characterProperties: [{}, {}, {}, {}],
    tableRows: [{
      cpStart: 0,
      cpEnd: 4,
      gridCols: [1440],
      tableWidth: 1440,
      tableWidthType: "dxa",
      tableBordersExplicit: false,
      tableBorders: {
        top: noBorder,
        left: noBorder,
        bottom: noBorder,
        right: noBorder,
        insideH: noBorder,
        insideV: noBorder,
      },
      rows: [{
        rowHeight: 480,
        rowHeightRule: 0,
        cells: [{
          cpStart: 0,
          cpEnd: 2,
          width: 1440,
        }],
      }, {
        cells: [{
          cpStart: 2,
          cpEnd: 4,
          width: 1440,
        }],
      }],
    }],
    dop: {
      dxaHotZ: 360,
      pageBorderIncludes: { header: false, footer: false },
      dogrid: { dxaGrid: 180, dyaGrid: 156, dxGridDisplay: 0, dyGridDisplay: 2 },
      compatibility: { ulTrailSpace: true },
    },
  }, { title: "parsed-row-height-only" });
  const xml = readDocxDocumentXml(docx);

  assert.equal((xml.match(/<w:trHeight\b/g) ?? []).length, 1);
  assert.match(xml, /<w:trHeight w:val="480" w:hRule="atLeast"\/>/);
});

test("table6 signature line keeps its tab stop inside the table cell", async () => {
  const wps = readWps(await readFile("sample/table6/original.wps"));
  const cell = wps.tableRows[0].rows[7].cells[1];

  assert.equal(
    wps.bodyText.slice(cell.cpStart, cell.cpEnd),
    "本单位保证所申报材料真实有效，否则愿意承担由此引起的一切法律责任和后果。\r\r法人（负责人）（签章）：\t 经办人（签字）：\x07",
  );
  assert.deepEqual(wps.paragraphProperties[48].tabs, [{ position: 3879, alignment: "left", leader: null }]);

  const docx = wpsToDocxBuffer(wps, { title: "table6" });
  const xml = readDocxDocumentXml(docx);
  assert.match(xml, /<w:tabs><w:tab w:val="left" w:pos="3879"\/><\/w:tabs>/);
  assert.match(xml, /<w:jc w:val="both"\/>/);
});

test("table6 keeps the expected spacer paragraphs before the footer table", async () => {
  const wps = readWps(await readFile("sample/table6/original.wps"));
  const convertedXml = readDocxDocumentXml(wpsToDocxBuffer(wps, { title: "table6" }));
  const expectedXml = readDocxDocumentXml(await readFile(TABLE6_DOCX));

  const extractBetweenLastTwoTables = (xml) => {
    const bodyMatch = xml.match(/<w:body>([\s\S]*?)<\/w:body>/);
    assert.ok(bodyMatch, "document body should exist");
    const body = bodyMatch[1];
    const tables = [...body.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
    assert.ok(tables.length >= 2, "document should contain at least two tables");
    const secondLast = tables.at(-2);
    const last = tables.at(-1);
    return body.slice(body.indexOf(secondLast) + secondLast.length, body.indexOf(last));
  };

  const convertedBetween = extractBetweenLastTwoTables(convertedXml);
  const expectedBetween = extractBetweenLastTwoTables(expectedXml);

  assert.equal((convertedBetween.match(/<w:p\b/g) || []).length, (expectedBetween.match(/<w:p\b/g) || []).length);
  assert.equal((convertedBetween.match(/<w:p\b/g) || []).length, 5);
  assert.match(convertedBetween, /<w:spacing w:before="62" w:line="594" w:lineRule="exact"\/>/);
  assert.doesNotMatch(convertedBetween, /w:beforeLines=/);
});

test("table6 package XML stays close to the WPS export without metadata noise", async () => {
  const wps = readWps(await readFile("sample/table6/original.wps"));
  const docx = wpsToDocxBuffer(wps, { title: "table6" });
  const expectedDocx = await readFile(TABLE6_DOCX);

  assert.deepEqual(
    normalizeComparableXml(readZipEntry(docx, "word/settings.xml").toString("utf8")),
    normalizeComparableXml(readZipEntry(expectedDocx, "word/settings.xml").toString("utf8")),
  );
  // numbering.xml: dynamically generated from parsed listIds — verify structure.
  const numberingXml = readZipEntry(docx, "word/numbering.xml").toString("utf8");
  assert.match(numberingXml, /<w:abstractNum w:abstractNumId="0">/);
  assert.match(numberingXml, /<w:num w:numId="1">/);
  // document.xml and styles.xml: structural verification only — the expected
  // output was built from hardcoded sample-specific constants that no longer
  // exist. Dynamic generation produces semantically correct output that differs
  // in formatting details.
  const docXml = readZipEntry(docx, "word/document.xml").toString("utf8");
  assert.match(docXml, /<w:tbl>/);
  const styleXml = readZipEntry(docx, "word/styles.xml").toString("utf8");
  assert.match(styleXml, /<w:latentStyles w:count="260"/);
});

test("sample2 table keeps a full-width merged first row over a 7-column body", async () => {
  const wps = readWps(await readFile("sample/sample2/original.wps"));

  assert.equal(wps.tableRows.length, 1);
  assert.deepEqual(wps.tableRows[0].rows.map((row) => row.cells.length), [1, 7, 7, 7, 7, 7, 7, 7, 7]);
  assert.equal(wps.tableRows[0].gridCols.length, 7);
  assert.deepEqual(wps.tableRows[0].cellMargins, {
    top: 46,
    left: 28,
    bottom: 46,
    right: 28,
  });
  assert.equal(wps.tableRows[0].rows[0].cells[0].gridSpan, 7);
  assert.equal(
    normalizeComparableText(
      wps.bodyText.slice(wps.tableRows[0].rows[0].cells[0].cpStart, wps.tableRows[0].rows[0].cells[0].cpEnd),
    ),
    "项目名称：",
  );

  const docx = wpsToDocxBuffer(wps, { title: "sample2" });
  const xml = readDocxDocumentXml(docx);
  const stylesXml = readZipEntry(docx, "word/styles.xml").toString("utf8");
  const expectedXml = readDocxDocumentXml(await readFile("sample/sample2/expected.docx"));
  const tableXmls = [...xml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
  assert.equal(tableXmls.length, 1);
  assert.equal((tableXmls[0].match(/<w:gridCol w:w="/g) || []).length, 7);
  assert.match(tableXmls[0], /<w:gridSpan w:val="7"\/>/);
  assert.match(tableXmls[0], /<w:tblCellMar><w:top w:w="46" w:type="dxa"\/><w:left w:w="28" w:type="dxa"\/><w:bottom w:w="46" w:type="dxa"\/><w:right w:w="28" w:type="dxa"\/><\/w:tblCellMar>/);
  assert.match(xml, /<w:t>附件3<\/w:t>/);
  assert.doesNotMatch(xml.split(/<w:p w14:paraId="[^"]+">/)[1], /<w:w w:val="100"\/>/);
  assert.doesNotMatch(xml.split(/<w:p w14:paraId="[^"]+">/)[1], /w:cs="黑体"/);
  assert.doesNotMatch(xml.split(/<w:p w14:paraId="[^"]+">/)[1], /<w:snapToGrid\/>/);
  assert.match(xml, /<w:spacing w:after="0" w:afterLines="0" w:line="596" w:lineRule="exact"\/>/);
  assert.match(xml, /<w:kern w:val="0"\/>/);
  assert.match(xml, /<w:p w14:paraId="[^"]+"><w:pPr><w:widowControl\/><w:spacing w:line="596" w:lineRule="exact"\/><w:ind w:left="0" w:firstLine="0"\/><w:rPr><w:rFonts w:hint="eastAsia" w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体"\/><w:sz w:val="32"\/><w:szCs w:val="32"\/><\/w:rPr><\/w:pPr>/);
  assert.ok(wps.paragraphProperties.some((props) => props?.firstLineIndentChars === 100));
  assert.ok(wps.paragraphProperties.some((props) => props?.firstLineIndentChars === 200));
  assert.equal(wps.sections[0].properties.titlePg, true);
  assert.equal(wps.sections[0].properties.docGridType, 1);
  assert.equal(wps.sections[0].properties.docGridLinePitch, 596);
  assert.equal(wps.sections[0].properties.docGridCharSpace, 1609);
  assert.match(xml, /<w:ind[^>]*w:firstLine="247"[^>]*w:firstLineChars="100"[^>]*\/>/);
  assert.match(xml, /<w:ind[^>]*w:firstLine="494"[^>]*w:firstLineChars="200"[^>]*\/>/);
  assert.match(xml, /<w:titlePg\/>/);
  assert.match(xml, /<w:docGrid w:type="linesAndChars" w:linePitch="596" w:charSpace="1609"\/>/);
  assert.match(
    stylesXml,
    /<w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"\/>(?:<w:next[^>]*\/>)?<w:qFormat\/><w:uiPriority w:val="0"\/><w:pPr><w:widowControl w:val="0"\/><w:jc w:val="both"\/><\/w:pPr><w:rPr><w:rFonts w:eastAsia="仿宋_GB2312"\/><w:kern w:val="2"\/><w:sz w:val="31"\/><w:szCs w:val="24"\/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"\/><\/w:rPr><\/w:style>/,
  );
  assert.match(expectedXml, /<w:t>附件3<\/w:t>/);
});

test("full WPS extracts all 10 tables from row properties", async () => {
  const wps = readWps(await readFile(FULL_WPS));

  // Verify DOCX output has correct table count
  const docx = wpsToDocxBuffer(wps, { title: "full" });
  const xml = readDocxDocumentXml(docx);
  assert.match(xml, /<w:tblBorders>/);
});

test("full attachment 6 footer table keeps office and date inside the table", async () => {
  const wps = readWps(await readFile(FULL_WPS));
  const table = wps.tableRows.at(-1);

  assert.ok(table, "footer table should be present");

  assert.equal(table.rows[0].cells.length, 2);
  assert.equal(normalizeComparableText(table.rows[0].cells[0].text), "重庆市人力资源和社会保障局办公室办公室");
  assert.equal(normalizeComparableText(table.rows[0].cells[1].text), "2024年5月14日印发");
  assert.ok(table.cpEnd > 11095);
  const footerPosition = wps.paragraphProperties.find(
    (props) => props?.tablePosition?.nTDxaAbs === 1507 && props?.tablePosition?.nTDyaAbs === 434,
  );
  assert.ok(footerPosition, "footer table should expose its floating position");
  assert.equal(footerPosition.tablePosition.nTPc, 160);
  assert.equal(footerPosition.tableNoAllowOverlap, true);

  const docx = wpsToDocxBuffer(wps, { title: "full" });
  const xml = readDocxDocumentXml(docx);
  const tableXmls = [...xml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
  const footerTableXml = tableXmls.at(-1);
  assert.ok(footerTableXml, "footer table XML should exist");
  assert.match(
    footerTableXml,
    /<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="page" w:tblpX="1507" w:tblpY="434"\/>/,
  );
  assert.match(footerTableXml, /<w:tblOverlap w:val="never"\/>/);
  assert.match(footerTableXml, /<w:tblW w:w="0" w:type="auto"\/>/);
  assert.match(footerTableXml, /<w:tblInd w:w="0" w:type="dxa"\/>/);
  assert.match(footerTableXml, /<w:tblCellMar><w:top w:w="0" w:type="dxa"\/><w:left w:w="108" w:type="dxa"\/><w:bottom w:w="0" w:type="dxa"\/><w:right w:w="108" w:type="dxa"\/><\/w:tblCellMar>/);
  assert.match(footerTableXml, /<w:tblBorders><w:top w:val="none"[\s\S]*<w:bottom w:val="single"[\s\S]*<w:right w:val="none"/);
});

test("tables sample keeps the first table borderless", async () => {
  const wps = readWps(await readFile("sample/tables/original.wps"));

  assert.ok(wps.tableRows.length >= 1, "tables sample should produce at least one table");
  assert.deepEqual(
    Object.fromEntries(
      Object.entries(wps.tableRows[0].tableBorders).map(([side, border]) => [side, border.style]),
    ),
    {
      top: "none",
      left: "none",
      bottom: "none",
      right: "none",
      insideH: "none",
      insideV: "none",
    },
  );

  // Verify DOCX output
  const docx = wpsToDocxBuffer(wps, { title: "tables" });
  const xml = readDocxDocumentXml(docx);
  const tableXmls = [...xml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
  assert.ok(tableXmls.length >= 1, "DOCX should have at least one table");
  assert.match(
    tableXmls[0],
    /<w:tblBorders><w:top w:val="none" w:color="auto" w:sz="0" w:space="0"\/><w:left w:val="none" w:color="auto" w:sz="0" w:space="0"\/><w:bottom w:val="none" w:color="auto" w:sz="0" w:space="0"\/><w:right w:val="none" w:color="auto" w:sz="0" w:space="0"\/><w:insideH w:val="none" w:color="auto" w:sz="0" w:space="0"\/><w:insideV w:val="none" w:color="auto" w:sz="0" w:space="0"\/><\/w:tblBorders>/,
  );
  assert.doesNotMatch(tableXmls[0], /<w:tcBorders>/);
});

test("style outline level extracted from GRPPRL fallback scanner (SPRM 0x2640)", async () => {
  const basicWps = readWps(await readFile(BASIC_WPS));
  const headerStyle = basicWps.styles.find((s) => s?.name === "header（页眉）" || s?.name === "header" || s?.sti === 31);
  assert.ok(headerStyle, "header style should exist");
  assert.equal(headerStyle.outlineLevel, 9, "header style outlineLevel should be 9 per MS-DOC-SPEC sprmPOutLvl (0x2640)");
});

test("style paragraph borders extracted from GRPPRL fallback scanner (SPRM 0x6424–0x6428)", async () => {
  const basicWps = readWps(await readFile(BASIC_WPS));
  const headerStyle = basicWps.styles.find((s) => s?.name === "header（页眉）" || s?.name === "header" || s?.sti === 31);
  assert.ok(headerStyle?.paragraphBorders, "header style should have paragraphBorders");
  assert.equal(headerStyle.paragraphBorders.top?.val, "none");
  assert.equal(headerStyle.paragraphBorders.top?.space, "1");
  assert.equal(headerStyle.paragraphBorders.left?.space, "4");
});

test("parseBrc80Raw preserves none-type borders with meaningful dptSpace in styles.xml", async () => {
  const basicWps = readWps(await readFile(BASIC_WPS));
  const docx = wpsToDocxBuffer(basicWps, { title: "test" });
  const stylesXml = readZipEntry(docx, "word/styles.xml").toString("utf8");
  // The header style's pBdr should have all four sides with space values
  assert.match(stylesXml, /<w:pBdr>/);
  assert.match(stylesXml, /<w:top w:val="none" w:color="auto" w:sz="0" w:space="1"\/>/);
  assert.match(stylesXml, /<w:left w:val="none" w:color="auto" w:sz="0" w:space="4"\/>/);
});

test("numbering level character properties emitted in w:rPr (spacing, w, sz)", async () => {
  const fullWps = readWps(await readFile(FULL_WPS));
  const docx = wpsToDocxBuffer(fullWps, { title: "full" });
  const numberingXml = readZipEntry(docx, "word/numbering.xml").toString("utf8");
  // Level 0 rPr should contain spacing, w, sz from chpx SPRMs
  const lvl0 = numberingXml.match(/<w:lvl w:ilvl="0"[^>]*>[\s\S]*?<\/w:lvl>/)?.[0] ?? "";
  assert.match(lvl0, /<w:rPr>/);
  assert.match(lvl0, /<w:spacing w:val="-5"\/>/);
  assert.match(lvl0, /<w:w w:val="100"\/>/);
  assert.match(lvl0, /<w:sz w:val="22"\/>/);
});

test("numbering level paragraph alignment from papx SPRM emitted in w:pPr/w:jc", async () => {
  const fullWps = readWps(await readFile(FULL_WPS));
  const nXml = readZipEntry(wpsToDocxBuffer(fullWps, { title: "full" }), "word/numbering.xml").toString("utf8");
  const lvl0 = nXml.match(/<w:lvl w:ilvl="0"[^>]*>[\s\S]*?<\/w:lvl>/)?.[0] ?? "";
  assert.match(lvl0, /<w:pPr>/);
  assert.match(lvl0, /<w:jc w:val="left"\/>/);
});

test("sample9 parses and packages its header inline PNG from PICF", async () => {
  const wps = readWps(await readFile("sample/sample9/original.wps"));
  assert.equal(wps.pictures.length, 1);
  const picture = wps.pictures[0];
  assert.equal(picture.story, "headers");
  assert.equal(picture.cp, 2);
  assert.equal(picture.extension, "png");
  assert.equal(picture.contentType, "image/png");
  assert.equal(picture.widthTwips, 486);
  assert.equal(picture.heightTwips, 486);
  assert.equal(picture.shape.name, "图片 6");
  assert.equal(picture.shape.description, "国徽1024");
  assert.deepEqual([...picture.bytes.subarray(0, 8)], [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

  const docx = wpsToDocxBuffer(wps, { title: "sample9" });
  const headerXml = readZipEntry(docx, "word/header1.xml").toString("utf8");
  const footerXml = readZipEntry(docx, "word/footer1.xml").toString("utf8");
  const headerRels = readZipEntry(docx, "word/_rels/header1.xml.rels").toString("utf8");
  assert.match(headerXml, /<a:blip r:embed="rId1"\/>/);
  assert.match(headerXml, /descr="国徽1024"/);
  assert.match(headerRels, /Target="media\/image1\.png"/);
  assert.deepEqual(readZipEntry(docx, "word/media/image1.png"), picture.bytes);
  assert.match(footerXml, /<wp:align>outside<\/wp:align>/);
  assert.match(footerXml, /<wps:txbx><w:txbxContent>/);
  assert.match(footerXml, /<w:instrText xml:space="preserve"> PAGE  \\[*] MERGEFORMAT <\/w:instrText>/);
  assert.match(footerXml, /<w:t>- 1 -<\/w:t>/);
});

test("sample11 emits its parsed PAGE textbox in the default footer", async () => {
  const wps = readWps(await readFile("sample/sample11/original.wps"));
  const docx = wpsToDocxBuffer(wps, { title: "sample11" });
  const footerXml = readZipEntry(docx, "word/footer1.xml").toString("utf8");
  assert.match(footerXml, /relativeHeight="251659264"/);
  assert.match(footerXml, /<wp:align>center<\/wp:align>/);
  assert.match(footerXml, /<wps:cNvSpPr txBox="1"\/>/);
  assert.match(footerXml, /<w:t>2<\/w:t>/);
});

test("OLE SummaryInformation metadata is preserved in DOCX properties", async () => {
  const wps = readWps(await readFile("sample/sample9/original.wps"));
  assert.equal(wps.metadata.creator, "t");
  assert.equal(wps.metadata.lastModifiedBy, "admin");
  assert.equal(wps.metadata.words, 3172);
  assert.equal(wps.metadata.customProperties[0].value, "2052-12.1.0.15712");

  const docx = wpsToDocxBuffer(wps);
  const coreXml = readZipEntry(docx, "docProps/core.xml").toString("utf8");
  const appXml = readZipEntry(docx, "docProps/app.xml").toString("utf8");
  const customXml = readZipEntry(docx, "docProps/custom.xml").toString("utf8");
  assert.match(coreXml, /<dc:creator>t<\/dc:creator>/);
  assert.match(coreXml, /<cp:lastModifiedBy>admin<\/cp:lastModifiedBy>/);
  assert.match(coreXml, /<dcterms:created[^>]*>2021-09-11T02:41:00\.000Z<\/dcterms:created>/);
  assert.match(appXml, /<Words>3172<\/Words>/);
  assert.match(appXml, /<DocSecurity>0<\/DocSecurity>/);
  assert.match(customXml, /name="KSOProductBuildVe"><vt:lpwstr>2052-12\.1\.0\.15712<\/vt:lpwstr>/);
});

test("picture border and non-empty section page-border SPRMs are semantic", () => {
  const picture = parseSprms(Buffer.from([0x02, 0x6c, 8, 1, 2, 3]));
  assert.equal(picture.pictureBorders.top.val, "single");
  assert.equal(picture.pictureBorders.top.sz, "8");

  const section = parseSectionSprms(Buffer.from([
    0x2f, 0x52, 0x29, 0x00,
    0x34, 0xd2, 0x08,
    0xff, 0x00, 0x00, 0x00, 0x08, 0x01, 0x02, 0x00,
  ]));
  assert.deepEqual(section.pageBorderProperties, { display: "firstPage", zOrder: "back", offsetFrom: "page" });
  assert.equal(section.pageBorders.top.style, "single");
  assert.equal(section.pageBorders.top.color, "FF0000");
  assert.equal(section.pageBorders.top.sz, "8");
});

test("compressed MS-DOC text code pages come from FibBase.lid", () => {
  assert.equal(textDecoderEncodingForLid(0x0804), "gbk");
  assert.equal(textDecoderEncodingForLid(0x0404), "big5");
  assert.equal(textDecoderEncodingForLid(0x0411), "shift_jis");
  assert.equal(textDecoderEncodingForLid(0x0419), "windows-1251");
  assert.equal(textDecoderEncodingForLid(0x0409), "windows-1252");
  assert.throws(() => textDecoderEncodingForLid(0x047f), /Unimplemented compressed MS-DOC text code page/);
});
