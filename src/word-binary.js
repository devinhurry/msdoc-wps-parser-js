import { parseSprms } from "./sprm.js";
import { BRC_TYPE_NAMES, brcColorFromIco, colorRefToHex } from "./sprm.js";
// sti→name mapping per MS-OI29500 §2.1.237 (Part 1 Section 17.7.4.9)
export const STI_NAMES = ["Normal","heading 1","heading 2","heading 3","heading 4","heading 5","heading 6","heading 7","heading 8","heading 9","index 1","index 2","index 3","index 4","index 5","index 6","index 7","index 8","index 9","toc 1","toc 2","toc 3","toc 4","toc 5","toc 6","toc 7","toc 8","toc 9","Normal Indent","footnote text","annotation text","header","footer","index heading","caption","table of figures","envelope address","envelope return","footnote reference","annotation reference","line number","page number","endnote reference","endnote text","table of authorities","macro","toa heading","List","List Bullet","List Number","List 2","List 3","List 4","List 5","List Bullet 2","List Bullet 3","List Bullet 4","List Bullet 5","List Number 2","List Number 3","List Number 4","List Number 5","Title","Closing","Signature","Default Paragraph Font","Body Text","Body Text Indent","List Continue","List Continue 2","List Continue 3","List Continue 4","List Continue 5","Message Header","Subtitle","Salutation","Date","Body Text First Indent","Body Text First Indent 2","Note Heading","Body Text 2","Body Text 3","Body Text Indent 2","Body Text Indent 3","Block Text","Hyperlink","FollowedHyperlink","Strong","Emphasis","Document Map","Plain Text","E-mail Signature","HTML Top of Form","HTML Bottom of Form","Normal (Web)","HTML Acronym","HTML Address","HTML Cite","HTML Code","HTML Definition","HTML Keyboard","HTML Preformatted","HTML Sample","HTML Typewriter","HTML Variable","Normal Table","annotation subject","No List","Outline List 1","Outline List 2","Outline List 3","Table Simple 1","Table Simple 2","Table Simple 3","Table Classic 1","Table Classic 2","Table Classic 3","Table Classic 4","Table Colorful 1","Table Colorful 2","Table Colorful 3","Table Columns 1","Table Columns 2","Table Columns 3","Table Columns 4","Table Columns 5","Table Grid 1","Table Grid 2","Table Grid 3","Table Grid 4","Table Grid 5","Table Grid 6","Table Grid 7","Table Grid 8","Table List 1","Table List 2","Table List 3","Table List 4","Table List 5","Table List 6","Table List 7","Table List 8","Table 3D effects 1","Table 3D effects 2","Table 3D effects 3","Table Contemporary","Table Elegant","Table Professional","Table Subtle 1","Table Subtle 2","Table Web 1","Table Web 2","Table Web 3","Balloon Text","Table Grid","Table Theme","Placeholder Text","No Spacing","Light Shading","Light List","Light Grid","Medium Shading 1","Medium Shading 2","Medium List 1","Medium List 2","Medium Grid 1","Medium Grid 2","Medium Grid 3","Dark List","Colorful Shading","Colorful List","Colorful Grid","Light Shading Accent 1","Light List Accent 1","Light Grid Accent 1","Medium Shading 1 Accent 1","Medium Shading 2 Accent 1","Medium List 1 Accent 1","Revision","List Paragraph","Quote","Intense Quote","Medium List 2 Accent 1","Medium Grid 1 Accent 1","Medium Grid 2 Accent 1","Medium Grid 3 Accent 1","Dark List Accent 1","Colorful Shading Accent 1","Colorful List Accent 1","Colorful Grid Accent 1","Light Shading Accent 2","Light List Accent 2","Light Grid Accent 2","Medium Shading 1 Accent 2","Medium Shading 2 Accent 2","Medium List 1 Accent 2","Medium List 2 Accent 2","Medium Grid 1 Accent 2","Medium Grid 2 Accent 2","Medium Grid 3 Accent 2","Dark List Accent 2","Colorful Shading Accent 2","Colorful List Accent 2","Colorful Grid Accent 2","Light Shading Accent 3","Light List Accent 3","Light Grid Accent 3","Medium Shading 1 Accent 3","Medium Shading 2 Accent 3","Medium List 1 Accent 3","Medium List 2 Accent 3","Medium Grid 1 Accent 3","Medium Grid 2 Accent 3","Medium Grid 3 Accent 3","Dark List Accent 3","Colorful Shading Accent 3","Colorful List Accent 3","Colorful Grid Accent 3","Light Shading Accent 4","Light List Accent 4","Light Grid Accent 4","Medium Shading 1 Accent 4","Medium Shading 2 Accent 4","Medium List 1 Accent 4","Medium List 2 Accent 4","Medium Grid 1 Accent 4","Medium Grid 2 Accent 4","Medium Grid 3 Accent 4","Dark List Accent 4","Colorful Shading Accent 4","Colorful List Accent 4","Colorful Grid Accent 4","Light Shading Accent 5","Light List Accent 5","Light Grid Accent 5","Medium Shading 1 Accent 5","Medium Shading 2 Accent 5","Medium List 1 Accent 5","Medium List 2 Accent 5","Medium Grid 1 Accent 5","Medium Grid 2 Accent 5","Medium Grid 3 Accent 5","Dark List Accent 5","Colorful Shading Accent 5","Colorful List Accent 5","Colorful Grid Accent 5","Light Shading Accent 6","Light List Accent 6","Light Grid Accent 6","Medium Shading 1 Accent 6","Medium Shading 2 Accent 6","Medium List 1 Accent 6","Medium List 2 Accent 6","Medium Grid 1 Accent 6","Medium Grid 2 Accent 6","Medium Grid 3 Accent 6","Dark List Accent 6","Colorful Shading Accent 6","Colorful List Accent 6","Colorful Grid Accent 6"];

const WORD_BINARY_MAGIC = 0xa5ec;
const FIB_FLAGS_OFFSET = 0x0a;
const FIB_F_COMPLEX = 0x0004;
const FIB_F_WHICH_TABLE_STREAM = 0x0200;

const FIB_CCP_TEXT_OFFSET = 0x4c;
const FIB_CCP_FTN_OFFSET = 0x50;
const FIB_CCP_HDD_OFFSET = 0x54;
const FIB_CCP_ATN_OFFSET = 0x5c;
const FIB_CCP_EDN_OFFSET = 0x60;
const FIB_CCP_TXBX_OFFSET = 0x64;
const FIB_CCP_HDR_TXBX_OFFSET = 0x68;

const FIB_FC_LCB_START = 0x9a;
const FIB_FC_LCB_COUNT_OFFSET = 0x98;
const FIB_FC_CLX_INDEX = 33;
const FIB_FC_PLCFSED_INDEX = 6;
const FIB_FC_PLCFHDD_INDEX = 8;
const FIB_FC_PAPX_INDEX = 13;
const FIB_FC_STSH_INDEX = 1;
const FIB_FC_FONT_TABLE_INDEX = 15;

const PHE_SIZE = 12;
const SPRM_WPS_DYA_LINE = 0x6412;
const SPRM_WPS_DYA_LINE_OPERAND_SIZE = 2;
const FIB_FC_CHPX_INDEX = 12;
const FIB_FC_PLCFLST_INDEX = 73;
const FIB_FC_PLCFLFO_INDEX = 74;
const FIB_FC_STTBFBKMK_INDEX = 21;
const FIB_FC_PLCFBKF_INDEX = 22;
const FIB_FC_PLCFBKL_INDEX = 23;
const FIB_FC_WSS_INDEX = 30;
const FIB_FC_DOP_INDEX = 31; // 0x1F per FibRgFcLcb97
const FIB_FC_STTBFRMARK_INDEX = 51; // MS-DOC-SPEC/15 FibRgFcLcb97 fcSttbfRMark at offset 0x232
const SELSF_SIZE = 36;

const STSH_NIL_BASE = 0xfff0;
const STSH_STD_HEADER_SIZE_WITH_POST2000 = 18;
const SPRM_OPERAND_SIZE_BY_SPRA = [1, 1, 2, 4, 2, 2, -1, 3];
const TABLE_BORDER_SIDES = ["top", "left", "bottom", "right", "insideH", "insideV"];

export function extractWordBinaryDocument({ wordDocument, table0, table1 = null, data = null }) {
  assertWordDocument(wordDocument);

  const fib = readFib(wordDocument);
  if (fib.fEncrypted) {
    throw new Error("Excluded Word binary document variant: encrypted/obfuscated files are outside this parser scope");
  }
  if ((fib.flags & FIB_F_COMPLEX) === 0) {
    throw new Error("Unimplemented Word binary document variant: non-complex files without a CLX piece table");
  }

  const tableStream = fib.whichTableStream === "1Table" ? table1 : table0;
  if (!tableStream) {
    throw new Error(`Missing required Word table stream: ${fib.whichTableStream}`);
  }

  const pieces = readPieceTable(tableStream, fib.fcClx, fib.lcbClx);
  const subdocuments = splitSubdocuments(wordDocument, pieces, fib.characterCounts);
  const rawText = readPieces(wordDocument, pieces);
  const bodyText = subdocuments.body.rawText;
  const dop = parseDop(tableStream, fib);
  const { styles, latentLsd, stiMaxWhenSaved } = extractStyleSheet(tableStream, fib);
  const fontTable = extractFontTable(tableStream, fib);
  const sections = extractSections(wordDocument, tableStream, fib);
  const listData = extractListData(tableStream, fib);
  const defaultTabStop = dop.dxaTab;
  const plcfHdd = parsePlcfHdd(tableStream, fib);
  const bookmarks = parseStandardBookmarks(tableStream, fib);
  const revisionAuthors = parseRevisionAuthors(tableStream, fib);
  const lastSelection = parseLastSelection(tableStream, fib, rawText.length);
  const paragraphProperties = extractParagraphProperties(wordDocument, tableStream, fib, bodyText, styles, pieces);
  const characterRuns = extractCharacterRuns(wordDocument, tableStream, fib, bodyText, pieces, styles);
  const characterProperties = expandCharacterRuns(characterRuns, bodyText.length);
  const tableRows = extractTableRows(wordDocument, tableStream, fib, pieces, bodyText, paragraphProperties, sections, data, styles);

  return {
    fib,
    pieces,
    text: normalizeWordText(bodyText),
    rawText,
    bodyText,
    paragraphs: paragraphsFromWordText(bodyText),
    paragraphProperties,
    characterProperties,
    characterRuns,
    styles,
    latentLsd,
    stiMaxWhenSaved,
    fontTable,
    sections,
    listData,
    defaultTabStop,
    subdocuments,
    tableRows,
    plcfHdd,
    bookmarks,
    revisionAuthors,
    lastSelection,
    dop,
  };
}

export function normalizeWordText(text) {
  return text
    .replace(/\r/g, "\n")
    .replace(/\x0c/g, "\n")
    .replace(/\x07/g, "")
    .replace(/[\x00-\x06\x08\x0b\x0e-\x1f]/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

export function normalizeComparableText(text) {
  return normalizeWordText(text).replace(/\s+/g, "");
}

function readFib(wordDocument) {
  const flags = wordDocument.readUInt16LE(FIB_FLAGS_OFFSET);
  const flags2 = wordDocument[0x13];
  const fcLcbCount = wordDocument.readUInt16LE(FIB_FC_LCB_COUNT_OFFSET);
  const tableStreamOffset = FIB_FC_LCB_START + FIB_FC_CLX_INDEX * 8;
  const plcfSedOffset = FIB_FC_LCB_START + FIB_FC_PLCFSED_INDEX * 8;
  const plcfHddOffset = FIB_FC_LCB_START + FIB_FC_PLCFHDD_INDEX * 8;
  const papxOffset = FIB_FC_LCB_START + FIB_FC_PAPX_INDEX * 8;
  const chpxOffset = FIB_FC_LCB_START + FIB_FC_CHPX_INDEX * 8;
  const fontTableOffset = FIB_FC_LCB_START + FIB_FC_FONT_TABLE_INDEX * 8;
  const stshOffset = FIB_FC_LCB_START + FIB_FC_STSH_INDEX * 8;
  const plcfLstOffset = FIB_FC_LCB_START + FIB_FC_PLCFLST_INDEX * 8;
  const plcfLfoOffset = FIB_FC_LCB_START + FIB_FC_PLCFLFO_INDEX * 8;
  const sttbfBkmkOffset = FIB_FC_LCB_START + FIB_FC_STTBFBKMK_INDEX * 8;
  const plcfBkfOffset = FIB_FC_LCB_START + FIB_FC_PLCFBKF_INDEX * 8;
  const plcfBklOffset = FIB_FC_LCB_START + FIB_FC_PLCFBKL_INDEX * 8;
  const wssOffset = FIB_FC_LCB_START + FIB_FC_WSS_INDEX * 8;
  const dopOffset = FIB_FC_LCB_START + FIB_FC_DOP_INDEX * 8;
  const sttbfRMarkOffset = FIB_FC_LCB_START + FIB_FC_STTBFRMARK_INDEX * 8;
  if (tableStreamOffset + 8 > FIB_FC_LCB_START + fcLcbCount * 4) {
    throw new Error("Unimplemented Word binary document variant: FIB does not contain fcClx/lcbClx");
  }

  return {
    nFib: wordDocument.readUInt16LE(0x02),
    lid: wordDocument.readUInt16LE(0x06),
    pnNext: wordDocument.readUInt16LE(0x08),
    flags,
    fDot: (flags & 0x0001) !== 0,
    fGlsy: (flags & 0x0002) !== 0,
    fComplex: (flags & FIB_F_COMPLEX) !== 0,
    fHasPic: (flags & 0x0008) !== 0,
    cQuickSaves: (flags >> 4) & 0x0f,
    fEncrypted: (flags & 0x0100) !== 0,
    fReadOnlyRecommended: (flags & 0x0400) !== 0,
    fWriteReservation: (flags & 0x0800) !== 0,
    fExtChar: (flags & 0x1000) !== 0,
    fLoadOverride: (flags & 0x2000) !== 0,
    fFarEast: (flags & 0x4000) !== 0,
    fObfuscated: (flags & 0x8000) !== 0,
    nFibBack: wordDocument.readUInt16LE(0x0c),
    lKey: wordDocument.readUInt32LE(0x0e),
    envr: wordDocument[0x12],
    fMac: (flags2 & 0x01) !== 0,
    fEmptySpecial: (flags2 & 0x02) !== 0,
    fLoadOverridePage: (flags2 & 0x04) !== 0,
    whichTableStream: (flags & FIB_F_WHICH_TABLE_STREAM) === 0 ? "0Table" : "1Table",
    fcMin: wordDocument.readUInt32LE(0x18),
    fcMac: wordDocument.readUInt32LE(0x1c),
    fcClx: wordDocument.readUInt32LE(tableStreamOffset),
    lcbClx: wordDocument.readUInt32LE(tableStreamOffset + 4),
    fcPlcfSed: wordDocument.readUInt32LE(plcfSedOffset),
    lcbPlcfSed: wordDocument.readUInt32LE(plcfSedOffset + 4),
    fcPlcfHdd: wordDocument.readUInt32LE(plcfHddOffset),
    lcbPlcfHdd: wordDocument.readUInt32LE(plcfHddOffset + 4),
    fcPapx: wordDocument.readUInt32LE(papxOffset),
    lcbPapx: wordDocument.readUInt32LE(papxOffset + 4),
    fcChpx: wordDocument.readUInt32LE(chpxOffset),
    lcbChpx: wordDocument.readUInt32LE(chpxOffset + 4),
    fcFontTable: wordDocument.readUInt32LE(fontTableOffset),
    lcbFontTable: wordDocument.readUInt32LE(fontTableOffset + 4),
    fcStsh: wordDocument.readUInt32LE(stshOffset),
    lcbStsh: wordDocument.readUInt32LE(stshOffset + 4),
    fcPlcfLst: wordDocument.readUInt32LE(plcfLstOffset),
    lcbPlcfLst: wordDocument.readUInt32LE(plcfLstOffset + 4),
    fcPlfLfo: wordDocument.readUInt32LE(plcfLfoOffset),
    lcbPlfLfo: wordDocument.readUInt32LE(plcfLfoOffset + 4),
    fcSttbfBkmk: wordDocument.readUInt32LE(sttbfBkmkOffset),
    lcbSttbfBkmk: wordDocument.readUInt32LE(sttbfBkmkOffset + 4),
    fcPlcfBkf: wordDocument.readUInt32LE(plcfBkfOffset),
    lcbPlcfBkf: wordDocument.readUInt32LE(plcfBkfOffset + 4),
    fcPlcfBkl: wordDocument.readUInt32LE(plcfBklOffset),
    lcbPlcfBkl: wordDocument.readUInt32LE(plcfBklOffset + 4),
    fcWss: wssOffset + 8 <= FIB_FC_LCB_START + fcLcbCount * 4 ? wordDocument.readUInt32LE(wssOffset) : 0,
    lcbWss: wssOffset + 8 <= FIB_FC_LCB_START + fcLcbCount * 4 ? wordDocument.readUInt32LE(wssOffset + 4) : 0,
    fcDop: dopOffset + 8 <= FIB_FC_LCB_START + fcLcbCount * 4 ? wordDocument.readUInt32LE(dopOffset) : 0,
    lcbDop: dopOffset + 8 <= FIB_FC_LCB_START + fcLcbCount * 4 ? wordDocument.readUInt32LE(dopOffset + 4) : 0,
    fcSttbfRMark: sttbfRMarkOffset + 8 <= FIB_FC_LCB_START + fcLcbCount * 4 ? wordDocument.readUInt32LE(sttbfRMarkOffset) : 0,
    lcbSttbfRMark: sttbfRMarkOffset + 8 <= FIB_FC_LCB_START + fcLcbCount * 4 ? wordDocument.readUInt32LE(sttbfRMarkOffset + 4) : 0,
    characterCounts: {
      body: wordDocument.readUInt32LE(FIB_CCP_TEXT_OFFSET),
      footnotes: wordDocument.readUInt32LE(FIB_CCP_FTN_OFFSET),
      headers: wordDocument.readUInt32LE(FIB_CCP_HDD_OFFSET),
      annotations: wordDocument.readUInt32LE(FIB_CCP_ATN_OFFSET),
      endnotes: wordDocument.readUInt32LE(FIB_CCP_EDN_OFFSET),
      textboxes: wordDocument.readUInt32LE(FIB_CCP_TXBX_OFFSET),
      headerTextboxes: wordDocument.readUInt32LE(FIB_CCP_HDR_TXBX_OFFSET),
    },
  };
}

// MS-DOC-SPEC/17 §DopBase, §Dop97, §Dop2002: parse DOP fields needed for
// OOXML settings.xml emission.
function parseDop(tableStream, fib) {
  const { fcDop, lcbDop } = fib;
  if (!lcbDop) {
    throw new Error("Invalid Word binary document: missing mandatory DOP");
  }
  if (fcDop > tableStream.length || lcbDop > tableStream.length - fcDop) {
    throw new Error("Invalid Word binary document: DOP is outside the table stream");
  }
  if (lcbDop < 84) {
    throw new Error(`Out-of-spec Word binary document: DOP is shorter than DopBase (${lcbDop} bytes)`);
  }
  const dop = tableStream.subarray(fcDop, fcDop + lcbDop);

  // ── DopBase fields (84 bytes minimum) ──────────────────────────────
  // MS-DOC-SPEC/17 §DopBase, diagram row 1 (bytes 0-3)
  const dopBase0 = dop.length >= 4 ? dop.readUInt32LE(0) : 0;
  const fFacingPages     = ((dopBase0 >>  0) & 1) !== 0;  // bit  0 = byte0 bit0 = A
  // bit 1 = unused1 (B), bit 2 = fPMHMainDoc (C), bits 3-4 = unused2 (D),
  // bits 5-6 = fpc, bit 7 = unused3 (E)
  const fpc              = (dopBase0 >>  5) & 0x03;        // bits 5-6
  // byte1 = unused4 (bits 8-15)
  const rncFtn           = (dopBase0 >> 16) & 0x03;        // bits 16-17 = F
  const nFtn             = (dopBase0 >> 18) & 0x3FFF;       // bits 18-31, 14 bits

  // MS-DOC-SPEC/17 §DopBase, diagram row 2 (bytes 4-7)
  const dopBase4 = dop.length >= 8 ? dop.readUInt32LE(4) : 0;
  const fSplAllDone      = ((dopBase4 >>  6) & 1) !== 0;   // bit 38 = M
  const fSplAllClean     = ((dopBase4 >>  7) & 1) !== 0;   // bit 39 = N
  const fSplHideErrors   = ((dopBase4 >>  8) & 1) !== 0;   // bit 40 = O
  const fGramHideErrors  = ((dopBase4 >>  9) & 1) !== 0;   // bit 41 = P
  const fLabelDoc        = ((dopBase4 >> 10) & 1) !== 0;   // bit 42 = Q
  const fHyphCapitals    = ((dopBase4 >> 11) & 1) !== 0;   // bit 43 = R
  const fAutoHyphen      = ((dopBase4 >> 12) & 1) !== 0;   // bit 44 = S
  const fFormNoFields    = ((dopBase4 >> 13) & 1) !== 0;   // bit 45 = T
  const fLinkStyles      = ((dopBase4 >> 14) & 1) !== 0;   // bit 46 = U
  const fRevMarking      = ((dopBase4 >> 15) & 1) !== 0;   // bit 47 = V
  // bits 48-49: W(unused11), X(fExactCWords)
  const fPagHidden       = ((dopBase4 >> 18) & 1) !== 0;   // bit 50 = Y
  const fPagResults      = ((dopBase4 >> 19) & 1) !== 0;   // bit 51 = Z
  const fLockAtn         = ((dopBase4 >> 20) & 1) !== 0;   // bit 52 = a
  const fMirrorMargins   = ((dopBase4 >> 21) & 1) !== 0;   // bit 53 = b
  const fWord97Compat    = ((dopBase4 >> 22) & 1) !== 0;   // bit 54 = c
  // bit 55 = unused12 (d), bit 56 = unused13 (e)
  const fProtEnabled     = ((dopBase4 >> 25) & 1) !== 0;   // bit 57 = f
  const fDispFormFldSel  = ((dopBase4 >> 26) & 1) !== 0;   // bit 58 = g
  const fRMView          = ((dopBase4 >> 27) & 1) !== 0;   // bit 59 = h
  const fRMPrint         = ((dopBase4 >> 28) & 1) !== 0;   // bit 60 = i
  const fLockVbaProj     = ((dopBase4 >> 29) & 1) !== 0;   // bit 61 = j
  const fLockRev         = ((dopBase4 >> 30) & 1) !== 0;   // bit 62 = k
  const fEmbedFonts      = ((dopBase4 >> 31) & 1) !== 0;   // bit 63 = l

  // MS-DOC-SPEC/17 DopBase.dxaTab is at bytes 10-11
  const dxaTab = dop.readUInt16LE(10);

  // MS-DOC-SPEC/17 DopBase.cpgWebOpt at bytes 12-13
  const cpgWebOpt = dop.readUInt16LE(12);

  // MS-DOC-SPEC/17 DopBase.dxaHotZ at bytes 14-15 → OOXML hyphenationZone
  const dxaHotZ = dop.readUInt16LE(14);

  // MS-DOC-SPEC/17 DopBase.cConsecHypLim at bytes 16-17 → OOXML consecutiveHyphenLimit
  const cConsecHypLim = dop.readUInt16LE(16);

  // MS-DOC-SPEC/17 §Copts60 at bytes 8-9: fDntULTrlSpc(bit14) is opposite of ulTrailSpace
  const copts60 = dop.readUInt16LE(8);
  const compatibility = {
    ulTrailSpace:            ((copts60 >> 14) & 1) === 0,  // P = fDntULTrlSpc, opposite sense
    noTabHangInd:            ((copts60 >>  0) & 1) !== 0,  // A = fNoTabForInd
    noSpaceRaiseLower:       ((copts60 >>  1) & 1) !== 0,  // B = fNoSpaceRaiseLower
    suppressSpBfAfterPgBrk: ((copts60 >>  2) & 1) !== 0,  // C = fSuppressSpBfAfterPgBrk
    wrapTrailSpaces:         ((copts60 >>  3) & 1) !== 0,  // D = fWrapTrailSpaces
    printColBlack:           ((copts60 >>  4) & 1) !== 0,  // E = fMapPrintTextColor
    noColumnBalance:         ((copts60 >>  5) & 1) !== 0,  // F = fNoColumnBalance
    convMailMergeEsc:        ((copts60 >>  6) & 1) !== 0,  // G = fConvMailMergeEsc
    suppressTopSpacing:      ((copts60 >>  7) & 1) !== 0,  // H = fSuppressTopSpacing
    useSingleBorderForContiguousCells: ((copts60 >>  8) & 1) !== 0,  // I = fOrigWordTableRules
    showBreaksInFrames:      ((copts60 >> 10) & 1) !== 0,  // K = fShowBreaksInFrames
    swapBordersFacingPages:  ((copts60 >> 11) & 1) !== 0,  // L = fSwapBordersFacingPgs
    doNotLeaveBackslashAlone: ((copts60 >> 12) & 1) === 0, // M = fLeaveBackslashAlone, opposite sense
    doNotExpandShiftReturn:  ((copts60 >> 13) & 1) === 0,  // N = fExpShRtn, opposite sense
    balanceSingleByteDoubleByteWidth: ((copts60 >> 15) & 1) === 0, // P2 = fDntBlnSbDbWid, opposite
  };

  // DopBase bytes 52-55: m=rncEdn(2), nEdn(14), epc(2), n=unused14(4), o=unused15(4),
  // p=fPrintFormData, q=fSaveFormData, r=fShadeFormData, s=fShadeMergeFields,
  // t=reserved2, u=fIncludeSubdocsInStats
  const dopBase52 = dop.length >= 56 ? dop.readUInt32LE(52) : 0;
  const rncEdn            = (dopBase52 >>  0) & 0x03;
  const nEdn              = (dopBase52 >>  2) & 0x3FFF;
  const epc               = (dopBase52 >> 16) & 0x03;
  const fPrintFormData    = ((dopBase52 >> 26) & 1) !== 0;  // bit 26 = p
  const fSaveFormData     = ((dopBase52 >> 27) & 1) !== 0;  // bit 27 = q
  const fShadeFormData    = ((dopBase52 >> 28) & 1) !== 0;  // bit 28 = r
  const fShadeMergeFields = ((dopBase52 >> 29) & 1) !== 0;  // bit 29 = s
  const fIncludeSubdocsInStats = ((dopBase52 >> 31) & 1) !== 0; // bit 31 = u

  // DopBase byte 51 (in diagram row 5): v=wvkoSaved(3), pctWwdSaved(9), w=zkSaved(2),
  // x=unused16, y=iGutterPos — stored in the high word of bytes 48-51 (cParas area)
  // Actually cParas is bytes 48-51. Let's read wvkoSaved/zkSaved from byte 51
  // according to the DopBase diagram row 5.
  // Row 5 layout (bytes 52-55) has: m, nEdn, epc, n, o, p, q, r, s, t, u
  // Row 6 starts with cLines (bytes 56-59)
  // The v/pctWwdSaved/w/x/y are at the end of the cParasWithSubdocs area.
  // Let's re-read the diagram:
  // Row 5: bytes 52-55 = | m |  | nEdn ... | epc |  | n ... | o ... | p | q | r | s | t | u |
  // The v, pctWwdSaved, w, x, y are in the row that goes:
  // | ... |  |  |  |  |  |  |  | v |  |  | pctWwdSaved |  |  |  |  | w |  | x | y |
  // This is at the end of byte 51 (since cPgWithSubdocs is at 80-81... actually it varies)
  // Let me just skip these for now as they're complex and low value for WPS output.

  // ── DopTypography at DOP offset 90 ────────────────────────────────
  // MS-DOC-SPEC/17 §DopTypography: maps fKerningPunct → noPunctuationKerning (opposite)
  // and iJustification → characterSpacingControl.
  // Also includes iLevelOfKinsoku, f2on1, iCustomKsu, fJapaneseUseLevel2.
  let typography = null;
  if (dop.length >= 400) {
    const flags = dop[90];
    typography = {
      fKerningPunct:     (flags & 0x01) !== 0,
      iJustification:    (flags >> 1) & 0x03,
      iLevelOfKinsoku:   (flags >> 4) & 0x03,  // bits 4-5 = C
      f2on1:             (flags >> 6) & 0x01,  // bit 6 = D
      iCustomKsu:        (flags >> 8) & 0x07,  // bits 8-10 = F (from byte 91? Let me check)
      // Actually the diagram shows DopTypography as a multi-row structure
      // starting at byte 90. Let me read the actual byte layout properly.
      // DopTypography diagram row 1 (byte 90): A B - C - D E F - - G reserved(5 bits)
      // Byte 90: bit0=A(fKerningPunct), bits1-2=B(iJustification), bits4-5=C(iLevelOfKinsoku)
      // bit6=D(f2on1), bits8-10=F(iCustomKsu)... wait, byte 90 only has bits 0-7
      // Byte 91: bits 0-2 = F(iCustomKsu continued), bit 3 = G(fJapaneseUseLevel2),
      // bits 4-7 = reserved
    };
    // Read full DopTypography flags from bytes 90-91
    const typoFlags = dop.readUInt16LE(90);
    typography = {
      fKerningPunct:        (typoFlags & 0x0001) !== 0,
      iJustification:       (typoFlags >> 1) & 0x03,
      iLevelOfKinsoku:      (typoFlags >> 4) & 0x03,
      f2on1:                (typoFlags >> 6) & 0x01,
      iCustomKsu:           (typoFlags >> 8) & 0x07,
      fJapaneseUseLevel2:   (typoFlags >> 11) & 0x01,
    };
  }

  // ── Dogrid at DOP offset 400 (Dop97) ──────────────────────────────
  // MS-DOC-SPEC/17 §Dogrid: drawing grid settings.
  let dogrid = null;
  if (dop.length >= 410) {
    const fFollowMargins = (dop[409] >> 7) & 1;
    dogrid = {
      xaGrid: dop.readUInt16LE(400),
      yaGrid: dop.readUInt16LE(402),
      dxaGrid: dop.readUInt16LE(404),
      dyaGrid: dop.readUInt16LE(406),
      dyGridDisplay: dop[408] & 0x7F,
      dxGridDisplay: dop[409] & 0x7F,
      fFollowMargins: !!fFollowMargins,
    };
  }

  // ── Dop97 shared flags at offset 410 ──────────────────────────────
  // MS-DOC-SPEC/17 §Dop97 row 11 (bytes 410-411):
  //   bit 4-7 : lvlDop           (4-bit outline-level view saved state)
  //   bit 5   : fGramAllDone     (B)
  //   bit 6   : fGramAllClean    (C)
  //   bit 7   : fSubsetFonts     (D)
  //   bit 9   : fHtmlDoc         (F)
  //   bit 11  : fSnapBorder       (G)
  //   bit 12  : fIncludeHeader   (H)
  //   bit 13  : fIncludeFooter   (I)
  //   bit 14  : fForcePageSizePag (J)
  //   bit 15  : fMinFontSizePag   (K)
  // Per MS-DOC-SPEC/17 §17.8.3.15 saveSubsetFonts emission depends on
  // DopBase.fEmbedFonts (true) AND Dop97.fSubsetFonts (true):
  //   <w:embedTrueTypeFonts/>   iff DopBase.fEmbedFonts
  //   <w:saveSubsetFonts/>      iff DopBase.fEmbedFonts AND Dop97.fSubsetFonts
  let pageBorderIncludes = null;
  let fSubsetFonts = false;
  if (dop.length >= 412) {
    const dop97Flags = dop.readUInt16LE(410);
    pageBorderIncludes = {
      header: ((dop97Flags >> 12) & 1) !== 0,
      footer: ((dop97Flags >> 13) & 1) !== 0,
    };
    fSubsetFonts = ((dop97Flags >> 7) & 1) !== 0;
  }

  // ── Dop2002 XML validation at offset 542 ──────────────────────────
  // fValidateXML (bit b) and fShowXMLErrors (bit d) have opposite sense from
  // OOXML doNotValidateAgainstSchema and doNotDemarcateInvalidXml.
  let xmlValidation = null;
  if (dop.length >= 546) {
    const dop2002Flags = dop.readUInt32LE(542);
    xmlValidation = {
      fValidateXML: ((dop2002Flags >> 12) & 1) !== 0,
      fShowXMLErrors: ((dop2002Flags >> 14) & 1) !== 0,
    };
  }

  // ── grfFmtFilter at DOP offset 554 (Dop2002) ─────────────────────
  // MS-DOC-SPEC/17; default per spec is 0x5024
  const grfFmtFilter = dop.length >= 556 ? dop.readUInt16LE(554) : null;

  return {
    fFacingPages,
    fRevMarking,
    fAutoHyphen,
    fHyphCapitals,
    fEmbedFonts,
    fMirrorMargins,
    fRMView,
    fRMPrint,
    fLockRev,
    fProtEnabled,
    fLockAtn,
    fLinkStyles,
    fSplAllDone,
    fSplAllClean,
    fSplHideErrors,
    fGramHideErrors,
    fLabelDoc,
    fFormNoFields,
    fPagHidden,
    fPagResults,
    fDispFormFldSel,
    fLockVbaProj,
    fWord97Compat,
    fpc,
    rncFtn,
    nFtn,
    dxaTab,
    cpgWebOpt,
    dxaHotZ,
    cConsecHypLim,
    rncEdn,
    nEdn,
    epc,
    fPrintFormData,
    fSaveFormData,
    fShadeFormData,
    fShadeMergeFields,
    fIncludeSubdocsInStats,
    compatibility,
    typography,
    dogrid,
    pageBorderIncludes,
    fSubsetFonts,
    xmlValidation,
    grfFmtFilter,
  };
}

function readPieceTable(tableStream, fcClx, lcbClx) {
  if (fcClx + lcbClx > tableStream.length) {
    throw new Error("Invalid Word binary document: CLX is outside the table stream");
  }

  const clx = tableStream.subarray(fcClx, fcClx + lcbClx);
  let offset = 0;
  while (offset < clx.length && clx[offset] === 0x01) {
    if (offset + 3 > clx.length) {
      throw new Error("Invalid Word binary document: truncated CLX Prc block");
    }
    offset += 3 + clx.readUInt16LE(offset + 1);
  }

  if (offset >= clx.length || clx[offset] !== 0x02) {
    throw new Error("Invalid Word binary document: CLX does not contain a Pcdt piece table");
  }
  if (offset + 5 > clx.length) {
    throw new Error("Invalid Word binary document: truncated Pcdt header");
  }

  const plcPcdLength = clx.readUInt32LE(offset + 1);
  const plcPcdStart = offset + 5;
  const plcPcdEnd = plcPcdStart + plcPcdLength;
  if (plcPcdEnd > clx.length) {
    throw new Error("Invalid Word binary document: truncated PlcPcd");
  }
  if ((plcPcdLength - 4) % 12 !== 0) {
    throw new Error("Invalid Word binary document: PlcPcd length is not valid");
  }

  const plcPcd = clx.subarray(plcPcdStart, plcPcdEnd);
  const pieceCount = (plcPcdLength - 4) / 12;
  const characterPositions = [];
  for (let i = 0; i <= pieceCount; i += 1) {
    characterPositions.push(plcPcd.readUInt32LE(i * 4));
  }

  const pcdStart = (pieceCount + 1) * 4;
  const pieces = [];
  for (let i = 0; i < pieceCount; i += 1) {
    const pcdOffset = pcdStart + i * 8;
    const fcCompressed = plcPcd.readUInt32LE(pcdOffset + 2);
    const compressed = (fcCompressed & 0x40000000) !== 0;
    const fileOffset = compressed ? ((fcCompressed & ~0x40000000) >>> 1) : fcCompressed;
    pieces.push({
      cpStart: characterPositions[i],
      cpEnd: characterPositions[i + 1],
      fileOffset,
      compressed,
    });
  }

  return pieces;
}

function splitSubdocuments(wordDocument, pieces, counts) {
  const ranges = {};
  let cp = 0;
  for (const [name, length] of Object.entries(counts)) {
    ranges[name] = {
      cpStart: cp,
      cpEnd: cp + length,
      rawText: readPieces(wordDocument, pieces, cp, cp + length),
    };
    cp += length;
  }
  return ranges;
}

function readPieces(wordDocument, pieces, cpStart = 0, cpEnd = Infinity) {
  let text = "";
  for (const piece of pieces) {
    const start = Math.max(piece.cpStart, cpStart);
    const end = Math.min(piece.cpEnd, cpEnd);
    if (end <= start) {
      continue;
    }

    const characterOffset = start - piece.cpStart;
    const characterLength = end - start;
    if (piece.compressed) {
      const startByte = piece.fileOffset + characterOffset;
      const endByte = startByte + characterLength;
      assertRange(wordDocument, startByte, endByte, "compressed text piece");
      text += decodeSingleByteText(wordDocument.subarray(startByte, endByte));
    } else {
      const startByte = piece.fileOffset + characterOffset * 2;
      const endByte = startByte + characterLength * 2;
      assertRange(wordDocument, startByte, endByte, "Unicode text piece");
      text += wordDocument.subarray(startByte, endByte).toString("utf16le");
    }
  }
  return text;
}

function paragraphsFromWordText(text) {
  return normalizeWordText(text)
    .split(/\n+/)
    .map((paragraph) => paragraph.trim())
    .filter(Boolean);
}

function assertWordDocument(wordDocument) {
  if (!Buffer.isBuffer(wordDocument)) {
    throw new TypeError("Expected WordDocument to be a Buffer");
  }
  if (wordDocument.length < 0x1aa || wordDocument.readUInt16LE(0) !== WORD_BINARY_MAGIC) {
    throw new Error("Invalid input file: WordDocument stream does not contain a Word binary FIB");
  }
}

function assertRange(buffer, start, end, label) {
  if (start < 0 || end > buffer.length) {
    throw new Error(`Invalid Word binary document: ${label} points outside WordDocument`);
  }
}

function decodeSingleByteText(buffer) {
  return new TextDecoder("windows-1252", { fatal: false }).decode(buffer);
}

function parsePlcfHdd(tableStream, fib) {
  // PlcfHdd is a PLC (CP array, no data entries) in the table stream.
  // It indexes stories in the header subdocument:
  //   First 0-6 entries: footnote/endnote separators
  //   Then groups of up to 6 entries per section (even/odd header/footer, first header/footer)
  if (!fib.fcPlcfHdd || !fib.lcbPlcfHdd || fib.lcbPlcfHdd < 4) {
    return { cpArray: [], separatorCount: 6 };
  }
  if (fib.fcPlcfHdd + fib.lcbPlcfHdd > tableStream.length) {
    return { cpArray: [], separatorCount: 6 };
  }

  const plcf = tableStream.subarray(fib.fcPlcfHdd, fib.fcPlcfHdd + fib.lcbPlcfHdd);
  const cpCount = fib.lcbPlcfHdd / 4;
  if (cpCount < 2) {
    return { cpArray: [], separatorCount: 6 };
  }

  const cpArray = [];
  for (let i = 0; i < cpCount; i += 1) {
    cpArray.push(plcf.readUInt32LE(i * 4));
  }
  // MS-DOC: the first 6 separators (footnote/endnote) always exist
  return { cpArray, separatorCount: 6 };
}

function parseStandardBookmarks(tableStream, fib) {
  if (!fib.lcbSttbfBkmk && !fib.lcbPlcfBkf && !fib.lcbPlcfBkl) return [];
  if (!fib.lcbSttbfBkmk || !fib.lcbPlcfBkf || !fib.lcbPlcfBkl) {
    throw new Error("Invalid Word binary document: incomplete standard bookmark tables");
  }
  assertTableRange(tableStream, fib.fcSttbfBkmk, fib.lcbSttbfBkmk, "SttbfBkmk");
  assertTableRange(tableStream, fib.fcPlcfBkf, fib.lcbPlcfBkf, "Plcfbkf");
  assertTableRange(tableStream, fib.fcPlcfBkl, fib.lcbPlcfBkl, "Plcfbkl");

  const names = parseSttbfBkmk(tableStream.subarray(fib.fcSttbfBkmk, fib.fcSttbfBkmk + fib.lcbSttbfBkmk));
  if (!names.length) return [];

  // MS-DOC-SPEC/15: SttbfBkmk, Plcfbkf, and Plcfbkl are parallel tables.
  // LibreOffice's WW8PLCFx_Book reads Plcfbkf with a 4-byte BKF.ibkl data
  // element that indexes the end CP in Plcfbkl.
  const bkf = tableStream.subarray(fib.fcPlcfBkf, fib.fcPlcfBkf + fib.lcbPlcfBkf);
  const bkl = tableStream.subarray(fib.fcPlcfBkl, fib.fcPlcfBkl + fib.lcbPlcfBkl);
  const dataBytes = fib.lcbPlcfBkf - (names.length + 1) * 4;
  if (dataBytes < 0 || dataBytes % names.length !== 0) {
    throw new Error("Invalid Word binary document: malformed Plcfbkf bookmark PLC");
  }
  const bkfDataSize = dataBytes / names.length;
  if (bkfDataSize < 4) {
    throw new Error(`Out-of-spec Word bookmark FBKF size ${bkfDataSize}`);
  }
  if (fib.lcbPlcfBkl !== (names.length + 1) * 4) {
    throw new Error("Invalid Word binary document: malformed Plcfbkl bookmark PLC");
  }

  return names.map((name, index) => {
    const start = bkf.readUInt32LE(index * 4);
    const dataOffset = (names.length + 1) * 4 + index * bkfDataSize;
    const endIndex = bkf.readUInt32LE(dataOffset);
    if (endIndex >= names.length) {
      throw new Error(`Invalid Word bookmark ${name}: end index ${endIndex} is outside Plcfbkl`);
    }
    const end = bkl.readUInt32LE(endIndex * 4);
    return { id: index, name, cpStart: start, cpEnd: end };
  });
}

function parseRevisionAuthors(tableStream, fib) {
  if (!fib.lcbSttbfRMark) {
    return ["Unknown"];
  }
  assertTableRange(tableStream, fib.fcSttbfRMark, fib.lcbSttbfRMark, "SttbfRMark");
  return parseSttbfRMark(tableStream.subarray(fib.fcSttbfRMark, fib.fcSttbfRMark + fib.lcbSttbfRMark));
}

function parseLastSelection(tableStream, fib, bodyTextLength) {
  if (!fib.lcbWss) return null;
  if (fib.lcbWss !== SELSF_SIZE) {
    throw new Error(`Out-of-spec Word binary document: expected ${SELSF_SIZE}-byte Selsf, got ${fib.lcbWss}`);
  }
  assertTableRange(tableStream, fib.fcWss, fib.lcbWss, "Selsf");

  // MS-DOC-SPEC/15 fcWss/lcbWss points to Selsf, and MS-DOC-SPEC/19
  // defines Selsf CPs relative to the document text piece, not just the
  // main-story body text that this converter emits to word/document.xml.
  const selsf = tableStream.subarray(fib.fcWss, fib.fcWss + SELSF_SIZE);
  const flags = selsf.readUInt32LE(0);
  const cpFirst = selsf.readInt32LE(4);
  const cpLim = selsf.readInt32LE(8);
  const selection = {
    fIns: ((flags >> 15) & 1) !== 0,
    fForward: (flags >> 16) & 0x7f,
    fInsEnd: (flags >> 24) & 0xff,
    cpFirst,
    cpLim,
    cpAnchor: selsf.readInt32LE(20),
    sty: selsf.readUInt16LE(24),
    cpAnchorShrink: selsf.readInt32LE(28),
    xaTableLeft: selsf.readInt16LE(32),
    xaTableRight: selsf.readInt16LE(34),
  };

  if (selection.cpFirst < 0 || selection.cpLim < selection.cpFirst || selection.cpLim > bodyTextLength) {
    throw new Error(`Invalid Word binary document: Selsf selection range ${selection.cpFirst}-${selection.cpLim} is outside the document text`);
  }
  if (selection.fIns && selection.cpFirst !== selection.cpLim) {
    throw new Error("Invalid Word binary document: Selsf insertion point has mismatched cpFirst/cpLim");
  }
  return selection;
}

function assertTableRange(tableStream, fc, lcb, label) {
  if (fc + lcb > tableStream.length) {
    throw new Error(`Invalid Word binary document: ${label} is outside the table stream`);
  }
}

export function parseSttbfRMark(sttbf) {
  // MS-DOC-SPEC/19 SttbfRMark: extended Unicode STTB, no extra data, first
  // entry MUST be "Unknown".
  const authors = parseUnicodeSttbNoExtra(sttbf, "SttbfRMark");
  if (authors[0] !== "Unknown") {
    throw new Error(`Out-of-spec SttbfRMark: first author must be Unknown, got ${JSON.stringify(authors[0] ?? "")}`);
  }
  return authors;
}

function parseSttbfBkmk(sttbf) {
  const names = parseUnicodeSttbNoExtra(sttbf, "SttbfBkmk");
  for (const name of names) {
    if (name.length === 0 || name.length >= 40) {
      throw new Error(`Invalid Word bookmark name length ${name.length}`);
    }
  }
  return names;
}

function parseUnicodeSttbNoExtra(sttbf, label) {
  if (sttbf.length < 6) {
    throw new Error(`Invalid Word binary document: truncated ${label}`);
  }
  const fExtend = sttbf.readUInt16LE(0);
  if (fExtend !== 0xffff) {
    throw new Error(`Out-of-spec ${label}: expected extended strings, got 0x${fExtend.toString(16)}`);
  }
  const count = sttbf.readUInt16LE(2);
  const cbExtra = sttbf.readUInt16LE(4);
  if (cbExtra !== 0) {
    throw new Error(`Out-of-spec ${label}: expected no extra data, got ${cbExtra} bytes`);
  }

  const strings = [];
  let off = 6;
  for (let i = 0; i < count; i += 1) {
    if (off + 2 > sttbf.length) {
      throw new Error(`Invalid Word binary document: truncated ${label} string length`);
    }
    const cch = sttbf.readUInt16LE(off);
    off += 2;
    const byteLength = cch * 2;
    if (off + byteLength > sttbf.length) {
      throw new Error(`Invalid Word binary document: truncated ${label} string data`);
    }
    strings.push(sttbf.subarray(off, off + byteLength).toString("utf16le"));
    off += byteLength;
  }
  return strings;
}

function extractParagraphProperties(wordDocument, tableStream, fib, bodyText, styles, pieces) {
  if (!fib.fcPapx || !fib.lcbPapx || fib.lcbPapx < 4) {
    return [];
  }
  if (fib.fcPapx + fib.lcbPapx > tableStream.length) {
    return [];
  }

  const plcf = tableStream.subarray(fib.fcPapx, fib.fcPapx + fib.lcbPapx);
  const binCount = (fib.lcbPapx - 4) / 8;
  if (binCount <= 0 || !Number.isInteger(binCount)) {
    return [];
  }

  const binFCs = [];
  for (let i = 0; i <= binCount; i += 1) {
    binFCs.push(plcf.readUInt32LE(i * 4));
  }

  const pageNumbers = [];
  for (let bi = 0; bi < binCount; bi += 1) {
    pageNumbers.push(plcf.readUInt32LE((binCount + 1) * 4 + bi * 4));
  }

  const bodyParagraphRanges = getParagraphRanges(bodyText);
  const allEntries = [];

  for (let bi = 0; bi < binCount; bi += 1) {
    const pageNumber = pageNumbers[bi];
    if (pageNumber === 0) continue;
    const pageOffset = pageNumber * 512;
    if (pageOffset + PHE_SIZE > wordDocument.length) continue;

    const page = wordDocument.subarray(pageOffset, Math.min(pageOffset + 512, wordDocument.length));
    const entries = parsePapxEntries(page, styles, pieces, fib.characterCounts.body);
    allEntries.push(...entries);
  }

  return bodyParagraphRanges.map((paragraphRange) => {
    const direct = allEntries.find((entry) => rangesOverlap(entry, paragraphRange)) ?? null;
    return direct?.properties ?? null;
  });
}

function countPageFCs(page, binStart, binEnd) {
  let count = 0;
  for (let i = 0; i * 4 + 4 <= page.length; i += 1) {
    const fc = page.readUInt32LE(i * 4);
    if (fc < binStart || fc > binEnd) break;
    count += 1;
  }
  return count;
}

function parsePapxEntries(page, styles, pieces, bodyCharacterCount) {
  const entries = [];
  const crun = page[511];
  if (crun <= 0) return entries;

  const fcBoundaries = readFkpFcBoundaries(page, crun);
  const bxStart = (crun + 1) * 4;

  for (let i = 0; i < crun; i += 1) {
    const bxOffset = bxStart + i * (1 + PHE_SIZE);
    if (bxOffset >= 511) break;
    const papxOffset = page[bxOffset] * 2;
    if (papxOffset <= 0 || papxOffset >= 511) continue;
    const cb = page[papxOffset];
    let dataOffset;
    let byteLength;
    if (cb === 0) {
      const cbx = page[papxOffset + 1];
      if (cbx === 0 || papxOffset + 2 + cbx * 2 > page.length) continue;
      dataOffset = papxOffset + 2;
      byteLength = cbx * 2;
    } else {
      byteLength = cb * 2;
      if (papxOffset + 1 + byteLength > page.length) continue;
      dataOffset = papxOffset + 1;
    }

    const cpStart = fileOffsetToCharacterPosition(fcBoundaries[i], pieces);
    const cpEnd = fileOffsetToCharacterPosition(fcBoundaries[i + 1], pieces);
    if (cpStart == null || cpEnd == null || cpStart >= bodyCharacterCount || cpEnd <= 0) continue;

    const data = page.subarray(dataOffset, dataOffset + byteLength);
    const properties = parseParagraphGrpprl(data);
    const istd = properties.istd ?? 0;
    properties.styleId = resolveStyleId(istd, styles);
    entries.push({
      cpStart: Math.max(0, cpStart),
      cpEnd: Math.min(bodyCharacterCount, cpEnd),
      properties,
    });
  }

  return entries;
}

function readFkpFcBoundaries(page, crun) {
  const boundaries = [];
  for (let i = 0; i <= crun; i += 1) {
    boundaries.push(page.readUInt32LE(i * 4));
  }
  return boundaries;
}

function resolveStyleId(istd, styles) {
  if (!styles) return null;
  if (istd === 0) return null;
  if (istd >= 0 && istd < styles.length && styles[istd]) {
    return styles[istd].styleId;
  }
  for (const style of styles) {
    if (style && style.type === STYLE_TYPE_PARAGRAPH && style.basedOn === null) {
      return style.styleId;
    }
  }
  return null;
}

function resolveCharacterStyleId(istd, styles) {
  if (!styles) return null;
  // MS-DOC-SPEC/16 sprmCIstd defaults to istd 0x000A, Default Paragraph Font;
  // WPS omits rStyle for that default character style.
  if (istd === 0x000a) return null;
  const style = styles[istd];
  return style?.type === STYLE_TYPE_CHARACTER ? style.styleId : null;
}

function parseParagraphGrpprl(data) {
  const parsed = parseSprms(data, true);
  return {
    istd: parsed.istd ?? null,
    styleId: null,
    lineSpacing: parsed.lineSpacing ?? null,
    alignment: parsed.alignment ?? null,
    leftIndentChars: parsed.leftIndentChars ?? null,
    rightIndentChars: parsed.rightIndentChars ?? null,
    firstLineIndentChars: parsed.firstLineIndentChars ?? null,
    leftIndent: parsed.leftIndent ?? null,
    rightIndent: parsed.rightIndent ?? null,
    firstLineIndent: parsed.firstLineIndent ?? null,
    spacingBefore: parsed.spacingBefore ?? null,
    spacingAfter: parsed.spacingAfter ?? null,
    spacingBeforeLines: parsed.spacingBeforeLines ?? null,
    spacingAfterLines: parsed.spacingAfterLines ?? null,
    spacingBeforeAuto: parsed.spacingBeforeAuto ?? null,
    spacingAfterAuto: parsed.spacingAfterAuto ?? null,
    listLevel: parsed.listLevel ?? null,
    listId: parsed.listId ?? null,
    tabs: parsed.tabs ?? null,
    inTable: parsed.inTable ?? false,
    keepLines: parsed.keepLines ?? null,
    keepNext: parsed.keepNext ?? null,
    pageBreakBefore: parsed.pageBreakBefore ?? null,
    suppressAutoHyphens: parsed.suppressAutoHyphens ?? null,
    widowControl: parsed.widowControl ?? null,
    contextualSpacing: parsed.contextualSpacing ?? null,
    mirrorIndents: parsed.mirrorIndents ?? null,
    bidi: parsed.bidi ?? null,
    snapToGrid: parsed.snapToGrid ?? null,
    textAlignment: parsed.textAlignment ?? null,
    kinsoku: parsed.kinsoku ?? null,
    wordWrap: parsed.wordWrap ?? null,
    overflowPunct: parsed.overflowPunct ?? null,
    topLinePunct: parsed.topLinePunct ?? null,
    autoSpaceDE: parsed.autoSpaceDE ?? null,
    autoSpaceDN: parsed.autoSpaceDN ?? null,
    adjustRightInd: parsed.adjustRightInd ?? null,
    lineNumberCount: parsed.lineNumberCount ?? null,
    paragraphBorders: parsed.paragraphBorders ?? null,
    paragraphShading: parsed.paragraphShading ?? null,
    tablePosition: parsed.tablePosition ?? null,
    tableNoAllowOverlap: parsed.tableNoAllowOverlap ?? null,
  };
}

const STYLE_TYPE_PARAGRAPH = "paragraph";
const STYLE_TYPE_CHARACTER = "character";
const STYLE_TYPE_TABLE = "table";
const STYLE_TYPE_LIST = "list";

function buildStyleId(sti, index, styles) {
  // styleId is sequential (index+1) so paragraph property istd references match
  return String(index + 1);
}

function buildStyleName(sti, name) {
  // MS-OI29500: English name from sti; fall back to raw name for custom styles
  if (sti < STI_NAMES.length && STI_NAMES[sti]) return STI_NAMES[sti];
  return name;
}

function buildStyleType(sgc) {
  // stk (style kind) per MS-DOC StdfBase: 1=paragraph, 2=character, 3=table, 4=list
  if (sgc === 2) return STYLE_TYPE_CHARACTER;
  if (sgc === 3) return STYLE_TYPE_TABLE;
  // sgc === 4 is list/numbering style; treat as paragraph for now
  return STYLE_TYPE_PARAGRAPH;
}

function extractStyleSheet(tableStream, fib) {
  if (!fib.lcbStsh || fib.lcbStsh < 6) {
    return [];
  }
  if (fib.fcStsh + fib.lcbStsh > tableStream.length) {
    return [];
  }

  const stsh = tableStream.subarray(fib.fcStsh, fib.fcStsh + fib.lcbStsh);
  const cbStshi = stsh.readUInt16LE(0);
  if (cbStshi < 4 || 2 + cbStshi > stsh.length) {
    return [];
  }

  const cstd = stsh.readUInt16LE(2);
  const cbSTDBaseInFile = stsh.readUInt16LE(4);
  if (cbSTDBaseInFile !== 10 && cbSTDBaseInFile !== 18) {
    throw new Error(`Out-of-spec Word stylesheet: invalid Stshif.cbSTDBaseInFile ${cbSTDBaseInFile}`);
  }
  const stiMaxWhenSaved = stsh.readUInt16LE(8); // Stshif offset 6 from stshi start
  const styles = new Array(cstd).fill(null);

  // Parse StshiLsd at STSH offset 22 (after cbStshi 2 + stshif 18 + ftcBi 2).
  // Each entry is a 4-byte LSD structure.  The count equals stiMaxWhenSaved.
  const latentLsd = [];
  const lsdOffset = 22;
  if (lsdOffset + 2 <= 2 + cbStshi) {
    const cbLSD = stsh.readUInt16LE(lsdOffset);
    if (cbLSD === 4 && stiMaxWhenSaved > 0) {
      const mpstiilsdStart = lsdOffset + 2;
      for (let i = 0; i < stiMaxWhenSaved; i++) {
        const off = mpstiilsdStart + i * 4;
        if (off + 4 > 2 + cbStshi) break;
        const bits = stsh.readUInt16LE(off);
        latentLsd.push({
          fLocked: (bits & 0x0001) !== 0,
          fSemiHidden: (bits & 0x0002) !== 0,
          fUnhideWhenUsed: (bits & 0x0004) !== 0,
          fQFormat: (bits & 0x0008) !== 0,
          iPriority: (bits >> 4) & 0x0FFF,
        });
      }
    }
  }


  let off = 2 + cbStshi;
  for (let i = 0; i < cstd && off + 2 <= stsh.length; i += 1) {
    const cbStd = stsh.readUInt16LE(off);
    off += 2;
    if (cbStd === 0) continue;
    const stdEnd = off + cbStd;
    if (stdEnd > stsh.length) break;

    const std = stsh.subarray(off, stdEnd);
    const style = parseStd(std, i, cbSTDBaseInFile);
    if (style) styles[i] = style;

    off = stdEnd;
  }

  // Assign non-null styles sequential order and attach latent data
  let order = 0;
  const nonNullStyles = [];
  const defaults = new Set([0, 65, 105]); // sti values of Normal, DPF, Normal Table
  for (let i = 0; i < styles.length; i++) {
    const style = styles[i];
    if (!style) continue;
    style.order = order;
    style._stshIndex = i; // preserve original STSH index for ordering
    // Attach parsed latent style data for this style's sti
    const latent = latentLsd[style.sti];
    if (latent) {
      style.latent = latent;
    }
    nonNullStyles.push(style);
    order += 1;
  }
  if (!nonNullStyles.some((style) => style.sti === 105)) {
    const normalTable = createSyntheticNormalTableStyle(styles.length);
    normalTable.order = order;
    normalTable._stshIndex = normalTable.index;
    normalTable.latent = latentLsd[normalTable.sti] ?? null;
    styles.push(normalTable);
    nonNullStyles.push(normalTable);
    order += 1;
  }

  // MS-DOC StdfBase stores style relationships as istd indexes into STSH.
  // MS-DOC-SPEC/19 StdfBase defines istdNext as the STSH index of the style
  // automatically applied after the current style. WPS promotes Normal's
  // non-self next style before the built-in paragraph block, then assigns
  // built-in paragraph styles before the list-style range, Normal Table,
  // built-in table styles, Default Paragraph Font, built-in character styles,
  // and finally the remaining STSH styles. MS-OI29500 fixes built-in style
  // identity through sti, so the sequence is based on parsed sti/type rather
  // than raw STSH slot numbers.
  const normalStyle = nonNullStyles.find((style) => style.sti === 0);
  const normalNextStyle = normalStyle?.nextCode != null
    && normalStyle.nextCode !== normalStyle.index
    && normalStyle.nextCode < 0xfff0
    ? styles[normalStyle.nextCode]
    : null;
  const promotedStyles = new Set();
  if (normalNextStyle?.type === STYLE_TYPE_PARAGRAPH && normalNextStyle.sti !== 0) {
    promotedStyles.add(normalNextStyle);
  }
  const styleIdOrder = [
    ...nonNullStyles.filter((style) => style.sti === 0),
    ...promotedStyles,
    ...nonNullStyles.filter((style) => (
      style.sti !== 0
      && !defaults.has(style.sti)
      && style.type === STYLE_TYPE_PARAGRAPH
      && style.sti < 179
      && !promotedStyles.has(style)
    )),
    ...nonNullStyles.filter((style) => style.sti === 105),
    ...nonNullStyles.filter((style) => (
      style.sti !== 105
      && style.type === STYLE_TYPE_TABLE
      && style.sti < 179
    )),
    ...nonNullStyles.filter((style) => style.sti === 65),
    ...nonNullStyles.filter((style) => (
      style.sti !== 65
      && style.sti !== 105
      && style.type === STYLE_TYPE_CHARACTER
      && style.sti < 179
    )),
    ...nonNullStyles.filter((style) => (
      style.sti !== 0
      && style.sti !== 65
      && style.sti !== 105
      && !promotedStyles.has(style)
      && !(style.type === STYLE_TYPE_PARAGRAPH && style.sti < 179)
      && !(style.type === STYLE_TYPE_TABLE && style.sti < 179)
      && !(style.type === STYLE_TYPE_CHARACTER && style.sti < 179)
    )),
  ];
  for (let i = 0; i < styleIdOrder.length; i += 1) {
    styleIdOrder[i].styleId = String(i + 1);
  }

  return { styles, latentLsd, stiMaxWhenSaved };
}

function createSyntheticNormalTableStyle(index) {
  return {
    index,
    name: "普通表格",
    styleName: "Normal Table",
    sti: 105,
    sgc: 3,
    type: STYLE_TYPE_TABLE,
    styleId: null,
    basedOn: null,
    next: index,
    baseCode: 0x0fff,
    nextCode: index,
    linkCode: 0,
    lineSpacing: null,
    alignment: null,
    leftIndentChars: null,
    rightIndentChars: null,
    firstLineIndentChars: null,
    leftIndent: null,
    rightIndent: null,
    firstLineIndent: null,
    spacingBefore: null,
    spacingAfter: null,
    spacingBeforeLines: null,
    spacingAfterLines: null,
    spacingBeforeAuto: null,
    spacingAfterAuto: null,
    listLevel: null,
    listId: null,
    tabs: null,
    keepLines: null,
    keepNext: null,
    pageBreakBefore: null,
    suppressAutoHyphens: null,
    widowControl: null,
    bidi: null,
    snapToGrid: null,
    textAlignment: null,
    kinsoku: null,
    wordWrap: null,
    overflowPunct: null,
    topLinePunct: null,
    autoSpaceDE: null,
    autoSpaceDN: null,
    adjustRightInd: null,
    lineNumberCount: null,
    paragraphBorders: null,
    paragraphShading: null,
    mirrorIndents: null,
    suppressOverlap: null,
    outlineLevel: null,
    frameWidth: null,
    frameHeight: null,
    frameHRule: null,
    frameX: null,
    frameY: null,
    frameXAlign: null,
    frameYAlign: null,
    frameHAnchor: null,
    frameVAnchor: null,
    frameWrap: null,
    frameLocked: null,
    frameHSpace: null,
    frameVSpace: null,
    frameDropCap: null,
    frameLines: null,
    runProperties: null,
    synthetic: true,
  };
}

function parseStd(std, index, cbSTDBaseInFile) {
  if (std.length < cbSTDBaseInFile + 2) return null;

  const sti = std.readUInt16LE(0) & 0x0fff;
  const sgc = std.readUInt16LE(2) & 0x000f;
  const istdBase = std.readUInt16LE(2) >> 4;
  const istdNext = std.readUInt16LE(4) >> 4;
  // MS-DOC-SPEC/19 StdfBase.grfstd stores the style behavior bits
  // (semiHidden, unhideWhenUsed, qFormat). StdfPost2000, when present
  // per Stshif.cbSTDBaseInFile, stores istdLink and iPriority.
  const grfstd = std.readUInt16LE(8);
  const hasStdfPost2000 = cbSTDBaseInFile === STSH_STD_HEADER_SIZE_WITH_POST2000;
  const istdLink = hasStdfPost2000 ? (std.readUInt16LE(10) & 0x0fff) : 0;
  const uiPriority = hasStdfPost2000 ? ((std.readUInt16LE(16) >> 4) & 0x0fff) : null;

  const cbName = std.readUInt16LE(cbSTDBaseInFile);
  if (cbName < 1 || cbName > 50) return null;

  const nameStart = cbSTDBaseInFile + 2;
  const nameEnd = nameStart + cbName * 2;
  if (nameEnd + 2 > std.length) return null;

  const name = std.subarray(nameStart, nameEnd).toString("utf16le");
  const grpprlStart = nameEnd + 2;
  const grpprl = std.subarray(grpprlStart);
  // cupx (count of Upx) per MS-DOC StdfBase: para=2, char=0, table=3, list=1
  const cupxMap = { 1: 2, 2: 0, 3: 3, 4: 1 };
  const cupx = cupxMap[sgc] || 0;
  const parsed = parseSprms(grpprl, true, cupx);
  // Fallback raw-grpprl scanners catch properties that parseSprms may miss
  // when Upx data (GrLPUpxSw) shifts the SPRM block. The scanners are cheap
  // O(n) passes on small buffers; the parsed.* == null guards prevent
  // overwriting values already correctly captured by parseSprms.
  if (parsed.outlineLevel == null) {
    parsed.outlineLevel = extractOutlineLevelFromGrpprl(grpprl);
  }
  if (parsed.paragraphBorders == null || Object.values(parsed.paragraphBorders).every(v => v == null)) {
    const fbBorders = extractParagraphBordersFromGrpprl(grpprl);
    if (fbBorders) parsed.paragraphBorders = fbBorders;
  }

  let runProperties = extractCharacterPropertiesFromGrpprl(grpprl);
  // extractCharacterPropertiesFromGrpprl scans from offset 0 (including istd+cupx),
  // which may miss or misread character SPRMs. For character styles, merge the
  // correctly-parsed character properties from parseSprms.
  if (sgc === 2) {
    const characterPropertyNames = [
      "underline",
      "underlineStyle",
      "underlineColor",
      "textColor",
      "border",
      "background",
      "verticalAlign",
      "bold",
      "italic",
      "outline",
      "shadow",
      "imprint",
      "emboss",
      "noProof",
      "webHidden",
      "specVanish",
    ];
    for (const name of characterPropertyNames) {
      if (parsed[name] != null) {
        (runProperties ??= {})[name] = parsed[name];
      }
    }
  }
  const type = buildStyleType(sgc);
  const styleName = buildStyleName(sti, name);

  return {
    index,
    name,
    styleName,
    sti,
    sgc,
    type,
    styleId: null, // assigned later in extractStyleSheet pass 2
    basedOn: resolveBasedOn(istdBase, index),
    next: istdNext,
    baseCode: istdBase,
    nextCode: istdNext,
    linkCode: istdLink,
    grfstd,
    styleFlags: {
      semiHidden: (grfstd & 0x0100) !== 0,
      unhideWhenUsed: (grfstd & 0x0800) !== 0,
      qFormat: (grfstd & 0x1000) !== 0,
    },
    uiPriority,
    lineSpacing: parsed.lineSpacing ?? extractLineSpacingFromGrpprl(grpprl),
    alignment: parsed.alignment ?? null,
    leftIndentChars: parsed.leftIndentChars ?? null,
    rightIndentChars: parsed.rightIndentChars ?? null,
    firstLineIndentChars: parsed.firstLineIndentChars ?? null,
    leftIndent: parsed.leftIndent ?? null,
    rightIndent: parsed.rightIndent ?? null,
    firstLineIndent: parsed.firstLineIndent ?? null,
    spacingBefore: parsed.spacingBefore ?? null,
    spacingAfter: parsed.spacingAfter ?? null,
    spacingBeforeLines: parsed.spacingBeforeLines ?? null,
    spacingAfterLines: parsed.spacingAfterLines ?? null,
    spacingBeforeAuto: parsed.spacingBeforeAuto ?? null,
    spacingAfterAuto: parsed.spacingAfterAuto ?? null,
    contextualSpacing: parsed.contextualSpacing ?? null,
    listLevel: parsed.listLevel ?? null,
    listId: parsed.listId ?? null,
    tabs: parsed.tabs ?? null,
    keepLines: parsed.keepLines ?? null,
    keepNext: parsed.keepNext ?? null,
    pageBreakBefore: parsed.pageBreakBefore ?? null,
    suppressAutoHyphens: parsed.suppressAutoHyphens ?? null,
    widowControl: parsed.widowControl ?? null,
    bidi: parsed.bidi ?? null,
    snapToGrid: parsed.snapToGrid ?? null,
    textAlignment: parsed.textAlignment ?? null,
    kinsoku: parsed.kinsoku ?? null,
    wordWrap: parsed.wordWrap ?? null,
    overflowPunct: parsed.overflowPunct ?? null,
    topLinePunct: parsed.topLinePunct ?? null,
    autoSpaceDE: parsed.autoSpaceDE ?? null,
    autoSpaceDN: parsed.autoSpaceDN ?? null,
    adjustRightInd: parsed.adjustRightInd ?? null,
    lineNumberCount: parsed.lineNumberCount ?? null,
    paragraphBorders: parsed.paragraphBorders ?? null,
    paragraphShading: parsed.paragraphShading ?? null,
    mirrorIndents: parsed.mirrorIndents ?? null,
    suppressOverlap: parsed.suppressOverlap ?? null,
    // MS-DOC-SPEC/16 sprmPOutLvl: "This MUST be ignored if the paragraph has
    // an istd that is greater than or equal to 0x1 and less than or equal
    // to 0x9." For these heading styles (sti 1–9), outlineLevel = sti - 1.
    outlineLevel: (sti >= 1 && sti <= 9) ? sti - 1 : (parsed.outlineLevel ?? null),
    frameWidth: parsed.frameWidth ?? null,
    frameHeight: parsed.frameHeight ?? null,
    frameHRule: parsed.frameHRule ?? null,
    frameX: parsed.frameX ?? null,
    frameY: parsed.frameY ?? null,
    frameXAlign: parsed.frameXAlign ?? null,
    frameYAlign: parsed.frameYAlign ?? null,
    frameHAnchor: parsed.frameHAnchor ?? null,
    frameVAnchor: parsed.frameVAnchor ?? null,
    frameWrap: parsed.frameWrap ?? null,
    frameLocked: parsed.frameLocked ?? null,
    frameHSpace: parsed.frameHSpace ?? null,
    frameVSpace: parsed.frameVSpace ?? null,
    frameDropCap: parsed.frameDropCap ?? null,
    frameLines: parsed.frameLines ?? null,
    runProperties,
  };
}

function resolveBasedOn(istdBase, index) {
  if (istdBase === 0x0fff || istdBase >= STSH_NIL_BASE) return null;
  if (istdBase === index) return null;
  return istdBase;
}

function extractLineSpacingFromGrpprl(grpprl) {
  for (let i = 0; i + 3 < grpprl.length; i += 1) {
    if (grpprl[i] === (SPRM_WPS_DYA_LINE & 0xff) && grpprl[i + 1] === ((SPRM_WPS_DYA_LINE >> 8) & 0xff)) {
      // Byte i+4 (if available) encodes the rule: 0=atLeast, 1=auto, 2+=exact
      const ruleByte = i + 4 < grpprl.length ? grpprl[i + 4] : null;
      let rule, twips;
      if (ruleByte === 1) {
        twips = grpprl[i + 2] | (grpprl[i + 3] << 8);
        rule = "auto";
      } else if (ruleByte != null && ruleByte >= 2) {
        twips = grpprl[i + 2] | (grpprl[i + 3] << 8);
        rule = "exact";
      } else {
        const raw = grpprl[i + 2] | (grpprl[i + 3] << 8);
        const signed = raw > 0x7fff ? raw - 0x10000 : raw;
        twips = Math.abs(signed);
        rule = signed < 0 ? "exact" : "atLeast";
      }
      return { twips, rule };
    }
  }
  return null;
}

// SPRM 0x2640: outline level (1-byte operand, per MS-DOC-SPEC §sprmPOutLvl)
function extractOutlineLevelFromGrpprl(grpprl) {
  // Scan the full buffer; the SPRM may appear before or after the Upx data.
  for (let i = 0; i + 2 < grpprl.length; i += 1) {
    const sprm = grpprl[i] | (grpprl[i + 1] << 8);
    if (sprm === 0x2640 && i + 3 <= grpprl.length) {
      const raw = grpprl[i + 2];
      return raw <= 9 ? raw : null; // match applySprm clamp per MS-DOC-SPEC sprmPOutLvl
    }
  }
  return null;
}

// SPRMs 0x6424–0x6428, 0x6629: paragraph border Brc80 (4-byte operands)
// per MS-DOC-SPEC §sprmPBrcTop80, sprmPBrcLeft80, etc.
// Brc80 per MS-DOC-SPEC §Brc80: dptLineWidth(1) + brcType(1) + ico(1) + dptSpace(1)
function extractParagraphBordersFromGrpprl(grpprl) {
  const borders = {};
  const BORDER_SPRMS = { 0x6424: "top", 0x6425: "left", 0x6426: "bottom", 0x6427: "right", 0x6428: "between", 0x6629: "bar" };
  for (let i = 0; i + 5 < grpprl.length; i += 1) {
    const sprm = grpprl[i] | (grpprl[i + 1] << 8);
    const side = BORDER_SPRMS[sprm];
    if (side && borders[side] == null) {
      borders[side] = parseBrc80Raw(grpprl[i + 2], grpprl[i + 3], grpprl[i + 4], grpprl[i + 5]);
      i += 5; // skip 4-byte operand (+ loop's i++ = net 6 = 2-byte SPRM + 4-byte Brc80)
    }
  }
  return Object.keys(borders).length > 0 ? borders : null;
}

// Parse Brc80 from raw bytes. Uses the shared BRC_TYPE_NAMES/brcColorFromIco from sprm.js.
// Unlike the primary parseBrc80(), this does NOT return null for brcType===0 (none),
// because "none" borders carry meaningful dptSpace values that OOXML needs.
function parseBrc80Raw(dptLineWidth, brcType, ico, dptSpace) {
  const borderName = BRC_TYPE_NAMES[brcType];
  if (!borderName) {
    throw new Error(`Out-of-spec paragraph border BrcType ${brcType}`);
  }
  return {
    val: borderName,
    sz: String(dptLineWidth),
    color: brcColorFromIco(ico),
    space: String(dptSpace & 0x1F), // dptSpace per MS-DOC-SPEC §Brc80: only bits 0-4
  };
}

function extractCharacterPropertiesFromGrpprl(grpprl) {
  const props = {};
  scanKnownSprm(grpprl, 0x4a43, 2, (value) => { props.fontSize = value.readUInt16LE(0); });
  scanKnownSprm(grpprl, 0x4845, 2, (value) => { props.textPosition = value.readInt16LE(0); });
  scanKnownSprm(grpprl, 0x2a48, 1, (value) => {
    if (value[0] === 1) props.verticalAlign = "superscript";
    else if (value[0] === 2) props.verticalAlign = "subscript";
    else if (value[0] > 2) throw new Error(`Out-of-spec MS-DOC Iss superscript/subscript value ${value[0]}`);
  });
  scanKnownSprm(grpprl, 0x4a61, 2, (value) => { props.fontSizeCs = value.readUInt16LE(0); });
  scanKnownSprm(grpprl, 0x4a4f, 2, (value) => { props.fontAscii = value.readUInt16LE(0); });
  scanKnownSprm(grpprl, 0x4a50, 2, (value) => { props.fontEastAsia = value.readUInt16LE(0); });
  scanKnownSprm(grpprl, 0x4a51, 2, (value) => { props.fontHAnsi = value.readUInt16LE(0); });
  scanKnownSprm(grpprl, 0x4a5e, 2, (value) => { props.fontCs = value.readUInt16LE(0); });
  scanKnownSprm(grpprl, 0x286f, 1, (value) => { props.fontHint = value[0] === 1 ? "eastAsia" : "default"; });
  scanKnownSprm(grpprl, 0x4852, 2, (value) => { props.charWidth = value.readUInt16LE(0); });
  scanKnownSprm(grpprl, 0x8840, 2, (value) => { props.charSpacing = value.readInt16LE(0); });
  scanKnownSprm(grpprl, 0x484b, 2, (value) => { props.kern = value.readUInt16LE(0); });
  scanKnownSprm(grpprl, 0x0835, 1, (value) => { props.bold = value[0] !== 0; });
  scanKnownSprm(grpprl, 0x2a02, 1, (value) => { props.bold = value[0] !== 0; });
  scanKnownSprm(grpprl, 0x085c, 1, (value) => { props.boldCs = value[0] !== 0; });
  scanKnownSprm(grpprl, 0x0838, 1, (value) => { props.outline = value[0] !== 0; });
  scanKnownSprm(grpprl, 0x0839, 1, (value) => { props.shadow = value[0] !== 0; });
  scanKnownSprm(grpprl, 0x0854, 1, (value) => { props.imprint = value[0] !== 0; });
  scanKnownSprm(grpprl, 0x0858, 1, (value) => { props.emboss = value[0] !== 0; });
  scanKnownSprm(grpprl, 0x0875, 1, (value) => { props.noProof = value[0] !== 0; });
  scanKnownSprm(grpprl, 0x0811, 1, (value) => { props.webHidden = value[0] !== 0; });
  scanKnownSprm(grpprl, 0x0818, 1, (value) => { props.specVanish = value[0] !== 0; });
  scanKnownSprm(grpprl, 0x2a42, 1, (value) => { props.textColor = styleTextColorHex(value[0]); });
  scanKnownSprm(grpprl, 0x6870, 4, (value) => { props.textColor = colorRefToHex(value); });
  scanKnownSprm(grpprl, 0x6877, 4, (value) => { props.underlineColor = colorRefToHex(value); });
  scanKnownSprm(grpprl, 0x2a0e, 1, (value) => {
    if (value[0] === 0) props.underline = false;
  });
  scanKnownSprm(grpprl, 0x2a3e, 1, (value) => {
    if (value[0] === 0) props.underline = false;
  });
  scanKnownSprm(grpprl, 0x0868, 1, (value) => { props.charSnapToGrid = value[0] !== 0; });
  scanKnownSprm(grpprl, 0x485f, 2, (value) => { props.langIdBidi = value.readUInt16LE(0); });
  scanKnownSprm(grpprl, 0x486d, 2, (value) => { props.langId = value.readUInt16LE(0); });
  scanKnownSprm(grpprl, 0x486e, 2, (value) => { props.langIdEastAsia = value.readUInt16LE(0); });
  scanKnownSprm(grpprl, 0x4873, 2, (value) => { props.langId = value.readUInt16LE(0); });
  scanKnownSprm(grpprl, 0x4874, 2, (value) => { props.langIdEastAsia = value.readUInt16LE(0); });
  return Object.keys(props).length ? props : null;
}

function styleTextColorHex(index) {
  return brcColorFromIco(index);
}

function scanKnownSprm(grpprl, sprm, operandSize, apply) {
  const lo = sprm & 0xff;
  const hi = (sprm >> 8) & 0xff;
  for (let i = 0; i + 2 + operandSize <= grpprl.length; i += 1) {
    if (grpprl[i] === lo && grpprl[i + 1] === hi) {
      apply(grpprl.subarray(i + 2, i + 2 + operandSize));
    }
  }
}

function extractFontTable(tableStream, fib) {
  if (!fib.lcbFontTable || fib.fcFontTable + fib.lcbFontTable > tableStream.length) return [];

  const data = tableStream.subarray(fib.fcFontTable, fib.fcFontTable + fib.lcbFontTable);
  const fonts = [];
  if (data.length < 4) return fonts;

  const count = data.readUInt16LE(0);
  const extraDataSize = data.readUInt16LE(2);
  let pos = 4;
  for (let i = 0; i < count && pos < data.length; i += 1) {
    const cbFfn = data[pos];
    if (cbFfn === 0 || pos + cbFfn > data.length) break;
    const ffn = data.subarray(pos, pos + cbFfn);
    const weight = ffn.length > 4 ? ffn.readUInt16LE(3) : 0;
    const charset = ffn.length > 6 ? ffn[6] : 0;
    const prq = ffn.length > 7 ? ffn[7] : 0;
    const name = readFfnString(ffn, 40);
    const alternateName = name ? readFfnString(ffn, 40 + (name.length + 1) * 2) : "";
    fonts.push({ index: fonts.length, name, alternateName, weight, charset, prq, extraDataSize });
    pos += cbFfn;
    if (pos % 2 === 1) pos += 1;
  }
  return fonts;
}

function readFfnString(ffn, offset) {
  let name = "";
  for (let pos = offset; pos + 1 < ffn.length; pos += 2) {
    const ch = ffn.readUInt16LE(pos);
    if (ch === 0) break;
    name += String.fromCharCode(ch);
  }
  return name;
}

function extractSections(wordDocument, tableStream, fib) {
  if (!fib.lcbPlcfSed || fib.lcbPlcfSed < 16) return [];
  if (fib.fcPlcfSed + fib.lcbPlcfSed > tableStream.length) return [];

  const plcf = tableStream.subarray(fib.fcPlcfSed, fib.fcPlcfSed + fib.lcbPlcfSed);
  const sedSize = 12;
  const sectionCount = (fib.lcbPlcfSed - 4) / (4 + sedSize);
  if (!Number.isInteger(sectionCount) || sectionCount <= 0) return [];

  const sections = [];
  const sedStart = (sectionCount + 1) * 4;
  for (let i = 0; i < sectionCount; i += 1) {
    const cpStart = plcf.readUInt32LE(i * 4);
    const cpEnd = plcf.readUInt32LE((i + 1) * 4);
    const sedOffset = sedStart + i * sedSize;
    const fcSepx = plcf.readUInt32LE(sedOffset + 2);
    const properties = fcSepx === 0xffffffff
      ? {}
      : readSectionProperties(wordDocument, fcSepx);

    sections.push({
      cpStart,
      cpEnd,
      properties,
    });
  }
  return sections;
}

function readSectionProperties(wordDocument, fcSepx) {
  if (fcSepx + 2 > wordDocument.length) {
    throw new Error("Invalid Word binary document: SEPX points outside WordDocument");
  }

  const cb = wordDocument.readUInt16LE(fcSepx);
  if (fcSepx + 2 + cb > wordDocument.length) {
    throw new Error("Invalid Word binary document: truncated SEPX");
  }

  return parseSectionSprms(wordDocument.subarray(fcSepx + 2, fcSepx + 2 + cb));
}

export function parseSectionSprms(grpprl) {
  const props = {};
  let off = 0;
  while (off + 2 <= grpprl.length) {
    const sprm = grpprl.readUInt16LE(off);
    off += 2;
    let size = sectionSprmOperandSize(sprm);
    if (size === -1) {
      if (off >= grpprl.length) {
        throw new Error(`Truncated section SPRM length byte for 0x${sprm.toString(16)}`);
      }
      size = grpprl[off] + 1;
    }
    if (off + size > grpprl.length) {
      throw new Error(`Truncated section SPRM operand for 0x${sprm.toString(16)}`);
    }
    const val = grpprl.subarray(off, off + size);
    applySectionSprm(props, sprm, val);
    off += size;
  }
  return props;
}

function sectionSprmOperandSize(sprm) {
  if (sprm === 0xd1ff) return 3;
  const spra = (sprm >> 13) & 0x7;
  const size = SPRM_OPERAND_SIZE_BY_SPRA[spra];
  if (size === -1) {
    return -1;
  }
  if (!size || size < 0) {
    throw new Error(`Out-of-spec section SPRM operand size for 0x${sprm.toString(16)}`);
  }
  return size;
}

function applySectionSprm(props, sprm, val) {
  switch (sprm) {
    case 0x3000:
      props.pageNumberChapterSeparator = sectionChapterSeparatorFromCns(val[0]);
      break;
    case 0x3001:
      props.pageNumberChapterStyle = sectionChapterStyleFromHeadingLevel(val[0]);
      break;
    case 0x3009:
      // sprmSBkc — section break code (MS-DOC §2.6.3)
      // 0=continuous, 1=newColumn, 2=newPage, 3=evenPage, 4=oddPage
      props.bkc = val[0];
      break;
    case 0x300a:
      props.titlePg = val[0] !== 0;
      break;
    case 0x3005:
      // MS-DOC-SPEC/16 sprmSFEvenlySpaced: whether section columns are evenly spaced.
      props.columnsEvenlySpaced = val[0] !== 0;
      break;
    case 0x3006:
      // MS-DOC-SPEC/16 sprmSFProtected is inverted: 1 means this section is
      // unprotected when document form protection is enabled.
      props.formProtection = val[0] === 0;
      break;
    case 0x5007:
      // MS-DOC-SPEC/16 sprmSDmBinFirst: paper source for the first page.
      props.paperSourceFirst = val.readUInt16LE(0);
      break;
    case 0x5008:
      // MS-DOC-SPEC/16 sprmSDmBinOther: paper source for non-first pages.
      props.paperSourceOther = val.readUInt16LE(0);
      break;
    case 0x500b:
      // MS-DOC-SPEC/16 sprmSCcolumns stores one less than the section column count.
      props.columnCount = val.readUInt16LE(0) + 1;
      break;
    case 0x900c:
      // MS-DOC-SPEC/16 sprmSDxaColumns: spacing between evenly spaced columns.
      props.columnSpacing = val.readUInt16LE(0);
      break;
    case 0x3019:
      // MS-DOC-SPEC/16 sprmSLBetween: draw lines between section columns.
      props.columnsLineBetween = val[0] !== 0;
      break;
    case 0x301a:
      props.verticalAlign = sectionVerticalAlignFromVjc(val[0]);
      break;
    case 0x3013:
      props.lineNumberRestart = sectionLineNumberRestartFromSlnc(val[0]);
      break;
    case 0x5015: {
      const countBy = val.readUInt16LE(0);
      if (countBy > 100) {
        throw new Error(`Out-of-spec section line-number countBy ${countBy}`);
      }
      props.lineNumberCountBy = countBy;
      break;
    }
    case 0x9016:
      props.lineNumberDistance = val.readUInt16LE(0);
      break;
    case 0x501b: {
      const startMinusOne = val.readUInt16LE(0);
      if (startMinusOne > 32766) {
        throw new Error(`Out-of-spec section line-number start ${startMinusOne + 1}`);
      }
      props.lineNumberStart = startMinusOne + 1;
      break;
    }
    case 0x3228:
      props.sectionBidi = val[0] !== 0;
      break;
    case 0x322a:
      props.rtlGutter = val[0] !== 0;
      break;
    case 0x3011:
      // MS-DOC-SPEC/16 sprmSFPgnRestart gates sprmSPgnStart97/sprmSPgnStart.
      props.pageNumberRestart = val[0] !== 0;
      break;
    case 0x3012:
      // MS-DOC-SPEC/16 sprmSFEndnote: 0 suppresses endnotes for this section.
      props.endnotesSuppressed = val[0] === 0;
      break;
    case 0x303b:
      props.footnotePosition = sectionFootnotePositionFromFpc(val[0]);
      break;
    case 0x303c:
      props.footnoteNumberRestart = sectionNoteNumberRestartFromRnc(val[0], "footnote");
      break;
    case 0x303e:
      props.endnoteNumberRestart = sectionNoteNumberRestartFromRnc(val[0], "endnote");
      break;
    case 0x503f: {
      const start = val.readUInt16LE(0);
      if (start > 16383) {
        throw new Error(`Out-of-spec footnote number start ${start}`);
      }
      props.footnoteNumberStart = start;
      break;
    }
    case 0x5040:
      props.footnoteNumberFormat = val.readUInt16LE(0);
      break;
    case 0x5041: {
      const start = val.readUInt16LE(0);
      if (start > 16383) {
        throw new Error(`Out-of-spec endnote number start ${start}`);
      }
      props.endnoteNumberStart = start;
      break;
    }
    case 0x5042:
      props.endnoteNumberFormat = val.readUInt16LE(0);
      break;
    case 0xf203: {
      // MS-DOC-SPEC/19 SDxaColWidthOperand: iCol, dxaCol width.
      const iCol = val[0];
      const width = val.readUInt16LE(1);
      props.columnWidths ??= [];
      props.columnWidths[iCol] = width;
      break;
    }
    case 0xf204: {
      // MS-DOC-SPEC/19 SDxaColSpacingOperand: iCol, dxaCol spacing.
      const iCol = val[0];
      const spacing = val.readUInt16LE(1);
      props.columnSpacings ??= [];
      props.columnSpacings[iCol] = spacing;
      break;
    }
    case 0x3014:
      // sprmSGprfIhdt — header/footer story flags per section (MS-DOC §2.6.3)
      // Bits: 0x01=evenHeader, 0x02=oddHeader, 0x04=evenFooter, 0x08=oddFooter,
      //       0x10=firstHeader, 0x20=firstFooter
      props.grpfIhdt = val[0];
      break;
    case 0x501c:
      props.pageNumberStart = val.readUInt16LE(0);
      break;
    case 0x7044:
      // MS-DOC-SPEC/16 sprmSPgnStart: 32-bit page number start
      props.pageNumberStart = val.readUInt32LE(0);
      break;
    case 0x301d:
      // MS-DOC-SPEC/16 sprmSBOrientation: 0=portrait, 1=landscape
      props.orientation = val[0] !== 0 ? "landscape" : "portrait";
      break;
    case 0x300e:
      // MS-DOC-SPEC/16 sprmSNfcPgn: MSONFC page-numbering format.
      props.pageNumberFormat = val[0];
      break;
    case 0x702b:
    case 0xd234:
      applyPageBorderSprm(props, "top", val);
      break;
    case 0x702c:
    case 0xd235:
      applyPageBorderSprm(props, "left", val);
      break;
    case 0x702d:
    case 0xd236:
      applyPageBorderSprm(props, "bottom", val);
      break;
    case 0x702e:
    case 0xd237:
      applyPageBorderSprm(props, "right", val);
      break;
    case 0xb01f:
      props.pageWidth = val.readUInt16LE(0);
      break;
    case 0xb020:
      props.pageHeight = val.readUInt16LE(0);
      break;
    case 0xb021:
      props.marginLeft = val.readUInt16LE(0);
      break;
    case 0xb022:
      props.marginRight = val.readUInt16LE(0);
      break;
    case 0x9023:
      props.marginTop = val.readUInt16LE(0);
      break;
    case 0x9024:
      props.marginBottom = val.readUInt16LE(0);
      break;
    case 0xb025:
      // MS-DOC-SPEC/16 sprmSDzaGutter: gutter margin in twips.
      props.gutterMargin = val.readUInt16LE(0);
      break;
    case 0x5026:
      // MS-DOC-SPEC/16 sprmSDmPaperReq is an implementation-specific paper
      // format tie-breaker that consumers MAY ignore; preserve it for callers
      // but do not map it to OOXML section markup.
      props.paperFormatTieBreaker = val.readUInt16LE(0);
      break;
    case 0xb017:
      props.headerMargin = val.readUInt16LE(0);
      break;
    case 0xb018:
      props.footerMargin = val.readUInt16LE(0);
      break;
    case 0x5032:
      props.docGridType = val.readUInt16LE(0);
      break;
    case 0x7030:
      props.docGridCharSpace = val.readInt32LE(0);
      break;
    case 0x9031:
      props.docGridLinePitch = val.readUInt16LE(0);
      break;
    default:
      break;
  }
}

function sectionVerticalAlignFromVjc(value) {
  const map = ["top", "center", "both", "bottom"];
  const align = map[value];
  if (!align) {
    throw new Error(`Out-of-spec section Vjc value ${value}`);
  }
  return align;
}

function sectionChapterSeparatorFromCns(value) {
  // MS-DOC-SPEC/19 CNS: chapter/page number separator enum.
  const map = ["hyphen", "period", "colon", "emDash", "enDash"];
  const separator = map[value];
  if (!separator) {
    throw new Error(`Out-of-spec section chapter separator CNS value ${value}`);
  }
  return separator;
}

function sectionChapterStyleFromHeadingLevel(value) {
  // MS-DOC-SPEC/16 sprmSiHeadingPgn: 0 disables chapter numbers; 1-9 map
  // to Heading 1 through Heading 9.
  if (value > 9) {
    throw new Error(`Out-of-spec section chapter heading level ${value}`);
  }
  return value;
}

function sectionFootnotePositionFromFpc(value) {
  // MS-DOC-SPEC/19 SFpcOperand: 1=bottom of page, 2=beneath text.
  if (value === 1) return "pageBottom";
  if (value === 2) return "beneathText";
  throw new Error(`Out-of-spec section footnote position Fpc value ${value}`);
}

function sectionNoteNumberRestartFromRnc(value, noteType) {
  // MS-DOC-SPEC/19 Rnc: 0=continuous, 1=restart each section,
  // 2=restart each page. Endnotes explicitly disallow rncRstPage.
  if (value === 0) return "continuous";
  if (value === 1) return "eachSect";
  if (value === 2 && noteType === "footnote") return "eachPage";
  throw new Error(`Out-of-spec ${noteType} numbering restart Rnc value ${value}`);
}

function sectionLineNumberRestartFromSlnc(value) {
  const map = ["newPage", "newSection", "continuous"];
  const restart = map[value];
  if (!restart) {
    throw new Error(`Out-of-spec section line-number restart mode ${value}`);
  }
  return restart;
}

function applyPageBorderSprm(props, side, val) {
  // MS-DOC-SPEC/16 defines sprmSBrc*80 as Brc80 and sprmSBrc* as
  // BrcOperand; MS-DOC-SPEC/19 requires BrcOperand.cb to be 8.
  const brc = val.length === 9 && val[0] === 8 ? val.subarray(1) : val;
  if (brc.length !== 4 && brc.length !== 8) {
    throw new Error(`Out-of-spec page border operand size ${val.length}`);
  }
  const isNone = brc.every((byte) => byte === 0) || (brc[0] === 0 && brc[1] === 0);
  if (!isNone) {
    throw new Error(`Unimplemented non-empty section page border parsing on ${side}`);
  }
  props.pageBorders ??= {};
  props.pageBorders[side] = { style: "none" };
}

function extractCharacterRuns(wordDocument, tableStream, fib, bodyText, pieces, styles) {
  if (!fib.lcbChpx || fib.lcbChpx < 4) return [];
  if (fib.fcChpx + fib.lcbChpx > tableStream.length) return [];

  const plcf = tableStream.subarray(fib.fcChpx, fib.fcChpx + fib.lcbChpx);
  const binCount = (fib.lcbChpx - 4) / 8;
  if (binCount <= 0 || !Number.isInteger(binCount)) return [];

  const binFCs = [];
  for (let i = 0; i <= binCount; i += 1) {
    binFCs.push(plcf.readUInt32LE(i * 4));
  }
  const pageNumbers = [];
  for (let bi = 0; bi < binCount; bi += 1) {
    pageNumbers.push(plcf.readUInt32LE((binCount + 1) * 4 + bi * 4));
  }

  const bodyCharacterCount = bodyText.length;
  const runs = [];

  for (let bi = 0; bi < binCount; bi += 1) {
    const pageNumber = pageNumbers[bi];
    if (pageNumber === 0) continue;
    const pageOffset = pageNumber * 512;
    if (pageOffset + 4 > wordDocument.length) continue;

    const page = wordDocument.subarray(pageOffset, Math.min(pageOffset + 512, wordDocument.length));
    const crun = page[511];
    if (crun <= 0) continue;

    const fcBoundaries = readFkpFcBoundaries(page, crun);
    const offsetArrayStart = (crun + 1) * 4;

    for (let i = 0; i < crun; i += 1) {
      const cpStart = fileOffsetToCharacterPosition(fcBoundaries[i], pieces);
      const cpEnd = fileOffsetToCharacterPosition(fcBoundaries[i + 1], pieces);
      if (cpStart == null || cpEnd == null || cpStart >= bodyCharacterCount || cpEnd <= 0) continue;

      const off = page[offsetArrayStart + i] * 2;
      if (off === 0 || off >= page.length) continue;
      const cb = page[off];
      if (cb === 0 || cb > 200 || off + 1 + cb > page.length) continue;
      const grpprl = page.subarray(off + 1, off + 1 + cb);
      const props = parseSprms(grpprl, false);
      // MS-DOC-SPEC/16: parseSprms skips 0x00 padding bytes, which breaks
      // SPRMs whose low byte is 0x00 (e.g. 0x0800 sprmCFRMarkDel). Scan
      // the GRPPRL for these known zero-byte SPRMs as a supplement.
      scanZeroByteSprms(grpprl, props);
      if (props.characterStyleIndex != null) {
        props.styleId = resolveCharacterStyleId(props.characterStyleIndex, styles);
      }
      runs.push({
        cpStart: Math.max(0, cpStart),
        cpEnd: Math.min(bodyCharacterCount, cpEnd),
        properties: props,
      });
    }
  }

  return runs;
}

// MS-DOC-SPEC/16: parseSprms skips 0x00 bytes as padding, preventing
// detection of SPRMs with 0x00 low byte (e.g. 0x0800 sprmCFRMarkDel).
const ZERO_BYTE_SPRM_PATTERNS = [
  { low: 0, high: 0x08, size: 1, name: "revisionMarkDel" },
  { low: 1, high: 0x08, size: 1, name: "revisionMarkIns" },
  { low: 2, high: 0x08, size: 1, name: "fldVanish" },
  { low: 4, high: 0x48, size: 2, name: "revisionAuthorIndex" },
  { low: 5, high: 0x68, size: 4, name: "revisionDate" },
  { low: 7, high: 0x48, size: 2, name: "revisionReason" },
  { low: 0x63, high: 0x48, size: 2, name: "revisionDelAuthorIndex" },
  { low: 0x64, high: 0x68, size: 4, name: "revisionDelDate" },
  { low: 0x67, high: 0x48, size: 2, name: "revisionDelReason" },
];
function scanZeroByteSprms(grpprl, props) {
  if (!grpprl || grpprl.length < 3) return;
  for (let i = 0; i + 2 < grpprl.length; i++) {
    if (grpprl[i] !== 0) continue;
    const low = grpprl[i];
    const high = grpprl[i + 1];
    for (const p of ZERO_BYTE_SPRM_PATTERNS) {
      if (p.low === low && p.high === high) {
        if (i + 2 + p.size > grpprl.length) continue;
        const val = grpprl.subarray(i + 2, i + 2 + p.size);
        if (p.size === 1) props[p.name] = val[0] !== 0;
        else if (p.size === 2) props[p.name] = val.readUInt16LE(0);
        else if (p.size === 4) props[p.name] = val.readUInt32LE(0);
        break;
      }
    }
  }
}

function expandCharacterRuns(characterRuns, bodyCharacterCount) {
  const properties = new Array(bodyCharacterCount).fill(null);
  for (const run of characterRuns) {
    for (let cp = run.cpStart; cp < run.cpEnd && cp < properties.length; cp += 1) {
      properties[cp] = run.properties;
    }
  }
  return properties;
}

function fileOffsetToCharacterPosition(fileOffset, pieces) {
  for (const piece of pieces) {
    const byteLength = (piece.cpEnd - piece.cpStart) * (piece.compressed ? 1 : 2);
    const start = piece.fileOffset;
    const end = start + byteLength;
    if (fileOffset >= start && fileOffset <= end) {
      const delta = fileOffset - start;
      return piece.cpStart + Math.floor(delta / (piece.compressed ? 1 : 2));
    }
  }
  return null;
}

function getParagraphRanges(text) {
  const ranges = [];
  let start = 0;
  for (let i = 0; i < text.length; i += 1) {
    if (text[i] === "\r" || text[i] === "\x07" || text[i] === "\x0c") {
      ranges.push({ cpStart: start, cpEnd: i + 1 });
      start = i + 1;
    }
  }
  if (start <= text.length) {
    ranges.push({ cpStart: start, cpEnd: text.length });
  }
  return ranges;
}

function rangesOverlap(a, b) {
  return a.cpStart < b.cpEnd && b.cpStart < a.cpEnd;
}

function extractTableRows(wordDocument, tableStream, fib, pieces, bodyText, paragraphProperties, sections = null, dataStream = null, styles = []) {
  const paragraphRanges = getParagraphRanges(bodyText);
  const rowProperties = collectTDefTableEntries(wordDocument, tableStream, fib, pieces, dataStream);
  return buildTablesFromInTableParagraphBlocks(bodyText, paragraphProperties, paragraphRanges, rowProperties, sections, styles);
}

function buildTablesFromInTableParagraphBlocks(bodyText, paragraphProperties, paragraphRanges, rowProperties, sections = null, styles = []) {
  if (!paragraphProperties || paragraphProperties.length !== paragraphRanges.length) return [];

  const tables = [];
  let i = 0;
  while (i < paragraphRanges.length) {
    while (i < paragraphRanges.length && !paragraphProperties[i]?.inTable) i += 1;
    if (i >= paragraphRanges.length) break;

    let blockStart = i;
    let blockEnd = i;
    while (blockEnd < paragraphRanges.length) {
      const rangeText = bodyText.substring(paragraphRanges[blockEnd].cpStart, paragraphRanges[blockEnd].cpEnd);
      if (paragraphProperties[blockEnd]?.inTable || rangeText === "\x07") {
        blockEnd += 1;
        continue;
      }
      break;
    }

    const rows = buildGenericTableRows(
      paragraphRanges.slice(blockStart, blockEnd),
      paragraphProperties.slice(blockStart, blockEnd),
      bodyText,
      rowProperties,
    );
    const meaningfulRows = rows.filter((row) =>
      row.rowColumns || row.cells.some((cell) => cleanCellText(cell.text).length > 0),
    );
    if (meaningfulRows.length === 0) {
      i = blockEnd;
      continue;
    }
    if (meaningfulRows.some((row) => !row.rowColumns)) {
      throw new Error("Unable to infer table grid positions: missing parsed row column geometry for a non-empty table block");
    }

    for (const tableRows of splitRowsIntoTables(meaningfulRows)) {
      if (tableRows.length === 0) continue;

      const gridPositions = buildTableGridPositions(tableRows);
      const gridCols = positionsToWidths(gridPositions);
      applyRowGeometry(tableRows, gridPositions, sections);
      const { tableWidth, tableWidthType } = inferTableWidth(tableRows);
      const tableIndent = inferTableIndent(tableRows, gridPositions);
      const tableAutofit = inferTableAutofit(tableRows);
      const cellMargins = inferTableCellMargins(tableRows);
      const tableBorders = inferTableBorders(tableRows);
      const tableBordersExplicit = inferTableBordersExplicit(tableRows);
      const tableStyleId = inferTableStyleId(tableRows, styles);
      tables.push({
        cpStart: tableRows[0].cpStart,
        cpEnd: tableRows.at(-1).cpEnd,
        gridCols,
        gridPositions,
        rows: tableRows,
        tableWidth,
        tableWidthType,
        tableIndent,
        tableAutofit,
        cellMargins,
        tableBorders,
        tableBordersExplicit,
        tableStyleId,
      });
    }

    i = blockEnd;
  }

  if (tables.length === 0 && rowProperties.length > 0) {
    // Some WPS exports do not mark table body paragraphs as inTable, but the row
    // properties still carry explicit table geometry. Use them as the fallback source.
    const rows = buildGenericTableRows(paragraphRanges, paragraphProperties, bodyText, rowProperties);
    const meaningfulRows = rows.filter((row) =>
      row.rowColumns || row.cells.some((cell) => cleanCellText(cell.text).length > 0),
    );
    if (meaningfulRows.some((row) => !row.rowColumns)) {
      throw new Error("Unable to infer table grid positions: missing parsed row column geometry for a non-empty table block");
    }
    for (const tableRows of splitRowsIntoTables(meaningfulRows)) {
      if (tableRows.length === 0) continue;

      const gridPositions = buildTableGridPositions(tableRows);
      const gridCols = positionsToWidths(gridPositions);
      applyRowGeometry(tableRows, gridPositions, sections);
      const { tableWidth, tableWidthType } = inferTableWidth(tableRows);
      const tableIndent = inferTableIndent(tableRows, gridPositions);
      const tableAutofit = inferTableAutofit(tableRows);
      const cellMargins = inferTableCellMargins(tableRows);
      const tableBorders = inferTableBorders(tableRows);
      const tableBordersExplicit = inferTableBordersExplicit(tableRows);
      const tableStyleId = inferTableStyleId(tableRows, styles);
      tables.push({
        cpStart: tableRows[0].cpStart,
        cpEnd: tableRows.at(-1).cpEnd,
        gridCols,
        gridPositions,
        rows: tableRows,
        tableWidth,
        tableWidthType,
        tableIndent,
        tableAutofit,
        cellMargins,
        tableBorders,
        tableBordersExplicit,
        tableStyleId,
      });
    }
  }

  return tables;
}

function splitRowsIntoTables(rows) {
  if (rows.length === 0) return [];
  const tables = [];
  let current = [];
  let currentWidthKey = null;
  let currentPositionKey = null;
  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    const widthKey = row.tableWidth == null
      ? null
      : `${row.tableWidth}:${row.tableWidthType ?? "dxa"}`;
    const indentKey = row.tableIndent ? `${row.tableIndent.width}:${row.tableIndent.type}` : "none";
    const positionKey = row.tablePosition ? "positioned" : row.tableNoAllowOverlap ? "positioned" : "flow";
    if (
      current.length > 0 &&
      ((widthKey != null && currentWidthKey != null && widthKey !== currentWidthKey) ||
        (currentPositionKey != null && positionKey !== currentPositionKey) ||
        current.some((currentRow) => (currentRow.tableIndent ? `${currentRow.tableIndent.width}:${currentRow.tableIndent.type}` : "none") !== indentKey))
    ) {
      tables.push(current);
      current = [];
      currentWidthKey = null;
      currentPositionKey = null;
    }
    if (currentWidthKey == null && widthKey != null) {
      currentWidthKey = widthKey;
    }
    if (currentPositionKey == null) {
      currentPositionKey = positionKey;
    }
    current.push(row);
  }
  if (current.length > 0) tables.push(current);
  return tables;
}

function buildGenericTableRows(paragraphRanges, paragraphProperties, bodyText, rowProperties) {
  const rows = [];
  const rowEndIndices = new Set();

  for (let i = 0; i < paragraphRanges.length; i += 1) {
    const rangeText = bodyText.substring(paragraphRanges[i].cpStart, paragraphRanges[i].cpEnd);
    const isSeparator = !paragraphProperties[i]?.inTable && rangeText === "\x07";
    if (isSeparator) rowEndIndices.add(i);
  }

  const blockStartCp = paragraphRanges[0]?.cpStart ?? 0;
  const blockEndCp = paragraphRanges.at(-1)?.cpEnd ?? 0;
  for (const rowProperty of rowProperties ?? []) {
    if (rowProperty.cpStart < blockStartCp || rowProperty.cpEnd > blockEndCp) continue;
    const rowEndIndex = paragraphRanges.findIndex((range) => range.cpStart === rowProperty.cpStart && range.cpEnd === rowProperty.cpEnd);
    if (rowEndIndex >= 0) rowEndIndices.add(rowEndIndex);
  }

  let rowStart = 0;
  for (const rowEndIndex of [...rowEndIndices].sort((a, b) => a - b)) {
    if (rowEndIndex < rowStart) continue;

    const row = buildGenericTableRow(
      paragraphRanges.slice(rowStart, rowEndIndex),
      paragraphProperties.slice(rowStart, rowEndIndex),
      bodyText,
      paragraphRanges[rowEndIndex],
      rowProperties,
    );
    if (row) rows.push(row);
    rowStart = rowEndIndex + 1;
  }

  const trailingRow = buildGenericTableRow(
    paragraphRanges.slice(rowStart),
    paragraphProperties.slice(rowStart),
    bodyText,
    null,
    rowProperties,
  );
  if (trailingRow) rows.push(trailingRow);

  return rows;
}

function buildGenericTableRow(paragraphRanges, paragraphProperties, bodyText, rowEndRange, rowProperties) {
  if (paragraphRanges.length === 0) return null;

  const cells = [];
  let cellStart = paragraphRanges[0].cpStart;
  let sawCell = false;

  for (const range of paragraphRanges) {
    const rangeText = bodyText.substring(range.cpStart, range.cpEnd);
    if (!rangeText.endsWith("\x07")) continue;

    const cellEnd = range.cpEnd;
      cells.push({
        cpStart: cellStart,
        cpEnd: cellEnd,
        text: cleanCellText(bodyText.substring(cellStart, cellEnd)),
        width: 0,
        gridSpan: 1,
        vMerge: null,
        vAlign: "center",
      });
    cellStart = cellEnd;
    sawCell = true;
  }

  if (!sawCell) return null;

  const cpStart = paragraphRanges[0].cpStart;
  const cpEnd = rowEndRange?.cpEnd ?? paragraphRanges.at(-1).cpEnd;
  const rowProperty = findRowPropertyForRange(rowProperties, cpStart, cpEnd, cells.length);
  const rowColumns = normalizeColumnPositions(rowProperty?.columns);
  const rowPosition = extractTablePositionForRange(paragraphRanges, paragraphProperties, cpStart, cpEnd);

  return {
    cpStart,
    cpEnd,
    cells,
    rowColumns,
    cellMargins: rowProperty?.cellMargins ?? null,
    cellFlags: rowProperty?.cellFlags ?? null,
    cellBorders: rowProperty?.cellBorders ?? null,
    cellShading: rowProperty?.cellShading ?? null,
    cellBorderSideArrays: rowProperty?.cellBorderSideArrays ?? null,
    cellBorderAssignments: rowProperty?.cellBorderAssignments ?? null,
    tableStyleIndex: rowProperty?.tableStyleIndex ?? null,
    tableBorders: rowProperty?.tableBorders ?? null,
    tableBordersExplicit: rowProperty?.tableBordersExplicit === true,
    tableWidth: rowProperty?.tableWidth ?? null,
    tableWidthType: rowProperty?.tableWidthType ?? null,
    tableIndent: rowProperty?.tableIndent ?? null,
    tableJustification: rowProperty?.tableJustification ?? null,
    tableAutofit: rowProperty?.tableAutofit ?? null,
    rowHeight: rowProperty?.rowHeight ?? null,
    rowHeightRule: rowProperty?.rowHeightRule ?? null,
    cantSplit: rowProperty?.cantSplit ?? null,
    repeatHeader: rowProperty?.repeatHeader ?? null,
    vMergeAssignments: rowProperty?.vMergeAssignments ?? null,
    vAlignAssignments: rowProperty?.vAlignAssignments ?? null,
    tablePosition: rowPosition?.tablePosition ?? null,
    tableNoAllowOverlap: rowPosition?.tableNoAllowOverlap ?? null,
  };
}

function buildTableGridPositions(rows) {
  const positionSet = new Set();
  for (const row of rows) {
    if (!row.rowColumns || row.rowColumns.length < 2) continue;
    for (const pos of row.rowColumns) positionSet.add(pos);
  }

  if (positionSet.size === 0) {
    throw new Error("Unable to infer table grid positions: no row column geometry was parsed");
  }

  const positions = [...positionSet].sort((a, b) => a - b);
  const merged = [positions[0]];
  for (let i = 1; i < positions.length; i += 1) {
    if (Math.abs(positions[i] - merged.at(-1)) > 20) {
      merged.push(positions[i]);
    }
  }
  return merged;
}

function positionsToWidths(positions) {
  if (!positions || positions.length < 2) {
    throw new Error("Invalid table grid: expected at least two positions");
  }
  const widths = [];
  for (let i = 1; i < positions.length; i += 1) {
    widths.push(Math.max(1, positions[i] - positions[i - 1]));
  }
  return widths;
}

function applyRowGeometry(rows, gridPositions, sections = null) {
  for (const row of rows) {
    if (!row.rowColumns || row.rowColumns.length !== row.cells.length + 1) {
      row.rowColumns = inferRowColumnsFromGrid(gridPositions, row);
    }
    if (!row.rowColumns || row.rowColumns.length !== row.cells.length + 1) {
      throw new Error("Unable to infer table row geometry: missing parsed row column boundaries");
    }

    for (let ci = 0; ci < row.cells.length; ci += 1) {
      const start = row.rowColumns[ci];
      const end = row.rowColumns[ci + 1];
      const startIndex = findGridPositionIndex(gridPositions, start);
      const endIndex = findGridPositionIndex(gridPositions, end);
      const span = startIndex != null && endIndex != null && endIndex > startIndex
        ? endIndex - startIndex
        : 1;
      row.cells[ci].width = Math.max(1, end - start);
      row.cells[ci].gridSpan = span;
      if (row.cellBorders?.[ci]) {
        row.cells[ci].borders = row.cellBorders[ci];
      }
      if (row.cellShading?.[ci]) {
        row.cells[ci].shading = row.cellShading[ci];
      }
    }
    applyCellFlags(row);
    applyCellBorderSideArrays(row);
    applyCellBorderAssignments(row);
    applyVerticalMergeAssignments(row);
    applyVerticalAlignAssignments(row);
  }
}

function applyCellFlags(row) {
  if (!row.cellFlags || row.cellFlags.length === 0) return;
  for (let ci = 0; ci < Math.min(row.cells.length, row.cellFlags.length); ci += 1) {
    const flags = row.cellFlags[ci];
    const cell = row.cells[ci];
    if (!cell || !flags) continue;
    if (flags.bVertRestart) cell.vMerge = "restart";
    else if (flags.bVertMerge) cell.vMerge = "continue";
    cell.noWrap = flags.fNoWrap === true;
    if (flags.nVertAlign === 2) cell.vAlign = "bottom";
    else if (flags.nVertAlign === 1) cell.vAlign = "center";
    else if (flags.nVertAlign === 0) cell.vAlign = "top";
  }
}

function applyVerticalMergeAssignments(row) {
  if (!row.vMergeAssignments || row.vMergeAssignments.length === 0) return;
  for (const assignment of row.vMergeAssignments) {
    const cell = row.cells[assignment.itc];
    if (!cell) continue;
    if (assignment.state === 3) cell.vMerge = "restart";
    else if (assignment.state === 1) cell.vMerge = "continue";
  }
}

function applyCellBorderAssignments(row) {
  if (!row.cellBorderAssignments || row.cellBorderAssignments.length === 0) return;
  for (const assignment of row.cellBorderAssignments) {
    const start = Math.max(0, assignment.itcFirst);
    const end = Math.min(row.cells.length, assignment.itcLim);
    for (let ci = start; ci < end; ci += 1) {
      const cell = row.cells[ci];
      if (!cell) continue;
      if (!cell.borders) {
        cell.borders = createDefaultTableBorders();
      }
      if (assignment.changeTop) cell.borders.top = assignment.border;
      if (assignment.changeLeft) cell.borders.left = assignment.border;
      if (assignment.changeBottom) cell.borders.bottom = assignment.border;
      if (assignment.changeRight) cell.borders.right = assignment.border;
    }
  }
}

function applyCellBorderSideArrays(row) {
  if (!row.cellBorderSideArrays) return;

  const sideArrays = [
    ["top", row.cellBorderSideArrays.top],
    ["left", row.cellBorderSideArrays.left],
    ["bottom", row.cellBorderSideArrays.bottom],
    ["right", row.cellBorderSideArrays.right],
  ];

  if (sideArrays.every(([, borders]) => !borders || borders.length === 0)) return;

  for (let ci = 0; ci < row.cells.length; ci += 1) {
    const cell = row.cells[ci];
    if (!cell) continue;
    for (const [side, borders] of sideArrays) {
      if (!borders || borders.length === 0) continue;
      const border = borders[ci];
      if (!border || border.style === "none") continue;
      if (!cell.borders) cell.borders = createDefaultTableBorders();
      cell.borders[side] = border;
    }
  }
}

function applyVerticalAlignAssignments(row) {
  if (!row.vAlignAssignments || row.vAlignAssignments.length === 0) return;
  for (const assignment of row.vAlignAssignments) {
    const start = Math.max(0, assignment.itcFirst);
    const end = Math.min(row.cells.length, assignment.itcLim);
    for (let ci = start; ci < end; ci += 1) {
      const cell = row.cells[ci];
      if (!cell) continue;
      cell.vAlign = assignment.valign === 2 ? "bottom" : assignment.valign === 1 ? "center" : "top";
    }
  }
}

function extractTablePositionForRange(paragraphRanges, paragraphProperties, cpStart, cpEnd) {
  const positions = [];
  let noAllowOverlap = false;

  for (let i = 0; i < paragraphRanges.length; i += 1) {
    const range = paragraphRanges[i];
    if (range.cpEnd <= cpStart || range.cpStart >= cpEnd) continue;
    const props = paragraphProperties[i];
    if (props?.tablePosition) positions.push(props.tablePosition);
    if (props?.tableNoAllowOverlap) noAllowOverlap = true;
  }

  if (positions.length === 0 && !noAllowOverlap) return null;

  const merged = {};
  for (const position of positions) {
    for (const [key, value] of Object.entries(position)) {
      if (value == null) continue;
      if (merged[key] == null) {
        merged[key] = value;
        continue;
      }
      if (merged[key] !== value) {
        throw new Error(`Conflicting table position sprms were parsed for a table row: ${key}`);
      }
    }
  }

  return { tablePosition: Object.keys(merged).length > 0 ? merged : null, tableNoAllowOverlap: noAllowOverlap };
}

function parseCellBorderSideArrayOperand(sprm, payload) {
  if (!payload || payload.length === 0) {
    throw new Error(`Truncated cell border side array SPRM 0x${sprm.toString(16)}`);
  }
  if (payload.length % 4 !== 0) {
    throw new Error(`Out-of-spec cell border side array length ${payload.length} in SPRM 0x${sprm.toString(16)}`);
  }

  let side = null;
  if (sprm === 0xD61A) side = "top";
  else if (sprm === 0xD61B) side = "left";
  else if (sprm === 0xD61C) side = "bottom";
  else if (sprm === 0xD61D) side = "right";
  else throw new Error(`Internal parser error: cell border side array called for SPRM 0x${sprm.toString(16)}`);

  const borders = [];
  for (let off = 0; off < payload.length; off += 4) {
    borders.push(parseBorderRecord(payload.subarray(off, off + 4), sprm));
  }
  return { side, borders };
}

function inferRowColumnsFromGrid(gridPositions, row) {
  if (!gridPositions || gridPositions.length < row.cells.length + 1) return null;

  const excessColumns = gridPositions.length - 1 - row.cells.length;
  if (excessColumns < 0) return null;

  const rowColumns = [gridPositions[0]];
  for (let i = excessColumns; i < gridPositions.length - 1; i += 1) {
    rowColumns.push(gridPositions[i]);
  }

  if (rowColumns.length !== row.cells.length + 1) return null;
  for (let i = 1; i < rowColumns.length; i += 1) {
    if (rowColumns[i] <= rowColumns[i - 1]) return null;
  }

  return rowColumns;
}

function findGridPositionIndex(gridPositions, value) {
  let bestIndex = null;
  let bestDistance = Infinity;
  for (let i = 0; i < gridPositions.length; i += 1) {
    const distance = Math.abs(gridPositions[i] - value);
    if (distance < bestDistance) {
      bestDistance = distance;
      bestIndex = i;
    }
  }
  return bestDistance <= 20 ? bestIndex : null;
}

function findRowPropertyForRange(rowProperties, cpStart, cpEnd, cellCount) {
  if (!rowProperties || rowProperties.length === 0) return null;

  const overlapping = rowProperties.filter((entry) => rangesOverlap(entry, { cpStart, cpEnd }));
  if (overlapping.length === 0) return null;

  const exactCellCount = overlapping.findLast((entry) => entry.columns?.length === cellCount + 1);
  if (exactCellCount) return exactCellCount;

  return overlapping.at(-1) ?? null;
}

function normalizeColumnPositions(columns) {
  if (!columns || columns.length < 2) return null;
  for (let i = 1; i < columns.length; i += 1) {
    if (columns[i] <= columns[i - 1]) return null;
  }
  return columns;
}

function collectTDefTableEntries(wordDocument, tableStream, fib, pieces, dataStream = null) {
  if (!fib.fcPapx || !fib.lcbPapx || fib.lcbPapx < 4) return [];
  if (fib.fcPapx + fib.lcbPapx > tableStream.length) return [];

  const plcf = tableStream.subarray(fib.fcPapx, fib.fcPapx + fib.lcbPapx);
  const binCount = (fib.lcbPapx - 4) / 8;
  if (binCount <= 0 || !Number.isInteger(binCount)) return [];

  const pageNumbers = [];
  for (let bi = 0; bi < binCount; bi += 1) {
    pageNumbers.push(plcf.readUInt32LE((binCount + 1) * 4 + bi * 4));
  }

  const entries = [];

  for (let bi = 0; bi < binCount; bi += 1) {
    const pageNumber = pageNumbers[bi];
    if (pageNumber === 0) continue;
    const pageOffset = pageNumber * 512;
    if (pageOffset + PHE_SIZE > wordDocument.length) continue;

    const page = wordDocument.subarray(pageOffset, Math.min(pageOffset + 512, wordDocument.length));
    const crun = page[511];
    if (crun <= 0) continue;

    const fcBoundaries = readFkpFcBoundaries(page, crun);
    const bxStart = (crun + 1) * 4;

    for (let i = 0; i < crun; i += 1) {
      const bxOffset = bxStart + i * (1 + PHE_SIZE);
      if (bxOffset >= 511) break;
      const papxOffset = page[bxOffset] * 2;
      if (papxOffset <= 0 || papxOffset >= 511) continue;
      const cb = page[papxOffset];
      let dataOffset;
      let byteLength;
      if (cb === 0) {
        const cbx = page[papxOffset + 1];
        if (cbx === 0 || papxOffset + 2 + cbx * 2 > page.length) continue;
        dataOffset = papxOffset + 2;
        byteLength = cbx * 2;
      } else {
        byteLength = cb * 2;
        if (papxOffset + 1 + byteLength > page.length) continue;
        dataOffset = papxOffset + 1;
      }

      const cpStart = fileOffsetToCharacterPosition(fcBoundaries[i], pieces);
      const cpEnd = fileOffsetToCharacterPosition(fcBoundaries[i + 1], pieces);
      if (cpStart == null || cpEnd == null) continue;

      const data = page.subarray(dataOffset, dataOffset + byteLength);
      const info = parseTableRowSprms(data, dataStream);
      if (!info) continue;

      entries.push({
        cpStart,
        cpEnd,
        columns: info.columns,
        cellFlags: info.cellFlags,
        cellBorders: info.cellBorders,
        cellShading: info.cellShading,
        cellBorderSideArrays: info.cellBorderSideArrays,
        cellBorderAssignments: info.cellBorderAssignments,
        tableWidth: info.tableWidth,
        tableWidthType: info.tableWidthType,
        tableIndent: info.tableIndent,
        tableJustification: info.tableJustification,
        tableAutofit: info.tableAutofit,
        rowHeight: info.rowHeight,
        rowHeightRule: info.rowHeightRule,
        cellMargins: info.cellMargins,
        tableBorders: info.tableBorders,
        tableBordersExplicit: info.tableBordersExplicit,
        vMergeAssignments: info.vMergeAssignments,
        vAlignAssignments: info.vAlignAssignments,
        cantSplit: info.cantSplit,
        repeatHeader: info.repeatHeader,
      });
    }
  }

  entries.sort((a, b) => a.cpStart - b.cpStart);
  return entries;
}

function cleanCellText(text) {
  return text
    .replace(/[\x07\x0c]/g, "")
    .replace(/[\x00-\x06\x08\x0b\x0e-\x1f]/g, "");
}

function isControlOnlyText(text) {
  return text.replace(/[\s\x00-\x1f]/g, "").length === 0;
}

function hasVisibleText(text) {
  return text.replace(/[\s\x00-\x1f]/g, "").length > 0;
}

function parseTableRowSprms(data, dataStream = null) {
  data = expandHugePapxTableSprms(data, dataStream);
  let off = 2;
  if (data.length >= 2 && data[off] === 0) off += 1;

  let tableDef = null;
  let cellFlags = [];
  let tableWidth = null;
  let tableWidthType = null;
  let tableIndent = null;
  let tableJustification = null;
  let tableStyleIndex = null;
  let tableAutofit = null;
  let rowHeight = null;
  let rowHeightRule = null;
  let cellMargins = null;
  let tableBorders = null;
  let tableBordersExplicit = false;
  let cellBorders = [];
  let cellShading = [];
  let cellBorderAssignments = [];
  let cellBorderSideArrays = {};
  const cellMarginCandidates = [];
  const vMergeAssignments = [];
  const vAlignAssignments = [];
  let cantSplit = null;
  let repeatHeader = null;

  while (off + 2 <= data.length) {
    if (data[off] === 0) { off += 1; continue; }
    const sprm = data.readUInt16LE(off);
    const spra = (sprm >> 13) & 0x7;
    let size = SPRM_OPERAND_SIZE_BY_SPRA[spra];

    if (sprm === 0xD608 || sprm === 0xD606) {
      if (off + 4 > data.length) {
        throw new Error(`Truncated table definition SPRM length word for 0x${sprm.toString(16)}`);
      }
      const cb = data.readUInt16LE(off + 2);
      size = cb + 1;
      if (off + 3 + cb <= data.length) {
        const tblData = data.subarray(off + 3, off + 3 + cb);
        const itc = tblData[1];
        if (itc > 0 && itc < 64 && 2 + (itc + 1) * 2 <= tblData.length) {
          const positions = [];
          for (let i = 0; i <= itc; i += 1) {
            positions.push(tblData.readInt16LE(2 + i * 2));
          }
          tableDef = positions;
          const descriptorStart = 2 + (itc + 1) * 2;
          const availableDescriptorBytes = tblData.length - descriptorStart;
          const descriptorCount = Math.min(itc, Math.floor(availableDescriptorBytes / 20));
          if (descriptorCount > 0) {
            cellFlags = [];
            cellBorders = [];
            for (let i = 0; i < descriptorCount; i += 1) {
              const base = descriptorStart + i * 20;
              const bits = tblData.readUInt16LE(base);
              cellFlags.push({
                bFirstMerged: (bits & 0x0001) !== 0,
                bMerged: (bits & 0x0002) !== 0,
                bVertical: (bits & 0x0004) !== 0,
                bBackward: (bits & 0x0008) !== 0,
                bRotateFont: (bits & 0x0010) !== 0,
                bVertMerge: (bits & 0x0020) !== 0,
                bVertRestart: (bits & 0x0040) !== 0,
                nVertAlign: (bits & 0x0180) >> 7,
                // MS-DOC-SPEC/19 TCGRF.fNoWrap is bit 13 of the TC80 flag word.
                fNoWrap: (bits & 0x2000) !== 0,
              });
              cellBorders.push(parseCellBordersFromDescriptor(tblData.subarray(base + 4, base + 20)));
            }
          }
        }
      }
    } else if (sprm === 0xF614) {
      if (off + 5 <= data.length) {
        const widthType = data[off + 2];
        const widthValue = data.readUInt16LE(off + 3);
        if (widthType === 3) {
          tableWidth = widthValue;
          tableWidthType = "dxa";
        } else if (widthType === 2) {
          tableWidth = widthValue;
          tableWidthType = "pct";
        } else if (widthType !== 1 && widthType !== 0) {
          throw new Error(`Out-of-spec table width type ${widthType} in sprmTTableWidth`);
        }
      }
    } else if (sprm === 0xF661) {
      if (off + 5 <= data.length) {
        tableIndent = parseTableIndent(data.subarray(off + 2, off + 5));
      }
    } else if (sprm === 0x548A || sprm === 0x5400) {
      if (off + 4 <= data.length) {
        tableJustification = parseTableJustification(data.readUInt16LE(off + 2), sprm);
      }
    } else if (sprm === 0x563A) {
      if (off + 4 <= data.length) {
        // MS-DOC-SPEC/16 sprmTIstd: istd of the table style to apply.
        tableStyleIndex = data.readUInt16LE(off + 2);
      }
    } else if (sprm === 0x3615) {
      if (off + 3 <= data.length) {
        tableAutofit = data[off + 2] !== 0;
      }
    } else if (sprm === 0xD605) {
      const cb = data[off + 2];
      size = cb + 1;
      if (off + 3 + cb <= data.length) {
        tableBorders = parseTableBordersOperand(sprm, data.subarray(off + 3, off + 3 + cb));
        tableBordersExplicit = true;
      }
    } else if (sprm === 0xD613) {
      const cb = data[off + 2];
      size = cb + 1;
      if (off + 3 + cb <= data.length) {
        tableBorders = parseTableBordersOperand(sprm, data.subarray(off + 3, off + 3 + cb));
        tableBordersExplicit = true;
      }
    } else if (sprm === 0xD620 || sprm === 0xD62F) {
      const cb = data[off + 2];
      size = cb + 1;
      if (off + 3 + cb <= data.length) {
        const payload = data.subarray(off + 3, off + 3 + cb);
        const assignment = parseCellBorderAssignmentOperand(sprm, payload);
        if (assignment) cellBorderAssignments.push(assignment);
      }
    } else if (sprm === 0xD612 || sprm === 0xD616 || sprm === 0xD60C || sprm === 0xD670 || sprm === 0xD671 || sprm === 0xD672) {
      const cb = data[off + 2];
      size = cb + 1;
      if (off + 3 + cb <= data.length) {
        // MS-DOC-SPEC/19 DefTableShdOperand applies 10-byte Shd entries
        // sequentially to cells 1-22, 23-44, or 45-63 depending on the SPRM.
        const startIndex = (sprm === 0xD616 || sprm === 0xD671) ? 22 : (sprm === 0xD60C || sprm === 0xD672) ? 44 : 0;
        const shading = parseDefTableShdOperand(data.subarray(off + 3, off + 3 + cb));
        for (let i = 0; i < shading.length; i += 1) {
          cellShading[startIndex + i] = shading[i];
        }
      }
    } else if ((sprm >= 0xD61A && sprm <= 0xD61D) || sprm === 0xD662 || (sprm >= 0xD680 && sprm <= 0xD686)) {
      const cb = data[off + 2];
      size = cb + 1;
      if (off + 3 + cb <= data.length && sprm >= 0xD61A && sprm <= 0xD61D) {
        const payload = data.subarray(off + 3, off + 3 + cb);
        const assignment = parseCellBorderSideArrayOperand(sprm, payload);
        if (assignment) cellBorderSideArrays[assignment.side] = assignment.borders;
      }
    } else if (sprm === 0xD632 || sprm === 0xD634) {
      const cb = data[off + 2];
      size = cb + 1;
      if (off + 3 + cb <= data.length) {
        const candidate = parseTableCellPadding(data.subarray(off + 3, off + 3 + cb), sprm === 0xD634);
        if (candidate) cellMarginCandidates.push(candidate);
      }
    } else if (sprm === 0x9407) {
      if (off + 4 <= data.length) {
        const height = data.readInt16LE(off + 2);
        rowHeight = Math.abs(height);
        rowHeightRule = height < 0 ? 1 : 0;
      }
    } else if (sprm === 0x3403 || sprm === 0x3466) {
      if (off + 3 <= data.length) {
        cantSplit = data[off + 2] !== 0;
      }
    } else if (sprm === 0x3404) {
      if (off + 3 <= data.length) {
        repeatHeader = data[off + 2] !== 0;
      }
    } else if (sprm === 0xD62B) {
      const cb = data[off + 2];
      size = cb + 1;
      if (off + 5 <= data.length) {
        const itc = data[off + 3];
        const state = data[off + 4] & 0x3;
        vMergeAssignments.push({ itc, state });
      }
    } else if (sprm === 0xD62C) {
      const cb = data[off + 2];
      size = cb + 1;
      if (off + 6 <= data.length) {
        const itcFirst = data[off + 3];
        const itcLim = data[off + 4];
        const valign = data[off + 5] & 0x3;
        vAlignAssignments.push({ itcFirst, itcLim, valign });
      }
    } else {
      if (size === -1) {
        if (off + 2 < data.length) size = data[off + 2] + 1;
        else break;
      }
    }

    if (!size || size < 0 || off + 2 + size > data.length) break;
    off += 2 + size;
  }

  if (!tableDef) return null;
  cellMargins = chooseTableCellMargins(cellMarginCandidates, tableDef.length - 1);
  return {
    columns: tableDef,
    cellFlags,
    cellBorders,
    cellShading,
    cellBorderSideArrays,
    cellBorderAssignments,
    tableWidth,
    tableWidthType,
    tableIndent,
    tableJustification,
    tableStyleIndex,
    tableAutofit,
    rowHeight,
    rowHeightRule,
    cellMargins,
    tableBorders: tableBorders ?? createDefaultTableBorders(),
    tableBordersExplicit,
    vMergeAssignments,
    vAlignAssignments,
    cantSplit,
    repeatHeader,
  };
}

function parseDefTableShdOperand(payload) {
  if (payload.length % 10 !== 0) {
    throw new Error(`Invalid DefTableShdOperand length ${payload.length}`);
  }
  const shades = [];
  for (let off = 0; off < payload.length; off += 10) {
    shades.push(parseTableShd(payload.subarray(off, off + 10)));
  }
  return shades;
}

function parseTableShd(shd) {
  const cvFore = shd.readUInt32LE(0);
  const cvBack = shd.readUInt32LE(4);
  const ipat = shd.readUInt16LE(8);
  if (ipat === 0xffff) return null;
  return {
    val: ipat === 0 ? "clear" : "clear",
    color: tableColorRefToHex(cvFore, "auto"),
    fill: tableColorRefToHex(cvBack, "auto"),
  };
}

function tableColorRefToHex(value, autoValue) {
  if (value === 0xff000000 || value === 0xffffffff) return autoValue;
  const red = value & 0xff;
  const green = (value >> 8) & 0xff;
  const blue = (value >> 16) & 0xff;
  return [red, green, blue].map((part) => part.toString(16).toUpperCase().padStart(2, "0")).join("");
}

function parseTableIndent(data) {
  if (!data || data.length < 3) {
    throw new Error("Truncated sprmTWidthIndent operand");
  }
  const widthType = data[0];
  const widthValue = data.readInt16LE(1);
  if (widthType === 0 || widthType === 1) {
    if (widthValue !== 0) {
      throw new Error(`Invalid sprmTWidthIndent width ${widthValue} for type ${widthType}`);
    }
    return null;
  }
  // MS-DOC-SPEC/19 FtsWWidth_Indent permits ftsDxa (0x03) and forbids
  // ftsPercent/ftsDxaSys; wWidth is signed because table indents may be negative.
  if (widthType === 3) {
    return { width: widthValue, type: "dxa" };
  }
  throw new Error(`Out-of-spec sprmTWidthIndent width type ${widthType}`);
}

function parseTableJustification(value, sprm) {
  // MS-DOC-SPEC/16 sprmTJc/sprmTJc90: table justification is an unsigned
  // 16-bit value; the documented default is logical/physical left.
  if (value === 0) return "left";
  if (value === 1) return "center";
  if (value === 2) return "right";
  throw new Error(`Out-of-spec table justification ${value} in SPRM 0x${sprm.toString(16)}`);
}

function inferTableWidth(rows) {
  let tableWidth = null;
  let tableWidthType = null;
  for (const row of rows) {
    if (row.tableWidth == null) continue;
    if (tableWidth == null) {
      tableWidth = row.tableWidth;
      tableWidthType = row.tableWidthType ?? "dxa";
      continue;
    }
    if (tableWidth !== row.tableWidth || tableWidthType !== (row.tableWidthType ?? "dxa")) {
      throw new Error("Conflicting table width sprms were parsed for a single table");
    }
  }
  return { tableWidth, tableWidthType };
}

function inferTableIndent(rows, gridPositions = null) {
  let indent = null;
  for (const row of rows) {
    if (!row.tableIndent) continue;
    if (!indent) {
      indent = row.tableIndent;
      continue;
    }
    if (indent.width !== row.tableIndent.width || indent.type !== row.tableIndent.type) {
      throw new Error("Conflicting table indent sprms were parsed for a single table");
    }
  }

  // MS-DOC-SPEC/19 TDefTableOperand: rgdxaCenter[0] is the logical-left
  // table edge, "indented from the logical left page margin"; preserve that
  // parsed edge as OOXML table indentation for flow tables.
  const geometryIndent = gridPositions?.[0] ? { width: gridPositions[0], type: "dxa" } : null;
  if (indent && indent.width !== 0 && geometryIndent && (indent.width !== geometryIndent.width || indent.type !== geometryIndent.type)) {
    throw new Error("Conflicting table indent and table geometry were parsed for a single table");
  }
  return geometryIndent ?? (indent?.width !== 0 ? indent : null);
}

function inferTableAutofit(rows) {
  let autofit = null;
  for (const row of rows) {
    if (row.tableAutofit == null) continue;
    if (autofit == null) {
      autofit = row.tableAutofit;
      continue;
    }
    if (autofit !== row.tableAutofit) {
      throw new Error("Conflicting table autofit sprms were parsed for a single table");
    }
  }
  return autofit;
}

function parseTableCellPadding(data, isDefault) {
  if (!data || data.length < 6) return null;
  const startCell = data[0];
  const endCell = isDefault ? 1 : data[1];
  const sideBits = data[2];
  const sizeType = isDefault ? 0x3 : data[3];
  const value = data.readUInt16LE(4);
  if (isDefault && startCell !== 0) return null;
  if (!isDefault && (startCell >= endCell || sizeType !== 0x3)) return null;
  const margins = { top: 0, left: 0, bottom: 0, right: 0 };
  let used = false;
  if (sideBits & 0x01) { margins.top = value; used = true; }
  if (sideBits & 0x02) { margins.left = value; used = true; }
  if (sideBits & 0x04) { margins.bottom = value; used = true; }
  if (sideBits & 0x08) { margins.right = value; used = true; }
  return used ? { margins, startCell, endCell, isDefault, sideBits } : null;
}

function chooseTableCellMargins(candidates, cellCount) {
  if (!candidates || candidates.length === 0) return null;
  const selected = { top: null, left: null, bottom: null, right: null };
  for (const candidate of candidates) {
    const coversAllCells = candidate.isDefault || (candidate.startCell === 0 && candidate.endCell >= cellCount);
    if (!coversAllCells) continue;
    for (const side of ["top", "left", "bottom", "right"]) {
      const bit = side === "top" ? 0x01 : side === "left" ? 0x02 : side === "bottom" ? 0x04 : 0x08;
      if (!(candidate.sideBits & bit)) continue;
      if (selected[side] == null) {
        selected[side] = candidate.margins[side];
        continue;
      }
      if (selected[side] !== candidate.margins[side]) {
        throw new Error("Conflicting table cell padding sprms were parsed for a single table");
      }
    }
  }
  return {
    top: selected.top ?? 0,
    left: selected.left ?? 0,
    bottom: selected.bottom ?? 0,
    right: selected.right ?? 0,
  };
}

function inferTableCellMargins(rows) {
  let margins = null;
  for (const row of rows) {
    if (!row.cellMargins) continue;
    if (!margins) {
      margins = row.cellMargins;
      continue;
    }
    if (
      margins.top !== row.cellMargins.top ||
      margins.left !== row.cellMargins.left ||
      margins.bottom !== row.cellMargins.bottom ||
      margins.right !== row.cellMargins.right
    ) {
      throw new Error("Conflicting table cell margins were parsed for a single table");
    }
  }
  return margins;
}

function inferTableBorders(rows) {
  let borders = null;
  for (const row of rows) {
    if (!row.tableBorders) continue;
    if (!borders) {
      borders = row.tableBorders;
      continue;
    }
    if (JSON.stringify(borders) !== JSON.stringify(row.tableBorders)) {
      throw new Error("Conflicting table border sprms were parsed for a single table");
    }
  }
  return borders ?? createDefaultTableBorders();
}

function inferTableBordersExplicit(rows) {
  return rows.some((row) => row.tableBordersExplicit === true);
}

function inferTableStyleId(rows, styles = []) {
  const styleIndexes = rows
    .map((row) => row.tableStyleIndex)
    .filter((styleIndex) => styleIndex != null);
  if (styleIndexes.length === 0) return null;
  const first = styleIndexes[0];
  if (!styleIndexes.every((styleIndex) => styleIndex === first)) {
    throw new Error("Conflicting table style sprms were parsed for a single table");
  }
  const style = styles[first];
  if (!style || style.type !== STYLE_TYPE_TABLE || !style.styleId) {
    throw new Error(`Invalid table style istd ${first} parsed from sprmTIstd`);
  }
  return style.styleId;
}

function createDefaultTableBorders() {
  const borders = {};
  for (const side of TABLE_BORDER_SIDES) {
    borders[side] = createEmptyTableBorder();
  }
  return borders;
}

function parseCellBordersFromDescriptor(raw) {
  if (!raw || raw.length < 16) {
    return null;
  }

  const borders = {};
  const cellBorderSides = ["top", "left", "bottom", "right"];
  for (let i = 0; i < cellBorderSides.length; i += 1) {
    borders[cellBorderSides[i]] = parseWw8CellBorderEntry(raw.subarray(i * 4, (i + 1) * 4));
  }
  return borders;
}

function parseWw8CellBorderEntry(raw) {
  if (raw.length !== 4) {
    throw new Error(`Invalid WW8 cell border entry length ${raw.length}`);
  }
  if (raw.every((byte) => byte === 0xff)) {
    return createNilTableBorder();
  }
  return normalizeTableBorderRecord({
    sprm: 0xd608,
    brcType: raw[1],
    dptLineWidth: raw[0],
    colorIndex: raw[2],
    space: raw[3] & 0x1f,
  });
}

function createEmptyTableBorder() {
  return {
    style: "none",
    width: 0,
    color: null,
    space: 0,
  };
}

function createNilTableBorder() {
  return {
    ...createEmptyTableBorder(),
    // MS-DOC-SPEC/19 Brc80MayBeNil and BrcMayBeNil use all-bits-set
    // NilBrc values to specify an explicit no-border region.
    nil: true,
  };
}

function parseTableBordersOperand(sprm, payload) {
  if (!payload || payload.length === 0) {
    throw new Error(`Truncated table border SPRM 0x${sprm.toString(16)}`);
  }

  const entrySize = payload.length / TABLE_BORDER_SIDES.length;
  if (!Number.isInteger(entrySize)) {
    throw new Error(`Out-of-spec table border SPRM length ${payload.length} for 0x${sprm.toString(16)}`);
  }

  if (entrySize !== 2 && entrySize !== 4 && entrySize !== 8) {
    throw new Error(`Unimplemented table border entry size ${entrySize} in SPRM 0x${sprm.toString(16)}`);
  }

  const borders = {};
  for (let i = 0; i < TABLE_BORDER_SIDES.length; i += 1) {
    borders[TABLE_BORDER_SIDES[i]] = parseTableBorderEntry(
      payload.subarray(i * entrySize, (i + 1) * entrySize),
      entrySize,
      sprm,
    );
  }
  return borders;
}

function parseCellBorderAssignmentOperand(sprm, payload) {
  if (!payload || payload.length < 3) {
    throw new Error(`Truncated table cell border assignment SPRM 0x${sprm.toString(16)}`);
  }

  const itcFirst = payload[0];
  const itcLim = payload[1];
  const nFlag = payload[2];
  if (itcLim <= itcFirst) {
    return null;
  }

  const borderRaw = payload.subarray(3);
  const border = parseBorderRecord(borderRaw, sprm);
  return {
    itcFirst,
    itcLim,
    changeTop: (nFlag & 0x01) !== 0,
    changeLeft: (nFlag & 0x02) !== 0,
    changeBottom: (nFlag & 0x04) !== 0,
    changeRight: (nFlag & 0x08) !== 0,
    border,
  };
}

function parseBorderRecord(raw, sprm) {
  if (!raw || raw.length === 0) {
    throw new Error(`Invalid border record length 0 in SPRM 0x${sprm.toString(16)}`);
  }
  if (raw.length === 4 && raw[0] === 0xff && raw[1] === 0xff) {
    return createNilTableBorder();
  }
  if (raw.length === 8 && raw.every((byte) => byte === 0xff)) {
    return createNilTableBorder();
  }
  if (raw.length === 2 && raw[0] === 0xff && raw[1] === 0xff) {
    return createNilTableBorder();
  }
  if (raw.length !== 2 && raw.length !== 4 && raw.length !== 8) {
    throw new Error(`Unimplemented border record length ${raw.length} in SPRM 0x${sprm.toString(16)}`);
  }
  return parseTableBorderEntry(raw, raw.length, sprm);
}

function parseTableBorderEntry(raw, entrySize, sprm) {
  if (entrySize === 2) {
    const bits = raw.readUInt16LE(0);
    let dptLineWidth = bits & 0x07;
    let brcType = (bits & 0x18) >> 3;
    const colorIndex = ((bits & 0xc0) >> 6) | ((raw[1] & 0x07) << 2);
    const space = raw[1] >> 3;
    if (dptLineWidth > 5) {
      brcType = dptLineWidth;
      dptLineWidth = 1;
    }
    dptLineWidth *= 6;
    return normalizeTableBorderRecord({
      sprm,
      brcType,
      dptLineWidth,
      colorIndex,
      space,
    });
  }

  if (entrySize === 4) {
    if (raw.every((byte) => byte === 0xff)) {
      return createNilTableBorder();
    }
    return normalizeTableBorderRecord({
      sprm,
      brcType: raw[1],
      dptLineWidth: raw[0],
      colorIndex: raw[2],
      space: raw[3] & 0x1f,
    });
  }

  if (entrySize === 8) {
    if (raw.subarray(4, 8).every((byte) => byte === 0xff)) {
      return createNilTableBorder();
    }
    return normalizeTableBorderRecord({
      sprm,
      brcType: raw[5],
      dptLineWidth: raw[4],
      color: raw.readUInt32LE(0),
      space: raw[6] & 0x1f,
    });
  }

  throw new Error(`Internal parser error: table border entry size ${entrySize} in SPRM 0x${sprm.toString(16)}`);
}

function normalizeTableBorderRecord({ sprm, brcType, dptLineWidth, colorIndex = null, color = null, space = 0 }) {
  return {
    style: tableBorderStyleFromBrcType(brcType, sprm),
    width: dptLineWidth,
    color: color != null ? colorToHexFromBgr(color) : colorIndexToHex(colorIndex),
    space,
  };
}

function tableBorderStyleFromBrcType(brcType, sprm) {
  const borderName = BRC_TYPE_NAMES[brcType];
  if (borderName) return borderName;
  throw new Error(`Out-of-spec table border BrcType ${brcType} in SPRM 0x${sprm.toString(16)}`);
}

function colorIndexToHex(colorIndex) {
  // MS-DOC-SPEC/19 Brc80.ico uses the Ico enumeration; 0 is automatic color.
  if (colorIndex == null || colorIndex === 0) {
    return null;
  }
  const color = brcColorFromIco(colorIndex);
  return color === "auto" ? null : color;
}

function colorToHexFromBgr(color) {
  // MS-DOC-SPEC/19 COLORREF with fAuto 0xFF is cvAuto.
  if (color === 0xff000000) {
    return null;
  }
  const blue = color & 0xff;
  const green = (color >> 8) & 0xff;
  const red = (color >> 16) & 0xff;
  return `${red.toString(16).padStart(2, "0")}${green.toString(16).padStart(2, "0")}${blue.toString(16).padStart(2, "0")}`.toUpperCase();
}

function expandHugePapxTableSprms(data, dataStream, depth = 0) {
  if (depth > 8) {
    throw new Error("Huge PAPX sprm expansion exceeded maximum recursion depth");
  }

  if (data.length <= 2) {
    return Buffer.from(data);
  }

  const expanded = [...data.subarray(0, 2)];
  let off = 2;
  if (data[off] === 0) {
    expanded.push(0);
    off += 1;
  }
  while (off + 2 <= data.length) {
    if (data[off] === 0) {
      expanded.push(0);
      off += 1;
      continue;
    }

    const sprm = data.readUInt16LE(off);
    if ((sprm === 0x6646 || sprm === 0x646B) && off + 6 <= data.length) {
      if (!dataStream) {
        throw new Error(`Encountered huge table sprm 0x${sprm.toString(16)} without Data stream`);
      }
      const dataOffset = data.readUInt32LE(off + 2);
      const nested = readDataStreamGrpprl(dataStream, dataOffset);
      const nestedExpanded = expandHugePapxTableSprms(nested, dataStream, depth + 1);
      expanded.push(...nestedExpanded);
      off += 6;
      continue;
    }

    let size = SPRM_OPERAND_SIZE_BY_SPRA[(sprm >> 13) & 0x7];
    if (sprm === 0xD608 || sprm === 0xD606) {
      if (off + 4 > data.length) {
        throw new Error(`Truncated table definition SPRM length word for 0x${sprm.toString(16)}`);
      }
      const cb = data.readUInt16LE(off + 2);
      size = cb + 1;
    } else if (sprm === 0xD605) {
      if (off + 3 > data.length) break;
      const cb = data[off + 2];
      size = cb + 1;
    } else if (size === -1) {
      if (off + 3 > data.length) break;
      size = data[off + 2] + 1;
    }

    if (!size || size < 0 || off + 2 + size > data.length) {
      throw new Error(`Truncated table sprm data at offset ${off}`);
    }

    expanded.push(...data.subarray(off, off + 2 + size));
    off += 2 + size;
  }

  return Buffer.from(expanded);
}

function readDataStreamGrpprl(dataStream, offset) {
  if (!dataStream) {
    throw new Error("Missing Data stream for huge PAPX table sprm expansion");
  }
  if (offset + 2 > dataStream.length) {
    throw new Error(`Huge PAPX data offset ${offset} is outside the Data stream`);
  }

  const length = dataStream.readUInt16LE(offset);
  const start = offset + 2;
  const end = start + length;
  if (end > dataStream.length) {
    throw new Error(`Huge PAPX data at offset ${offset} exceeds the Data stream`);
  }

  return dataStream.subarray(start, end);
}

// --- List/numbering parsing (LSTF, LVL, LFO, LFOData) ---

// Number format codes per MS-DOC-SPEC MSONFC
const NFC_TO_NUMFMT = {
  0: "decimal", 1: "upperRoman", 2: "lowerRoman", 3: "upperLetter",
  4: "lowerLetter", 5: "ordinal", 6: "cardinalText", 7: "ordinalText",
  14: "decimal", 23: "bullet", 255: "none",
};
function nfcToNumFmt(nfc) { return NFC_TO_NUMFMT[nfc] ?? "decimal"; }

// ixchFollow: what follows the number
const IXCH_FOLLOW = { 0: "tab", 1: "space", 2: "nothing" };

function extractListData(tableStream, fib) {
  const result = { lstfList: [], lfoList: [] };

  // Parse PlfLst (list format templates)
  if (fib.lcbPlcfLst >= 2 && fib.fcPlcfLst + fib.lcbPlcfLst <= tableStream.length) {
    const plfLst = tableStream.subarray(fib.fcPlcfLst, fib.fcPlcfLst + fib.lcbPlcfLst);
    const cLst = plfLst.readUInt16LE(0);
    let off = 2;
    for (let i = 0; i < cLst && off + 28 <= plfLst.length; i++) {
      const lstf = {
        lsid: plfLst.readUInt32LE(off),
        tplc: plfLst.readUInt32LE(off + 4),
        rgistdPara: [],
        fSimpleList: !!(plfLst[off + 26] & 1),
        fHybrid: !!(plfLst[off + 26] & 16),
        lvlList: [],
      };
      for (let j = 0; j < 9; j++) {
        lstf.rgistdPara.push(plfLst.readUInt16LE(off + 8 + j * 2));
      }
      off += 28;
      result.lstfList.push(lstf);
    }

    // LVL arrays follow the PlfLst in the stream immediately after the LSTF entries
    let lvlOff = fib.fcPlcfLst + off; // off = 2 + cLst*28 from LSTF loop
    for (const lstf of result.lstfList) {
      const numLvls = lstf.fSimpleList ? 1 : 9;
      for (let i = 0; i < numLvls; i++) {
        if (lvlOff + 20 > tableStream.length) break;
        const iStartAt = tableStream.readUInt32LE(lvlOff);
        const nfc = tableStream[lvlOff + 4];
        const flags = tableStream[lvlOff + 5];
        const jc = flags & 3;
        const fLegal = !!(flags & 0x04);
        const fNoRestart = !!(flags & 0x08);
        const fTentative = !!(flags & 0x80);
        const rgbxchNums = [];
        for (let j = 0; j < 9; j++) rgbxchNums.push(tableStream[lvlOff + 6 + j]);
        const ixchFollow = tableStream[lvlOff + 15];
        // skip dxaIndentSav(4) + unused2(4) = 8 bytes at offset 16
        const cbGrpprlChpx = tableStream[lvlOff + 24];
        const cbGrpprlPapx = tableStream[lvlOff + 25];
        const ilvlRestartLim = tableStream[lvlOff + 26];
        const grfhic = tableStream[lvlOff + 27];
        lvlOff += 28; // fixed LVLF header size (28 bytes per MS-DOC-SPEC §LVLF)

        // Read grpprlPapx then grpprlChpx (LVL stores papx first, then chpx)
        const grpprlPapx = tableStream.subarray(lvlOff, lvlOff + cbGrpprlPapx);
        lvlOff += cbGrpprlPapx;
        const grpprlChpx = tableStream.subarray(lvlOff, lvlOff + cbGrpprlChpx);
        lvlOff += cbGrpprlChpx;

        // Read xst (level text): Xst structure per MS-DOC-SPEC §Xst.
        // cch (2 bytes, unsigned) = character count, followed by cch × 2 bytes of rgtchar.
        // Xst is NOT null-terminated (unlike Xstz).
        // cch=0 means empty string.
        const cch = lvlOff + 2 <= tableStream.length ? tableStream.readUInt16LE(lvlOff) : 0;
        const xstDataLen = Math.min(cch * 2, tableStream.length - lvlOff - 2);
        const xst = cch > 0 && xstDataLen > 0
          ? tableStream.subarray(lvlOff + 2, lvlOff + 2 + xstDataLen).toString("utf16le")
          : "";
        lvlOff += 2 + cch * 2; // skip cch + rgtchar

        const lvl = {
          ilvl: i, iStartAt, nfc, jc, fLegal, fNoRestart, fTentative, ilvlRestartLim, grfhic,
          ixchFollow, rgbxchNums, xst,
          grpprlPapx, grpprlChpx,
        };
        // Parse paragraph properties from grpprlPapx for indents.
        // Some LVLs have a phantom SPRM at the end (0x8460) whose operand is missing;
        // try with the full size first, then with reduced size if truncated.
        if (cbGrpprlPapx > 0) {
          let papxProps = null;
          try { papxProps = parseSprms(grpprlPapx, false); } catch(e) { /* ignore */ }
          // Retry with 2 fewer bytes if the last SPRM was truncated
          if (!papxProps && cbGrpprlPapx > 2) {
            try { papxProps = parseSprms(grpprlPapx.subarray(0, cbGrpprlPapx - 2), false); } catch(e2) { /* ignore */ }
          }
          if (papxProps) {
            lvl.leftIndent = papxProps.leftIndent ?? null;
            lvl.firstLineIndent = papxProps.firstLineIndent ?? null;
            lvl.rightIndent = papxProps.rightIndent ?? null;
            lvl.tabs = papxProps.tabs ?? null;
            lvl.papxAlignment = papxProps.alignment ?? null;
          }
        }
        // Parse character properties from grpprlChpx for font and rPr
        if (cbGrpprlChpx > 0) {
          try {
            const chpxProps = parseSprms(grpprlChpx, false);
            lvl.fontAscii = chpxProps.fontAscii ?? null;
            lvl.fontHAnsi = chpxProps.fontHAnsi ?? null;
            lvl.fontEastAsia = chpxProps.fontEastAsia ?? null;
            lvl.fontCs = chpxProps.fontCs ?? null;
            lvl.fontSizeCs = chpxProps.fontSizeCs ?? null;
            lvl.fontHint = chpxProps.fontHint ?? null;
	            lvl.textColor = chpxProps.textColor ?? null;
	            lvl.bold = chpxProps.bold ?? null;
	            lvl.boldCs = chpxProps.boldCs ?? null;
	            lvl.italic = chpxProps.italic ?? null;
	            lvl.italicCs = chpxProps.italicCs ?? null;
	            lvl.caps = chpxProps.caps ?? null;
	            lvl.smallCaps = chpxProps.smallCaps ?? null;
	            lvl.strike = chpxProps.strike ?? null;
	            lvl.dstrike = chpxProps.dstrike ?? null;
	            lvl.vanish = chpxProps.vanish ?? null;
	            lvl.outline = chpxProps.outline ?? null;
	            lvl.shadow = chpxProps.shadow ?? null;
	            lvl.imprint = chpxProps.imprint ?? null;
	            lvl.emboss = chpxProps.emboss ?? null;
	            lvl.noProof = chpxProps.noProof ?? null;
	            lvl.webHidden = chpxProps.webHidden ?? null;
	            lvl.specVanish = chpxProps.specVanish ?? null;
	            lvl.charSpacing = chpxProps.charSpacing ?? null;
	            lvl.charWidth = chpxProps.charWidth ?? null;
	            lvl.fontSize = chpxProps.fontSize ?? null;
	            lvl.kern = chpxProps.kern ?? null;
	            lvl.textPosition = chpxProps.textPosition ?? null;
	            lvl.verticalAlign = chpxProps.verticalAlign ?? null;
	            lvl.underline = chpxProps.underline ?? null;
	            lvl.underlineStyle = chpxProps.underlineStyle ?? null;
	            lvl.underlineColor = chpxProps.underlineColor ?? null;
	            lvl.border = chpxProps.border ?? null;
	          } catch(e) { /* ignore */ }
        }
        lstf.lvlList.push(lvl);
      }
    }
  }

  // Parse PlfLfo (list format overrides)
  if (fib.lcbPlfLfo >= 4 && fib.fcPlfLfo + fib.lcbPlfLfo <= tableStream.length) {
    const plfLfo = tableStream.subarray(fib.fcPlfLfo, fib.fcPlfLfo + fib.lcbPlfLfo);
    const lfoMac = plfLfo.readUInt32LE(0);
    let off = 4;
    for (let i = 0; i < lfoMac && off + 16 <= plfLfo.length; i++) {
      const lsid = plfLfo.readUInt32LE(off);
      const clfolvl = plfLfo[off + 12];
      result.lfoList.push({ lsid, clfolvl, index: i, lfolvlList: [] });
      off += 16;
    }

    // MS-DOC-SPEC/19 PlfLfo stores LFOData immediately after rgLfo, in
    // parallel order with rgLfo. LFOLVL entries can override level starts even
    // when they do not carry a full LVL formatting override.
    for (const lfo of result.lfoList) {
      if (off + 4 > plfLfo.length) {
        throw new Error("Invalid Word binary document: truncated LFOData");
      }
      const cp = plfLfo.readUInt32LE(off);
      off += 4;
      lfo.cp = cp;
      for (let j = 0; j < lfo.clfolvl; j += 1) {
        if (off + 8 > plfLfo.length) {
          throw new Error("Invalid Word binary document: truncated LFOLVL");
        }
        const iStartAt = plfLfo.readInt32LE(off);
        const flags = plfLfo.readUInt32LE(off + 4);
        const iLvl = flags & 0x0f;
        const fStartAt = !!(flags & 0x10);
        const fFormatting = !!(flags & 0x20);
        if (iLvl > 0x08) {
          throw new Error(`Out-of-spec LFOLVL level ${iLvl}`);
        }
        const lfolvl = { iLvl, fStartAt, fFormatting };
        if (fStartAt) {
          if (iStartAt < 0 || iStartAt > 0x7fff) {
            throw new Error(`Out-of-spec LFOLVL start-at value ${iStartAt}`);
          }
          lfolvl.iStartAt = iStartAt;
        }
        off += 8;
        if (fFormatting) {
          throw new Error("Unimplemented LFOLVL full LVL formatting override");
        }
        lfo.lfolvlList.push(lfolvl);
      }
    }
  }

  return result;
}
