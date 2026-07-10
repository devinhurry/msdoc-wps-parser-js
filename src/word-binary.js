import { parseSprms, scanSprmRecords } from "./sprm.js";
import { BRC_TYPE_NAMES, brcColorFromIco, colorRefToHex } from "./sprm.js";
import { inflateRawSync } from "node:zlib";
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
const FIB_FC_PLCFFNDREF_INDEX = 2;
const FIB_FC_PLCFFNDTXT_INDEX = 3;
const FIB_FC_PLCFANDREF_INDEX = 4;
const FIB_FC_PLCFANDTXT_INDEX = 5;
const FIB_FC_PAPX_INDEX = 13;
const FIB_FC_STSH_INDEX = 1;
const FIB_FC_FONT_TABLE_INDEX = 15;

const PHE_SIZE = 12;
const SPRM_WPS_DYA_LINE = 0x6412;
const SPRM_WPS_DYA_LINE_OPERAND_SIZE = 2;
const FIB_FC_CHPX_INDEX = 12;
const FIB_FC_PLCFFLDMOM_INDEX = 16;
const FIB_FC_PLCFFLDHDR_INDEX = 17;
const FIB_FC_PLCFFLDFTN_INDEX = 18;
const FIB_FC_PLCFFLDATN_INDEX = 19;
const FIB_FC_PLCFFLDEDN_INDEX = 48;
const FIB_FC_PLCFFLDTXBX_INDEX = 57;
const FIB_FC_PLCFFLDHDRTXBX_INDEX = 59;
const PNFPN_MASK = 0x003fffff;
const FKP_PAGE_SIZE = 512;
const FIB_FC_PLCFLST_INDEX = 73;
const FIB_FC_PLCFLFO_INDEX = 74;
const FIB_FC_STTBFBKMK_INDEX = 21;
const FIB_FC_PLCFBKF_INDEX = 22;
const FIB_FC_PLCFBKL_INDEX = 23;
const FIB_FC_WSS_INDEX = 30;
const FIB_FC_DOP_INDEX = 31; // 0x1F per FibRgFcLcb97
const FIB_FC_PLCFSPAMOM_INDEX = 40;
const FIB_FC_PLCFSPAHDR_INDEX = 41;
const FIB_FC_DGGINFO_INDEX = 50;
const FIB_FC_GRPXSTATNOWNERS_INDEX = 36;
const FIB_FC_STTBFATNBKMK_INDEX = 37;
const FIB_FC_PLCFATNBKF_INDEX = 42;
const FIB_FC_PLCFATNBKL_INDEX = 43;
const FIB_FC_PLCFENDREF_INDEX = 46;
const FIB_FC_PLCFENDTXT_INDEX = 47;
const FIB_FC_PLCFTXBXTXT_INDEX = 56;
const FIB_FC_PLCFHDRTXBXTXT_INDEX = 58;
const FIB_FC_STTBFRMARK_INDEX = 51; // MS-DOC-SPEC/15 FibRgFcLcb97 fcSttbfRMark at offset 0x232
const SELSF_SIZE = 36;
const OFFICE_ART_DGG_CONTAINER = 0xf000;
const OFFICE_ART_DG_CONTAINER = 0xf002;
const OFFICE_ART_SPGR_CONTAINER = 0xf003;
const OFFICE_ART_SP_CONTAINER = 0xf004;
const OFFICE_ART_FSP = 0xf00a;
const OFFICE_ART_FOPT = 0xf00b;
const OFFICE_ART_BSE = 0xf007;
const OFFICE_ART_BLIP_EMF = 0xf01a;
const OFFICE_ART_BLIP_WMF = 0xf01b;
const OFFICE_ART_BLIP_PICT = 0xf01c;
const OFFICE_ART_BLIP_JPEG = 0xf01d;
const OFFICE_ART_BLIP_PNG = 0xf01e;
const OFFICE_ART_BLIP_DIB = 0xf01f;
const OFFICE_ART_BLIP_TIFF = 0xf029;
const OFFICE_ART_CLIENT_DATA = 0xf012;
const OFFICE_ART_TERTIARY_FOPT = 0xf122;
const OFFICE_ART_SHAPE_NAME_PID = 0x0380;
const OFFICE_ART_SHAPE_DESCRIPTION_PID = 0x0381;
const OFFICE_ART_GFXDATA_PID = 0x03a9;
const WPS_DRAWING_RELATIVE_HEIGHT_BASE = 0x0f000000;
const WPS_DRAWING_RELATIVE_HEIGHT_STEP = 0x400;

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
  const tableStream = fib.whichTableStream === "1Table" ? table1 : table0;
  if (!tableStream) {
    throw new Error(`Missing required Word table stream: ${fib.whichTableStream}`);
  }

  const pieces = readDocumentPieces(wordDocument, tableStream, fib);
  const subdocuments = splitSubdocuments(wordDocument, pieces, fib.characterCounts);
  const rawText = readPieces(wordDocument, pieces);
  const bodyText = subdocuments.body.rawText;
  const dop = parseDop(tableStream, fib);
  const { styles, latentLsd, stiMaxWhenSaved, styleSheetInfo } = extractStyleSheet(tableStream, fib);
  const fontTable = extractFontTable(tableStream, fib);
  if ((styleSheetInfo?.nVerBuiltInNamesWhenSaved ?? 0) >= 7) {
    for (const font of fontTable) {
      if (font?.alternateName && /^[\x20-\x7e]+$/.test(font.alternateName) && /[^\x00-\x7f]/.test(font.name ?? "")) {
        // MS-DOC-SPEC/19 Stshif.nVerBuiltInNamesWhenSaved identifies the
        // built-in style-name generation. In newer-generation WPS exports,
        // ASCII FFN.xszAlt values are used as OOXML face names for localized
        // fonts, while older generations retain xszFfn as the face name.
        font.preferredName = font.alternateName;
      }
    }
  }
  const sections = extractSections(wordDocument, tableStream, fib, bodyText);
  const listData = extractListData(tableStream, fib);
  const defaultTabStop = dop.dxaTab;
  const plcfHdd = parsePlcfHdd(tableStream, fib, subdocuments.headers.rawText, sections.length);
  const fields = parseFieldTables(tableStream, fib, subdocuments);
  const footnotes = parseNoteCollection(tableStream, fib.fcPlcffndRef, fib.lcbPlcffndRef, fib.fcPlcffndTxt, fib.lcbPlcffndTxt, bodyText, subdocuments.footnotes.rawText, "footnote");
  const endnotes = parseNoteCollection(tableStream, fib.fcPlcfendRef, fib.lcbPlcfendRef, fib.fcPlcfendTxt, fib.lcbPlcfendTxt, bodyText, subdocuments.endnotes.rawText, "endnote");
  const comments = parseComments(tableStream, fib, bodyText, subdocuments.annotations.rawText);
  const textboxes = {
    body: parseTextboxTextPlc(tableStream, fib.fcPlcftxbxTxt, fib.lcbPlcftxbxTxt, subdocuments.textboxes.rawText, "Textbox Document"),
    headers: parseTextboxTextPlc(tableStream, fib.fcPlcfHdrtxbxTxt, fib.lcbPlcfHdrtxbxTxt, subdocuments.headerTextboxes.rawText, "Header Textbox Document"),
  };
  const bookmarks = parseStandardBookmarks(tableStream, fib);
  const officeArtDrawings = parseOfficeArtContent(tableStream, fib);
  const shapeAnchors = attachOfficeArtToShapeAnchors(
    parseShapeAnchors(tableStream, fib.fcPlcSpaMom, fib.lcbPlcSpaMom, fib.characterCounts.body, "PlcfSpaMom"),
    officeArtDrawings.body,
  );
  const headerShapeAnchors = attachOfficeArtToShapeAnchors(
    parseShapeAnchors(tableStream, fib.fcPlcSpaHdr, fib.lcbPlcSpaHdr, fib.characterCounts.headers, "PlcfSpaHdr"),
    officeArtDrawings.headers,
  );
  const revisionAuthors = parseRevisionAuthors(tableStream, fib);
  const lastSelection = parseLastSelection(tableStream, fib, rawText.length);
  const allParagraphPropertyEntries = extractParagraphPropertyEntries(wordDocument, tableStream, fib, rawText.length, styles, pieces);
  const subdocumentParagraphProperties = Object.fromEntries(
    Object.entries(subdocuments).map(([name, story]) => [name, paragraphPropertiesForStory(story, allParagraphPropertyEntries)]),
  );
  const paragraphProperties = subdocumentParagraphProperties.body;
  const allCharacterRuns = extractCharacterRuns(wordDocument, tableStream, fib, rawText, pieces, styles);
  const allCharacterProperties = expandCharacterRuns(allCharacterRuns, rawText.length);
  const pictures = parseInlinePictures(data, subdocuments, allCharacterProperties);
  const characterRuns = allCharacterRuns
    .filter((run) => run.cpStart < bodyText.length && run.cpEnd > 0)
    .map((run) => ({ ...run, cpStart: Math.max(0, run.cpStart), cpEnd: Math.min(bodyText.length, run.cpEnd) }));
  const characterProperties = allCharacterProperties.slice(0, bodyText.length);
  const subdocumentCharacterProperties = Object.fromEntries(
    Object.entries(subdocuments).map(([name, story]) => [name, allCharacterProperties.slice(story.cpStart, story.cpEnd)]),
  );
  const tableRows = extractTableRows(wordDocument, tableStream, fib, pieces, bodyText, paragraphProperties, sections, data, styles);

  return {
    fib,
    pieces,
    text: normalizeWordText(bodyText),
    rawText,
    bodyText,
    paragraphs: paragraphsFromWordText(bodyText),
    paragraphProperties,
    subdocumentParagraphProperties,
    characterProperties,
    characterRuns,
    allCharacterRuns,
    subdocumentCharacterProperties,
    styles,
    latentLsd,
    stiMaxWhenSaved,
    styleSheetInfo,
    fontTable,
    sections,
    listData,
    defaultTabStop,
    subdocuments,
    tableRows,
    plcfHdd,
    fields,
    footnotes,
    endnotes,
    comments,
    textboxes,
    bookmarks,
    shapeAnchors,
    headerShapeAnchors,
    officeArtDrawings,
    pictures,
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
  const plcfSpaMomOffset = FIB_FC_LCB_START + FIB_FC_PLCFSPAMOM_INDEX * 8;
  const plcfSpaHdrOffset = FIB_FC_LCB_START + FIB_FC_PLCFSPAHDR_INDEX * 8;
  const dggInfoOffset = FIB_FC_LCB_START + FIB_FC_DGGINFO_INDEX * 8;
  const sttbfRMarkOffset = FIB_FC_LCB_START + FIB_FC_STTBFRMARK_INDEX * 8;
  const fcLcbLimit = FIB_FC_LCB_START + fcLcbCount * 4;
  const readOptionalFcLcb = (index) => {
    const offset = FIB_FC_LCB_START + index * 8;
    return offset + 8 <= fcLcbLimit
      ? { fc: wordDocument.readUInt32LE(offset), lcb: wordDocument.readUInt32LE(offset + 4) }
      : { fc: 0, lcb: 0 };
  };
  const fieldMom = readOptionalFcLcb(FIB_FC_PLCFFLDMOM_INDEX);
  const fieldHdr = readOptionalFcLcb(FIB_FC_PLCFFLDHDR_INDEX);
  const fieldFtn = readOptionalFcLcb(FIB_FC_PLCFFLDFTN_INDEX);
  const fieldAtn = readOptionalFcLcb(FIB_FC_PLCFFLDATN_INDEX);
  const fieldEdn = readOptionalFcLcb(FIB_FC_PLCFFLDEDN_INDEX);
  const fieldTxbx = readOptionalFcLcb(FIB_FC_PLCFFLDTXBX_INDEX);
  const fieldHdrTxbx = readOptionalFcLcb(FIB_FC_PLCFFLDHDRTXBX_INDEX);
  const footnoteRef = readOptionalFcLcb(FIB_FC_PLCFFNDREF_INDEX);
  const footnoteText = readOptionalFcLcb(FIB_FC_PLCFFNDTXT_INDEX);
  const annotationRef = readOptionalFcLcb(FIB_FC_PLCFANDREF_INDEX);
  const annotationText = readOptionalFcLcb(FIB_FC_PLCFANDTXT_INDEX);
  const annotationOwners = readOptionalFcLcb(FIB_FC_GRPXSTATNOWNERS_INDEX);
  const annotationBookmarkNames = readOptionalFcLcb(FIB_FC_STTBFATNBKMK_INDEX);
  const annotationBookmarkStarts = readOptionalFcLcb(FIB_FC_PLCFATNBKF_INDEX);
  const annotationBookmarkEnds = readOptionalFcLcb(FIB_FC_PLCFATNBKL_INDEX);
  const endnoteRef = readOptionalFcLcb(FIB_FC_PLCFENDREF_INDEX);
  const endnoteText = readOptionalFcLcb(FIB_FC_PLCFENDTXT_INDEX);
  const textboxText = readOptionalFcLcb(FIB_FC_PLCFTXBXTXT_INDEX);
  const headerTextboxText = readOptionalFcLcb(FIB_FC_PLCFHDRTXBXTXT_INDEX);
  if (tableStreamOffset + 8 > fcLcbLimit) {
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
    fcPlcffndRef: footnoteRef.fc,
    lcbPlcffndRef: footnoteRef.lcb,
    fcPlcffndTxt: footnoteText.fc,
    lcbPlcffndTxt: footnoteText.lcb,
    fcPlcfandRef: annotationRef.fc,
    lcbPlcfandRef: annotationRef.lcb,
    fcPlcfandTxt: annotationText.fc,
    lcbPlcfandTxt: annotationText.lcb,
    fcGrpXstAtnOwners: annotationOwners.fc,
    lcbGrpXstAtnOwners: annotationOwners.lcb,
    fcSttbfAtnBkmk: annotationBookmarkNames.fc,
    lcbSttbfAtnBkmk: annotationBookmarkNames.lcb,
    fcPlcfAtnBkf: annotationBookmarkStarts.fc,
    lcbPlcfAtnBkf: annotationBookmarkStarts.lcb,
    fcPlcfAtnBkl: annotationBookmarkEnds.fc,
    lcbPlcfAtnBkl: annotationBookmarkEnds.lcb,
    fcPlcfendRef: endnoteRef.fc,
    lcbPlcfendRef: endnoteRef.lcb,
    fcPlcfendTxt: endnoteText.fc,
    lcbPlcfendTxt: endnoteText.lcb,
    fcPlcftxbxTxt: textboxText.fc,
    lcbPlcftxbxTxt: textboxText.lcb,
    fcPlcfHdrtxbxTxt: headerTextboxText.fc,
    lcbPlcfHdrtxbxTxt: headerTextboxText.lcb,
    fcPlcfFldMom: fieldMom.fc,
    lcbPlcfFldMom: fieldMom.lcb,
    fcPlcfFldHdr: fieldHdr.fc,
    lcbPlcfFldHdr: fieldHdr.lcb,
    fcPlcfFldFtn: fieldFtn.fc,
    lcbPlcfFldFtn: fieldFtn.lcb,
    fcPlcfFldAtn: fieldAtn.fc,
    lcbPlcfFldAtn: fieldAtn.lcb,
    fcPlcfFldEdn: fieldEdn.fc,
    lcbPlcfFldEdn: fieldEdn.lcb,
    fcPlcfFldTxbx: fieldTxbx.fc,
    lcbPlcfFldTxbx: fieldTxbx.lcb,
    fcPlcfFldHdrTxbx: fieldHdrTxbx.fc,
    lcbPlcfFldHdrTxbx: fieldHdrTxbx.lcb,
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
    fcPlcSpaMom: plcfSpaMomOffset + 8 <= FIB_FC_LCB_START + fcLcbCount * 4 ? wordDocument.readUInt32LE(plcfSpaMomOffset) : 0,
    lcbPlcSpaMom: plcfSpaMomOffset + 8 <= FIB_FC_LCB_START + fcLcbCount * 4 ? wordDocument.readUInt32LE(plcfSpaMomOffset + 4) : 0,
    fcPlcSpaHdr: plcfSpaHdrOffset + 8 <= FIB_FC_LCB_START + fcLcbCount * 4 ? wordDocument.readUInt32LE(plcfSpaHdrOffset) : 0,
    lcbPlcSpaHdr: plcfSpaHdrOffset + 8 <= FIB_FC_LCB_START + fcLcbCount * 4 ? wordDocument.readUInt32LE(plcfSpaHdrOffset + 4) : 0,
    fcDggInfo: dggInfoOffset + 8 <= FIB_FC_LCB_START + fcLcbCount * 4 ? wordDocument.readUInt32LE(dggInfoOffset) : 0,
    lcbDggInfo: dggInfoOffset + 8 <= FIB_FC_LCB_START + fcLcbCount * 4 ? wordDocument.readUInt32LE(dggInfoOffset + 4) : 0,
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

  // ── Dop2000 shared flags at offset 504 ────────────────────────────
  // MS-DOC-SPEC/17 Dop2000 bits A-O begin after the 500-byte Dop97,
  // ilvlLastBulletMain, ilvlLastNumberMain, and istdClickParaType fields.
  // Bit N (bit 30 of the 32-bit flags field) is fCharLineUnits: if zero,
  // character-unit indents and line-unit spacing MUST NOT be in use. Preserve
  // the parsed bit as evidence; do not use it to synthesize paragraph values.
  let fCharLineUnits = null;
  if (dop.length >= 508) {
    const dop2000Flags = dop.readUInt32LE(504);
    fCharLineUnits = ((dop2000Flags >> 30) & 1) !== 0;
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
    fCharLineUnits,
    xmlValidation,
    grfFmtFilter,
  };
}

function readDocumentPieces(wordDocument, tableStream, fib) {
  if (fib.fComplex) {
    const encoding = textDecoderEncodingForLid(fib.lid);
    return readPieceTable(tableStream, fib.fcClx, fib.lcbClx).map((piece) => piece.compressed ? { ...piece, encoding } : piece);
  }
  return readNonComplexPieces(wordDocument, tableStream, fib);
}

function readNonComplexPieces(wordDocument, tableStream, fib) {
  if (!fib.fExtChar) {
    throw new Error("Unimplemented Word binary document variant: non-complex non-Unicode text");
  }
  const characterCount = totalDocumentCharacterCount(fib.characterCounts);
  const byteLength = characterCount * 2;
  if (fib.fcMin < 0 || fib.fcMac > wordDocument.length || fib.fcMin > fib.fcMac) {
    throw new Error("Invalid Word binary document: non-complex fcMin/fcMac range is outside WordDocument");
  }

  const fkpInfo = collectFkpInfo(tableStream, wordDocument, fib);
  const textRanges = trimNonComplexTextPadding(
    subtractRanges(
      [{ start: fib.fcMin, end: fib.fcMac }],
      fkpInfo.pageRanges
        .map((range) => ({
          start: Math.max(range.start, fib.fcMin),
          end: Math.min(range.end, fib.fcMac),
        }))
        .filter((range) => range.start < range.end),
    ),
    fkpInfo.fcBoundaries,
    fkpInfo.pageRanges,
    wordDocument,
  );

  const pieces = [];
  let cp = 0;
  let remainingBytes = byteLength;
  for (const range of textRanges) {
    if (remainingBytes === 0) break;
    const availableBytes = range.end - range.start;
    if (availableBytes % 2 !== 0) {
      throw new Error("Invalid Word binary document: non-complex Unicode text range has odd byte length");
    }
    const usedBytes = Math.min(availableBytes, remainingBytes);
    if (usedBytes === 0) continue;
    pieces.push({
      cpStart: cp,
      cpEnd: cp + usedBytes / 2,
      fileOffset: range.start,
      compressed: false,
      nonComplex: true,
    });
    cp += usedBytes / 2;
    remainingBytes -= usedBytes;
  }

  if (remainingBytes !== 0) {
    throw new Error("Invalid Word binary document: non-complex text ranges do not cover FibRgLw97 character counts");
  }

  // MS-DOC-SPEC/15 FibBase.fComplex records whether the last save was an
  // incremental save. When it is zero, no CLX/Pcd piece table is used. For
  // Unicode simple files (FibBase.fExtChar set), consume exactly the
  // FibRgLw97 ccp* character count from the FibBase.fcMin..fcMac region,
  // excluding parsed ChpxFkp/PapxFkp pages explicitly pointed to by the
  // mandatory PlcBteChpx/PlcBtePapx structures (MS-DOC-SPEC/19 PnFkp*).
  return pieces;
}

function collectFkpInfo(tableStream, wordDocument, fib) {
  const chpx = collectFkpInfoFromPlc(tableStream, wordDocument, fib.fcChpx, fib.lcbChpx, "PlcBteChpx");
  const papx = collectFkpInfoFromPlc(tableStream, wordDocument, fib.fcPapx, fib.lcbPapx, "PlcBtePapx");
  return {
    pageRanges: mergeRanges([...chpx.pageRanges, ...papx.pageRanges]),
    fcBoundaries: [...new Set([...chpx.fcBoundaries, ...papx.fcBoundaries])].sort((a, b) => a - b),
  };
}

function collectFkpInfoFromPlc(tableStream, wordDocument, fc, lcb, label) {
  if (lcb < 4) {
    throw new Error(`Invalid Word binary document: missing mandatory ${label}`);
  }
  if (fc + lcb > tableStream.length) {
    throw new Error(`Invalid Word binary document: ${label} is outside the table stream`);
  }
  const binCount = (lcb - 4) / 8;
  if (binCount <= 0 || !Number.isInteger(binCount)) {
    throw new Error(`Invalid Word binary document: malformed ${label}`);
  }
  const ranges = [];
  const fcBoundaries = [];
  for (let i = 0; i <= binCount; i += 1) {
    fcBoundaries.push(tableStream.readUInt32LE(fc + i * 4));
  }
  const pageNumberOffset = fc + (binCount + 1) * 4;
  for (let i = 0; i < binCount; i += 1) {
    const rawPn = tableStream.readUInt32LE(pageNumberOffset + i * 4);
    const pageNumber = rawPn & PNFPN_MASK;
    const pageStart = pageNumber * FKP_PAGE_SIZE;
    const pageEnd = pageStart + FKP_PAGE_SIZE;
    if (pageEnd > wordDocument.length) {
      throw new Error(`Invalid Word binary document: ${label} FKP page is outside WordDocument`);
    }
    const page = wordDocument.subarray(pageStart, pageEnd);
    const crun = page[FKP_PAGE_SIZE - 1];
    if (crun > 0 && (crun + 1) * 4 < FKP_PAGE_SIZE) {
      for (let f = 0; f <= crun; f += 1) {
        fcBoundaries.push(page.readUInt32LE(f * 4));
      }
    }
    ranges.push({ start: pageStart, end: pageEnd });
  }
  return { pageRanges: mergeRanges(ranges), fcBoundaries };
}

function trimNonComplexTextPadding(textRanges, fcBoundaries, pageRanges, wordDocument) {
  const fkpPageStarts = new Set(pageRanges.map((range) => range.start));
  return textRanges.map((range) => {
    if (!fkpPageStarts.has(range.end)) return range;
    const parsedEnd = fcBoundaries
      .filter((fc) => fc >= range.start && fc <= range.end)
      .at(-1);
    if (parsedEnd == null || parsedEnd === range.end) return range;
    const padding = wordDocument.subarray(parsedEnd, range.end);
    if (padding.some((byte) => byte !== 0)) {
      throw new Error("Invalid Word binary document: non-complex text padding before FKP page is non-zero");
    }
    // MS-DOC-SPEC/19 PapxFkp/ChpxFkp rgfc arrays contain the parsed FC
    // boundaries for text before a page-aligned FKP. Bytes after the last
    // parsed FC boundary and before the next 512-byte FKP page are padding,
    // not document characters. Trim only zero padding proven by those FCs.
    return { start: range.start, end: parsedEnd };
  }).filter((range) => range.start < range.end);
}

function subtractRanges(baseRanges, removeRanges) {
  let ranges = baseRanges;
  for (const remove of mergeRanges(removeRanges)) {
    const next = [];
    for (const range of ranges) {
      if (remove.end <= range.start || remove.start >= range.end) {
        next.push(range);
        continue;
      }
      if (remove.start > range.start) {
        next.push({ start: range.start, end: remove.start });
      }
      if (remove.end < range.end) {
        next.push({ start: remove.end, end: range.end });
      }
    }
    ranges = next;
  }
  return ranges;
}

function mergeRanges(ranges) {
  const sorted = ranges
    .filter((range) => range.start < range.end)
    .sort((a, b) => a.start - b.start || a.end - b.end);
  const merged = [];
  for (const range of sorted) {
    const last = merged.at(-1);
    if (last && range.start <= last.end) {
      last.end = Math.max(last.end, range.end);
    } else {
      merged.push({ start: range.start, end: range.end });
    }
  }
  return merged;
}

function totalDocumentCharacterCount(counts = {}) {
  return (counts.body ?? 0)
    + (counts.footnotes ?? 0)
    + (counts.headers ?? 0)
    + (counts.annotations ?? 0)
    + (counts.endnotes ?? 0)
    + (counts.textboxes ?? 0)
    + (counts.headerTextboxes ?? 0);
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
      text += decodeSingleByteText(wordDocument.subarray(startByte, endByte), piece.encoding);
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

function decodeSingleByteText(buffer, encoding) {
  if (!encoding) throw new Error("Invalid compressed text piece: missing parsed LID code page");
  return new TextDecoder(encoding, { fatal: false }).decode(buffer);
}

export function textDecoderEncodingForLid(lid) {
  // FibBase.lid identifies the language of the document. MS-DOC compressed
  // pieces use that language's Windows ANSI code page; this explicit table is
  // the Windows LCID/ANSI-code-page mapping, not a byte-content guess.
  // https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-lcid/
  const exact = new Map([
    [0x0404, "big5"], [0x0c04, "big5"], [0x1404, "big5"], [0x7c04, "big5"],
    [0x0804, "gbk"], [0x1004, "gbk"], [0x0004, "gbk"], [0x7804, "gbk"],
  ]);
  if (exact.has(lid)) return exact.get(lid);
  const primary = lid & 0x03ff;
  if (primary === 0x11) return "shift_jis";
  if (primary === 0x12) return "euc-kr";
  if (primary === 0x1e) return "windows-874";
  if (primary === 0x0d) return "windows-1255";
  if ([0x01, 0x20, 0x29].includes(primary)) return "windows-1256";
  if ([0x02, 0x19, 0x22, 0x23, 0x28, 0x2f, 0x3f, 0x40, 0x44].includes(primary)) return "windows-1251";
  if (primary === 0x08) return "windows-1253";
  if ([0x1f, 0x2c].includes(primary)) return "windows-1254";
  if ([0x25, 0x26, 0x27].includes(primary)) return "windows-1257";
  if (primary === 0x2a) return "windows-1258";
  if ([0x05, 0x0e, 0x15, 0x18, 0x1a, 0x1b, 0x1c, 0x24].includes(primary)) return "windows-1250";
  if ([0x03, 0x06, 0x07, 0x09, 0x0a, 0x0b, 0x0c, 0x0f, 0x10, 0x13, 0x14, 0x16, 0x17, 0x1d, 0x21, 0x2d, 0x36, 0x38, 0x3a, 0x3b, 0x3c, 0x3e, 0x41, 0x45, 0x46, 0x4a, 0x52, 0x56, 0x60, 0x61, 0x62, 0x64].includes(primary)) return "windows-1252";
  throw new Error(`Unimplemented compressed MS-DOC text code page for FibBase.lid 0x${lid.toString(16)}`);
}

export function parseNoteReferencePlc(tableStream, fc, lcb, bodyText, label = "note") {
  if (!lcb) return { cpArray: [], references: [] };
  if (lcb < 4 || (lcb - 4) % 6 !== 0) {
    throw new Error(`Invalid Word binary document: ${label} reference PLC length ${lcb} is not 6n+4 bytes`);
  }
  assertTableRange(tableStream, fc, lcb, `${label} reference PLC`);
  const count = (lcb - 4) / 6;
  const cpBytes = (count + 1) * 4;
  const cpArray = Array.from({ length: count + 1 }, (_, index) => tableStream.readUInt32LE(fc + index * 4));
  for (let i = 1; i < cpArray.length; i += 1) {
    if (cpArray[i] <= cpArray[i - 1]) throw new Error(`Out-of-spec ${label} reference PLC has duplicate or descending CP at index ${i}`);
  }
  const references = Array.from({ length: count }, (_, index) => {
    const cp = cpArray[index];
    if (cp >= bodyText.length) throw new Error(`Out-of-spec ${label} reference CP ${cp} exceeds main document`);
    const indexValue = tableStream.readUInt16LE(fc + cpBytes + index * 2);
    if (indexValue !== 0 && bodyText.charCodeAt(cp) !== 0x02) {
      throw new Error(`Out-of-spec automatic ${label} reference at CP ${cp} is not character 0x02`);
    }
    return { id: index + 1, cp, indexValue, automatic: indexValue !== 0 };
  });
  return { cpArray, references };
}

export function parseNoteTextPlc(tableStream, fc, lcb, storyText, expectedCount, label = "note") {
  if (!lcb) {
    if (storyText.length || expectedCount) throw new Error(`Invalid Word binary document: missing ${label} text PLC`);
    return { cpArray: [], stories: [] };
  }
  if (lcb < 8 || lcb % 4 !== 0) throw new Error(`Invalid Word binary document: malformed ${label} text PLC length ${lcb}`);
  assertTableRange(tableStream, fc, lcb, `${label} text PLC`);
  const cpArray = Array.from({ length: lcb / 4 }, (_, index) => tableStream.readUInt32LE(fc + index * 4));
  for (let i = 1; i < cpArray.length; i += 1) {
    if (cpArray[i] <= cpArray[i - 1]) throw new Error(`Out-of-spec ${label} text PLC has duplicate or descending CP at index ${i}`);
  }
  const storyCount = cpArray.length - 2;
  if (storyCount !== expectedCount) throw new Error(`Invalid Word binary document: ${label} reference/text count mismatch (${expectedCount} vs ${storyCount})`);
  if (cpArray.at(-2) !== storyText.length - 1) {
    throw new Error(`Out-of-spec ${label} text PLC penultimate CP ${cpArray.at(-2)} does not equal story length minus one ${storyText.length - 1}`);
  }
  const stories = Array.from({ length: storyCount }, (_, index) => {
    const cpStart = cpArray[index];
    const cpEnd = cpArray[index + 1];
    if (cpEnd > storyText.length || cpStart >= cpEnd || storyText[cpEnd - 1] !== "\r") {
      throw new Error(`Out-of-spec ${label} text range ${index} does not end with a paragraph mark`);
    }
    const rawText = storyText.slice(cpStart, cpEnd);
    return { id: index + 1, cpStart, cpEnd, rawText, text: rawText.slice(0, -1) };
  });
  return { cpArray, stories };
}

function parseNoteCollection(tableStream, refFc, refLcb, textFc, textLcb, bodyText, storyText, label) {
  const references = parseNoteReferencePlc(tableStream, refFc, refLcb, bodyText, label);
  const texts = parseNoteTextPlc(tableStream, textFc, textLcb, storyText, references.references.length, label);
  return {
    references: references.references.map((reference, index) => ({ ...reference, story: texts.stories[index] })),
    stories: texts.stories,
    referenceCpArray: references.cpArray,
    textCpArray: texts.cpArray,
  };
}

export function parseTextboxTextPlc(tableStream, fc, lcb, storyText, label = "Textbox Document") {
  if (!lcb) {
    if (storyText.length) throw new Error(`Invalid Word binary document: missing ${label} text PLC`);
    return { cpArray: [], entries: [] };
  }
  if (lcb < 30 || (lcb - 4) % 26 !== 0) throw new Error(`Invalid Word binary document: malformed ${label} PLC length ${lcb}`);
  assertTableRange(tableStream, fc, lcb, `${label} PLC`);
  const count = (lcb - 4) / 26;
  const cpBytes = (count + 1) * 4;
  const cpArray = Array.from({ length: count + 1 }, (_, index) => tableStream.readUInt32LE(fc + index * 4));
  for (let i = 1; i < cpArray.length; i += 1) {
    if (cpArray[i] <= cpArray[i - 1]) throw new Error(`Out-of-spec ${label} PLC has duplicate or descending CP at index ${i}`);
  }
  const entries = Array.from({ length: count }, (_, index) => {
    const off = fc + cpBytes + index * 22;
    const fReusableRaw = tableStream.readUInt16LE(off + 8);
    const isLast = index === count - 1;
    const reusable = isLast || fReusableRaw !== 0;
    if (fReusableRaw !== 0 && (fReusableRaw & 1) === 0) throw new Error(`Out-of-spec ${label} FTXBXS ${index} has invalid fReusable`);
    const cpStart = cpArray[index];
    const cpEnd = cpArray[index + 1];
    const lid = tableStream.readUInt32LE(off + 14);
    const entry = {
      index,
      cpStart,
      cpEnd,
      reusable,
      fReusableRaw,
      itxbxsDest: tableStream.readUInt32LE(off + 10),
      lid,
      txidUndo: tableStream.readUInt32LE(off + 18),
      cTxbx: tableStream.readUInt32LE(off),
      cTxbxEdit: tableStream.readUInt32LE(off + 4),
      rawText: storyText.slice(cpStart, Math.min(cpEnd, storyText.length)),
    };
    if (entry.txidUndo !== 0) throw new Error(`Out-of-spec ${label} FTXBXS ${index} has nonzero txidUndo`);
    if (reusable) {
      if ((!isLast && cpEnd - cpStart !== 1) || lid !== 0) throw new Error(`Out-of-spec ${label} reusable FTXBXS ${index}`);
    } else {
      if (cpEnd - cpStart <= 1 || cpEnd > storyText.length || storyText[cpEnd - 1] !== "\r") {
        throw new Error(`Out-of-spec ${label} textbox range ${index}`);
      }
      if (entry.cTxbx <= 0 || entry.cTxbxEdit !== 0 || lid === 0) throw new Error(`Out-of-spec ${label} actual FTXBXS ${index}`);
      entry.text = entry.rawText.slice(0, -1);
    }
    return entry;
  });
  return { cpArray, entries, textboxes: entries.filter((entry) => !entry.reusable) };
}

function parseCommentAuthors(tableStream, fib) {
  if (!fib.lcbGrpXstAtnOwners) return [];
  assertTableRange(tableStream, fib.fcGrpXstAtnOwners, fib.lcbGrpXstAtnOwners, "comment author XST array");
  const end = fib.fcGrpXstAtnOwners + fib.lcbGrpXstAtnOwners;
  const authors = [];
  let off = fib.fcGrpXstAtnOwners;
  while (off < end) {
    if (off + 2 > end) throw new Error("Invalid Word binary document: truncated comment author XST");
    const cch = tableStream.readUInt16LE(off);
    if (cch >= 56 || off + 2 + cch * 2 > end) throw new Error("Out-of-spec comment author XST");
    authors.push(tableStream.subarray(off + 2, off + 2 + cch * 2).toString("utf16le"));
    off += 2 + cch * 2;
  }
  if (new Set(authors).size !== authors.length) throw new Error("Out-of-spec duplicate comment author names");
  return authors;
}

function parseAnnotationBookmarks(tableStream, fib) {
  if (!fib.lcbSttbfAtnBkmk && !fib.lcbPlcfAtnBkf && !fib.lcbPlcfAtnBkl) return new Map();
  if (!fib.lcbSttbfAtnBkmk || !fib.lcbPlcfAtnBkf || !fib.lcbPlcfAtnBkl) throw new Error("Invalid Word binary document: incomplete annotation bookmark tables");
  assertTableRange(tableStream, fib.fcSttbfAtnBkmk, fib.lcbSttbfAtnBkmk, "SttbfAtnBkmk");
  const sttb = tableStream.subarray(fib.fcSttbfAtnBkmk, fib.fcSttbfAtnBkmk + fib.lcbSttbfAtnBkmk);
  if (sttb.length < 6 || sttb.readUInt16LE(0) !== 0xffff || sttb.readUInt16LE(4) !== 10) throw new Error("Out-of-spec SttbfAtnBkmk header");
  const count = sttb.readUInt16LE(2);
  const tags = [];
  let off = 6;
  for (let i = 0; i < count; i += 1) {
    if (off + 12 > sttb.length || sttb.readUInt16LE(off) !== 0) throw new Error("Out-of-spec SttbfAtnBkmk entry");
    const bmc = sttb.readUInt16LE(off + 2);
    const lTag = sttb.readUInt32LE(off + 4);
    const lTagOld = sttb.readInt32LE(off + 8);
    if (bmc !== 0x0100 || lTagOld !== -1) throw new Error("Out-of-spec annotation ATNBE");
    tags.push(lTag);
    off += 12;
  }
  if (off !== sttb.length || new Set(tags).size !== tags.length) throw new Error("Out-of-spec SttbfAtnBkmk length or duplicate lTag");
  if (fib.lcbPlcfAtnBkl !== (count + 1) * 4) throw new Error("Invalid annotation Plcfbkl length");
  const dataBytes = fib.lcbPlcfAtnBkf - (count + 1) * 4;
  if (dataBytes !== count * 4) throw new Error("Invalid annotation Plcfbkf length");
  const bkf = tableStream.subarray(fib.fcPlcfAtnBkf, fib.fcPlcfAtnBkf + fib.lcbPlcfAtnBkf);
  const bkl = tableStream.subarray(fib.fcPlcfAtnBkl, fib.fcPlcfAtnBkl + fib.lcbPlcfAtnBkl);
  const starts = readBookmarkPlcCps(bkf, count, "annotation Plcfbkf");
  const ends = readBookmarkPlcCps(bkl, count, "annotation Plcfbkl");
  const map = new Map();
  for (let i = 0; i < count; i += 1) {
    const ibkl = bkf.readUInt16LE((count + 1) * 4 + i * 4);
    parseBkc(bkf.readUInt16LE((count + 1) * 4 + i * 4 + 2), `annotation ${tags[i]}`);
    if (ibkl >= count || starts[i] > ends[ibkl]) throw new Error("Out-of-spec annotation bookmark linkage");
    map.set(tags[i], { cpStart: starts[i], cpEnd: ends[ibkl] });
  }
  return map;
}

export function parseComments(tableStream, fib, bodyText, storyText) {
  if (!fib.lcbPlcfandRef && !fib.lcbPlcfandTxt) {
    if (storyText.length) throw new Error("Invalid Word binary document: comment story exists without comment PLCs");
    return { comments: [], authors: [] };
  }
  if (!fib.lcbPlcfandRef || !fib.lcbPlcfandTxt) throw new Error("Invalid Word binary document: incomplete comment PLCs");
  const authors = parseCommentAuthors(tableStream, fib);
  const bookmarks = parseAnnotationBookmarks(tableStream, fib);
  if (fib.lcbPlcfandRef < 4 || (fib.lcbPlcfandRef - 4) % 34 !== 0) throw new Error("Invalid comment reference PLC length");
  assertTableRange(tableStream, fib.fcPlcfandRef, fib.lcbPlcfandRef, "comment reference PLC");
  const count = (fib.lcbPlcfandRef - 4) / 34;
  const cpBytes = (count + 1) * 4;
  const cps = Array.from({ length: count + 1 }, (_, index) => tableStream.readUInt32LE(fib.fcPlcfandRef + index * 4));
  for (let i = 1; i < cps.length; i += 1) if (cps[i] <= cps[i - 1]) throw new Error("Out-of-spec comment reference CP array");
  const textPlc = parseNoteTextPlc(tableStream, fib.fcPlcfandTxt, fib.lcbPlcfandTxt, storyText, count, "comment");
  const comments = Array.from({ length: count }, (_, index) => {
    const cp = cps[index];
    if (cp >= bodyText.length || bodyText.charCodeAt(cp) !== 0x05) throw new Error(`Out-of-spec comment reference at CP ${cp}`);
    const off = fib.fcPlcfandRef + cpBytes + index * 30;
    const cch = tableStream.readUInt16LE(off);
    if (cch > 9) throw new Error("Out-of-spec comment initials");
    const initials = tableStream.subarray(off + 2, off + 2 + cch * 2).toString("utf16le");
    const ibst = tableStream.readUInt16LE(off + 20);
    if (ibst >= authors.length) throw new Error(`Out-of-spec comment author index ${ibst}`);
    if (tableStream.readUInt16LE(off + 22) !== 0 || tableStream.readUInt16LE(off + 24) !== 0) throw new Error("Out-of-spec comment ATRD reserved fields");
    const lTagBkmk = tableStream.readInt32LE(off + 26);
    const range = lTagBkmk === -1 ? { cpStart: cp, cpEnd: cp } : bookmarks.get(lTagBkmk);
    if (!range) throw new Error(`Invalid comment annotation bookmark tag ${lTagBkmk}`);
    const story = textPlc.stories[index];
    if (story.text.charCodeAt(0) !== 0x05) throw new Error(`Out-of-spec comment text ${index} does not begin with 0x05`);
    return { id: index, cp, initials, author: authors[ibst], authorIndex: ibst, lTagBkmk, ...range, story: { ...story, text: story.text.slice(1), rawText: story.rawText.slice(1) } };
  });
  return { comments, authors, referenceCpArray: cps, textCpArray: textPlc.cpArray };
}

function parseFieldTables(tableStream, fib, subdocuments) {
  return {
    body: parseFieldPlc(tableStream, fib.fcPlcfFldMom, fib.lcbPlcfFldMom, subdocuments.body.rawText, "Main Document"),
    headers: parseFieldPlc(tableStream, fib.fcPlcfFldHdr, fib.lcbPlcfFldHdr, subdocuments.headers.rawText, "Header Document"),
    footnotes: parseFieldPlc(tableStream, fib.fcPlcfFldFtn, fib.lcbPlcfFldFtn, subdocuments.footnotes.rawText, "Footnote Document"),
    annotations: parseFieldPlc(tableStream, fib.fcPlcfFldAtn, fib.lcbPlcfFldAtn, subdocuments.annotations.rawText, "Comment Document"),
    endnotes: parseFieldPlc(tableStream, fib.fcPlcfFldEdn, fib.lcbPlcfFldEdn, subdocuments.endnotes.rawText, "Endnote Document"),
    textboxes: parseFieldPlc(tableStream, fib.fcPlcfFldTxbx, fib.lcbPlcfFldTxbx, subdocuments.textboxes.rawText, "Textbox Document"),
    headerTextboxes: parseFieldPlc(tableStream, fib.fcPlcfFldHdrTxbx, fib.lcbPlcfFldHdrTxbx, subdocuments.headerTextboxes.rawText, "Header Textbox Document"),
  };
}

export function parseFieldPlc(tableStream, fc, lcb, storyText = "", label = "document part") {
  // MS-DOC-SPEC/18 Plcfld is a PLC with 4-byte CPs and 2-byte Fld records:
  // lcb = 4 * (recordCount + 1) + 2 * recordCount. The terminal CP is
  // deliberately not interpreted as a character position; the specification
  // defines it as sorted/largest but otherwise undefined.
  if (!lcb) return { cpArray: [], records: [], fields: [] };
  if (lcb < 4 || (lcb - 4) % 6 !== 0) {
    throw new Error(`Invalid Word binary document: ${label} Plcfld length ${lcb} is not 6n+4 bytes`);
  }
  assertTableRange(tableStream, fc, lcb, `${label} Plcfld`);
  const recordCount = (lcb - 4) / 6;
  const cpBytes = (recordCount + 1) * 4;
  const cpArray = Array.from({ length: recordCount + 1 }, (_, index) => tableStream.readUInt32LE(fc + index * 4));
  for (let i = 1; i < cpArray.length; i += 1) {
    if (cpArray[i] <= cpArray[i - 1]) {
      throw new Error(`Out-of-spec ${label} Plcfld CP array is not strictly ascending at index ${i}`);
    }
  }

  const records = [];
  const fields = [];
  const stack = [];
  for (let i = 0; i < recordCount; i += 1) {
    const cp = cpArray[i];
    if (cp < 0 || cp >= storyText.length) {
      throw new Error(`Out-of-spec ${label} Plcfld field CP ${cp} exceeds story character range ${storyText.length}`);
    }
    const fldOffset = fc + cpBytes + i * 2;
    const fldch = tableStream[fldOffset];
    const ch = fldch & 0x1f;
    const grffld = tableStream[fldOffset + 1];
    if (ch !== 0x13 && ch !== 0x14 && ch !== 0x15) {
      throw new Error(`Out-of-spec ${label} Plcfld Fld ${i} has fldch ${ch}`);
    }
    if (storyText.charCodeAt(cp) !== ch) {
      throw new Error(`Out-of-spec ${label} Plcfld Fld ${i} does not match story character at CP ${cp}`);
    }
    const record = { index: i, cp, fldch, ch, grffld };
    records.push(record);

    if (ch === 0x13) {
      record.fieldType = grffld;
      const field = { begin: record, separator: null, end: null, nested: stack.length > 0 };
      fields.push(field);
      stack.push(field);
      continue;
    }
    if (ch === 0x14) {
      if (!stack.length) {
        throw new Error(`Out-of-spec ${label} Plcfld separator at CP ${cp} has no open field`);
      }
      const field = stack.at(-1);
      if (field.separator) {
        throw new Error(`Out-of-spec ${label} Plcfld field beginning at CP ${field.begin.cp} has multiple separators`);
      }
      field.separator = record;
      record.field = field;
      continue;
    }
    if (!stack.length) {
      throw new Error(`Out-of-spec ${label} Plcfld end at CP ${cp} has no open field`);
    }
    const field = stack.pop();
    field.end = record;
    record.field = field;
    record.endFlags = {
      differ: (grffld & 0x01) !== 0,
      zombieEmbed: (grffld & 0x02) !== 0,
      resultsDirty: (grffld & 0x04) !== 0,
      resultsEdited: (grffld & 0x08) !== 0,
      locked: (grffld & 0x10) !== 0,
      privateResult: (grffld & 0x20) !== 0,
      nested: (grffld & 0x40) !== 0,
      hasSeparator: (grffld & 0x80) !== 0,
    };
    if (record.endFlags.hasSeparator !== Boolean(field.separator)) {
      throw new Error(`Out-of-spec ${label} Plcfld end at CP ${cp} has inconsistent fHasSep`);
    }
    if (record.endFlags.nested !== field.nested) {
      throw new Error(`Out-of-spec ${label} Plcfld end at CP ${cp} has inconsistent fNested`);
    }
    field.begin.field = field;
    field.begin.endFlags = record.endFlags;
  }
  if (stack.length) {
    throw new Error(`Out-of-spec ${label} Plcfld contains ${stack.length} unterminated field(s)`);
  }
  return { cpArray, records, fields };
}

export function parsePlcfHdd(tableStream, fib, headerText = "", sectionCount = null) {
  // MS-DOC-SPEC/13 Headers: PlcfHdd is a PLC containing only CPs. The first
  // six ranges are note separator stories, followed by exactly six
  // header/footer stories per section in this order: even header, odd header,
  // even footer, odd footer, first header, first footer.
  if (!fib.lcbPlcfHdd) {
    return { cpArray: [], separatorCount: 6, stories: [], sections: [] };
  }
  if (fib.lcbPlcfHdd < 8 || fib.lcbPlcfHdd % 4 !== 0) {
    throw new Error(`Invalid Word binary document: PlcfHdd length ${fib.lcbPlcfHdd} is not a CP PLC`);
  }
  assertTableRange(tableStream, fib.fcPlcfHdd, fib.lcbPlcfHdd, "PlcfHdd");

  const plcf = tableStream.subarray(fib.fcPlcfHdd, fib.fcPlcfHdd + fib.lcbPlcfHdd);
  const cpArray = [];
  for (let off = 0; off < plcf.length; off += 4) {
    cpArray.push(plcf.readUInt32LE(off));
  }
  if (cpArray[0] !== 0) {
    throw new Error(`Out-of-spec PlcfHdd first CP ${cpArray[0]} is not zero`);
  }
  for (let i = 1; i < cpArray.length; i += 1) {
    if (cpArray[i] < cpArray[i - 1]) {
      throw new Error(`Out-of-spec PlcfHdd CP array is not ascending at index ${i}`);
    }
  }

  const headerCharacterCount = fib.characterCounts?.headers ?? headerText.length;
  if (cpArray.at(-1) !== headerCharacterCount) {
    throw new Error(`Out-of-spec PlcfHdd final CP ${cpArray.at(-1)} does not equal ccpHdd ${headerCharacterCount}`);
  }
  if (headerText.length !== headerCharacterCount) {
    throw new Error(`Invalid Word binary document: header story text length ${headerText.length} does not equal ccpHdd ${headerCharacterCount}`);
  }

  const storyCount = cpArray.length - 1;
  if (storyCount < 6 || (storyCount - 6) % 6 !== 0) {
    throw new Error(`Out-of-spec PlcfHdd story count ${storyCount} is not six separators plus six stories per section`);
  }
  const parsedSectionCount = (storyCount - 6) / 6;
  if (sectionCount != null && parsedSectionCount !== sectionCount) {
    throw new Error(`Out-of-spec PlcfHdd has ${parsedSectionCount} section groups, expected ${sectionCount}`);
  }

  const separatorKinds = [
    "footnoteSeparator",
    "footnoteContinuationSeparator",
    "footnoteContinuationNotice",
    "endnoteSeparator",
    "endnoteContinuationSeparator",
    "endnoteContinuationNotice",
  ];
  const sectionKinds = [
    "evenHeader",
    "defaultHeader",
    "evenFooter",
    "defaultFooter",
    "firstHeader",
    "firstFooter",
  ];
  const stories = [];
  for (let i = 0; i < storyCount; i += 1) {
    const cpStart = cpArray[i];
    const cpEnd = cpArray[i + 1];
    const rawText = headerText.slice(cpStart, cpEnd);
    if (rawText && !rawText.endsWith("\r")) {
      throw new Error(`Out-of-spec PlcfHdd story ${i} is non-empty without a guard paragraph mark`);
    }
    const sectionIndex = i < 6 ? null : Math.floor((i - 6) / 6);
    const kind = i < 6 ? separatorKinds[i] : sectionKinds[(i - 6) % 6];
    stories.push({
      index: i,
      kind,
      sectionIndex,
      cpStart,
      cpEnd,
      rawText,
      // The final paragraph mark is a story guard and is not part of the
      // header/footer content (MS-DOC-SPEC/13 Headers).
      text: rawText ? rawText.slice(0, -1) : "",
      empty: cpStart === cpEnd,
    });
  }

  return {
    cpArray,
    separatorCount: 6,
    stories,
    separators: stories.slice(0, 6),
    sections: Array.from({ length: parsedSectionCount }, (_, sectionIndex) => {
      const group = stories.slice(6 + sectionIndex * 6, 12 + sectionIndex * 6);
      return Object.fromEntries(group.map((story) => [story.kind, story]));
    }),
  };
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
  // MS-DOC-SPEC/19 FBKF stores a 2-byte ibkl followed by a 2-byte BKC.
  // ibkl indexes the end CP in Plcfbkl.
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

  const bkf = tableStream.subarray(fib.fcPlcfBkf, fib.fcPlcfBkf + fib.lcbPlcfBkf);
  const bkl = tableStream.subarray(fib.fcPlcfBkl, fib.fcPlcfBkl + fib.lcbPlcfBkl);
  const startCps = readBookmarkPlcCps(bkf, names.length, "Plcfbkf");
  const endCps = readBookmarkPlcCps(bkl, names.length, "Plcfbkl");

  const seenIbkl = new Set();
  return names.map((name, index) => {
    const start = startCps[index];
    const dataOffset = (names.length + 1) * 4 + index * bkfDataSize;
    const ibkl = bkf.readUInt16LE(dataOffset);
    if (ibkl >= names.length) {
      throw new Error(`Invalid Word bookmark ${name}: end index ${ibkl} is outside Plcfbkl`);
    }
    if (seenIbkl.has(ibkl)) {
      throw new Error(`Out-of-spec Word bookmark ${name}: duplicate FBKF.ibkl ${ibkl}`);
    }
    seenIbkl.add(ibkl);
    const bkc = bkf.readUInt16LE(dataOffset + 2);
    const bkcInfo = parseBkc(bkc, name);
    const end = endCps[ibkl];
    if (start > end) {
      throw new Error(`Out-of-spec Word bookmark ${name}: start CP ${start} is greater than end CP ${end}`);
    }
    return {
      id: index,
      name,
      cpStart: start,
      cpEnd: end,
      // MS-DOC-SPEC/19 FBKF.ibkl is the zero-based index into the paired
      // Plcfbkl whose CP gives this bookmark's limit. Preserve it because
      // duplicate start CPs are legal in Plcfbkf, so ibkl is the explicit
      // parsed linkage rather than an inferred row/order relationship.
      ibkl,
      bkc,
      bkcInfo,
    };
  });
}

function parseBkc(raw, bookmarkName) {
  const bkc = {
    raw,
    // MS-DOC-SPEC/19 BKC: low 7 bits are itcFirst. It is ignored unless
    // fCol is set, but preserving it avoids inferring table-column state.
    itcFirst: raw & 0x007f,
    fPub: ((raw >> 7) & 0x0001) !== 0,
    itcLim: (raw >> 8) & 0x003f,
    fNative: ((raw >> 14) & 0x0001) !== 0,
    fCol: ((raw >> 15) & 0x0001) !== 0,
  };
  if (bkc.fPub) {
    // MS-DOC-SPEC/19 BKC.fPub MUST be zero and ignored.
    throw new Error(`Out-of-spec Word bookmark ${bookmarkName}: BKC.fPub must be zero`);
  }
  if (bkc.fCol && bkc.itcFirst >= bkc.itcLim) {
    // MS-DOC-SPEC/19 BKC: for all bookmark types, itcFirst MUST be less
    // than itcLim when fCol is nonzero.
    throw new Error(`Out-of-spec Word bookmark ${bookmarkName}: BKC itcFirst must be less than itcLim`);
  }
  return bkc;
}

function readBookmarkPlcCps(plc, bookmarkCount, label) {
  const cps = [];
  for (let i = 0; i <= bookmarkCount; i += 1) {
    cps.push(plc.readUInt32LE(i * 4));
  }
  for (let i = 1; i < cps.length; i += 1) {
    if (cps[i] < cps[i - 1]) {
      // MS-DOC-SPEC/12 PLC: CP arrays MUST appear in ascending order.
      // MS-DOC-SPEC/18 bookmark PLCs may contain duplicate CPs, so equality
      // is allowed but decreasing CPs are invalid and must fail fast.
      throw new Error(`Out-of-spec Word bookmark ${label}: CP array is not ascending at index ${i}`);
    }
  }
  return cps;
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

function parseInlinePictures(data, subdocuments, allCharacterProperties) {
  const pictures = [];
  const parsedByLocation = new Map();
  for (const [storyName, story] of Object.entries(subdocuments)) {
    for (let localCp = 0; localCp < story.rawText.length; localCp += 1) {
      if (story.rawText.charCodeAt(localCp) !== 0x0001) continue;
      const globalCp = story.cpStart + localCp;
      const properties = allCharacterProperties[globalCp] ?? {};
      if (!Number.isInteger(properties.pictureLocation)) {
        throw new Error(`Out-of-spec picture character in ${storyName} at CP ${localCp}: missing sprmCPicLocation`);
      }
      if (properties.specialCharacter !== true) {
        throw new Error(`Out-of-spec picture character in ${storyName} at CP ${localCp}: sprmCFSpec is not set`);
      }
      if (properties.binaryData) {
        throw new Error(`Excluded binary-data picture character in ${storyName} at CP ${localCp}`);
      }
      if (properties.ole2 || properties.embeddedObject) {
        throw new Error(`Excluded OLE object picture character in ${storyName} at CP ${localCp}`);
      }
      if (!data) {
        throw new Error(`Invalid Word binary document: picture character in ${storyName} references a missing Data stream`);
      }
      let picture = parsedByLocation.get(properties.pictureLocation);
      if (!picture) {
        picture = parsePicfAndOfficeArtData(data, properties.pictureLocation);
        parsedByLocation.set(properties.pictureLocation, picture);
      }
      pictures.push({
        ...picture,
        borders: { ...picture.borders, ...(properties.pictureBorders ?? {}) },
        story: storyName,
        cp: localCp,
        globalCp,
        properties,
      });
    }
  }
  return pictures;
}

export function parsePicfAndOfficeArtData(data, offset = 0) {
  if (!Buffer.isBuffer(data)) data = Buffer.from(data ?? []);
  if (!Number.isInteger(offset) || offset < 0 || offset + 68 > data.length) {
    throw new Error(`Invalid PICFAndOfficeArtData offset ${offset}`);
  }
  const lcb = data.readInt32LE(offset);
  if (lcb < 68 || offset + lcb > data.length) {
    throw new Error(`Out-of-spec PICF length ${lcb} at Data offset ${offset}`);
  }
  const end = offset + lcb;
  const cbHeader = data.readUInt16LE(offset + 4);
  if (cbHeader !== 0x44) {
    throw new Error(`Out-of-spec PICF.cbHeader 0x${cbHeader.toString(16)}`);
  }
  const mm = data.readInt16LE(offset + 6);
  if (mm !== 0x0064 && mm !== 0x0066) {
    throw new Error(`Out-of-spec PICF MFPF.mm 0x${(mm & 0xffff).toString(16)}`);
  }
  if (data.readUInt16LE(offset + 12) !== 0) {
    throw new Error("Out-of-spec PICF MFPF.swHMF must be zero");
  }
  if (data.readUInt32LE(offset + 18) !== 0 || data.readUInt32LE(offset + 24) !== 0) {
    throw new Error("Out-of-spec PICF_Shape padding must be zero");
  }
  const dxaGoal = data.readInt16LE(offset + 28);
  const dyaGoal = data.readInt16LE(offset + 30);
  const mx = data.readUInt16LE(offset + 32);
  const my = data.readUInt16LE(offset + 34);
  if (dxaGoal <= 0 || dyaGoal <= 0) {
    throw new Error(`Out-of-spec PICMID goal size ${dxaGoal}x${dyaGoal}`);
  }
  for (const reservedOffset of [36, 38, 40, 42, 62, 64]) {
    if (data.readInt16LE(offset + reservedOffset) !== 0) {
      throw new Error(`Out-of-spec PICMID reserved field at offset ${reservedOffset}`);
    }
  }
  if (data[offset + 44] !== 0) {
    throw new Error("Out-of-spec PICMID.fReserved must be zero");
  }
  if (data.readUInt16LE(offset + 66) !== 0) {
    throw new Error("Out-of-spec PICF.cProps must be zero");
  }

  let pictureOffset = offset + cbHeader;
  let sourceName = null;
  if (mm === 0x0066) {
    if (pictureOffset >= end) throw new Error("Invalid PICF shape-file name");
    const cchPicName = data[pictureOffset++];
    if (pictureOffset + cchPicName > end) throw new Error("Invalid PICF shape-file name length");
    sourceName = data.subarray(pictureOffset, pictureOffset + cchPicName).toString("latin1");
    pictureOffset += cchPicName;
  }

  const shapeRecord = readOfficeArtRecord(data, pictureOffset, end);
  if (shapeRecord.type !== OFFICE_ART_SP_CONTAINER || shapeRecord.ver !== 0x0f) {
    throw new Error(`Out-of-spec OfficeArtInlineSpContainer shape record 0x${shapeRecord.type.toString(16)}`);
  }
  const shape = parseOfficeArtShapeContainer(data, shapeRecord);
  const bseRecord = readOfficeArtRecord(data, shapeRecord.end, end);
  if (bseRecord.type !== OFFICE_ART_BSE || bseRecord.ver !== 0x02) {
    throw new Error(`Out-of-spec OfficeArtInlineSpContainer BSE record 0x${bseRecord.type.toString(16)}`);
  }
  if (bseRecord.end !== end) {
    throw new Error("Out-of-spec PICFAndOfficeArtData has trailing bytes after OfficeArtBSE");
  }
  const image = parseOfficeArtBse(data, bseRecord);
  const widthTwips = Math.round(dxaGoal * mx / 1000);
  const heightTwips = Math.round(dyaGoal * my / 1000);
  if (widthTwips < 15 || widthTwips > 31680 || heightTwips < 15 || heightTwips > 31680) {
    throw new Error(`Out-of-spec PICMID display size ${widthTwips}x${heightTwips} twips`);
  }
  return {
    dataOffset: offset,
    byteLength: lcb,
    sourceName,
    dxaGoal,
    dyaGoal,
    mx,
    my,
    widthTwips,
    heightTwips,
    borders: {
      top: parseBrc80Raw(data[offset + 46], data[offset + 47], data[offset + 48], data[offset + 49] & 0x1f),
      left: parseBrc80Raw(data[offset + 50], data[offset + 51], data[offset + 52], data[offset + 53] & 0x1f),
      bottom: parseBrc80Raw(data[offset + 54], data[offset + 55], data[offset + 56], data[offset + 57] & 0x1f),
      right: parseBrc80Raw(data[offset + 58], data[offset + 59], data[offset + 60], data[offset + 61] & 0x1f),
    },
    shape,
    ...image,
  };
}

function parseOfficeArtBse(buffer, record) {
  if (record.len < 36) throw new Error("Out-of-spec OfficeArtBSE is shorter than 36 bytes");
  const off = record.content;
  const btWin32 = buffer[off];
  const btMacOS = buffer[off + 1];
  const cb = buffer.readUInt32LE(off + 20);
  const cRef = buffer.readUInt32LE(off + 24);
  const foDelay = buffer.readUInt32LE(off + 28);
  const usage = buffer[off + 32];
  const cbName = buffer[off + 33];
  if (buffer[off + 34] !== 0 || buffer[off + 35] !== 0) {
    throw new Error("Out-of-spec OfficeArtBSE unused fields must be zero");
  }
  const nameEnd = off + 36 + cbName;
  if (nameEnd > record.end) throw new Error("Invalid OfficeArtBSE name length");
  if (foDelay !== 0) {
    throw new Error(`Unimplemented delayed OfficeArt blip at offset ${foDelay}`);
  }
  const blip = readOfficeArtRecord(buffer, nameEnd, record.end);
  if (blip.end !== record.end) throw new Error("Out-of-spec OfficeArtBSE has trailing bytes after blip");
  const image = parseOfficeArtBlip(buffer, blip);
  if (cb !== blip.len + 8) {
    throw new Error(`Out-of-spec OfficeArtBSE.cb ${cb} does not match blip size ${blip.len + 8}`);
  }
  return { btWin32, btMacOS, cRef, usage, ...image };
}

function parseOfficeArtBlip(buffer, record) {
  // [MS-ODRAW] 2.2.21-2.2.30 define the record-instance values and
  // UID/tag prefixes for each bitmap blip. Keep this explicit: searching for
  // an image signature would make malformed records look valid.
  // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-odraw/
  const raster = new Map([
    [OFFICE_ART_BLIP_JPEG, new Map([[0x046a, { prefix: 17, extension: "jpg", contentType: "image/jpeg" }], [0x06e2, { prefix: 33, extension: "jpg", contentType: "image/jpeg" }]])],
    [OFFICE_ART_BLIP_PNG, new Map([[0x06e0, { prefix: 17, extension: "png", contentType: "image/png" }], [0x06e1, { prefix: 33, extension: "png", contentType: "image/png" }]])],
    [OFFICE_ART_BLIP_DIB, new Map([[0x07a8, { prefix: 17, extension: "bmp", contentType: "image/bmp", dib: true }]])],
    [OFFICE_ART_BLIP_TIFF, new Map([[0x06e4, { prefix: 17, extension: "tif", contentType: "image/tiff" }], [0x06e5, { prefix: 33, extension: "tif", contentType: "image/tiff" }]])],
  ]);
  const byInstance = raster.get(record.type);
  if (byInstance) {
    const format = byInstance.get(record.inst);
    if (!format) throw new Error(`Out-of-spec OfficeArt bitmap blip instance 0x${record.inst.toString(16)}`);
    if (record.len <= format.prefix) throw new Error("Invalid OfficeArt bitmap blip payload");
    let bytes = Buffer.from(buffer.subarray(record.content + format.prefix, record.end));
    if (format.dib) bytes = dibToBmp(bytes);
    validateImageSignature(bytes, format.extension);
    return { extension: format.extension, contentType: format.contentType, bytes };
  }
  if (record.type === OFFICE_ART_BLIP_EMF || record.type === OFFICE_ART_BLIP_WMF || record.type === OFFICE_ART_BLIP_PICT) {
    if (record.len <= 50) throw new Error("Invalid OfficeArt metafile blip payload");
    const header = record.content + 16;
    const uncompressedSize = buffer.readUInt32LE(header);
    const compressedSize = buffer.readUInt32LE(header + 28);
    const compression = buffer[header + 32];
    const filter = buffer[header + 33];
    if (filter !== 0xfe) throw new Error(`Unimplemented OfficeArt metafile filter 0x${filter.toString(16)}`);
    const stored = buffer.subarray(header + 34, record.end);
    if (stored.length !== compressedSize) throw new Error("Out-of-spec OfficeArt metafile compressed size");
    const bytes = compression === 0x00 ? inflateRawSync(stored)
      : compression === 0xfe ? Buffer.from(stored)
        : (() => { throw new Error(`Out-of-spec OfficeArt metafile compression 0x${compression.toString(16)}`); })();
    if (bytes.length !== uncompressedSize) throw new Error("Out-of-spec OfficeArt metafile uncompressed size");
    if (record.type === OFFICE_ART_BLIP_PICT) throw new Error("Excluded Macintosh PICT image: no interoperable DOCX image part mapping");
    return record.type === OFFICE_ART_BLIP_EMF
      ? { extension: "emf", contentType: "image/x-emf", bytes }
      : { extension: "wmf", contentType: "image/x-wmf", bytes };
  }
  throw new Error(`Unimplemented OfficeArt blip record type 0x${record.type.toString(16)}`);
}

function validateImageSignature(bytes, extension) {
  const valid = extension === "png" ? bytes.subarray(0, 8).equals(Buffer.from([0x89,0x50,0x4e,0x47,0x0d,0x0a,0x1a,0x0a]))
    : extension === "jpg" ? bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff
      : extension === "tif" ? bytes.length >= 4 && ((bytes[0] === 0x49 && bytes[1] === 0x49 && bytes[2] === 0x2a && bytes[3] === 0) || (bytes[0] === 0x4d && bytes[1] === 0x4d && bytes[2] === 0 && bytes[3] === 0x2a))
        : extension === "bmp" ? bytes.length >= 2 && bytes[0] === 0x42 && bytes[1] === 0x4d
          : true;
  if (!valid) throw new Error(`Invalid OfficeArt ${extension} image signature`);
}

function dibToBmp(dib) {
  if (dib.length < 40) throw new Error("Unsupported truncated DIB image");
  const headerSize = dib.readUInt32LE(0);
  if (headerSize < 40 || headerSize > dib.length) throw new Error(`Unsupported DIB header size ${headerSize}`);
  const bitCount = dib.readUInt16LE(14);
  const compression = dib.readUInt32LE(16);
  const colorsUsed = dib.readUInt32LE(32);
  const paletteEntries = colorsUsed || (bitCount <= 8 ? 2 ** bitCount : 0);
  const bitMasks = compression === 3 && headerSize === 40 ? 12 : 0;
  const pixelOffset = 14 + headerSize + bitMasks + paletteEntries * 4;
  if (pixelOffset - 14 > dib.length) throw new Error("Out-of-spec DIB palette exceeds payload");
  const bmp = Buffer.alloc(14 + dib.length);
  bmp.write("BM", 0, "ascii");
  bmp.writeUInt32LE(bmp.length, 2);
  bmp.writeUInt32LE(pixelOffset, 10);
  dib.copy(bmp, 14);
  return bmp;
}

function parseOfficeArtContent(tableStream, fib) {
  if (!fib.lcbDggInfo) return { body: new Map(), headers: new Map() };
  assertTableRange(tableStream, fib.fcDggInfo, fib.lcbDggInfo, "OfficeArtContent");
  const content = tableStream.subarray(fib.fcDggInfo, fib.fcDggInfo + fib.lcbDggInfo);
  if (content.length < 8) {
    throw new Error("Invalid Word binary document: truncated OfficeArtContent");
  }

  // MS-DOC-SPEC/15 fcDggInfo/lcbDggInfo point to OfficeArtContent.
  // MS-DOC-SPEC/19 OfficeArtContent begins with an OfficeArtDggContainer
  // followed by OfficeArtWordDrawing records, each prefixed by dgglbl.
  const dgg = readOfficeArtRecord(content, 0, content.length);
  if (dgg.type !== OFFICE_ART_DGG_CONTAINER || dgg.ver !== 0x0f) {
    throw new Error(`Out-of-spec OfficeArtContent: expected DggContainer, got type 0x${dgg.type.toString(16)}`);
  }

  const drawings = { body: new Map(), headers: new Map() };
  let off = dgg.end;
  while (off < content.length) {
    const dgglbl = content[off];
    off += 1;
    const drawing = readOfficeArtRecord(content, off, content.length);
    if (drawing.type !== OFFICE_ART_DG_CONTAINER || drawing.ver !== 0x0f) {
      throw new Error(`Out-of-spec OfficeArtWordDrawing: expected DgContainer, got type 0x${drawing.type.toString(16)}`);
    }
    if (dgglbl !== 0 && dgglbl !== 1) {
      throw new Error(`Out-of-spec OfficeArtWordDrawing dgglbl ${dgglbl}`);
    }
    const target = dgglbl === 0 ? drawings.body : drawings.headers;
    for (const shape of parseOfficeArtShapeContainers(content, drawing.content, drawing.end)) {
      if (shape.spid != null) target.set(shape.spid, shape);
    }
    off = drawing.end;
  }
  return drawings;
}

function parseOfficeArtShapeContainers(buffer, start, end) {
  const shapes = [];
  const clientDataBySpid = parseOfficeArtClientData(buffer, start, end);
  let drawingOrder = 0;

  function walk(off, limit) {
    while (off + 8 <= limit) {
      const record = readOfficeArtRecord(buffer, off, limit);
      if (record.type === OFFICE_ART_SP_CONTAINER) {
        const shape = parseOfficeArtShapeContainer(buffer, record);
        if (shape?.spid != null) {
          const clientData = clientDataBySpid.get(shape.spid) ?? null;
          shapes.push({
            ...shape,
            drawingOrder: clientData?.drawingOrder ?? drawingOrder,
            clientData,
          });
          drawingOrder += 1;
        }
      } else if (record.ver === 0x0f || record.type === OFFICE_ART_SPGR_CONTAINER) {
        walk(record.content, record.end);
      }
      off = record.end;
    }
    if (off !== limit) {
      throw new Error("Invalid Word binary document: OfficeArt container has trailing partial record");
    }
  }

  walk(start, end);
  return shapes;
}

function parseOfficeArtShapeContainer(buffer, record) {
  let off = record.content;
  let fsp = null;
  let fopt = null;
  let tertiaryFopt = null;
  while (off + 8 <= record.end) {
    const child = readOfficeArtRecord(buffer, off, record.end);
    if (child.type === OFFICE_ART_FSP) fsp = child;
    if (child.type === OFFICE_ART_FOPT) fopt = child;
    if (child.type === OFFICE_ART_TERTIARY_FOPT) tertiaryFopt = child;
    off = child.end;
  }
  if (off !== record.end) {
    throw new Error("Invalid Word binary document: OfficeArtSpContainer has trailing partial record");
  }
  if (!fsp) return null;
  if (fsp.len !== 8) {
    throw new Error(`Out-of-spec OfficeArtFSP: expected 8 bytes, got ${fsp.len}`);
  }

  const shape = {
    spid: buffer.readUInt32LE(fsp.content),
    shapeType: fsp.inst,
    fspFlags: buffer.readUInt32LE(fsp.content + 4),
  };
  if (fopt) {
    const properties = parseOfficeArtFopt(buffer, fopt, "OfficeArtFOPT");
    shape.properties = properties;
    const nameProperty = properties.find((prop) => prop.pid === OFFICE_ART_SHAPE_NAME_PID && prop.complexData);
    if (nameProperty) {
      shape.name = stripNullTerminator(nameProperty.complexData.toString("utf16le"));
      const idMatch = shape.name.match(/(\d+)$/);
      if (idMatch) shape.docPrId = Number.parseInt(idMatch[1], 10);
    }
    const descriptionProperty = properties.find((prop) => prop.pid === OFFICE_ART_SHAPE_DESCRIPTION_PID && prop.complexData);
    if (descriptionProperty) {
      shape.description = stripNullTerminator(descriptionProperty.complexData.toString("utf16le"));
    }
  }
  if (tertiaryFopt) {
    const properties = parseOfficeArtFopt(buffer, tertiaryFopt, "OfficeArtTertiaryFOPT");
    shape.tertiaryProperties = properties;
    const gfxProperty = properties.find((prop) => prop.pid === OFFICE_ART_GFXDATA_PID && prop.complexData);
    if (gfxProperty) {
      shape.gfxData = extractOfficeArtGfxData(gfxProperty.complexData);
      const e2oDoc = parseZipEntries(shape.gfxData).get("drs/e2oDoc.xml");
      if (e2oDoc) {
        shape.e2oXml = e2oDoc.toString("utf8");
        const cNvPr = parseEmbeddedShapeCNvPr(shape.e2oXml);
        if (cNvPr) {
          shape.docPrId = cNvPr.id;
          shape.name = cNvPr.name;
        }
      }
    }
  }
  return shape;
}

function parseOfficeArtClientData(buffer, start, end) {
  const clientDataBySpid = new Map();
  let drawingOrder = 0;

  function walk(off, limit) {
    while (off + 8 <= limit) {
      const record = readOfficeArtRecord(buffer, off, limit);
      if (record.type === OFFICE_ART_CLIENT_DATA) {
        if (record.len >= 16) {
          const storedOrder = buffer.readUInt32LE(record.content);
          const spid = buffer.readUInt32LE(record.content + 12);
          clientDataBySpid.set(spid, { drawingOrder: storedOrder, rawOrder: drawingOrder });
          drawingOrder += 1;
        }
      } else if (record.ver === 0x0f) {
        walk(record.content, record.end);
      }
      off = record.end;
    }
  }

  walk(start, end);
  return clientDataBySpid;
}

function attachOfficeArtToShapeAnchors(shapeAnchors, officeArtShapes) {
  const matched = shapeAnchors.map((shape) => officeArtShapes.get(shape.lid)).filter(Boolean);
  const matchedSpids = matched.map((shape) => shape.spid).filter(Number.isInteger);
  const matchedOrders = matched.map((shape) => shape.drawingOrder).filter(Number.isInteger);
  const fallbackSpid = matchedSpids.length ? Math.min(...matchedSpids) : null;
  const firstDrawingOrder = matchedOrders.length ? Math.min(...matchedOrders) : 0;
  return shapeAnchors.map((shape) => {
    const officeArt = officeArtShapes.get(shape.lid);
    if (!officeArt) return shape;
    return {
      ...shape,
      officeArt: {
        spid: officeArt.spid,
        shapeType: officeArt.shapeType,
        fspFlags: officeArt.fspFlags,
        drawingOrder: officeArt.drawingOrder,
        relativeHeight: WPS_DRAWING_RELATIVE_HEIGHT_BASE + (officeArt.drawingOrder - firstDrawingOrder + 1) * WPS_DRAWING_RELATIVE_HEIGHT_STEP,
        docPrId: officeArt.docPrId,
        name: officeArt.name,
        description: officeArt.description,
        fallbackSpid,
        gfxData: officeArt.gfxData,
        e2oXml: officeArt.e2oXml,
        properties: officeArt.properties,
        tertiaryProperties: officeArt.tertiaryProperties,
      },
    };
  });
}

function readOfficeArtRecord(buffer, off, limit) {
  if (off + 8 > limit) {
    throw new Error("Invalid Word binary document: truncated OfficeArt record header");
  }
  const verInst = buffer.readUInt16LE(off);
  const len = buffer.readUInt32LE(off + 4);
  const end = off + 8 + len;
  if (end > limit) {
    throw new Error(`Invalid Word binary document: OfficeArt record 0x${buffer.readUInt16LE(off + 2).toString(16)} exceeds its container`);
  }
  return {
    off,
    ver: verInst & 0x000f,
    inst: verInst >>> 4,
    type: buffer.readUInt16LE(off + 2),
    len,
    content: off + 8,
    end,
  };
}

function parseOfficeArtFopt(buffer, record, label) {
  const fixedEnd = record.content + record.inst * 6;
  if (fixedEnd > record.end) {
    throw new Error(`Invalid Word binary document: ${label} property table exceeds record`);
  }

  const properties = [];
  let complexOff = fixedEnd;
  for (let i = 0; i < record.inst; i += 1) {
    const off = record.content + i * 6;
    const opid = buffer.readUInt16LE(off);
    const op = buffer.readUInt32LE(off + 2);
    const isComplex = (opid & 0x8000) !== 0;
    const property = {
      opid,
      pid: opid & 0x3fff,
      isBlip: (opid & 0x4000) !== 0,
      value: op,
      complexData: null,
    };
    if (isComplex) {
      if (complexOff + op > record.end) {
        throw new Error(`Invalid Word binary document: ${label} complex property exceeds record`);
      }
      property.complexData = buffer.subarray(complexOff, complexOff + op);
      complexOff += op;
    }
    properties.push(property);
  }
  return properties;
}

function extractOfficeArtGfxData(data) {
  const zipOffset = data.indexOf(Buffer.from([0x50, 0x4b, 0x03, 0x04]));
  if (zipOffset < 0) {
    throw new Error("Invalid Word binary document: OfficeArt gfxdata property does not contain an OOXML package");
  }
  return Buffer.from(data.subarray(zipOffset));
}

function parseZipEntries(zip) {
  const entries = new Map();
  let off = 0;
  while (off + 30 <= zip.length) {
    const signature = zip.readUInt32LE(off);
    if (signature !== 0x04034b50) {
      off += 1;
      continue;
    }
    const method = zip.readUInt16LE(off + 8);
    const compressedSize = zip.readUInt32LE(off + 18);
    const fileNameLength = zip.readUInt16LE(off + 26);
    const extraLength = zip.readUInt16LE(off + 28);
    const nameStart = off + 30;
    const dataStart = nameStart + fileNameLength + extraLength;
    const dataEnd = dataStart + compressedSize;
    if (dataEnd > zip.length) {
      throw new Error("Invalid Word binary document: embedded OfficeArt package entry exceeds package length");
    }
    const name = zip.subarray(nameStart, nameStart + fileNameLength).toString("utf8");
    const data = zip.subarray(dataStart, dataEnd);
    if (method === 0) {
      entries.set(name, Buffer.from(data));
    } else if (method === 8) {
      entries.set(name, inflateRawSync(data));
    } else {
      throw new Error(`Unimplemented embedded OfficeArt package compression method ${method}`);
    }
    off = dataEnd;
  }
  return entries;
}

function parseEmbeddedShapeCNvPr(xml) {
  const match = xml.match(/<wps:cNvPr\b[^>]*\bid="([^"]+)"[^>]*\bname="([^"]*)"/);
  if (!match) return null;
  const id = Number.parseInt(match[1], 10);
  if (!Number.isInteger(id)) {
    throw new Error(`Invalid Word binary document: embedded shape cNvPr id ${JSON.stringify(match[1])} is not an integer`);
  }
  return { id, name: unescapeXml(match[2]) };
}

function stripNullTerminator(value) {
  return value.replace(/\u0000+$/g, "");
}

function unescapeXml(value) {
  return value
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
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

function parseShapeAnchors(tableStream, fc, lcb, characterCount, label) {
  if (!lcb) return [];
  assertTableRange(tableStream, fc, lcb, label);
  const plcf = tableStream.subarray(fc, fc + lcb);
  // MS-DOC-SPEC/15 fcPlcSpaMom/lcbPlcSpaMom point to PlcfSpa. A PlcfSpa is
  // a PLC: (n + 1) CPs followed by n fixed-size Spa records. MS-DOC-SPEC/19
  // Spa is 26 bytes.
  const spaSize = 26;
  if (plcf.length < 4 || (plcf.length - 4) % (4 + spaSize) !== 0) {
    throw new Error(`Invalid Word binary document: ${label} length ${plcf.length} is not valid`);
  }
  const count = (plcf.length - 4) / (4 + spaSize);
  const cps = [];
  for (let i = 0; i <= count; i += 1) {
    const cp = plcf.readUInt32LE(i * 4);
    if (i < count && cp > characterCount) {
      throw new Error(`Out-of-spec ${label} CP ${cp} exceeds story character count ${characterCount}`);
    }
    if (i > 0 && cp < cps[i - 1]) {
      throw new Error(`Out-of-spec ${label} CPs are not sorted`);
    }
    cps.push(cp);
  }
  const shapes = [];
  let off = (count + 1) * 4;
  for (let i = 0; i < count; i += 1, off += spaSize) {
    const spa = plcf.subarray(off, off + spaSize);
    const flags = spa.readUInt16LE(20);
    const bx = (flags >> 1) & 0x03;
    const by = (flags >> 3) & 0x03;
    const wr = (flags >> 5) & 0x0f;
    const wrk = (flags >> 9) & 0x0f;
    const fRcaSimple = (flags >> 13) & 0x01;
    if (fRcaSimple !== 0) {
      throw new Error("Out-of-spec Spa.fRcaSimple must be zero");
    }
    shapes.push({
      cpStart: cps[i],
      cpEnd: cps[i + 1],
      lid: spa.readUInt32LE(0),
      xaLeft: spa.readInt32LE(4),
      yaTop: spa.readInt32LE(8),
      xaRight: spa.readInt32LE(12),
      yaBottom: spa.readInt32LE(16),
      bx,
      by,
      wr,
      wrk,
      fBelowText: ((flags >> 14) & 0x01) === 1,
      fAnchorLock: ((flags >> 15) & 0x01) === 1,
    });
  }
  return shapes;
}

function parseSttbfBkmk(sttbf) {
  const names = parseUnicodeSttbNoExtra(sttbf, "SttbfBkmk");
  const seen = new Set();
  for (const name of names) {
    if (name.length === 0 || name.length >= 40) {
      throw new Error(`Invalid Word bookmark name length ${name.length}`);
    }
    if (seen.has(name)) {
      throw new Error(`Out-of-spec Word bookmark duplicate name: ${name}`);
    }
    seen.add(name);
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

function extractParagraphPropertyEntries(wordDocument, tableStream, fib, documentCharacterCount, styles, pieces) {
  if (!fib.fcPapx || !fib.lcbPapx || fib.lcbPapx < 4) return [];
  if (fib.fcPapx + fib.lcbPapx > tableStream.length) return [];
  const plcf = tableStream.subarray(fib.fcPapx, fib.fcPapx + fib.lcbPapx);
  const binCount = (fib.lcbPapx - 4) / 8;
  if (binCount <= 0 || !Number.isInteger(binCount)) return [];
  const pageNumbers = [];
  for (let bi = 0; bi < binCount; bi += 1) {
    pageNumbers.push(plcf.readUInt32LE((binCount + 1) * 4 + bi * 4));
  }
  const allEntries = [];
  for (const pageNumber of pageNumbers) {
    if (pageNumber === 0) continue;
    const pageOffset = pageNumber * 512;
    if (pageOffset + PHE_SIZE > wordDocument.length) continue;
    const page = wordDocument.subarray(pageOffset, Math.min(pageOffset + 512, wordDocument.length));
    allEntries.push(...parsePapxEntries(page, styles, pieces, documentCharacterCount));
  }
  return allEntries;
}

function paragraphPropertiesForStory(story, allEntries) {
  const paragraphRanges = getParagraphRanges(story.rawText);
  return paragraphRanges.map((paragraphRange) => {
    const globalRange = { cpStart: story.cpStart + paragraphRange.cpStart, cpEnd: story.cpStart + paragraphRange.cpEnd };
    const direct = allEntries.find((entry) => rangesOverlap(entry, globalRange)) ?? null;
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
  const properties = buildParagraphPropertiesFromSprms(parsed);
  const paragraphPropertyChange = parseParagraphPropertyChange(data, parsed);
  if (paragraphPropertyChange) {
    properties.paragraphPropertyChange = paragraphPropertyChange;
  }
  return properties;
}

function buildParagraphPropertiesFromSprms(parsed) {
  return {
    istd: parsed.istd ?? null,
    styleId: null,
    lineSpacing: parsed.lineSpacing ?? null,
    alignment: parsed.alignment ?? null,
    leftIndentChars: parsed.leftIndentChars ?? null,
    rightIndentChars: parsed.rightIndentChars ?? null,
    firstLineIndentChars: parsed.firstLineIndentChars ?? null,
    firstLineIndentCharsSprm: parsed.firstLineIndentCharsSprm ?? null,
    leftIndent: parsed.leftIndent ?? null,
    rightIndent: parsed.rightIndent ?? null,
    firstLineIndent: parsed.firstLineIndent ?? null,
    firstLineIndentSprm: parsed.firstLineIndentSprm ?? null,
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
    outlineLevel: parsed.outlineLevel ?? null,
    tablePosition: parsed.tablePosition ?? null,
    tableNoAllowOverlap: parsed.tableNoAllowOverlap ?? null,
  };
}

function parseParagraphPropertyChange(data, parsed) {
  const wallRecords = scanSprmRecords(data, true)
    .filter((record) => record.sprm === 0x2664);
  const activeWall = wallRecords.at(-1);
  if (!activeWall) return null;
  if (activeWall.val[0] === 0) return null;
  if (activeWall.val[0] !== 1) {
    throw new Error(`Out-of-spec MS-DOC sprmPWall Bool8 value ${activeWall.val[0]}`);
  }

  const beforeWall = parseSprms(data.subarray(0, activeWall.offset), true);
  const mark = beforeWall.paragraphPropRMark ?? parsed.paragraphPropRMark;
  if (!mark?.fPropRMark) return null;

  return {
    // MS-DOC-SPEC/16 sprmPWall: SPRMs before sprmPWall specify the state
    // before paragraph-property revision marking was enabled.
    previous: buildParagraphPropertiesFromSprms(beforeWall),
    // MS-DOC-SPEC/16 sprmPPropRMark: PropRMark stores author and DTTM.
    mark,
  };
}

const STYLE_TYPE_PARAGRAPH = "paragraph";
const STYLE_TYPE_CHARACTER = "character";
const STYLE_TYPE_TABLE = "table";
const STYLE_TYPE_NUMBERING = "numbering";
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
  // stk (style kind) per MS-DOC StdfBase: 1=paragraph, 2=character, 3=table, 4=numbering
  if (sgc === 2) return STYLE_TYPE_CHARACTER;
  if (sgc === 3) return STYLE_TYPE_TABLE;
  if (sgc === 4) return STYLE_TYPE_NUMBERING;
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
  const styleSheetInfo = {
    // MS-DOC-SPEC/19 Stshif.ftcAsci/ftcFE/ftcOther/ftcBi are the
    // sprmCRgFtc* / sprmCFtcBi operands for default document formatting.
    ftcAsci: stsh.readInt16LE(14),
    ftcFE: stsh.readInt16LE(16),
    ftcOther: stsh.readInt16LE(18),
    ftcBi: stsh.readInt16LE(20),
    nVerBuiltInNamesWhenSaved: stsh.readUInt16LE(12),
  };
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
  // MS-DOC-SPEC/19 StdfBase.stk value 4 is a numbering style, not a
  // paragraph style. This converter does not emit OOXML numbering styles,
  // so exclude them from WPS paragraph/table/character style-id assignment.
  const styleIdStyles = nonNullStyles.filter((style) => style.type !== STYLE_TYPE_NUMBERING);
  const hasNumberingStyle = nonNullStyles.some((style) => style.type === STYLE_TYPE_NUMBERING);
  const styleIdOrder = hasNumberingStyle ? [
    ...styleIdStyles.filter((style) => style.sti === 0),
    ...styleIdStyles.filter((style) => style.sti >= 1 && style.sti <= 9),
    ...styleIdStyles.filter((style) => style.sti === 65),
    ...styleIdStyles.filter((style) => style.sti === 105),
    ...styleIdStyles.filter((style) => style.type === STYLE_TYPE_PARAGRAPH && style.sti !== 0 && !(style.sti >= 1 && style.sti <= 9)),
    ...styleIdStyles.filter((style) => style.type === STYLE_TYPE_TABLE && style.sti !== 105),
    ...styleIdStyles.filter((style) => style.type === STYLE_TYPE_CHARACTER && style.sti !== 65),
  ] : [
    ...styleIdStyles.filter((style) => style.sti === 0),
    ...[...promotedStyles].filter((style) => style.type !== STYLE_TYPE_NUMBERING),
    ...styleIdStyles.filter((style) => (
      style.sti !== 0
      && !defaults.has(style.sti)
      && style.type === STYLE_TYPE_PARAGRAPH
      && style.sti < 179
      && !promotedStyles.has(style)
    )),
    ...styleIdStyles.filter((style) => style.sti === 105),
    ...styleIdStyles.filter((style) => (
      style.sti !== 105
      && style.type === STYLE_TYPE_TABLE
      && style.sti < 179
    )),
    ...styleIdStyles.filter((style) => style.sti === 65),
    ...styleIdStyles.filter((style) => (
      style.sti !== 65
      && style.sti !== 105
      && style.type === STYLE_TYPE_CHARACTER
      && style.sti < 179
    )),
    ...styleIdStyles.filter((style) => (
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

  return { styles, latentLsd, stiMaxWhenSaved, styleSheetInfo };
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
    firstLineIndentCharsSprm: parsed.firstLineIndentCharsSprm ?? null,
    leftIndent: parsed.leftIndent ?? null,
    rightIndent: parsed.rightIndent ?? null,
    firstLineIndent: parsed.firstLineIndent ?? null,
    firstLineIndentSprm: parsed.firstLineIndentSprm ?? null,
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

function extractSections(wordDocument, tableStream, fib, bodyText = "") {
  if (!fib.lcbPlcfSed || fib.lcbPlcfSed < 16) return [];
  if (fib.fcPlcfSed + fib.lcbPlcfSed > tableStream.length) return [];

  const plcf = tableStream.subarray(fib.fcPlcfSed, fib.fcPlcfSed + fib.lcbPlcfSed);
  const sedSize = 12;
  const sectionCount = (fib.lcbPlcfSed - 4) / (4 + sedSize);
  if (!Number.isInteger(sectionCount) || sectionCount <= 0) return [];

  const sectionCps = [];
  for (let i = 0; i <= sectionCount; i += 1) {
    sectionCps.push(plcf.readUInt32LE(i * 4));
  }
  validatePlcfSedCps(sectionCps, bodyText);

  const sections = [];
  const sedStart = (sectionCount + 1) * 4;
  for (let i = 0; i < sectionCount; i += 1) {
    const cpStart = sectionCps[i];
    const cpEnd = sectionCps[i + 1];
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

function validatePlcfSedCps(cps, bodyText) {
  for (let i = 1; i < cps.length; i += 1) {
    if (cps[i] <= cps[i - 1]) {
      // MS-DOC-SPEC/18 PlcfSed: section CPs MUST NOT contain duplicates,
      // and each section range ends immediately before the next CP.
      throw new Error(`Out-of-spec PlcfSed CP array is not strictly ascending at index ${i}`);
    }
  }
  const lastCp = cps.at(-1);
  if (lastCp < bodyText.length) {
    // MS-DOC-SPEC/18 PlcfSed: the final CP does not start a section and
    // MUST be at or beyond the end of the main document.
    throw new Error("Out-of-spec PlcfSed final CP is before the end of the main document");
  }
  for (let i = 0; i < cps.length - 2; i += 1) {
    const cpEnd = cps[i + 1];
    if (cpEnd > bodyText.length || bodyText[cpEnd - 1] !== "\f") {
      // MS-DOC-SPEC/18 PlcfSed: every non-final section's last character
      // in the main-document text range MUST be end-of-section 0x0C.
      throw new Error(`Out-of-spec PlcfSed section ${i} does not end with an end-of-section character`);
    }
  }
}

function readSectionProperties(wordDocument, fcSepx) {
  if (fcSepx + 2 > wordDocument.length) {
    throw new Error("Invalid Word binary document: SEPX points outside WordDocument");
  }

  const cb = wordDocument.readUInt16LE(fcSepx);
  if (fcSepx + 2 + cb > wordDocument.length) {
    throw new Error("Invalid Word binary document: truncated SEPX");
  }

  return parseSectionSprms(wordDocument.subarray(fcSepx + 2, fcSepx + 2 + cb), { validateRequiredSectionProperties: true });
}

export function parseSectionSprms(grpprl, options = {}) {
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
  if (off !== grpprl.length) {
    throw new Error("Truncated section SPRM code at end of SEPX grpprl");
  }
  validateSectionProperties(props, options);
  return props;
}

function validateSectionProperties(props, options = {}) {
  validateSectionPageAndMarginProperties(props, options);
  if (props.docGridType != null && ![0, 1, 2, 3].includes(props.docGridType)) {
    // MS-DOC-SPEC/19 SClmOperand enumerates only clmUseDefault,
    // clmCharsAndLines, clmLinesOnly, and clmEnforceGrid.
    throw new Error(`Out-of-spec section document grid mode ${props.docGridType}`);
  }
  if (props.docGridCharSpace != null
    && (props.docGridCharSpace < -670925 || props.docGridCharSpace > 6488064)) {
    // MS-DOC-SPEC/16 sprmSDxtCharSpace MUST be in this range.
    throw new Error(`Out-of-spec section document grid character spacing ${props.docGridCharSpace}`);
  }
  if (props.docGridLinePitch != null
    && (props.docGridLinePitch < 1 || props.docGridLinePitch > 31680)) {
    // MS-DOC-SPEC/16 sprmSDyaLinePitch MUST be 1..31680 twips.
    throw new Error(`Out-of-spec section document grid line pitch ${props.docGridLinePitch}`);
  }
  if (props.docGridType === 1 || props.docGridType === 2 || props.docGridType === 3) {
    if (props.docGridLinePitch == null) {
      // MS-DOC-SPEC/16 sprmSDyaLinePitch: if the document grid is enabled
      // by sprmSClm, the section MUST specify the grid line height.
      throw new Error("Out-of-spec section document grid is enabled without sprmSDyaLinePitch");
    }
  }
}

function validateSectionPageAndMarginProperties(props, options = {}) {
  if (props.pageWidth != null && (props.pageWidth < 144 || props.pageWidth > 31680)) {
    // MS-DOC-SPEC/16 sprmSXaPage MUST be in [144, 31680] twips.
    throw new Error(`Out-of-spec section page width ${props.pageWidth}`);
  }
  if (props.pageHeight != null && (props.pageHeight < 144 || props.pageHeight > 31680)) {
    // MS-DOC-SPEC/16 sprmSYaPage MUST be in [144, 31680] twips.
    throw new Error(`Out-of-spec section page height ${props.pageHeight}`);
  }
  for (const [key, label] of [["marginLeft", "left"], ["marginRight", "right"], ["gutterMargin", "gutter"]]) {
    if (props[key] != null && (props[key] < 0 || props[key] > 31680)) {
      // MS-DOC-SPEC/19 XAS_nonNeg values MUST be <= 31680 twips.
      throw new Error(`Out-of-spec section ${label} margin ${props[key]}`);
    }
  }
  for (const [key, label] of [["headerMargin", "header"], ["footerMargin", "footer"]]) {
    if (props[key] != null && props[key] > 31680) {
      // MS-DOC-SPEC/19 YAS_nonNeg values MUST be <= 31680 twips.
      throw new Error(`Out-of-spec section ${label} margin ${props[key]}`);
    }
  }
  for (const [key, label] of [["marginTop", "top"], ["marginBottom", "bottom"]]) {
    if (props[key] != null && (props[key] < -31665 || props[key] > 31665)) {
      // MS-DOC-SPEC/16 sprmSDyaTop/sprmSDyaBottom MUST be [-31665, 31665].
      throw new Error(`Out-of-spec section ${label} margin ${props[key]}`);
    }
  }
  if (options.validateRequiredSectionProperties) {
    for (const [key, sprmName] of [
      ["marginLeft", "sprmSDxaLeft"],
      ["marginRight", "sprmSDxaRight"],
      ["marginTop", "sprmSDyaTop"],
      ["marginBottom", "sprmSDyaBottom"],
    ]) {
      if (props[key] == null) {
        // MS-DOC-SPEC/16 requires each section to explicitly specify these
        // implementation-dependent margins; fail instead of emitting guessed
        // OOXML page margins.
        throw new Error(`Out-of-spec section is missing required ${sprmName}`);
      }
    }
    Object.defineProperty(props, "_msDocRequiredSectionPropertiesValidated", {
      value: true,
      enumerable: false,
      configurable: true,
    });
  }
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
      // MS-DOC-SPEC/16 sprmSFRTLGutter is a Bool8. Preserve presence so an
      // explicit false value can be emitted instead of disappearing as default.
      props.rtlGutterSpecified = true;
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
    case 0x522f: {
      const raw = val.readUInt16LE(0);
      if ((raw & 0xff00) !== 0) throw new Error("Out-of-spec SPgbPropOperand reserved byte must be zero");
      const applyTo = raw & 0x0007;
      const pageDepth = (raw >> 3) & 0x0003;
      const offsetFrom = (raw >> 5) & 0x0007;
      if (applyTo > 2 || pageDepth > 1 || offsetFrom > 1) throw new Error(`Out-of-spec SPgbPropOperand 0x${raw.toString(16)}`);
      props.pageBorderProperties = {
        display: ["allPages", "firstPage", "notFirstPage"][applyTo],
        zOrder: ["front", "back"][pageDepth],
        offsetFrom: ["text", "page"][offsetFrom],
      };
      break;
    }
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
      // MS-DOC-SPEC/16 sprmSDyaTop is a signed YAS; negative values
      // specify fixed margins rather than minimum margins.
      props.marginTop = val.readInt16LE(0);
      break;
    case 0x9024:
      // MS-DOC-SPEC/16 sprmSDyaBottom is a signed YAS; negative values
      // specify fixed margins rather than minimum margins.
      props.marginBottom = val.readInt16LE(0);
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
  let border;
  if (brc.length === 4) {
    border = parseBrc80Raw(brc[0], brc[1], brc[2], brc[3] & 0x1f);
  } else {
    const brcType = brc[5];
    const valName = BRC_TYPE_NAMES[brcType];
    if (!valName) throw new Error(`Out-of-spec page border BrcType ${brcType}`);
    border = {
      val: valName,
      color: colorRefToHex(brc.subarray(0, 4)),
      sz: String(brc[4]),
      space: String(brc[6] & 0x1f),
      shadow: (brc[7] & 0x20) !== 0,
      frame: (brc[7] & 0x40) !== 0,
    };
  }
  props.pageBorders ??= {};
  props.pageBorders[side] = border.val === "none" ? { style: "none" } : { ...border, style: border.val };
}

function extractCharacterRuns(wordDocument, tableStream, fib, documentText, pieces, styles) {
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

  const documentCharacterCount = documentText.length;
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
      if (cpStart == null || cpEnd == null || cpStart >= documentCharacterCount || cpEnd <= 0) continue;

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
        cpEnd: Math.min(documentCharacterCount, cpEnd),
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
    cellWidthAssignments: rowProperty?.cellWidthAssignments ?? null,
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

  // MS-DOC-SPEC/19 TDefTable stores explicit table-cell boundary positions.
  // Preserve even very narrow parsed columns; WPS exports them as OOXML grid
  // columns instead of merging nearby boundaries.
  return [...positionSet].sort((a, b) => a - b);
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
    applyCellWidthAssignments(row);
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
    if (flags.horzMerge === 2 || flags.horzMerge === 3) cell.hMerge = "restart";
    else if (flags.horzMerge === 1) cell.hMerge = "continue";
    const textDirectionByFlow = { 0: null, 1: "tbRl", 3: "btLr", 4: "lrTbV", 5: "tbRlV" };
    cell.textDirection = textDirectionByFlow[flags.textFlow];
    cell.fitText = flags.fFitText === true;
    cell.hideMark = flags.fHideMark === true;
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

function applyCellWidthAssignments(row) {
  if (!row.cellWidthAssignments || row.cellWidthAssignments.length === 0) return;
  for (const assignment of row.cellWidthAssignments) {
    const start = Math.max(0, assignment.itcFirst);
    const end = Math.min(row.cells.length, assignment.itcLim);
    for (let ci = start; ci < end; ci += 1) {
      const cell = row.cells[ci];
      if (!cell) continue;
      cell.preferredWidth = assignment.width;
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
  if (overlapping.length === 0) {
    // MS-DOC-SPEC/19 PapxFkp.rgfc can point at an end-of-row mark, and the
    // corresponding PapxInFkp contains the table properties for the row whose
    // end-of-row mark is at that offset. In simple/non-complex files, a final
    // body table row can have that row-mark PAPX at the first CP after the
    // body subdocument; accept only that adjacent one-character row mark and
    // only when its parsed TDefTable column count exactly matches the row.
    return rowProperties.findLast((entry) =>
      entry.cpStart === cpEnd &&
      entry.cpEnd === cpEnd + 1 &&
      entry.columns?.length === cellCount + 1
    ) ?? null;
  }

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
        cellWidthAssignments: info.cellWidthAssignments,
        tableWidth: info.tableWidth,
        tableWidthType: info.tableWidthType,
        tableIndent: info.tableIndent,
        tableJustification: info.tableJustification,
        tableStyleIndex: info.tableStyleIndex,
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
  let cellWidthAssignments = [];
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
              const horzMerge = bits & 0x0003;
              const textFlow = (bits >> 2) & 0x0007;
              const vertMerge = (bits >> 5) & 0x0003;
              if (![0, 1, 3, 4, 5].includes(textFlow)) {
                throw new Error(`Out-of-spec TCGRF TextFlow ${textFlow}`);
              }
              cellFlags.push({
                horzMerge,
                textFlow,
                vertMerge,
                bFirstMerged: horzMerge === 2 || horzMerge === 3,
                bMerged: horzMerge === 1,
                bVertical: textFlow !== 0,
                bBackward: textFlow === 3,
                bRotateFont: textFlow === 4 || textFlow === 5,
                bVertMerge: vertMerge === 1,
                bVertRestart: vertMerge === 3,
                nVertAlign: (bits & 0x0180) >> 7,
                fFitText: (bits & 0x1000) !== 0,
                // MS-DOC-SPEC/19 TCGRF.fNoWrap is bit 13 of the TC80 flag word.
                fNoWrap: (bits & 0x2000) !== 0,
                fHideMark: (bits & 0x4000) !== 0,
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
    } else if (sprm === 0xD635) {
      const cb = data[off + 2];
      size = cb + 1;
      if (off + 3 + cb <= data.length) {
        cellWidthAssignments.push(parseTableCellWidthOperand(data.subarray(off + 3, off + 3 + cb)));
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
    cellWidthAssignments,
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

function parseTableCellWidthOperand(payload) {
  // MS-DOC-SPEC/19 TableCellWidthOperand: cb MUST be 5; after cb the
  // operand stores ItcFirstLim plus an FtsWWidth_TablePart.
  if (!payload || payload.length !== 5) {
    throw new Error(`Invalid sprmTCellWidth operand length ${payload?.length ?? 0}`);
  }
  const itcFirst = payload[0];
  const itcLim = payload[1];
  if (itcLim <= itcFirst || itcLim > 64) {
    throw new Error(`Out-of-spec sprmTCellWidth cell range ${itcFirst}-${itcLim}`);
  }
  const width = parseTablePartWidth(payload.subarray(2, 5), "sprmTCellWidth");
  return { itcFirst, itcLim, width };
}

function parseTablePartWidth(data, label) {
  if (!data || data.length < 3) {
    throw new Error(`Truncated ${label} FtsWWidth_TablePart`);
  }
  const widthType = data[0];
  const widthValue = data.readUInt16LE(1);
  if (widthType === 0) {
    return null;
  }
  if (widthType === 1) {
    if (widthValue !== 0) {
      throw new Error(`Invalid ${label} auto width ${widthValue}`);
    }
    return { type: "auto", value: 0 };
  }
  if (widthType === 2) {
    if (widthValue > 5000) {
      throw new Error(`Out-of-spec ${label} percent width ${widthValue}`);
    }
    return { type: "pct", value: widthValue };
  }
  if (widthType === 3) {
    if (widthValue > 31680) {
      throw new Error(`Out-of-spec ${label} dxa width ${widthValue}`);
    }
    return { type: "dxa", value: widthValue };
  }
  throw new Error(`Out-of-spec ${label} width type ${widthType}`);
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

  // MS-DOC-SPEC/16 sprmTWidthIndent is the table's preferred leading
  // indent, so when it is explicitly present preserve it as OOXML tblInd.
  // If it is absent, use MS-DOC-SPEC/19 TDefTableOperand.rgdxaCenter[0],
  // the parsed logical-left table edge, as the table indentation source.
  const geometryIndent = gridPositions?.[0] ? { width: gridPositions[0], type: "dxa" } : null;
  return (indent?.width !== 0 ? indent : null) ?? geometryIndent;
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
