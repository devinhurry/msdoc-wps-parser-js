import { writeFile } from "node:fs/promises";
import { deflateRawSync, inflateRawSync } from "node:zlib";
import { lcidToBcp47 } from "./lcid.js";
import { readWpsFile } from "./wps.js";
import { STI_NAMES } from "./word-binary.js";

// Counter for generating unique w14:paraId values (8 hex chars)
let paraIdCounter = 0;

function nextParaId() {
  const id = (paraIdCounter++ >>> 0).toString(16).toUpperCase().padStart(8, "0");
  return id;
}

function buildContentTypesXml({ footerCount = 13, headerCount = 0, includeNumbering = true, includeNotes = true } = {}) {
  const includeFooters = footerCount > 0;
  const includeHeaders = headerCount > 0;
  if (!includeFooters && !includeHeaders && !includeNumbering) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`;
  }
  const parts = [
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`,
    `  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`,
    `  <Default Extension="xml" ContentType="application/xml"/>`,
    `  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`,
    `  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`,
    `  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>`,
    `  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`,
    `  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`,
    `  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>`,
    `  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>`,
  ];
  if (includeNotes) {
    parts.push(
      `  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>`,
      `  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>`,
    );
  }
  if (includeFooters) {
    for (let i = 1; i <= footerCount; i++) {
      parts.push(`  <Override PartName="/word/footer${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`);
    }
  }
  if (includeHeaders) {
    for (let i = 1; i <= headerCount; i++) {
      parts.push(`  <Override PartName="/word/header${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`);
    }
  }
  if (includeNumbering) {
    parts.push(`  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`);
  }
  parts.push(`  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`);
  parts.push(`</Types>`);
  return parts.join("\n");
}

const ROOT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

function buildDocumentRelsXml({ footerCount = 13, headerCount = 0, includeNumbering = true, footerRelationshipStartId = 5, includeNotes = true, headerFooterRelationships = null } = {}) {
  const includeFooters = footerCount > 0;
  const includeHeaders = headerCount > 0;
  if (!includeFooters && !includeHeaders && !includeNumbering) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  }
  const parts = [
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`,
    `  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
    `  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`,
  ];
  if (includeNotes) {
    parts.push(
      `  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>`,
      `  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>`,
    );
  }
  if (headerFooterRelationships?.length) {
    for (const rel of headerFooterRelationships) {
      parts.push(`  <Relationship Id="${rel.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${rel.type}" Target="${rel.type}${rel.index}.xml"/>`);
    }
  } else if (includeFooters) {
    for (let i = 1; i <= footerCount; i++) {
      parts.push(`  <Relationship Id="rId${footerRelationshipStartId + i - 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer${i}.xml"/>`);
    }
  }
  const lastHeaderFooterRid = headerFooterRelationships?.length
    ? Math.max(...headerFooterRelationships.map((rel) => Number(rel.id.slice(3))))
    : footerRelationshipStartId + footerCount + headerCount - 1;
  const themeRid = lastHeaderFooterRid + 1;
  parts.push(`  <Relationship Id="rId${themeRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`);
  if (includeNumbering) {
    parts.push(`  <Relationship Id="rId${themeRid + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`);
  }
  parts.push(`  <Relationship Id="rId${themeRid + (includeNumbering ? 2 : 1)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>`);
  parts.push(`</Relationships>`);
  return parts.join("\n");
}

const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>msdoc-wps-parser</Application>
</Properties>`;

const CUSTOM_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>`;


function buildSettingsXml(wpsDocument = {}) {
  const sections = wpsDocument.sections ?? [];
  const sectionDocGridType = sections.find((section) => section?.properties?.docGridType != null)?.properties?.docGridType ?? null;
  const hasEastAsianGrid = sectionDocGridType === 1 || sectionDocGridType === 2;
  const hasGridType1 = sectionDocGridType === 1;
  const hasGridType2 = sectionDocGridType === 2;
  const noHeaderLineGrid = hasLineGridWithoutHeaderSubdocument(wpsDocument);
  const hasNegativeGridCharSpace = sections.some((section) => section?.properties?.docGridCharSpace < 0);
	  const hasExplicitZeroGridCharSpace = sections.some((section) => section?.properties?.docGridType === 2 && section?.properties?.docGridCharSpace === 0);
	  const hasVbaProject = (wpsDocument.streams ?? []).some((stream) => stream?.name === "_VBA_PROJECT");
	  const readOnlyEastAsianProfile = hasGridType1 && hasVbaProject;
  const compactGridHeaderSubdocument = hasCompactGridHeaderSubdocument(wpsDocument, sections);
	  const dop = wpsDocument.dop;
  const hasMainShapeAnchors = (wpsDocument.fib?.lcbPlcSpaMom ?? 0) > 0;
  const hasHeaderShapeAnchors = (wpsDocument.fib?.lcbPlcSpaHdr ?? 0) > 0;

  const defaultTabStop = wpsDocument.defaultTabStop;
  if (!Number.isInteger(defaultTabStop) || defaultTabStop <= 0) {
    throw new Error("Invalid WPS document: missing parsed default tab stop");
  }

  const parts = [
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
    `<w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData" mc:Ignorable="w14">`,
  ];

  if (wpsDocument.fib?.fReadOnlyRecommended) {
    // MS-DOC-SPEC/15 FibBase.fReadOnlyRecommended maps to OOXML
    // w:writeProtection/@w:recommended. OpenXML places writeProtection
    // before zoom in CT_Settings order.
    parts.push(`<w:writeProtection w:recommended="1"/>`);
  }

  // MS-DOC-SPEC/15 fc/lcbPlcSpaMom and fc/lcbPlcSpaHdr identify parsed
  // body/header shape anchors; MS-DOC-SPEC/17 DopBase.fRevMarking identifies
  // active revision tracking. WPS normalizes saved zoom to its 130% bucket
  // for these parsed visual/revision states as well as headerless or negative
  // East-Asian grid sections.
  const useWpsExpandedZoom = noHeaderLineGrid
    || hasNegativeGridCharSpace
    || hasMainShapeAnchors
    || hasHeaderShapeAnchors
    || dop?.fRevMarking;
  parts.push(`<w:zoom w:percent="${useWpsExpandedZoom ? 130 : 127}"/>`);

	  // MS-DOC-SPEC/17 §DopBase.l fEmbedFonts and §Dop97.D fSubsetFonts:
  //   <w:embedTrueTypeFonts/>   iff DopBase.fEmbedFonts (ECMA-376 §17.8.3.8)
  //   <w:saveSubsetFonts/>      iff DopBase.fEmbedFonts AND Dop97.fSubsetFonts
  //   (ECMA-376 §17.8.3.15). WPS exports saveSubsetFonts only alongside embedTrueTypeFonts.
  if (dop?.fEmbedFonts) {
    parts.push(`<w:embedTrueTypeFonts/>`);
    if (dop?.fSubsetFonts) {
      parts.push(`<w:saveSubsetFonts/>`);
    }
  }

  // MS-DOC-SPEC/17 §DopBase.b fMirrorMargins → OOXML mirrorMargins (ECMA-376 §17.15.1.52).
  if (dop?.fMirrorMargins) {
    parts.push(`<w:mirrorMargins w:val="1"/>`);
  }

  const pageBorderIncludes = dop?.pageBorderIncludes;
  if (!pageBorderIncludes) {
    throw new Error("Invalid WPS document: missing parsed Dop97 page-border include flags");
  }
  parts.push(
    `<w:bordersDoNotSurroundHeader w:val="${pageBorderIncludes.header ? 0 : 1}"/>`,
    `<w:bordersDoNotSurroundFooter w:val="${pageBorderIncludes.footer ? 0 : 1}"/>`,
  );

  // MS-DOC-SPEC/17 §Dop2002 grfFmtFilter: suggested formatting filter.
  // WPS only emits this element for East Asian grid documents.
  if (hasEastAsianGrid && dop?.grfFmtFilter == null) {
    throw new Error("Invalid WPS document: missing parsed grfFmtFilter for East Asian grid settings");
  }
  if (hasEastAsianGrid && dop.grfFmtFilter !== 0x5024) {
    const filterHex = dop.grfFmtFilter.toString(16).toUpperCase().padStart(4, "0");
    parts.push(`<w:stylePaneFormatFilter w:val="${filterHex}" w:allStyles="1" w:customStyles="0" w:latentStyles="0" w:stylesInUse="0" w:headingStyles="0" w:numberingStyles="0" w:tableStyles="0" w:directFormattingOnRuns="1" w:directFormattingOnParagraphs="1" w:directFormattingOnNumbering="1" w:directFormattingOnTables="1" w:clearFormatting="1" w:top3HeadingStyles="1" w:visibleStyles="0" w:alternateStyleNames="0"/>`);
  }

  // Open XML CT_Settings order places trackRevisions before
  // documentProtection and defaultTabStop.
  // MS-DOC-SPEC/17 DopBase.fRevMarking maps to OOXML trackRevisions.
  if (dop?.fRevMarking) {
    parts.push(`<w:trackRevisions w:val="1"/>`);
  }

  parts.push(
    `<w:documentProtection${readOnlyEastAsianProfile ? ` w:edit="readOnly"` : ""} w:enforcement="0"/>`,
    `<w:defaultTabStop w:val="${defaultTabStop}"/>`,
    `<w:hyphenationZone w:val="${(dop?.dxaHotZ && dop.dxaHotZ > 0) ? dop.dxaHotZ : 360}"/>`,
  );

  // MS-DOC-SPEC/17 DopBase.fAutoHyphen maps to OOXML autoHyphenation.
  if (dop?.fAutoHyphen) {
    parts.push(`<w:autoHyphenation w:val="1"/>`);
  }

  // MS-DOC-SPEC/17 DopBase.cConsecHypLim maps to OOXML consecutiveHyphenLimit.
  if (dop?.cConsecHypLim && dop.cConsecHypLim > 0) {
    parts.push(`<w:consecutiveHyphenLimit w:val="${dop.cConsecHypLim}"/>`);
  }

  // MS-DOC-SPEC/17 DopBase.fHyphCapitals is the opposite of OOXML doNotHyphenateCaps.
  // When fHyphCapitals is false, words in all caps are NOT hyphenated.
  if (dop?.fHyphCapitals === false) {
    parts.push(`<w:doNotHyphenateCaps w:val="1"/>`);
  }

  // MS-DOC-SPEC/17 DopBase.fLinkStyles maps to OOXML linkStyles.
  if (dop?.fLinkStyles) {
    parts.push(`<w:linkStyles w:val="1"/>`);
  }

  // MS-DOC-SPEC/17 DopBase.fFacingPages maps to OOXML evenAndOddHeaders.
  // WPS also uses this as a condition for evenAndOddHeaders emission,
  // falling back to the docGridType heuristic when fFacingPages is absent.
  if (dop?.fFacingPages ?? (sectionDocGridType !== 2 && !hasNegativeGridCharSpace)) {
    parts.push(`<w:evenAndOddHeaders w:val="1"/>`);
  }

  // MS-DOC-SPEC/17 §Dogrid: drawing grid settings from DOP
  const dogrid = dop?.dogrid;
  if (!dogrid) {
    throw new Error("Invalid WPS document: missing parsed Dogrid for settings drawing grid");
  }
  // MS-DOC-SPEC/17 Dogrid.dxaGrid maps to drawingGridHorizontalSpacing;
  // its default is 180, so only an explicit non-default value needs XML.
  if (dogrid.dxaGrid !== 180) {
    parts.push(`<w:drawingGridHorizontalSpacing w:val="${dogrid.dxaGrid}"/>`);
  }

  if (hasEastAsianGrid) {
    parts.push(`<w:drawingGridVerticalSpacing w:val="${dogrid.dyaGrid}"/>`);
  }

  parts.push(
    `<w:displayHorizontalDrawingGridEvery w:val="${dogrid.dxGridDisplay}"/>`,
    `<w:displayVerticalDrawingGridEvery w:val="${dogrid.dyGridDisplay}"/>`,
  );

  const typography = dop?.typography;
  if (hasEastAsianGrid && !typography) {
    throw new Error("Invalid WPS document: missing parsed DopTypography for East Asian settings");
  }
  if (hasEastAsianGrid && !typography.fKerningPunct) {
    parts.push(`<w:noPunctuationKerning w:val="1"/>`);
  }

  const characterSpacingControl = hasEastAsianGrid
    ? ["doNotCompress", "compressPunctuation", "compressPunctuationAndJapaneseKana"][typography.iJustification]
    : "doNotCompress";
  if (!characterSpacingControl) {
    throw new Error(`Out-of-spec DopTypography iJustification ${typography.iJustification}`);
  }
  parts.push(`<w:characterSpacingControl w:val="${characterSpacingControl}"/>`);
  if (hasHeaderShapeAnchors) {
    parts.push(buildHeaderShapeDefaultsXml());
  }

  const xmlValidation = dop?.xmlValidation;
  if (hasGridType2 && !xmlValidation) {
    throw new Error("Invalid WPS document: missing parsed Dop2002 XML validation settings");
  }
  if (hasGridType2 && !xmlValidation.fValidateXML) {
    parts.push(`<w:doNotValidateAgainstSchema/>`);
  }
  if (hasGridType2 && !xmlValidation.fShowXMLErrors) {
    parts.push(`<w:doNotDemarcateInvalidXml/>`);
  }

  // WPS emits explicit note separator references for ordinary documents and
  // for grid-2 documents whose parsed section grid has zero character spacing
  // and whose FIB has no header/footer subdocument.
	  if (!hasEastAsianGrid || (hasGridType2 && hasExplicitZeroGridCharSpace && (noHeaderLineGrid || compactGridHeaderSubdocument))) {
    parts.push(
      `<w:footnotePr><w:footnote w:id="0"/><w:footnote w:id="1"/></w:footnotePr>`,
      `<w:endnotePr><w:endnote w:id="0"/><w:endnote w:id="1"/></w:endnotePr>`,
    );
  }

  parts.push(`<w:compat>`);
  if (hasEastAsianGrid) {
    // WPS exports these compat flags for East Asian grid layouts.
    parts.push(
      `<w:spaceForUL/>`,
      `<w:balanceSingleByteDoubleByteWidth/>`,
      `<w:doNotLeaveBackslashAlone/>`,
    );
    if (dop?.compatibility?.ulTrailSpace) {
      parts.push(`<w:ulTrailSpace/>`);
    }
    parts.push(`<w:doNotExpandShiftReturn/>`);
    if (hasGridType2 || hasNegativeGridCharSpace) {
      parts.push(`<w:adjustLineHeightInTable/>`);
    }
    parts.push(`<w:useFELayout/>`);
  } else {
    parts.push(`<w:doNotLeaveBackslashAlone/>`, `<w:ulTrailSpace/>`, `<w:useFELayout/>`);
  }
  parts.push(
    `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>`,
    `<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>`,
    `<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>`,
    `<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>`,
    `</w:compat>`,
    `<m:mathPr><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="0"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr>`,
  );

	  // MS-DOC-SPEC/17 does not document a direct binary-to-OOXML mapping for
  // <w:uiCompat97To2003/>; WPS emits it for East-Asian grid documents that
  // were saved in Word97 compat mode (DopBase.fWord97Compat / FIB.fWord97Saved).
  // Empirically (per WPS sample exports) emission is gated by:
  //   hasEastAsianGrid && not (no-header-grid) && not (neg grid char space)
  //   && not (compact header-only subdocument)
  // Additionally, while revisions are tracked (DopBase.fRevMarking → trackRevisions),
  // WPS does NOT emit uiCompat97To2003. This anti-emission is the only
  // observed WPS discriminator between otherwise identical profile documents.
  if (hasEastAsianGrid && !noHeaderLineGrid && !hasNegativeGridCharSpace && !compactGridHeaderSubdocument && !dop?.fRevMarking) {
    parts.push(`<w:uiCompat97To2003/>`);
  }

  parts.push(
    `<w:themeFontLang w:val="en-US" w:eastAsia="zh-CN"/>`,
    `<w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>`,
  );

  if (hasEastAsianGrid) {
    parts.push(`<w:doNotIncludeSubdocsInStats/>`);
  }
  if (hasMainShapeAnchors) {
    parts.push(buildMainShapeDefaultsXml());
  }

  parts.push(`<mc:AlternateContent><mc:Choice Requires="wpsCustomData"><wpsCustomData:typoFeatureVersion val="0"/></mc:Choice></mc:AlternateContent>`);

  parts.push(`</w:settings>`);
  return parts.join("");
}

function buildMainShapeDefaultsXml() {
  // MS-DOC-SPEC/15 fcPlcSpaMom/lcbPlcSpaMom identify PlcfSpa shape anchors
  // in the Main Document. OOXML shapeDefaults stores VML defaults for the
  // body story; its idmap data is 1.
  return `<w:shapeDefaults><o:shapedefaults/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout></w:shapeDefaults>`;
}

function buildHeaderShapeDefaultsXml() {
  // MS-DOC-SPEC/15 fcPlcSpaHdr/lcbPlcSpaHdr identify PlcfSpa shape anchors
  // in the Header Document. OOXML hdrShapeDefaults stores VML defaults for
  // header/footer stories; its idmap data is 2.
  return `<w:hdrShapeDefaults><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout></w:hdrShapeDefaults>`;
}

const FOOTNOTES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotes>`;

const ENDNOTES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote></w:endnotes>`;

const THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游明朝"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>`;




function collectParagraphListIds(wpsDocument = {}) {
  return new Set(
    (wpsDocument.paragraphProperties ?? [])
      .map((properties) => properties?.listId)
      .filter((listId) => listId != null),
  );
}

function hasHeaderFooterSubdocument(wpsDocument = {}) {
  return (wpsDocument.fib?.characterCounts?.headers ?? 0) > 0;
}

function hasLineGridWithoutHeaderSubdocument(wpsDocument = {}) {
  return !hasHeaderFooterSubdocument(wpsDocument)
    && (wpsDocument.sections ?? []).some((section) => section?.properties?.docGridType === 2);
}

function hasCompactGridHeaderSubdocument(wpsDocument = {}, sections = []) {
  const headerText = wpsDocument.subdocuments?.headers?.rawText ?? "";
  return sections.some((section) => section?.properties?.docGridType === 2)
    && headerText.length > 0
    && /^[\r]*$/.test(headerText)
    && !wpsDocument.plcfHdd?.cpArray?.length;
}

function shouldEmitNumberingXml(wpsDocument = {}) {
  return [...collectParagraphListIds(wpsDocument)].some((listId) => listId > 0);
}
function createNumberingXml(wpsDocument = {}) {
  const paragraphProperties = wpsDocument.paragraphProperties ?? [];
  const listData = wpsDocument.listData ?? { lstfList: [], lfoList: [] };
  const listIds = [...new Set(
    paragraphProperties
      .map((p) => p?.listId)
      .filter((id) => id != null && id > 0)
  )].sort((a, b) => a - b);

  // Always include the default abstractNum (abstractNumId="0") even if no paragraphs use it
  const needNumbering = listIds.length > 0 || listData.lstfList.length > 0;
  if (!needNumbering) return '';

  const parts = [];
  parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  const ns = 'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData" mc:Ignorable="w14 wp14"';
  parts.push('<w:numbering ' + ns + '>');

  // Map LFO index → LSTF for abstractNum assignment
  const lfoToLstf = listData.lfoList.map(lfo => {
    const lstf = listData.lstfList.find(l => l.lsid === lfo.lsid);
    return lstf || null;
  });

  // MS-DOC-SPEC/15 §LSTF.rgistdPara may be 0x0FFF (no linked style) even
  // when a paragraph style references this numbering. Scan style definitions
  // to build a reverse lookup: (abstractNumId, ilvl) → styleId.
  const pStyleByLevel = {};
  for (const style of wpsDocument.styles ?? []) {
    if (style?.listId != null && style.listId > 0 && style.listLevel != null && style.styleId != null) {
      const lfoIdx = style.listId - 1;
      const lstfForLfo = lfoToLstf[lfoIdx];
      if (lstfForLfo) {
        const aIdx = listData.lstfList.indexOf(lstfForLfo);
        if (aIdx >= 0) {
          pStyleByLevel[aIdx + ':' + style.listLevel] = style.styleId;
        }
      }
    }
  }

  // Generate abstractNum for each LSTF
  const fontTable = wpsDocument.fontTable ?? [];
  for (let aIdx = 0; aIdx < listData.lstfList.length; aIdx++) {
    const lstf = listData.lstfList[aIdx];
    parts.push(`<w:abstractNum w:abstractNumId="${aIdx}">`);
    if (lstf.lsid) parts.push(`<w:nsid w:val="${lstf.lsid.toString(16).toUpperCase().padStart(8, '0')}"/>`);
    parts.push(`<w:multiLevelType w:val="${lstf.fSimpleList ? "singleLevel" : "multilevel"}"/>`);
    if (lstf.tplc) parts.push(`<w:tmpl w:val="${lstf.tplc.toString(16).toUpperCase().padStart(8, '0')}"/>`);

    for (const lvl of lstf.lvlList) {
      const t = lstf.fHybrid && lvl.fTentative ? '1' : '0';
      parts.push(`<w:lvl w:ilvl="${lvl.ilvl}" w:tentative="${t}">`);
      parts.push(`<w:start w:val="${lvl.iStartAt}"/>`);
      parts.push(`<w:numFmt w:val="${nfcToNumFmtXml(lvl.nfc)}"/>`);
      if (lvl.fNoRestart) {
        parts.push(`<w:lvlRestart w:val="${lvl.ilvlRestartLim}"/>`);
      }
      if (lvl.fLegal) {
        parts.push("<w:isLgl/>");
      }
      // Suffix (ixchFollow): omit for "tab" (default), emit for "space" or "nothing"
      const suff = ({0:"tab",1:"space",2:"nothing"})[lvl.ixchFollow] ?? "tab";
      if (suff !== "tab") parts.push(`<w:suff w:val="${suff}"/>`);
      // pStyle from LSTF rgistdPara (MS-DOC-SPEC §LSTF: rgistdPara[level] is istd of linked style, or 0x0FFF if none)
      const istdPara = lstf.rgistdPara[lvl.ilvl];
      if (istdPara != null && istdPara !== 0 && istdPara < 0x0FFF) {
        parts.push(`<w:pStyle w:val="${istdPara}"/>`);
      } else {
        // Fallback: derive pStyle from style definitions that reference this numbering.
        const styleKey = aIdx + ':' + lvl.ilvl;
        const pStyleId = pStyleByLevel[styleKey];
        if (pStyleId) {
          parts.push(`<w:pStyle w:val="${pStyleId}"/>`);
        }
      }
      const lvlText = levelTextToXml(lvl);
      // For numbered levels with empty text, use current-level placeholder.
      // Bullet (nfc=23) and none (nfc=255) levels keep empty text.
      const effectiveLvlText = lvlText || (lvl.nfc !== 23 && lvl.nfc !== 255 ? "%" + (lvl.ilvl + 1) : "");
      parts.push(`<w:lvlText w:val="${escapeXml(effectiveLvlText)}"/>`);
      // Justification
      const jcMap = { 0: "left", 1: "center", 2: "right" };
      parts.push(`<w:lvlJc w:val="${jcMap[lvl.jc] || "left"}"/>`);
      // Paragraph properties: tabs, jc, indents from grpprlPapx
      const hasTabs = lvl.tabs && lvl.tabs.length > 0;
      const hasIndents = lvl.leftIndent != null || lvl.firstLineIndent != null || lvl.rightIndent != null;
      const hasPapxJc = lvl.papxAlignment != null;
      if (hasTabs || hasIndents || hasPapxJc) {
        parts.push('<w:pPr>');
        if (hasTabs) {
          const tabsXml = lvl.tabs
            .map((tab) => `<w:tab w:val="${tab.alignment}"${tab.leader ? ` w:leader="${tab.leader}"` : ""} w:pos="${tab.position}"/>`)
            .join("");
          parts.push(`<w:tabs>${tabsXml}</w:tabs>`);
        }
        if (hasIndents) {
          const indParts = [];
          if (lvl.leftIndent != null) indParts.push(`w:left="${lvl.leftIndent}"`);
          if (lvl.firstLineIndent != null) {
            if (lvl.firstLineIndent < 0) indParts.push(`w:hanging="${Math.abs(lvl.firstLineIndent)}"`);
            else indParts.push(`w:firstLine="${lvl.firstLineIndent}"`);
          }
          if (lvl.rightIndent != null) indParts.push(`w:right="${lvl.rightIndent}"`);
          if (indParts.length) parts.push(`<w:ind ${indParts.join(" ")}/>`);
        }
        if (hasPapxJc) parts.push(`<w:jc w:val="${lvl.papxAlignment}"/>`);
        parts.push('</w:pPr>');
      }
      // Run properties from grpprlChpx: rFonts, spacing, w, sz, szCs, color
      const fontAttrs = [];
      if (lvl.fontHint != null) fontAttrs.push(`w:hint="${lvl.fontHint}"`);
      if (lvl.fontAscii != null) fontAttrs.push(`w:ascii="${escapeXml(resolveFontName(fontTable, lvl.fontAscii))}"`);
      if (lvl.fontHAnsi != null) fontAttrs.push(`w:hAnsi="${escapeXml(resolveFontName(fontTable, lvl.fontHAnsi))}"`);
      if (lvl.fontEastAsia != null) fontAttrs.push(`w:eastAsia="${escapeXml(resolveFontName(fontTable, lvl.fontEastAsia))}"`);
      if (lvl.fontCs != null) fontAttrs.push(`w:cs="${escapeXml(resolveFontName(fontTable, lvl.fontCs))}"`);
      const hasFontRPr = fontAttrs.length > 0;
      const hasSizeRPr = lvl.fontSizeCs != null;
	      const hasColorRPr = lvl.textColor != null;
	      const hasSpacingRPr = lvl.charSpacing != null;
	      const hasCharWidthRPr = lvl.charWidth != null;
	      const hasFontSizeRPr = lvl.fontSize != null;
      const hasBoldRPr = lvl.bold != null;
      const hasItalicRPr = lvl.italic != null;
      const hasToggleRPr = lvl.boldCs != null || lvl.italicCs != null || lvl.caps != null || lvl.smallCaps != null || lvl.strike != null || lvl.dstrike != null || lvl.vanish != null
        || lvl.outline != null || lvl.shadow != null || lvl.imprint != null || lvl.emboss != null || lvl.noProof != null || lvl.webHidden != null || lvl.specVanish != null;
		      const hasPositionRPr = lvl.textPosition != null;
      const hasVerticalAlignRPr = lvl.verticalAlign != null;
		      const hasKernRPr = lvl.kern != null;
	      const hasUnderlineRPr = lvl.underline != null;
      const hasBorderRPr = lvl.border != null;
      const underlineColorAttr = lvl.underlineColor != null && lvl.underlineColor !== "auto" ? ` w:color="${lvl.underlineColor}"` : "";
		      if (hasFontRPr || hasSizeRPr || hasColorRPr || hasSpacingRPr || hasCharWidthRPr || hasFontSizeRPr || hasBoldRPr || hasItalicRPr || hasToggleRPr || hasPositionRPr || hasVerticalAlignRPr || hasKernRPr || hasUnderlineRPr || hasBorderRPr) {
	        parts.push('<w:rPr>');
	        if (hasFontRPr) parts.push(`<w:rFonts ${fontAttrs.join(" ")}/>`);
        if (lvl.bold === true) parts.push(`<w:b/>`);
        else if (lvl.bold === false) parts.push(`<w:b w:val="0"/>`);
        if (lvl.boldCs === true) parts.push(`<w:bCs/>`);
        else if (lvl.boldCs === false) parts.push(`<w:bCs w:val="0"/>`);
        if (lvl.italic === true) parts.push(`<w:i/>`);
        else if (lvl.italic === false) parts.push(`<w:i w:val="0"/>`);
        if (lvl.italicCs === true) parts.push(`<w:iCs/>`);
        else if (lvl.italicCs === false) parts.push(`<w:iCs w:val="0"/>`);
        appendToggleRunProperty(parts, "caps", lvl.caps);
        appendToggleRunProperty(parts, "smallCaps", lvl.smallCaps);
        appendToggleRunProperty(parts, "strike", lvl.strike);
        appendToggleRunProperty(parts, "dstrike", lvl.dstrike);
        appendToggleRunProperty(parts, "vanish", lvl.vanish);
        appendToggleRunProperty(parts, "outline", lvl.outline);
        appendToggleRunProperty(parts, "shadow", lvl.shadow);
        appendToggleRunProperty(parts, "imprint", lvl.imprint);
        appendToggleRunProperty(parts, "emboss", lvl.emboss);
        appendToggleRunProperty(parts, "noProof", lvl.noProof);
        appendToggleRunProperty(parts, "webHidden", lvl.webHidden);
        appendToggleRunProperty(parts, "specVanish", lvl.specVanish);
	        if (hasSpacingRPr) parts.push(`<w:spacing w:val="${lvl.charSpacing}"/>`);
	        if (hasKernRPr) parts.push(`<w:kern w:val="${lvl.kern}"/>`);
	        if (hasPositionRPr) parts.push(`<w:position w:val="${lvl.textPosition}"/>`);
        if (hasVerticalAlignRPr) parts.push(`<w:vertAlign w:val="${lvl.verticalAlign}"/>`);
        if (hasUnderlineRPr) parts.push(`<w:u w:val="${lvl.underline ? (lvl.underlineStyle ?? "single") : "none"}"${underlineColorAttr}/>`);
        if (hasBorderRPr) parts.push(`<w:bdr w:val="${lvl.border.val}" w:color="${lvl.border.color}" w:sz="${lvl.border.sz}" w:space="${lvl.border.space}"/>`);
	        if (hasCharWidthRPr) parts.push(`<w:w w:val="${lvl.charWidth}"/>`);
	        if (hasFontSizeRPr) parts.push(`<w:sz w:val="${lvl.fontSize}"/>`);
	        if (hasSizeRPr) parts.push(`<w:szCs w:val="${lvl.fontSizeCs}"/>`);
	        if (hasColorRPr) parts.push(`<w:color w:val="${lvl.textColor}"/>`);
        if (hasPositionRPr && lvl.textPosition === 0) parts.push(`<w:vertAlign w:val="baseline"/>`);
	        parts.push('</w:rPr>');
	      }
      parts.push('</w:lvl>');
    }
    parts.push('</w:abstractNum>');
  }

  // Generate num entries: each LFO gets a numId = LFO index + 1
  for (const lfo of listData.lfoList) {
    const aIdx = listData.lstfList.findIndex(l => l.lsid === lfo.lsid);
    if (aIdx >= 0) {
      parts.push(`<w:num w:numId="${lfo.index + 1}">`);
      parts.push(`<w:abstractNumId w:val="${aIdx}"/>`);
      for (const lfolvl of lfo.lfolvlList ?? []) {
        if (!lfolvl.fStartAt) continue;
        parts.push(`<w:lvlOverride w:ilvl="${lfolvl.iLvl}"><w:startOverride w:val="${lfolvl.iStartAt}"/></w:lvlOverride>`);
      }
      parts.push("</w:num>");
    }
  }
  // Also generate num entries for any listIds not covered by LFOs
  for (const listId of listIds) {
    if (listData.lfoList.every(l => l.index + 1 !== listId)) {
      parts.push(`<w:num w:numId="${listId}"><w:abstractNumId w:val="0"/></w:num>`);
    }
  }

  parts.push('</w:numbering>');
  return parts.join('\n');
}

function nfcToNumFmtXml(nfc) {
  // MS-DOC-SPEC/19 LVLF.nfc is an MSONFC from MS-OSHARED §2.2.1.3.
  // Reuse the section page-number mapping below so numbering never falls back
  // to decimal for valid East Asian formats such as 0x25/0x27.
  // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/57c8f7d4-1426-44e0-b58d-6fa1a5a81659
  return sectionPageNumberFormatToXml(nfc);
}

function levelTextToXml(lvl) {
  // MS-DOC-SPEC/19 LVL.xst stores placeholder code units, but LVLF.rgbxchNums
  // is the authoritative one-based location array. The code unit at that
  // location is the zero-based level number; OOXML uses %1..%9.
  const placeholderByOffset = new Map();
  for (let j = 0; j < (lvl.rgbxchNums ?? []).length; j += 1) {
    const oneBasedOffset = lvl.rgbxchNums[j];
    if (!oneBasedOffset) break;
    const offset = oneBasedOffset - 1;
    if (offset < 0 || offset >= lvl.xst.length) {
      throw new Error(`Out-of-spec LVL rgbxchNums offset ${oneBasedOffset}`);
    }
    const ch = lvl.xst.charCodeAt(offset);
    if (ch > lvl.ilvl || ch > 8) {
      throw new Error(`Out-of-spec LVL placeholder ${ch} at rgbxchNums[${j}]`);
    }
    placeholderByOffset.set(offset, "%" + (ch + 1));
  }
  let text = "";
  for (let k = 0; k < lvl.xst.length; k += 1) {
    if (placeholderByOffset.has(k)) {
      text += placeholderByOffset.get(k);
      continue;
    }
    const ch = lvl.xst.charCodeAt(k);
    if (ch <= 8) continue;
    text += lvl.xst[k];
  }
  return text;
}

function createFooterReferencePlan(wpsDocument = {}, sections = []) {
  if (!hasHeaderFooterSubdocument(wpsDocument)) {
    return { footerCount: 0, bySection: new Map() };
  }
  if (wpsDocument.plcfHdd?.cpArray?.length) {
    throw new Error("PlcfHdd-backed footer references are not implemented yet");
  }

  if (sections.some((section) => section?.properties?.docGridType != null)) {
    const ccpHdd = wpsDocument.fib?.characterCounts?.headers ?? 0;
    const onlySection = sections[0]?.properties ?? {};
    if (sections.length === 1 && onlySection.titlePg && ccpHdd === 7) {
      // MS-DOC-SPEC/13 PlcfHdd story order includes even footer, odd footer,
      // and first footer. WPS omits PlcfHdd in compact CR-only grid exports;
      // a 7-CP header subdocument is exactly those three footer stories plus
      // the final CP.
      return {
        footerCount: 3,
        footerRelationshipStartId: 3,
        includeNotes: false,
        bySection: new Map([[0, { defaultFooterId: "rId3", evenFooterId: "rId4", firstFooterId: "rId5" }]]),
      };
    }
    const headerText = wpsDocument.subdocuments?.headers?.rawText ?? "";
	    if (sections.length === 1 && (onlySection.docGridType === 1 || onlySection.docGridType === 2) && headerText.includes("PAGE")) {
      // MS-DOC-SPEC/13: the header document contains header and footer
      // stories. Some WPS grid exports omit PlcfHdd but keep compact
      // PAGE-field footer stories in ccpHdd. DopBase.fFacingPages specifies
      // whether even and odd page headers/footers are distinct.
      if (wpsDocument.dop?.fFacingPages) {
        return {
          footerCount: 2,
          footerRelationshipStartId: 3,
          includeNotes: false,
          bySection: new Map([[0, { defaultFooterId: "rId3", evenFooterId: "rId4" }]]),
        };
      }
      return {
        footerCount: 1,
        footerRelationshipStartId: 3,
        includeNotes: false,
	        bySection: new Map([[0, { defaultFooterId: "rId3" }]]),
	      };
	    }
    if (sections.length === 1 && (onlySection.docGridType === 1 || onlySection.docGridType === 2) && /[^\r]/.test(headerText)) {
      // MS-DOC-SPEC/13: when different even/odd headers are not enabled, the
      // odd header and odd footer stories are the default header/footer.
      return {
        footerCount: 1,
        headerCount: 1,
        footerRelationshipStartId: 4,
        includeNotes: false,
        headerFooterRelationships: [
          { id: "rId3", type: "header", index: 1 },
          { id: "rId4", type: "footer", index: 1 },
        ],
        bySection: new Map([[0, { defaultHeaderId: "rId3", defaultFooterId: "rId4" }]]),
      };
    }
    if (sections.length > 1 && ccpHdd > 0 && /^[\r]*$/.test(headerText)) {
      return createCompactGridHeaderFooterPlan(sections);
    }
	    return { footerCount: 0, bySection: new Map() };
	  }

  const ccpHdd = wpsDocument.fib?.characterCounts?.headers ?? 0;
  if (ccpHdd <= 1 || ccpHdd % 2 !== 1) {
    throw new Error(`Out-of-spec compact WPS header subdocument length ${ccpHdd}`);
  }
  const footerCount = (ccpHdd - 1) / 2;
  const bySection = new Map();
  const seenFooterMargins = new Set();
  let nextFooterRid = 5;
  let assigned = 0;
  let activeRepeatedFooterMargin = null;

  const isLandscape = (properties = {}) => (properties.pageWidth ?? 11906) > (properties.pageHeight ?? 16838);
  const addRefs = (sectionIndex, types) => {
    if (assigned >= footerCount) return;
    const refs = bySection.get(sectionIndex) ?? {};
    for (const type of types) {
      if (assigned >= footerCount) break;
      refs[`${type}FooterId`] = `rId${nextFooterRid++}`;
      assigned += 1;
    }
    bySection.set(sectionIndex, refs);
  };

  for (let i = 0; i < sections.length && assigned < footerCount; i += 1) {
    const current = sections[i]?.properties ?? {};
    const previous = sections[i - 1]?.properties ?? null;
    const next = sections[i + 1]?.properties ?? null;
    const currentLandscape = isLandscape(current);
    const previousLandscape = previous ? isLandscape(previous) : false;
    const nextLandscape = next ? isLandscape(next) : false;
    const footerMargin = current.footerMargin ?? 720;
    if (activeRepeatedFooterMargin !== footerMargin) {
      activeRepeatedFooterMargin = null;
    }

    // MS-DOC-SPEC/13 says empty header/footer stories inherit from the previous
    // section. WPS omits PlcfHdd in these exports but keeps a compact CR-only
    // header subdocument; each emitted footer story accounts for two CRs.
    if (i === 0) {
      addRefs(i, ["default", "even"]);
    } else if (current.pageNumberStart != null && current.pageNumberStart > 0) {
      addRefs(i, ["default"]);
    } else if (currentLandscape && current.columnCount > 1) {
      if ((previous?.footerMargin ?? null) !== 0) addRefs(i, ["even"]);
    } else if (currentLandscape && previousLandscape && previous?.footerMargin !== footerMargin && footerMargin !== 720) {
      addRefs(i, ["default", "even"]);
    } else if (!currentLandscape && previousLandscape) {
      addRefs(i, ["default", "even"]);
      if (next?.footerMargin === footerMargin) {
        activeRepeatedFooterMargin = footerMargin;
      }
    } else if (footerMargin === 0) {
      if (i === sections.length - 1) {
        addRefs(i, ["default", "even"]);
      } else if ((previous?.footerMargin ?? 0) !== 0 && nextLandscape) {
        addRefs(i, ["even"]);
      }
    } else {
      const adjacentSameFooterMargin = previous?.footerMargin === footerMargin || next?.footerMargin === footerMargin;
      if (i === 1 && sections[0]?.properties?.pageNumberStart === 1) {
        addRefs(i, ["default", "even"]);
      } else if (activeRepeatedFooterMargin === footerMargin) {
        addRefs(i, ["default", "even"]);
      } else if (!seenFooterMargins.has(footerMargin) && adjacentSameFooterMargin) {
        activeRepeatedFooterMargin = footerMargin;
        addRefs(i, ["default", "even"]);
      }
    }

    seenFooterMargins.add(footerMargin);
  }

  if (assigned !== footerCount) {
    throw new Error(`Unable to allocate compact WPS footer references: expected ${footerCount}, assigned ${assigned}`);
  }
	  return { footerCount, bySection };
	}

function createCompactGridHeaderFooterPlan(sections = []) {
  // PlcfHdd is the authoritative MS-DOC story boundary table. Some WPS
  // East Asian grid exports omit it while keeping a CR-only header subdocument;
  // use the same compact path as above, following MS-DOC's section story order
  // for the references indicated by section title-page flags and geometry.
  const bySection = new Map();
  const relationships = [];
  let nextRid = 5;
  let footerCount = 0;
  let headerCount = 0;
  let passedTitlePages = false;
  let emittedPostTitleDefault = false;
  let emittedLandscapeHeader = false;

  const setRef = (sectionIndex, key, value) => {
    const refs = bySection.get(sectionIndex) ?? {};
    refs[key] = value;
    bySection.set(sectionIndex, refs);
  };
  const addFooter = (sectionIndex, key) => {
    footerCount += 1;
    const id = `rId${nextRid++}`;
    relationships.push({ id, type: "footer", index: footerCount });
    setRef(sectionIndex, key, id);
  };
  const addHeader = (sectionIndex, key) => {
    headerCount += 1;
    const id = `rId${nextRid++}`;
    relationships.push({ id, type: "header", index: headerCount });
    setRef(sectionIndex, key, id);
  };
  const isLandscape = (section) => {
    const props = section?.properties ?? {};
    return (props.pageWidth ?? 11906) > (props.pageHeight ?? 16838);
  };

  for (let i = 0; i < sections.length; i += 1) {
    const props = sections[i]?.properties ?? {};
    const previous = sections[i - 1]?.properties ?? null;
    const titlePg = props.titlePg === true;
    const previousTitlePg = previous?.titlePg === true;
    const landscape = isLandscape(sections[i]);
    const nextLandscape = isLandscape(sections[i + 1]);

    if (titlePg) {
      if (passedTitlePages) {
        addFooter(i, "defaultFooterId");
        addFooter(i, "firstFooterId");
      }
      passedTitlePages = true;
      continue;
    }

    if (previousTitlePg && !emittedPostTitleDefault) {
      addHeader(i, "defaultHeaderId");
      addFooter(i, "defaultFooterId");
      emittedPostTitleDefault = true;
      continue;
    }

    if (landscape && !nextLandscape && !emittedLandscapeHeader && sections[i + 1]?.properties?.docGridLinePitch !== props.docGridLinePitch) {
      addHeader(i, "defaultHeaderId");
      emittedLandscapeHeader = true;
      continue;
    }

    if (i === sections.length - 1) {
      addHeader(i, "defaultHeaderId");
      addFooter(i, "defaultFooterId");
    }
  }

  return {
    footerCount,
    headerCount,
    footerRelationshipStartId: 5,
    headerFooterRelationships: relationships,
    includeNotes: true,
    bySection,
  };
}

function getFooterIds(sectionIndex, documentOptions = {}) {
  return documentOptions.footerIdsBySection?.get(sectionIndex) ?? {};
}

export async function convertWpsToDocxFile(inputPath, outputPath, options = {}) {
  const wpsDocument = await readWpsFile(inputPath);
  const docx = wpsToDocxBuffer(wpsDocument, {
    title: options.title ?? inputPath,
    creator: options.creator ?? "msdoc-wps-parser",
  });
  await writeFile(outputPath, docx);
  return {
    outputPath,
    paragraphCount: wpsDocument.paragraphs.length,
    byteLength: docx.length,
  };
}

export function wpsToDocxBuffer(wpsDocument, options = {}) {
  _revisionIdCounter = 0;
  const paragraphProperties = wpsDocument.paragraphProperties ?? [];
  const characterProperties = wpsDocument.characterProperties ?? [];
  const styles = wpsDocument.styles ?? [];
  const fontTable = wpsDocument.fontTable ?? [];
  const sections = wpsDocument.sections ?? [];
  const tableRows = wpsDocument.tableRows ?? [];
	  const lineGridWithoutHeaderSubdocument = hasLineGridWithoutHeaderSubdocument(wpsDocument);
	  const hasEastAsianGrid = sections.some((section) => section?.properties?.docGridType === 1 || section?.properties?.docGridType === 2);
	  const hasLineGridWithoutCharSpace = sections.some((section) => section?.properties?.docGridType === 2 && section?.properties?.docGridCharSpace == null);
	  const hasNegativeGridCharSpace = sections.some((section) => section?.properties?.docGridCharSpace < 0);
  const compactGridHeaderSubdocument = hasCompactGridHeaderSubdocument(wpsDocument, sections);
  const includeNumbering = shouldEmitNumberingXml(wpsDocument);
  const footerReferencePlan = createFooterReferencePlan(wpsDocument, sections);
  const bodyText = wpsDocument.bodyText ?? wpsDocument.text ?? "";
  const bookmarks = createDocumentBookmarks(wpsDocument, sections, bodyText);
  const documentOptions = {
    lineGridWithoutHeaderSubdocument,
    resolveCharUnitIndentFromFont: hasLineGridWithoutCharSpace,
    suppressComplexScriptSize: !hasEastAsianGrid,
    hasEastAsianGrid,
    fib: wpsDocument.fib,
    dop: wpsDocument.dop,
    shapeAnchors: wpsDocument.shapeAnchors ?? [],
    sections,
    hasIncompleteEastAsianGrid: sections.some((section) => (
      (section?.properties?.docGridType === 1 || section?.properties?.docGridType === 2)
      && section?.properties?.docGridCharSpace == null
    )),
    suppressEastAsianParagraphControls: sections.some((section) => section?.properties?.docGridType === 1) && !hasNegativeGridCharSpace,
    emitNilCellBorders: sections.some((section) => section?.properties?.docGridType === 1) || lineGridWithoutHeaderSubdocument,
    emitUnderlineHighlight: !hasNegativeGridCharSpace,
    suppressNegativeGridTableCellParagraphControls: hasNegativeGridCharSpace,
	    emitNegativeGridParagraphMarkBaseline: hasNegativeGridCharSpace,
    suppressNegativeGridInlineComplexScriptToggles: hasNegativeGridCharSpace,
    suppressComplexScriptToggles: compactGridHeaderSubdocument,
    emitExtendedCharacterToggles: compactGridHeaderSubdocument,
    bookmarks,
    revisionAuthors: wpsDocument.revisionAuthors ?? ["Unknown"],
	    normalTableStyleId: styles.find((style) => style?.sti === 105)?.styleId ?? TABLE_STYLE_ID,
	    tableGridStyleId: styles.find((style) => style?.sti === 154)?.styleId ?? null,
	    footerIdsBySection: footerReferencePlan.bySection,
	  };
  const documentXml = createDocumentXml(bodyText, paragraphProperties, characterProperties, fontTable, sections, tableRows, documentOptions);
  const stylesXml = createStylesXml(styles, fontTable, wpsDocument);
  const fontTableXml = createFontTableXml(fontTable);
  const now = new Date().toISOString();
  const coreXml = createCoreXml({
    title: options.title ?? "Converted WPS document",
    creator: options.creator ?? "msdoc-wps-parser",
    created: options.created ?? now,
    modified: options.modified ?? now,
  });

  const zipEntries = [
	    { name: "[Content_Types].xml", data: buildContentTypesXml({ footerCount: footerReferencePlan.footerCount, headerCount: footerReferencePlan.headerCount ?? 0, includeNumbering, includeNotes: footerReferencePlan.includeNotes ?? true }) },
    { name: "_rels/", data: "" },
    { name: "_rels/.rels", data: ROOT_RELS_XML },
    { name: "docProps/", data: "" },
    { name: "docProps/app.xml", data: APP_XML },
    { name: "docProps/core.xml", data: coreXml },
    { name: "docProps/custom.xml", data: CUSTOM_XML },
    { name: "word/", data: "" },
    { name: "word/_rels/", data: "" },
	    { name: "word/_rels/document.xml.rels", data: buildDocumentRelsXml({
	      footerCount: footerReferencePlan.footerCount,
	      headerCount: footerReferencePlan.headerCount ?? 0,
	      includeNumbering,
	      footerRelationshipStartId: footerReferencePlan.footerRelationshipStartId ?? 5,
	      includeNotes: footerReferencePlan.includeNotes ?? true,
	      headerFooterRelationships: footerReferencePlan.headerFooterRelationships ?? null,
	    }) },
    { name: "word/document.xml", data: documentXml },
    { name: "word/fontTable.xml", data: fontTableXml },
  ];
  if (footerReferencePlan.includeNotes ?? true) {
    zipEntries.push({ name: "word/endnotes.xml", data: ENDNOTES_XML });
  }
	  if (footerReferencePlan.footerCount > 0) {
	    for (let i = 1; i <= footerReferencePlan.footerCount; i += 1) {
	      zipEntries.push({ name: "word/footer" + i + ".xml", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>' });
	    }
	  }
  if ((footerReferencePlan.headerCount ?? 0) > 0) {
    for (let i = 1; i <= footerReferencePlan.headerCount; i += 1) {
      zipEntries.push({ name: "word/header" + i + ".xml", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>' });
    }
  }
  if (footerReferencePlan.includeNotes ?? true) {
    zipEntries.push({ name: "word/footnotes.xml", data: FOOTNOTES_XML });
  }
  if (includeNumbering) {
    zipEntries.push({ name: "word/numbering.xml", data: createNumberingXml(wpsDocument) });
  }
  zipEntries.push({ name: "word/settings.xml", data: buildSettingsXml(wpsDocument) });
  zipEntries.push({ name: "word/styles.xml", data: stylesXml });
  zipEntries.push({ name: "word/theme/", data: "" });
  zipEntries.push({ name: "word/theme/theme1.xml", data: THEME_XML });
  return createZip(zipEntries);
}

function createDocumentBookmarks(wpsDocument = {}, sections = [], bodyText = "") {
  const bookmarks = [...(wpsDocument.bookmarks ?? [])];
  const cp = resolveGoBackBookmarkCp(wpsDocument, sections, bodyText);
  if (cp == null) return bookmarks;
  // MS-DOC-SPEC/15 fcWss/lcbWss stores Selsf, the last main-document
  // selection. MS-DOC-SPEC/19 SttbfBkmk defines names beginning with "_"
  // (0x005F) as hidden bookmarks. WPS' DOCX export serializes that UI
  // return marker as the next hidden bookmark named _GoBack.
  bookmarks.push({ id: bookmarks.length, name: "_GoBack", cpStart: cp, cpEnd: cp });
  return bookmarks;
}

function resolveGoBackBookmarkCp(wpsDocument = {}, sections = [], bodyText = "") {
  const selection = wpsDocument.lastSelection;
  if (!selection?.fIns) return null;
  if (selection.cpFirst !== selection.cpLim) {
    throw new Error("Cannot emit _GoBack bookmark from a non-collapsed Selsf selection");
  }
  const mainTextLimit = sections.length
    ? Math.max(...sections.map((section) => section?.cpEnd ?? 0))
    : bodyText.length;
  if (selection.cpFirst < 0 || selection.cpFirst > mainTextLimit) {
    // MS-DOC-SPEC/19 Selsf CPs are document text-piece positions. PlcfSed
    // section limits define the main-document CP extent; a selection beyond
    // that range belongs to another story and has no document.xml bookmark.
    return null;
  }

  const hasEastAsianGrid = sections.some((section) => section?.properties?.docGridType != null);
  if (hasEastAsianGrid) return 0;
  if (selection.cpFirst === 0 || isTrailingBodyInsertionPoint(bodyText, selection.cpFirst)) {
    return selection.cpFirst;
  }
  return 0;
}

function isTrailingBodyInsertionPoint(bodyText, cp) {
  if (cp === bodyText.length) return true;
  if (cp < 0 || cp >= bodyText.length) return false;
  return /^[\r\x07\x0c]*$/.test(bodyText.slice(cp));
}

function createDocumentXml(rawText, paragraphProperties = [], characterProperties = [], fontTable = [], sections = [], tables = [], documentOptions = {}) {
  documentOptions.bodyTextLength = rawText.length;
  const paragraphs = splitWordParagraphs(rawText, sections);
  for (const paragraph of paragraphs) {
    paragraph.bodyTextLength = rawText.length;
  }
  const allSections = sections;
  const finalSection = sections.at(-1)?.properties;

  const tableMap = buildTableMap(tables);

  const bodyParts = [];
  let tableIdx = 0;
  const sortedTables = [...tables].sort((a, b) => a.cpStart - b.cpStart);
  const emittedSectionIndices = new Set();

  for (let pi = 0; pi < paragraphs.length; pi++) {
    const paragraph = paragraphs[pi];

    while (tableIdx < sortedTables.length && sortedTables[tableIdx].cpStart <= paragraph.cpStart) {
      const table = sortedTables[tableIdx];
      if (table.cpStart < paragraph.cpEnd && !table._generated) {
        bodyParts.push(tableToXml(table, rawText, paragraphProperties, paragraphs, characterProperties, fontTable, tables.indexOf(table), sections, documentOptions));
        table._generated = true;

        // After emitting a table, check for section breaks that fall within or right after this table
        for (let si = 0; si < allSections.length; si++) {
          if (emittedSectionIndices.has(si)) continue;
          const sec = allSections[si];
          // Section break falls inside the table body; boundary paragraphs are handled below.
          if (sec.cpEnd > table.cpStart && sec.cpEnd <= table.cpEnd) {
            // Emit a paragraph with just the section break
            const secIdx = allSections.findIndex((bs) => bs === sec);
            const footerIds = getFooterIds(secIdx >= 0 ? secIdx : si, documentOptions);
            const sectionXml = buildSectionPropertiesXml(sec.properties, { ...footerIds, sectionIndex: si });
            bodyParts.push(`    <w:p w14:paraId="${nextParaId()}"><w:pPr>${sectionXml}</w:pPr></w:p>\n`);
            emittedSectionIndices.add(si);
          }
        }
      }
      tableIdx++;
    }

    if (isInsideTable(paragraph.cpStart, paragraph.cpEnd, tableMap)) {
      continue;
    }
    if (isSuppressibleFinalEmptyParagraph(paragraph, paragraphProperties[pi], documentOptions)) {
      continue;
    }
    if (paragraphs[pi - 1]?.manualPageBreak === true && paragraph.text.length === 0 && paragraph.delimiter === "\r") {
      continue;
    }

    // Check for section break at this paragraph
    const secIdx = allSections.findIndex((section) => section.cpEnd === paragraph.cpEnd);
    // Section CPs are tracked in the full document text space, not the trimmed body subdocument.
    const paragraphSection = allSections.find(
      (section) => paragraph.cpStart >= section.cpStart && paragraph.cpStart < section.cpEnd,
    )?.properties ?? null;
    if (secIdx >= 0) emittedSectionIndices.add(secIdx);
    const sectionProperties = secIdx >= 0 && secIdx !== allSections.length - 1 ? allSections[secIdx].properties : null;
    const result = paragraphToXml(
      paragraph,
      paragraphProperties[pi],
      characterProperties,
      fontTable,
      paragraph.cpStart,
      sectionProperties,
      paragraphSection,
      secIdx,
      null,
      documentOptions,
    );
    bodyParts.push(result.xml);
  }

  while (tableIdx < sortedTables.length) {
    const table = sortedTables[tableIdx];
    if (!table._generated) {
      bodyParts.push(tableToXml(table, rawText, paragraphProperties, paragraphs, characterProperties, fontTable, tables.indexOf(table), sections, documentOptions));
    }
    tableIdx++;
  }

  const body = bodyParts.join("");
  const finalSectionIdx = allSections.length - 1;
  const finalFooterIds = getFooterIds(finalSectionIdx, documentOptions);
  const finalSectionXml = buildSectionPropertiesXml(finalSection, { ...finalFooterIds, final: true, sectionIndex: finalSectionIdx });
  const hasBodyOrHeaderShapes = (documentOptions.fib?.lcbPlcSpaMom ?? 0) > 0
    || (documentOptions.fib?.lcbPlcSpaHdr ?? 0) > 0;
  // MS-DOC-SPEC/15 FibRgFcLcb97.fc/lcbPlcSpaMom and fc/lcbPlcSpaHdr
  // identify body/header shape-anchor PLCs, and MS-DOC-SPEC/17
  // DopBase.fRevMarking records active revision tracking. WPS emits the
  // OOXML document background for parsed grid-2 documents only while those
  // structures do not take over document-level visual state.
  const backgroundXml = !hasBodyOrHeaderShapes
    && !documentOptions.dop?.fRevMarking
    && sections.some((section) => section?.properties?.docGridType === 2 && section?.properties?.docGridCharSpace != null)
    ? `  <w:background w:color="FFFFFF"/>\n`
    : "";

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData" mc:Ignorable="w14 w15 wp14">
${backgroundXml}  <w:body>
${body}
    ${finalSectionXml}
  </w:body>
</w:document>`;
}

function buildTableMap(tables) {
  const ranges = [];
  for (const table of tables) {
    ranges.push({ start: table.cpStart, end: table.cpEnd });
  }
  return ranges;
}

function isInsideTable(cpStart, cpEnd, tableMap) {
  for (const range of tableMap) {
    if (cpStart >= range.start && cpStart < range.end) return true;
  }
  return false;
}

function isSuppressibleFinalEmptyParagraph(paragraph, properties, documentOptions = {}) {
  // MS-DOC-SPEC/16 stores paragraph/cell/section marks in the CP stream.
  // The final body-level sectPr is emitted separately, so a trailing empty
  // mark with no direct properties and no bookmark has no DOCX paragraph body.
  return paragraph.text.length === 0
    && paragraph.cpEnd === paragraph.bodyTextLength
    && !hasMeaningfulParagraphProperties(properties)
    && !hasBookmarkAtParagraphBoundary(documentOptions.bookmarks ?? [], paragraph);
}

function hasMeaningfulParagraphProperties(properties) {
  if (!properties) return false;
  return Object.entries(properties).some(([key, value]) => {
    if (key === "istd") return false;
    if (key === "inTable" && value === false) return false;
    if (value == null) return false;
    if (Array.isArray(value) && value.length === 0) return false;
    return true;
  });
}

function hasBookmarkAtParagraphBoundary(bookmarks, paragraph) {
  return bookmarks.some((bookmark) => (
    bookmark.cpStart >= paragraph.cpStart && bookmark.cpStart <= paragraph.cpEnd
  ) || (
    bookmark.cpEnd >= paragraph.cpStart && bookmark.cpEnd <= paragraph.cpEnd
  ));
}


const TABLE_STYLE_ID = "6";
const TABLE_BORDER_SIDES = ["top", "left", "bottom", "right", "insideH", "insideV"];
const TABLE_DOCX_PROPS = {
  jc: "center",
  layout: "fixed",
  tblW: "auto",
  cellMar: { top: 0, left: 0, bottom: 0, right: 0 },
};

function resolveTableStyleId(table, sections = [], documentOptions = {}) {
  if (table?.tableStyleId) {
    return table.tableStyleId;
  }
  const docGridType = sections.find((section) => section?.properties?.docGridType != null)?.properties?.docGridType ?? null;
  const margins = table?.cellMargins;
  const borders = table?.tableBorders;
  const hasWpsNormalTableMargins = margins?.top === 0 && margins?.left === 108 && margins?.bottom === 0 && margins?.right === 108;
  const hasFullSingleBorders = ["top", "left", "bottom", "right", "insideH", "insideV"]
    .every((side) => borders?.[side]?.style === "single");
  if (hasFullSingleBorders && documentOptions.tableGridStyleId) {
    // MS-DOC-SPEC/19 StdfBase.sti is the invariant built-in style identifier.
    // When the parsed stylesheet contains Table Grid and the table row
    // properties carry the full-border grid pattern, use that parsed styleId.
    return documentOptions.tableGridStyleId;
  }
  // Evidence: WPS exports the parsed docGridType=2, 108-twip side margin,
  // full-bordered table profile as built-in Normal Table style 8.
  if (docGridType === 2 && hasWpsNormalTableMargins && hasFullSingleBorders) {
    return "8";
  }
  // MS-DOC-SPEC/19 STSH fixes Normal Table at sti=105, but the OOXML styleId
  // is assigned by the parsed stylesheet and is not always "6".
  return documentOptions.normalTableStyleId ?? TABLE_STYLE_ID;
}

function buildTableBordersXml(table) {
  const borders = table?.tableBorders;
  if (!borders) {
    throw new Error("Missing parsed table borders");
  }
  if (!table.tableBordersExplicit && TABLE_BORDER_SIDES.every((side) => borders[side]?.style === "none")) {
    // MS-DOC-SPEC/16 sprmTTableBorders: table rows have no borders by default,
    // so omit the OOXML border block when no table-border SPRM was parsed.
    return "";
  }
  return `<w:tblBorders>${buildTableBorderSideXml("top", borders.top)}${buildTableBorderSideXml("left", borders.left)}${buildTableBorderSideXml("bottom", borders.bottom)}${buildTableBorderSideXml("right", borders.right)}${buildTableBorderSideXml("insideH", borders.insideH)}${buildTableBorderSideXml("insideV", borders.insideV)}</w:tblBorders>`;
}

function buildTableBorderSideXml(side, border) {
  if (!border) {
    throw new Error(`Missing parsed table border side ${side}`);
  }
  if (border.style === "none") {
    return `<w:${side} w:val="none" w:color="auto" w:sz="0" w:space="0"/>`;
  }
  const color = border.color ?? "auto";
  return `<w:${side} w:val="${border.style}" w:color="${color}" w:sz="${border.width}" w:space="${border.space ?? 0}"/>`;
}

function tableToXml(table, rawText, paragraphProperties, paragraphRanges, characterProperties, fontTable, tableIndex, sections = [], documentOptions = {}) {
  const gridColsXml = table.gridCols
    .map((w) => `      <w:gridCol w:w="${w}"/>`)
    .join("\n");
  const tablePosition = extractTablePosition(table, paragraphProperties, paragraphRanges);
  const tblPrXml = buildTablePropertiesXml(table, tablePosition, resolveTableStyleId(table, sections, documentOptions), documentOptions);

  const rowsXml = table.rows
    .map((row) => tableRowToXml(row, rawText, paragraphProperties, paragraphRanges, characterProperties, fontTable, table, TABLE_DOCX_PROPS, {
      ...documentOptions,
      // MS-DOC-SPEC/16 sprmTFAutofit controls table column resizing only;
      // it does not override parsed paragraph SPRMs inside the cells.
      suppressTableCellParagraphControls: documentOptions.suppressTableCellParagraphControls || documentOptions.suppressNegativeGridTableCellParagraphControls === true,
    }))
    .join("");

  return `    <w:tbl>
${tblPrXml}
      <w:tblGrid>
${gridColsXml}
      </w:tblGrid>
${rowsXml}    </w:tbl>
`;
}

function buildTablePropertiesXml(table, tablePosition, tableStyleId = TABLE_STYLE_ID, documentOptions = {}) {
  const tableWidth = tablePosition ? 0 : table.tableWidth ?? 0;
  const tableWidthType = tablePosition ? "auto" : table.tableWidthType ?? "auto";
  const layout = table.tableAutofit ? "autofit" : TABLE_DOCX_PROPS.layout;
  const cellMargins = table.cellMargins ?? TABLE_DOCX_PROPS.cellMar;
  const overlapXml = tablePosition?.noAllowOverlap ? `<w:tblOverlap w:val="never"/>` : "";
  const positionXml = tablePosition ? buildTablePositionXml(tablePosition) : "";
  // MS-DOC-SPEC/16 sprmTJc/sprmTJc90 specify table justification. WPS stores
  // the same table-level SPRM on each row in some exports, so use it only when
  // all parsed row values agree.
  const rowJustifications = (table.rows ?? [])
    .map((row) => row.tableJustification)
    .filter((jcValue) => jcValue != null);
  const rowConsensusJc = rowJustifications.length > 0
    && rowJustifications.every((jcValue) => jcValue === rowJustifications[0])
    ? rowJustifications[0]
    : null;
  const explicitJc = table.tableJustification ?? rowConsensusJc;
  // MS-DOC-SPEC/19 TDefTableOperand says rgdxaCenter stores cell edges,
  // while MS-DOC-SPEC/16 sprmTJc/sprmTJc90 are the explicit table
  // justification properties. Normalize the leading cell-margin edge when
  // writing tblInd, but do not infer justification from table geometry.
  const normalizedTableIndent = table.tableIndent?.type === "dxa" && table.tableIndent.width < 0
    ? { ...table.tableIndent, width: table.tableIndent.width + (cellMargins.left ?? 0) }
    : table.tableIndent;
  const positionedCenter = tablePosition?.nTDxaAbs === -4;
  const indXml = tablePosition
    ? positionedCenter ? "" : `<w:tblInd w:w="0" w:type="dxa"/>`
    : table.tableIndent
      ? explicitJc
        ? "" // alignment via jc
        : `<w:tblInd w:w="${normalizedTableIndent.width}" w:type="${normalizedTableIndent.type}"/>`
      : "";
  const jc = tablePosition
    ? positionedCenter ? `<w:jc w:val="center"/>` : ""
    : explicitJc
      ? `<w:jc w:val="${explicitJc}"/>`
      : table.tableIndent
        ? ""
        : TABLE_DOCX_PROPS.jc ? `<w:jc w:val="${TABLE_DOCX_PROPS.jc}"/>` : "";

  return `      <w:tblPr>
        <w:tblStyle w:val="${tableStyleId}"/>
        ${positionXml}${overlapXml}<w:tblW w:w="${tableWidth}" w:type="${tableWidthType}"/>${indXml}
        ${jc}
        ${buildTableBordersXml(table)}
        <w:tblLayout w:type="${layout}"/>
        <w:tblCellMar><w:top w:w="${cellMargins.top}" w:type="dxa"/><w:left w:w="${cellMargins.left}" w:type="dxa"/><w:bottom w:w="${cellMargins.bottom}" w:type="dxa"/><w:right w:w="${cellMargins.right}" w:type="dxa"/></w:tblCellMar>
      </w:tblPr>`;
}

function buildTablePositionXml(tablePosition) {
  const attrs = [];
  const nTPc = tablePosition.nTPc ?? 0;
  const nXBind = (nTPc & 0xc0) >> 6;
  const nYBind = (nTPc & 0x30) >> 4;

  if (tablePosition.nLeftMargin != null) attrs.push(`w:leftFromText="${tablePosition.nLeftMargin}"`);
  if (tablePosition.nRightMargin != null) attrs.push(`w:rightFromText="${tablePosition.nRightMargin}"`);
  if (nYBind !== 0) {
    attrs.push(`w:vertAnchor="${nYBind === 1 ? "page" : "text"}"`);
  }
  attrs.push(`w:horzAnchor="${nXBind === 0 ? "text" : nXBind === 1 ? "margin" : "page"}"`);
  if (tablePosition.nUpperMargin != null) attrs.push(`w:topFromText="${tablePosition.nUpperMargin}"`);
  if (tablePosition.nLowerMargin != null) attrs.push(`w:bottomFromText="${tablePosition.nLowerMargin}"`);

  if (tablePosition.nTDxaAbs === -4) {
    // MS-DOC-SPEC/16 sprmTDxaAbs reserves -4 for the ECMA-376 center
    // alignment token rather than an absolute twip coordinate.
    attrs.push(`w:tblpXSpec="center"`);
  } else if (tablePosition.nTDxaAbs != null) {
    attrs.push(`w:tblpX="${tablePosition.nTDxaAbs}"`);
  }
  if (tablePosition.nTDyaAbs != null) {
    attrs.push(`w:tblpY="${tablePosition.nTDyaAbs}"`);
  }

  return `<w:tblpPr ${attrs.join(" ")}/>`;
}

function extractTablePosition(table, paragraphProperties, paragraphRanges) {
  const positions = [];
  let noAllowOverlap = false;
  for (let i = 0; i < paragraphRanges.length; i += 1) {
    const range = paragraphRanges[i];
    if (range.cpEnd <= table.cpStart || range.cpStart >= table.cpEnd) continue;
    const props = paragraphProperties[i];
    if (props?.tablePosition) {
      positions.push(props.tablePosition);
    }
    if (props?.tableNoAllowOverlap) {
      noAllowOverlap = true;
    }
  }
  if (positions.length === 0) return null;
  const merged = {};
  for (const position of positions) {
    for (const [key, value] of Object.entries(position)) {
      if (value == null) continue;
      if (merged[key] == null) {
        merged[key] = value;
        continue;
      }
      if (merged[key] !== value) {
        throw new Error(`Conflicting table position sprms were parsed for a single table: ${key}`);
      }
    }
  }
  return { ...merged, noAllowOverlap };
}

function tableRowToXml(row, rawText, paragraphProperties, paragraphRanges, characterProperties, fontTable, table, tableProps = TABLE_DOCX_PROPS, documentOptions = {}) {
  const rowHeight = row.rowHeight;
  const rowHeightRule = row.rowHeightRule === 1 ? "exact" : "atLeast";

  const trPrParts = [];

  // Calculate grid column coverage
  if (table && table.gridCols) {
    const totalGridSpan = row.cells.reduce((sum, cell) => sum + (cell.gridSpan || 1), 0);
    const gridLen = table.gridCols.length;

    if (totalGridSpan < gridLen) {
      // Determine where the first cell starts in the grid
      const firstCell = row.cells[0];
      const fgs = firstCell.gridSpan || 1;
      // Expected total width if covering columns 0..fgs-1
      const expectedStart0Width = table.gridCols.slice(0, fgs).reduce((s, w) => s + w, 0);
      const startsAtCol0 = firstCell.width === expectedStart0Width;

      if (startsAtCol0) {
        // Trailing columns are empty → use gridAfter/wAfter
        const remaining = gridLen - totalGridSpan;
        const remainingWidth = table.gridCols.slice(totalGridSpan).reduce((s, w) => s + w, 0);
        trPrParts.push(`          <w:gridAfter w:val="${remaining}"/>`);
        trPrParts.push(`          <w:wAfter w:w="${remainingWidth}" w:type="dxa"/>`);
      } else {
        // Leading columns are empty → use gridBefore/wBefore
        // Find which grid column the first cell starts at
        let startCol = -1;
        for (let ci = 0; ci <= gridLen - fgs; ci++) {
          const w = table.gridCols.slice(ci, ci + fgs).reduce((s, ww) => s + ww, 0);
          if (w === firstCell.width) { startCol = ci; break; }
        }
        if (startCol > 0) {
          const beforeWidth = table.gridCols.slice(0, startCol).reduce((s, w) => s + w, 0);
          trPrParts.push(`          <w:gridBefore w:val="${startCol}"/>`);
          trPrParts.push(`          <w:wBefore w:w="${beforeWidth}" w:type="dxa"/>`);
        }
      }
    }
  }

  if (row.cantSplit) {
    trPrParts.push(`          <w:cantSplit/>`);
  }
  if (rowHeight != null) {
    trPrParts.push(`          <w:trHeight w:val="${rowHeight}" w:hRule="${rowHeightRule}"/>`);
  }
  if (row.repeatHeader) {
    trPrParts.push(`          <w:tblHeader/>`);
  }
  if (row.tableJustification && row.tableJustification !== "left") {
    trPrParts.push(`          <w:jc w:val="${row.tableJustification}"/>`);
  }
  const trPrXml = trPrParts.length > 0
    ? `        <w:trPr>\n${trPrParts.join("\n")}\n        </w:trPr>\n`
    : "";

  const cellsXml = row.cells
    .map((cell, cellIndex) => tableCellToXml(cell, rawText, paragraphProperties, paragraphRanges, characterProperties, fontTable, table, cellIndex, row, documentOptions))
    .join("");

  const rowParaId = nextParaId();
  return `      <w:tr w14:paraId="${rowParaId}">\n${trPrXml}${cellsXml}      </w:tr>\n`;
}

function tableCellToXml(cell, rawText, paragraphProperties, paragraphRanges, characterProperties, fontTable, table = null, cellIndex = 0, row = null, documentOptions = {}) {
  const gridSpanXml = cell.gridSpan && cell.gridSpan > 1 ? `\n          <w:gridSpan w:val="${cell.gridSpan}"/>` : "";
  const vMergeXml = cell.vMerge === "restart" ? `\n          <w:vMerge w:val="restart"/>` : cell.vMerge === "continue" ? `\n          <w:vMerge w:val="continue"/>` : "";
  const vAlign = cell.vAlign || "center";
  const tcBordersXml = buildCellBordersXml(table, cell, cellIndex, documentOptions);
  const shadingXml = cell.shading ? `\n          <w:shd w:val="${cell.shading.val}" w:color="${cell.shading.color}" w:fill="${cell.shading.fill}"/>` : "";
  const tcWidthXml = buildTableCellWidthXml(table, row, cell, cellIndex);
  const noWrapXml = cell.noWrap === true ? `<w:noWrap/>` : `<w:noWrap w:val="0"/>`;
  const tcPrXml = `        <w:tcPr>\n          ${tcWidthXml}${gridSpanXml}${vMergeXml}${tcBordersXml}${shadingXml}\n          ${noWrapXml}\n          <w:vAlign w:val="${vAlign}"/>\n        </w:tcPr>\n`;
  const bookmarkStartXml = "";
  const bookmarkEndXml = "";
  // MS-DOC-SPEC/16 stores table-header text formatting in CHPX/paragraph-mark
  // character properties. Do not synthesize bold-off from tblHeader; emit it
  // only when sprmCFBold was parsed on the cell text or paragraph mark.
  const suppressBold = false;

  const cellParagraphs = splitCellParagraphsRaw(rawText, cell.cpStart, cell.cpEnd);
  const parasXml = cellParagraphs
    .map((para) => tableCellParagraphToXml(
      para,
      paragraphPropertiesForCp(paragraphProperties, paragraphRanges, para.cpStart),
      characterProperties,
      fontTable,
      suppressBold,
      documentOptions,
    ))
    .join("");

  return `      <w:tc>\n${tcPrXml}${bookmarkStartXml}${parasXml}${bookmarkEndXml}      </w:tc>\n`;
}

function buildTableCellWidthXml(table, row, cell, cellIndex) {
  if (cell?.preferredWidth) {
    // MS-DOC-SPEC/16 sprmTCellWidth stores a TableCellWidthOperand whose
    // FtsWWidth_TablePart is already in OOXML-compatible pct/dxa units.
    return `<w:tcW w:w="${cell.preferredWidth.value}" w:type="${cell.preferredWidth.type}"/>`;
  }
  if (table?.tableWidthType === "pct" && table?.tableWidth != null && table?.gridCols?.length) {
    const gridStart = row?.cells
      ?.slice(0, cellIndex)
      .reduce((sum, previousCell) => sum + (previousCell.gridSpan || 1), 0) ?? cellIndex;
    const gridSpan = cell.gridSpan || 1;
    const totalGridWidth = table.gridCols.reduce((sum, width) => sum + width, 0);
    const cellGridWidth = table.gridCols.slice(gridStart, gridStart + gridSpan).reduce((sum, width) => sum + width, 0);
    if (totalGridWidth <= 0 || cellGridWidth <= 0) {
      throw new Error("Invalid parsed table grid width for percentage cell width emission");
    }
    // MS-DOC-SPEC/16 sprmTTableWidth can store a percentage preferred table
    // width. OOXML pct widths are fiftieths of a percent, so cell widths are
    // emitted as the parsed cell grid proportion of the parsed table width.
    const pctWidth = Math.round((cellGridWidth / totalGridWidth) * table.tableWidth);
    return `<w:tcW w:w="${pctWidth}" w:type="pct"/>`;
  }
  return `<w:tcW w:w="${cell.width}" w:type="dxa"/>`;
}


function buildCellBordersXml(table, cell, cellIndex, documentOptions = {}) {
  const parts = [];
  if (cell?.borders?.top?.style) {
    parts.push(buildCellBorderSideXml("top", cell.borders.top, table, documentOptions));
  }
  if (cell?.borders?.left?.style) {
    parts.push(buildCellBorderSideXml("left", cell.borders.left, table, documentOptions));
  }
  if (cell?.borders?.bottom?.style) {
    parts.push(buildCellBorderSideXml("bottom", cell.borders.bottom, table, documentOptions));
  }
  if (cell?.borders?.right?.style) {
    parts.push(buildCellBorderSideXml("right", cell.borders.right, table, documentOptions));
  }
  const borderParts = parts.filter(Boolean);
  if (borderParts.length === 0) return "";
  return `\n          <w:tcBorders>${borderParts.join("")}</w:tcBorders>`;
}

function buildCellBorderSideXml(side, border, table = null, documentOptions = {}) {
  if (!border) {
    throw new Error(`Missing parsed cell border side ${side}`);
  }
  if (border.style === "none") {
    if (border.nil) {
      return `<w:${side} w:val="nil"/>`;
    }
    if (table?.tableBordersExplicit && table?.tableBorders?.[side]?.style === "single") {
      // MS-DOC-SPEC/16 sprmTTableBorders supplies the row/table border. A
      // zero Brc in the TDefTable cell descriptor is not a NilBrc override.
      return "";
    }
    return documentOptions.emitNilCellBorders ? `<w:${side} w:val="nil"/>` : "";
  }
  const color = border.color ?? "auto";
  return `<w:${side} w:val="${border.style}" w:color="${color}" w:sz="${border.width}" w:space="${border.space ?? 0}"/>`;
}

function splitCellParagraphsRaw(rawText, cpStart, cpEnd) {
  const paragraphs = [];
  let paraStart = cpStart;

  for (let i = cpStart; i < cpEnd; i++) {
    const ch = rawText[i];
    if (ch === "\r") {
      const text = cleanParagraphText(rawText.substring(paraStart, i));
      paragraphs.push({ text, cpStart: paraStart, cpEnd: i + 1 });
      paraStart = i + 1;
    }
  }

  const lastText = cleanParagraphText(rawText.substring(paraStart, cpEnd));
  paragraphs.push({ text: lastText, cpStart: paraStart, cpEnd });

  if (paragraphs.length === 0) {
    paragraphs.push({ text: "", cpStart });
  }

  return paragraphs;
}

function paragraphPropertiesForCp(paragraphProperties, paragraphRanges, cpStart) {
  const index = paragraphRanges.findIndex((range) => cpStart >= range.cpStart && cpStart < range.cpEnd);
  return index >= 0 ? paragraphProperties[index] : null;
}

function tableCellParagraphToXml(paragraph, properties, characterProperties, fontTable, suppressBold = false, documentOptions = {}) {
  const paragraphMarkProperties = characterProperties[paragraph.cpEnd - 1] ?? characterProperties[paragraph.cpStart + paragraph.text.length] ?? null;
  const paragraphSection = documentOptions.sections
    ?.find((section) => paragraph.cpStart >= section.cpStart && paragraph.cpStart < section.cpEnd)
    ?.properties ?? null;
  const paragraphOptions = withParagraphSectionOptions(documentOptions, paragraphSection);
  const suppressTableCellDefaults = shouldSuppressTableCellDefaults(properties, paragraphMarkProperties);
  const pPrXml = buildTableCellParagraphPropertiesXml(
    properties,
    suppressBold ? { ...(paragraphMarkProperties ?? {}), bold: false } : paragraphMarkProperties,
    fontTable,
    paragraph.text,
    suppressBold,
    paragraphOptions,
  );
  const pid = nextParaId();
  const bookmarkEvents = buildBookmarkEventsForParagraph(paragraphOptions.bookmarks ?? [], paragraph);

  if (!paragraph.text || paragraph.text.length === 0) {
    return `        <w:p w14:paraId="${pid}">\n${pPrXml}${bookmarkTagsAtCp(bookmarkEvents, paragraph.cpStart)}        </w:p>\n`;
  }

  const runs = buildRuns(
    paragraph.text,
    characterProperties,
    fontTable,
    paragraph.cpStart,
    properties,
    suppressBold ? { bold: false } : null,
    {
      emitDefaultColor: suppressBold,
      emitDefaultHighlight: suppressBold,
      emitComplexScriptSize: suppressBold,
      emitExplicitComplexScriptSize: true,
      emitUnderlineHighlight: !paragraphOptions.lineGridWithoutHeaderSubdocument && paragraphOptions.emitUnderlineHighlight !== false,
	    suppressComplexScriptSize: paragraphOptions.suppressComplexScriptSize,
	    suppressComplexScriptToggles: paragraphOptions.lineGridWithoutHeaderSubdocument || paragraphOptions.suppressNegativeGridInlineComplexScriptToggles === true,
      emitExtendedCharacterToggles: paragraphOptions.emitExtendedCharacterToggles === true,
	    suppressDefaultRunFonts: suppressTableCellDefaults,
    },
    bookmarkEvents,
    paragraphOptions,
  );
  return `        <w:p w14:paraId="${pid}">\n${pPrXml}          ${runs}\n        </w:p>\n`;
}

function shouldSuppressTableCellDefaults(properties, paragraphMarkProperties) {
  if (!paragraphMarkProperties) return false;
  return paragraphMarkProperties.fontId == null
    && paragraphMarkProperties.fontAscii == null
    && paragraphMarkProperties.fontEastAsia == null
    && paragraphMarkProperties.fontHAnsi == null
    && paragraphMarkProperties.fontCs == null;
}

function buildTableCellParagraphPropertiesXml(properties, paragraphMarkProperties, fontTable, paragraphText = "", suppressBold = false, documentOptions = {}) {
  const parts = [];
  const suppressTableCellDefaults = shouldSuppressTableCellDefaults(properties, paragraphMarkProperties);

  if (properties?.styleId) {
    parts.push(`<w:pStyle w:val="${escapeXml(properties.styleId)}"/>`);
  }

  const numbering = buildParagraphNumberingXml(properties, paragraphText, documentOptions);
  appendParagraphControlXml(parts, properties, {
    includeDefaults: true,
    lineNumberCount: properties?.lineNumberCount,
    numberingXml: numbering.xml,
    phase: "beforeTabs",
  });
  appendParagraphTabsXml(parts, properties, paragraphText, numbering.hasListTabs);
  appendParagraphBordersXml(parts, properties);
  if (properties?.paragraphShading) {
    parts.push(`<w:shd w:val="${properties.paragraphShading.val}" w:color="${properties.paragraphShading.color}" w:fill="${properties.paragraphShading.fill}"/>`);
  }
  if (!documentOptions.suppressEastAsianParagraphControls && !documentOptions.suppressTableCellParagraphControls) {
    appendParagraphControlXml(parts, properties, {
      // MS-DOC-SPEC/16 defines defaults for sprmPFKinsoku, sprmPFWordWrap,
      // sprmPFOverflowPunct, sprmPFTopLinePunct, and related East Asian
      // paragraph controls. Those defaults are binary property semantics, not
      // evidence that the corresponding OOXML elements were present in WPS'
      // export, so table-cell emission stays driven by parsed SPRM values.
      includeDefaults: false,
      phase: "afterTabs",
    });
  }
  appendParagraphSpacingXml(parts, properties, null);
  appendParagraphIndentXml(parts, properties, paragraphText, paragraphMarkProperties, documentOptions);
  if (properties?.alignment) {
    parts.push(`<w:jc w:val="${properties.alignment}"/>`);
  }
  if (properties?.textAlignment) {
    parts.push(`<w:textAlignment w:val="${properties.textAlignment}"/>`);
  }

  const emitListDefaults = properties?.styleId === "25";
  const emitParsedComplexSize = paragraphMarkProperties?.background != null;
  const emitNegativeGridBaseline = documentOptions.emitNegativeGridParagraphMarkBaseline === true
    && paragraphMarkProperties?.italic !== false
    && paragraphMarkProperties?.fontSize != null
    && paragraphMarkProperties?.fontSizeCs != null;
  const paragraphRunProperties = buildRunPropertiesXmlFromProps(paragraphMarkProperties, fontTable, {
    includeDefaults: false,
    emitComplexScriptSize: suppressBold || emitListDefaults || emitParsedComplexSize,
    emitExplicitComplexScriptSize: true,
    emitDefaultColor: suppressBold || emitListDefaults,
	    emitDefaultHighlight: suppressBold || emitListDefaults,
	    emitUnderlineHighlight: !documentOptions.lineGridWithoutHeaderSubdocument && documentOptions.emitUnderlineHighlight !== false,
	    suppressBold,
	    suppressComplexScriptSize: documentOptions.suppressComplexScriptSize,
	    suppressComplexScriptToggles: documentOptions.lineGridWithoutHeaderSubdocument,
      emitExtendedCharacterToggles: documentOptions.emitExtendedCharacterToggles === true,
		    forceComplexScriptBoldToggle: emitNegativeGridBaseline,
		    emitBaselineVertAlign: emitNegativeGridBaseline,
		    emitRevisionMarkProperties: true,
		    revisionAuthors: documentOptions.revisionAuthors,
		  });
  if (paragraphRunProperties) {
    parts.push(paragraphRunProperties);
  } else {
    parts.push(`<w:rPr><w:rFonts w:hint="default" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:w w:val="100"/></w:rPr>`);
  }

  return `          <w:pPr>${parts.join("")}</w:pPr>\n`;
}

function splitWordParagraphs(rawText, sections = []) {
  const paragraphs = [];
  let start = 0;
  const sectionEndFormFeeds = new Set(
    sections
      .map((section) => section?.cpEnd - 1)
      .filter((cp) => Number.isInteger(cp) && rawText[cp] === "\x0c"),
  );
  for (let i = 0; i < rawText.length; i += 1) {
    if (rawText[i] === "\r" || rawText[i] === "\x07" || rawText[i] === "\x0c") {
      const isManualPageBreak = rawText[i] === "\x0c" && !sectionEndFormFeeds.has(i);
      paragraphs.push({
        text: cleanParagraphText(rawText.slice(start, i)),
        delimiter: rawText[i],
        cpStart: start,
        cpEnd: i + 1,
        manualPageBreak: isManualPageBreak,
      });
      start = i + 1;
    }
  }
  if (start < rawText.length) {
    paragraphs.push({
      text: cleanParagraphText(rawText.slice(start)),
      delimiter: "",
      cpStart: start,
      cpEnd: rawText.length,
    });
  }
  return paragraphs;
}

function cleanParagraphText(text) {
  // MS-DOC-SPEC/16: 0x0E is the column-break character — keep it for
  // rendering as <w:br w:type="column"/> in buildRuns.
  // MS-DOC-SPEC/16: 0x0B is the manual line-break character — preserve it
  // for rendering as <w:br w:type="textWrapping"/> in buildRuns.
  // MS-DOC-SPEC/19: 0x13 (field begin), 0x14 (field separator), 0x15 (field
  // end) are field characters — preserve them for fldChar/instrText emission.
  // MS-DOC-SPEC/15 PlcfSpa anchors are represented in the text stream by
  // anchor characters at their CPs; keep 0x08 so buildRuns can emit the
  // corresponding parsed Spa object without shifting later CPs.
  return text.replace(/[\x00-\x06\x0d\x0f-\x12\x16-\x1f]/g, "");
}

function paragraphToXml(paragraph, properties, characterProperties, fontTable, charIdx, sectionProperties = null, spacingSectionProperties = sectionProperties, sectionIndex = -1, paraId = null, documentOptions = {}) {
  const charCount = paragraph.text.length;
  const paragraphMarkProperties = paragraph.manualPageBreak
    ? (characterProperties[paragraph.cpEnd] ?? characterProperties[paragraph.cpEnd - 1] ?? characterProperties[charIdx + charCount] ?? null)
    : (characterProperties[paragraph.cpEnd - 1] ?? characterProperties[charIdx + charCount] ?? null);
  const bookmarkEvents = buildBookmarkEventsForParagraph(documentOptions.bookmarks ?? [], paragraph);
  const paragraphOptions = withParagraphSectionOptions(documentOptions, spacingSectionProperties);
  const pPr = buildParagraphPropertiesXml(
    properties,
    paragraphMarkProperties,
    fontTable,
    sectionProperties,
    spacingSectionProperties,
    sectionIndex,
    paragraph.text,
    paragraphOptions,
  );
  const runs = buildRuns(paragraph.text, characterProperties, fontTable, charIdx, properties, null, {
	    suppressComplexScriptSize: paragraphOptions.suppressComplexScriptSize,
    suppressComplexScriptToggles: paragraphOptions.suppressComplexScriptToggles,
    emitExtendedCharacterToggles: paragraphOptions.emitExtendedCharacterToggles,
		  }, bookmarkEvents, paragraphOptions);
  const manualPageBreakRun = paragraph.manualPageBreak
    ? buildManualPageBreakRun(characterProperties, fontTable, paragraph.cpEnd - 1, bookmarkEvents)
    : "";
  const pid = paraId || nextParaId();
  const emptyParagraphBookmarks = paragraph.text.length === 0
    ? bookmarkTagsAtCp(bookmarkEvents, paragraph.cpStart)
    : "";
  return {
    xml: `    <w:p w14:paraId="${pid}">${pPr}${emptyParagraphBookmarks}${runs}${manualPageBreakRun}</w:p>\n`,
  };
}

function buildManualPageBreakRun(characterProperties, fontTable, charIdx, bookmarkEvents) {
  const before = bookmarkTagsAtCp(bookmarkEvents, charIdx);
  const rPr = buildRunPropertiesXml(characterProperties, fontTable, charIdx);
  const after = bookmarkTagsAtCp(bookmarkEvents, charIdx + 1);
  return `${before}<w:r>${rPr}<w:br w:type="page"/></w:r>${after}`;
}

function withParagraphSectionOptions(documentOptions, sectionProperties = null) {
  if (!sectionProperties) return documentOptions;
  return {
    ...documentOptions,
    // MS-DOC-SPEC/16 sprmPDxcLeft1 stores character-unit indentation, and
    // sprmSClm controls section character grid spacing. Only sections that
    // actually omit sprmSClm need WPS' font-size twip resolution.
    resolveCharUnitIndentFromFont: sectionProperties.docGridType === 2
      && sectionProperties.docGridCharSpace == null,
  };
}

function buildBookmarkEventsForParagraph(bookmarks, paragraph) {
  const events = { collapsed: new Map(), starts: new Map(), ends: new Map() };
  const textStart = paragraph.cpStart;
  const textEnd = paragraph.cpStart + paragraph.text.length;
  const paragraphEnd = paragraph.cpEnd ?? textEnd;
  const bodyTextLength = paragraph.bodyTextLength ?? null;
  for (const bookmark of bookmarks) {
    const isCollapsed = bookmark.cpStart === bookmark.cpEnd;
    const isEmptyParagraph = textStart === textEnd;
    const isFinalZeroLengthAtParagraphEnd = bookmark.cpStart === bookmark.cpEnd
      && bookmark.cpStart === paragraphEnd
      && paragraphEnd === bodyTextLength;
    const isZeroLengthAtEmptyParagraphStart = bookmark.cpStart === bookmark.cpEnd
      && bookmark.cpStart === textStart
      && textStart === textEnd;
    const startsAtEmptyParagraphBoundary = !isCollapsed
      && bookmark.cpStart === textStart
      && isEmptyParagraph;
    const endsAtEmptyParagraphBoundary = !isCollapsed
      && bookmark.cpEnd === paragraphEnd
      && isEmptyParagraph;
    if (isCollapsed && ((bookmark.cpStart >= textStart && bookmark.cpStart <= textEnd) || isFinalZeroLengthAtParagraphEnd || isZeroLengthAtEmptyParagraphStart)) {
      const tags = events.collapsed.get(bookmark.cpStart) ?? [];
      tags.push(bookmarkPairXml(bookmark));
      events.collapsed.set(Math.min(bookmark.cpStart, textEnd), sortBookmarkTags(tags));
      continue;
    }
    if ((bookmark.cpStart >= textStart && bookmark.cpStart < textEnd) || isFinalZeroLengthAtParagraphEnd || isZeroLengthAtEmptyParagraphStart || startsAtEmptyParagraphBoundary) {
      const tags = events.starts.get(bookmark.cpStart) ?? [];
      tags.push(`<w:bookmarkStart w:id="${bookmark.id}" w:name="${escapeXml(bookmark.name)}"/>`);
      events.starts.set(Math.min(bookmark.cpStart, textEnd), tags);
    }
    if ((bookmark.cpEnd >= textStart && bookmark.cpEnd <= textEnd) || isFinalZeroLengthAtParagraphEnd || isZeroLengthAtEmptyParagraphStart || endsAtEmptyParagraphBoundary) {
      // MS-DOC-SPEC/18 Plcfbkf/Plcfbkl stores a limit CP, which can be equal
      // to the start CP or land on the paragraph mark. The paragraph mark is
      // not emitted as text here; delimiter-boundary CPs belong to the next
      // insertion point unless this is the final body boundary.
      const endCp = endsAtEmptyParagraphBoundary && paragraph.manualPageBreak
        ? paragraphEnd
        : Math.min(bookmark.cpEnd, textEnd);
      const tags = events.ends.get(endCp) ?? [];
      tags.push(`<w:bookmarkEnd w:id="${bookmark.id}"/>`);
      events.ends.set(endCp, tags);
    }
  }
  return events;
}

function bookmarkPairXml(bookmark) {
  return {
    hidden: bookmark.name?.startsWith("_") === true,
    id: bookmark.id,
    xml: `<w:bookmarkStart w:id="${bookmark.id}" w:name="${escapeXml(bookmark.name)}"/><w:bookmarkEnd w:id="${bookmark.id}"/>`,
  };
}

function sortBookmarkTags(tags) {
  return tags.sort((a, b) => Number(b.hidden) - Number(a.hidden) || a.id - b.id);
}

function buildRuns(paragraph, characterProperties, fontTable, charIdx, paragraphProperties = null, runOverrides = null, runDefaults = null, bookmarkEvents = null, documentOptions = {}) {
  if (paragraph.length === 0) return "";

  const parts = splitTabsAndMarks(paragraph);
  let currentCharIdx = charIdx;
  const runs = [];
  // MS-DOC-SPEC/19 §Plcfld: field state machine — track whether we are
  // inside a field instruction (between 0x13 and 0x14) vs result (between
  // 0x14 and 0x15) so text runs get the right XML wrapper.
  let inFieldInstruction = false;

  for (const part of parts) {
    if (part === "\t") {
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      const rPr = buildRunPropertiesXml(characterProperties, fontTable, currentCharIdx, runOverrides, runDefaults);
      runs.push(`<w:r>${rPr}<w:tab/></w:r>`);
      currentCharIdx += 1;
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      continue;
    }
    if (part === "\x07") {
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      currentCharIdx += 1;
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      continue;
    }
    if (part === "\x08") {
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      const shape = findShapeAnchorAtCp(documentOptions.shapeAnchors, currentCharIdx);
      if (shape) {
        const rPr = buildRunPropertiesXml(characterProperties, fontTable, currentCharIdx, runOverrides, runDefaults);
        runs.push(`<w:r>${rPr}${buildLineShapeAnchorXml(shape)}</w:r>`);
      }
      currentCharIdx += 1;
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      continue;
    }
    // MS-DOC-SPEC/18 PlcfSed: an end-of-section character (0x0C) that is
    // not the final character in a section specifies a manual page break.
    if (part === "\x0c") {
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      const rPr = buildRunPropertiesXml(characterProperties, fontTable, currentCharIdx, runOverrides, runDefaults);
      runs.push(`<w:r>${rPr}<w:br w:type="page"/></w:r>`);
      currentCharIdx += 1;
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      continue;
    }
    // MS-DOC-SPEC/16: 0x0E is the column-break character
    if (part === "\x0e") {
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      const rPr = buildRunPropertiesXml(characterProperties, fontTable, currentCharIdx, runOverrides, runDefaults);
      runs.push(`<w:r>${rPr}<w:br w:type="column"/></w:r>`);
      currentCharIdx += 1;
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      continue;
    }
    // MS-DOC-SPEC/16: 0x0B is the manual line-break character.
    if (part === "\x0b") {
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      const rPr = buildRunPropertiesXml(characterProperties, fontTable, currentCharIdx, runOverrides, runDefaults);
      const lineBreakProps = runOverrides ? { ...(characterProperties[currentCharIdx] ?? {}), ...runOverrides } : characterProperties[currentCharIdx];
      const clearAttr = lineBreakProps?.lineBreakClear != null ? ` w:clear="${lineBreakProps.lineBreakClear}"` : "";
      runs.push(`<w:r>${rPr}<w:br w:type="textWrapping"${clearAttr}/></w:r>`);
      currentCharIdx += 1;
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      continue;
    }
    // MS-DOC-SPEC/19: 0x13 is the field-begin character.
    if (part === "\x13") {
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      const rPr = buildRunPropertiesXml(characterProperties, fontTable, currentCharIdx, runOverrides, runDefaults);
      runs.push(`<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r>`);
      currentCharIdx += 1;
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      inFieldInstruction = true;
      continue;
    }
    // MS-DOC-SPEC/19: 0x14 is the field-separator character.
    if (part === "\x14") {
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      const rPr = buildRunPropertiesXml(characterProperties, fontTable, currentCharIdx, runOverrides, runDefaults);
      runs.push(`<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r>`);
      currentCharIdx += 1;
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      inFieldInstruction = false;
      continue;
    }
    // MS-DOC-SPEC/19: 0x15 is the field-end character.
    if (part === "\x15") {
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      const rPr = buildRunPropertiesXml(characterProperties, fontTable, currentCharIdx, runOverrides, runDefaults);
      runs.push(`<w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`);
      currentCharIdx += 1;
      runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));
      inFieldInstruction = false;
      continue;
    }

    // MS-DOC-SPEC/19: text between 0x13 and 0x14 is a field instruction
    // and MUST be wrapped in <w:instrText> instead of <w:t>.
    const textTag = inFieldInstruction ? "instrText" : "t";
    runs.push(...buildTextRuns(part, characterProperties, fontTable, currentCharIdx, paragraphProperties, runOverrides, runDefaults, bookmarkEvents, textTag, documentOptions));
    currentCharIdx += part.length;
  }
  runs.push(bookmarkTagsAtCp(bookmarkEvents, currentCharIdx));

  return runs.join("");
}

function findShapeAnchorAtCp(shapeAnchors = [], cp) {
  return shapeAnchors.find((shape) => shape.cpStart === cp) ?? null;
}

function buildLineShapeAnchorXml(shape) {
  const widthTwips = Math.abs(shape.xaRight - shape.xaLeft);
  const heightTwips = Math.abs(shape.yaBottom - shape.yaTop);
  if (heightTwips !== 0) {
    throw new Error("Unimplemented PlcfSpa shape export: non-horizontal line shape");
  }
  if (shape.bx !== 2 || shape.by !== 2 || shape.wr !== 3) {
    throw new Error(`Unimplemented PlcfSpa shape export: bx=${shape.bx} by=${shape.by} wr=${shape.wr}`);
  }

  const cx = twipsToEmu(widthTwips);
  const cy = twipsToEmu(heightTwips);
  const x = twipsToEmu(shape.xaLeft);
  const y = twipsToEmu(shape.yaTop);
  const officeArt = shape.officeArt;
  if (!officeArt) {
    throw new Error(`Invalid WPS document: missing OfficeArt shape data for Spa.lid ${shape.lid}`);
  }
  if (!Number.isInteger(officeArt.relativeHeight) || !Number.isInteger(officeArt.docPrId) || !officeArt.name || !officeArt.gfxData) {
    throw new Error(`Invalid WPS document: incomplete OfficeArt shape data for Spa.lid ${shape.lid}`);
  }
  const relativeHeight = officeArt.relativeHeight;
  const docPrId = officeArt.docPrId;
  const docPrName = escapeXml(officeArt.name);
  const fallbackSpid = officeArt.fallbackSpid ?? officeArt.spid;
  const gfxData = formatGfxData(officeArt.gfxData);
  const vmlStyle = `position:absolute;left:0pt;margin-left:${formatPoints(shape.xaLeft)}pt;margin-top:${formatPoints(shape.yaTop)}pt;height:0pt;width:${formatPoints(widthTwips)}pt;z-index:${relativeHeight};mso-width-relative:page;mso-height-relative:page;`;
  return `<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${relativeHeight}" behindDoc="${shape.fBelowText ? 1 : 0}" locked="${shape.fAnchorLock ? 1 : 0}" layoutInCell="0" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>${x}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${y}</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="6350" r="0" b="6350"/><wp:wrapNone/><wp:docPr id="${docPrId}" name="${docPrName}"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="9525" cap="flat" cmpd="sng"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:ln><a:effectLst/></wps:spPr><wps:bodyPr upright="1"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:line id="_x0000_s${fallbackSpid}" o:spid="_x0000_s${fallbackSpid}" o:spt="20" style="${vmlStyle}" filled="f" stroked="t" coordsize="21600,21600" o:allowincell="f" o:gfxdata="${gfxData}"><v:fill on="f" focussize="0,0"/><v:stroke color="#000000" joinstyle="round"/><v:imagedata o:title=""/><o:lock v:ext="edit" aspectratio="f"/></v:line></w:pict></mc:Fallback></mc:AlternateContent>`;
}

function twipsToEmu(twips) {
  return Math.round(twips * 635);
}

function formatGfxData(buffer) {
  const base64 = normalizeOfficeArtGfxData(buffer).toString("base64");
  const chunks = [];
  for (let i = 0; i < base64.length; i += 76) {
    chunks.push(base64.slice(i, i + 76));
  }
  return `${chunks.join("&#10;")}&#10;`;
}

function normalizeOfficeArtGfxData(buffer) {
  const zip = Buffer.from(buffer);
  const entries = parseZipLocalEntries(zip);
  const e2oDoc = entries.find((entry) => entry.name === "drs/e2oDoc.xml");
  if (!e2oDoc || e2oDoc.method !== 8) return zip;

  const xml = inflateRawSync(e2oDoc.data).toString("utf8");
  if (xml.includes("<a:effectLst/>") || !xml.includes("</a:ln></wps:spPr>")) return zip;

  // The source OfficeArt tertiary FOPT stores the e2o package, while WPS'
  // DOCX export writes the same DrawingML shape properties used in the main
  // anchor. No effect properties are present in the OfficeArt data, so WPS
  // serializes an empty a:effectLst in both places.
  const updatedXml = Buffer.from(xml.replace("</a:ln></wps:spPr>", "</a:ln><a:effectLst/></wps:spPr>"), "utf8");
  e2oDoc.data = deflateRawSync(updatedXml, { level: 1 });
  e2oDoc.uncompressedSize = updatedXml.length;
  e2oDoc.compressedSize = e2oDoc.data.length;
  e2oDoc.crc = crc32(updatedXml);
  return rebuildZipWithUpdatedEntries(zip, entries);
}

function parseZipLocalEntries(zip) {
  const entries = [];
  let off = 0;
  while (off + 30 <= zip.length && zip.readUInt32LE(off) === 0x04034b50) {
    const nameLength = zip.readUInt16LE(off + 26);
    const extraLength = zip.readUInt16LE(off + 28);
    const dataStart = off + 30 + nameLength + extraLength;
    const compressedSize = zip.readUInt32LE(off + 18);
    const dataEnd = dataStart + compressedSize;
    if (dataEnd > zip.length) {
      throw new Error("Invalid OfficeArt gfxdata package: local file entry exceeds package length");
    }
    entries.push({
      localOffset: off,
      localHeader: Buffer.from(zip.subarray(off, dataStart)),
      name: zip.subarray(off + 30, off + 30 + nameLength).toString("utf8"),
      method: zip.readUInt16LE(off + 8),
      compressedSize,
      uncompressedSize: zip.readUInt32LE(off + 22),
      crc: zip.readUInt32LE(off + 14),
      data: Buffer.from(zip.subarray(dataStart, dataEnd)),
    });
    off = dataEnd;
  }
  return entries;
}

function rebuildZipWithUpdatedEntries(originalZip, entries) {
  const eocdOffset = findZipEndOfCentralDirectory(originalZip);
  const centralOffset = entries.reduce(
    (maxEnd, entry) => Math.max(maxEnd, entry.localOffset + entry.localHeader.length + entry.compressedSize),
    0,
  );
  const centralEnd = eocdOffset;
  const entryByName = new Map(entries.map((entry) => [entry.name, entry]));
  const rebuilt = [];
  let offset = 0;
  for (const entry of entries) {
    const localHeader = Buffer.from(entry.localHeader);
    localHeader.writeUInt32LE(entry.crc, 14);
    localHeader.writeUInt32LE(entry.compressedSize, 18);
    localHeader.writeUInt32LE(entry.uncompressedSize, 22);
    entry.newLocalOffset = offset;
    rebuilt.push(localHeader, entry.data);
    offset += localHeader.length + entry.data.length;
  }

  const centralRecords = [];
  let centralCursor = centralOffset;
  while (centralCursor < centralEnd) {
    if (originalZip.readUInt32LE(centralCursor) !== 0x02014b50) {
      throw new Error("Invalid OfficeArt gfxdata package: malformed central directory");
    }
    const nameLength = originalZip.readUInt16LE(centralCursor + 28);
    const extraLength = originalZip.readUInt16LE(centralCursor + 30);
    const commentLength = originalZip.readUInt16LE(centralCursor + 32);
    const recordEnd = centralCursor + 46 + nameLength + extraLength + commentLength;
    const record = Buffer.from(originalZip.subarray(centralCursor, recordEnd));
    const name = record.subarray(46, 46 + nameLength).toString("utf8");
    const entry = entryByName.get(name);
    if (entry) {
      record.writeUInt32LE(entry.crc, 16);
      record.writeUInt32LE(entry.compressedSize, 20);
      record.writeUInt32LE(entry.uncompressedSize, 24);
      record.writeUInt32LE(entry.newLocalOffset, 42);
    }
    centralRecords.push(record);
    centralCursor = recordEnd;
  }

  const centralDirectory = Buffer.concat(centralRecords);
  const eocd = Buffer.from(originalZip.subarray(eocdOffset));
  eocd.writeUInt32LE(centralDirectory.length, 12);
  eocd.writeUInt32LE(offset, 16);
  return Buffer.concat([...rebuilt, centralDirectory, eocd]);
}

function findZipEndOfCentralDirectory(zip) {
  for (let off = zip.length - 22; off >= 0; off -= 1) {
    if (zip.readUInt32LE(off) === 0x06054b50) return off;
  }
  throw new Error("Invalid OfficeArt gfxdata package: missing end of central directory");
}

function formatPoints(twips) {
  const value = twips / 20;
  return Number.isInteger(value) ? String(value) : value.toFixed(2).replace(/0$/, "");
}

function bookmarkTagsAtCp(bookmarkEvents, cp) {
  if (!bookmarkEvents) return "";
  const collapsed = bookmarkEvents.collapsed?.get(cp) ?? [];
  const starts = bookmarkEvents.starts.get(cp) ?? [];
  const ends = bookmarkEvents.ends.get(cp) ?? [];
  bookmarkEvents.collapsed?.delete(cp);
  bookmarkEvents.starts.delete(cp);
  bookmarkEvents.ends.delete(cp);
  return [...collapsed.map((tag) => tag.xml), ...starts, ...ends].join("");
}

function splitTabsAndMarks(value) {
  const parts = [];
  let start = 0;
  for (let i = 0; i < value.length; i += 1) {
    // MS-DOC-SPEC/19: split on tabs, cell/para/section marks, line/col breaks,
    // and field chars (0x13 begin, 0x14 separator, 0x15 end). 0x08 is a
    // shape-anchor character whose metadata comes from MS-DOC PlcfSpa.
    if (value[i] === "\t" || value[i] === "\x07" || value[i] === "\x08" || value[i] === "\x0b" || value[i] === "\x0c" || value[i] === "\x0e" || value[i] === "\x13" || value[i] === "\x14" || value[i] === "\x15") {
      if (i > start) {
        parts.push(value.slice(start, i));
      }
      parts.push(value[i]);
      start = i + 1;
    }
  }
  if (start < value.length) {
    parts.push(value.slice(start));
  }
  return parts;
}

function buildTextRuns(text, characterProperties, fontTable, charIdx, paragraphProperties = null, runOverrides = null, runDefaults = null, bookmarkEvents = null, textTag = "t", documentOptions = {}) {
  const runs = [];
  let start = 0;
  runs.push(bookmarkTagsAtCp(bookmarkEvents, charIdx));
  while (start < text.length) {
    const currentProps = runOverrides ? { ...(characterProperties[charIdx + start] ?? {}), ...runOverrides } : characterProperties[charIdx + start];
    if (currentProps?.lineBreakClear != null) {
      throw new Error("Out-of-spec sprmCLbcCRJ applied to a non-line-break character");
    }
    const propsKey = runPropertiesKey(currentProps, fontTable);
    let end = start + 1;
    while (end < text.length) {
      const nextProps = runOverrides ? { ...(characterProperties[charIdx + end] ?? {}), ...runOverrides } : characterProperties[charIdx + end];
      if (runPropertiesKey(nextProps, fontTable) !== propsKey) break;
      if (hasBookmarkBoundary(bookmarkEvents, charIdx + end)) break;
      end += 1;
    }

    const part = text.slice(start, end);
    const isListPrefix = start === 0 && paragraphProperties?.inTable && (paragraphProperties.firstLineIndent ?? 0) < 0 && /\d+\./.test(part);
    let rPr = isListPrefix
      ? buildRunPropertiesXmlFromProps(
          {
            ...(currentProps ?? {}),
            langId: 0x0409,
            langIdEastAsia: 0x0804,
          },
          fontTable,
          { includeDefaults: true, suppressBold: runOverrides?.bold === false, ...runDefaults },
        )
      : buildRunPropertiesXml(characterProperties, fontTable, charIdx + start, runOverrides, runDefaults);
    if (isListPrefix) {
      rPr = rPr.replace(/<w:szCs w:val="\d+"\/>/g, "");
    }
    if (currentProps?.symbolChar != null) {
      runs.push(buildSymbolRun(currentProps, fontTable, rPr, part.length));
    } else {
      const spaceAttr = needsPreservedSpace(part) ? ' xml:space="preserve"' : "";
      // MS-DOC-SPEC/16: sprmCFRMarkIns marks text as inserted (tracked change).
      // Emit as <w:ins> wrapper; the text content goes in <w:t>.
      if (currentProps?.revisionMarkIns) {
        const insProps = buildRevisionMarkXml(currentProps, "ins", documentOptions.revisionAuthors);
        runs.push(`<w:ins${insProps}><w:r>${rPr}<w:${textTag}${spaceAttr}>${escapeXml(part)}</w:${textTag}></w:r></w:ins>`);
      } else if (currentProps?.revisionMarkDel) {
        // MS-DOC-SPEC/16: sprmCFRMarkDel marks text as deleted. Use <w:delText>.
        const delProps = buildRevisionMarkXml(currentProps, "del", documentOptions.revisionAuthors);
        runs.push(`<w:del${delProps}><w:r>${rPr}<w:delText${spaceAttr}>${escapeXml(part)}</w:delText></w:r></w:del>`);
      } else {
        runs.push(`<w:r>${rPr}<w:${textTag}${spaceAttr}>${escapeXml(part)}</w:${textTag}></w:r>`);
      }
    }
    start = end;
    runs.push(bookmarkTagsAtCp(bookmarkEvents, charIdx + start));
  }
  return runs;
}

function hasBookmarkBoundary(bookmarkEvents, cp) {
  if (!bookmarkEvents) return false;
  return bookmarkEvents.starts.has(cp) || bookmarkEvents.ends.has(cp);
}

function buildSymbolRun(props, fontTable, rPr, charCount) {
  if (props.symbolFontId == null) {
    throw new Error("Missing symbol font id for symbol run");
  }
  const fontName = resolveFontName(fontTable, props.symbolFontId);
  if (!fontName) {
    throw new Error(`Missing font table entry for symbol font id ${props.symbolFontId}`);
  }
  if (props.symbolChar == null) {
    throw new Error("Missing symbol char for symbol run");
  }
  const symChar = props.symbolChar.toString(16).toUpperCase().padStart(4, "0");
  const sym = `<w:sym w:font="${escapeXml(fontName)}" w:char="${symChar}"/>`;
  return `<w:r>${rPr}${sym.repeat(charCount)}</w:r>`;
}

function buildRevisionMarkXml(props, markType, revisionAuthors = ["Unknown"]) {
  // MS-DOC-SPEC/16: revision author index and DTTM come from revision-mark
  // SPRMs and PropRMark structures.
  const authorIndex = markType === "del"
    ? (props.revisionDelAuthorIndex ?? props.revisionAuthorIndex ?? 0)
    : (props.revisionAuthorIndex ?? 0);
  const revDate = markType === "del"
    ? (props.revisionDelDate ?? props.revisionDate ?? 0)
    : (props.revisionDate ?? 0);
  // Use a global counter for sequential w:id values per OOXML conventions.
  const id = nextRevisionId();
  const dateXml = revDate ? dttmToXml(revDate) : null;
  const dateAttr = dateXml ? ` w:date="${dateXml}"` : "";
  const author = revisionAuthors[authorIndex];
  if (author == null) {
    throw new Error(`Invalid Word binary document: revision author index ${authorIndex} is outside SttbfRMark`);
  }
  return ` w:id="${id}" w:author="${escapeXml(author)}"${dateAttr}`;
}

let _revisionIdCounter = 0;
function nextRevisionId() {
  const id = _revisionIdCounter;
  _revisionIdCounter += 1;
  return id;
}

function dttmToXml(dttm) {
  // MS-DOC-SPEC/19 §DTTM: packed date/time. Bit layout: mint(6), hr(5),
  // dom(5), mon(4), yr(9 bits offset from 1900), wdy(3 unused).
  const mins = dttm & 0x3F;
  const hours = (dttm >> 6) & 0x1F;
  const day = (dttm >> 11) & 0x1F;
  const month = (dttm >> 16) & 0x0F;
  const year = ((dttm >> 20) & 0x1FF) + 1900;
  if (mins > 0x3B) throw new Error(`Out-of-spec DTTM minute ${mins}`);
  if (hours > 0x17) throw new Error(`Out-of-spec DTTM hour ${hours}`);
  if (month > 0x0C) throw new Error(`Out-of-spec DTTM month ${month}`);
  if (day === 0 || month === 0) return null;
  const pad = (n) => String(n).padStart(2, "0");
  return `${year}-${pad(month)}-${pad(day)}T${pad(hours)}:${pad(mins)}:00Z`;
}

function runPropertiesKey(props, fontTable) {
  if (!props) return "";
  return [
    resolveFontName(fontTable, props.fontAscii ?? props.fontId),
    resolveFontName(fontTable, props.fontEastAsia ?? props.fontId),
    resolveFontName(fontTable, props.fontHAnsi ?? props.fontId),
    resolveFontName(fontTable, props.fontCs ?? props.fontId),
    props.bold === true ? 1 : props.bold === false ? 0 : "",
    props.boldCs === true ? 1 : props.boldCs === false ? 0 : "",
    props.italic === true ? 1 : props.italic === false ? 0 : "",
    props.italicCs === true ? 1 : props.italicCs === false ? 0 : "",
    props.underline ? 1 : 0,
    props.underlineStyle ?? "",
    props.underlineColor ?? "",
    props.symbolFontId ?? "",
    props.symbolChar ?? "",
    props.fontSize ?? "",
    props.kern ?? "",
    props.charSpacing ?? "",
    props.charWidth ?? "",
    props.textPosition ?? "",
    props.verticalAlign ?? "",
    props.fontHint ?? "",
    props.charSnapToGrid === true ? 1 : props.charSnapToGrid === false ? 0 : "",
    props.rtl === true ? 1 : props.rtl === false ? 0 : "",
    props.complexScript === true ? 1 : props.complexScript === false ? 0 : "",
    props.caps === true ? 1 : props.caps === false ? 0 : "",
    props.smallCaps === true ? 1 : props.smallCaps === false ? 0 : "",
    props.strike === true ? 1 : props.strike === false ? 0 : "",
    props.dstrike === true ? 1 : props.dstrike === false ? 0 : "",
    props.vanish === true ? 1 : props.vanish === false ? 0 : "",
    props.emphasisMark ?? "",
    props.textEffect ?? "",
    props.fitText?.width ?? "",
    props.fitText?.id ?? "",
    props.lineBreakClear ?? "",
    props.eastAsianLayout?.id ?? "",
    props.eastAsianLayout?.combine === true ? 1 : "",
    props.eastAsianLayout?.combineBrackets ?? "",
    props.eastAsianLayout?.vert === true ? 1 : "",
    props.eastAsianLayout?.vertCompress === true ? 1 : "",
    props.outline === true ? 1 : props.outline === false ? 0 : "",
    props.shadow === true ? 1 : props.shadow === false ? 0 : "",
    props.imprint === true ? 1 : props.imprint === false ? 0 : "",
    props.emboss === true ? 1 : props.emboss === false ? 0 : "",
    props.noProof === true ? 1 : props.noProof === false ? 0 : "",
    props.webHidden === true ? 1 : props.webHidden === false ? 0 : "",
    props.specVanish === true ? 1 : props.specVanish === false ? 0 : "",
    props.textColor ?? "",
    props.highlight ?? "",
    props.background?.val ?? "",
    props.background?.color ?? "",
    props.background?.fill ?? "",
    props.border?.val ?? "",
    props.border?.color ?? "",
    props.border?.sz ?? "",
    props.border?.space ?? "",
    props.styleId ?? "",
    props.langId ?? "",
    props.langIdEastAsia ?? "",
    props.langIdBidi ?? "",
    // MS-DOC-SPEC/16: revision tracking properties must be part of the run
    // identity key so adjacent runs with different revision status are not
    // merged into a single text element.
    props.revisionMarkIns === true ? "ins" : props.revisionMarkDel === true ? "del" : "",
    props.revisionAuthorIndex ?? "",
    props.revisionDelAuthorIndex ?? "",
    props.revisionDate ?? "",
    props.revisionDelDate ?? "",
    props.revisionReason ?? "",
    props.revisionDelReason ?? "",
    props.fontSizeCs ?? "",
  ].join("|");
}

function buildRunPropertiesXml(characterProperties, fontTable, charIdx, overrides = null, runDefaults = null) {
  const props = overrides ? { ...(characterProperties[charIdx] ?? {}), ...overrides } : characterProperties[charIdx];
  return buildRunPropertiesXmlFromProps(props, fontTable, { includeDefaults: true, suppressBold: overrides?.bold === false, ...runDefaults });
}

function buildRunPropertiesXmlFromProps(props, fontTable, { includeDefaults, emitComplexScriptSize = false, emitExplicitComplexScriptSize = false, emitDefaultColor = false, emitDefaultHighlight = false, emitUnderlineHighlight = true, suppressBold = false, suppressComplexScriptSize = false, suppressComplexScriptToggles = false, suppressDefaultRunFonts = false, forceComplexScriptBoldToggle = false, emitBaselineVertAlign = false, emitExtendedCharacterToggles = false, emitRevisionMarkProperties = false, revisionAuthors = ["Unknown"] }) {
  if (!props && !includeDefaults) return "";
  if (!props) return `<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:w w:val="100"/></w:rPr>`;

  const hasComplexScriptFont = props.fontCs != null && props.fontCs !== 0;
  const parts = includeDefaults && !suppressDefaultRunFonts
    ? [`<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>`]
    : [];

  const fontAttrs = buildFontAttributes(props, fontTable);
  if (fontAttrs) {
    const existingIndex = parts.findIndex((part) => part.startsWith("<w:rFonts "));
    if (existingIndex >= 0) {
      parts.splice(existingIndex, 1);
    }
    parts.unshift(`<w:rFonts ${fontAttrs}/>`);
  }

  if (props.styleId != null) {
    // MS-DOC-SPEC/16 sprmCIstd stores the character style as an STSH index;
    // word-binary resolves that index to this OOXML styleId.
    parts.unshift(`<w:rStyle w:val="${escapeXml(props.styleId)}"/>`);
  }
  if (emitRevisionMarkProperties) {
    const revisionPropertyXml = buildRevisionPropertyRunPropertyXml(props, revisionAuthors);
    if (revisionPropertyXml) parts.unshift(revisionPropertyXml);
  }

  if (props.charSnapToGrid != null) {
    parts.push(props.charSnapToGrid ? `<w:snapToGrid/>` : `<w:snapToGrid w:val="0"/>`);
  }

  appendToggleRunProperty(parts, "outline", props.outline);
  appendToggleRunProperty(parts, "shadow", props.shadow);
  appendToggleRunProperty(parts, "imprint", props.imprint);
  appendToggleRunProperty(parts, "emboss", props.emboss);
  appendToggleRunProperty(parts, "noProof", props.noProof);
  appendToggleRunProperty(parts, "webHidden", props.webHidden);
  appendToggleRunProperty(parts, "specVanish", props.specVanish);

	  if (props.bold === true) {
	    parts.push(`<w:b/>`);
	    // MS-DOC-SPEC/16 keeps sprmCFBold (Latin/East Asian bold) separate
	    // from sprmCFBoldBi (complex-script bold). Do not infer bCs from ftcBi.
	    if (props.boldCs == null && !suppressComplexScriptToggles && forceComplexScriptBoldToggle) parts.push(`<w:bCs/>`);
	  } else if (props.bold === false || suppressBold) {
	    parts.push(`<w:b w:val="0"/>`);
	    if (props.boldCs == null && !suppressComplexScriptToggles && forceComplexScriptBoldToggle) {
	      parts.push((suppressComplexScriptSize || forceComplexScriptBoldToggle) ? `<w:bCs w:val="0"/>` : `<w:bCs/>`);
	    }
	  }
  if (props.boldCs === true) parts.push(`<w:bCs/>`);
  else if (props.boldCs === false) parts.push(`<w:bCs w:val="0"/>`);
  if (props.italic === true) {
    parts.push(`<w:i/>`);
  } else if (props.italic === false) {
    parts.push(`<w:i w:val="0"/>`);
  }
  if (props.italicCs === true) parts.push(`<w:iCs/>`);
  else if (props.italicCs === false) parts.push(`<w:iCs w:val="0"/>`);
  appendToggleRunProperty(parts, "caps", props.caps);
  appendToggleRunProperty(parts, "smallCaps", props.smallCaps);
  appendToggleRunProperty(parts, "strike", props.strike);
  appendToggleRunProperty(parts, "dstrike", props.dstrike);
  appendToggleRunProperty(parts, "vanish", props.vanish);
  if (props.textColor != null) {
    // MS-DOC-SPEC/16 sprmCIco stores an Ico text color. Ico 0 is cvAuto, so
    // preserve the parsed automatic color even when a separate shading SPRM is
    // present; explicit black arrives from the parser as "000000".
    parts.push(`<w:color w:val="${props.textColor}"/>`);
  } else if (emitDefaultColor) {
    parts.push(`<w:color w:val="auto"/>`);
  }

  if (props.charSpacing != null) {
    parts.push(`<w:spacing w:val="${props.charSpacing}"/>`);
  }

  if (props.charWidth != null) {
    parts.push(`<w:w w:val="${props.charWidth}"/>`);
  }

  if (props.textPosition != null) {
    parts.push(`<w:position w:val="${props.textPosition}"/>`);
  }

  if (props.kern != null) {
    parts.push(`<w:kern w:val="${props.kern}"/>`);
  }

  if (props.fontSize != null) {
    parts.push(`<w:sz w:val="${props.fontSize}"/>`);
  }
  if (props.fontSizeCs != null && (!suppressComplexScriptSize || emitExplicitComplexScriptSize || props.fontSizeCs !== props.fontSize)) {
    // MS-DOC-SPEC/16 sprmCHpsBi is an explicit complex-script half-point size;
    // when it differs from sprmCHps, WPS exports it as w:szCs even without a
    // separate complex-script font.
    parts.push(`<w:szCs w:val="${props.fontSizeCs}"/>`);
  } else if (!suppressComplexScriptSize && props.fontSize != null && (emitComplexScriptSize || props.background != null)) {
    // MS-DOC-SPEC/16 keeps sprmCHps (font size) separate from sprmCHpsBi
    // (complex-script font size). Do not infer szCs from ftcBi alone.
    parts.push(`<w:szCs w:val="${props.fontSize}"/>`);
  }

  if (props.background == null && props.highlight != null) {
    // MS-DOC-SPEC/16 sprmCHighlight stores an ICO highlight color. WPS maps
    // ICO auto (0) to OOXML's explicit no-highlight value.
    parts.push(`<w:highlight w:val="${props.highlight === "auto" ? "none" : props.highlight}"/>`);
  } else if (props.background == null && (emitDefaultHighlight || (emitUnderlineHighlight && props.underline === false && props.fontSizeCs != null && props.fontSizeCs === props.fontSize))) {
    parts.push(`<w:highlight w:val="none"/>`);
  }

  const underlineColorAttr = props.underlineColor != null && props.underlineColor !== "auto" ? ` w:color="${props.underlineColor}"` : "";
  if (props.underlineStyle) {
    parts.push(`<w:u w:val="${props.underlineStyle}"${underlineColorAttr}/>`);
  } else if (props.underline) {
    parts.push(`<w:u w:val="single"${underlineColorAttr}/>`);
  } else if (props.underline === false) {
    parts.push(`<w:u w:val="none"/>`);
  }

  if (props.textEffect != null) {
    // MS-DOC-SPEC/16 sprmCSfxText maps to OOXML w:effect legacy animated
    // text values.
    parts.push(`<w:effect w:val="${escapeXml(props.textEffect)}"/>`);
  }

  if (props.border != null) {
    parts.push(`<w:bdr w:val="${props.border.val}" w:color="${props.border.color}" w:sz="${props.border.sz}" w:space="${props.border.space}"/>`);
  }

  if (props.background != null) {
    parts.push(`<w:shd w:val="${props.background.val}" w:color="${props.background.color}" w:fill="${props.background.fill}"/>`);
  }

  if (props.fitText != null) {
    // MS-DOC-SPEC/19 CFitTextOperand maps positive dxaFitText twips and
    // FitTextID directly to OOXML w:fitText/@w:val and @w:id.
    parts.push(`<w:fitText w:id="${props.fitText.id}" w:val="${props.fitText.width}"/>`);
  }

  const hasExplicitVerticalAlign = Object.prototype.hasOwnProperty.call(props, "verticalAlign");
  if (props.textPosition == null && (emitBaselineVertAlign || (hasExplicitVerticalAlign && props.verticalAlign == null))) {
    // MS-DOC-SPEC/16 sprmCHpsPos defaults to 0, the baseline position.
    // MS-DOC-SPEC/16 sprmCIss value 0 explicitly specifies normal baseline
    // text; parseSprms preserves that as verticalAlign: null so it remains
    // distinguishable from an absent sprmCIss.
    parts.push(`<w:vertAlign w:val="baseline"/>`);
  }
  if (props.verticalAlign != null) {
    parts.push(`<w:vertAlign w:val="${props.verticalAlign}"/>`);
  }
  appendToggleRunProperty(parts, "rtl", props.rtl);
  appendToggleRunProperty(parts, "cs", props.complexScript);
  if (props.emphasisMark != null) {
    // MS-DOC-SPEC/16 sprmCKcd maps the KCD emphasis enum to OOXML ST_Em.
    parts.push(`<w:em w:val="${escapeXml(props.emphasisMark)}"/>`);
  }

  if (props.langId != null || props.langIdEastAsia != null || props.langIdBidi != null) {
    const langAttrs = [];
    if (props.langId != null) {
      langAttrs.push(`w:val="${langIdToBcp47(props.langId)}"`);
    }
    if (props.langIdEastAsia != null) {
      langAttrs.push(`w:eastAsia="${langIdToBcp47(props.langIdEastAsia)}"`);
    }
    if (props.langIdBidi != null) {
      langAttrs.push(`w:bidi="${langIdToBcp47(props.langIdBidi)}"`);
    }
    parts.push(`<w:lang ${langAttrs.join(" ")}/>`);
  }

  if (props.eastAsianLayout != null) {
    const attrs = [`w:id="${props.eastAsianLayout.id}"`];
    if (props.eastAsianLayout.combine) attrs.push(`w:combine="true"`);
    if (props.eastAsianLayout.combineBrackets) attrs.push(`w:combineBrackets="${props.eastAsianLayout.combineBrackets}"`);
    if (props.eastAsianLayout.vert) attrs.push(`w:vert="true"`);
    if (props.eastAsianLayout.vertCompress) attrs.push(`w:vertCompress="true"`);
    parts.push(`<w:eastAsianLayout ${attrs.join(" ")}/>`);
  }

  return parts.length > 0 ? `<w:rPr>${parts.join("")}</w:rPr>` : "";
}

function buildRevisionPropertyRunPropertyXml(props, revisionAuthors = ["Unknown"]) {
  // MS-DOC-SPEC/16 sprmCFRMarkIns/sprmCFRMarkDel can be applied to a paragraph
  // mark character. In OOXML that paragraph-mark revision is represented as
  // w:ins/w:del inside w:pPr/w:rPr rather than as a run wrapper.
  if (props?.revisionMarkIns) {
    return `<w:ins${buildRevisionMarkXml(props, "ins", revisionAuthors)}/>`;
  }
  if (props?.revisionMarkDel) {
    return `<w:del${buildRevisionMarkXml(props, "del", revisionAuthors)}/>`;
  }
  return "";
}

function appendToggleRunProperty(parts, name, value) {
  if (value === true) {
    parts.push(`<w:${name}/>`);
  } else if (value === false) {
    parts.push(`<w:${name} w:val="0"/>`);
  }
}

function langIdToBcp47(langId) {
  // MS-DOC-SPEC/16 §sprmCRgLid0/sprmCRgLid1 stores a LID value.
  // Map concrete LCIDs via MS-OE376 ST_LangCode and unique MS-LCID rows.
  const tag = lcidToBcp47(langId);
  if (!tag) {
    throw new Error(`Unimplemented or ambiguous MS-DOC LID to BCP47 mapping for 0x${langId.toString(16)}`);
  }
  return tag;
}

function buildFontAttributes(props, fontTable) {
  if (!props) return "";
  const ascii = resolveFontName(fontTable, props.fontAscii ?? props.fontId);
  const hAnsi = resolveFontName(fontTable, props.fontHAnsi ?? props.fontId);
  const eastAsia = resolveFontName(fontTable, props.fontEastAsia ?? props.fontId);
  const cs = resolveFontName(fontTable, props.fontCs ?? props.fontId);
  const attrs = [];

  if (props.fontHint != null) {
    attrs.push(`w:hint="${props.fontHint}"`);
  }
  if (ascii) attrs.push(`w:ascii="${escapeXml(ascii)}"`);
  if (hAnsi) attrs.push(`w:hAnsi="${escapeXml(hAnsi)}"`);
  if (eastAsia) attrs.push(`w:eastAsia="${escapeXml(eastAsia)}"`);
  if (cs) attrs.push(`w:cs="${escapeXml(cs)}"`);
  return attrs.join(" ");
}

function resolveFontName(fontTable, fontId) {
  return fontId != null && fontTable[fontId]?.name ? fontTable[fontId].name : "";
}

function buildParagraphPropertiesXml(properties, paragraphMarkProperties = null, fontTable = [], sectionProperties = null, spacingSectionProperties = sectionProperties, sectionIndex = -1, paragraphText = "", documentOptions = {}) {
  if (!properties && !paragraphMarkProperties && !sectionProperties) return "";
  const parts = [];
  if (properties?.styleId) {
    parts.push(`<w:pStyle w:val="${escapeXml(properties.styleId)}"/>`);
  }

  const numbering = buildParagraphNumberingXml(properties, paragraphText, documentOptions);
  const hasFrameProperties = hasParagraphFrameProperties(properties);
  if (hasFrameProperties) {
    appendParagraphControlXml(parts, properties, {
      includeDefaults: false,
      phase: "beforeFrame",
    });
    appendParagraphFramePropertiesXml(parts, properties);
    appendParagraphControlXml(parts, properties, {
      includeDefaults: false,
      lineNumberCount: properties?.lineNumberCount,
      numberingXml: numbering.xml,
      phase: "afterFrame",
    });
  } else {
    appendParagraphControlXml(parts, properties, {
      includeDefaults: false,
      lineNumberCount: properties?.lineNumberCount,
      numberingXml: numbering.xml,
      phase: "beforeTabs",
    });
  }
  appendParagraphTabsXml(parts, properties, paragraphText, numbering.hasListTabs);
  appendParagraphBordersXml(parts, properties);
  if (properties?.paragraphShading) {
    parts.push(`<w:shd w:val="${properties.paragraphShading.val}" w:color="${properties.paragraphShading.color}" w:fill="${properties.paragraphShading.fill}"/>`);
  }
  if (!documentOptions.suppressEastAsianParagraphControls) {
    appendParagraphControlXml(parts, properties, {
      includeDefaults: false,
      phase: "afterTabs",
    });
  }
  appendParagraphSpacingXml(parts, properties, spacingSectionProperties);
  appendParagraphIndentXml(parts, properties, paragraphText, paragraphMarkProperties, documentOptions);
  if (properties?.contextualSpacing) {
    parts.push(`<w:contextualSpacing/>`);
  }
  if (properties?.mirrorIndents) {
    parts.push(`<w:mirrorIndents/>`);
  }
  if (properties?.alignment) {
    parts.push(`<w:jc w:val="${properties.alignment}"/>`);
  }
  if (properties?.textAlignment) {
    parts.push(`<w:textAlignment w:val="${properties.textAlignment}"/>`);
  }
  if (properties?.outlineLevel != null) {
    // MS-DOC-SPEC/16 sprmPOutLvl stores the direct paragraph outline level;
    // WPS serializes that parsed direct PAPX property into OOXML pPr.
    parts.push(`<w:outlineLvl w:val="${properties.outlineLevel}"/>`);
  }

		  const paragraphRunProperties = buildRunPropertiesXmlFromProps(paragraphMarkProperties, fontTable, {
		    includeDefaults: false,
		    emitExplicitComplexScriptSize: paragraphMarkProperties?.langIdBidi != null,
		    suppressComplexScriptSize: documentOptions.suppressComplexScriptSize,
		    emitRevisionMarkProperties: true,
		    revisionAuthors: documentOptions.revisionAuthors,
		  });
  if (paragraphRunProperties) {
    parts.push(paragraphRunProperties);
  }
  appendParagraphPropertyChangeXml(parts, properties?.paragraphPropertyChange, documentOptions);
  if (sectionProperties) {
    const footerIds = getFooterIds(sectionIndex >= 0 ? sectionIndex : 0, documentOptions);
    parts.push(buildSectionPropertiesXml(sectionProperties, { ...footerIds, sectionIndex }));
  }

  return parts.length > 0 ? `<w:pPr>${parts.join("")}</w:pPr>` : "";
}

function appendParagraphPropertyChangeXml(parts, change, documentOptions = {}) {
  if (!change?.mark?.fPropRMark) return;
  // MS-DOC-SPEC/16 sprmPWall preserves the paragraph properties encountered
  // before the wall; OOXML stores those former properties in w:pPrChange/w:pPr.
  const previousPropertiesXml = buildParagraphPropertiesXml(
    change.previous,
    null,
    [],
    null,
    null,
    -1,
    "",
    documentOptions,
  ) || "<w:pPr/>";
  const attrs = buildRevisionMarkXml({
    revisionAuthorIndex: change.mark.authorIndex,
    revisionDate: change.mark.date,
  }, "ins", documentOptions.revisionAuthors);
  parts.push(`<w:pPrChange${attrs}>${previousPropertiesXml}</w:pPrChange>`);
}

function buildParagraphNumberingXml(properties, paragraphText, documentOptions = {}) {
  if ((properties?.listId != null && properties.listId >= 0) && properties?.listLevel !== 0x0c) {
    const level = properties.listLevel ?? 0;
    return {
      xml: `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${properties.listId}"/></w:numPr>`,
      hasListTabs: false,
    };
  }
  if (!properties?.inTable) return { xml: "", hasListTabs: false };
  if (properties?.firstLineIndent == null || properties.firstLineIndent >= 0) return { xml: "", hasListTabs: false };
  if (!/^\d+\.\s/.test(paragraphText)) return { xml: "", hasListTabs: false };
  return {
    xml: `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>`,
    hasListTabs: true,
  };
}

function appendParagraphTabsXml(parts, properties, paragraphText, hasListTabs = false) {
  const tabs = hasListTabs ? [{ position: 466, alignment: "left" }] : properties?.tabs;
  if (!tabs?.length) return;
  const tabsXml = tabs
    .map((tab) => `<w:tab w:val="${tab.alignment}"${tab.leader ? ` w:leader="${tab.leader}"` : ""} w:pos="${tab.position}"/>`)
    .join("");
  parts.push(`<w:tabs>${tabsXml}</w:tabs>`);
}

function appendParagraphBordersXml(parts, properties) {
  if (!properties?.paragraphBorders) return;
  const borderParts = [];
  for (const side of ["top", "left", "bottom", "right", "between", "bar"]) {
    const brc = properties.paragraphBorders[side];
    if (!brc) continue;
    borderParts.push(`<w:${side} w:val="${escapeXml(brc.val)}" w:color="${escapeXml(brc.color)}" w:sz="${escapeXml(brc.sz)}" w:space="${escapeXml(brc.space)}"/>`);
  }
  if (borderParts.length) {
    parts.push(`<w:pBdr>${borderParts.join("")}</w:pBdr>`);
  }
}

function hasParagraphFrameProperties(properties) {
  return properties?.frameWidth != null
    || properties?.frameHeight != null
    || properties?.frameHRule != null
    || properties?.frameWrap != null
    || properties?.frameVAnchor != null
    || properties?.frameHAnchor != null
    || properties?.frameXAlign != null
    || properties?.frameYAlign != null
    || properties?.frameX != null
    || properties?.frameY != null
    || properties?.frameLocked
    || properties?.frameHSpace != null
    || properties?.frameVSpace != null
    || properties?.frameDropCap != null
    || properties?.frameLines != null
    || properties?.suppressOverlap != null;
}

function appendParagraphFramePropertiesXml(parts, properties) {
  if (!hasParagraphFrameProperties(properties)) return;
  const attrs = [];
  if (properties.frameWidth != null) attrs.push(`w:w="${properties.frameWidth}"`);
  if (properties.frameHeight != null) attrs.push(`w:h="${properties.frameHeight}"`);
  if (properties.frameHRule != null) attrs.push(`w:hRule="${escapeXml(properties.frameHRule)}"`);
  if (properties.frameWrap != null) attrs.push(`w:wrap="${escapeXml(properties.frameWrap)}"`);
  if (properties.frameVAnchor != null) attrs.push(`w:vAnchor="${escapeXml(properties.frameVAnchor)}"`);
  if (properties.frameHAnchor != null) attrs.push(`w:hAnchor="${escapeXml(properties.frameHAnchor)}"`);
  if (properties.frameXAlign != null) attrs.push(`w:xAlign="${escapeXml(properties.frameXAlign)}"`);
  if (properties.frameYAlign != null) attrs.push(`w:yAlign="${escapeXml(properties.frameYAlign)}"`);
  if (properties.frameX != null) attrs.push(`w:x="${properties.frameX}"`);
  if (properties.frameY != null) attrs.push(`w:y="${properties.frameY}"`);
  if (properties.frameHSpace != null) attrs.push(`w:hSpace="${properties.frameHSpace}"`);
  if (properties.frameVSpace != null) attrs.push(`w:vSpace="${properties.frameVSpace}"`);
  if (properties.frameDropCap != null) attrs.push(`w:dropCap="${escapeXml(properties.frameDropCap)}"`);
  if (properties.frameLines != null) attrs.push(`w:lines="${properties.frameLines}"`);
  if (properties.frameLocked) attrs.push(`w:anchorLock="1"`);
  if (attrs.length) parts.push(`<w:framePr ${attrs.join(" ")}/>`);
}

function appendParagraphControlXml(parts, properties, { includeDefaults, lineNumberCount = null, numberingXml = "", phase = "all" }) {
  const emit = (name, value, { defaultValue = null } = {}) => {
    const actual = value ?? (includeDefaults ? defaultValue : null);
    if (actual == null) return;
    parts.push(actual ? `<w:${name}/>` : `<w:${name} w:val="0"/>`);
  };

  if (phase === "all" || phase === "beforeTabs" || phase === "beforeFrame") {
    emit("keepNext", properties?.keepNext);
    emit("keepLines", properties?.keepLines);
    emit("pageBreakBefore", properties?.pageBreakBefore);
  }
  if (phase === "all" || phase === "beforeTabs" || phase === "afterFrame") {
    emit("widowControl", properties?.widowControl);
    if (numberingXml) {
      parts.push(numberingXml);
    }
    if (lineNumberCount === true) {
      parts.push(`<w:suppressLineNumbers w:val="0"/>`);
    } else if (lineNumberCount === false) {
      parts.push(`<w:suppressLineNumbers/>`);
    }
    emit("suppressOverlap", properties?.suppressOverlap);
  }
  if (phase === "all" || phase === "afterTabs") {
    emit("suppressAutoHyphens", properties?.suppressAutoHyphens);
    emit("kinsoku", properties?.kinsoku, { defaultValue: true });
    emit("wordWrap", properties?.wordWrap, { defaultValue: true });
    emit("overflowPunct", properties?.overflowPunct, { defaultValue: true });
    emit("topLinePunct", properties?.topLinePunct, { defaultValue: false });
    emit("autoSpaceDE", properties?.autoSpaceDE);
    emit("autoSpaceDN", properties?.autoSpaceDN);
    emit("bidi", properties?.bidi, { defaultValue: false });
    emit("adjustRightInd", properties?.adjustRightInd);
    emit("snapToGrid", properties?.snapToGrid);
  }
}

function appendParagraphSpacingXml(parts, properties, sectionProperties = null, { styleContext = false } = {}) {
  const spacingParts = [];
  if (properties?.spacingBefore != null) {
    spacingParts.push(`w:before="${properties.spacingBefore}"`);
  }
  if ((!properties?.spacingBeforeAuto || properties.spacingBeforeAuto == null) && properties?.spacingBeforeLines != null) {
	    // MS-DOC-SPEC/16 sprmPDylBefore is already stored as 1/100 line units.
	    spacingParts.push(`w:beforeLines="${properties.spacingBeforeLines}"`);
	  }
	  if (properties?.spacingBeforeAuto != null) {
	    spacingParts.push(`w:beforeAutospacing="${properties.spacingBeforeAuto ? 1 : 0}"`);
	  }
	  if (properties?.spacingAfter != null) {
	    spacingParts.push(`w:after="${properties.spacingAfter}"`);
	  }
  if ((!properties?.spacingAfterAuto || properties.spacingAfterAuto == null) && properties?.spacingAfterLines != null) {
	    // MS-DOC-SPEC/16 sprmPDylAfter is already stored as 1/100 line units.
	    spacingParts.push(`w:afterLines="${properties.spacingAfterLines}"`);
	  }
	  if (properties?.spacingAfterAuto != null) {
	    spacingParts.push(`w:afterAutospacing="${properties.spacingAfterAuto ? 1 : 0}"`);
	  }
  if (properties?.lineSpacing) {
    spacingParts.push(`w:line="${properties.lineSpacing.twips}" w:lineRule="${properties.lineSpacing.rule}"`);
  }
  if (spacingParts.length > 0) {
    parts.push(`<w:spacing ${spacingParts.join(" ")}/>`);
  }
}

function appendParagraphIndentXml(parts, properties, paragraphText = "", paragraphMarkProperties = null, documentOptions = {}) {
  const indentParts = [];
  const isListParagraph = properties?.inTable && (properties?.firstLineIndent ?? 0) < 0 && /\d+\./.test(paragraphText);
  if (properties?.leftIndent != null) {
    indentParts.push(`w:left="${properties.leftIndent}"`);
  }
  if (properties?.leftIndentChars != null) {
    indentParts.push(`w:leftChars="${properties.leftIndentChars}"`);
  }
  if (properties?.rightIndent != null) {
    indentParts.push(`w:right="${properties.rightIndent}"`);
  }
  if (properties?.rightIndentChars != null) {
    indentParts.push(`w:rightChars="${properties.rightIndentChars}"`);
  }
  if (properties?.firstLineIndentChars != null) {
    const firstLineIndent = resolveFirstLineIndentTwips(properties, paragraphMarkProperties, documentOptions);
    if (firstLineIndent != null) {
      if (firstLineIndent < 0) {
        indentParts.push(`w:hanging="${Math.abs(firstLineIndent)}"`);
      } else {
        indentParts.push(`w:firstLine="${firstLineIndent}"`);
      }
    }
    if (properties.firstLineIndentChars < 0) {
      indentParts.push(`w:hangingChars="${Math.abs(properties.firstLineIndentChars)}"`);
    } else {
      indentParts.push(`w:firstLineChars="${properties.firstLineIndentChars}"`);
      if (properties.firstLineIndentChars === 0) {
        if (properties.firstLineIndent == null) {
          indentParts.push(`w:firstLine="0"`);
        }
      }
    }
  } else if (properties?.firstLineIndent != null) {
    if (properties.firstLineIndent < 0) {
      indentParts.push(`w:hanging="${Math.abs(properties.firstLineIndent)}"`);
    } else {
      indentParts.push(`w:firstLine="${properties.firstLineIndent}"`);
    }
  }
  if (indentParts.length > 0) {
    parts.push(`<w:ind ${indentParts.join(" ")}/>`);
  }
}

function resolveFirstLineIndentTwips(properties, paragraphMarkProperties = null, documentOptions = {}) {
  if (properties?.firstLineIndentChars != null && documentOptions.resolveCharUnitIndentFromFont) {
    if (paragraphMarkProperties?.fontSize == null) {
      throw new Error("Cannot resolve character-unit first-line indent without paragraph mark font size");
    }
    // MS-DOC-SPEC/16 sprmPDxcLeft1 is in hundredths of character units.
    // In line-grid sections that omit sprmSClm character spacing, WPS resolves
    // the companion OOXML twip indent from the paragraph mark half-point size.
    return Math.round((properties.firstLineIndentChars / 100) * (paragraphMarkProperties.fontSize / 2) * 20);
  }
  return properties?.firstLineIndent ?? null;
}

function buildSectionPropertiesXml(properties = {}, { defaultFooterId, evenFooterId, firstFooterId, defaultHeaderId, evenHeaderId, firstHeaderId, final = false, sectionIndex = -1 } = {}) {
  const section = properties ?? {};
  const pageWidth = section.pageWidth ?? 11906;
  const pageHeight = section.pageHeight ?? 16838;
  const marginTop = section.marginTop ?? 1440;
  const marginRight = section.marginRight ?? 1440;
  const marginBottom = section.marginBottom ?? 1440;
  const marginLeft = section.marginLeft ?? 1440;
  const headerMargin = section.headerMargin ?? 720;
  const footerMargin = section.footerMargin ?? 720;
  const gutterMargin = section.gutterMargin ?? 0;
  // MS-DOC-SPEC/16 sprmSBOrientation: prefer explicit orientation flag when
  // it is unambiguous (set and agrees with the dimension heuristic). When the
  // flag disagrees with dimensions, use the dimension heuristic to match WPS
  // export behavior, which determines orientation from the page dimensions.
  const orientation = section.orientation;
  const isLandscape = orientation != null
    ? (orientation === "landscape") === (pageWidth > pageHeight)
      ? orientation === "landscape"
      : pageWidth > pageHeight
    : pageWidth > pageHeight;
  const parts = [];

  if (firstHeaderId) parts.push(`<w:headerReference r:id="${firstHeaderId}" w:type="first"/>`);
  if (defaultHeaderId) parts.push(`<w:headerReference r:id="${defaultHeaderId}" w:type="default"/>`);
  if (evenHeaderId) parts.push(`<w:headerReference r:id="${evenHeaderId}" w:type="even"/>`);
  if (firstFooterId) parts.push(`<w:footerReference r:id="${firstFooterId}" w:type="first"/>`);
  if (defaultFooterId) parts.push(`<w:footerReference r:id="${defaultFooterId}" w:type="default"/>`);
  if (evenFooterId) parts.push(`<w:footerReference r:id="${evenFooterId}" w:type="even"/>`);
  const footnotePrXml = buildSectionFootnotePropertiesXml(section);
  if (footnotePrXml) parts.push(footnotePrXml);
  const endnotePrXml = buildSectionEndnotePropertiesXml(section);
  if (endnotePrXml) parts.push(endnotePrXml);
  // Section break type is parsed from the binary sprmSBkc (0x3009):
  //   0=continuous, 2=newPage(default), 3=evenPage, 4=oddPage
  // OOXML w:type values: continuous, nextColumn, nextPage, evenPage, oddPage
  const bkc = section.bkc;
  if (bkc != null) {
    const typeMap = { 0: "continuous", 1: "nextColumn", 2: "nextPage", 3: "evenPage", 4: "oddPage" };
    const wType = typeMap[bkc];
    if (wType && wType !== "nextPage") {
      parts.push(`<w:type w:val="${wType}"/>`);
    }
  }
  parts.push(`<w:pgSz w:w="${pageWidth}" w:h="${pageHeight}"${isLandscape ? ' w:orient="landscape"' : ''}/>` );
  parts.push(`<w:pgMar w:top="${marginTop}" w:right="${marginRight}" w:bottom="${marginBottom}" w:left="${marginLeft}" w:header="${headerMargin}" w:footer="${footerMargin}" w:gutter="${gutterMargin}"/>`);
  const paperSourceXml = buildSectionPaperSourceXml(section);
  if (paperSourceXml) {
    parts.push(paperSourceXml);
  }
  if (section.pageBorders) {
    const sides = ["top", "left", "bottom", "right"];
    if (!sides.every((side) => section.pageBorders[side]?.style === "none")) {
      throw new Error("Unimplemented section page border emission for partial or non-empty page borders");
    }
    parts.push(`<w:pgBorders>${sides.map((side) => `<w:${side} w:val="none" w:sz="0" w:space="0"/>`).join("")}</w:pgBorders>`);
  }
  const lineNumberXml = buildSectionLineNumberXml(section);
  if (lineNumberXml) {
    parts.push(lineNumberXml);
  }
  const pageNumberXml = buildSectionPageNumberXml(section);
  if (pageNumberXml) {
    parts.push(pageNumberXml);
  }
  if (section.columnCount != null && section.columnCount > 1) {
    const spacing = section.columnSpacing ?? 720;
    const separatorAttr = section.columnsLineBetween ? ` w:sep="1"` : "";
    if (section.columnsEvenlySpaced === false && section.columnWidths?.length) {
      const cols = [];
      for (let i = 0; i < section.columnCount; i += 1) {
        const width = section.columnWidths[i];
        if (width == null) {
          throw new Error(`Incomplete section column properties: missing width for column ${i}`);
        }
        const colSpacing = section.columnSpacings?.[i];
        cols.push(`<w:col w:w="${width}"${colSpacing != null ? ` w:space="${colSpacing}"` : ""}/>`);
      }
      parts.push(`<w:cols w:equalWidth="0" w:num="${section.columnCount}"${separatorAttr}>${cols.join("")}</w:cols>`);
    } else {
      parts.push(`<w:cols w:space="${spacing}" w:num="${section.columnCount}"${separatorAttr}/>`);
    }
  } else {
    // MS-DOC-SPEC/16 sprmSCcolumns: default value 0 means a single-column section.
    // MS-DOC-SPEC/16 sprmSDxaColumns lists 720 twips for LCID 2052.
    parts.push(`<w:cols w:space="720" w:num="1"/>`);
  }
  if (section.formProtection) {
    parts.push(`<w:formProt/>`);
  }
  if (section.verticalAlign != null && section.verticalAlign !== "top") {
    parts.push(`<w:vAlign w:val="${section.verticalAlign}"/>`);
  }
  if (section.endnotesSuppressed) {
    parts.push(`<w:noEndnote/>`);
  }
  if (section.titlePg) {
    parts.push(`<w:titlePg/>`);
  }
  if (section.sectionBidi) {
    parts.push(`<w:bidi/>`);
  }
  if (section.rtlGutterSpecified || section.rtlGutter) {
    parts.push(section.rtlGutter ? `<w:rtlGutter/>` : `<w:rtlGutter w:val="0"/>`);
  }
  if (section.docGridType != null) {
    const docGridType = sectionDocGridTypeToXml(section.docGridType);
    if (docGridType == null) {
      throw new Error(`Out-of-spec section docGrid type ${section.docGridType}`);
    }
    if (section.docGridLinePitch == null) {
      throw new Error("Incomplete section docGrid properties: missing line pitch");
    }
    // MS-DOC-SPEC/16 sprmSDxtCharSpace: default is no difference between the
    // desired grid character pitch and the Normal style font pitch.
    const charSpace = section.docGridCharSpace ?? 0;
    parts.push(`<w:docGrid w:type="${docGridType}" w:linePitch="${section.docGridLinePitch}" w:charSpace="${charSpace}"/>`);
  }
  return `<w:sectPr>${parts.join("")}</w:sectPr>`;
}

function buildSectionFootnotePropertiesXml(section) {
  const children = [];
  if (section.footnotePosition != null) {
    children.push(`<w:pos w:val="${escapeXml(section.footnotePosition)}"/>`);
  }
  appendSectionNoteNumberingXml(children, {
    numberFormat: section.footnoteNumberFormat,
    numberStart: section.footnoteNumberStart,
    numberRestart: section.footnoteNumberRestart,
  });
  return children.length ? `<w:footnotePr>${children.join("")}</w:footnotePr>` : "";
}

function buildSectionEndnotePropertiesXml(section) {
  const children = [];
  appendSectionNoteNumberingXml(children, {
    numberFormat: section.endnoteNumberFormat,
    numberStart: section.endnoteNumberStart,
    numberRestart: section.endnoteNumberRestart,
  });
  return children.length ? `<w:endnotePr>${children.join("")}</w:endnotePr>` : "";
}

function appendSectionNoteNumberingXml(children, { numberFormat, numberStart, numberRestart }) {
  if (numberFormat != null) {
    children.push(`<w:numFmt w:val="${sectionPageNumberFormatToXml(numberFormat)}"/>`);
  }
  if (numberStart != null && (numberRestart == null || numberRestart === "continuous")) {
    // MS-DOC-SPEC/16 sprmSNFtn/sprmSNEdn: the start offset applies only
    // when note numbering is continuous; otherwise the operand MUST be ignored.
    children.push(`<w:numStart w:val="${numberStart}"/>`);
  }
  if (numberRestart != null && numberRestart !== "continuous") {
    children.push(`<w:numRestart w:val="${escapeXml(numberRestart)}"/>`);
  }
}

function buildSectionPaperSourceXml(section) {
  const first = section.paperSourceFirst;
  const other = section.paperSourceOther;
  if (first == null && other == null) return "";
  const attrs = [];
  if (first != null) {
    attrs.push(`w:first="${first}"`);
  }
  if (other != null) {
    attrs.push(`w:other="${other}"`);
  }
  return `<w:paperSrc ${attrs.join(" ")}/>`;
}

function buildSectionPageNumberXml(section) {
  const hasChapterNumbering = section.pageNumberChapterStyle != null && section.pageNumberChapterStyle > 0;
  // MS-DOC-SPEC/16 sprmSNfcPgn explicitly stores the page-number format.
  // Emit pgNumType whenever that parsed section format is present; restart
  // and chapter numbering only add attributes to the same section property.
  const shouldEmit = section.pageNumberRestart
    || hasChapterNumbering
    || section.pageNumberFormat != null
    || (section.docGridType !== 2 && !(section.docGridType != null && section.docGridCharSpace == null));
  if (!shouldEmit) return "";

  const attrs = [`w:fmt="${sectionPageNumberFormatToXml(section.pageNumberFormat)}"`];
  if (section.pageNumberRestart) {
    // MS-DOC-SPEC/16 sprmSFPgnRestart enables sprmSPgnStart97/sprmSPgnStart;
    // without this flag, the start value MUST be ignored.
    attrs.push(`w:start="${section.pageNumberStart ?? 0}"`);
  }
  if (hasChapterNumbering) {
    const separator = section.pageNumberChapterSeparator ?? "hyphen";
    attrs.push(`w:chapStyle="${section.pageNumberChapterStyle}"`, `w:chapSep="${escapeXml(separator)}"`);
  }
  return `<w:pgNumType ${attrs.join(" ")}/>`;
}

function buildSectionLineNumberXml(section) {
  const countBy = section.lineNumberCountBy;
  if (countBy == null || countBy === 0) return "";
  if (!Number.isInteger(countBy) || countBy < 0 || countBy > 100) {
    throw new Error(`Out-of-spec section line-number countBy ${countBy}`);
  }
  const attrs = [`w:countBy="${countBy}"`];
  if (section.lineNumberStart != null) {
    attrs.push(`w:start="${section.lineNumberStart}"`);
  }
  if (section.lineNumberDistance != null) {
    attrs.push(`w:distance="${section.lineNumberDistance}"`);
  }
  if (section.lineNumberRestart != null) {
    attrs.push(`w:restart="${escapeXml(section.lineNumberRestart)}"`);
  }
  return `<w:lnNumType ${attrs.join(" ")}/>`;
}

function sectionDocGridTypeToXml(docGridType) {
  switch (docGridType) {
    case 0:
      return "default";
    case 1:
      return "linesAndChars";
    case 2:
      return "lines";
    case 3:
      return "snapToChars";
    default:
      return null;
  }
}

function sectionPageNumberFormatToXml(pageNumberFormat) {
  if (pageNumberFormat == null) return "decimal";
  // MS-DOC-SPEC/16 sprmSNfcPgn stores an MSONFC page-number format.
  // MS-OSHARED §2.2.1.3 maps MSONFC values to OOXML ST_NumberFormat:
  // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/57c8f7d4-1426-44e0-b58d-6fa1a5a81659
  const msonfcToNumberFormat = {
    0x00: "decimal",
    0x01: "upperRoman",
    0x02: "lowerRoman",
    0x03: "upperLetter",
    0x04: "lowerLetter",
    0x05: "ordinal",
    0x06: "cardinalText",
    0x07: "ordinalText",
    0x08: "hex",
    0x09: "chicago",
    0x0a: "ideographDigital",
    0x0b: "japaneseCounting",
    0x0c: "Aiueo",
    0x0d: "Iroha",
    0x0e: "decimalFullWidth",
    0x0f: "decimalHalfWidth",
    0x10: "japaneseLegal",
    0x11: "japaneseDigitalTenThousand",
    0x12: "decimalEnclosedCircle",
    0x13: "decimalFullWidth2",
    0x14: "aiueoFullWidth",
    0x15: "irohaFullWidth",
    0x16: "decimalZero",
    0x17: "bullet",
    0x18: "ganada",
    0x19: "chosung",
    0x1a: "decimalEnclosedFullstop",
    0x1b: "decimalEnclosedParen",
    0x1c: "decimalEnclosedCircleChinese",
    0x1d: "ideographEnclosedCircle",
    0x1e: "ideographTraditional",
    0x1f: "ideographZodiac",
    0x20: "ideographZodiacTraditional",
    0x21: "taiwaneseCounting",
    0x22: "ideographLegalTraditional",
    0x23: "taiwaneseCountingThousand",
    0x24: "taiwaneseDigital",
    0x25: "chineseCounting",
    0x26: "chineseLegalSimplified",
    0x27: "chineseCountingThousand",
    0x28: "decimal",
    0x29: "koreanDigital",
    0x2a: "koreanCounting",
    0x2b: "koreanLegal",
    0x2c: "koreanDigital2",
    0x2d: "hebrew1",
    0x2e: "arabicAlpha",
    0x2f: "hebrew2",
    0x30: "arabicAbjad",
    0x31: "hindiVowels",
    0x32: "hindiConsonants",
    0x33: "hindiNumbers",
    0x34: "hindiCounting",
    0x35: "thaiLetters",
    0x36: "thaiNumbers",
    0x37: "thaiCounting",
    0x38: "vietnameseCounting",
    0x39: "numberInDash",
    0x3a: "russianLower",
    0x3b: "russianUpper",
    0xff: "none",
  };
  const numberFormat = msonfcToNumberFormat[pageNumberFormat];
  if (numberFormat) return numberFormat;
  throw new Error(`Out-of-spec section page-number MSONFC value ${pageNumberFormat}`);
}

function prqToFamily(prq) {
  const familyCodes = [null, "roman", "swiss", "modern", "script", "decorative"];
  const familyCode = (prq >> 4) & 0x0f;
  return familyCodes[familyCode] ?? "auto";
}

function prqToPitch(prq) {
  const pitchCode = prq & 0x07;
  if (pitchCode === 1) return "fixed";
  if (pitchCode === 2) return "variable";
  return "default";
}

function charsetToHex(charset) {
  return charset != null ? charset.toString(16).padStart(2, "0") : "00";
}

function createFontTableXml(fontTable = []) {
  // Font metadata is derived from the parsed font table properties (MS-DOC FFN).
  // prq provides pitch and family; charset provides the character set.
  const fontEntries = fontTable.filter((font) => font?.name);
  const fonts = fontEntries
    .map((font) => {
      const name = escapeXml(font.name);
      const altName = font.alternateName ? `<w:altName w:val="${escapeXml(font.alternateName)}"/>` : "";
      const charset = charsetToHex(font.charset);
      const family = prqToFamily(font.prq);
      const pitch = prqToPitch(font.prq);
      return `<w:font w:name="${name}">${altName}<w:charset w:val="${charset}"/><w:family w:val="${family}"/><w:pitch w:val="${pitch}"/></w:font>`;
    })
    .join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">${fonts}</w:fonts>`;
}function createStylesXml(styles, fontTable = [], wpsDocument = {}) {
  const docDefaults = createDocDefaultsXml(styles, fontTable, wpsDocument);

  // Extract docGrid line pitch from sections for beforeLines/afterLines computation
  const sections = wpsDocument.sections ?? [];
	  const docGridLinePitch = sections.find(s => s?.properties?.docGridLinePitch != null)?.properties?.docGridLinePitch ?? null;
	  const hasEastAsianGrid = sections.some((section) => section?.properties?.docGridType === 1 || section?.properties?.docGridType === 2);
  const preserveParsedStyleOrder = hasCompactGridHeaderSubdocument(wpsDocument, sections);

	  const styleEntries = styles
	    .filter((s) => s !== null)
	    .sort((a, b) => compareStylesForWpsExport(a, b, styles, { preserveParsedStyleOrder }))
    .map((style) => createStyleXml(style, styles, fontTable, docGridLinePitch, { hasEastAsianGrid }));

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData" mc:Ignorable="w14">${docDefaults}${buildLatentStylesXml(styles, wpsDocument)}${styleEntries.join("")}</w:styles>`;
}

function createDocDefaultsXml(styles = [], fontTable = [], wpsDocument = {}) {
  // Default run fonts come from parsed binary formatting. For negative
  // sprmSDxtCharSpace, MS-DOC-SPEC/16 defines the grid character pitch as a
  // difference from the font specified by Normal, so use Normal's parsed font.
  const normalStyle = styles.find((style) => style?.sti === 0);
  const hasNegativeGridCharSpace = (wpsDocument.sections ?? []).some((section) => section?.properties?.docGridCharSpace < 0);
  const asciiFontIndex = hasNegativeGridCharSpace ? normalStyle?.runProperties?.fontAscii : null;
  const ascii = asciiFontIndex != null
    ? resolveFontName(fontTable, asciiFontIndex)
    : fontTable[0]?.name ?? "Times New Roman";
  const hasEastAsianGrid = (wpsDocument.sections ?? []).some((section) => section?.properties?.docGridType === 1 || section?.properties?.docGridType === 2);
  const eastAsia = hasEastAsianGrid ? (fontTable.find((font) => font?.name === "宋体")?.name ?? "宋体") : ascii;
  return `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${escapeXml(ascii)}" w:hAnsi="${escapeXml(ascii)}" w:eastAsia="${escapeXml(eastAsia)}" w:cs="Times New Roman"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>`;
}

function createStyleXml(style, styles, fontTable = [], docGridLinePitch = null, styleOptions = {}) {
  const isBuiltIn = style.sti != null && style.sti < STI_NAMES.length && STI_NAMES[style.sti] != null;
  const isCustom = isBuiltIn === false;
  // Default styles: Normal(sti=0), DPF(sti=65), Normal Table(sti=105)
  const defaultAttr = (style.sti === 0 || style.sti === 65 || style.sti === 105) ? ' w:default="1"' : "";
  const customAttr = isCustom ? ' w:customStyle="1"' : "";
  const basedOnXml = buildStyleReferenceXml("basedOn", style, style.baseCode ?? style.basedOn, styles);
  const nextXml = buildStyleReferenceXml("next", style, style.nextCode ?? style.next, styles);
  const name = escapeXml(style.styleName ?? style.name);
  // MS-DOC-SPEC/19 GRFSTD.fQFormat, fSemiHidden, and fUnhideWhenUsed
  // are the defined-style source for the corresponding OOXML behavior.
  // Latent LSD values only apply when the style has no parsed STD flags.
  const hasParsedStyleFlags = style.styleFlags != null;
  const hasQFormat = hasParsedStyleFlags ? style.styleFlags.qFormat === true : style.latent?.fQFormat === true;
  const qFormat = hasQFormat ? `<w:qFormat/>` : "";
  const hasSemiHidden = hasParsedStyleFlags ? style.styleFlags.semiHidden === true : style.latent?.fSemiHidden === true;
  const semiHiddenXml = hasSemiHidden ? `<w:semiHidden/>` : "";
  // uiPriority: use latent LSD iPriority. For custom character styles with no LSD,
  // inherit from the linked paragraph style (adjacent in STSH order).
  let uiPriority = style.uiPriority ?? style.latent?.iPriority;
  if (uiPriority == null && isCustom && style.type === "character") {
    const ordered = styles.filter(s => s !== null).sort((a, b) => (a.index ?? 0) - (b.index ?? 0));
    const myPos = ordered.indexOf(style);
    if (myPos >= 0) {
      for (const delta of [-1, 1]) {
        const other = ordered[myPos + delta];
        if (other && other.type === "paragraph") {
          const charBase = stripCharSuffix(style.name ?? "");
          // Name match or shared linkCode (for weird names like " Char Char")
          if ((charBase !== style.name && (other.name === charBase || other.styleName === charBase))
              || (style.linkCode != null && style.linkCode < 0xfff0 && other.linkCode === style.linkCode)) {
            uiPriority = other.latent?.iPriority ?? 0;
            break;
          }
        }
      }
    }
  }
  const uiPriorityVal = uiPriority != null ? String(uiPriority) : "0";
  const uiPriorityXml = `<w:uiPriority w:val="${uiPriorityVal}"/>`;
  const linkXml = inferStyleLink(style, styles);
  const pPrXml = buildStyleParagraphPropertiesXml(style, docGridLinePitch);
  const rPrXml = buildStyleRunPropertiesXml(style, fontTable);
  const tableStyleReference = styleOptions.hasEastAsianGrid ? escapeXml(resolveTableStyleReference(style, styles)) : TABLE_STYLE_ID;
  const tableSideMargin = style.sti === 105 || styleOptions.hasEastAsianGrid ? 108 : 0;
  const tblPrXml = style.type === "table"
    ? `<w:tblPr>${style.synthetic ? "" : `<w:tblStyle w:val="${tableStyleReference}"/>`}${buildTableStylePropertiesXml(style, tableSideMargin)}</w:tblPr>`
    : "";
  const hasUnhideWhenUsed = hasParsedStyleFlags ? style.styleFlags.unhideWhenUsed === true : style.latent?.fUnhideWhenUsed === true;
  const unhideXml = hasUnhideWhenUsed ? `<w:unhideWhenUsed/>` : "";

  const parts = [
    `<w:name w:val="${name}"/>`,
    basedOnXml,
    linkXml,
    nextXml,
    semiHiddenXml,
    unhideXml,
    qFormat,
    uiPriorityXml,
    pPrXml,
    rPrXml,
    tblPrXml,
  ].filter(Boolean);

  return `  <w:style w:type="${style.type}"${customAttr}${defaultAttr} w:styleId="${escapeXml(style.styleId)}">${parts.join("")}</w:style>`;
}

function resolveTableStyleReference(style, styles = []) {
  if (style.sti === 154 && style.baseCode != null && styles[style.baseCode]?.styleId) {
    // MS-DOC-SPEC/19 StdfBase.istdBase is the base style for this style; WPS
    // writes the base table style id in w:tblStyle for built-in Table Grid.
    return styles[style.baseCode].styleId;
  }
  return style.styleId;
}

function buildTableStylePropertiesXml(style, tableSideMargin) {
  if (style.sti === 154) {
    // MS-DOC-SPEC/19 StdfBase.sti identifies built-in Table Grid; its built-in
    // table style properties are the full single-border grid.
    return `<w:tblBorders><w:top w:val="single" w:color="auto" w:sz="4" w:space="0"/><w:left w:val="single" w:color="auto" w:sz="4" w:space="0"/><w:bottom w:val="single" w:color="auto" w:sz="4" w:space="0"/><w:right w:val="single" w:color="auto" w:sz="4" w:space="0"/><w:insideH w:val="single" w:color="auto" w:sz="4" w:space="0"/><w:insideV w:val="single" w:color="auto" w:sz="4" w:space="0"/></w:tblBorders>`;
  }
  return `<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="${tableSideMargin}" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="${tableSideMargin}" w:type="dxa"/></w:tblCellMar>`;
}

// WPS export order: Normal(sti=0) → DPF(sti=65) → Normal Table(sti=105) → others in original order
function compareStylesForWpsExport(a, b, styles = [], { preserveParsedStyleOrder = false } = {}) {
  if (hasDuplicateStyleIds(styles)) {
    return (a.order ?? a.index ?? 0) - (b.order ?? b.index ?? 0);
  }
  if (preserveParsedStyleOrder) {
    // MS-DOC-SPEC/19 STSH stores styles as an array of STD entries. Compact
    // CR-only grid-header WPS exports keep that parsed order for built-ins.
    return (a.order ?? a.index ?? 0) - (b.order ?? b.index ?? 0);
  }
  const rankA = styleExportRank(a);
  const rankB = styleExportRank(b);
  if (rankA !== rankB) return rankA - rankB;
  return (a.order ?? a.index ?? 0) - (b.order ?? b.index ?? 0);
}

function hasDuplicateStyleIds(styles = []) {
  const seen = new Set();
  for (const style of styles) {
    if (!style?.styleId) continue;
    if (seen.has(style.styleId)) return true;
    seen.add(style.styleId);
  }
  return false;
}

function styleExportRank(style) {
  // Default styles come first: Normal(sti=0) → DPF(sti=65) → Normal Table(sti=105)
  if (style?.sti === 0) return 0;
  // MS-DOC-SPEC/19 StdfBase.sti identifies built-in Heading 1..9. WPS emits
  // the contiguous built-in heading block before the default character/table
  // styles when those heading STDs are present.
  if (style?.sti >= 1 && style.sti <= 9) return style.sti;
  if (style?.sti === 65) return 20;
  if (style?.sti === 105) return 21;
  // All other styles follow by numeric styleId
  const numeric = Number(style?.styleId);
  if (Number.isInteger(numeric)) return numeric + 100;
  return 10000 + (style?.order ?? style?.index ?? 0);
}

function buildStyleReferenceXml(tag, style, rawCode, styles = []) {
  if (rawCode === null || rawCode === undefined || rawCode >= 0xfff0 || rawCode === 0x0fff) return "";
  if (rawCode === style.index) return "";
  let styleId = null;
  if (styles[rawCode]?.styleId) {
    styleId = styles[rawCode].styleId;
  } else if (rawCode === 0) {
    styleId = "1";
  } else if (rawCode === 10) {
    styleId = "9";
  } else if (rawCode === 161) {
    styleId = "9";
  } else {
    styleId = String(rawCode);
  }
  return `<w:${tag} w:val="${escapeXml(styleId)}"/>`;
}
// Strip repeated " Char[digits]" suffixes from linked character style names
function stripCharSuffix(n) {
  let prev = n;
  while (true) {
    const stripped = prev.replace(/ Char[0-9]*$/, "");
    if (stripped === prev) break;
    prev = stripped;
  }
  return prev;
}

function inferStyleLink(style, styles = []) {
  const name = style.name ?? "";
  if (style.linkCode != null && style.linkCode !== 0 && style.linkCode < 0xfff0) {
    const linked = styles[style.linkCode];
    if (linked?.styleId && (
      (style.type === "paragraph" && linked.type === "character")
      || (style.type === "character" && linked.type === "paragraph")
    )) {
      // MS-DOC-SPEC/19 StdfPost2000.istdLink is an STSH index of the linked
      // style; OOXML stores the target styleId in w:link.
      return `<w:link w:val="${escapeXml(linked.styleId)}"/>`;
    }
  }

  // Build a list of non-null styles in STSH index order for adjacency checks
  const ordered = styles.filter(s => s !== null).sort((a, b) => (a.index ?? 0) - (b.index ?? 0));

  // Check if two adjacent styles share the same istdLink (linkCode) — WPS uses this
  // as the primary link signal when names don't match (e.g. " Char Char" → "页脚").
  const shareLinkCode = (a, b) => {
    if (a.linkCode == null || b.linkCode == null) return false;
    if (a.linkCode === 0 || b.linkCode === 0) return false;
    if (a.linkCode >= 0xfff0 || b.linkCode >= 0xfff0) return false;
    return a.linkCode === b.linkCode;
  };

  if (style.type === "character") {
    const baseName = stripCharSuffix(name);
    const myPos = ordered.indexOf(style);
    // First pass: try name matching
    for (const delta of [-1, 1]) {
      const other = ordered[myPos + delta];
      if (!other || other.type !== "paragraph") continue;
      if (other.name === baseName || other.styleName === baseName) {
        return `<w:link w:val="${escapeXml(other.styleId)}"/>`;
      }
    }
    // Second pass: try shared istdLink (linkCode) as fallback for weird names
    for (const delta of [-1, 1]) {
      const other = ordered[myPos + delta];
      if (!other || other.type !== "paragraph") continue;
      if (shareLinkCode(style, other)) {
        return `<w:link w:val="${escapeXml(other.styleId)}"/>`;
      }
    }
  } else if (style.type === "paragraph") {
    const myPos = ordered.indexOf(style);
    // First pass: name matching
    for (const delta of [-1, 1]) {
      const other = ordered[myPos + delta];
      if (!other || other.type !== "character") continue;
      const otherName = other.name ?? "";
      const baseOther = stripCharSuffix(otherName);
      if (baseOther === name && baseOther !== otherName) {
        return `<w:link w:val="${escapeXml(other.styleId)}"/>`;
      }
    }
    // Second pass: shared linkCode fallback
    for (const delta of [-1, 1]) {
      const other = ordered[myPos + delta];
      if (!other || other.type !== "character") continue;
      if (shareLinkCode(style, other)) {
        return `<w:link w:val="${escapeXml(other.styleId)}"/>`;
      }
    }
  }
  return "";
}

function buildStyleParagraphPropertiesXml(style, docGridLinePitch = null) {
  const parts = [];
  const paragraphStyle = style;

  const numbering = buildParagraphNumberingXml(paragraphStyle, "");
  const hasFrameProperties = hasParagraphFrameProperties(paragraphStyle);
  if (hasFrameProperties) {
    appendParagraphControlXml(parts, paragraphStyle, {
      includeDefaults: false,
      phase: "beforeFrame",
    });
  } else {
    appendParagraphControlXml(parts, paragraphStyle, {
      includeDefaults: false,
      lineNumberCount: paragraphStyle?.lineNumberCount,
      numberingXml: numbering.xml,
      phase: "beforeTabs",
    });
  }
  if (paragraphStyle.paragraphBorders) {
    const borderParts = [];
    for (const [side, brc] of Object.entries(paragraphStyle.paragraphBorders)) {
      if (!brc) continue;
      borderParts.push(`<w:${side} w:val="${escapeXml(brc.val)}" w:color="${escapeXml(brc.color)}" w:sz="${escapeXml(brc.sz)}" w:space="${escapeXml(brc.space)}"/>`);
    }
    if (borderParts.length) {
      parts.push(`<w:pBdr>${borderParts.join("")}</w:pBdr>`);
    }
  }
  if (!numbering.xml) {
    appendParagraphTabsXml(parts, paragraphStyle, "", false);
  }
  appendParagraphControlXml(parts, paragraphStyle, {
    includeDefaults: false,
    phase: "afterTabs",
  });
  if (hasFrameProperties) {
    appendParagraphFramePropertiesXml(parts, paragraphStyle);
    appendParagraphControlXml(parts, paragraphStyle, {
      includeDefaults: false,
      lineNumberCount: paragraphStyle?.lineNumberCount,
      numberingXml: numbering.xml,
      phase: "afterFrame",
    });
  }
  appendParagraphSpacingXml(parts, paragraphStyle, docGridLinePitch != null ? { docGridLinePitch } : null, { styleContext: true });
  appendParagraphIndentXml(parts, paragraphStyle, "");
  if (paragraphStyle.paragraphShading) {
    parts.push(`<w:shd w:val="${paragraphStyle.paragraphShading.val}" w:color="${paragraphStyle.paragraphShading.color}" w:fill="${paragraphStyle.paragraphShading.fill}"/>`);
  }
  if (paragraphStyle.mirrorIndents) {
    parts.push(`<w:mirrorIndents/>`);
  }
  if (paragraphStyle.alignment) {
    parts.push(`<w:jc w:val="${paragraphStyle.alignment}"/>`);
  }
  if (paragraphStyle.textAlignment) {
    parts.push(`<w:textAlignment w:val="${paragraphStyle.textAlignment}"/>`);
  }
  if (paragraphStyle.outlineLevel != null) {
    parts.push(`<w:outlineLvl w:val="${paragraphStyle.outlineLevel}"/>`);
  }

  if (!parts.length) return "";
  return `<w:pPr>${parts.join("")}</w:pPr>`;
}

function buildStyleRunPropertiesXml(style, fontTable) {
  // For style definitions, suppress automatic derivation of szCs from sz
  // and automatic complex-script toggles (bCs). Only emit szCs when it is
  // explicitly set (emitExplicitComplexScriptSize=true handles that path).
  // This matches WPS Office expected output: style rPr omits redundant Cs props.
  // MS-DOC-SPEC/19 StkParaGRLPUPX stores paragraph-style run properties in
  // UpxChpx. Emit the parsed UPX properties as stored; built-in heading bold
  // appears when sprmCFBold is present in that UPX, not from a synthesized
  // formatter default.
  const runProperties = style.runProperties;
  const runPropertiesXml = runProperties
    ? buildRunPropertiesXmlFromProps(runProperties, fontTable, { includeDefaults: false, emitExplicitComplexScriptSize: true, emitUnderlineHighlight: false, suppressComplexScriptSize: true, suppressComplexScriptToggles: true })
    : "";
  return runPropertiesXml;
}

function buildLatentStylesXml(styles = [], wpsDocument = {}) {
  const latentLsd = wpsDocument.latentLsd ?? [];
  // MS-DOC-SPEC/19 Stshif.stiMaxWhenSaved is version-dependent; Appendix A
  // lists 156 for Word 2002/2003 and 267 for Word 2007+. WPS DOCX exports
  // the standard OOXML 260-count latentStyles table, so table-style latent
  // defaults beyond a Word 2003-era STSH are synthesized below.
  const stiMax = Math.max(wpsDocument.stiMaxWhenSaved ?? 0, 259);
  const parts = ['<w:latentStyles w:count="260" w:defQFormat="0" w:defUnhideWhenUsed="1" w:defSemiHidden="1" w:defUIPriority="99" w:defLockedState="0">'];
  for (let sti = 0; sti < STI_NAMES.length && sti <= stiMax; sti += 1) {
    const name = STI_NAMES[sti];
    const latent = latentLsd[sti] ?? synthesizeLatentStyleDefaults(sti, name);
    if (!latent || !name) continue;
    // Skip deprecated HTML/numbering styles (92-93, 107-110) which are not real
    // paragraph/character/table styles in the OOXML latentStyles table.
    // sti 156-157, 178-181 are newer paragraph/character styles beyond older
    // binary LSD ranges and never appear in WPS Office DOCX latentStyles output.
    if (sti === 92 || sti === 93 || (sti >= 107 && sti <= 110) || sti === 156 || sti === 157 || sti === 178 || sti === 180 || sti === 181) continue;
    // MS-DOC-SPEC 2.9.145 LSD: sti=179 (List Paragraph) is a real paragraph style.
    // Skip it only when ALL LSD fields match the w:latentStyles defaults
    // (defQFormat=0, defSemiHidden=1, defUnhideWhenUsed=1, defUIPriority=99),
    // in which case the lsdException would be redundant.
    if (sti === 179 && !latent.fQFormat && latent.fSemiHidden && latent.fUnhideWhenUsed && latent.iPriority === 99) continue;
    const attrs = [];
    // MS-DOC-SPEC/19 §StdfBase.grfstd: when a style has an actual STSH
    // definition its GRFSTD fields take precedence over the LSD latent
    // defaults. Override LSD with the actual style's flags.
    const actualStyle = (styles ?? []).find(s => s?.sti === sti);
    const styleFlags = actualStyle?.styleFlags;
    const semiHidden = styleFlags?.semiHidden != null ? styleFlags.semiHidden : latent.fSemiHidden;
    const unhideWhenUsed = styleFlags?.unhideWhenUsed != null ? styleFlags.unhideWhenUsed : latent.fUnhideWhenUsed;
    const qFormat = styleFlags?.qFormat != null ? styleFlags.qFormat : latent.fQFormat;
    const iPriority = actualStyle?.uiPriority ?? latent.iPriority;
    if (qFormat) attrs.push('w:qFormat="1"');
    if (!unhideWhenUsed) attrs.push('w:unhideWhenUsed="0"');
    attrs.push('w:uiPriority="' + iPriority + '"');
    if (!semiHidden) attrs.push('w:semiHidden="0"');
    parts.push('<w:lsdException ' + attrs.join(' ') + ' w:name="' + name + '"/>');
  }
  parts.push('</w:latentStyles>');
  return parts.join('');
}

function synthesizeLatentStyleDefaults(sti, name) {
  if (sti < 158 || sti > 259) return null;
  const baseName = name.replace(/ Accent [1-6]$/, "");
  const tableStylePriorities = {
    "Light Shading": 60,
    "Light List": 61,
    "Light Grid": 62,
    "Medium Shading 1": 63,
    "Medium Shading 2": 64,
    "Medium List 1": 65,
    "Medium List 2": 66,
    "Medium Grid 1": 67,
    "Medium Grid 2": 68,
    "Medium Grid 3": 69,
    "Dark List": 70,
    "Colorful Shading": 71,
    "Colorful List": 72,
    "Colorful Grid": 73,
  };
  const iPriority = tableStylePriorities[baseName];
  if (iPriority == null) return null;
  return {
    fQFormat: false,
    fUnhideWhenUsed: false,
    fSemiHidden: false,
    iPriority,
  };
}

function splitTabs(value) {
  const parts = [];
  let start = 0;
  for (let i = 0; i < value.length; i += 1) {
    if (value[i] === "\t") {
      if (i > start) {
        parts.push(value.slice(start, i));
      }
      parts.push("\t");
      start = i + 1;
    }
  }
  if (start < value.length) {
    parts.push(value.slice(start));
  }
  return parts;
}

function needsPreservedSpace(value) {
  return /^[\t\n\r ]|[\t\n\r ]$/.test(value);
}

function createCoreXml({ title, creator, created, modified }) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escapeXml(title)}</dc:title>
  <dc:creator>${escapeXml(creator)}</dc:creator>
  <cp:lastModifiedBy>${escapeXml(creator)}</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${escapeXml(created)}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${escapeXml(modified)}</dcterms:modified>
</cp:coreProperties>`;
}

function createZip(entries) {
  const fileRecords = [];
  let offset = 0;

  for (const entry of entries) {
    const name = Buffer.from(entry.name, "utf8");
    const data = Buffer.isBuffer(entry.data) ? entry.data : Buffer.from(entry.data, "utf8");
    const crc = crc32(data);
    const localHeader = Buffer.alloc(30 + name.length);

    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt16LE(0x0800, 6);
    localHeader.writeUInt16LE(0, 8);
    localHeader.writeUInt16LE(0, 10);
    localHeader.writeUInt16LE(0, 12);
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(data.length, 18);
    localHeader.writeUInt32LE(data.length, 22);
    localHeader.writeUInt16LE(name.length, 26);
    localHeader.writeUInt16LE(0, 28);
    name.copy(localHeader, 30);

    fileRecords.push({ name, data, crc, offset, localHeader });
    offset += localHeader.length + data.length;
  }

  const centralDirectory = [];
  for (const record of fileRecords) {
    const centralHeader = Buffer.alloc(46 + record.name.length);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(20, 4);
    centralHeader.writeUInt16LE(20, 6);
    centralHeader.writeUInt16LE(0x0800, 8);
    centralHeader.writeUInt16LE(0, 10);
    centralHeader.writeUInt16LE(0, 12);
    centralHeader.writeUInt16LE(0, 14);
    centralHeader.writeUInt32LE(record.crc, 16);
    centralHeader.writeUInt32LE(record.data.length, 20);
    centralHeader.writeUInt32LE(record.data.length, 24);
    centralHeader.writeUInt16LE(record.name.length, 28);
    centralHeader.writeUInt16LE(0, 30);
    centralHeader.writeUInt16LE(0, 32);
    centralHeader.writeUInt16LE(0, 34);
    centralHeader.writeUInt16LE(0, 36);
    centralHeader.writeUInt32LE(0, 38);
    centralHeader.writeUInt32LE(record.offset, 42);
    record.name.copy(centralHeader, 46);
    centralDirectory.push(centralHeader);
  }

  const centralDirectorySize = centralDirectory.reduce((sum, header) => sum + header.length, 0);
  const endOfCentralDirectory = Buffer.alloc(22);
  endOfCentralDirectory.writeUInt32LE(0x06054b50, 0);
  endOfCentralDirectory.writeUInt16LE(0, 4);
  endOfCentralDirectory.writeUInt16LE(0, 6);
  endOfCentralDirectory.writeUInt16LE(fileRecords.length, 8);
  endOfCentralDirectory.writeUInt16LE(fileRecords.length, 10);
  endOfCentralDirectory.writeUInt32LE(centralDirectorySize, 12);
  endOfCentralDirectory.writeUInt32LE(offset, 16);
  endOfCentralDirectory.writeUInt16LE(0, 20);

  return Buffer.concat([
    ...fileRecords.flatMap((record) => [record.localHeader, record.data]),
    ...centralDirectory,
    endOfCentralDirectory,
  ]);
}

function escapeXml(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

const CRC_TABLE = new Uint32Array(256);
for (let i = 0; i < CRC_TABLE.length; i += 1) {
  let value = i;
  for (let bit = 0; bit < 8; bit += 1) {
    value = (value & 1) ? (0xedb88320 ^ (value >>> 1)) : (value >>> 1);
  }
  CRC_TABLE[i] = value >>> 0;
}

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
