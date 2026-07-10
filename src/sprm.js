const SPRM_SIZES = {
  0x0800: 1, 0x0801: 1, 0x0802: 1, 0x0806: 1, 0x080A: 1,
  0x0855: 1, 0x0856: 1, 0x6A03: 4,
  0x6C02: 4, 0x6C03: 4, 0x6C04: 4, 0x6C05: 4,
  0xCE08: -1, 0xCE09: -1, 0xCE0A: -1, 0xCE0B: -1,
  0x2400: 4, 0x2401: 4, 0x2402: 2, 0x2403: 1, 0x2404: 1,
  0x2405: 1, 0x2406: 1, 0x2407: 1, 0x2408: 1, 0x2409: 1,
  0x240A: 4, 0x240B: 1, 0x240C: 4, 0x240D: 4, 0x240E: 4,
  0x2410: 1, 0x2411: 1, 0x2413: 1, 0x2414: 1, 0x2415: 1,
  0x240C: 1,
  0x2416: 1, 0x2417: 4, 0x2418: 4, 0x2419: 1, 0x241A: 1,
  0x241B: 1, 0x241C: 1, 0x241D: 1, 0x241E: 1, 0x241F: 1,
  0x2420: 1, 0x2421: 1, 0x2422: 1, 0x2423: 1, 0x2424: 1,
  0x2431: 1, 0x2433: 1, 0x2434: 1, 0x2435: 1, 0x2436: 1,
  0x2437: 1, 0x2438: 1, 0x2439: 2, 0x243A: 2, 0x243B: 2,
  0x243C: 1, 0x243D: 1, 0x243E: 1, 0x243F: 1,
  0x2441: 1, 0x2447: 1, 0x4439: 2,
  0x2425: 1, 0x2426: 1, 0x2427: 1, 0x2428: 1, 0x2429: 1,
  0x242A: 1, 0x242B: 1, 0x242C: 1, 0x242D: 1, 0x242E: 1,
  0x242F: 1, 0x2430: 1, 0x2432: 1,
  0x2440: 1, 0x2441: 1, 0x2443: 1, 0x2444: 1, 0x2445: 4,
  0x2446: 4, 0x2447: 1, 0x2448: 1, 0x2449: 4, 0x244A: 4,
  0x244B: 1, 0x244C: 1, 0x244D: 4, 0x244E: 4, 0x244F: 4,
  0x2450: 4, 0x2451: 4, 0x2452: 4, 0x2453: 4, 0x2454: 4,
  0x245A: 1, 0x245B: 1, 0x245C: 1, 0x246D: 1,
  0x260A: 1, 0x460B: 2, 0x2664: 1,
  0x261B: 1, 0x2423: 1, 0x2430: 1,
  0x8418: 2, 0x8419: 2, 0x841A: 2, 0x442B: 2,
  0x2640: 1,
  0x2461: 1, 0x2462: 1, 0x2463: 1, 0x2464: 1, 0x2465: 4,
  0x6412: 4,
  0x6424: 4, 0x6425: 4, 0x6426: 4, 0x6427: 4, 0x6428: 4, 0x6629: 4,
  0x6A0C: 4, 0x6A0D: 4, 0x6A0E: 4, 0x6A0F: 4,
  0x4455: 2, 0x4456: 2, 0x4457: 2, 0x4458: 2, 0x4459: 2,
  0x840E: 2, 0x840F: 2, 0x8411: 2, 0x8458: 2, 0x845D: 2,
  0x845E: 2, 0x8460: 2,
  0xA413: 2, 0xA414: 2,
  0x3465: 1, 0x360D: 1, 0x940E: 2, 0x940F: 2, 0x9410: 2,
  0x9411: 2, 0x941E: 2, 0x941F: 2,
  0x2A02: 1, 0x2A03: 1, 0x2A04: 1, 0x2A05: 1, 0x2A06: 1,
  0x2A07: 1, 0x2A08: 1, 0x2A09: 1, 0x2A0A: 1, 0x2A0B: 1,
  0x2A0C: 1, 0x2A0D: 1, 0x2A0E: 1, 0x2A0F: 1, 0x2A10: 1,
  0x2A11: 1, 0x2A12: 1, 0x2A13: 1, 0x2A14: 1, 0x2A15: 1,
  0x2A16: 1, 0x2A17: 1, 0x2A18: 1, 0x2A19: 1, 0x2A1A: 1,
  0x2A1B: 1, 0x2A1C: 1, 0x2A1D: 1, 0x2A1E: 1, 0x2A1F: 1,
  0x2A20: 1, 0x2A21: 1, 0x2A22: 1, 0x2A23: 1, 0x2A24: 1,
  0x2A25: 1, 0x2A26: 1, 0x2A27: 1, 0x2A28: 1, 0x2A29: 1,
  0x2A2A: 1, 0x2A2B: 1, 0x2A2C: 1, 0x2A2D: 1, 0x2A2E: 1,
  0x2A2F: 1, 0x2A30: 1, 0x2A31: 1, 0x2A32: 1, 0x2A33: 1,
  0x2A34: 1, 0x2A35: 1, 0x2A36: 1, 0x2A37: 1, 0x2A38: 1,
  0x2A39: 1, 0x2A3A: 1, 0x2A3B: 1, 0x2A3C: 1, 0x2A3D: 1,
  0x2A3E: 1, 0x2A3F: 1, 0x2A40: 1, 0x2A41: 1, 0x2A42: 1,
  0x2A43: 1, 0x2A44: 1, 0x2A45: 1, 0x2A46: 1, 0x2A47: 1,
  0x2A48: 1, 0x2A49: 1, 0x2A4A: 1, 0x2A4B: 1, 0x2A4C: 1,
  0x2A4D: 1, 0x2A4E: 1, 0x2A4F: 1, 0x2A50: 1, 0x2A51: 1,
  0x2A52: 1, 0x2A53: 1, 0x2A54: 1, 0x2A55: 1, 0x2A56: 1,
  0x0868: 1,
  0x2879: 1,
  0x286F: 1,
  0x4A30: 2, 0x4A31: 2, 0x4A32: 2, 0x4A33: 2, 0x4A34: 2,
  0x4A35: 2, 0x4A36: 2, 0x4A37: 2, 0x4A38: 2, 0x4A39: 2,
  0x4A3A: 2, 0x4A3B: 2, 0x4A3C: 2, 0x4A3D: 2, 0x4A3E: 2,
  0x4A3F: 2, 0x4A40: 2, 0x4A41: 2, 0x4A42: 2, 0x4A43: 2,
  0x4A44: 2, 0x4A45: 2, 0x4A46: 2, 0x4A47: 2, 0x4A48: 2,
  0x4A49: 2, 0x4A4A: 2, 0x4A4B: 2, 0x4A4C: 2, 0x4A4D: 2,
  0x4A4E: 2, 0x4A4F: 2, 0x4A50: 2, 0x4A51: 2, 0x4A52: 2,
  0x4A53: 2, 0x4A54: 2, 0x4A55: 2, 0x4A56: 2, 0x4A57: 2,
  0x4A58: 2, 0x4A59: 2, 0x4A5A: 2, 0x4A5B: 2, 0x4A5C: 2,
  0x4A5D: 2, 0x4A5E: 2, 0x4A5F: 2,
  0x6A09: 4,
  0x6A40: 4, 0x6A41: 4, 0x6A42: 4, 0x6A43: 4, 0x6A44: 4,
  0x6A45: 4, 0x6A46: 4, 0x6A47: 4, 0x6A48: 4, 0x6A49: 4,
  0x4852: 2, 0x4853: 2, 0x4854: 2, 0x4855: 2, 0x4856: 2,
  0x484B: 2, 0x4857: 2, 0x4858: 2, 0x4859: 2, 0x485A: 2, 0x485B: 2,
  0x8840: 2, 0x8841: 2, 0x8842: 2, 0x8843: 2, 0x8844: 2,
  0x8845: 2, 0x8846: 2, 0x8847: 2, 0x8848: 2, 0x8849: 2,
  0xC60D: -1, 0xC66F: -1, 0xCA76: -1, 0xCA78: -1,
};

const SPRM_CATEGORIES = {
  0x2403: "jc",
  0x6412: "lineSpacing",
  0x840E: "rightIndent",
  0x840F: "leftIndent",
  0x4610: "nest80",
  0x8411: "firstLineIndent",
  0x4455: "rightIndentChars",
  0x4456: "leftIndentChars",
  0x4457: "firstLineIndentChars",
  0x4458: "spacingBeforeLines",
  0x4459: "spacingAfterLines",
  0x8458: "firstLineIndent",
  0x845D: "rightIndent",
  0x845E: "leftIndent",
  0x465F: "nest",
  0x8460: "firstLineIndent",
  0x2400: "spacingBefore",
  0x2401: "spacingAfter",
  0x4A43: "fontSize",
  0x4845: "textPosition",
  0x4A30: "characterStyleIndex",
  0x4A4F: "fontAscii",
  0x4A50: "fontEastAsia",
  0x4A51: "fontHAnsi",
  0x4A61: "fontSizeCs",
  0x4A5E: "fontCs",
  0x6A03: "pictureLocation",
  0x0806: "binaryData",
  0x080A: "ole2",
  0x0855: "specialCharacter",
  0x0856: "embeddedObject",
  0x6C02: "pictureBorderTop", 0x6C03: "pictureBorderLeft", 0x6C04: "pictureBorderBottom", 0x6C05: "pictureBorderRight",
  0xCE08: "pictureBorderTop", 0xCE09: "pictureBorderLeft", 0xCE0A: "pictureBorderBottom", 0xCE0B: "pictureBorderRight",
  0x6A09: "symbol",
  0x484B: "kern",
  0x485F: "langIdBidi",
  0x8840: "charSpacing",
  0x4852: "charWidth",
  0x2A02: "bold",
  0x2A03: "italic",
  0x2A0E: "underline",
  0x2A3E: "underline",
  0x2879: "lineBreakClear",
  0xCA76: "fitText",
  0xCA78: "eastAsianLayout",
  0x0868: "charSnapToGrid",
  0x286F: "fontHint",
  0x486D: "langId",
  0x486E: "langIdEastAsia",
  0x4873: "langId",
  0x4874: "langIdEastAsia",
  0xC60D: "tabs",
  0x3465: "tableNoAllowOverlap",
  0x360D: "tablePositionCode",
  0x940E: "tablePositionX",
  0x940F: "tablePositionY",
  0x9410: "tablePositionLeftMargin",
  0x9411: "tablePositionTopMargin",
  0x941E: "tablePositionRightMargin",
  0x941F: "tablePositionBottomMargin",
  0x6A0C: "fontIdWps",
  // LibreOffice WW8 maps these paragraph flags as:
  // PFKeep -> keepLines, PFKeepFollow -> keepNext, PFPageBreakBefore -> pageBreakBefore.
  0x2405: "keepLines",
  0x2406: "keepNext",
  0x2407: "pageBreakBefore",
  0x2441: "bidi",
  0x2447: "snapToGrid",
  0x4439: "textAlignment",
  0x2431: "widowControl",
  0x2433: "kinsoku",
  0x2434: "wordWrap",
  0x2435: "overflowPunct",
  0x2436: "topLinePunct",
  0x2437: "autoSpaceDE",
  0x2438: "autoSpaceDN",
  0x2448: "adjustRightInd",
  0x245B: "spacingBeforeAuto",
  0x245C: "spacingAfterAuto",
  0x260A: "listLevel",
  0x460B: "listId",
  0x2664: "paragraphWall",
  0xC66F: "paragraphPropRMark",
};

const ALIGNMENT_MAP = { 0: "left", 1: "right", 2: "center", 3: "both", 4: "distribute", 5: "numTab" };
const WORD_ALIGNMENT_MAP = { 0: "left", 1: "center", 2: "right", 3: "both", 4: "distribute", 5: "numTab" };

const SPRM_OPERAND_SIZE_BY_SPRA = [1, 1, 2, 4, 2, 2, -1, 3];

// Known SPRM codes that can have a 0x00 low byte (to prevent the padding-
// byte skipper from consuming their first byte). Every SPRM in SPRM_SIZES
// is implicitly known; this extra set covers SPRMs handled only by switch
// cases in applySprm.
const KNOWN_ZEROBYTE_SPRMS = new Set([
  0x0800, 0x0801, 0x0802,  // character revision marks
  0x4804, 0x4807, 0x4863, 0x4867,  // revision author/reason
  0x6805, 0x6864,  // revision date
  0xCA57, 0xCA89,  // PropRMark operands
  0xC66F,  // paragraph PropRMark
  0x3000, 0x3001, 0x3005, 0x3006, 0x3009, 0x300A, 0x300E,  // section (0x30xx)
  0x5400, 0x548A,  // table justification
  0x2417, 0x2416,  // table paragraph flags
]);

function isKnownSprm(sprm) {
  return sprm in SPRM_SIZES || KNOWN_ZEROBYTE_SPRMS.has(sprm);
}
// MS-DOC-SPEC/19 Ico: value MUST be less than 0x11. Use RGB hex for
// OOXML w:color; 0 is automatic color.
const WW8_TEXT_COLOR_INDEXES = [
  "auto",
  "000000",
  "0000FF",
  "00FFFF",
  "00FF00",
  "FF00FF",
  "FF0000",
  "FFFF00",
  "FFFFFF",
  "000080",
  "008080",
  "008000",
  "800080",
  "800000",
  "808000",
  "808080",
  "C0C0C0",
];

const WW8_SHADING_STRENGTHS = [
  0,
  1000,
  50,
  100,
  200,
  250,
  300,
  400,
  500,
  600,
  700,
  750,
  800,
  900,
  333,
  333,
  333,
  333,
  333,
  333,
  333,
  333,
  333,
  333,
  333,
  333,
  500,
  500,
  500,
  500,
  500,
  500,
  500,
  500,
  500,
  25,
  75,
  125,
  150,
  175,
  225,
  275,
  325,
  350,
  375,
  425,
  450,
  475,
  525,
  550,
  575,
  625,
  650,
  675,
  725,
  775,
  825,
  850,
  875,
  925,
  950,
  975,
];

function isLowByteOfKnownSprm(grpprl, off) {
  // Check if the two-byte little-endian value at (off, off+1) is a SPRM
  // code that this parser knows about and that has 0x00 as its low byte.
  // This prevents the padding-byte skipper from consuming the first byte
  // of valid SPRMs like 0x0800 (sprmCFRMarkDel).
  if (off + 2 > grpprl.length) return false;
  const candidate = grpprl.readUInt16LE(off);
  return (candidate & 0xFF) === 0 && isKnownSprm(candidate);
}

export function parseSprms(grpprl, skipIstd = false, cupx = 0) {
  const props = {};
  if (skipIstd && grpprl.length >= 2) {
    props.istd = grpprl.readUInt16LE(0);
  }
  for (const { sprm, val, size } of scanSprmRecords(grpprl, skipIstd, cupx)) {
    applySprm(props, sprm, val, size);
  }
  return props;
}

export function scanSprmRecords(grpprl, skipIstd = false, cupx = 0) {
  const records = [];
  let off = skipIstd ? 2 : 0;
  // Skip Upx section (user-defined properties) in style GRPPRLs.
  // cupx values per MS-DOC StdfBase: 1=char, 2=para, 3=table
  off += cupx;
  if (skipIstd && off < grpprl.length && grpprl[off] === 0) {
    off += 1;
  }
  while (off + 2 <= grpprl.length) {
    // Style GRPPRLs can contain zero padding that is ambiguous with
    // low-byte-zero character SPRMs; only recover those SPRMs in normal GRPPRLs.
    if (grpprl[off] === 0 && (skipIstd || !isLowByteOfKnownSprm(grpprl, off))) {
      off += 1;
      continue;
    }
    const recordOffset = off;
    const sprm = grpprl.readUInt16LE(off);
    off += 2;
    let size = SPRM_SIZES[sprm];
    if (size === undefined) {
      const spra = (sprm >> 13) & 0x7;
      size = SPRM_OPERAND_SIZE_BY_SPRA[spra];
    }
    if (size === -1) {
      if (off >= grpprl.length) break;
      size = grpprl[off] + 1;
    }
    if (!size || size < 0) {
      throw new Error(`Out-of-spec SPRM operand size for 0x${sprm.toString(16)}`);
    }
    if (off + size > grpprl.length) {
      if (sprm === 0xC60D) {
        size = grpprl.length - off;
      } else if (SPRM_CATEGORIES[sprm]) {
        throw new Error(`Truncated SPRM operand for 0x${sprm.toString(16)}`);
      } else {
        break;
      }
    }
    const val = grpprl.subarray(off, off + size);
    records.push({ sprm, offset: recordOffset, operandOffset: off, operandEnd: off + size, val, size });
    off += size;
  }
  return records;
}

function applySprm(props, sprm, val, size) {
  switch (sprm) {
    case 0x6412: {
      // Byte 2 of the 4-byte operand encodes the rule: 0=atLeast, 1=auto, 2+=exact.
      // When byte 2 is absent (old 2-byte format), use signed twips convention.
      const ruleByte = val.length >= 3 ? val[2] : null;
      let rule, twips;
      if (ruleByte === 1) {
        twips = val.readUInt16LE(0);
        rule = "auto";
      } else if (ruleByte != null && ruleByte >= 2) {
        twips = val.readUInt16LE(0);
        rule = "exact";
      } else {
        const raw = val.readUInt16LE(0);
        const signed = raw > 0x7fff ? raw - 0x10000 : raw;
        twips = Math.abs(signed);
        rule = signed < 0 ? "exact" : "atLeast";
      }
      props.lineSpacing = { twips, rule };
      break;
    }
    case 0x2403:
      props.alignment = WORD_ALIGNMENT_MAP[val[0]] || ALIGNMENT_MAP[val[0]] || "left";
      break;
    case 0xA413:
      props.spacingBefore = val.readInt16LE(0);
      break;
    case 0xA414:
      props.spacingAfter = val.readInt16LE(0);
      break;
    case 0x840E:
    case 0x845D:
      props.rightIndent = val.readInt16LE(0);
      break;
    case 0x840F:
      setParagraphIndentPart(props, "leftIndent80", val.readInt16LE(0));
      break;
    case 0x845E:
      setParagraphIndentPart(props, "leftIndent", val.readInt16LE(0));
      break;
    case 0x4610:
      setParagraphIndentPart(props, "nest80", val.readInt16LE(0));
      break;
    case 0x465F:
      setParagraphIndentPart(props, "nest", val.readInt16LE(0));
      break;
    case 0x4455:
      props.rightIndentChars = val.readInt16LE(0);
      break;
    case 0x4456:
      props.leftIndentChars = val.readInt16LE(0);
      break;
    case 0x4457:
      props.firstLineIndentChars = val.readInt16LE(0);
      // MS-DOC-SPEC/16 sprmPDxcLeft1 stores the first-line indent in
      // hundredths of character units. Preserve the SPRM id as parsed
      // evidence so DOCX emission can avoid inferred source attribution.
      props.firstLineIndentCharsSprm = sprm;
      break;
    case 0x4458:
      // MS-DOC-SPEC/16 sprmPDylBefore is already in 1/100 line units.
      props.spacingBeforeLines = val.readInt16LE(0);
      break;
    case 0x4459:
      // MS-DOC-SPEC/16 sprmPDylAfter is already in 1/100 line units.
      props.spacingAfterLines = val.readInt16LE(0);
      break;
    case 0x8411:
    case 0x8458:
    case 0x8460:
      props.firstLineIndent = val.readInt16LE(0);
      // MS-DOC-SPEC/16 has multiple twip first-line indent SPRMs
      // (sprmPDxaLeft180, sprmPDxaLeft1). Preserve the winning SPRM id
      // from the parsed grpprl order instead of guessing the source later.
      props.firstLineIndentSprm = sprm;
      break;
    case 0x2400:
      props.spacingBefore = val.readInt32LE(0);
      break;
    case 0x2401:
      props.spacingAfter = val.readInt32LE(0);
      break;
    case 0x2664:
      props.paragraphWall = val[0] !== 0;
      break;
    case 0xC66F:
      // MS-DOC-SPEC/16 sprmPPropRMark stores a PropRMarkOperand for a
      // paragraph property revision mark.
      props.paragraphPropRMark = parsePropRMarkOperand(val, "paragraph property revision mark");
      break;
    case 0x4A43:
      props.fontSize = val.readUInt16LE(0);
      break;
    case 0x4845:
      // MS-DOC-SPEC/16 sprmCHpsPos: signed half-point vertical text position.
      props.textPosition = val.readInt16LE(0);
      break;
    case 0x2A48:
      props.verticalAlign = verticalAlignFromIss(val[0]);
      break;
    case 0x2A34:
      props.emphasisMark = emphasisMarkFromKcd(val[0]);
      break;
    case 0x2859:
      props.textEffect = textEffectFromSfxText(val[0]);
      break;
    case 0x4A61:
      props.fontSizeCs = val.readUInt16LE(0);
      break;
    case 0x4A30:
      // MS-DOC-SPEC/16 sprmCIstd: STSH index of the character style to apply.
      props.characterStyleIndex = val.readUInt16LE(0);
      break;
    case 0x6A0C:
      props.fontId = val.readUInt16LE(0);
      break;
    case 0x4A4F:
      props.fontAscii = val.readUInt16LE(0);
      break;
    case 0x4A50:
      props.fontEastAsia = val.readUInt16LE(0);
      break;
    case 0x4A51:
      props.fontHAnsi = val.readUInt16LE(0);
      break;
    case 0x4A5E:
      props.fontCs = val.readUInt16LE(0);
      break;
    case 0x6A09:
      props.symbolFontId = val.readUInt16LE(0);
      props.symbolChar = val.readUInt16LE(2);
      break;
    case 0x8840:
      props.charSpacing = val.readInt16LE(0);
      break;
    case 0x6A03:
      props.pictureLocation = val.readInt32LE(0);
      break;
    case 0x0806:
      props.binaryData = val[0] !== 0;
      break;
    case 0x080A:
      props.ole2 = val[0] !== 0;
      break;
    case 0x0855:
      props.specialCharacter = val[0] !== 0;
      break;
    case 0x0856:
      props.embeddedObject = val[0] !== 0;
      break;
    case 0x6C02:
    case 0x6C03:
    case 0x6C04:
    case 0x6C05: {
      const side = ["top", "left", "bottom", "right"][sprm - 0x6C02];
      props.pictureBorders ??= {};
      props.pictureBorders[side] = parseBrc80(val, { preserveNone: true });
      break;
    }
    case 0xCE08:
    case 0xCE09:
    case 0xCE0A:
    case 0xCE0B: {
      const side = ["top", "left", "bottom", "right"][sprm - 0xCE08];
      props.pictureBorders ??= {};
      props.pictureBorders[side] = parseBrcOperand(val, `inline picture ${side} border`, { preserveNone: true });
      break;
    }
    case 0x4852:
      props.charWidth = val.readUInt16LE(0);
      break;
    case 0x484B:
      props.kern = val.readUInt16LE(0);
      break;
    case 0x485F:
      props.langIdBidi = val.readUInt16LE(0);
      break;
    case 0x0835:
    case 0x2A02:
      props.bold = val[0] !== 0;
      break;
    case 0x085C:
      props.boldCs = val[0] !== 0;
      break;
    case 0x0836:
    case 0x2A03:
      props.italic = val[0] !== 0;
      break;
    case 0x085D:
      props.italicCs = val[0] !== 0;
      break;
    case 0x0837:
      props.strike = val[0] !== 0;
      break;
    case 0x0800:
      // MS-DOC-SPEC/16 sprmCFRMarkDel: deleted revision mark text
      props.revisionMarkDel = val[0] !== 0;
      break;
    case 0x0801:
      // MS-DOC-SPEC/16 sprmCFRMarkIns: inserted revision mark text
      props.revisionMarkIns = val[0] !== 0;
      break;
    case 0x0802:
      // MS-DOC-SPEC/16 sprmCFFldVanish: hide field text
      props.fldVanish = val[0] !== 0;
      break;
    case 0x0838:
      props.outline = val[0] !== 0;
      break;
    case 0x4804:
      // MS-DOC-SPEC/16 sprmCIbstRMark: author index into SttbfRMark for insertion
      props.revisionAuthorIndex = val.readUInt16LE(0);
      break;
    case 0x6805:
      // MS-DOC-SPEC/16 sprmCDttmRMark: date/time of insertion (DTTM)
      props.revisionDate = val.readUInt32LE(0);
      break;
    case 0x4807:
      // MS-DOC-SPEC/16 sprmCIdslRMark: reason for the revision mark
      props.revisionReason = val.readUInt16LE(0);
      break;
    case 0x4863:
      // MS-DOC-SPEC/16 sprmCIbstRMarkDel: author index for deletion
      props.revisionDelAuthorIndex = val.readUInt16LE(0);
      break;
    case 0x6864:
      // MS-DOC-SPEC/16 sprmCDttmRMarkDel: date/time of deletion (DTTM)
      props.revisionDelDate = val.readUInt32LE(0);
      break;
    case 0x4867:
      // MS-DOC-SPEC/16 sprmCIdslRMarkDel: reason for deletion
      props.revisionDelReason = val.readUInt16LE(0);
      break;
    case 0x0839:
      props.shadow = val[0] !== 0;
      break;
    case 0x2A53:
      props.dstrike = val[0] !== 0;
      break;
    case 0x083A:
      props.smallCaps = val[0] !== 0;
      break;
    case 0x083B:
      props.caps = val[0] !== 0;
      break;
    case 0x083C:
      props.vanish = val[0] !== 0;
      break;
    case 0x0811:
      props.webHidden = val[0] !== 0;
      break;
    case 0x0818:
      props.specVanish = val[0] !== 0;
      break;
    case 0x085A:
      props.rtl = val[0] !== 0;
      break;
    case 0x0882:
      props.complexScript = val[0] !== 0;
      break;
    case 0x0854:
      props.imprint = val[0] !== 0;
      break;
    case 0x0858:
      props.emboss = val[0] !== 0;
      break;
    case 0x0875:
      props.noProof = val[0] !== 0;
      break;
    case 0x2879:
      props.lineBreakClear = lineBreakClearFromLbcOperand(val[0]);
      break;
    case 0xCA76:
      props.fitText = parseCFitTextOperand(val);
      break;
    case 0xCA78:
      props.eastAsianLayout = parseFarEastLayoutOperand(val);
      break;
    case 0x2A0E:
    case 0x2A3E:
      props.underlineStyle = underlineStyleFromCode(val[0]);
      props.underline = props.underlineStyle != null;
      break;
    case 0x2A42:
    case 0x4A60:
      props.textColor = ww8TextColorHex(val[0]);
      break;
    case 0x6870:
      props.textColor = colorRefToHex(val);
      break;
    case 0x6877:
      props.underlineColor = colorRefToHex(val);
      break;
    case 0x2A0C:
      props.highlight = ww8HighlightName(val[0]);
      break;
    case 0x0868:
      // MS-DOC-SPEC/16 sprmCFUsePgsuSettings: character-level document
      // grid usage, mapped to OOXML run snapToGrid.
      props.charSnapToGrid = val[0] !== 0;
      break;
    case 0x6865:
      props.border = parseBrc80(val, { ignoreSpace: true });
      break;
    case 0xCA72:
      props.border = parseBrcOperand(val, "character border", { ignoreSpace: true });
      break;
    case 0x4866:
      props.background = parseWw8Shade(val);
      break;
    case 0xCA71:
      props.background = parseShdOperand(val, "character shading");
      break;
    case 0x442D:
      props.paragraphShading = parseWw8Shade(val);
      break;
    case 0xC64D:
      props.paragraphShading = parseShdOperand(val, "paragraph shading");
      break;
    case 0x286F:
      props.fontHint = val[0] === 1 ? "eastAsia" : "default";
      break;
    case 0x486D:
    case 0x4873:
      props.langId = val.readUInt16LE(0);
      break;
    case 0x486E:
    case 0x4874:
      props.langIdEastAsia = val.readUInt16LE(0);
      break;
    case 0xCA57:
    case 0xCA89:
      // MS-DOC-SPEC/16 sprmCPropRMark90 / sprmCPropRMark: character property
      // revision mark. Operand is PropRMarkOperand: cb(1) + PropRMark(7).
      props.charPropRMark = parsePropRMarkOperand(val, "character property revision mark");
      break;
    case 0xC60D:
      props.tabs = parseTabsOperand(val);
      break;
    case 0x3465:
      props.tableNoAllowOverlap = val[0] !== 0;
      break;
    case 0x360D:
      props.tablePosition ??= {};
      props.tablePosition.nTPc = val[0];
      break;
    case 0x940E:
      props.tablePosition ??= {};
      props.tablePosition.nTDxaAbs = val.readInt16LE(0);
      break;
    case 0x940F:
      props.tablePosition ??= {};
      props.tablePosition.nTDyaAbs = val.readInt16LE(0);
      break;
    case 0x9410:
      props.tablePosition ??= {};
      props.tablePosition.nLeftMargin = val.readUInt16LE(0);
      break;
    case 0x9411:
      props.tablePosition ??= {};
      props.tablePosition.nUpperMargin = val.readUInt16LE(0);
      break;
    case 0x941E:
      props.tablePosition ??= {};
      props.tablePosition.nRightMargin = val.readUInt16LE(0);
      break;
    case 0x941F:
      props.tablePosition ??= {};
      props.tablePosition.nLowerMargin = val.readUInt16LE(0);
      break;
    case 0x2416:
      props.inTable = val[0] !== 0;
      break;
    case 0x2417:
      // MS-DOC-SPEC/16 sprmPFTtp: Table Terminating Paragraph mark
      props.isTtp = val[0] !== 0;
      break;
    case 0x6649:
      // MS-DOC-SPEC/16 sprmPItap: table depth (integer)
      props.tableDepth = val.readUInt32LE(0);
      break;
    case 0x664A: {
      // MS-DOC-SPEC/16 sprmPDtap: signed adjustment to the existing table depth.
      const tableDepth = (props.tableDepth ?? 0) + val.readInt32LE(0);
      if (tableDepth < 0) {
        throw new Error(`Out-of-spec paragraph table depth ${tableDepth}`);
      }
      props.tableDepth = tableDepth;
      break;
    }
    case 0x244B:
      // MS-DOC-SPEC/16 sprmPFInnerTableCell
      props.innerTableCell = val[0] !== 0;
      break;
    case 0x244C:
      // MS-DOC-SPEC/16 sprmPFInnerTtp
      props.innerTtp = val[0] !== 0;
      break;
    case 0x245A:
      // MS-DOC-SPEC/16 sprmPFOpenTch: table cell mark display state after a nested table.
      props.openTch = val[0] !== 0;
      break;
    case 0x240C:
      // LibreOffice WW8 maps pap.fNoLnn to RES_LINENUMBER and writes
      // <w:suppressLineNumbers w:val="0"/> when the paragraph count flag is set.
      props.lineNumberCount = val[0] === 0;
      break;
    case 0x2405:
      props.keepLines = val[0] !== 0;
      break;
    case 0x2406:
      props.keepNext = val[0] !== 0;
      break;
    case 0x2407:
      props.pageBreakBefore = val[0] !== 0;
      break;
    case 0x242A:
      props.suppressAutoHyphens = val[0] !== 0;
      break;
    case 0x2431:
      props.widowControl = val[0] !== 0;
      break;
    case 0x2433:
      props.kinsoku = val[0] !== 0;
      break;
    case 0x2434:
      props.wordWrap = val[0] !== 0;
      break;
    case 0x2435:
      props.overflowPunct = val[0] !== 0;
      break;
    case 0x2436:
      props.topLinePunct = val[0] !== 0;
      break;
    case 0x2437:
      props.autoSpaceDE = val[0] !== 0;
      break;
    case 0x2438:
      props.autoSpaceDN = val[0] !== 0;
      break;
    case 0x245B:
      props.spacingBeforeAuto = val[0] !== 0;
      break;
    case 0x245C:
      props.spacingAfterAuto = val[0] !== 0;
      break;
    case 0x246D:
      // sprmPFContextualSpacing (MS-DOC-SPEC/16): Bool8 — ignore space between same-style paragraphs
      props.contextualSpacing = val[0] !== 0;
      break;
    case 0x2470:
      props.mirrorIndents = val[0] !== 0;
      break;
    case 0x2462:
      // sprmPFNoAllowOverlap: prevent this paragraph frame from overlapping other frames.
      props.suppressOverlap = val[0] !== 0;
      break;
    case 0x260A:
      props.listLevel = val[0];
      break;
    case 0x2640:
      // sprmPOutLvl: WW8 values 0-9 → OOXML outlineLvl 0-9 (0=Level1, …, 8=Level9, 9=BodyText).
      // OOXML §17.3.1.20 outlineLvl explicitly allows 9; WPS exports it.
      props.outlineLevel = val[0] <= 9 ? val[0] : null;
      break;
    // Frame properties per MS-DOC-SPEC §2.6.2
    case 0x2423: // sprmPWr: text wrap around frame
      props.frameWrap = FRAME_WRAP_NAMES[val[0]] ?? null;
      break;
    case 0x2430: // sprmPFLocked: anchor lock
      props.frameLocked = val[0] !== 0;
      break;
    case 0x8418: { // sprmPDxaAbs: horizontal position (XAS_plusOne)
      const x = val.readInt16LE(0);
      if (x === 0) props.frameXAlign = "left";
      else if (x === -4) props.frameXAlign = "center";
      else if (x === -8) props.frameXAlign = "right";
      else if (x === -12) props.frameXAlign = "inside";
      else if (x === -16) props.frameXAlign = "outside";
      else props.frameX = x;
      break;
    }
    case 0x8419: { // sprmPDyaAbs: vertical position (YAS_plusOne)
      const y = val.readInt16LE(0);
      if (y === 0) props.frameYAlign = "inline";
      else if (y === -4) props.frameYAlign = "top";
      else if (y === -8) props.frameYAlign = "center";
      else if (y === -12) props.frameYAlign = "bottom";
      else if (y === -16) props.frameYAlign = "inside";
      else if (y === -20) props.frameYAlign = "outside";
      else props.frameY = y;
      break;
    }
    case 0x841A: // sprmPDxaWidth: frame width in twips
      props.frameWidth = val.readUInt16LE(0);
      break;
    case 0x442B: { // sprmPWHeightAbs: frame height; bit 15 = fMinHeight (atLeast vs exact)
      const raw = val.readUInt16LE(0);
      props.frameHeight = raw & 0x7FFF;
      props.frameHRule = (raw & 0x8000) ? "atLeast" : "exact";
      break;
    }
    case 0x442C: { // sprmPDcs: drop cap properties for the paragraph frame
      Object.assign(props, parseDcs(val));
      break;
    }
    case 0x842E: // sprmPDyaFromText: minimum vertical frame-to-text distance in twips
      props.frameVSpace = val.readUInt16LE(0);
      break;
    case 0x842F: // sprmPDxaFromText: minimum horizontal frame-to-text distance in twips
      props.frameHSpace = val.readUInt16LE(0);
      break;
    case 0x261B: { // sprmPPc: frame anchor (PositionCodeOperand)
      const pc = val[0];
      const pcVert = (pc >> 4) & 3;
      const pcHorz = (pc >> 6) & 3;
      props.frameVAnchor = FRAME_V_ANCHOR[pcVert] ?? "page";
      props.frameHAnchor = FRAME_H_ANCHOR[pcHorz] ?? "page";
      break;
    }
    case 0x460B:
      props.listId = val.readInt16LE(0);
      break;
    case 0x2448:
      props.adjustRightInd = val[0] !== 0;
      break;
    case 0x2441:
      props.bidi = val[0] !== 0;
      break;
    case 0x2447:
      props.snapToGrid = val[0] !== 0;
      break;
    case 0x4439: {
      const alignMap = {
        0: "top",
        1: "center",
        2: "baseline",
        3: "bottom",
        4: "auto",
      };
      props.textAlignment = alignMap[val.readUInt16LE(0)] ?? "auto";
      break;
    }
    // Paragraph borders (Brc80 format: 4 bytes — dptLineWidth, brcType, ico, dptSpace|flags)
    case 0x6424: // sprmPBrcTop80
      props.paragraphBorders ??= {};
      props.paragraphBorders.top = parseBrc80(val, { preserveNone: true });
      break;
    case 0x6425: // sprmPBrcLeft80
      props.paragraphBorders ??= {};
      props.paragraphBorders.left = parseBrc80(val, { preserveNone: true });
      break;
    case 0x6426: // sprmPBrcBottom80
      props.paragraphBorders ??= {};
      props.paragraphBorders.bottom = parseBrc80(val, { preserveNone: true });
      break;
    case 0x6427: // sprmPBrcRight80
      props.paragraphBorders ??= {};
      props.paragraphBorders.right = parseBrc80(val, { preserveNone: true });
      break;
    case 0x6428: // sprmPBrcBetween80
      props.paragraphBorders ??= {};
      props.paragraphBorders.between = parseBrc80(val, { preserveNone: true });
      break;
    case 0x6629: // sprmPBrcBar80
      props.paragraphBorders ??= {};
      props.paragraphBorders.bar = parseBrc80(val, { preserveNone: true });
      break;
    case 0xC64E: // sprmPBrcTop
      props.paragraphBorders ??= {};
      props.paragraphBorders.top = parseBrcOperand(val, "paragraph top border", { preserveNone: true });
      break;
    case 0xC64F: // sprmPBrcLeft
      props.paragraphBorders ??= {};
      props.paragraphBorders.left = parseBrcOperand(val, "paragraph left border", { preserveNone: true });
      break;
    case 0xC650: // sprmPBrcBottom
      props.paragraphBorders ??= {};
      props.paragraphBorders.bottom = parseBrcOperand(val, "paragraph bottom border", { preserveNone: true });
      break;
    case 0xC651: // sprmPBrcRight
      props.paragraphBorders ??= {};
      props.paragraphBorders.right = parseBrcOperand(val, "paragraph right border", { preserveNone: true });
      break;
    case 0xC652: // sprmPBrcBetween
      props.paragraphBorders ??= {};
      props.paragraphBorders.between = parseBrcOperand(val, "paragraph between border", { preserveNone: true });
      break;
    case 0xC653: // sprmPBrcBar
      props.paragraphBorders ??= {};
      props.paragraphBorders.bar = parseBrcOperand(val, "paragraph bar border", { preserveNone: true });
      break;
    default:
      break;
  }
}

function parsePropRMarkOperand(val, context) {
  if (val.length < 8) {
    throw new Error(`Truncated MS-DOC ${context} PropRMarkOperand`);
  }
  const cb = val[0];
  if (cb !== 7) {
    throw new Error(`Out-of-spec MS-DOC ${context} PropRMarkOperand cb ${cb}`);
  }
  return {
    fPropRMark: val[1] !== 0,
    authorIndex: val.readInt16LE(2),
    date: val.readUInt32LE(4),
  };
}

// Brc80 structure (4 bytes): dptLineWidth(1), brcType(1), ico(1), dptSpace+flags(1)
// per MS-DOC-SPEC 19-structures-basic-types.md §Brc80
// Frame wrap styles per MS-DOC-SPEC sprmPWr
const FRAME_WRAP_NAMES = {
  0: "auto",
  1: "notBeside",
  2: "around",
  3: "none",
  4: "tight",
  5: "through",
};

// Frame vertical anchor per MS-DOC-SPEC sprmPPc pcVert
const FRAME_V_ANCHOR = ["margin", "page", "text", "margin"];

// Frame horizontal anchor per MS-DOC-SPEC sprmPPc pcHorz
const FRAME_H_ANCHOR = ["text", "margin", "page", "margin"];

function parseBrc80(val, { ignoreSpace = false, preserveNone = false } = {}) {
  if (val.length < 4) return null;
  const dptLineWidth = val[0];
  const brcType = val[1];
  const ico = val[2];
  const dptSpace = ignoreSpace ? 0 : (val[3] & 0x1F);
  if (brcType === 0 && !preserveNone) return null; // no border
  return {
    val: brcType === 0 ? "none" : brcTypeToBorderName(brcType, "paragraph border"),
    color: brcColorFromIco(ico),
    sz: String(dptLineWidth),
    space: String(dptSpace),
  };
}

function parseDcs(val) {
  const raw = val.readUInt16LE(0);
  const fdct = raw & 0x0007;
  const cl = (raw >> 3) & 0x001f;
  if (fdct !== 1 && fdct !== 2) {
    throw new Error(`Out-of-spec MS-DOC DCS fdct ${fdct}`);
  }
  if (cl < 1 || cl > 10) {
    throw new Error(`Out-of-spec MS-DOC DCS lines ${cl}`);
  }
  // MS-DOC-SPEC/19 DCS.fdct maps to ECMA-376 framePr@dropCap.
  return {
    frameDropCap: fdct === 1 ? "drop" : "margin",
    frameLines: cl,
  };
}

function setParagraphIndentPart(props, key, value) {
  Object.defineProperty(props, `_${key}`, {
    value,
    writable: true,
    configurable: true,
    enumerable: false,
  });
  updateParagraphLeftIndent(props);
}

function updateParagraphLeftIndent(props) {
  const hasLogicalNest = props._nest != null;
  const hasLogicalLeft = props._leftIndent != null;
  if (hasLogicalNest || hasLogicalLeft) {
    // MS-DOC-SPEC/16 sprmPNest adds to sprmPDxaLeft and supersedes sprmPNest80.
    props.leftIndent = (props._leftIndent ?? 0) + (props._nest ?? 0);
    return;
  }
  if (props._leftIndent80 != null || props._nest80 != null) {
    // MS-DOC-SPEC/16 sprmPNest80 adds to sprmPDxaLeft80.
    props.leftIndent = (props._leftIndent80 ?? 0) + (props._nest80 ?? 0);
  }
}

function parseBrcOperand(val, context, { ignoreSpace = false, preserveNone = false } = {}) {
  if (!val || val.length < 9 || val[0] !== 8) {
    throw new Error(`Out-of-spec ${context} BrcOperand length ${val?.length ?? 0}`);
  }
  const brc = val.subarray(1, 9);
  const brcType = brc[5];
  if (brcType === 0 && !preserveNone) return null;
  // MS-DOC-SPEC/19 BrcOperand contains a Brc, and BrcType 0x00 is
  // explicitly "No border" / ECMA-376 "none"; paragraph SPRMs can carry
  // that explicit record and WPS serializes it as a w:pBdr side.
  return {
    val: brcType === 0 ? "none" : brcTypeToBorderName(brcType, context),
    color: colorRefToHex(brc.subarray(0, 4)),
    sz: String(brc[4]),
    space: String(ignoreSpace ? 0 : (brc[6] & 0x1f)),
  };
}

function parseCFitTextOperand(val) {
  if (!val || val.length < 9 || val[0] !== 8) {
    throw new Error(`Out-of-spec character fitText CFitTextOperand length ${val?.length ?? 0}`);
  }
  const width = val.readInt32LE(1);
  if (width === 0) return null;
  if (width < 0) {
    // MS-DOC-SPEC/19 CFitTextOperand defines negative dxaFitText as the
    // minimum-width variant. OOXML w:fitText/@w:val is an unsigned
    // ST_TwipsMeasure, so do not guess a lossy mapping.
    throw new Error(`Unsupported MS-DOC character fitText minimum-width variant ${width}`);
  }
  return {
    width,
    id: val.readInt32LE(5),
  };
}

function parseFarEastLayoutOperand(val) {
  if (!val || val.length !== 7 || val[0] !== 6) {
    throw new Error(`Out-of-spec character FarEastLayoutOperand length ${val?.length ?? 0}`);
  }
  const ufel = val.readUInt16LE(1);
  const mustBeZero = 0x683c;
  if ((ufel & mustBeZero) !== 0) {
    throw new Error(`Out-of-spec character UFEL MUST-zero bits 0x${(ufel & mustBeZero).toString(16)}`);
  }

  const fTny = (ufel & 0x0001) !== 0;
  const fWarichu = (ufel & 0x0002) !== 0;
  const bracketCode = (ufel >> 8) & 0x7;
  const fTnyCompress = (ufel & 0x1000) !== 0;
  if (fWarichu && bracketCode > 4) {
    throw new Error(`Out-of-spec character UFEL iWarichuBracket value ${bracketCode}`);
  }
  if (!fTny && !fWarichu) return null;

  const layout = {
    id: val.readInt32LE(3),
  };
  if (fTny) {
    layout.vert = true;
    if (fTnyCompress) layout.vertCompress = true;
  }
  if (fWarichu) {
    layout.combine = true;
    const combineBrackets = eastAsianLayoutBracketFromCode(bracketCode);
    if (combineBrackets) layout.combineBrackets = combineBrackets;
  }
  return layout;
}

function eastAsianLayoutBracketFromCode(code) {
  switch (code) {
    case 0:
      return null;
    case 1:
      return "round";
    case 2:
      return "square";
    case 3:
      return "angle";
    case 4:
      return "curly";
    default:
      throw new Error(`Out-of-spec character UFEL iWarichuBracket value ${code}`);
  }
}

// MS-DOC-SPEC/19 BrcType values and their ST_Border references.
// 0x05 has no ECMA-376 reference but is specified as a thin single solid line,
// so it maps to the closest OOXML border value, single.
export const BRC_TYPE_NAMES = {
  0x00: "none",
  0x01: "single",
  0x03: "double",
  0x05: "single",
  0x06: "dotted",
  0x07: "dashed",
  0x08: "dotDash",
  0x09: "dotDotDash",
  0x0A: "triple",
  0x0B: "thinThickSmallGap",
  0x0C: "thickThinSmallGap",
  0x0D: "thinThickThinSmallGap",
  0x0E: "thinThickMediumGap",
  0x0F: "thickThinMediumGap",
  0x10: "thinThickThinMediumGap",
  0x11: "thinThickLargeGap",
  0x12: "thickThinLargeGap",
  0x13: "thinThickThinLargeGap",
  0x14: "wave",
  0x15: "doubleWave",
  0x16: "dashSmallGap",
  0x17: "dashDotStroked",
  0x18: "threeDEmboss",
  0x19: "threeDEngrave",
  0x1A: "outset",
  0x1B: "inset",
};

function brcTypeToBorderName(brcType, context) {
  const borderName = BRC_TYPE_NAMES[brcType];
  if (borderName) return borderName;
  throw new Error(`Out-of-spec ${context} BrcType ${brcType}`);
}

// ico color palette index per MS-DOC-SPEC §ICO (0x00=auto)
export function brcColorFromIco(ico) {
  if (ico === 0) return "auto";
  const colors = [
    null, "000000", "0000FF", "00FFFF", "00FF00", "FF00FF",
    "FF0000", "FFFF00", "FFFFFF", "000080", "008080",
    "008000", "800080", "800000", "808000", "808080", "C0C0C0",
  ];
  const color = colors[ico];
  if (color) return color;
  throw new Error(`Out-of-spec MS-DOC Ico color index ${ico}`);
}

function ww8TextColorHex(index) {
  if (index == null) return null;
  if (index >= 0 && index < WW8_TEXT_COLOR_INDEXES.length) {
    return WW8_TEXT_COLOR_INDEXES[index];
  }
  throw new Error(`Out-of-spec MS-DOC Ico text color index ${index}`);
}

export function colorRefToHex(colorRef) {
  if (!colorRef || colorRef.length < 4) {
    return null;
  }
  // MS-DOC-SPEC/19 COLORREF stores red, green, blue, fAuto bytes. fAuto MUST
  // be 0x00 or 0xFF; 0xFF is cvAuto, the application default color.
  const fAuto = colorRef[3];
  if (fAuto === 0xff) return "auto";
  if (fAuto !== 0x00) {
    throw new Error(`Out-of-spec MS-DOC COLORREF fAuto ${fAuto}`);
  }
  return `${colorRef[0].toString(16).padStart(2, "0")}${colorRef[1].toString(16).padStart(2, "0")}${colorRef[2].toString(16).padStart(2, "0")}`.toUpperCase();
}

function ww8ShadeColorHex(index, autoHex) {
  if (index == null) return null;
  if (index === 0) return autoHex;
  if (index >= 0 && index < WW8_TEXT_COLOR_INDEXES.length) {
    return WW8_TEXT_COLOR_INDEXES[index];
  }
  throw new Error(`Out-of-spec MS-DOC Ico shading color index ${index}`);
}

function parseWw8Shade(val) {
  if (val.length < 2) {
    return null;
  }

  const raw = val.readUInt16LE(0);
  const foreIndex = raw & 0x1f;
  const backIndex = (raw >> 5) & 0x1f;
  const styleIndex = (raw >> 10) & 0x3f;
  const foreHex = ww8ShadeColorHex(foreIndex, "000000");
  const backHex = ww8ShadeColorHex(backIndex, "FFFFFF");
  const strength = WW8_SHADING_STRENGTHS[styleIndex] ?? 0;

  if (styleIndex === 0) {
    return {
      val: "clear",
      color: "auto",
      fill: backHex,
    };
  }

  if (strength === 0) {
    return {
      val: "clear",
      color: "auto",
      fill: backHex,
    };
  }

  if (strength === 1000) {
    return {
      val: "clear",
      color: "auto",
      fill: foreHex,
    };
  }

  const mix = (fore, back) => Math.round((fore * strength + back * (1000 - strength)) / 1000);
  const fill = [0, 1, 2].map((i) => {
    const foreChannel = Number.parseInt(foreHex.slice(i * 2, i * 2 + 2), 16);
    const backChannel = Number.parseInt(backHex.slice(i * 2, i * 2 + 2), 16);
    return mix(foreChannel, backChannel).toString(16).toUpperCase().padStart(2, "0");
  }).join("");

  return {
    val: styleIndex === 38 ? "pct15" : "clear",
    color: "auto",
    fill,
  };
}

const IPAT_TO_ST_SHD = {
  0x0000: "clear",
  0x0001: "solid",
  0x0002: "pct5",
  0x0003: "pct10",
  0x0004: "pct20",
  0x0005: "pct25",
  0x0006: "pct30",
  0x0007: "pct40",
  0x0008: "pct50",
  0x0009: "pct60",
  0x000a: "pct70",
  0x000b: "pct75",
  0x000c: "pct80",
  0x000d: "pct90",
  0x000e: "horzStripe",
  0x000f: "vertStripe",
  0x0010: "reverseDiagStripe",
  0x0011: "diagStripe",
  0x0012: "horzCross",
  0x0013: "diagCross",
  0x0014: "thinHorzStripe",
  0x0015: "thinVertStripe",
  0x0016: "thinReverseDiagStripe",
  0x0017: "thinDiagStripe",
  0x0018: "thinHorzCross",
  0x0019: "thinDiagCross",
  0x0025: "pct12",
  0x0026: "pct15",
  0x002b: "pct35",
  0x002c: "pct37",
  0x002e: "pct45",
  0x0031: "pct55",
  0x0033: "pct62",
  0x0034: "pct65",
  0x0039: "pct85",
  0x003a: "pct87",
  0x003c: "pct95",
  0xffff: "nil",
};

const UNMAPPED_IPAT_VALUES = new Set([
  0x0023, 0x0024, 0x0027, 0x0028, 0x0029, 0x002a, 0x002d, 0x002f,
  0x0030, 0x0032, 0x0035, 0x0036, 0x0037, 0x0038, 0x003b, 0x003d,
]);

function parseShdOperand(val, context) {
  if (!val || val.length < 11 || val[0] !== 10) {
    throw new Error(`Out-of-spec ${context} SHDOperand length ${val?.length ?? 0}`);
  }
  const ipat = val.readUInt16LE(9);
  if (ipat === 0xffff) return null;
  const shdVal = IPAT_TO_ST_SHD[ipat];
  if (!shdVal) {
    if (UNMAPPED_IPAT_VALUES.has(ipat)) {
      throw new Error(`Unimplemented ${context} Ipat 0x${ipat.toString(16)} has no ECMA-376 ST_Shd mapping`);
    }
    throw new Error(`Out-of-spec ${context} Ipat 0x${ipat.toString(16)}`);
  }
  return {
    val: shdVal,
    color: colorRefToHex(val.subarray(1, 5)),
    fill: colorRefToHex(val.subarray(5, 9)),
  };
}

function ww8HighlightName(index) {
  const map = {
    0: "auto",
    1: "black",
    2: "blue",
    3: "cyan",
    4: "green",
    5: "magenta",
    6: "red",
    7: "yellow",
    8: "white",
    9: "darkBlue",
    10: "darkCyan",
    11: "darkGreen",
    12: "darkMagenta",
    13: "darkRed",
    14: "darkYellow",
    15: "darkGray",
    16: "lightGray",
  };
  return map[index] ?? "auto";
}

function underlineStyleFromCode(code) {
  switch (code) {
    case 0:
      return null;
    case 1:
    case 2:
      return "single";
    case 3:
      return "double";
    case 4:
      return "dotted";
    case 5:
      return null;
    case 6:
      return "thick";
    case 7:
      return "dash";
    case 9:
      return "dotDash";
    case 10:
      return "dotDotDash";
    case 11:
      return "wave";
    case 20:
      return "dottedHeavy";
    case 23:
      return "dashedHeavy";
    case 25:
      return "dashDotHeavy";
    case 26:
      return "dashDotDotHeavy";
    case 27:
      return "wavyHeavy";
    case 39:
      return "dashLong";
    case 55:
      return "dashLongHeavy";
    case 43:
      return "wavyDouble";
    default:
      throw new Error(`Out-of-spec MS-DOC Kul underline value ${code}`);
  }
}

function verticalAlignFromIss(code) {
  switch (code) {
    case 0:
      return null;
    case 1:
      return "superscript";
    case 2:
      return "subscript";
    default:
      throw new Error(`Out-of-spec MS-DOC Iss superscript/subscript value ${code}`);
  }
}

function emphasisMarkFromKcd(code) {
  // MS-DOC-SPEC/16 sprmCKcd: KCD values 0-4 map directly to OOXML
  // ST_Em values for emphasis marks.
  switch (code) {
    case 0:
      return "none";
    case 1:
      return "dot";
    case 2:
      return "comma";
    case 3:
      return "circle";
    case 4:
      return "underDot";
    default:
      throw new Error(`Out-of-spec MS-DOC Kcd emphasis value ${code}`);
  }
}

function textEffectFromSfxText(code) {
  // MS-DOC-SPEC/16 sprmCSfxText and LibreOffice WW8 import document Word's
  // legacy animation enum: 0 none, 1 Las Vegas lights, 2 blinking
  // background, 3 sparkle, 4 black ants, 5 red ants, 6 shimmer. OOXML
  // w:effect uses the same semantic values:
  // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.texteffect
  switch (code) {
    case 0:
      return "none";
    case 1:
      return "lights";
    case 2:
      return "blinkBackground";
    case 3:
      return "sparkle";
    case 4:
      return "antsBlack";
    case 5:
      return "antsRed";
    case 6:
      return "shimmer";
    default:
      throw new Error(`Out-of-spec MS-DOC SfxText text effect value ${code}`);
  }
}

function lineBreakClearFromLbcOperand(code) {
  // MS-DOC-SPEC/19 LBCOperand values map directly to OOXML w:br clear.
  switch (code) {
    case 0:
      return "none";
    case 1:
      return "left";
    case 2:
      return "right";
    case 3:
      return "all";
    default:
      throw new Error(`Out-of-spec MS-DOC LBCOperand value ${code}`);
  }
}

function parseTabsOperand(val) {
  const payload = val.subarray(1);
  if (payload.length < 2) return [];
  const deletedCount = payload[0];
  const addedCount = payload[1 + deletedCount * 2] ?? 0;
  const positionsOffset = 2 + deletedCount * 2;
  if (positionsOffset + addedCount * 2 > payload.length) {
    throw new Error("Truncated tab-stop SPRM: missing tab positions");
  }

  const tabs = [];
  const tbdOffset = positionsOffset + addedCount * 2;
  if (tbdOffset + addedCount > payload.length) {
    throw new Error("Truncated tab-stop SPRM: missing tab descriptors");
  }
  for (let i = 0; i < addedCount; i += 1) {
    const descriptor = payload[tbdOffset + i];
    const rawAlign = descriptor & 0x07;
    const rawLeader = (descriptor >> 3) & 0x07;
    // WPS sometimes stores the alignment in bits 3-5 and leader in bits 0-2;
    // if rawAlign is unrecognised, try the swapped layout.
    let alignCode = rawAlign;
    let leaderCode = rawLeader;
    if (!(rawAlign in TAB_ALIGNMENT_MAP) && (rawLeader in TAB_ALIGNMENT_MAP)) {
      alignCode = rawLeader;
      leaderCode = rawAlign;
    }
    tabs.push({
      position: payload.readInt16LE(positionsOffset + i * 2),
      alignment: TAB_ALIGNMENT_MAP[alignCode] ?? "left",
      leader: TAB_LEADER_MAP[leaderCode] ?? null,
    });
  }
  return tabs;
}

const TAB_ALIGNMENT_MAP = { 0: "left", 1: "center", 2: "right", 3: "decimal", 4: "bar" };
const TAB_LEADER_MAP = { 0: null, 1: "dot", 2: "hyphen", 3: "underscore", 4: "heavy", 5: "middleDot" };

function tabAlignmentFromCode(code) {
  return TAB_ALIGNMENT_MAP[code] ?? "left";
}

function tabLeaderFromCode(code) {
  return TAB_LEADER_MAP[code] ?? null;
}
