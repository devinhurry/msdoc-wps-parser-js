# MS-DOC / WPS Feature Audit

> Audit date: **2026-07-10**
> Audited change set: **feature-completion commit based on `30791b0`**
> Verification: **142/142 tests pass**; sample 9 and sample 11 packages pass ZIP/XML validation and render to PDF with LibreOffice.
> Scope: `src/cfb.js`, `src/wps.js`, `src/property-set.js`, `src/word-binary.js`, `src/sprm.js`, `src/docx.js`, `src/lcid.js`, `test/wps.test.js`, and the checked-in sample corpus.

This audit supersedes the earlier audit. The earlier version described fields, secondary stories, inline pictures, metadata, page borders, and header/footer drawings as missing; those features are now implemented.

The project remains a focused WPS/Word 97–2003 binary-to-DOCX converter rather than a complete implementation of every historical MS-DOC feature. The supported path is intentionally strict: required structures are parsed explicitly, malformed inputs fail fast, and unsupported variants are rejected instead of guessed.

## Status definitions

| Status | Meaning |
|---|---|
| ✅ | Parsed from an explicit binary structure and emitted or exposed semantically. |
| 🟡 | Implemented for a documented subset; unsupported subvariants fail fast. |
| 🔹 | Parsed/exposed, but not fully represented in DOCX. |
| ❌ | Not implemented. |
| 🚫 | Deliberately outside the approved scope or explicitly rejected. |

## 1. Executive assessment

### Completed conversion areas

- CFB stream reading, FIB/DOP parsing, complex CLX text, and non-complex Unicode text.
- Story-wide PAPX and CHPX extraction for the main story and all supported secondary stories.
- Paragraph/run formatting, styles, fonts, numbering, revisions, sections, tables, bookmarks, and settings.
- Authoritative `PlcfHdd` header/footer reconstruction and compact WPS omitted-PLC reconstruction based on parsed story/shape linkage.
- Fields in all seven MS-DOC field stories, including nested fields and PAGE fields inside textboxes.
- Footnotes, endnotes, comments, main-story references, and comment metadata.
- Main and header textboxes with parsed `FTXBXS` shape linkage.
- Inline PICF/OfficeArt images and DOCX media relationships.
- Header/footer OfficeArt lines, textboxes, inline images, and general relationship-free DrawingML shapes.
- SummaryInformation, DocumentSummaryInformation, and custom properties.
- Non-empty section page borders and picture border SPRMs.
- LID-driven decoding of compressed complex-file text pieces.
- Expanded table cell flags, including horizontal merge, text direction, fit-text, no-wrap, vertical merge, and hide-mark.

### Deliberate or remaining boundaries

- Encrypted/obfuscated documents, VBA, digital signatures, IRM/protected-content recovery, and OLE object conversion are outside this implementation pass.
- Non-complex single-byte documents are rejected; the non-complex path currently requires Unicode text.
- PICT image payloads are rejected because no interoperable DOCX media mapping is implemented.
- DrawingML shapes whose stored `e2o` graphic requires external relationships are not transplanted yet.
- Tight/through wrap requiring a contour polygon is rejected when no explicit polygon was parsed.
- CFB writing and hierarchical storage path APIs are not implemented.
- Glossary/AutoText, embedded fonts, picture bullets, charts, SmartArt, and arbitrary embedded packages are not converted.
- Complete conditional table-style evaluation and every obsolete/revision-only MS-DOC SPRM are not implemented.

## 2. Public API and container

### 2.1 Public API

| Feature | Status | Notes |
|---|---:|---|
| Read `Buffer` / `Uint8Array` | ✅ | `readWps`. |
| Read file | ✅ | `readWpsFile`. |
| Return semantic parsed model | ✅ | Stories, properties, tables, styles, lists, fields, drawings, metadata, and references. |
| Convert model to DOCX buffer | ✅ | `wpsToDocxBuffer`. |
| Convert source file to DOCX file | ✅ | `convertWpsToDocxFile`. |
| CLI text inspection | ✅ | `bin/msdoc-wps-parser.js`. |
| CLI WPS/DOC-to-DOCX conversion | ✅ | `bin/msdoc-wps-to-docx.js`. |

### 2.2 Compound File Binary reader

| Capability | Status | Notes |
|---|---:|---|
| Signature, byte order, sector validation | ✅ | Invalid/big-endian containers fail fast. |
| 512-byte and 4096-byte sectors | ✅ | Read from the CFB header. |
| DIFAT/FAT chains | ✅ | Includes chained DIFAT sectors. |
| MiniFAT and mini stream | ✅ | Small streams are read correctly. |
| Directory entries | ✅ | Names, types, siblings, children, CLSID, state, timestamps, sector, size. |
| Stream enumeration and exact-name lookup | ✅ | Sufficient for standard Word streams. |
| Hierarchical storage path lookup | ❌ | Current lookup is flat by exact stream name. |
| CFB writer | ❌ | DOCX output is ZIP-based; source CFB rewriting is not provided. |

### 2.3 Consumed streams

| Stream | Status | Use |
|---|---:|---|
| `WordDocument` | ✅ | FIB, text, FKP, SEPX, and non-complex text. |
| `0Table` / `1Table` | ✅ | Selected from `FibBase.fWhichTblStm`. |
| `Data` | ✅ | Huge PAPX/table grpprl expansion plus PICF/OfficeArt image payloads. |
| `\u0005SummaryInformation` | ✅ | Core and extended metadata. |
| `\u0005DocumentSummaryInformation` | ✅ | Extended and custom metadata. |
| Object/VBA/signature/protection storages | 🚫 | Outside the approved scope. |

## 3. Text and story model

### 3.1 Text extraction

| Feature | Status | Notes |
|---|---:|---|
| Complex CLX/Pcd piece table | ✅ | Strict CLX and CP validation. |
| Unicode pieces | ✅ | UTF-16LE. |
| Compressed complex pieces | ✅ | Explicit ANSI decoder selected from parsed `FibBase.lid`. |
| Supported compressed code pages | ✅ | GBK, Big5, Shift-JIS, EUC-KR, Windows-874, and Windows-1250 through Windows-1258. |
| Unknown compressed code page | 🚫 | Fails rather than assuming Windows-1252. |
| Non-complex Unicode | ✅ | Explicit `fcMin`/`fcMac` and story-count validation. |
| Non-complex single-byte | 🚫 | Rejected because the current FIB/version path does not establish a safe decoder. |

### 3.2 Story ranges and formatting

| Story | Text | PAPX | CHPX | Semantic DOCX output |
|---|---:|---:|---:|---:|
| Main document | ✅ | ✅ | ✅ | ✅ |
| Footnotes | ✅ | ✅ | ✅ | ✅ |
| Headers/footers | ✅ | ✅ | ✅ | ✅ |
| Comments/annotations | ✅ | ✅ | ✅ | ✅ |
| Endnotes | ✅ | ✅ | ✅ | ✅ |
| Main textboxes | ✅ | ✅ | ✅ | ✅ |
| Header textboxes | ✅ | ✅ | ✅ | ✅ |

Story boundaries come from `FibRgLw97` character counts. PAPX/CHPX are extracted over the complete raw story sequence and then sliced by explicit story offsets.

## 4. Fields, notes, comments, and textboxes

### 4.1 Fields

| Feature | Status | Notes |
|---|---:|---|
| Main-story `PlcffldMom` | ✅ | Begin/separate/end characters validated against story text. |
| Header `PlcffldHdr` | ✅ | Emitted in real header/footer parts. |
| Footnote `PlcffldFtn` | ✅ | Emitted in footnotes. |
| Comment `PlcffldAtn` | ✅ | Emitted in comments. |
| Endnote `PlcffldEdn` | ✅ | Emitted in endnotes. |
| Main textbox `PlcffldTxbx` | ✅ | Emitted inside `w:txbxContent`. |
| Header textbox `PlcffldHdrTxbx` | ✅ | Includes PAGE footer textboxes. |
| Nested fields | ✅ | Strict stack/state-machine validation. |
| Field flags | ✅ | Parsed authoritative flags are preserved where DOCX has an equivalent. |

### 4.2 Notes and comments

| Feature | Status | Notes |
|---|---:|---|
| Footnote reference/text PLCs | ✅ | Count, CP ordering, markers, and paragraph endings validated. |
| Endnote reference/text PLCs | ✅ | Same strict validation. |
| Comment reference/text PLCs | ✅ | Supports zero-length comment ranges. |
| Comment owners and initials | ✅ | Emitted in `comments.xml`. |
| Main-story note/comment references | ✅ | Corresponding OOXML reference elements emitted. |
| Note separator/continuation variants | 🟡 | Standard DOCX separator entries are produced; uncommon source-specific custom separators are not separately modeled. |

### 4.3 Textboxes

| Feature | Status | Notes |
|---|---:|---|
| `PlcftxbxTxt` / `PlcfHdrtxbxTxt` | ✅ | Strict length and CP ordering. |
| `FTXBXS.lid` shape linkage | ✅ | Text is assigned by parsed shape id, not row-count/signature inference. |
| Reusable textbox tail | ✅ | Explicit reusable flag handled. |
| Main-story textboxes | ✅ | Drawing anchor plus `wps:txbx/w:txbxContent`. |
| Header/footer textboxes | ✅ | PAGE field footer samples are reconstructed. |

## 5. Headers, footers, and sections

| Feature | Status | Notes |
|---|---:|---|
| Authoritative `PlcfHdd` parsing | ✅ | Boundary count/order and story extent validated. |
| Odd/even/first headers and footers | ✅ | References follow section flags and parsed header/footer ranges. |
| Multiple sections | ✅ | Section-specific relationships and `sectPr` references. |
| Compact WPS omitted `PlcfHdd` | 🟡 | Reconstructed from parsed story counts and parsed `FTXBXS`/shape-anchor linkage; malformed ambiguity fails. |
| Page size/orientation/margins/gutter | ✅ | From SEPX section SPRMs. |
| Columns and separator | ✅ | Explicit column widths/spaces where present. |
| Vertical alignment | ✅ | Emitted from parsed section property. |
| Page-number restart/format | ✅ | Restart is honored only when explicitly enabled. |
| Document grid | ✅ | Line/character grids and settings dependencies. |
| Footnote/endnote section settings | ✅ | Position, numbering, restart, start values. |
| Non-empty page borders | ✅ | BRC80/BRC operands and `SPgbPropOperand` display/z-order/offset. |
| Default/empty page borders | ✅ | No synthetic border attributes are emitted. |

## 6. Drawings and pictures

### 6.1 Inline pictures

| Feature | Status | Notes |
|---|---:|---|
| Picture marker validation | ✅ | Requires U+0001, `sprmCPicLocation`, and `sprmCFSpec`. |
| PICF parsing | ✅ | `lcb`, `cbHeader`, MFPF, PICMID, scaling, crop, and bounds validated. |
| OfficeArt inline container | ✅ | Parses SpContainer, BSE, and embedded blip records. |
| PNG | ✅ | Embedded as DOCX media. |
| JPEG | ✅ | Embedded as DOCX media. |
| DIB | ✅ | Converted to a BMP file with an explicit file header. |
| TIFF | ✅ | Embedded as DOCX media. |
| EMF / WMF | ✅ | Handles explicit compression mode and raw-deflate payloads. |
| PICT | 🚫 | Rejected; no interoperable DOCX target is implemented. |
| Crop | ✅ | Explicit OfficeArt crop properties 256–259. |
| Name/description | ✅ | From OfficeArt properties 896/897. |
| Picture borders | ✅ | All four BRC80 and all four modern picture-border SPRMs. |
| Binary data / OLE placeholders | 🚫 | Rejected rather than emitted as a guessed image. |

### 6.2 Floating shapes

| Feature | Status | Notes |
|---|---:|---|
| Main/header `PlcSpa` anchors | ✅ | CP, shape id, bounds, wrap flags, and drawing order. |
| Horizontal line (`shapeType=20`) | ✅ | DrawingML plus VML fallback when available. |
| Textbox (`shapeType=202`) | ✅ | Parsed or synthesized relationship-free rectangle geometry. |
| Other relationship-free `e2o` graphics | ✅ | Existing DrawingML graphic is normalized and anchored. |
| Position/wrap mapping | ✅ | Explicit `bx`, `by`, `wr`, `wrk`, and OfficeArt `posh` mapping. |
| Shape relative height | ✅ | Normalized from explicit story drawing order. |
| External shape relationships | 🚫 | `r:embed`/`r:link` graphics are rejected until their relationship parts are transplanted. |
| Tight/through contour without polygon | 🚫 | Fails rather than inventing a wrap polygon. |

## 7. Formatting, styles, numbering, and revisions

### 7.1 Character formatting

Implemented semantic groups include:

- Latin, East Asian, high-ANSI, and complex-script fonts and LIDs.
- Font size, complex-script size, position, vertical alignment, kerning, spacing, width scaling, and fit-text.
- Bold/italic and complex-script variants, strike/double-strike, outline, shadow, caps/small caps, emboss, imprint.
- Underline style/color, text color, highlight, shading, character borders, and automatic color handling.
- Visibility/no-proof/field-vanish/special state and East Asian layout/emphasis.
- Revision author/date/property marks and insertion/deletion states.

Restricted variants fail explicitly, including negative minimum-width character fit-text.

### 7.2 Paragraph formatting

Implemented semantic groups include:

- Style id, alignment, bidi, keep controls, page break, widow/orphan, contextual spacing, outline level.
- Twip and character-unit indents with explicit source provenance.
- Line/before/after spacing, automatic spacing flags, tabs, line numbering suppression.
- Paragraph borders, shading, frame properties, table depth/cell markers, and paragraph-mark run properties.
- Paragraph property revisions using `sprmPWall` and `sprmPPropRMark`.

### 7.3 Styles and numbering

| Feature | Status | Notes |
|---|---:|---|
| STSH styles | ✅ | Paragraph, character, table, and numbering styles. |
| Built-in style mapping | ✅ | Uses the documented `sti` mapping rather than name inference. |
| Based-on/next/link relationships | ✅ | Validated and emitted. |
| Latent style metadata | ✅ | Emitted where parsed. |
| FFN font table | ✅ | Font names, families, charset, pitch, Panose/signature data. |
| LSTF/LVL lists | ✅ | Levels, numbering text, alignment, indents, and run properties. |
| LFO/LFOLVL overrides | 🟡 | Start-at overrides implemented; full embedded LVL formatting override is rejected. |
| Picture bullets | ❌ | Not implemented. |

### 7.4 Revisions and bookmarks

| Feature | Status | Notes |
|---|---:|---|
| Standard bookmarks | ✅ | Start/end ordering, collapsed bookmarks, BKC validation. |
| `_GoBack` from `Selsf` | ✅ | Only when the parsed selection lies in the authoritative main-story extent. |
| Revision author table | ✅ | `SttbfRMark`. |
| Insert/delete revisions | ✅ | Author and DTTM metadata. |
| Character property revisions | ✅ | Emitted as run-property changes where applicable. |
| Paragraph property revisions | ✅ | Emitted as `w:pPrChange`. |
| Revision reason ids | 🔹 | Parsed but OOXML has no direct equivalent in the current output model. |

## 8. Tables

| Area | Status | Notes |
|---|---:|---|
| Table/row discovery | ✅ | PAPX table depth and row-end properties. |
| Row/cell geometry | ✅ | `TDefTable`, grid boundaries, preferred widths, and row heights. |
| Horizontal/vertical merge | ✅ | Includes TCGRF horizontal and vertical merge states. |
| Cell text direction | ✅ | Strict TCGRF enum mapping. |
| Fit text / no wrap / hide mark | ✅ | From explicit cell flags. |
| Table/cell width types | ✅ | Auto, percent, and dxa with strict range checks. |
| Table indent/alignment/autofit | ✅ | No inferred alignment from geometry. |
| Cell margins/padding | ✅ | From explicit table SPRMs. |
| Table and cell borders | ✅ | BRC/BRC80, including explicit `nil`. |
| Cell shading | ✅ | Explicit parsed shading only. |
| Header rows and cant-split | ✅ | Emitted from row flags. |
| Nested tables | ✅ | Table depth and inner-cell markers. |
| Table styles | ✅ | Parsed style id and supported table style definitions. |
| Full conditional table-style evaluation | 🟡 | Core style output exists; every historical conditional/mutation SPRM is not evaluated. |
| Table revision records | ❌ | Not converted. |

## 9. Document settings and metadata

### 9.1 Settings

- Parsed default tab stop; `WpsCustomData` is not used as the source of truth.
- Facing pages/even-and-odd headers, mirror margins, protection flags, track revisions, font embedding flags.
- Hyphenation zone, compatibility flags, XML validation flags, typography, drawing grid, display grid, and shape defaults.
- East Asian document-grid requirements fail fast when mandatory DOP structures are absent.

### 9.2 Metadata

| Feature | Status | Notes |
|---|---:|---|
| OLE property-set header/sections | ✅ | Strict offsets, sizes, property table, code page. |
| Common scalar types | ✅ | `VT_I2`, `VT_I4`, `VT_UI4`, `VT_R8`, `VT_BOOL`, `VT_LPSTR`, `VT_LPWSTR`, `VT_FILETIME`. |
| Core properties | ✅ | Title, subject, creator, keywords, description, revision, dates, category, status, version. |
| Extended properties | ✅ | Template, manager, company, pages, words, characters, lines, paragraphs, application, security. |
| Custom property dictionary | ✅ | Emitted as `docProps/custom.xml`. |
| Converter option overrides | ✅ | Source metadata is overridden only by explicitly supplied options. |
| Total editing time | 🔹 | Parsed, but not emitted until source-unit behavior is validated across WPS variants. |
| Arbitrary vector/array/blob property types | ❌ | Unsupported property types are not converted. |

## 10. DOCX package output

| Part/capability | Status |
|---|---:|
| `[Content_Types].xml` and package relationships | ✅ |
| `word/document.xml` | ✅ |
| `word/styles.xml`, `numbering.xml`, `fontTable.xml`, `settings.xml`, theme | ✅ |
| Header/footer parts and relationships | ✅ |
| Footnotes/endnotes/comments parts | ✅ |
| Media parts and image relationships | ✅ |
| Core/extended/custom property parts | ✅ |
| XML escaping and preserved whitespace | ✅ |
| ZIP package integrity | ✅ |
| OPC signatures/encryption | 🚫 |
| Macro-enabled DOCM output | 🚫 |

## 11. Fail-fast guarantees

The parser rejects, rather than guesses, at least the following conditions:

- encrypted/obfuscated FIB variants;
- missing mandatory streams, DOP, FKP/PLC structures, or invalid offsets/counts;
- malformed CLX/Pcd, story ranges, fields, notes, comments, textboxes, and shape linkage;
- unknown compressed-text LID/code page;
- non-complex non-Unicode text;
- unsupported PICF/blip/container records and invalid image bounds;
- PICT and OLE/binary picture placeholders;
- unsupported drawing relationship dependencies or contour wrapping;
- invalid section/table/border/width enum values and conflicting SPRMs;
- LFO full formatting overrides that are not yet implemented.

This behavior is intentional and follows the project rule that missing binary evidence must not be replaced by a heuristic default.

## 12. Verification record

The feature-completion working tree was validated with:

```text
node --check src/*.js
npm test
# 142 tests, 142 pass, 0 fail

git diff --check

# For sample9 and sample11:
node bin/msdoc-wps-to-docx.js <input.wps> <output.docx>
unzip -t <output.docx>
xmllint --noout on every XML/.rels part
soffice --headless --convert-to pdf <output.docx>
```

Both generated DOCX packages passed ZIP and XML validation and LibreOffice produced PDFs successfully.

## 13. Regression coverage added by this pass

- Authoritative header/footer story parsing and malformed boundary rejection.
- Nested field PLC parsing and field-text validation.
- Footnote/endnote/comment/textbox PLC parsing and DOCX emission.
- Sample 9 inline PNG extraction, header relationship/media packaging, line shapes, and footer textbox.
- Sample 11 PAGE-field textbox footer.
- OLE SummaryInformation metadata preservation.
- Picture border and non-empty section page-border SPRMs.
- Compressed-piece code page selection from `FibBase.lid`.
- Story-wide PAPX/CHPX behavior across secondary stories.

## 14. Recommended next expansion order

If the converter is expanded beyond the approved scope, the highest-value sequence is:

1. external relationships for arbitrary `e2o` floating graphics;
2. non-complex single-byte documents with version- and LID-authoritative decoding;
3. remaining LFO full-level overrides and conditional table styles;
4. embedded font and picture-bullet conversion;
5. glossary/AutoText and additional auxiliary stories;
6. only then, if explicitly required, OLE/VBA/protection-related package preservation.
