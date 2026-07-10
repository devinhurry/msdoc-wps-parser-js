# Implementation Status

> Status date: **2026-07-10**
>
> Implementation snapshot: **feature-completion commit based on `30791b0`**
>
> Detailed audit: [`MS-DOC-SPEC/feature-audit.md`](MS-DOC-SPEC/feature-audit.md)

## 1. Completion summary

The approved feature-completion pass is complete.

The broader audit contains 158 explicit status entries. Nine entries are deliberately excluded or rejected and are not counted as unfinished in-scope work.

| Status | Count | Percentage of 149 in-scope entries |
|---|---:|---:|
| Fully implemented | 138 | **92.6%** |
| Partial or parsed-only | 6 | **4.0%** |
| Not implemented | 5 | **3.4%** |
| Total in-scope | 149 | **100%** |
| Deliberately excluded/rejected | 9 | Not included |

The implementation is therefore:

- **100% complete for the approved implementation scope**;
- **92.6% fully complete across the broader in-scope audit**;
- **96.6% implemented or partially implemented across that audit**.

## 2. Completion by area

Percentages below exclude deliberately rejected/out-of-scope entries.

| Area | Complete | Partial | Missing | Fully complete |
|---|---:|---:|---:|---:|
| Public API and CFB container | 18 | 0 | 2 | **90.0%** |
| Text and story model | 33 | 0 | 0 | **100%** |
| Fields, notes, comments, and textboxes | 19 | 1 | 0 | **95.0%** |
| Headers, footers, and sections | 11 | 1 | 0 | **91.7%** |
| Drawings and pictures | 17 | 0 | 0 | **100%** |
| Formatting, styles, numbering, and revisions | 12 | 2 | 1 | **80.0%** |
| Tables | 13 | 1 | 1 | **86.7%** |
| Settings and metadata | 6 | 1 | 1 | **75.0%** |
| DOCX package output | 9 | 0 | 0 | **100%** |

## 3. Major completed features

- Complex CLX/Pcd text and non-complex Unicode text extraction.
- LID-authoritative decoding for compressed complex-file text pieces.
- Story-wide PAPX and CHPX for the main and all supported secondary stories.
- Paragraph/run formatting, sections, styles, fonts, lists, tables, settings, bookmarks, and revisions.
- Authoritative `PlcfHdd` header/footer story conversion.
- Fields in all seven MS-DOC field stories, including nested fields.
- Footnotes, endnotes, comments, and main-story references.
- Main and header/footer textboxes with parsed `FTXBXS` shape linkage.
- Inline PICF/OfficeArt images and DOCX media relationships.
- PNG, JPEG, TIFF, DIB/BMP, EMF, and WMF payloads.
- Image crop, dimensions, name, description, and picture borders.
- Header/footer OfficeArt lines, textboxes, images, and relationship-free DrawingML shapes.
- Non-empty section page borders and `SPgbPropOperand` behavior.
- OLE SummaryInformation, DocumentSummaryInformation, and custom properties.
- Table TCGRF horizontal/vertical merge, text direction, fit-text, no-wrap, and hide-mark.
- ZIP-based DOCX output with header/footer, note, comment, media, metadata, style, and numbering parts.

## 4. Explicitly missing features

### 4.1 Hierarchical CFB storage-path lookup

**Current state:** CFB directory entries are parsed, but streams are looked up by flat exact name.

**Why missing:** The Word streams used by the supported converter are uniquely named at the root. Reconstructing the CFB directory red-black trees is primarily required for nested OLE, VBA, signature, and embedded-package storages, which are outside the current conversion path.

**Needed:** Validated parent/child hierarchy reconstruction, path-based stream APIs, duplicate-name handling, cycle detection, and nested-storage fixtures.

### 4.2 CFB writing

**Current state:** The project reads source CFB and writes DOCX as an OPC ZIP package.

**Why missing:** Writing legacy DOC/WPS files is not required for WPS-to-DOCX conversion. A correct CFB writer would require FAT, DIFAT, MiniFAT, sector allocation, directory-tree generation, and stream-threshold handling as a separate storage subsystem.

**Needed:** Adopt or build a fully tested CFB writer with Office, WPS, and LibreOffice round-trip validation.

### 4.3 Picture bullets

**Current state:** Text and symbol numbering are supported; `sprmCPbiIBullet` and `sprmCPbiGrf` are not semantically converted.

**Why missing:** Picture bullets require the hidden `_PictureBullets` bookmark, Bullet Pictures document CP resolution, PICF extraction, list-level association, and numbering-picture relationships. This is a different pipeline from ordinary inline pictures, and no authoritative WPS picture-bullet fixture is currently available.

**Needed:** Parse both SPRMs, resolve the hidden bookmark and image CP, emit numbering-picture definitions/relationships, and add WPS regression fixtures.

### 4.4 Table revision records

**Current state:** Text, character-property, and paragraph-property revisions are supported; table-property revisions such as `sprmTPropRMark` are not.

**Why missing:** The current table model stores final effective geometry and formatting. Table revisions require previous and new table/row/cell property states, structural insertion/deletion semantics, author/date metadata, and correct `tblPrChange`, `trPrChange`, or `tcPrChange` scope.

**Needed:** A revision-aware table history model and WPS fixtures containing structural and formatting table revisions.

### 4.5 Arbitrary OLE property-set vector, array, and blob types

**Current state:** Scalar types required by current samples are supported: `VT_I2`, `VT_I4`, `VT_UI4`, `VT_R8`, `VT_BOOL`, `VT_LPSTR`, `VT_LPWSTR`, and `VT_FILETIME`.

**Why missing:** Vector, array, variant, blob, clipboard, stream, storage, and object types need recursive alignment-sensitive parsing. Several types have no safe lossless mapping to DOCX custom properties, so converting them to strings or Base64 would invent semantics.

**Needed:** Preserve source variant types, add recursive vector/array parsing, define exact `docPropsVTypes` mappings, and explicitly preserve or reject binary/object values.

## 5. Partial or parsed-only features

### 5.1 Custom footnote/endnote separator stories

**Implemented:** `PlcfHdd` identifies all six separator and continuation stories.

**Incomplete:** DOCX generation currently emits standard separator and continuation-separator paragraphs instead of rendering custom source separator story content.

**Reason:** The separator stories are not yet wired through their PAPX, CHPX, fields, drawings, and special note IDs. There is no current WPS fixture with non-empty custom separator content.

### 5.2 Compact WPS files that omit `PlcfHdd`

**Implemented:** Compatibility paths reconstruct existing WPS samples from parsed `ccpHdd`, sections, fields, textbox linkage, and shape-anchor CPs.

**Incomplete:** The path cannot be considered a general authoritative implementation because the source omits the header/footer boundary PLC.

**Reason:** Some compact CR-only files do not expose unambiguous story allocation through currently parsed structures. Additional text/geometry pattern rules would violate the project rule against heuristics. A WPS-specific authoritative boundary property must be identified.

### 5.3 Full LFO embedded-level formatting overrides

**Implemented:** `LFOLVL` start-number overrides.

**Incomplete:** Entries with `fFormatting` are rejected.

**Reason:** Such entries contain a complete variable-length embedded LVL with paragraph and character grpprls, level text, numbering format, alignment, and restart behavior. It must be parsed and merged with the base list level as a unit; applying only part of it would be incorrect.

### 5.4 Revision-reason IDs

**Implemented:** Insertion/deletion reason IDs are parsed and kept in run identity so differently marked runs are not merged.

**Incomplete:** They are not emitted into DOCX revision markup.

**Reason:** Standard WordprocessingML `w:ins` and `w:del` elements have no direct standard attribute for the MS-DOC revision-reason ID. Emitting it as a comment or custom value would invent a mapping without evidence.

### 5.5 Full conditional table-style evaluation

**Implemented:** Base table styles, style IDs, direct table/cell formatting, selected built-in behavior, and Table Grid output.

**Incomplete:** The complete `w:tblStylePr` condition cascade is not generated.

**Reason:** Correct evaluation requires conditional style UPX data, CNF masks, table-look flags, row/cell position, style inheritance, banding, corners, and direct-formatting precedence. The current model does not retain this complete cascade, and implementing isolated conditions would be sample-specific.

### 5.6 Total editing-time emission

**Implemented:** SummaryInformation property 10 is parsed as a FILETIME duration and converted to minutes.

**Incomplete:** The value is not emitted as DOCX `<TotalTime>`.

**Reason:** Current WPS samples are inconsistent: one produces a plausible one-minute duration, while another produces `157270560` minutes. Emitting, clamping, or reinterpreting that value without establishing WPS behavior would be heuristic.

## 6. Deliberately excluded or rejected features

The following features are not counted as unfinished work in the approved scope:

- encrypted or obfuscated Word binary documents;
- VBA preservation or conversion;
- digital signatures and IRM/protected-content recovery;
- OLE object conversion;
- PICT image payloads;
- non-complex single-byte documents without an authoritative decoder path;
- DrawingML graphics requiring untransplanted external relationships;
- tight/through shape wrapping without a parsed contour polygon;
- OPC package signatures/encryption and macro-enabled DOCM output.

These cases fail explicitly instead of using guessed fallback behavior.

## 7. Verification status

The current implementation has been validated with:

```text
node --check src/*.js
npm test
git diff --check
```

Current regression result:

```text
142 tests
142 passed
0 failed
```

Sample 9 and sample 11 were also converted independently and validated by:

- ZIP package integrity checks;
- `xmllint` over every XML and relationship part;
- successful LibreOffice DOCX-to-PDF rendering.

## 8. Recommended next implementation order

1. Custom note separator story emission.
2. Full LFOLVL formatting overrides.
3. Conditional table-style cascade.
4. Picture bullets, after obtaining WPS fixtures.
5. Table-property revisions with a before/after table model.
6. External relationships for arbitrary `e2o` graphics.
7. Non-complex single-byte decoding based on authoritative FIB/version evidence.
8. Additional OLE property-set types.
9. Hierarchical CFB paths only when embedded-object support is approved.
10. CFB writing only if legacy DOC/WPS output becomes a project requirement.
