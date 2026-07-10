import { readFile } from "node:fs/promises";
import { CompoundFile } from "./cfb.js";
import { extractWordBinaryDocument } from "./word-binary.js";
import { parseOlePropertySet } from "./property-set.js";

export async function readWpsFile(filePath) {
  const buffer = await readFile(filePath);
  return readWps(buffer);
}

export function readWps(input) {
  const cfb = new CompoundFile(input);
  if (!cfb.hasStream("WordDocument")) {
    throw new Error("Invalid WPS file: missing WordDocument stream");
  }

  const wordDocument = cfb.readStream("WordDocument");
  const table0 = cfb.hasStream("0Table") ? cfb.readStream("0Table") : null;
  const table1 = cfb.hasStream("1Table") ? cfb.readStream("1Table") : null;
  const data = cfb.hasStream("Data") ? cfb.readStream("Data") : null;
  const document = extractWordBinaryDocument({ wordDocument, table0, table1, data });
  const summaryInformation = cfb.hasStream("\u0005SummaryInformation")
    ? parseOlePropertySet(cfb.readStream("\u0005SummaryInformation"), { kind: "summary" })
    : null;
  const documentSummaryInformation = cfb.hasStream("\u0005DocumentSummaryInformation")
    ? parseOlePropertySet(cfb.readStream("\u0005DocumentSummaryInformation"), { kind: "documentSummary" })
    : null;

  return {
    type: "wps-ole2-word-binary",
    streams: cfb.listStreams(),
    metadata: {
      ...(summaryInformation?.properties ?? {}),
      ...(documentSummaryInformation?.properties ?? {}),
      customProperties: documentSummaryInformation?.customProperties ?? [],
    },
    ...document,
  };
}
