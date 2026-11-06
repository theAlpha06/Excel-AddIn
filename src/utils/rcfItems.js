import SalesforceService from "../services/SalesforceService";

// The quote cost item sheets carry an "RCF Item" column for QuoteCostItem__c.RCF_Item__c
// (checkbox). The bulk-insert Apex does not set it, so after a full push the add-in sets it on the
// new cost lines itself, using the Ids just written into column A.

export const RCF_ITEM_HEADER = "rcf item";

const normalizeHeader = (value) =>
  String(value ?? "").replace(/(\r\n|\n|\r)/g, " ").replace(/\s+/g, " ").trim().toLowerCase();

export const isRcfItemHeader = (header) => normalizeHeader(header) === RCF_ITEM_HEADER;

// An Excel checkbox cell comes back as true/false; a typed cell as text. Yes/Y/True/1/X/✓ count as
// ticked, everything else (blank included) as not ticked.
export function toCheckbox(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value === 1;
  return ["yes", "y", "true", "1", "x", "✓", "✔"].includes(String(value ?? "").trim().toLowerCase());
}

// Same cost-sheet layout the full push and escalationFetcher read (header on row 5).
const COST_HEADER_ROW = 5;
const COMPOSITE_BATCH_SIZE = 25;

// Reads the RCF Item column for the given cost-sheet row ranges ({ sheetName, rangeAddress:
// "start:end" }) and sets RCF_Item__c = true on every cost line whose box is ticked. Unticked rows
// are left alone: a new QuoteCostItem__c is already false. Returns { attempted, failures }.
export async function syncRcfItems(ranges) {
  const usable = (ranges ?? []).filter((r) => r?.sheetName && r?.rangeAddress);
  if (usable.length === 0) return { attempted: 0, failures: [] };

  const ticked = await Excel.run({ delayForCellEdit: true }, async (context) => {
    const bySheet = new Map();
    for (const { sheetName, rangeAddress } of usable) {
      const [start, end] = rangeAddress.split(":").map(Number);
      if (!bySheet.has(sheetName)) bySheet.set(sheetName, []);
      bySheet.get(sheetName).push({ start, end });
    }

    const reads = [];
    for (const [sheetName, spans] of bySheet) {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const header = sheet.getRange(`${COST_HEADER_ROW}:${COST_HEADER_ROW}`).getUsedRangeOrNullObject(true);
      header.load(["values", "columnIndex", "isNullObject"]);
      const maxRow = Math.max(...spans.map((s) => s.end));
      reads.push({ sheetName, sheet, header, spans, maxRow });
    }
    await context.sync();

    const blocks = [];
    for (const read of reads) {
      if (read.header.isNullObject) continue;
      const offset = read.header.columnIndex;
      const rcfIndex = (read.header.values[0] ?? []).findIndex(isRcfItemHeader);
      if (rcfIndex === -1) continue; // this cost sheet has no RCF Item column
      const ids = read.sheet.getRangeByIndexes(0, 0, read.maxRow, 1);
      const flags = read.sheet.getRangeByIndexes(0, offset + rcfIndex, read.maxRow, 1);
      ids.load("values");
      flags.load("values");
      blocks.push({ ...read, ids, flags });
    }
    await context.sync();

    const result = new Set();
    for (const { spans, ids, flags } of blocks) {
      for (const { start, end } of spans) {
        for (let row = start; row <= end; row++) {
          const id = String(ids.values[row - 1]?.[0] ?? "").trim();
          if (id && toCheckbox(flags.values[row - 1]?.[0])) result.add(id);
        }
      }
    }
    return [...result];
  });

  const failures = [];
  for (let i = 0; i < ticked.length; i += COMPOSITE_BATCH_SIZE) {
    const chunk = ticked.slice(i, i + COMPOSITE_BATCH_SIZE);
    const response = await SalesforceService.composite({
      allOrNone: false,
      compositeRequest: chunk.map((id, n) => ({
        method: "PATCH",
        url: `/services/data/v59.0/sobjects/QuoteCostItem__c/${id}`,
        referenceId: `rcf_${n}_${id}`,
        body: { RCF_Item__c: true },
      })),
    });
    (response?.compositeResponse ?? []).forEach((r, n) => {
      if (r.httpStatusCode !== 204 && r.httpStatusCode !== 200) {
        const body = r.body;
        failures.push({
          id: chunk[n],
          message: Array.isArray(body) ? body.map((e) => `${e.errorCode}: ${e.message}`).join("; ") : JSON.stringify(body),
        });
      }
    });
  }

  return { attempted: ticked.length, failures };
}
