import SalesforceService from "../services/SalesforceService";
import { pauseListener, resumeListener } from "./diff";

// How many rows one read/write request may span. Bounded so a quote with thousands of cost lines
// never hangs its Ids on a single oversized Excel request.
const ROWS_PER_BLOCK = 2000;

export async function populateSalesforceIds(apiResult) {
  pauseListener();
  try {
    const results = apiResult.results.filter(
      (r) => r.status === "created" && r.cell && r.salesforceId
    );

    const quoteId = apiResult.insertedIds?.quoteIds?.[0];
    const quoteNumberPromise = quoteId
      ? SalesforceService.query(`SELECT QuoteNumber FROM Quote WHERE Id = '${quoteId}'`)
      : Promise.resolve(null);

    // Group writes by sheet: sheetName -> Map<rowNumber, salesforceId>
    const sheetMap = new Map();
    for (const record of results) {
      const [sheetName, rowStr] = record.cell.split("!");
      const row = parseInt(rowStr, 10);
      if (!sheetName || !row) continue;
      if (!sheetMap.has(sheetName)) sheetMap.set(sheetName, new Map());
      sheetMap.get(sheetName).set(row, record.salesforceId);
    }

    // Kick off the Salesforce query in parallel with building the sheet map
    const quoteNumberResult = await quoteNumberPromise;
    const quoteNumberValue = quoteNumberResult?.records?.[0]?.QuoteNumber || "";

    await Excel.run({ delayForCellEdit: true }, async (context) => {
      const worksheets = context.workbook.worksheets;

      const sheetTargets = [];
      for (const [sheetName, rowValueMap] of sheetMap.entries()) {
        const ws = worksheets.getItemOrNullObject(sheetName);
        // isNullObject is only meaningful once something on the sheet itself has been loaded.
        // Without this the check below always read false, and a missing sheet failed the whole
        // batch at sync instead of being reported and skipped.
        ws.load("isNullObject");
        // Rows sorted so the write blocks below walk them in one pass. Math.min(...rows) spreads
        // one argument per row, which has a hard ceiling — a plain sort has none.
        const entries = [...rowValueMap.entries()].sort((a, b) => a[0] - b[0]);
        sheetTargets.push({ sheetName, ws, entries });
      }

      const quoteSheet = quoteId
        ? worksheets.getItemOrNullObject("Final Quote Header & Lines")
        : null;
      if (quoteSheet) quoteSheet.load("isNullObject");

      await context.sync();

      let written = 0;
      let mismatched = 0;

      // The whole A{min}:A{max} window used to be read, patched and written as one request. On a
      // ten-thousand-line quote that is a single oversized payload covering every row between the
      // first and last Id, including thousands with nothing to write. Bounded blocks that each
      // start on a row that actually has an Id keep every request small and skip the gaps.
      for (const { sheetName, ws, entries } of sheetTargets) {
        if (ws.isNullObject) {
          console.warn(`Sheet "${sheetName}" not found — ${entries.length} Id(s) not written.`);
          continue;
        }

        let cursor = 0;
        while (cursor < entries.length) {
          const blockStart = entries[cursor][0];
          const blockEnd = blockStart + ROWS_PER_BLOCK - 1;

          const blockEntries = [];
          while (cursor < entries.length && entries[cursor][0] <= blockEnd) {
            blockEntries.push(entries[cursor]);
            cursor++;
          }

          const lastRow = blockEntries[blockEntries.length - 1][0];
          const range = ws.getRange(`A${blockStart}:A${lastRow}`);
          range.load("values");
          await context.sync();

          const existing = range.values; // 2D array, single column
          for (const [row, value] of blockEntries) {
            existing[row - blockStart][0] = value;
          }
          range.values = existing;

          // Read the block back rather than trusting a sync that reported no error. A silently
          // short write is the one failure mode that would otherwise look like "the Ids just did
          // not appear", with nothing in the console to say so.
          const verify = ws.getRange(`A${blockStart}:A${lastRow}`);
          verify.load("values");
          await context.sync();

          for (const [row, value] of blockEntries) {
            if (String(verify.values[row - blockStart][0] ?? "").trim() === String(value).trim()) {
              written++;
            } else {
              mismatched++;
              if (mismatched <= 10) {
                console.warn(`Id did not land at ${sheetName}!A${row} (expected ${value}).`);
              }
            }
          }
        }
      }

      if (quoteSheet && !quoteSheet.isNullObject) {
        quoteSheet.getRange("J7").values = [[quoteId]];
        quoteSheet.getRange("J14").values = [[quoteNumberValue]];
        await context.sync();
      }

      console.log(`Wrote ${written} Salesforce Id(s) to Excel.`);
      if (mismatched > 0) {
        console.error(`${mismatched} Salesforce Id(s) failed to write. See warnings above.`);
      }
    });
  } finally {
    resumeListener();
  }
}