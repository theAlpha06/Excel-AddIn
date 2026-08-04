import SalesforceService from "../services/SalesforceService";
import { pauseListener, resumeListener } from "./diff";

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

      // One bounding range per sheet, loaded together in a single sync
      const sheetRangeInfo = [];
      for (const [sheetName, rowValueMap] of sheetMap.entries()) {
        const rows = [...rowValueMap.keys()];
        const minRow = Math.min(...rows);
        const maxRow = Math.max(...rows);
        const ws = worksheets.getItemOrNullObject(sheetName);
        const range = ws.getRange(`A${minRow}:A${maxRow}`);
        range.load("values");
        sheetRangeInfo.push({ ws, minRow, range, rowValueMap });
      }

      const quoteSheet = quoteId
        ? worksheets.getItemOrNullObject("Final Quote Header & Lines")
        : null;

      await context.sync(); // <-- single round-trip: reads every bounding range at once

      // Patch in-memory, then queue one write per sheet
      for (const { ws, minRow, range, rowValueMap } of sheetRangeInfo) {
        if (ws.isNullObject) continue;
        const existing = range.values; // 2D array, single column
        for (const [row, value] of rowValueMap.entries()) {
          existing[row - minRow][0] = value;
        }
        range.values = existing;
      }

      if (quoteSheet && !quoteSheet.isNullObject) {
        quoteSheet.getRange("J7").values = [[quoteId]];
        quoteSheet.getRange("J14").values = [[quoteNumberValue]];
      }

      await context.sync(); // <-- single round-trip: writes every sheet at once
    });
  } finally {
    resumeListener();
  }
}