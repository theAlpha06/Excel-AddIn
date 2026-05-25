import { pauseListener, resumeListener } from "./diff";

export async function populateSalesforceIds(apiResult) {
  pauseListener();
  try {
    await Excel.run(async (context) => {
      const results = apiResult.results.filter(
        (r) => r.status === "created" && r.cell && r.salesforceId
      );

      const sheetMap = new Map();
      for (const record of results) {
        const [sheetName, rowStr] = record.cell.split("!");
        if (!sheetName || !rowStr) continue;
        const cellAddress = `R${rowStr}`; 
        if (!sheetMap.has(sheetName)) sheetMap.set(sheetName, []);
        sheetMap.get(sheetName).push({ address: cellAddress, value: record.salesforceId });
      }

      const worksheets = context.workbook.worksheets;
      worksheets.load("items/name");
      await context.sync();

      const sheetByName = new Map(
        worksheets.items.map((ws) => [ws.name, ws])
      );

      for (const [sheetName, entries] of sheetMap.entries()) {
        const ws = sheetByName.get(sheetName);
        if (!ws) {
          continue;
        }

        for (const { address, value } of entries) {
          ws.getRange(address).values = [[value]];
        }
      }

      const quoteId = apiResult.insertedIds?.quoteIds?.[0];
      if (quoteId) {
        const quoteSheet = sheetByName.get("Final Quote Header & Lines");
        if (quoteSheet) {
          console.log(`Writing Quote ID ${quoteId} → Final Quote Header & Lines!J7`);
          quoteSheet.getRange("J7").values = [[quoteId]];
        } else {
          console.warn(`Sheet not found: "Final Quote Header & Lines"`);
        }
      }
      await context.sync();
    });
  } finally {
    resumeListener();
  }
}