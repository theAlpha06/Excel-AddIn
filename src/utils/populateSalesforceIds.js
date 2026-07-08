import SalesforceService from "../services/SalesforceService";
import { pauseListener, resumeListener } from "./diff";

export async function populateSalesforceIds(apiResult) {
  pauseListener();
  try {
      const results = apiResult.results.filter(
        (r) => r.status === "created" && r.cell && r.salesforceId
      );

      const quoteId = apiResult.insertedIds?.quoteIds?.[0];
      const quoteNumber = await SalesforceService.query(`SELECT QuoteNumber FROM Quote WHERE Id = '${quoteId}'`);

    await Excel.run({ delayForCellEdit: true }, async (context) => {
      const sheetMap = new Map();
      for (const record of results) {
        const [sheetName, rowStr] = record.cell.split("!");
        if (!sheetName || !rowStr) continue;
        const cellAddress = `A${rowStr}`; 
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
      
      if (quoteId) {
        const quoteSheet = sheetByName.get("Final Quote Header & Lines");
        if (quoteSheet) {
          quoteSheet.getRange("J7").values = [[quoteId]];
          quoteSheet.getRange("J14").values = [[quoteNumber?.records?.[0]?.QuoteNumber || ""]];
        } 
      }
      await context.sync();
    });
  } finally {
    resumeListener();
  }
}