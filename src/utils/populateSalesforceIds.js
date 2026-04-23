export async function populateSalesforceIds(apiResult) {
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
        console.warn(`Worksheet "${sheetName}" not found, skipping.`);
        continue;
      }

      for (const { address, value } of entries) {
        ws.getRange(address).values = [[value]];
      }
    }

    await context.sync();
  });
}