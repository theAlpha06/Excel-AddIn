import SalesforceService from "../services/SalesforceService";

export async function fetchEscalationData(quantityRanges) {
    console.log("Fetching escalation data for quantity ranges:", quantityRanges);
  return Excel.run({ delayForCellEdit: true }, async (context) => {
    const workbook = context.workbook;
    const sheetCache = new Map();

    for (const { sheetName } of quantityRanges) {
      if (!sheetCache.has(sheetName)) {
        const sheet = workbook.worksheets.getItem(sheetName);
        const usedRange = sheet.getUsedRange();

        usedRange.load(["values", "rowCount", "columnCount"]);

        sheetCache.set(sheetName, {
          sheet,
          usedRange,
        });
      }
    }

    await context.sync();

    const result = [];

    for (const { sheetName, rangeAddress } of quantityRanges) {
      const { usedRange } = sheetCache.get(sheetName);

      const values = usedRange.values;
      console.log(`Processing sheet: ${sheetName}, range: ${rangeAddress}, values:`, values);

      if (!values.length) continue;

      const headers = values[4];

      const escalationColumns = [];

      headers.forEach((header, index) => {
        console.log(`Checking header: ${header} at index ${index}`);
        if (
          typeof header === "string" &&
          (
            header.includes("Sale Escalation (") ||
            header.includes("Cost Escalation (")
          )
        ) {
          escalationColumns.push({
            index,
            header,
          });
        }
      });

      const [startRow, endRow] = rangeAddress
        .split(":")
        .map(Number);

      for (let excelRow = startRow; excelRow <= endRow; excelRow++) {
        const row = values[excelRow - 1];

        if (!row) continue;

        const escalationValues = {};

        escalationColumns.forEach((col) => {
          escalationValues[col.header] = row[col.index];
        });

        result.push({
          sheetName,
          rowNumber: excelRow,
          salesforceId: row[0], 
          escalations: escalationValues,
        });
      }
    }

    await updateEscalations(result);
  });
}
export async function updateEscalations(escalationData) {
  const parentIds = [
    ...new Set(
      escalationData
        .map((r) => r.salesforceId)
        .filter((id) => id && id.trim() !== "")
    ),
  ];

  if (parentIds.length === 0) return;

  const CHUNK_SIZE = 100;
  const allRecords = [];

  for (let i = 0; i < parentIds.length; i += CHUNK_SIZE) {
    const chunk = parentIds.slice(i, i + CHUNK_SIZE);

    const soql = `
      SELECT Id,
             ParentQuoteCostLine__c,
             YearNumber__c
      FROM QuoteCostItem__c
      WHERE ParentQuoteCostLine__c IN ('${chunk.join("','")}')
    `;

    const queryResult = await SalesforceService.query(soql);

    allRecords.push(...queryResult.records);
  }

  const yearMap = new Map();

  for (const record of allRecords) {
    if (!yearMap.has(record.ParentQuoteCostLine__c)) {
      yearMap.set(record.ParentQuoteCostLine__c, new Map());
    }

    yearMap
      .get(record.ParentQuoteCostLine__c)
      .set(Number(record.YearNumber__c), record.Id);
  }

  const updateMap = new Map();

  for (const row of escalationData) {
    if (!row.salesforceId) continue;

    const yearLookup = yearMap.get(row.salesforceId);

    if (!yearLookup) continue;

    for (const [header, value] of Object.entries(row.escalations)) {
      if (value === "" || value == null) continue;

      const match = header.match(
        /(Sale|Cost) Escalation \((\d+)(?:st|nd|rd|th) Year\)/
      );

      if (!match) continue;

      const type = match[1];
      const year = Number(match[2]);

      const recordId = yearLookup.get(year);

      if (!recordId) continue;

      if (!updateMap.has(recordId)) {
        updateMap.set(recordId, {
          Id: recordId,
          UniqueText__c: `${recordId}-${Date.now()}`
        });
      }

      const updateRecord = updateMap.get(recordId);

      if (type === "Sale") {
        updateRecord.SaleEscalation__c = Number(value);
      } else {
        updateRecord.CostEscalation__c = Number(value);
      }
    }
  }

  const updates = [...updateMap.values()];


  if (!updates.length) {
    console.log("No escalation records to update.");
    return;
  }

  await SalesforceService.updateEscalationRecords(updates);
}