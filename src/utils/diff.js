const QUOTE_HEADER_MAP = {
  "Opportunity Name *": "Name",
  "Contract Start Date": "StartDate__c",
  "Manual Contract End Date": "ManualContractEnddate__c",
  "Expiration Date": "ExpirationDate__c",
  "Billing Frequency": "BillingFrequency__c",
  "Account Name * (Also Add Account Number)": null,
  LOB: "LOB__c",
  "Sub LOB": "Sub_LOB__c",
  "Service Type": "ServiceType__c",
  "Sale Price Book": null,
  "Visit Frequency": "Visit_frequency__c",
};

const QUOTE_LINE_ITEM_MAP = {
  Part: null,
  Description: "Description",
  Quantity: "Quantity",
  "Sale Price": "UnitPrice",
  "Manual Sales Price": "ManualSalesPrice__c",
  "Start Date": "StartDate__c",
  "End Date": "EndDate__c",
};

const QUOTE_COST_ITEM_MAP = {
  Part: null,
  "Part Description": "LocalPartDescription__c",
  "Cost Price": "CostPrice__c",
  "Manual Cost Price": "ManualCostPrice__c",
  "Manual Sale Price": "ManualSalePrice__c",
  Quantity: "ManualQuantity__c",
};

const OBJECT_API_NAMES = {
  QuoteHeader: "Quote",
  QuoteLineItem: "QuoteLineItem",
  QuoteCostItem: "QuoteCostItem__c",
};

let _setPayload = null;
export function registerPayloadSetter(fn) {
  _setPayload = fn;
}
const SALE_ESCALATION_PATTERN = /sale escalation \((\d+)(?:st|nd|rd|th) year\)/i;
const COST_ESCALATION_PATTERN = /cost escalation \((\d+)(?:st|nd|rd|th) year\)/i;



const pendingChanges = {
  quoteHeader: {},
  quoteLineItems: {},
  quoteCostItems: {},
  escalations:    {},
  meta: {
    quoteExists: false, 
    hasNewRows:  false, 
  },
};


export let isListenerPaused = false;

function findSectionRow(colAValues, marker) {
  const normalize = (s) =>
    String(s).replace(/[—–-]/g, "-").replace(/\s+/g, " ").trim().toLowerCase();

  const normalizedMarker = normalize(marker);

  for (let i = 0; i < colAValues.length; i++) {
    if (normalize(colAValues[i][0]).includes(normalizedMarker)) {
      return i + 1; 
    }
  }
  return -1;
}

function resolveField(map, excelLabel) {
  const trimmed = String(excelLabel ?? "").trim();
  if (!(trimmed in map)) return undefined;
  return map[trimmed];
}


async function handleMasterSheetChange(event, context) {
  const sheet = context.workbook.worksheets.getItem(event.worksheetId);
  const colA = sheet.getRange("A1:A100");
  colA.load("values");
  await context.sync();
  const colAVals = colA.values;

  const sectionBRow = findSectionRow(colAVals, "SECTION B - QUOTE LINE ITEMS");
  if (sectionBRow === -1) {
    console.warn("Could not find SECTION B marker on master sheet");
    return;
  }

  const lineItemHeaderRow = sectionBRow + 1;

  const changedRange = sheet.getRange(event.address);
  changedRange.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
  await context.sync();

  const startRow = changedRange.rowIndex + 1; 
  const startCol = changedRange.columnIndex + 1; 
  const { values } = changedRange;

  for (let r = 0; r < changedRange.rowCount; r++) {
    const absoluteRow = startRow + r; 

    if (absoluteRow < sectionBRow) {
      for (let c = 0; c < changedRange.columnCount; c++) {
        const absoluteCol = startCol + c; 

        let labelCol = absoluteCol - 2; 
        let label = "";

        while (labelCol >= 0) {
          const labelCell = sheet.getCell(absoluteRow - 1, labelCol); 
          labelCell.load("values");
          await context.sync();

          const cellVal = String(labelCell.values[0][0] ?? "").trim();
          if (cellVal) {
            label = cellVal;
            break; 
          }
          labelCol--; 
        }

        console.log(`Quote Header — row ${absoluteRow}, label: "${label}", value: ${values[r][c]}`);

        if (!label) {
          console.warn(
            `No label found to the left of row ${absoluteRow}, col ${absoluteCol} — skipping`
          );
          continue;
        }

        const sfField = resolveField(QUOTE_HEADER_MAP, label);
        if (sfField === undefined) {
          console.warn(`Label "${label}" not in QUOTE_HEADER_MAP — skipping`);
          continue;
        }
        if (sfField === null) {
          console.log(`Label "${label}" is read-only — skipping`);
          continue;
        }

        pendingChanges.quoteHeader[sfField] = values[r][c];
        console.log(`  Quote Header: "${label}" → ${sfField} = ${values[r][c]}`);
      }
      continue;
    }

    if (absoluteRow === sectionBRow || absoluteRow === lineItemHeaderRow) continue;

    if (absoluteRow > lineItemHeaderRow) {
      const headerRange = sheet.getRange(`${lineItemHeaderRow}:${lineItemHeaderRow}`);
      headerRange.load("values");
      await context.sync();
      const headers = headerRange.values[0];

      const sfIdCell = sheet.getCell(absoluteRow - 1, 0);
      sfIdCell.load("values");
      await context.sync();
      const sfId = String(sfIdCell.values[0][0] ?? "").trim();

      if (!sfId) {
        pendingChanges.meta.hasNewRows = true;
        console.warn(`No Salesforce ID at row ${absoluteRow}, col R — skipping`);
        continue;
      }

      if (!pendingChanges.quoteLineItems[sfId]) {
        pendingChanges.quoteLineItems[sfId] = {};
      }

      for (let c = 0; c < changedRange.columnCount; c++) {
        const absoluteCol = startCol + c; 
        const headerIndex = absoluteCol - 1; 
        const label = String(headers[headerIndex] ?? "").trim();
        const sfField = resolveField(QUOTE_LINE_ITEM_MAP, label);
        if (!sfField) continue;

        pendingChanges.quoteLineItems[sfId][sfField] = values[r][c];
        console.log(`  Quote Line Item field mapped: "${label}" → ${sfField} = ${values[r][c]}`);
      }
    }
  }
}


async function handleCostItemSheetChange(event, context) {
  const sheet = context.workbook.worksheets.getItem(event.worksheetId);

  const colA = sheet.getRange("A1:A20");
  colA.load("values");
  await context.sync();
  const colAVals = colA.values;

  const sectionCRow = findSectionRow(colAVals, "SECTION C - QUOTE COST ITEM");
  const HEADER_ROW  = sectionCRow !== -1 ? sectionCRow + 1 : 5;

  const headerRange = sheet.getRange(`${HEADER_ROW}:${HEADER_ROW}`);
  headerRange.load("values");
  await context.sync();
  const headers = headerRange.values[0];

  const changedRange = sheet.getRange(event.address);
  changedRange.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
  await context.sync();

  const startRow = changedRange.rowIndex + 1; 
  const startCol = changedRange.columnIndex + 1;
  const { values } = changedRange;

if (startCol === 4) { 
  const usedRange = sheet.getUsedRange();
  usedRange.load("columnCount");
  await context.sync();

  const headerRange = sheet.getRangeByIndexes(4, 0, 1, usedRange.columnCount);
  headerRange.load("values");
  await context.sync();

  const headers = headerRange.values[0];

    for (let r = 0; r < changedRange.rowCount; r++) {
      const rowNumber = startRow + r;

      for (let c = 0; c < headers.length; c++) {
        const header = headers[c];

        if (
            header === "" ||
            header === null ||
            header === undefined ||
            !Number.isInteger(Number(header))
        ) {
            continue;
        }

        sheet.getCell(rowNumber - 1, c).formulas = [[
          `=IF(INDEX(5:5,COLUMN())>'Final Quote Header & Lines'!J15,"",IF(INDEX(5:5,COLUMN())='Final Quote Header & Lines'!J15,INDEX(D:D,ROW())-FLOOR(INDEX(D:D,ROW())/'Final Quote Header & Lines'!J15,1)*('Final Quote Header & Lines'!J15-1),FLOOR(INDEX(D:D,ROW())/'Final Quote Header & Lines'!J15,1)))`
        ]];
      }
    }

    await context.sync();
  }

  for (let r = 0; r < changedRange.rowCount; r++) {
    const absoluteRow = startRow + r;
    if (absoluteRow <= HEADER_ROW) continue;
    const sfIdCell = sheet.getCell(absoluteRow - 1, 0);
    const parentIdCell = sheet.getCell(absoluteRow - 1, 0);

    sfIdCell.load("values");
    parentIdCell.load("values");
    await context.sync();

    const sfId    = String(sfIdCell.values[0][0]    ?? "").trim();
    const parentId = String(parentIdCell.values[0][0] ?? "").trim();

    for (let c = 0; c < changedRange.columnCount; c++) {
      const absoluteCol = startCol + c; 
      const headerIndex = absoluteCol - 1; 
      const label       = String(headers[headerIndex] ?? "").trim();
      const cellValue   = values[r][c];

      const saleMatch = SALE_ESCALATION_PATTERN.exec(label);
      const costMatch = COST_ESCALATION_PATTERN.exec(label);

      if (saleMatch || costMatch) {
        const yearNumber = parseInt((saleMatch ?? costMatch)[1]);
        const fieldName  = saleMatch ? "SaleEscalation__c" : "CostEscalation__c";

        if (!parentId) {
          console.warn(`No ParentQuoteCostLine__c at row ${absoluteRow}, col A — skipping escalation`);
          continue;
        }

        if (!pendingChanges.escalations[parentId]) {
          pendingChanges.escalations[parentId] = {};
        }
        if (!pendingChanges.escalations[parentId][yearNumber]) {
          pendingChanges.escalations[parentId][yearNumber] = {};
        }

        pendingChanges.escalations[parentId][yearNumber][fieldName] = cellValue;
        console.log(`  Escalation: parentId=${parentId}, year=${yearNumber}, ${fieldName}=${cellValue}`);
        continue;
      }

      if (!sfId) {
        pendingChanges.meta.hasNewRows = true;
        console.warn(`No Salesforce ID at row ${absoluteRow}, col R — skipping`);
        continue;
      }

      const sfField = resolveField(QUOTE_COST_ITEM_MAP, label);
      if (!sfField) continue;

      if (!pendingChanges.quoteCostItems[sfId]) {
        pendingChanges.quoteCostItems[sfId] = {};
      }

      pendingChanges.quoteCostItems[sfId][sfField] = cellValue;
      console.log(`  Cost Item: "${label}" → ${sfField} = ${cellValue}`);
    }
  }
}


export async function buildPayload() {
  let headerId = null;

  await Excel.run({ delayForCellEdit: true }, async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const master = sheets.items.find((s) => s.name === "Final Quote Header & Lines");
    if (master) {
      const idCell = master.getRange("J7");
      idCell.load("values");
      await context.sync();
      headerId = String(idCell.values[0][0] ?? "").trim() || null;
      pendingChanges.meta.quoteExists = !!headerId;
    }
  });

  const quoteHeader = Object.keys(pendingChanges.quoteHeader).length > 0 ? {
          objectApiName: OBJECT_API_NAMES.QuoteHeader,
          id:            headerId,
          fields:        { ...pendingChanges.quoteHeader, UniqueText__c: `ExcelSync-${Date.now()}`, isEditable__c: false },
        }: null;

  const quoteLineItems = Object.entries(pendingChanges.quoteLineItems).map(
    ([id, fields]) => ({
      objectApiName: OBJECT_API_NAMES.QuoteLineItem,
      id,
      fields: {
        ...fields,
        UniqueText__c: `ExcelSync-${Date.now()}`,
        isEditable__c: false
      }
    })
  );

  const quoteCostItems = Object.entries(pendingChanges.quoteCostItems).map(
    ([id, fields]) => ({
      objectApiName: OBJECT_API_NAMES.QuoteCostItem,
      id,
      fields: {
        ...fields,
        UniqueText__c: `ExcelSync-${Date.now()}`,
        IsEditable__c: false
      }
    })
  );

  const escalations = Object.entries(pendingChanges.escalations).map(
    ([parentId, yearMap]) => ({
      parentId,
      years: Object.entries(yearMap).map(([year, fields]) => ({
        yearNumber: parseInt(year),
        fields: {
          ...fields,
          UniqueText__c: `ExcelSync-${Date.now()}`,
          IsEditable__c: false
        }
      }))
    })
  );

  return { quoteHeader, quoteLineItems, quoteCostItems, escalations, meta: { ...pendingChanges.meta } };
}

export function clearPendingChanges() {
  pendingChanges.quoteHeader    = {};
  pendingChanges.quoteLineItems = {};
  pendingChanges.quoteCostItems = {};
  pendingChanges.escalations    = {};
  pendingChanges.meta.hasNewRows = false;
}

export async function handleSheetChange(event) {
  if (isListenerPaused) return;
  if (event.changeType !== "RangeEdited") return;
  console.log("Detected change event:", event);
  console.log(`Change detected on sheet ${event.worksheetId}, address ${event.address}, type ${event.changeType}, name ${event.worksheetName}`);

  try {
    await Excel.run({ delayForCellEdit: true }, async (context) => {
      const sheet = context.workbook.worksheets.getItem(event.worksheetId);
      sheet.load("name");
      await context.sync();

      const isMaster = sheet.name === "Final Quote Header & Lines"; 

      if (isMaster) {
        await handleMasterSheetChange(event, context);
      } else {
        await handleCostItemSheetChange(event, context);
      }
    });

    const payload = await buildPayload();
    if (_setPayload) _setPayload(payload);
    console.log("📦 Payload to push to Salesforce:", JSON.stringify(payload, null, 2));
  } catch (err) {
    console.error("Change handler error:", err);
  }
}

export const pauseListener = () => {
  isListenerPaused = true;
};
export const resumeListener = () => {
  isListenerPaused = false;
};
