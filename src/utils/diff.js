// ─── Field Mappings ───────────────────────────────────────────────────────────

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


// ─── In-memory pending changes ────────────────────────────────────────────────

const pendingChanges = {
  quoteHeader: {}, // { sfApiName: newValue }
  quoteLineItems: {}, // { sfId: { sfApiName: newValue } }
  quoteCostItems: {}, // { sfId: { sfApiName: newValue } }
  escalations:    {},
  meta: {
    quoteExists: false,  // true if J7 has a value
    hasNewRows:  false,  // true if any data row has empty A col
  },
};

// ─── Helpers ──────────────────────────────────────────────────────────────────

export let isListenerPaused = false;

/**
 * Find the 1-based row index of a section marker in a column-A values array.
 * Uses loose matching to handle em-dash vs regular dash variations.
 */
function findSectionRow(colAValues, marker) {
  // Normalize dashes so "—" (em), "–" (en), "-" all match
  const normalize = (s) =>
    String(s).replace(/[—–-]/g, "-").replace(/\s+/g, " ").trim().toLowerCase();

  const normalizedMarker = normalize(marker);

  for (let i = 0; i < colAValues.length; i++) {
    if (normalize(colAValues[i][0]).includes(normalizedMarker)) {
      return i + 1; // 1-based
    }
  }
  return -1;
}

/**
 * Resolve an Excel column label to a Salesforce API field name.
 * @returns {string}    sfApiName  — use this field
 * @returns {null}      explicitly null-mapped — skip
 * @returns {undefined} label not in map — ignore silently
 */
function resolveField(map, excelLabel) {
  const trimmed = String(excelLabel ?? "").trim();
  if (!(trimmed in map)) return undefined;
  return map[trimmed];
}

// ─── Master sheet handler ─────────────────────────────────────────────────────

async function handleMasterSheetChange(event, context) {
  const sheet = context.workbook.worksheets.getItem(event.worksheetId);

  // Scan first 100 rows for section marker
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

  const startRow = changedRange.rowIndex + 1; // 1-based
  const startCol = changedRange.columnIndex + 1; // 1-based
  const { values } = changedRange;

  for (let r = 0; r < changedRange.rowCount; r++) {
    const absoluteRow = startRow + r; // 1-based

    // ── Quote Header region ───────────────────────────────────────────────────
    // Inside handleMasterSheetChange, replace the Quote Header region block:

    if (absoluteRow < sectionBRow) {
      for (let c = 0; c < changedRange.columnCount; c++) {
        const absoluteCol = startCol + c; // 1-based

        // Scan leftward from the changed cell until we find a non-empty cell — that's the label
        let labelCol = absoluteCol - 2; // 0-based, start one col to the left
        let label = "";

        while (labelCol >= 0) {
          const labelCell = sheet.getCell(absoluteRow - 1, labelCol); // 0-based
          labelCell.load("values");
          await context.sync();

          const cellVal = String(labelCell.values[0][0] ?? "").trim();
          if (cellVal) {
            label = cellVal;
            break; // found the label
          }
          labelCol--; // keep going left
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

    // ── Skip section marker row and line item header row ──────────────────────
    if (absoluteRow === sectionBRow || absoluteRow === lineItemHeaderRow) continue;

    // ── Quote Line Item data rows ─────────────────────────────────────────────
    if (absoluteRow > lineItemHeaderRow) {
      const headerRange = sheet.getRange(`${lineItemHeaderRow}:${lineItemHeaderRow}`);
      headerRange.load("values");
      await context.sync();
      const headers = headerRange.values[0];

      // Salesforce ID in column R (0-based index 17)
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
        const absoluteCol = startCol + c; // 1-based
        const headerIndex = absoluteCol - 1; // 0-based into headers array
        const label = String(headers[headerIndex] ?? "").trim();
        const sfField = resolveField(QUOTE_LINE_ITEM_MAP, label);
        if (!sfField) continue;

        pendingChanges.quoteLineItems[sfId][sfField] = values[r][c];
        console.log(`  Quote Line Item field mapped: "${label}" → ${sfField} = ${values[r][c]}`);
      }
    }
  }
}

// ─── Cost item sheet handler ──────────────────────────────────────────────────

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

  const startRow = changedRange.rowIndex + 1; // 1-based
  const startCol = changedRange.columnIndex + 1;
  const { values } = changedRange;

  for (let r = 0; r < changedRange.rowCount; r++) {
    const absoluteRow = startRow + r;
    if (absoluteRow <= HEADER_ROW) continue;

    // Salesforce ID in column R (0-based index 17)
    const sfIdCell = sheet.getCell(absoluteRow - 1, 0);
    // ParentQuoteCostLine__c ID in column A (0-based index 0)
    const parentIdCell = sheet.getCell(absoluteRow - 1, 0);

    sfIdCell.load("values");
    parentIdCell.load("values");
    await context.sync();

    const sfId    = String(sfIdCell.values[0][0]    ?? "").trim();
    const parentId = String(parentIdCell.values[0][0] ?? "").trim();

    for (let c = 0; c < changedRange.columnCount; c++) {
      const absoluteCol = startCol + c; // 1-based
      const headerIndex = absoluteCol - 1; // 0-based
      const label       = String(headers[headerIndex] ?? "").trim();
      const cellValue   = values[r][c];

      // ── Check if it's an escalation column ──────────────────────────────
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

      // ── Regular cost item field ──────────────────────────────────────────
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

// ─── Payload builder ──────────────────────────────────────────────────────────

export async function buildPayload() {
  let headerId = null;

  await Excel.run(async (context) => {
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

  const quoteHeader = {
          objectApiName: OBJECT_API_NAMES.QuoteHeader,
          id:            headerId,
          fields:        { ...pendingChanges.quoteHeader, UniqueText__c: `ExcelSync-${Date.now()}`, isEditable__c: false },
        };

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

  // Escalations — resolve actual SF IDs at push time via DiffSyncService
  // Here we just pass the raw structure; DiffSyncService queries SF for the IDs
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

// ─── Clear state after a successful push ─────────────────────────────────────

// AFTER
export function clearPendingChanges() {
  pendingChanges.quoteHeader    = {};
  pendingChanges.quoteLineItems = {};
  pendingChanges.quoteCostItems = {};
  pendingChanges.escalations    = {};
  pendingChanges.meta.hasNewRows = false;
}

// ─── Main listener ────────────────────────────────────────────────────────────

export async function handleSheetChange(event) {
  if (isListenerPaused) return;
  if (event.changeType !== "RangeEdited") return;
  console.log(`Change detected on sheet ${event.worksheetId}, address ${event.address}, type ${event.changeType}`);

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(event.worksheetId);
      sheet.load("name");
      await context.sync();

      const isMaster = sheet.name === "Final Quote Header & Lines"; // ← adjust to your master sheet tab name

      if (isMaster) {
        await handleMasterSheetChange(event, context);
      } else {
        await handleCostItemSheetChange(event, context);
      }
    });

    // buildPayload() called OUTSIDE Excel.run() — opens its own context internally
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
