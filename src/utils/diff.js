import { isKamLabel } from "./kamLookup";

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

// The cost sheet's header row is located from the "SECTION C" marker; row 5 is the fallback
// the rest of the codebase (escalationFetcher, SalesforceSync) assumes.
const COST_HEADER_ROW_FALLBACK = 5;

// The per-year split is driven by the "Total Quantity" column. It is the 4th column today,
// which is the fallback, but it is looked up by header text first so the distribution keeps
// working if the column ever moves.
const TOTAL_QUANTITY_HEADER = "Total Quantity";
const TOTAL_QUANTITY_COLUMN_INDEX = 3;

// The tenure that the total quantity is spread over.
const QUOTE_TENURE_REF = "'Final Quote Header & Lines'!J15";

// Excel headers routinely carry line breaks and padding, so every header comparison runs
// against a normalised copy of the text.
const normalizeHeaderText = (value) =>
  String(value ?? "")
    .replace(/(\r\n|\n|\r)/g, " ")
    .replace(/\s+/g, " ")
    .trim();

// Year columns are headed by the year number itself, which Excel returns as a number or as
// text depending on the cell format. Number("") and Number(" ") are both 0 — an integer — so
// blank headers have to be rejected before the numeric test, or every empty column past the
// data would be treated as a year column and have a formula written into it.
const parseYearHeader = (value) => {
  const normalized = normalizeHeaderText(value);
  if (normalized === "") return null;
  const year = Number(normalized);
  return Number.isInteger(year) && year >= 1 ? year : null;
};

const columnIndexToLetter = (index) => {
  let letter = "";
  for (let n = index; n >= 0; n = Math.floor(n / 26) - 1) {
    letter = String.fromCharCode((n % 26) + 65) + letter;
  }
  return letter;
};

// Collapses a sorted list of indexes into { start, length } runs of adjacent values so a
// block of rows or of year columns can be written with one range assignment.
const contiguousRuns = (sortedValues) => {
  const runs = [];
  for (const value of sortedValues) {
    const last = runs[runs.length - 1];
    if (last && value === last.start + last.length) last.length++;
    else runs.push({ start: value, length: 1 });
  }
  return runs;
};

// Same distribution as before: years past the tenure stay blank, the final year absorbs the
// remainder, every earlier year gets the floored equal share. Only the header row and the
// total-quantity column are parameterised so a relocated layout cannot silently break it.
const buildYearQuantityFormula = (headerRow, quantityColumnLetter) => {
  const yearOfColumn = `INDEX(${headerRow}:${headerRow},COLUMN())`;
  const total = `INDEX(${quantityColumnLetter}:${quantityColumnLetter},ROW())`;
  const perYear = `FLOOR(${total}/${QUOTE_TENURE_REF},1)`;
  const lastYear = `${total}-${perYear}*(${QUOTE_TENURE_REF}-1)`;
  return (
    `=IF(${yearOfColumn}>${QUOTE_TENURE_REF},"",` +
    `IF(${yearOfColumn}=${QUOTE_TENURE_REF},${lastYear},${perYear}))`
  );
};

const stripSheetName = (address) => {
  const text = String(address ?? "").replace(/\$/g, "").trim();
  const bang = text.lastIndexOf("!");
  return bang === -1 ? text : text.slice(bang + 1).trim();
};

// Excel reports a non-contiguous edit as a comma separated list of areas ("D6:D9,D14"), and
// Worksheet.getRange() cannot load one of those — it throws, which used to abort the whole
// handler and drop the edit. Each area is read on its own instead.
const splitChangeAddress = (address) =>
  String(address ?? "")
    .split(",")
    .map(stripSheetName)
    .filter(Boolean);

// Largest row number named by a set of addresses. Deleting the contents of the bottom rows
// shrinks the used range, so the one bulk read of column A is sized off the edit as well.
const maxAddressRow = (addresses) => {
  const pattern = /[A-Za-z]+(\d+)/g;
  let max = 0;
  for (const address of addresses) {
    let match;
    while ((match = pattern.exec(address)) !== null) {
      max = Math.max(max, Number(match[1]));
    }
  }
  return max;
};

// Excel raises onChanged for the add-in's own writes too, so distributing the quantity
// formulas re-enters this handler, and each re-entry costs a full round trip to Excel before
// it concludes there is nothing to record. Hosts that report triggerSource are filtered on
// that; on the rest, the blocks written below are remembered and matched by address.
const SELF_WRITE_TTL_MS = 5000;
const selfWrites = [];

function noteSelfWrite(worksheetId, address) {
  selfWrites.push({ key: `${worksheetId}!${stripSheetName(address)}`, at: Date.now() });
}

function isSelfTriggered(event) {
  if (String(event.triggerSource ?? "").toLowerCase() === "thislocaladdin") return true;

  // Expire stale entries first: a write that never produced an event must not swallow a
  // later edit at the same address.
  const now = Date.now();
  for (let i = selfWrites.length - 1; i >= 0; i--) {
    if (now - selfWrites[i].at > SELF_WRITE_TTL_MS) selfWrites.splice(i, 1);
  }

  const key = `${event.worksheetId}!${stripSheetName(event.address)}`;
  const index = selfWrites.findIndex((w) => w.key === key);
  if (index === -1) return false;
  selfWrites.splice(index, 1);
  return true;
}



const pendingChanges = {
  quoteHeader: {},
  quoteLineItems: {},
  quoteCostItems: {},
  escalations:    {},
  // Cost sheet name → Set of Excel rows whose Total Quantity or year (1, 2, 3, …) columns were
  // edited. Those values have no field mapping of their own — the year split is written by this
  // add-in as formulas — so the rows are remembered here and their year quantities are read and
  // sent at push time, the same way the full push does it (fetchEscalationData).
  costQuantityRows: {},
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

        // KAM - Account holds the KAM's email, but Quote.KAM__c is a User lookup. The email is
        // queued as-is; DiffSyncService resolves it to the User Id when the push runs.
        if (isKamLabel(label)) {
          pendingChanges.quoteHeader.KAM__c = String(values[r][c] ?? "").trim();
          console.log(`  Quote Header: "${label}" → KAM__c (email ${values[r][c]}, resolved at push)`);
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

  const areaAddresses = splitChangeAddress(event.address);

  // Round trip 1: the section marker that locates the header row, and the used range that
  // bounds how far the sheet's rows and columns actually go.
  const colA = sheet.getRange("A1:A20");
  const usedRange = sheet.getUsedRangeOrNullObject();
  sheet.load("name");
  colA.load("values");
  usedRange.load(["isNullObject", "rowIndex", "rowCount", "columnIndex", "columnCount"]);
  await context.sync();

  const sectionCRow = findSectionRow(colA.values, "SECTION C - QUOTE COST ITEM");
  const HEADER_ROW  = sectionCRow !== -1 ? sectionCRow + 1 : COST_HEADER_ROW_FALLBACK;

  // getUsedRange().columnCount is a width, not a last index: content that does not start in
  // column A needs the offset added back, or the trailing year columns are never even read
  // and their formulas never written. Same correction as escalationFetcher.js.
  const usedWidth = usedRange.isNullObject ? 0 : usedRange.columnIndex + usedRange.columnCount;
  const usedHeight = usedRange.isNullObject ? 0 : usedRange.rowIndex + usedRange.rowCount;

  const sheetWidth = Math.max(usedWidth, TOTAL_QUANTITY_COLUMN_INDEX + 1);
  const sheetHeight = Math.max(usedHeight, maxAddressRow(areaAddresses), HEADER_ROW);

  // Round trip 2: the header row, column A for every row at once, and each changed area.
  // The header row used to be read as a whole worksheet row (16,384 cells) twice per event,
  // and column A one cell — one sync — per changed row, which is what made a multi-row edit
  // crawl.
  const headerRange = sheet.getRangeByIndexes(HEADER_ROW - 1, 0, 1, sheetWidth);
  headerRange.load("values");

  // Read column A from row 1 so an array index is the Excel row number.
  const idColumn = sheet.getRangeByIndexes(0, 0, sheetHeight, 1);
  idColumn.load("values");

  const areas = areaAddresses.map((address) => {
    const range = sheet.getRange(address);
    range.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    return range;
  });

  await context.sync();

  const headers = headerRange.values[0] ?? [];
  const idValues = idColumn.values ?? [];

  const wantedQuantityHeader = TOTAL_QUANTITY_HEADER.toLowerCase();
  const namedQuantityColumn = headers.findIndex(
    (header) => normalizeHeaderText(header).toLowerCase() === wantedQuantityHeader
  );
  const quantityColumnIndex =
    namedQuantityColumn === -1 ? TOTAL_QUANTITY_COLUMN_INDEX : namedQuantityColumn;

  const yearColumnIndexes = [];
  headers.forEach((header, index) => {
    if (parseYearHeader(header) !== null) yearColumnIndexes.push(index);
  });

  // Every row whose Total Quantity cell was touched gets its per-year split rewritten. The
  // old check only fired when the edited range *started* on that column, so an edit that
  // merely spanned it — a paste across C:E, a two-column fill, a multi-cell selection — left
  // the year columns untouched. That is the "sometimes it does not work".
  const rowsToDistribute = new Set();
  for (const area of areas) {
    const firstCol = area.columnIndex;
    const lastCol  = area.columnIndex + area.columnCount - 1;
    if (quantityColumnIndex < firstCol || quantityColumnIndex > lastCol) continue;

    for (let r = 0; r < area.rowCount; r++) {
      const rowNumber = area.rowIndex + 1 + r;
      // Never write into the header row or above it: that would overwrite the year headers,
      // and since the formula reads that row it would also make every cell self-referential.
      if (rowNumber > HEADER_ROW) rowsToDistribute.add(rowNumber);
    }
  }

  // Rows whose quantities changed — through Total Quantity (which rewrites the year split) or by
  // typing into a year column directly — are queued for the diff push. Only rows that already
  // exist in Salesforce (an Id in column A) can be updated; a new row goes through the full push.
  const quantityColumnSet = new Set([quantityColumnIndex, ...yearColumnIndexes]);
  for (const area of areas) {
    const firstCol = area.columnIndex;
    const lastCol  = area.columnIndex + area.columnCount - 1;
    let touchesQuantity = false;
    for (const index of quantityColumnSet) {
      if (index >= firstCol && index <= lastCol) { touchesQuantity = true; break; }
    }
    if (!touchesQuantity) continue;

    for (let r = 0; r < area.rowCount; r++) {
      const rowNumber = area.rowIndex + 1 + r;
      if (rowNumber <= HEADER_ROW) continue;
      if (!String(idValues[rowNumber - 1]?.[0] ?? "").trim()) continue;
      if (!pendingChanges.costQuantityRows[sheet.name]) {
        pendingChanges.costQuantityRows[sheet.name] = new Set();
      }
      pendingChanges.costQuantityRows[sheet.name].add(rowNumber);
    }
  }

  if (rowsToDistribute.size > 0 && yearColumnIndexes.length > 0) {
    const formula = buildYearQuantityFormula(
      HEADER_ROW,
      columnIndexToLetter(quantityColumnIndex)
    );
    const rowRuns = contiguousRuns([...rowsToDistribute].sort((a, b) => a - b));
    const columnRuns = contiguousRuns(yearColumnIndexes);
    const written = [];

    // One assignment per rectangular block instead of one per cell: a 20-row × 5-year edit
    // was 100 separate writes, each of which Excel also reported back as its own change
    // event that re-entered this handler.
    for (const rowRun of rowRuns) {
      for (const columnRun of columnRuns) {
        const block = sheet.getRangeByIndexes(
          rowRun.start - 1,
          columnRun.start,
          rowRun.length,
          columnRun.length
        );
        block.formulas = Array.from({ length: rowRun.length }, () =>
          Array.from({ length: columnRun.length }, () => formula)
        );
        block.load("address");
        written.push(block);
      }
    }

    await context.sync();

    for (const block of written) noteSelfWrite(event.worksheetId, block.address);
  }

  for (const area of areas) {
    const startRow = area.rowIndex + 1;
    const startCol = area.columnIndex + 1;
    const { values } = area;

    for (let r = 0; r < area.rowCount; r++) {
      const absoluteRow = startRow + r;
      if (absoluteRow <= HEADER_ROW) continue;

      // Column A holds the cost line's Salesforce Id, which is both the record to update and
      // the ParentQuoteCostLine__c of its year lines.
      const sfId     = String(idValues[absoluteRow - 1]?.[0] ?? "").trim();
      const parentId = sfId;

      for (let c = 0; c < area.columnCount; c++) {
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

  // One "start:end" range per contiguous block of rows, the shape fetchEscalationData reads.
  const costQuantityRanges = Object.entries(pendingChanges.costQuantityRows).flatMap(
    ([sheetName, rows]) =>
      contiguousRuns([...rows].sort((a, b) => a - b)).map(({ start, length }) => ({
        sheetName,
        rangeAddress: `${start}:${start + length - 1}`,
      }))
  );

  return {
    quoteHeader,
    quoteLineItems,
    quoteCostItems,
    escalations,
    costQuantityRanges,
    meta: { ...pendingChanges.meta },
  };
}

export function clearPendingChanges() {
  pendingChanges.quoteHeader    = {};
  pendingChanges.quoteLineItems = {};
  pendingChanges.quoteCostItems = {};
  pendingChanges.escalations    = {};
  pendingChanges.costQuantityRows = {};
  pendingChanges.meta.hasNewRows = false;
}

// Excel delivers change events concurrently, and every one of them mutates the same
// pendingChanges object and writes back to the sheet. Running them one at a time keeps a
// later event from reading the sheet halfway through an earlier one's write, and lets the
// payload be built once per burst instead of once per changed cell.
let changeQueue = Promise.resolve();
let queuedChanges = 0;

export async function handleSheetChange(event) {
  if (isListenerPaused) return;
  if (event.changeType !== "RangeEdited") return;
  // Changes this add-in made itself carry no user intent — processing them only costs round
  // trips and, when a write lands mid-handler, produces stale reads.
  if (isSelfTriggered(event)) return;

  console.log("Detected change event:", event);
  console.log(`Change detected on sheet ${event.worksheetId}, address ${event.address}, type ${event.changeType}, name ${event.worksheetName}`);

  queuedChanges++;
  changeQueue = changeQueue.then(() => processSheetChange(event));
  return changeQueue;
}

async function processSheetChange(event) {
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
  } catch (err) {
    console.error("Change handler error:", err);
  } finally {
    queuedChanges--;
  }

  // A burst of events — a paste, a fill, a multi-cell delete — all accumulate into the same
  // pendingChanges, so only the last one needs to publish it. Skipping the intermediate
  // builds drops an Excel round trip and a full payload stringify per changed cell.
  if (queuedChanges > 0) return;

  try {
    const payload = await buildPayload();
    if (_setPayload) _setPayload(payload);
    console.log("📦 Payload to push to Salesforce:", JSON.stringify(payload, null, 2));
  } catch (err) {
    console.error("Payload build error:", err);
  }
}

export const pauseListener = () => {
  isListenerPaused = true;
};
export const resumeListener = () => {
  isListenerPaused = false;
};
