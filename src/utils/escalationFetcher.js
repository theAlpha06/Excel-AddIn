import SalesforceService from "../services/SalesforceService";

// Matches the cost-sheet layout that fetchDynamicStructuredRows reads the cost lines from.
const HEADER_ROW = 5;

// Excel headers routinely carry line breaks and padding ("Sale Escalation\n(2nd Year)"),
// so every comparison below runs against a normalised copy of the header text.
const normalizeHeader = (header) =>
  String(header ?? "")
    .replace(/(\r\n|\n|\r)/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const ESCALATION_HEADER_PATTERN =
  /^(sale|cost)\s*escalation\s*\(\s*(\d+)\s*(?:st|nd|rd|th)?\s*year\s*\)$/i;

// Single source of truth for header → (field, year). Parsing the header in one place keeps
// column detection and field mapping from drifting apart.
const parseEscalationHeader = (header) => {
  const match = ESCALATION_HEADER_PATTERN.exec(normalizeHeader(header));
  if (!match) return null;
  return {
    field: match[1].toLowerCase() === "sale" ? "SaleEscalation__c" : "CostEscalation__c",
    year: Number(match[2]),
  };
};

// Year columns are headed by the year number itself, which Excel returns as a number or as
// text depending on how the cell is formatted — hence the coercion, matching the same
// header row's handling in diff.js.
const parseYearHeader = (header) => {
  const normalized = normalizeHeader(header);
  if (normalized === "") return null;
  const year = Number(normalized);
  return Number.isInteger(year) && year >= 1 ? year : null;
};

// Salesforce needs a real number; NaN would serialise to null and silently clear the field.
const toNumber = (value) => {
  if (value === "" || value == null) return null;
  const num = typeof value === "number" ? value : Number(String(value).trim());
  return Number.isFinite(num) ? num : null;
};

const QUERY_CHUNK_SIZE = 100;

// One getRangeByIndexes(0, 0, maxRow, width) over a ten-thousand-row cost sheet asks Excel for
// hundreds of thousands of cells in a single request. Excel bundles everything queued since the
// last sync into one payload, so the rows are read — and synced — in bounded blocks instead.
const MAX_CELLS_PER_READ = 50000;

// Child year lines are written by Apex, so one that is missing right now may simply not have
// been created yet. Re-read before writing a value off as unmatched. A large quote needs longer
// than a small one before Apex has created them all, hence the doubling backoff rather than a
// flat wait.
const MAX_RESOLVE_ATTEMPTS = 5;
const RESOLVE_RETRY_DELAY_MS = 1500;

// The one-at-a-time lookup below costs a round trip per record, so it is only worth doing for a
// handful of stragglers. On a ten-thousand-line quote whose year lines Apex has not finished
// creating, nearly every target is unresolved, and running it unbounded turned into tens of
// thousands of sequential queries per year — which is the escalation step that never returns.
const MAX_SINGLE_LOOKUPS = 25;

// Read-only queries, so they can overlap. The DML further down deliberately does not.
const QUERY_CONCURRENCY = 6;

// Pretty-printing thousands of records blocks the UI thread for longer than the request it is
// describing; past this many, log the count and let the console expand the array lazily.
const MAX_LOGGED_RECORDS = 25;

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// Runs `worker` over `items` with at most `limit` in flight, preserving input order. A hundred
// chunked queries sent strictly one after another is the difference between a few seconds and
// half a minute before any DML starts.
async function mapWithConcurrency(items, limit, worker) {
  const results = new Array(items.length);
  let next = 0;

  const runners = Array.from({ length: Math.min(limit, items.length) }, async () => {
    for (;;) {
      const index = next++;
      if (index >= items.length) return;
      results[index] = await worker(items[index], index);
    }
  });

  await Promise.all(runners);
  return results;
}

// Salesforce caps a query response at 2,000 records and the backend exposes no queryMore, so a
// chunk whose year lines overflow that comes back marked not done. Splitting it and asking again
// is what keeps those year lines from going missing — a truncated read is indistinguishable from
// "this cost line has no year lines", which is exactly what sends the resolver into its
// expensive fallback path.
async function queryYearRecords(parentIds) {
  const soql = `
      SELECT Id,
             ParentQuoteCostLine__c,
             YearNumber__c
      FROM QuoteCostItem__c
      WHERE ParentQuoteCostLine__c IN ('${parentIds.join("','")}')
    `;

  const queryResult = await SalesforceService.query(soql);

  if (queryResult.done === false && parentIds.length > 1) {
    const mid = Math.ceil(parentIds.length / 2);
    console.warn(
      `Year-line query for ${parentIds.length} cost line(s) came back truncated — splitting it.`
    );
    const halves = await Promise.all([
      queryYearRecords(parentIds.slice(0, mid)),
      queryYearRecords(parentIds.slice(mid)),
    ]);
    return halves.flat();
  }

  return queryResult.records ?? [];
}

// parent cost line Id → (year number → QuoteCostItem__c Id). Year 1 is the cost line itself
// and so never appears here; the map only ever covers years 2..n.
async function fetchYearMap(parentIds) {
  const yearMap = new Map();

  const chunks = [];
  for (let i = 0; i < parentIds.length; i += QUERY_CHUNK_SIZE) {
    chunks.push(parentIds.slice(i, i + QUERY_CHUNK_SIZE));
  }

  const perChunk = await mapWithConcurrency(chunks, QUERY_CONCURRENCY, queryYearRecords);

  for (const records of perChunk) {
    for (const record of records) {
      if (!yearMap.has(record.ParentQuoteCostLine__c)) {
        yearMap.set(record.ParentQuoteCostLine__c, new Map());
      }
      yearMap.get(record.ParentQuoteCostLine__c).set(Number(record.YearNumber__c), record.Id);
    }
  }

  return yearMap;
}

// Last resort for a (parent, year) the bulk read did not return — asks for that one record
// directly so a single unmatched line cannot be blamed on the batched query.
async function fetchSingleYearId(parentId, year) {
  const soql = `
    SELECT Id
    FROM QuoteCostItem__c
    WHERE ParentQuoteCostLine__c = '${parentId}'
      AND YearNumber__c = ${year}
    LIMIT 1
  `;

  try {
    const queryResult = await SalesforceService.query(soql);
    return queryResult.records?.[0]?.Id ?? null;
  } catch (err) {
    console.warn(`Lookup failed for parent=${parentId}, year=${year}:`, err);
    return null;
  }
}

// `quantitiesOnly` sends just the year quantity columns (ManualQuantity__c) and leaves the
// escalation columns out. The diff push uses it for cost rows whose Total Quantity or year
// columns were edited: escalation edits already travel through the diff's own escalation path,
// and re-sending every escalation on those rows would change how that path behaves.
export async function fetchEscalationData(quantityRanges, { quantitiesOnly = false } = {}) {
  const ranges = (quantityRanges ?? []).filter((r) => r?.sheetName && r?.rangeAddress);
  if (ranges.length === 0) {
    console.warn("fetchEscalationData: no usable quantity ranges supplied", quantityRanges);
    return { attempted: 0, updated: 0, skipped: [], failures: [] };
  }
  console.log("Fetching escalation data for quantity ranges:", ranges);

  const escalationData = await Excel.run({ delayForCellEdit: true }, async (context) => {
    const workbook = context.workbook;
    const sheetInfo = new Map();

    for (const { sheetName, rangeAddress } of ranges) {
      const [, endRow] = rangeAddress.split(":").map(Number);
      if (!sheetInfo.has(sheetName)) {
        const sheet = workbook.worksheets.getItem(sheetName);
        const usedRange = sheet.getUsedRange();
        usedRange.load(["columnIndex", "columnCount"]);
        sheetInfo.set(sheetName, { sheet, usedRange, maxRow: HEADER_ROW });
      }
      const info = sheetInfo.get(sheetName);
      info.maxRow = Math.max(info.maxRow, endRow);
    }

    await context.sync();

    // Read each sheet from A1 so array indexes line up with real Excel rows and columns.
    // Indexing into getUsedRange().values instead shifts every lookup by that range's own
    // offset whenever a sheet's content does not begin at A1.
    //
    // The rows are read in blocks, each with its own sync, because Excel bundles everything
    // queued since the last sync into a single request: a ten-thousand-row sheet read in one go
    // is a payload big enough to fail or to sit there long enough to look like a hang. The
    // blocks are appended in order, so info.values[excelRow - 1] still addresses the real row.
    for (const info of sheetInfo.values()) {
      const width = Math.max(1, info.usedRange.columnIndex + info.usedRange.columnCount);
      const rowsPerBlock = Math.max(1, Math.floor(MAX_CELLS_PER_READ / width));
      info.values = [];

      for (let startRow = 0; startRow < info.maxRow; startRow += rowsPerBlock) {
        const height = Math.min(rowsPerBlock, info.maxRow - startRow);
        const block = info.sheet.getRangeByIndexes(startRow, 0, height, width);
        block.load("values");
        await context.sync();

        for (const row of block.values) info.values.push(row);
      }
    }

    const columnsBySheet = new Map();
    for (const [sheetName, info] of sheetInfo.entries()) {
      const headers = info.values[HEADER_ROW - 1] ?? [];
      const escalationColumns = [];
      const quantityColumns = [];

      headers.forEach((header, index) => {
        const escalation = parseEscalationHeader(header);
        if (escalation) {
          escalationColumns.push({ index, ...escalation });
          return;
        }
        const year = parseYearHeader(header);
        if (year !== null) quantityColumns.push({ index, year });
      });

      console.log(
        `"${sheetName}" row ${HEADER_ROW}: year columns [${quantityColumns
          .map((c) => c.year)
          .join(", ")}], escalation columns [${escalationColumns
          .map((c) => `${c.field}@${c.year}`)
          .join(", ")}]`
      );

      if (escalationColumns.length === 0 && quantityColumns.length === 0) {
        console.warn(
          `No escalation or year columns found on "${sheetName}" row ${HEADER_ROW}. Headers:`,
          headers.map(normalizeHeader)
        );
      }

      columnsBySheet.set(sheetName, { escalationColumns, quantityColumns });
    }

    const rows = [];
    // Quote-line windows can cover the same cost row from more than one range.
    const seenRows = new Set();

    for (const { sheetName, rangeAddress } of ranges) {
      const values = sheetInfo.get(sheetName).values;
      const { escalationColumns, quantityColumns } = columnsBySheet.get(sheetName);
      const [startRow, endRow] = rangeAddress.split(":").map(Number);

      for (let excelRow = startRow; excelRow <= endRow; excelRow++) {
        const rowKey = `${sheetName}!${excelRow}`;
        if (seenRows.has(rowKey)) continue;

        const row = values[excelRow - 1];
        if (!row) continue;
        seenRows.add(rowKey);

        // Year 1 is the cost line itself; column A holds its Salesforce Id.
        const salesforceId = String(row[0] ?? "").trim();

        const escalations = {};
        (quantitiesOnly ? [] : escalationColumns).forEach(({ index, field, year }) => {
          const num = toNumber(row[index]);
          if (num === null) return;
          escalations[year] = { ...escalations[year], [field]: num };
        });

        const quantities = {};
        quantityColumns.forEach(({ index, year }) => {
          const num = toNumber(row[index]);
          if (num !== null) quantities[year] = num;
        });

        rows.push({ sheetName, rowNumber: excelRow, salesforceId, escalations, quantities });
      }
    }

    return rows;
  });

  return updateEscalations(escalationData);
}

export async function updateEscalations(escalationData) {
  const skipped = [];

  const rows = (escalationData ?? []).filter((row) => {
    const hasValues =
      Object.keys(row.escalations ?? {}).length > 0 || Object.keys(row.quantities ?? {}).length > 0;
    if (!hasValues) return false;
    if (!row.salesforceId) {
      skipped.push({
        row: `${row.sheetName}!${row.rowNumber}`,
        reason: "No Salesforce Id in column A",
      });
      return false;
    }
    return true;
  });

  if (rows.length === 0) {
    console.log("No escalation rows to update.");
    if (skipped.length) console.warn(`${skipped.length} row(s) skipped:`, skipped);
    return { attempted: 0, updated: 0, skipped, failures: [] };
  }

  const parentIds = [...new Set(rows.map((r) => r.salesforceId))];

  // One target per (cost line, year): the same year of the same cost line can be named by
  // more than one spreadsheet row, and those have to merge into a single record update.
  const targets = new Map();

  const targetFor = (row, year) => {
    const key = `${row.salesforceId}|${year}`;
    if (!targets.has(key)) {
      targets.set(key, { parentId: row.salesforceId, year, fields: {}, requestedBy: [] });
    }
    const target = targets.get(key);
    // One entry per contributing value, so an unresolved target reports the same number of
    // skipped values as the rows that asked for it.
    target.requestedBy.push(`${row.sheetName}!${row.rowNumber}`);
    return target;
  };

  for (const row of rows) {
    for (const [yearStr, fields] of Object.entries(row.escalations)) {
      Object.assign(targetFor(row, Number(yearStr)).fields, fields);
    }

    for (const [yearStr, value] of Object.entries(row.quantities)) {
      targetFor(row, Number(yearStr)).fields.ManualQuantity__c = value;
    }
  }

  const years = [...new Set([...targets.values()].map((t) => t.year))].sort((a, b) => a - b);

  let yearMap = await fetchYearMap(parentIds);

  const failures = [];
  let attempted = 0;

  // QuoteCostItem_After cascades from whichever year line was updated to every other year
  // line in the same family, which rules out sending a cost line and its own child years
  // together: they share one Trigger.new, so the Apex collects the same child twice and
  // Salesforce rejects the whole request with "Duplicate id in list". Worse, when the request
  // does go through, the parent's cascade runs after every record in it is written and
  // redistributes the parent's quantity over the year lines — overwriting the ManualQuantity__c
  // just set on year 2. One request per year, ascending, keeps each family to a single record
  // per DML and puts every explicit value after the cascade that would otherwise clobber it.
  for (const year of years) {
    let pending = [...targets.values()].filter((t) => t.year === year);
    const updates = [];

    for (let attempt = 1; ; attempt++) {
      const unresolved = [];

      for (const target of pending) {
        // Year 1 is the cost line itself, so it resolves even for a line that has no child
        // year records at all.
        const recordId =
          year === 1 ? target.parentId : (yearMap.get(target.parentId)?.get(year) ?? null);

        if (recordId) {
          updates.push({
            Id: recordId,
            UniqueText__c: `${recordId}-${Date.now()}`,
            ...target.fields,
          });
        } else {
          unresolved.push(target);
        }
      }

      pending = unresolved;
      if (pending.length === 0 || attempt >= MAX_RESOLVE_ATTEMPTS) break;

      // A cost line that already has year lines, just not this year, is not a read that came
      // too early — the quote simply holds fewer years than the sheet has year columns, and no
      // amount of re-reading will produce the record. Only chase the lines that came back with
      // nothing at all, which is what reading before Apex has written them looks like.
      const retriable = pending.filter((t) => (yearMap.get(t.parentId)?.size ?? 0) === 0);
      if (retriable.length === 0) break;

      if (attempt > 1) await delay(RESOLVE_RETRY_DELAY_MS * 2 ** (attempt - 2));
      console.log(
        `Re-reading year ${year} records for ${retriable.length} unresolved cost line(s) (attempt ${attempt + 1}).`
      );

      // Merge rather than replace: the map is shared by every year, and this read only covers
      // the cost lines still missing one.
      const refreshed = await fetchYearMap([...new Set(retriable.map((t) => t.parentId))]);
      for (const [parentId, byYear] of refreshed) {
        if (!yearMap.has(parentId)) yearMap.set(parentId, new Map());
        for (const [refreshedYear, recordId] of byYear) {
          yearMap.get(parentId).set(refreshedYear, recordId);
        }
      }

      // The batched read can still come back short; ask for the stragglers one at a time. This
      // costs a round trip per record, so it is capped: when thousands are still missing the
      // cause is Apex not having created them yet, not a short read, and another bulk re-read
      // on the next attempt is both cheaper and more likely to find them.
      const stragglers = retriable.filter((t) => !yearMap.get(t.parentId)?.has(year));

      if (stragglers.length > MAX_SINGLE_LOOKUPS) {
        console.warn(
          `${stragglers.length} cost line(s) still missing year ${year} after the bulk re-read — ` +
            `skipping the per-record lookup (cap ${MAX_SINGLE_LOOKUPS}) and re-reading in bulk instead.`
        );
      } else {
        const found = await mapWithConcurrency(stragglers, QUERY_CONCURRENCY, (target) =>
          fetchSingleYearId(target.parentId, year).then((recordId) => ({ target, recordId }))
        );

        for (const { target, recordId } of found) {
          if (!recordId) continue;
          if (!yearMap.has(target.parentId)) yearMap.set(target.parentId, new Map());
          yearMap.get(target.parentId).set(year, recordId);
        }
      }
    }

    for (const target of pending) {
      // Naming the years the cost line does have separates "the quote is shorter than the
      // sheet's year columns" from "the year line has not been created yet".
      const knownYears = [...(yearMap.get(target.parentId)?.keys() ?? [])].sort((a, b) => a - b);
      const detail = knownYears.length
        ? ` (Salesforce holds year ${knownYears.join(", ")} for this cost line)`
        : "";

      for (const rowLabel of target.requestedBy) {
        skipped.push({
          row: rowLabel,
          reason: `No QuoteCostItem__c for parent=${target.parentId}, year=${year}${detail}`,
        });
      }
    }

    if (updates.length === 0) continue;

    console.log(`Sending ${updates.length} record(s) for year ${year} to Salesforce.`);
    if (updates.length <= MAX_LOGGED_RECORDS) console.log(updates);
    const result = await SalesforceService.updateEscalationRecords(updates);
    failures.push(...result.failures);
    attempted += updates.length;
  }

  if (skipped.length) console.warn(`${skipped.length} value(s) skipped:`, skipped);

  if (attempted === 0) {
    console.log("No records to update.");
    return { attempted: 0, updated: 0, skipped, failures: [] };
  }

  return {
    attempted,
    updated: attempted - failures.length,
    skipped,
    failures,
  };
}
