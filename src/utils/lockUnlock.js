const MASTER_SHEET = "Final Quote Header & Lines";
const COST_SECTION_MARKER = "section c - quote cost item";

const normalizeMarker = (value) =>
  String(value ?? "").replace(/[—–-]/g, "-").replace(/\s+/g, " ").trim().toLowerCase();

// Columns kept read-only for the user. restoreDefaultProtection() resets every cell's lock flag,
// so locks set by hand in Excel do not survive it: anything that must stay protected has to be
// listed here. Column A holds the Salesforce Ids, which only the add-in writes (see
// unprotectSheets(), used while populateSalesforceIds runs). Do not add Total Quantity or the year
// columns of the cost sheets: the user edits the first and the add-in writes the second.
const LOCKED_COLUMNS = {
  master: ["A:A"],
  costSheets: ["A:A"],
};

// Protection used while the user is editing. Rows can be added and deleted; Excel still refuses
// to delete a row that contains a locked cell, so the locked header rows stay put.
const EDITING_PROTECTION = {
  allowAutoFilter: true,
  allowSort: true,
  allowInsertRows: true,
  allowDeleteRows: true,
  allowInsertColumns: true,
};

// Protection used while a push is running: nothing can be inserted or deleted mid-sync.
const PUSH_PROTECTION = {
  allowAutoFilter: true,
  allowSort: true,
  allowInsertRows: true,
};

// Finds the master sheet and every quote cost item sheet. Cost sheets used to be recognised only
// by a name starting with "Template", so a cost sheet named any other way was never unlocked: it
// kept Excel's default of every cell locked, and Total Quantity and the year columns could not be
// edited. They are now also recognised by the "SECTION C - QUOTE COST ITEM" marker in column A —
// the same marker the change handler in diff.js uses to find their header row.
async function getProtectedSheets(context) {
  const worksheets = context.workbook.worksheets;
  worksheets.load("items/name");
  await context.sync();

  const master = worksheets.items.find((sheet) => sheet.name === MASTER_SHEET) ?? null;
  const candidates = worksheets.items.filter((sheet) => sheet.name !== MASTER_SHEET);

  const markerCells = candidates.map((sheet) => {
    const range = sheet.getRange("A1:A20");
    range.load("values");
    return range;
  });
  await context.sync();

  const costSheets = [];
  const costHeaderRows = new Map();
  candidates.forEach((sheet, i) => {
    const markerIndex = (markerCells[i].values ?? []).findIndex((row) =>
      normalizeMarker(row[0]).includes(COST_SECTION_MARKER)
    );
    if (markerIndex === -1 && !sheet.name.startsWith("Template")) return;
    costSheets.push(sheet);
    // Header row = the row after the marker, row 5 when there is no marker — same rule as diff.js.
    costHeaderRows.set(sheet, markerIndex === -1 ? 5 : markerIndex + 2);
  });

  return { master, costSheets, costHeaderRows };
}

// Unprotects each sheet in its own sync. A single failure — typically a sheet protected by hand
// with a password — used to fail the whole batch, so no sheet was unlocked at all. Returns the
// sheets that are now unprotected; the rest are left untouched and reported.
async function unprotectEach(context, sheets) {
  const unprotected = [];
  for (const sheet of sheets) {
    try {
      sheet.protection.unprotect();
      await context.sync();
      unprotected.push(sheet);
    } catch (err) {
      console.warn(
        `Could not unprotect "${sheet.name}" — if it was protected with a password, remove the ` +
          `password so the add-in can manage it. Leaving it as it is.`,
        err
      );
    }
  }
  return unprotected;
}

export async function lockEntireSheet() {
  await Excel.run({ delayForCellEdit: true }, async (context) => {
    const { master, costSheets } = await getProtectedSheets(context);
    const sheets = await unprotectEach(context, [master, ...costSheets].filter(Boolean));

    for (const sheet of sheets) {
      if (sheet === master) sheet.getUsedRange().format.protection.locked = true;
      sheet.protection.protect(PUSH_PROTECTION);
    }

    await context.sync();
  });
}

// Removes protection from the master and cost sheets without touching any lock flags, so the
// add-in can write into locked cells (the Salesforce Ids in column A, J7, J14). The caller must
// protect the sheets again afterwards — lockEntireSheet() or restoreDefaultProtection().
export async function unprotectSheets() {
  await Excel.run({ delayForCellEdit: true }, async (context) => {
    const { master, costSheets } = await getProtectedSheets(context);
    await unprotectEach(context, [master, ...costSheets].filter(Boolean));
  });
}

export async function restoreDefaultProtection() {
  await Excel.run({ delayForCellEdit: true }, async (context) => {
    const { master, costSheets, costHeaderRows } = await getProtectedSheets(context);
    const sheets = await unprotectEach(context, [master, ...costSheets].filter(Boolean));

    for (const sheet of sheets) {
      // Unlock the whole sheet, not just the used range, before re-locking only the ranges
      // below. Excel cells are locked by default, so anything outside the used range — new rows,
      // blank cells, the year columns of a row added at the bottom — stayed read-only on the
      // protected sheet. This also clears the sheet-wide lock lockEntireSheet() applies; without
      // that the push fails writing the Salesforce Ids into J7/J14.
      sheet.getRange().format.protection.locked = false;

      if (sheet === master) {
        sheet.getRange("A4:A18").format.protection.locked = true;
        sheet.getRange("A20").format.protection.locked = true;
        sheet.getRange("A21:I21").format.protection.locked = true;
        sheet.getRange("G6:G13").format.protection.locked = true;
        sheet.getRange("G15:L15").format.protection.locked = true;
        sheet.getRange("G5:L5").format.protection.locked = true;
        sheet.getRange("G14:I14").format.protection.locked = true;
        LOCKED_COLUMNS.master.forEach((address) => {
          sheet.getRange(address).format.protection.locked = true;
        });
      } else {
        // Only the header row and the Id column stay locked on a cost sheet; Total Quantity and
        // the year columns (1, 2, 3, …) must remain editable.
        const headerRow = costHeaderRows.get(sheet);
        sheet.getRange(`A${headerRow}:D${headerRow}`).format.protection.locked = true;
        LOCKED_COLUMNS.costSheets.forEach((address) => {
          sheet.getRange(address).format.protection.locked = true;
        });
      }

      sheet.protection.protect(EDITING_PROTECTION);
    }

    await context.sync();
  });
}
