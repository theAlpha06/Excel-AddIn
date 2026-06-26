export async function lockEntireSheet() {
  await Excel.run(async (context) => {
    const sheet1 = context.workbook.worksheets.getItem("Final Quote Header & Lines");
    const sheet2 = context.workbook.worksheets.getItem("Lines Template");

    // ✅ Always try to unprotect — throws if no password set, which is fine
    try { sheet1.protection.unprotect(); } catch (_) {}
    try { sheet2.protection.unprotect(); } catch (_) {}

    sheet1.getUsedRange().format.protection.locked = true;
    sheet2.getUsedRange().format.protection.locked = true;

    sheet1.protection.protect({
      allowAutoFilter: true,
      allowSort: true,
      allowInsertRows: true,
    });

    sheet2.protection.protect({
      allowAutoFilter: true,
      allowSort: true,
      allowInsertRows: true,
    });

    await context.sync();
  });
}

export async function restoreDefaultProtection() {
  await Excel.run(async (context) => {
    const sheet1 = context.workbook.worksheets.getItem("Final Quote Header & Lines");
    const sheet2 = context.workbook.worksheets.getItem("Lines Template");

    // ✅ Always unprotect unconditionally
    try { sheet1.protection.unprotect(); } catch (_) {}
    try { sheet2.protection.unprotect(); } catch (_) {}

    // Unlock all cells first
    sheet1.getUsedRange().format.protection.locked = false;
    sheet2.getUsedRange().format.protection.locked = false;

    // Lock specific ranges
    sheet1.getRange("A4:A18").format.protection.locked = true;
    sheet1.getRange("A20").format.protection.locked = true;
    sheet1.getRange("A21:I21").format.protection.locked = true;
    sheet1.getRange("G6:G13").format.protection.locked = true;
    sheet2.getRange("A5:D5").format.protection.locked = true;

    sheet1.protection.protect({
      allowAutoFilter: true,
      allowSort: true,
      allowInsertRows: true,
    });

    sheet2.protection.protect({
      allowAutoFilter: true,
      allowSort: true,
      allowInsertRows: true,
    });

    await context.sync();
  });
}