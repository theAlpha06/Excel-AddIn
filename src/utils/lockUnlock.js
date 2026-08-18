export async function lockEntireSheet() {
  await Excel.run({ delayForCellEdit: true }, async (context) => {
    const sheet1 = context.workbook.worksheets.getItem("Final Quote Header & Lines");

    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");

    await context.sync();

    const lineTemplateSheets = worksheets.items.filter(sheet =>
      sheet.name.startsWith("Template")
    );

    try {
      sheet1.protection.unprotect();
    } catch (_) {}

    lineTemplateSheets.forEach(sheet => {
      try {
        sheet.protection.unprotect();
      } catch (_) {}
    });

    sheet1.getUsedRange().format.protection.locked = true;

    lineTemplateSheets.forEach(sheet => {
      sheet.getUsedRange().format.protection.locked = true;
    });

    sheet1.protection.protect({
      allowAutoFilter: true,
      allowSort: true,
      allowInsertRows: true,
    });

    lineTemplateSheets.forEach(sheet => {
      sheet.protection.protect({
        allowAutoFilter: true,
        allowSort: true,
        allowInsertRows: true,
      });
    });

    await context.sync();
  });
}

export async function restoreDefaultProtection() {
  await Excel.run({ delayForCellEdit: true }, async (context) => {
    const sheet1 = context.workbook.worksheets.getItem("Final Quote Header & Lines");

    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");

    await context.sync();

    const lineTemplateSheets = worksheets.items.filter(sheet =>
      sheet.name.startsWith("Template")
    );

    try {
      sheet1.protection.unprotect();
    } catch (_) {}

    lineTemplateSheets.forEach(sheet => {
      try {
        sheet.protection.unprotect();
      } catch (_) {}
    });

    sheet1.getUsedRange().format.protection.locked = false;
    sheet1.getRange("A4:A18").format.protection.locked = true;
    sheet1.getRange("A20").format.protection.locked = true;
    sheet1.getRange("A21:I21").format.protection.locked = true;
    sheet1.getRange("G6:G13").format.protection.locked = true;
    sheet1.getRange("G15:L15").format.protection.locked = true;
    sheet1.getRange("G5:L5").format.protection.locked = true;
    sheet1.getRange("G14:I14").format.protection.locked = true;

    lineTemplateSheets.forEach(sheet => {
      sheet.getUsedRange().format.protection.locked = false;
      sheet.getRange("A5:D5").format.protection.locked = true;
    });

    sheet1.protection.protect({
      allowAutoFilter: true,
      allowSort: true,
      allowInsertRows: true,
      allowInsertColumns: true,
    });

    lineTemplateSheets.forEach(sheet => {
      sheet.protection.protect({
        allowAutoFilter: true,
        allowSort: true,
        allowInsertRows: true,
        allowInsertColumns: true,
      });
    });

    await context.sync();
  });
}