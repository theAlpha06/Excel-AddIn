export let isListenerPaused = false;

export async function handleSheetChange(event) {
    if (isListenerPaused) return;
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(event.worksheetId);

            console.log("Change detected:");
            console.log("Sheet ID:", event.worksheetId);
            console.log("Address:", event.address);
            console.log("Change Type:", event.changeType);
            console.log("Source:", event.source); 

            switch (event.changeType) {

                case "RangeEdited": {
                    const range = sheet.getRange(event.address);
                    range.load(["values", "address", "formulas"]);
                    await context.sync();

                    console.log("Cell Edited");
                    console.log("Address:", range.address);
                    console.log("New Value:", range.values);
                    console.log("Formula:", range.formulas);
                    break;
                }

                case "RowInserted": {
                    const range = sheet.getRange(event.address);
                    range.load(["address", "rowIndex", "rowCount"]);
                    await context.sync();

                    console.log("Row Inserted");
                    console.log("At row index:", range.rowIndex);
                    console.log("Row count:", range.rowCount);
                    break;
                }

                case "RowDeleted": {
                    console.log("Row Deleted");
                    console.log("Address:", event.address);
                    break;
                }

                case "ColumnInserted": {
                    const range = sheet.getRange(event.address);
                    range.load(["address", "columnIndex", "columnCount"]);
                    await context.sync();

                    console.log("Column Inserted");
                    console.log("At column index:", range.columnIndex);
                    console.log("Column count:", range.columnCount);
                    console.log("Address:", range.address);
                    break;
                }

                case "ColumnDeleted": {
                    console.log("Column Deleted");
                    console.log("Address:", event.address);
                    break;
                }

                case "CellInserted": {      
                    const range = sheet.getRange(event.address);
                    range.load(["values", "address"]);
                    await context.sync();

                    console.log("Cell(s) Inserted");
                    console.log("Address:", range.address);
                    console.log("Values:", range.values);
                    break;
                }

                case "CellDeleted": {
                    console.log("Cell(s) Deleted");
                    console.log("Address:", event.address);
                    break;
                }

                default: {
                    console.log("Unknown change type:", event.changeType);
                    break;
                }
            }
        });
    } catch (err) {
        console.error("Change handler error:", err);
    }
}

export const pauseListener = () => { isListenerPaused = true; };
export const resumeListener = () => { isListenerPaused = false; };