export async function runWithRetry(callback) {
    while (true) {
        try {
            return await Excel.run({ delayForCellEdit: true }, callback);
        } catch (e) {
            if (
                e.code === Excel.ErrorCodes.invalidOperation ||
                e.message.includes("cell-editing mode")
            ) {
                console.log("Waiting for user to finish editing...");
                await new Promise(r => setTimeout(r, 1000));
                continue;
            }

            throw e;
        }
    }
}