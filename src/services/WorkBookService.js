import SalesforceService from "../services/SalesforceService";

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const getFileWithRetry = async (maxRetries = 30, delayMs = 1000) => {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const file = await new Promise((resolve, reject) => {
        Office.context.document.getFileAsync(
          Office.FileType.Compressed,
          { sliceSize: 4 * 1024 * 1024 },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result.value);
            } else {
              reject(new Error(result.error?.message || "Failed to get workbook"));
            }
          }
        );
      });

      return file;
    } catch (err) {
      const msg = err.message?.toLowerCase() || "";

      const retryable =
        msg.includes("closeasync") ||
        msg.includes("before retrieving another") ||
        msg.includes("network problem") ||
        msg.includes("busy");

      if (!retryable || attempt === maxRetries) {
        throw err;
      }

      console.warn(
        `[uploadWorkbook] Workbook busy. Retry ${attempt}/${maxRetries}...`
      );

      await delay(delayMs);
    }
  }
};

const getSliceAsync = (file, index) =>
  new Promise((resolve, reject) => {
    file.getSliceAsync(index, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });

const closeFileAsync = (file) =>
  new Promise((resolve) => {
    file.closeAsync(() => resolve());
  });

export async function uploadWorkbook(parentId) {
  // Give Excel a moment to finish writing changes
  await Excel.run({ delayForCellEdit: true }, async (context) => {
    await context.sync();
  });

  await delay(1000);

  const file = await getFileWithRetry();

  try {
    const slices = [];

    for (let i = 0; i < file.sliceCount; i++) {
      const slice = await getSliceAsync(file, i);
      slices.push(new Uint8Array(slice.data));
    }

    const blob = new Blob(slices, {
      type:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const base64 = await new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = () => {
        resolve(reader.result.split(",")[1]);
      };

      reader.onerror = reject;

      reader.readAsDataURL(blob);
    });

    return await SalesforceService.uploadToSalesforce(
      "Workbook.xlsx",
      base64,
      parentId
    );
  } finally {
    await closeFileAsync(file);
  }
}