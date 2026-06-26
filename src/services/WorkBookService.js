import SalesforceService from "../services/SalesforceService";

export async function uploadWorkbook(parentId) {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 4194304 },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error.message));
          return;
        }

        const file = result.value;
        let sliceIndex = 0;
        const slices = [];

        const getNextSlice = () => {
          file.getSliceAsync(sliceIndex, async (sliceResult) => {
            if (
              sliceResult.status !== Office.AsyncResultStatus.Succeeded
            ) {
              reject(new Error(sliceResult.error.message));
              return;
            }

            slices.push(new Uint8Array(sliceResult.value.data));

            sliceIndex++;

            if (sliceIndex < file.sliceCount) {
              getNextSlice();
              return;
            }

            file.closeAsync();

            const blob = new Blob(slices, {
              type:
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            });

            const reader = new FileReader();

            reader.onload = async () => {
              try {
                const base64Data = reader.result.split(",")[1];

                const response =
                  await SalesforceService.uploadToSalesforce(
                    "Workbook.xlsx",
                    base64Data,
                    parentId
                  );

                resolve(response);
              } catch (err) {
                reject(err);
              }
            };

            reader.readAsDataURL(blob);
          });
        };

        getNextSlice();
      }
    );
  });
}