import React from "react";
import SalesforceService from "../services/SalesforceService";
import {
  DefaultButton,
  Dialog,
  DialogFooter,
  PrimaryButton,
} from "@fluentui/react";

const WorkBookAPI = () => {
  const [results, setResults] = React.useState(null);
  const [isLoading, setIsLoading] = React.useState(false);
  const [parentId, setParentId] = React.useState("");

  React.useEffect(() => {
    fetchParentId();
  }, []);

  async function fetchParentId() {
    try {
      await Excel.run({ delayForCellEdit: true }, async (context) => {
        const sheet = context.workbook.worksheets.getItem(
          "Final Quote Header & Lines"
        );

        const cell = sheet.getRange("J7");

        cell.load("values");

        await context.sync();

        const value = cell.values?.[0]?.[0];

        setParentId(value ? String(value).trim() : "");
      });
    } catch (error) {
      console.error("Error fetching J7 value:", error);
      setParentId("");
    }
  }

  async function passwordProtectDataInWorksheet() {
    await Excel.run({ delayForCellEdit: true }, async (context) => {
      let activeSheet = context.workbook.worksheets.getActiveWorksheet();

      activeSheet.load("protection/protected");

      await context.sync();

      if (!activeSheet.protection.protected) {
        activeSheet.protection.protect(null, "testing");
      }
    });
  }

  async function addOverlayOnUsedRange() {
    await Excel.run({ delayForCellEdit: true }, async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      const usedRange = sheet.getUsedRange();

      usedRange.load("address, format/*, columnCount, rowCount");

      await context.sync();

      const shape = sheet.shapes.addGeometricShape(
        Excel.GeometricShapeType.rectangle
      );

      const cell0 = usedRange.getCell(0, 0);

      const cellLast = usedRange.getCell(
        usedRange.rowCount - 1,
        usedRange.columnCount - 1
      );

      cell0.load("left, top");
      cellLast.load("left, top, width, height");

      await context.sync();

      const left = cell0.left;
      const top = cell0.top;

      const width = cellLast.left + cellLast.width - cell0.left;

      const height = cellLast.top + cellLast.height - cell0.top;

      shape.left = left;
      shape.top = top;
      shape.width = width;
      shape.height = height;

      shape.fill.setSolidColor("#000");
      shape.fill.transparency = 1.0;

      shape.lineFormat.weight = 10;
    });
  }

  const pushToSalesforce = async () => {
    if (!parentId) {
      return;
    }

    setIsLoading(true);

    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 4194304 },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const file = result.value;

          let sliceIndex = 0;

          let slices = [];

          const getNextSlice = () => {
            file.getSliceAsync(sliceIndex, async (sliceResult) => {
              if (
                sliceResult.status === Office.AsyncResultStatus.Succeeded
              ) {
                slices.push(new Uint8Array(sliceResult.value.data));

                sliceIndex++;

                if (sliceIndex < file.sliceCount) {
                  getNextSlice();
                } else {
                  file.closeAsync();

                  const blob = new Blob(slices, {
                    type:
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  });

                  const reader = new FileReader();

                  reader.onload = async function () {
                    try {
                      const base64Data = reader.result.split(",")[1];

                      console.log("Parent Id:", parentId);

                      const response =
                        await SalesforceService.uploadToSalesforce(
                          "Workbook.xlsx",
                          base64Data,
                          parentId
                        );

                      setResults(response);
                    } catch (error) {
                      console.error("Upload Error:", error);

                      setResults({
                        success: false,
                        message:
                          error?.message || "Workbook upload failed",
                      });
                    } finally {
                      setIsLoading(false);
                    }
                  };

                  reader.readAsDataURL(blob);
                }
              } else {
                console.error(
                  "getSliceAsync failed:",
                  sliceResult.error.message
                );

                setResults({
                  success: false,
                  message: sliceResult.error.message,
                });

                setIsLoading(false);
              }
            });
          };

          getNextSlice();
        } else {
          console.error(
            "getFileAsync failed:",
            result.error.message
          );

          setResults({
            success: false,
            message: result.error.message,
          });

          setIsLoading(false);
        }
      }
    );
  };

  return (
    <div style={{ width: "100%" }}>
      {results && (
        <Dialog
          hidden={!results}
          onDismiss={() => setResults(null)}
          dialogContentProps={{
            title: results.success
              ? "Workbook Uploaded Successfully"
              : "Workbook Upload Failed",

            subText: results.success
              ? "The workbook has been uploaded and linked successfully."
              : results.message || "Something went wrong.",
          }}
        >
          <div>
            {results.success ? (
              <>
                <h4>Content Version Id</h4>
                <pre>{results.contentVersionId}</pre>

                <h4>Content Document Id</h4>
                <pre>{results.contentDocumentId}</pre>

                <h4>Linked Record Id</h4>
                <pre>{parentId}</pre>
              </>
            ) : (
              <pre>{JSON.stringify(results, null, 2)}</pre>
            )}
          </div>

          <DialogFooter>
            <DefaultButton
              onClick={() => setResults(null)}
              text="Close"
            />
          </DialogFooter>
        </Dialog>
      )}

      <PrimaryButton
        style={{ width: "100%" }}
        text={
          isLoading
            ? "Uploading..."
            : !parentId
            ? "Parent Id Missing in J7"
            : "Push Workbook"
        }
        onClick={pushToSalesforce}
        disabled={isLoading || !parentId}
      />

      {/* <PrimaryButton
        style={{ width: "100%" }}
        text="Protect"
        onClick={passwordProtectDataInWorksheet}
      /> */}

      {/* <PrimaryButton
        style={{ width: "100%" }}
        text="Add Overlay"
        onClick={addOverlayOnUsedRange}
      /> */}
    </div>
  );
};

export default WorkBookAPI;