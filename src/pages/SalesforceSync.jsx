import React, { useEffect, useState } from "react";
import Progress from "../components/Progress";
import SalesforceService from "../services/SalesforceService";
import { Button } from "@fluentui/react-components";
import {
  Dismiss16Regular,
  ArrowSync16Regular,
  ArrowUpload16Regular,
  ArrowClockwise16Regular,
} from "@fluentui/react-icons";
import { populateSalesforceIds } from "../utils/populateSalesforceIds";
import { useDiff } from "../context/DiffContext";
import { pushDiffToSalesforce } from "../services/DiffSyncService";
import { clearPendingChanges } from "../utils/diff";
import WorkBookAPI from "../components/WorkBookAPI";

const SalesforceSync = () => {
  const [pushState, setPushState] = useState(null);
  const [isPushing, setIsPushing] = useState(false);
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [error, setError] = useState(null);
  const [quoteData, setQuoteData] = useState([]);
  const [quoteExistsOnLoad, setQuoteExistsOnLoad] = useState(false);
  const sheetName = "Final Quote Header & Lines";

  // Diff context
  const { pendingPayload, clearPayload } = useDiff();

  const hasPendingChanges =
    pendingPayload &&
    (pendingPayload.quoteHeader !== null ||
      (pendingPayload.quoteLineItems?.length ?? 0) > 0 ||
      (pendingPayload.quoteCostItems?.length ?? 0) > 0 ||
      (pendingPayload.escalations?.length ?? 0) > 0);

  useEffect(() => {
    readExcelData();
  }, []);

  // ── Helpers ────────────────────────────────────────────────────────────────

  const headerIndex = (data) => {
    for (let i = 0; i < data.length; i++) if (data[i][0] === "SECTION A — QUOTE HEADER") return i;
    return -1;
  };

  const linesIndex = (data) => {
    for (let i = 0; i < data.length; i++)
      if (data[i][0] === "SECTION B — QUOTE LINE ITEMS") return i;
    return -1;
  };

  const headerData = (data, index, end) => {
    const header = {};
    for (let i = index; i < end - 2; i++) {
      let field = "",
        value = "";
      for (let j = 0; j < data[i].length; j++) {
        if (data[i][j] === "") continue;
        if (field) {
          value = data[i][j];
        } else {
          field = data[i][j];
        }
        if (field && value) {
          header[field] = value;
          field = "";
          value = "";
        }
      }
    }
    return header;
  };

  function excelDateToYYYYMMDD(value) {
    if (!value) return null;
    if (typeof value === "string") {
      const parsed = new Date(value);
      if (isNaN(parsed.getTime())) return null;
      return parsed.toISOString().split("T")[0];
    }
    if (typeof value === "number") {
      const utc_days = Math.floor(value - 25569);
      const utc_value = utc_days * 86400;
      const date = new Date(utc_value * 1000);
      if (isNaN(date.getTime())) return null;
      return date.toISOString().split("T")[0];
    }
    return null;
  }

  const fetchProductMapByPart = async (parts) => {
    if (!parts || parts.length === 0) return {};
    const uniqueParts = [...new Set(parts.map((p) => String(p).trim()))];
    const formattedParts = uniqueParts.map((p) => `'${p.replace(/'/g, "\\'")}'`).join(",");
    const data = await SalesforceService.query(
      `SELECT Id, ProductCode FROM Product2 WHERE ProductCode IN (${formattedParts})`
    );
    const productMap = {};
    (data.records || []).forEach((rec) => {
      productMap[rec.ProductCode] = rec.Id;
    });
    return productMap;
  };

  const extractRangeFromFormula = (cellValue) => {
    if (typeof cellValue !== "string" || !cellValue.startsWith("=")) return null;
    const match =
      cellValue.match(/[A-Z]+\('?([^'!]+)'?!([$A-Z]+\d+:[$A-Z]+\d+)\)/i) ??
      cellValue.match(/'([^']+)'!([$A-Z]+\d+:[$A-Z]+\d+)/i) ??
      cellValue.match(/([^'!(]+)!([$A-Z]+\d+:[$A-Z]+\d+)/i);
    if (!match) return null;
    const [rowStart, rowEnd] = match[2]
      .replace(/\$/g, "")
      .split(":")
      .map((a) => a.replace(/[A-Z]+/gi, ""));
    return {
      sheetName: match[1].trim(),
      rangeAddress: `${rowStart}:${rowEnd}`,
    };
  };

  const linesData = (formulas, values, linesStartIndex, linesHeaderRow) => {
    const QUANTITY_COL = 3;
    const COST_PRICE_COL = 5;

    return formulas
      .slice(linesStartIndex)
      .map((formulaRow, offset) => {
        const valueRow = values[linesStartIndex + offset];
        const line = {};

        formulaRow.forEach((cell, j) => {
          const header = linesHeaderRow[j]?.replace(/(\r\n|\n|\r)/g, " ");
          if (!header) return;

          if (j === QUANTITY_COL || j === COST_PRICE_COL) {
            const formula = formulaRow[j];
            const value = valueRow?.[j];
            const resolved = typeof formula === "string" && formula.startsWith("=") ? value : value;
            if (resolved !== "" && resolved != null) line[header] = resolved;
          } else {
            if (cell !== "") line[header] = cell;
          }
        });

        line.cell = `${sheetName}!${linesStartIndex + offset + 1}`;
        line.rowIndex = linesStartIndex + offset + 1;

        const rawQty = formulaRow[QUANTITY_COL];
        const range = extractRangeFromFormula(String(rawQty ?? ""));
        if (range) line.quantityRange = range;

        return line;
      })
      .filter((line) => Object.keys(line).length > 1);
  };

  const ensureOpportunity = async (quoteData) => {
    try {
      const accountNameRaw = quoteData["Account Name * (Also Add Account Number)"];
      const opportunityName = quoteData["Opportunity Name *"];
      const LOB = quoteData["LOB"];
      const subLob = quoteData["Sub LOB"];
      const serviceType = quoteData["Service Type"];
      const salePriceBookName = quoteData["Sale Price Book"];
      const accountName = accountNameRaw.replace("(Lookup)", "").trim();

      const oppRes = await SalesforceService.query(
        `SELECT Id, Name, AccountId, Pricebook2Id, StageName FROM Opportunity WHERE Name = '${opportunityName}' LIMIT 1`
      );

      if (oppRes.records.length) {
        const stage = oppRes.records[0].StageName;
        if (stage && stage != "Working Platform") {
          const msg = `Opportunity is not in Working Platform stage`;
          setError(msg);
          return { error: msg };
        }
      }

      let accountId, priceBookId;

      if (!oppRes.records.length) {
        const accountRes = await SalesforceService.query(
          `SELECT Id, Name FROM Account WHERE Name = '${accountName}' LIMIT 1`
        );
        if (!accountRes.records.length) {
          const message = `Account not found: ${accountName}`;
          setError(message);
          return { error: message };
        }
        accountId = accountRes.records[0].Id;

        const priceRes = await SalesforceService.query(
          `SELECT Id, Name FROM PriceBook2 WHERE Name = '${salePriceBookName}' LIMIT 1`
        );
        if (!priceRes.records.length) {
          const message = `Price Book not found: ${salePriceBookName}`;
          setError(message);
          return { error: message };
        }
        priceBookId = priceRes.records[0].Id;
      }

      let opportunityId;

      if (oppRes.records.length) {
        opportunityId = oppRes.records[0].Id;
        accountId = oppRes.records[0].AccountId;
        priceBookId = oppRes.records[0].Pricebook2Id;
      } else {
        const closeDate = new Date();
        closeDate.setDate(closeDate.getDate() + 60);
        const createRes = await SalesforceService.createRecords("Opportunity", {
          Name: opportunityName,
          AccountId: accountId,
          CloseDate: closeDate.toISOString().split("T")[0],
          StageName: "Working Platform",
          Pricebook2Id: priceBookId,
          LOB__c: LOB,
          Sub_LOB__c: subLob,
          ServiceType__c: serviceType,
        });
        opportunityId = createRes[0].id;
      }

      return { opportunityId, accountId, priceBookId };
    } catch (err) {
      const message =
        err?.body?.message || err?.message || err[0]?.message || "Failed to ensure Opportunity";
      setError(message);
      return { error: message };
    }
  };

  const columnIndexToLetter = (index) => {
    let letter = "",
      n = index + 1;
    while (n > 0) {
      const rem = (n - 1) % 26;
      letter = String.fromCharCode(65 + rem) + letter;
      n = Math.floor((n - 1) / 26);
    }
    return letter;
  };

  const fetchDynamicStructuredRows = async (ranges) => {
    return Excel.run(async (context) => {
      const HEADER_ROW = 5;
      const loaded = [];

      for (const { sheetName: sName, rangeAddress } of ranges) {
        const [rowStartStr, rowEndStr] = rangeAddress.split(":");
        const rowStart = parseInt(rowStartStr);
        const rowEnd = parseInt(rowEndStr);
        const sheet = context.workbook.worksheets.getItem(sName);
        const usedRange = sheet.getUsedRange();
        usedRange.load(["columnCount", "columnIndex"]);
        loaded.push({ sName, sheet, rowStart, rowEnd, usedRange });
      }

      await context.sync();

      const toLoad = [];

      for (const { sName, sheet, rowStart, rowEnd, usedRange } of loaded) {
        const colCount = usedRange.columnCount;
        const colStart = usedRange.columnIndex;
        const startLetter = columnIndexToLetter(colStart);
        const endLetter = columnIndexToLetter(colStart + colCount - 1);

        const headerRange = sheet.getRange(`${startLetter}${HEADER_ROW}:${endLetter}${HEADER_ROW}`);
        const dataRange = sheet.getRange(`${startLetter}${rowStart}:${endLetter}${rowEnd}`);

        headerRange.load("values");
        dataRange.load("values");

        toLoad.push({ sName, headerRange, dataRange, rowStart });
      }

      await context.sync();

      const results = [];

      for (const { sName, headerRange, dataRange, rowStart } of toLoad) {
        const headers = headerRange.values[0].map((h) =>
          (h || "")
            .toString()
            .trim()
            .replace(/(\r\n|\n|\r)/g, " ")
        );
        dataRange.values.forEach((row, i) => {
          if (!row || row.every((cell) => cell === "" || cell === null)) return;
          const rowData = {};
          row.forEach((cell, colIdx) => {
            if (cell !== "") rowData[headers[colIdx] || `Col_${colIdx + 1}`] = cell;
          });
          rowData.cell = `${sName}!${rowStart + i}`;
          rowData.rowIndex = rowStart + i;
          results.push(rowData);
        });
      }

      return results;
    });
  };

  const prepareData = async (formulas, values) => {
    const headerStartIndex = headerIndex(formulas) + 2;
    const linesStartIndex = linesIndex(formulas) + 3;
    const header = headerData(formulas, headerStartIndex, linesStartIndex);
    const linesHeaderRow = formulas[linesStartIndex - 2];
    const rawLines = linesData(formulas, values, linesStartIndex, linesHeaderRow);

    const opp = await ensureOpportunity(header);
    if (!opp.opportunityId || !opp.accountId || opp.error) return;

    const rangeMap = new Map();
    rawLines.forEach((line) => {
      if (line.quantityRange) {
        const key = `${line.quantityRange.sheetName}!${line.quantityRange.rangeAddress}`;
        rangeMap.set(key, line.quantityRange);
      }
    });

    let allCostRows = [];
    if (rangeMap.size > 0) {
      allCostRows = await fetchDynamicStructuredRows([...rangeMap.values()]);
    }

    const partSet = new Set();
    rawLines.forEach((line) => {
      if (line.Part) partSet.add(String(line.Part).trim());
    });
    allCostRows.forEach((r) => {
      if (r["Part"]) partSet.add(String(r["Part"]).trim());
    });

    const productMap = await fetchProductMapByPart([...partSet]);

    const quoteLines = rawLines.map((line) => {
      const { quantityRange, Part, ...lineFields } = line;
      const productId = productMap[String(Part ?? "").trim()];
      if (!productId) console.warn(`No Product found for Part: ${Part}`);

      if (quantityRange) {
        const { sheetName: costSheet, rangeAddress } = quantityRange;
        const [rowStartStr, rowEndStr] = rangeAddress.split(":");
        const rowStart = parseInt(rowStartStr);
        const rowEnd = parseInt(rowEndStr);

        const costRows = allCostRows.filter((r) => {
          const colonIdx = r.cell.lastIndexOf("!");
          const rSheet = r.cell.substring(0, colonIdx);
          const rRowNum = parseInt(r.cell.substring(colonIdx + 1));
          return rSheet === costSheet && rRowNum >= rowStart && rRowNum <= rowEnd;
        });

        return {
          ...lineFields,
          Product2Id: productId,
          PricebookEntryId: opp.priceBookId,
          quoteLineCosts: costRows
            .filter((r) => {
              const quantity = r["Quantity"];
              const part = r["Part"];
              const partDescription = r["Part Description"];

              return (
                quantity !== null &&
                quantity !== undefined &&
                quantity !== "" &&
                part !== null &&
                part !== undefined &&
                String(part).trim() !== "" &&
                partDescription !== null &&
                partDescription !== undefined &&
                String(partDescription).trim() !== ""
              );
            })
            .map((r) => ({
              cell: r.cell,
              rowIndex: r.rowIndex,
              costType: r["Cost Type"] ?? r["Type"] ?? "Material",
              amount: parseFloat(r["Amount"] ?? r["Cost Price"] ?? r["Cost"] ?? 0),
              part: r["Part"],
              partDescription: r["Part Description"],
              manualCostPrice: r["Manual Cost Price"],
              manualQuantity: r["Quantity"],
              Product2Id: productMap[String(r["Part"] ?? "").trim()] ?? null,
            })),
        };
      }

      return {
        ...lineFields,
        Product2Id: productId,
        PricebookEntryId: opp.priceBookId
      };
    });

    const quoteData = [
      {
        rowIndex: headerStartIndex,
        name: header["Opportunity Name *"],
        OpportunityId: opp.opportunityId,
        accountId: opp.accountId,
        contractStartDate: excelDateToYYYYMMDD(header["Contract Start Date"]),
        manualContractEndDate: excelDateToYYYYMMDD(header["Manual Contract End Date"]),
        expirationDate: excelDateToYYYYMMDD(header["Expiration Date"]),
        billingFrequency: header["Billing Frequency"],
        contractScope: header["Contract Scope"],
        visitFrequency: header["Visit Frequency"],
        quoteRef: header["Quote Ref"],
        quoteLines,
      },
    ];

    setQuoteData(quoteData);
    console.log(quoteData);
    return quoteData;
  };

const readExcelData = async () => {
  setIsRefreshing(true);
  let result = null;
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const idCell = sheet.getRange("J7");
      const usedRange = sheet.getUsedRange();
      idCell.load("values");
      usedRange.load(["formulas", "values"]);
      await context.sync();

      const j7Value = String(idCell.values[0][0] ?? "").trim();
      setQuoteExistsOnLoad(!!j7Value);

      if (usedRange.formulas.length > 0) {
        result = await prepareData(usedRange.formulas, usedRange.values); // ← capture
      }
    });
  } catch (err) {
    console.error("Error reading Excel data:", err);
  } finally {
    setIsRefreshing(false);
  }
  return result; // ← return it
};

  // ── Full push (existing) ───────────────────────────────────────────────────
const handlePush = async () => {
    setIsPushing(true);
    let dataToPush = quoteData;

    if (hasNewRows && dataToPush[0].quoteRef) {
      try {
        const quoteHeader = await SalesforceService.query(
          `SELECT Id, ContractId, Status FROM Quote WHERE Id = '${quoteData[0].quoteRef}' LIMIT 1`
        );
        if (quoteHeader.records[0]?.Status === "Presented" && quoteHeader.records[0]?.ContractId != null) {
          setPushState((prev) => ({ ...prev, status: "Failed", message: "Quote is in Presented Stage. Can't add/delete new rows" }));
          setIsPushing(false);
          return;
        }
        if (quoteHeader.records[0]?.ContractId == null) {
          await SalesforceService.updateRecord("Quote", quoteData[0].quoteRef, {
            Status: "Rejected",
          });
        }
        dataToPush = await readExcelData();
      } catch (err) {
        setPushState((prev) => ({ ...prev, status: "Failed", message: `Failed to prepare data: ${err.message}` }));
        setIsPushing(false);
        return;
      }
    }

    try {
      await SalesforceService.pushDataToSalesforce(dataToPush, async (update) => {
        const recordErrors = update.results?.filter((r) => r.status === "error") ?? [];

        if (update.status === "Complete" && recordErrors.length > 0) {
          const uniqueMessages = [...new Set(recordErrors.map((r) => r.errorMessage))];
          setPushState({
            ...update,
            status: "Failed",
            message: `Sync completed with ${recordErrors.length} record error(s): ${uniqueMessages.join(", ")}`,
            recordErrors,
          });
          setIsPushing(false);
          return;
        }

        setPushState(update);

        if (update.status === "Complete") {
          try {
            setPushState((prev) => ({ ...prev, message: "Writing IDs to Excel…" }));
            await populateSalesforceIds(update);
            setPushState((prev) => ({ ...prev, message: "Done" }));
          } catch (err) {
            setPushState((prev) => ({ ...prev, status: "Failed", message: `Failed to write IDs to Excel: ${err.message}` }));
          } finally {
            setIsPushing(false);
          }
        }

        if (update.status === "Failed" || update.status === "RolledBack") {
          setIsPushing(false);
        }
      });
    } catch (err) {
      setPushState((prev) => ({ ...prev, status: "Failed", message: err.message }));
      setIsPushing(false);
    }
  };

  // ── Diff push (new) ────────────────────────────────────────────────────────

  const handleDiffPush = async () => {
    if (!pendingPayload) return;
    setIsPushing(true);
    setPushState({ status: "Pending", message: "Preparing…", total: 0, done: 0, errors: [] });

    try {
      await pushDiffToSalesforce(pendingPayload, (update) => {
        setPushState(update);
        if (update.status === "Complete") {
          clearPendingChanges();
          clearPayload();
          setIsPushing(false);
        }
        if (update.status === "Failed") {
          setIsPushing(false);
        }
      });
    } catch (err) {
      setPushState({ status: "Failed", message: err.message, total: 0, done: 0, errors: [] });
      setIsPushing(false);
    }
  };

  // ── Derived state ──────────────────────────────────────────────────────────
  const isFailed = pushState?.status === "Failed" || pushState?.status === "RolledBack";

  const quoteExists = quoteExistsOnLoad || (pendingPayload?.meta?.quoteExists ?? false);
  const hasNewRows = pendingPayload?.meta?.hasNewRows ?? false;
  let isPushBlocked = quoteExists && !hasNewRows;

  const pendingSummary = hasPendingChanges
    ? [
        pendingPayload.quoteHeader ? "1 quote header" : null,
        (pendingPayload.quoteLineItems?.length ?? 0)
          ? `${pendingPayload.quoteLineItems.length} line item(s)`
          : null,
        (pendingPayload.quoteCostItems?.length ?? 0)
          ? `${pendingPayload.quoteCostItems.length} cost item(s)`
          : null,
        (pendingPayload.escalations?.length ?? 0)
          ? `${pendingPayload.escalations.length} escalation(s)`
          : null,
      ]
        .filter(Boolean)
        .join(", ")
    : "";

  // ── Error screen ───────────────────────────────────────────────────────────

  if (error) {
    return (
      <div
        style={{
          margin: "20px",
          padding: "12px 16px",
          borderRadius: "6px",
          background: "#fef2f2",
          border: "1px solid #fecaca",
          color: "#b91c1c",
          fontSize: "13px",
        }}
      >
        {error}
      </div>
    );
  }

  // ── Render ─────────────────────────────────────────────────────────────────

  return (
    <>
      <style>{`
        @keyframes spin {
          from { transform: rotate(0deg); }
          to   { transform: rotate(360deg); }
        }
        .spin-icon { animation: spin 0.8s linear infinite; display: inline-flex; }
      `}</style>

      <div style={{ padding: "20px" }}>
        <div
          style={{
            background: "#f9fafb",
            border: "1px solid #e5e7eb",
            borderRadius: "8px",
            padding: "10px 5px",
            display: "flex",
            flexDirection: "column",
            gap: "14px",
          }}
        >
          {/* ── Row 1: Full Push + Refresh ── */}
          <div style={{ display: "flex", alignItems: "center", gap: "20px" }}>
            <Button
              appearance="primary"
              icon={isFailed ? <ArrowSync16Regular /> : <ArrowUpload16Regular />}
              onClick={handlePush}
              disabled={isPushing || !quoteData || quoteData.length === 0 || isPushBlocked}
              style={isFailed ? { borderColor: "#dc2626" } : {}}
              title={isPushBlocked ? "Quote already exists in Salesforce" : undefined}
            >
              {!quoteData || quoteData.length === 0
                ? "Preparing Data…"
                : isPushing
                  ? "Pushing…"
                  : isFailed
                    ? "Retry"
                    : isPushBlocked
                      ? "Push (Locked)"
                      : "Push"}
            </Button>

            <div style={{ width: "1px", height: "28px", background: "#d1d5db", margin: "0 6px" }} />

            <Button
              appearance="primary"
              icon={
                <span className={isRefreshing ? "spin-icon" : ""}>
                  <ArrowSync16Regular />
                </span>
              }
              onClick={() => readExcelData()}
              disabled={isPushing || isRefreshing}
            >
              Refresh Data
            </Button>

            {pushState && (
              <Button
                appearance="subtle"
                icon={<Dismiss16Regular />}
                title="Clear progress"
                onClick={() => {
                  setPushState(null);
                  setIsPushing(false);
                }}
                style={{ color: "#6b7280" }}
              />
            )}
          </div>

          {/* ── Row 2: Pending diff banner — only when there are unsync'd changes ── */}
          {hasPendingChanges && (
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: "10px",
                padding: "8px 12px",
                borderRadius: "6px",
                background: "#fffbeb",
                border: "1px solid #fcd34d",
              }}
            >
              <span style={{ fontSize: "12px", color: "#92400e", flex: 1 }}>
                <strong>{pendingSummary}</strong> changed — not yet synced to Salesforce
              </span>

              <Button
                appearance="primary"
                size="small"
                icon={<ArrowClockwise16Regular />}
                onClick={handleDiffPush}
                disabled={isPushing}
                style={{ background: "#d97706", borderColor: "#d97706", whiteSpace: "nowrap" }}
              >
                {isPushing ? "Syncing…" : "Sync Changes"}
              </Button>

              <Button
                appearance="subtle"
                size="small"
                icon={<Dismiss16Regular />}
                title="Discard pending changes"
                onClick={() => {
                  clearPendingChanges();
                  clearPayload();
                }}
                style={{ color: "#92400e" }}
              />
            </div>
          )}

          {/* ── Progress bar (shared) ── */}
          {pushState && <Progress pushState={pushState} />}

          {/* ── Per-record error list ── */}
          {(pushState?.errors?.length ?? 0) > 0 && (
            <div
              style={{
                background: "#fef2f2",
                border: "1px solid #fecaca",
                borderRadius: "6px",
                padding: "10px 12px",
                fontSize: "12px",
                color: "#b91c1c",
                display: "flex",
                flexDirection: "column",
                gap: "4px",
              }}
            >
              <strong style={{ marginBottom: "4px" }}>{pushState.errors.length} error(s):</strong>
              {pushState.errors.map((e, i) => (
                <div key={i} style={{ paddingLeft: "8px", borderLeft: "2px solid #fca5a5" }}>
                  <span style={{ fontWeight: 500 }}>
                    {e.objectApiName} · {e.id}:
                  </span>{" "}
                  {e.message}
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
      {quoteData && quoteData.length > 0 && quoteData[0].quoteRef && 
        <WorkBookAPI />
      }
    </>
  );
};

export default SalesforceSync;
