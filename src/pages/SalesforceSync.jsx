import React, { useEffect, useState } from 'react';
import Profile from '../components/Profile';
import Progress from '../components/Progress';
import SalesforceService from '../services/SalesforceService';
import { Button } from '@fluentui/react-components';
import { populateSalesforceIds } from '../utils/populateSalesforceIds';

const SalesforceSync = () => {
  const [pushState, setPushState] = useState(null);
  const [isPushing, setIsPushing] = useState(false);
  const [error, setError] = useState(null);
  const sheetName = "Final Quote Header & Lines";

  useEffect(() => {
    readExcelData();
  }, []);

  const headerIndex = (data) => {
    for (let i = 0; i < data.length; i++)
      if (data[i][0] === "SECTION A — QUOTE HEADER") return i;
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
      let field = '', value = '';
      for (let j = 0; j < data[i].length; j++) {
        if (data[i][j] === '') continue;
        if (field) { value = data[i][j]; }
        else { field = data[i][j]; }
        if (field && value) {
          header[field] = value;
          field = ''; value = '';
        }
      }
    }
    return header;
  };
   
  function excelDateToYYYYMMDD(serial) {
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date = new Date(utc_value * 1000);

  return date.toISOString().split('T')[0];
}

const fetchProductMapByPart = async (parts) => {
  if (!parts || parts.length === 0) return {};

  const uniqueParts = [...new Set(parts.map(p => String(p).trim()))];

  const formattedParts = uniqueParts
    .map(p => `'${p.replace(/'/g, "\\'")}'`) 
    .join(',');

  const data = await SalesforceService.query(
      `SELECT Id, ProductCode FROM Product2 WHERE ProductCode IN (${formattedParts})`
    );
    console.log(data);
  const productMap = {};
  (data.records || []).forEach(rec => {
    productMap[rec.ProductCode] = rec.Id;
  });

  return productMap;
};
const extractRangeFromFormula = (cellValue) => {
  if (typeof cellValue !== 'string' || !cellValue.startsWith('=')) return null;

  const match = cellValue.match(/[A-Z]+\('?([^'!]+)'?!([$A-Z]+\d+:[$A-Z]+\d+)\)/i)
    ?? cellValue.match(/'([^']+)'!([$A-Z]+\d+:[$A-Z]+\d+)/i)
    ?? cellValue.match(/([^'!(]+)!([$A-Z]+\d+:[$A-Z]+\d+)/i);

  if (!match) return null;

  const [rowStart, rowEnd] = match[2].replace(/\$/g, '').split(':').map(a => a.replace(/[A-Z]+/gi, ''));

  return {
    sheetName:    match[1].trim(),
    rangeAddress: `${rowStart}:${rowEnd}`,
  };
};

  const linesData = (formulas, values, linesStartIndex, linesHeaderRow) => {
    const QUANTITY_COL   = 2; 
    const COST_PRICE_COL = 4; 

    return formulas.slice(linesStartIndex).map((formulaRow, offset) => {
      const valueRow = values[linesStartIndex + offset];
      const line = {};

      formulaRow.forEach((cell, j) => {
        const header = linesHeaderRow[j]?.replace(/(\r\n|\n|\r)/g, ' ');
        if (!header) return;

        if (j === QUANTITY_COL || j === COST_PRICE_COL) {
          const formula   = formulaRow[j];
          const value     = valueRow?.[j];
          const isFormula = typeof formula === 'string' && formula.startsWith('=');
          const resolved  = isFormula ? value : value;
          if (resolved !== '' && resolved != null) line[header] = resolved;
        } else {
          if (cell !== '') line[header] = cell;
        }
      });

      line.cell = `${sheetName}!${linesStartIndex + offset + 1}`;
      line.rowIndex = linesStartIndex + offset + 1;

      const rawQty = formulaRow[QUANTITY_COL];
      const range  = extractRangeFromFormula(String(rawQty ?? ''));
      if (range) line.quantityRange = range;

      return line;
    }).filter(line => Object.keys(line).length > 1); 
  };

  const ensureOpportunity = async (quoteData) => {
    const accountNameRaw  = quoteData["Account Name * (Also Add Account Number)"];
    const opportunityName = quoteData["Opportunity Name *"];
    const priceBookName   = quoteData["Price Book *"];
    const LOB = quoteData["LOB"];
    const subLob = quoteData["Sub LOB"];
    const serviceType = quoteData["Service Type"];
    const salePriceBookName = quoteData["Sale Price Book"];
    const accountName = accountNameRaw.replace("(Lookup)", "").trim();

    const accountRes = await SalesforceService.query(
      `SELECT Id, Name FROM Account WHERE Name = '${accountName}' LIMIT 1`
    );
    if (!accountRes.records.length) {
      setError(`Account not found: ${accountName}`);
      return {};
    }
    const accountId = accountRes.records[0].Id;

    const priceRes = await SalesforceService.query(
      `SELECT Id, Name FROM PriceBook2 WHERE Name = '${salePriceBookName}' LIMIT 1`
    );
    if (!priceRes.records.length) {
      setError(`Price Book not found: ${salePriceBookName}`);
      return {};
    }
    const priceBookId = priceRes.records[0].Id;

    const oppRes = await SalesforceService.query(
      `SELECT Id, Name FROM Opportunity WHERE Name = '${opportunityName}' LIMIT 1`
    );

    let opportunityId;
    if (oppRes.records.length) {
      opportunityId = oppRes.records[0].Id;
    } else {
      const closeDate = new Date();
      closeDate.setDate(closeDate.getDate() + 60);
      const createRes = await SalesforceService.createRecords("Opportunity", {
        Name:         opportunityName,
        AccountId:    accountId,
        CloseDate:    closeDate.toISOString().split("T")[0],
        StageName:    "Working Platform",
        Pricebook2Id: priceBookId,
        LOB__c:       LOB,
        Sub_LOB__c: subLob,
        ServiceType__c: serviceType,
      });
      console.log(createRes);
      opportunityId = createRes.id;
    }

    return { opportunityId, accountId, priceBookId };
  };

  const columnIndexToLetter = (index) => {
    let letter = '', n = index + 1;
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
        const [rowStartStr, rowEndStr] = rangeAddress.split(':');
        const rowStart = parseInt(rowStartStr);
        const rowEnd   = parseInt(rowEndStr);

        const sheet     = context.workbook.worksheets.getItem(sName);
        const usedRange = sheet.getUsedRange();
        usedRange.load(["columnCount", "columnIndex"]);

        loaded.push({ sName, sheet, rowStart, rowEnd, usedRange });
      }

      await context.sync();

      const toLoad = [];

      for (const { sName, sheet, rowStart, rowEnd, usedRange } of loaded) {
        const colCount    = usedRange.columnCount;
        const colStart    = usedRange.columnIndex;
        const startLetter = columnIndexToLetter(colStart);
        const endLetter   = columnIndexToLetter(colStart + colCount - 1);

        const headerRange = sheet.getRange(`${startLetter}${HEADER_ROW}:${endLetter}${HEADER_ROW}`);
        const dataRange   = sheet.getRange(`${startLetter}${rowStart}:${endLetter}${rowEnd}`);

        headerRange.load("values");
        dataRange.load("values");

        toLoad.push({ sName, headerRange, dataRange, rowStart });
      }

      await context.sync();

      const results = [];

      for (const { sName, headerRange, dataRange, rowStart } of toLoad) {
        const headers = headerRange.values[0].map(h =>
          (h || "").toString().trim().replace(/(\r\n|\n|\r)/g, " ")
        );

        dataRange.values.forEach((row, i) => {
          if (!row || row.every(cell => cell === "" || cell === null)) return;
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
  const linesStartIndex  = linesIndex(formulas) + 3;
  const header           = headerData(formulas, headerStartIndex, linesStartIndex);
  const linesHeaderRow   = formulas[linesStartIndex - 2];
  const rawLines         = linesData(formulas, values, linesStartIndex, linesHeaderRow);

  const opp = await ensureOpportunity(header);
  if (!opp.opportunityId || !opp.accountId) return;

  const partSet = new Set();
  rawLines.forEach(line => {
    if (line.Part) partSet.add(String(line.Part));
  });

  const productMap = await fetchProductMapByPart([...partSet]); 

  const rangeMap = new Map();
  rawLines.forEach(line => {
    if (line.quantityRange) {
      const key = `${line.quantityRange.sheetName}!${line.quantityRange.rangeAddress}`;
      rangeMap.set(key, line.quantityRange);
    }
  });

  let allCostRows = [];
  if (rangeMap.size > 0) {
    allCostRows = await fetchDynamicStructuredRows([...rangeMap.values()]);
  }

  const quoteLines = rawLines.map(line => {
    const { quantityRange, Part, ...lineFields } = line;

    const productId = productMap[String(Part)];

    if (!productId) {
      console.warn(`No Product found for Part: ${Part}`);
    }

    if (quantityRange) {
      const { sheetName: costSheet, rangeAddress } = quantityRange;
      const [rowStartStr, rowEndStr] = rangeAddress.split(':');
      const rowStart = parseInt(rowStartStr);
      const rowEnd   = parseInt(rowEndStr);

      const costRows = allCostRows.filter(r => {
        const colonIdx = r.cell.lastIndexOf('!');
        const rSheet   = r.cell.substring(0, colonIdx);
        const rRowNum  = parseInt(r.cell.substring(colonIdx + 1));
        return rSheet === costSheet && rRowNum >= rowStart && rRowNum <= rowEnd;
      });

      return {
        ...lineFields,
        Product2Id: productId,
        PricebookEntryId: opp.priceBookId,
        quoteLineCosts: costRows.map(r => ({
          cell: r.cell,
          rowIndex: r.rowIndex,
          costType: r["Cost Type"] ?? r["Type"] ?? "Material",
          amount: parseFloat(r["Amount"] ?? r["Cost Price"] ?? r["Cost"] ?? 0),
          part: r["Part"],
          partDescription: r["Part Description"]
        })),
      };
    }

    return {
      ...lineFields,
      Product2Id: productId,
      PricebookEntryId: opp.priceBookId
    };
  });
  
  console.log({...header, "quoteLines": quoteLines});
  const quoteData = [{
    rowIndex: headerStartIndex,
    name: header["Opportunity Name *"],
    OpportunityId: opp.opportunityId,
    accountId: opp.accountId,
    contractStartDate: excelDateToYYYYMMDD(header["Contract Start Date"]),
    manualContractEndDate: excelDateToYYYYMMDD(header["Manual Contract End Date"]),
    expirationDate: excelDateToYYYYMMDD(header["Expiration Date"]),
    billingFrequency: header["Billing Frequency"],
    "quoteLines": quoteLines
  }];
  console.log((quoteData));
  return quoteData;
};

  const readExcelData = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet     = context.workbook.worksheets.getItem(sheetName);
        const usedRange = sheet.getUsedRange();
        usedRange.load(["formulas", "values"]); 
        await context.sync();
        if (usedRange.formulas.length > 0) {
          await prepareData(usedRange.formulas, usedRange.values);
        }
      });
    } catch (err) {
      console.error("Error reading Excel data:", err);
    }
  };

  const handlePush = async () => {
    setIsPushing(true);
    setPushState({ status: "Pending", message: "Submitting…" });
    try {
      await SalesforceService.pushDataToSalesforce(async (update) => {
        const recordErrors = update.results?.filter(r => r.status === "error") ?? [];

        if (update.status === "Complete" && recordErrors.length > 0) {
          const uniqueMessages = [...new Set(recordErrors.map(r => r.errorMessage))];
          setPushState({
            ...update,
            status:  "Failed",
            message: `Sync completed with ${recordErrors.length} record error(s): ${uniqueMessages.join(", ")}`,
            recordErrors,
          });
          setIsPushing(false);
          return;
        }

        setPushState(update);

        if (update.status === "Complete") {
          setPushState(prev => ({ ...prev, message: "Writing IDs to Excel…" }));
          await populateSalesforceIds(update);
          setPushState(prev => ({ ...prev, message: "Done" }));
          setIsPushing(false);
        }
        if (update.status === "Failed" || update.status === "RolledBack") {
          setIsPushing(false);
        }
      });
    } catch (err) {
      setPushState(prev => ({ ...prev, status: "Failed", message: err.message }));
      setIsPushing(false);
    }
  };

  if (error) {
    return <div style={{ padding: "20px", color: "red" }}>{error}</div>;
  }

  return (
    <div style={{ padding: "20px" }}>
      <Button onClick={handlePush} disabled={isPushing}>
        {isPushing
          ? "Pushing…"
          : pushState?.status === "Failed"
          || pushState?.status === "Complete"
          || pushState?.status === "RolledBack"
          ? "Retry"
          : "Push"}
      </Button>
      <Progress pushState={pushState} />
    </div>
  );
};

export default SalesforceSync;