import SalesforceService from "./SalesforceService";

const BATCH_SIZE = 25;

// ── Flatten standard records ──────────────────────────────────────────────────

function flattenPayload(payload) {
  const records = [];

  if (payload.quoteHeader?.id && Object.keys(payload.quoteHeader.fields).length > 0) {
    records.push({
      objectApiName: payload.quoteHeader.objectApiName,
      id:            payload.quoteHeader.id,
      fields:        payload.quoteHeader.fields,
    });
  }

  (payload.quoteLineItems ?? []).forEach((item) => {
    if (item.id && Object.keys(item.fields).length > 0) {
      records.push({ objectApiName: item.objectApiName, id: item.id, fields: item.fields });
    }
  });

  (payload.quoteCostItems ?? []).forEach((item) => {
    if (item.id && Object.keys(item.fields).length > 0) {
      records.push({ objectApiName: item.objectApiName, id: item.id, fields: item.fields });
    }
  });

  return records;
}

// ── Resolve escalation records from Salesforce ────────────────────────────────

/**
 * For each escalation entry, query SF to get the actual QuoteCostItem__c ID
 * using ParentQuoteCostLine__c + YearNumber__c, then return as flat update records.
 *
 * If only the 2nd year column was edited → constant escalation →
 * apply the same value to ALL subsequent year records under that parent.
 *
 * If multiple year columns were edited → apply each year's value individually.
 */
async function resolveEscalationRecords(escalations) {
  if (!escalations || escalations.length === 0) return [];

  const resolved = [];

  for (const { parentId, years } of escalations) {
    const isConstant = years.length === 1 && years[0].yearNumber === 2;

    if (isConstant) {
      // Single 2nd-year value → fetch ALL child records for this parent and apply to all
      const result = await SalesforceService.query(
        `SELECT Id, YearNumber__c FROM QuoteCostItem__c 
         WHERE ParentQuoteCostLine__c = '${parentId}' 
         ORDER BY YearNumber__c ASC`
      );

      const { fields } = years[0];

      (result.records ?? []).forEach((rec) => {
        resolved.push({
          objectApiName: "QuoteCostItem__c",
          id:            rec.Id,
          fields,
        });
      });

      console.log(
        `✅ Constant escalation for parent ${parentId} → applied to ${result.records?.length ?? 0} year records`
      );
    } else {
      // Variable escalation → query each year individually
      for (const { yearNumber, fields } of years) {
        const result = await SalesforceService.query(
          `SELECT Id, SaleEscalation__c, CostEscalation__c FROM QuoteCostItem__c 
           WHERE ParentQuoteCostLine__c = '${parentId}' 
           AND YearNumber__c = ${yearNumber} 
           LIMIT 1`
        );

        if (!result.records?.length) {
          console.warn(`No QuoteCostItem__c found for parent=${parentId}, year=${yearNumber}`);
          continue;
        }

        resolved.push({
          objectApiName: "QuoteCostItem__c",
          id:            result.records[0].Id,
          fields,
        });

        console.log(
          `✅ Variable escalation for parent=${parentId}, year=${yearNumber} → ${result.records[0].Id}`
        );
      }
    }
  }
 console.log(resolved);
  return resolved;
}

// ── Composite API helpers ─────────────────────────────────────────────────────

function buildCompositeBody(chunk) {
  return {
    allOrNone: false,
    compositeRequest: chunk.map((rec, i) => ({
      method:      "PATCH",
      url:         `/services/data/v59.0/sobjects/${rec.objectApiName}/${rec.id}`,
      referenceId: `ref_${i}_${rec.id}`,
      body:        rec.fields,
    })),
  };
}

function parseCompositeErrors(compositeResults, chunk) {
  const errors = [];
  (compositeResults ?? []).forEach((result, i) => {
    if (result.httpStatusCode !== 204 && result.httpStatusCode !== 200) {
      const rec  = chunk[i];
      const body = result.body;
      const msgs = Array.isArray(body)
        ? body.map((e) => `${e.errorCode}: ${e.message}`).join("; ")
        : JSON.stringify(body);
      errors.push({
        referenceId:   result.referenceId,
        objectApiName: rec.objectApiName,
        id:            rec.id,
        message:       msgs,
      });
    }
  });
  return errors;
}

async function sendInBatches(records, onProgress, total, doneOffset, allErrors) {
  const chunks = [];
  for (let i = 0; i < records.length; i += BATCH_SIZE) {
    chunks.push(records.slice(i, i + BATCH_SIZE));
  }

  let done = doneOffset;

  for (let batchIdx = 0; batchIdx < chunks.length; batchIdx++) {
    const chunk = chunks[batchIdx];

    onProgress({
      status:  "InProgress",
      message: `Updating batch ${batchIdx + 1} of ${chunks.length}…`,
      total,
      done,
      errors:  allErrors,
    });

    try {
      const body     = buildCompositeBody(chunk);
      const response = await SalesforceService.composite(body);
      const errors   = parseCompositeErrors(response.compositeResponse, chunk);
      allErrors.push(...errors);
      done += chunk.length;
    } catch (err) {
      const message = err?.message ?? "Unknown error";
      onProgress({
        status:  "Failed",
        message: `Batch ${batchIdx + 1} failed: ${message}`,
        total,
        done,
        errors:  allErrors,
      });
      return { done, failed: true };
    }
  }

  return { done, failed: false };
}

// ── Main entry point ──────────────────────────────────────────────────────────

export async function pushDiffToSalesforce(payload, onProgress) {
  const allErrors = [];

  // Step 1 — flatten standard records
  const standardRecords = flattenPayload(payload);

  // Step 2 — resolve escalation records (requires SF queries)
  onProgress({
    status:  "Pending",
    message: "Resolving escalation records…",
    total:   0,
    done:    0,
    errors:  [],
  });

  let escalationRecords = [];
  try {
    escalationRecords = await resolveEscalationRecords(payload.escalations ?? []);
  } catch (err) {
    onProgress({
      status:  "Failed",
      message: `Failed to resolve escalations: ${err.message}`,
      total:   0,
      done:    0,
      errors:  [],
    });
    return;
  }

  const allRecords = [...standardRecords, ...escalationRecords];
  const total      = allRecords.length;

  if (total === 0) {
    onProgress({ status: "Complete", message: "Nothing to update.", total: 0, done: 0, errors: [] });
    return;
  }

  onProgress({ status: "Pending", message: "Preparing updates…", total, done: 0, errors: [] });

  // Step 3 — send standard records
  let { done, failed } = await sendInBatches(
    standardRecords,
    onProgress,
    total,
    0,
    allErrors
  );
  if (failed) return;

  // Step 4 — send escalation records
  if (escalationRecords.length > 0) {
    onProgress({
      status:  "InProgress",
      message: "Updating escalation records…",
      total,
      done,
      errors:  allErrors,
    });

    ({ done, failed } = await sendInBatches(
      escalationRecords,
      onProgress,
      total,
      done,
      allErrors
    ));
    if (failed) return;
  }

  // Step 5 — final status
  if (allErrors.length > 0) {
    onProgress({
      status:  "Failed",
      message: `Completed with ${allErrors.length} error(s).`,
      total,
      done,
      errors:  allErrors,
    });
  } else {
    onProgress({
      status:  "Complete",
      message: `Successfully updated ${done} record(s).`,
      total,
      done,
      errors:  [],
    });
  }
}