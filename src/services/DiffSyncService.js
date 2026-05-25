import SalesforceService from "./SalesforceService";

/**
 * Bulkified update service for diff payloads.
 *
 * Salesforce Composite API allows up to 25 subrequests per call.
 * We chunk all records and fire them in batches, reporting progress
 * after each batch via the onProgress callback.
 *
 * onProgress shape:
 * {
 *   status:   "Pending" | "InProgress" | "Complete" | "Failed",
 *   message:  string,
 *   total:    number,   // total records to update
 *   done:     number,   // records processed so far
 *   errors:   [ { referenceId, objectApiName, id, message } ]
 * }
 */

const BATCH_SIZE = 25; // Salesforce Composite limit

/**
 * Flatten the payload into a single array of update requests.
 * Each item: { objectApiName, id, fields }
 */
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
      records.push({
        objectApiName: item.objectApiName,
        id:            item.id,
        fields:        item.fields,
      });
    }
  });

  (payload.quoteCostItems ?? []).forEach((item) => {
    if (item.id && Object.keys(item.fields).length > 0) {
      records.push({
        objectApiName: item.objectApiName,
        id:            item.id,
        fields:        item.fields,
      });
    }
  });

  return records;
}

/**
 * Build a Composite API request body from a chunk of records.
 * Each subrequest is a PATCH to /sobjects/{objectApiName}/{id}
 */
function buildCompositeBody(chunk) {
  return {
    allOrNone: false, // partial success — collect per-record errors
    compositeRequest: chunk.map((rec, i) => ({
      method:      "PATCH",
      url:         `/services/data/v59.0/sobjects/${rec.objectApiName}/${rec.id}`,
      referenceId: `ref_${i}_${rec.id}`,
      body:        rec.fields,
    })),
  };
}

/**
 * Parse errors out of a Composite API response.
 * Salesforce returns 204 for success and 4xx for errors per subrequest.
 */
function parseCompositeErrors(compositeResults, chunk) {
  const errors = [];
  (compositeResults ?? []).forEach((result, i) => {
    if (result.httpStatusCode !== 204 && result.httpStatusCode !== 200) {
      const rec = chunk[i];
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

/**
 * Main entry point.
 * @param {object}   payload    - built by buildPayload() in diff.js
 * @param {function} onProgress - called after every batch with progress update
 */
export async function pushDiffToSalesforce(payload, onProgress) {
  const records = flattenPayload(payload);
  const total   = records.length;

  if (total === 0) {
    onProgress({ status: "Complete", message: "Nothing to update.", total: 0, done: 0, errors: [] });
    return;
  }

  onProgress({ status: "Pending", message: "Preparing updates…", total, done: 0, errors: [] });

  // Chunk into batches of BATCH_SIZE
  const chunks = [];
  for (let i = 0; i < records.length; i += BATCH_SIZE) {
    chunks.push(records.slice(i, i + BATCH_SIZE));
  }

  let done       = 0;
  const allErrors = [];

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
      console.log(body);
      const response = await SalesforceService.composite(body);
      const errors   = parseCompositeErrors(response.compositeResponse, chunk);
      allErrors.push(...errors);
      done += chunk.length;
    } catch (err) {
      const message = err?.message ?? "Unknown error during batch update";
      onProgress({
        status:  "Failed",
        message: `Batch ${batchIdx + 1} failed: ${message}`,
        total,
        done,
        errors:  allErrors,
      });
      return; // stop processing further batches on hard failure
    }
  }

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