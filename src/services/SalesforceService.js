// Excel hosts the task pane in a webview, and a hidden webview gets its timers throttled hard
// — Chromium drops background timers to roughly one a minute and throttles further after a few
// minutes, and the host can suspend the pane outright. A lone setTimeout(interval) therefore
// stalls the poll loop for as long as the user is away from the pane. Waiting in short ticks
// against a wall-clock deadline fixes that: however long a tick is delayed, the next one sees
// the deadline has already passed and polls immediately.
const POLL_TICK_MS = 1000;

// Coming back to the pane is exactly when the status is most likely stale, so a wake-up polls
// straight away rather than waiting out the rest of a throttled interval. The floor keeps a
// burst of focus events from turning into a burst of requests.
const WAKE_MIN_GAP_MS = 3000;

// A poll that fails because the pane was suspended, the machine slept, or the local service
// blipped must not end the job — it is still running server side and the next tick can pick it
// up. Only an unbroken run of failures is treated as the job being unreachable.
const MAX_CONSECUTIVE_POLL_ERRORS = 5;

function waitForNextPoll(deadline) {
  return new Promise((resolve) => {
    const startedAt = Date.now();
    let timer = null;
    let settled = false;

    const finish = () => {
      if (settled) return;
      settled = true;
      if (timer !== null) clearTimeout(timer);
      document.removeEventListener("visibilitychange", onWake);
      window.removeEventListener("focus", onWake);
      resolve();
    };

    const tick = () => {
      const remaining = deadline - Date.now();
      if (remaining <= 0) return finish();
      timer = setTimeout(tick, Math.min(POLL_TICK_MS, remaining));
    };

    const onWake = () => {
      if (document.visibilityState === "hidden") return;
      if (Date.now() - startedAt < WAKE_MIN_GAP_MS) return;
      finish();
    };

    document.addEventListener("visibilitychange", onWake);
    window.addEventListener("focus", onWake);
    tick();
  });
}

class SalesforceService {
  static async pushDataToSalesforce(data, onStatusUpdate) {
    await Office.context.document.settings.refreshAsync();
    const accessToken = Office.context.document.settings.get("salesforce_access_token");
    const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");

    await onStatusUpdate({ status: "Pending", message: "Submitting records…" });
    const pushRes = await fetch("http://localhost:5000/salesforce/v2/pushRecords", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "sf-access-token": accessToken,
        "sf-instance-url": instanceUrl,
      },
      body: JSON.stringify(data),
    });

    if (!pushRes.ok) throw new Error("Push request failed");

    const { jobKey, message } = await pushRes.json();
    await onStatusUpdate({ status: "Pending", jobKey, message });    

    return SalesforceService.pollStatus({ jobKey, accessToken, instanceUrl, onStatusUpdate });
  }

  static async checkExistingOpportunity({ opportunityName, accessToken, instanceUrl }) {
    const res = await fetch(`http://localhost:5000/salesforce/v2/checkOpportunity/${opportunityName}`, {
      headers: {
        "sf-access-token": accessToken,
        "sf-instance-url": instanceUrl,
      },
    });
    if (!res.ok) throw new Error("Check request failed");
    const { exists } = await res.json();
    return exists;
  }


static async pollStatus({ jobKey, accessToken, instanceUrl, onStatusUpdate, interval = 20000, maxAttempts = 3000 }) {
  let consecutiveErrors = 0;

  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    await waitForNextPoll(Date.now() + interval);

    let data;
    try {
      const res = await fetch(`http://localhost:5000/salesforce/v2/bulk-status/${jobKey}`, {
        headers: {
          "sf-access-token": accessToken,
          "sf-instance-url": instanceUrl,
        },
      });

      if (!res.ok) throw new Error(`Poll request failed: ${res.status}`);

      const raw = await res.json();
      data = typeof raw === "string" ? JSON.parse(raw) : raw;
    } catch (err) {
      // Throwing here used to abandon the job for good: the error propagated out of
      // pushDataToSalesforce and nothing ever polled again, so a single blip while the user
      // was away looked like the push had silently stopped.
      consecutiveErrors++;
      console.warn(
        `Status poll ${attempt + 1} failed (${consecutiveErrors}/${MAX_CONSECUTIVE_POLL_ERRORS}):`,
        err?.message ?? err
      );
      if (consecutiveErrors >= MAX_CONSECUTIVE_POLL_ERRORS) throw err;
      continue;
    }

    consecutiveErrors = 0;
    await onStatusUpdate({ ...data, jobKey });

    if (data.status === "Complete" || data.status === "Failed") {
      return data;
    }

    if (data.status === "RolledBack") {
      // setTimeout(() => {
      //   onStatusUpdate({ status: null });
      // }, 5000);
      return data;
    }
  }
  throw new Error("Polling timed out");
}

  static async query(soqlQuery) {
    const accessToken = Office.context.document.settings.get("salesforce_access_token");
    const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");
    const response = await fetch("http://localhost:5000/salesforce/services/data/v58.0/query", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "sf-access-token": accessToken,
        "sf-instance-url": instanceUrl,
      },
      body: JSON.stringify({ q: soqlQuery }),
    });
    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.message || `SOQL query failed: ${response.status}`);
    }
    return response.json();
  }

  static async createRecords(objectName, records) {
    const accessToken = Office.context.document.settings.get("salesforce_access_token");
    const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");

    const response = await fetch("http://localhost:5000/salesforce/create", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "sf-access-token": accessToken,
        "sf-instance-url": instanceUrl,
      },
      body: JSON.stringify({ objectName, records }),
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error[0].message || `Failed to create records: ${response.status}`);
    }

    return response.json();
  }

  static async composite(body) {
    const accessToken = Office.context.document.settings.get("salesforce_access_token");
    const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");

    const response = await fetch(`http://localhost:5000/salesforce/composite`, {
      method: "POST",
      headers: {
        "Content-Type":    "application/json",
        "sf-access-token": accessToken,
        "sf-instance-url": instanceUrl,
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err?.message ?? `Composite API failed: ${response.status}`);
    }

    return response.json();
  }

  static async uploadToSalesforce (fileName, base64Data, parentId) {
    const accessToken = Office.context.document.settings.get("salesforce_access_token");
    const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");

    const response = await fetch("http://localhost:5000/salesforce/uploadWorkbook", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "sf-access-token": accessToken,
        "sf-instance-url": instanceUrl,
      },
      body: JSON.stringify({
        fileName: fileName,
        base64Data,
        parentId: parentId,
      }),
    });
    if (!response.ok) {
      const errorData = await response.json();                                                    
      throw new Error(errorData.message || `Upload failed: ${response.status}`);
    }
    return response.json();
  }

  static async updateRecord(objectName, recordId, fields) {
    const accessToken = Office.context.document.settings.get("salesforce_access_token");
    const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");
    const response = await fetch("http://localhost:5000/salesforce/update", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "sf-access-token": accessToken,
        "sf-instance-url": instanceUrl,
      },
      body: JSON.stringify({ objectName, recordId, fields }),
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      console.error(`Failed to update record ${objectName}/${recordId}:`, errorData);
      throw new Error(errorData.message || `Failed to update record: ${response.status}`);
    }

    return response.json();
  }

  static async updateEscalationRecords(records) {
    await Office.context.document.settings.refreshAsync();

    const accessToken = Office.context.document.settings.get("salesforce_access_token");
    const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");

    const CHUNK_SIZE = 500;
    const failures = [];
    const batchCount = Math.ceil(records.length / CHUNK_SIZE);

    // Pretty-printing ten thousand records built a multi-megabyte string on the UI thread before
    // a single request went out. The array itself is still logged — the console expands it
    // lazily — but only while it is small enough for that to be readable.
    console.log(`Updating ${records.length} escalation record(s) in ${batchCount} batch(es).`);
    if (records.length <= 25) console.log(records);

    for (let i = 0; i < records.length; i += CHUNK_SIZE) {
      const chunk = records.slice(i, i + CHUNK_SIZE);

      const response = await fetch(
        "http://localhost:5000/salesforce/updateEscalations",
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "sf-access-token": accessToken,
            "sf-instance-url": instanceUrl,
          },
          body: JSON.stringify({
            records: chunk,
          }),
        }
      );

      if (!response.ok) {
        const error = await response.json().catch(() => ({}));
        throw new Error(error.message || "Failed to update records");
      }

      const result = await response.json();
      console.log(`Updated batch ${Math.floor(i / CHUNK_SIZE) + 1} of ${batchCount}.`, result);

      // The route sends allOrNone:false, so Salesforce reports rejections per record and
      // still returns 200. Without this check a failed update is indistinguishable from a
      // successful one.
      (Array.isArray(result) ? result : []).forEach((r, idx) => {
        if (r && r.success === false) {
          const record = chunk[idx];
          console.error(`Escalation update rejected for ${record?.Id}:`, r.errors);
          failures.push({ id: record?.Id, fields: record, errors: r.errors });
        }
      });
    }

    return { failures };
  }
}

export default SalesforceService;