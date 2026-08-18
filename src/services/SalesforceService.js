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
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    await new Promise((r) => setTimeout(r, interval));

    const res = await fetch(`http://localhost:5000/salesforce/v2/bulk-status/${jobKey}`, {
      headers: {
        "sf-access-token": accessToken,
        "sf-instance-url": instanceUrl,
      },
    });

    if (!res.ok) throw new Error("Poll request failed");

    const raw = await res.json();
    const data = typeof raw === "string" ? JSON.parse(raw) : raw;

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
    console.log("Updating escalation records:", JSON.stringify(records, null, 2));
    await Office.context.document.settings.refreshAsync();

    const accessToken = Office.context.document.settings.get("salesforce_access_token");
    const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");

    const CHUNK_SIZE = 100;
    const failures = [];

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
      console.log(`Updated batch ${Math.floor(i / CHUNK_SIZE) + 1}:`, result);

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