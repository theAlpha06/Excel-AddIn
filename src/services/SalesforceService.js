class SalesforceService {
  static async pushDataToSalesforce(data, onStatusUpdate) {
    await Office.context.document.settings.refreshAsync();
    const accessToken = Office.context.document.settings.get("salesforce_access_token");
    const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");

    onStatusUpdate({ status: "Pending", message: "Submitting records…" });

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
    onStatusUpdate({ status: "Pending", jobKey, message });

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

  static async pollStatus({ jobKey, accessToken, instanceUrl, onStatusUpdate, interval = 3000, maxAttempts = 100 }) {
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

      onStatusUpdate({ ...data, jobKey });

      if (data.status === "Complete" || data.status === "Failed") {
        return data;
      }

      if (data.status === "RolledBack") {
        return data;
      }
    }
    throw new Error("Polling timed out");
  }

  static async query(soqlQuery) {
    const accessToken = Office.context.document.settings.get("salesforce_access_token");
    const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");
    console.log(accessToken);
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

  /**
   * PATCH a single Salesforce record via the local proxy.
   *
   * The proxy is expected to forward this to:
   *   PATCH {instanceUrl}/services/data/v58.0/sobjects/{objectName}/{recordId}
   *
   * @param {string} objectName  - Salesforce object API name, e.g. "SBQQ__Quote__c"
   * @param {string} recordId    - 18-char Salesforce record ID
   * @param {object} fields      - Plain object of { sfApiFieldName: value } to update
   */
  static async updateRecord(objectName, recordId, fields) {
    const accessToken = Office.context.document.settings.get("salesforce_access_token");
    const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");

    const response = await fetch(
      `http://localhost:5000/salesforce/services/data/v58.0/sobjects/${objectName}/${recordId}`,
      {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json",
          "sf-access-token": accessToken,
          "sf-instance-url": instanceUrl,
        },
        body: JSON.stringify(fields),
      }
    );

    // Salesforce returns 204 No Content on a successful PATCH – treat that as OK too
    if (!response.ok && response.status !== 204) {
      let errorMessage = `Failed to update ${objectName} ${recordId}: HTTP ${response.status}`;
      try {
        const errorData = await response.json();
        errorMessage = errorData?.[0]?.message || errorData?.message || errorMessage;
      } catch {
        // response body wasn't JSON – keep the generic message
      }
      throw new Error(errorMessage);
    }

    // 204 has no body; return an empty object so callers don't blow up on .json()
    if (response.status === 204) return {};
    return response.json();
  }
}

export default SalesforceService;