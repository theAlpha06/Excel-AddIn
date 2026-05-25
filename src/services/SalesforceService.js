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
  // Add inside SalesforceService class

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
}



export default SalesforceService;