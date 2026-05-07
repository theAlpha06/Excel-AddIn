class SalesforceService {
  static async pushDataToSalesforce(onStatusUpdate) {
    await Office.context.document.settings.refreshAsync();
    const accessToken = Office.context.document.settings.get("salesforce_access_token");
    const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");

  const raw = JSON.stringify([
  {
    "rowIndex": 5,
    "name": "Opportunity Test",
    "OpportunityId": "0067100000K3ix3AAB",
    "accountId": "0017100000R8QKYAA3",
    "contractStartDate": "2026-04-01",
    "manualContractEndDate": "2028-03-31",
    "expirationDate": "2028-03-01",
    "billingFrequency": "Quarterly",
    "quoteLines": [
      {
        "Quantity": 6,
        "Cost Price": 51670.16,
        "cell": "Final Quote Header & Lines!23",
        "rowIndex": 23,
        "Product2Id": "01t2s0000012QatAAE",
        "OpportunityId": "0067100000K3ix3AAB",
        "quoteLineCosts": [
          {
            "cell": "Lines Template!7",
            "rowIndex": 7,
            "costType": "Material",
            "amount": 50014.16,
            "part": 3238515,
            "partDescription": "Aircleaner Kit"
          },
          {
            "cell": "Lines Template!8",
            "rowIndex": 8,
            "costType": "Material",
            "amount": 1656,
            "part": "Service Engineer Charges for 70 - 82.5 KVA CPCB II",
            "partDescription": "Service Engineer Charges for 70 - 82.5 KVA CPCB II"
          }
        ]
      },
      {
        "Quantity": 2,
        "Cost Price": 10000,
        "cell": "Final Quote Header & Lines!24",
        "rowIndex": 24,
        "Product2Id": "01t2s0000012QatAAE",
        "OpportunityId": "0067100000K3ix3AAB"
      }
    ]
  }
]);

    onStatusUpdate({ status: "Pending", message: "Submitting records…" });

    const pushRes = await fetch("http://localhost:5000/salesforce/v2/pushRecords", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "sf-access-token": accessToken,
        "sf-instance-url": instanceUrl,
      },
      body: raw,
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
      throw new Error(errorData.message || `Failed to create records: ${response.status}`);
    }

    return response.json();
  }
}

export default SalesforceService;