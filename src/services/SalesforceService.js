class SalesforceService {
  static async pushDataToSalesforce(onStatusUpdate) {
    await Office.context.document.settings.refreshAsync();
    const accessToken = Office.context.document.settings.get("salesforce_access_token");
    const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");

    const raw = JSON.stringify([
      {
        rowIndex: 2,
        cell: 'VTA1710!11',
        name: "Quote-001",
        OpportunityId: "0067100000GL5I1AAL",
        status: "Draft",
        accountId: "001xx000003GYn1AAG",
        expirationDate: "2026-12-31",
        quoteLines: [
          {
            rowIndex: 3,
            cell: 'VTA1710!12',
            OpportunityId: "0067100000GL5I1AAL",
            Quantity: 2,
            Product2Id: "01t2s0000012QjJAAU",
            unitPrice: 100.0,
            quoteLineCosts: [{ rowIndex: 4, cell: 'VTA1710!13', costType: "Material", amount: 50.0 }],
          },
        ],
      },
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

  static async pollStatus({ jobKey, accessToken, instanceUrl, onStatusUpdate, interval = 3000, maxAttempts = 40 }) {
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

      if (data.status === "Complete" || data.status === "Failed") return data;
    }

    throw new Error("Polling timed out");
  }
}

export default SalesforceService;