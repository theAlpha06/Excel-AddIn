// Sub LOB is not sent to Salesforce when the LOB is Railway. The org's picklist value is
// "Railways"; the sheet may say either, in any case.
export const isRailwayLob = (lob) => /^railways?$/i.test(String(lob ?? "").trim());
