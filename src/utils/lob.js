// Sub LOB is not sent to Salesforce when the LOB is Railway. The org's picklist value is
// "Railways"; the sheet may say either, in any case.
export const isRailwayLob = (lob) => /^railways?$/i.test(String(lob ?? "").trim());

// The org's Opportunity validation rule: Special Contract Type is mandatory when the LOB is
// Railways or Oil & Gas, and must be blank for every other LOB.
export const requiresSpecialContractType = (lob) =>
  isRailwayLob(lob) || /^oil\s*(&|and)\s*gas$/i.test(String(lob ?? "").trim());
