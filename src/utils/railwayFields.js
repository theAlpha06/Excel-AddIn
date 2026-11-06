import { isRailwayLob, requiresSpecialContractType } from "./lob";
import { resolveActiveUserId } from "./kamLookup";

// Quote fields that are only sent when the LOB is Railways.
//   "Service Area Manager"  → Quote.AreaServiceManager__c  (User lookup, no lookup filter; the
//                              sheet holds the person's email)
//   "Special Contract Type" → Quote.Special_Contract_Type__c (restricted picklist)
// The sheet's label says "Service Area Manager" while the field is "Area Service Manager", so
// both word orders are accepted.
const AREA_SERVICE_MANAGER_LABEL = /^(service\s+area|area\s+service)\s+manager\b/i;
const SPECIAL_CONTRACT_TYPE_LABEL = /^special\s+contract\s+type\b/i;

export const isAreaServiceManagerLabel = (label) =>
  AREA_SERVICE_MANAGER_LABEL.test(String(label ?? "").trim());
export const isSpecialContractTypeLabel = (label) =>
  SPECIAL_CONTRACT_TYPE_LABEL.test(String(label ?? "").trim());

// Active values of Quote.Special_Contract_Type__c in the org. The picklist is restricted, so a
// value typed slightly differently ("rcf contract", "RCF-Non RCF combination") is mapped onto the
// exact one, and anything else is reported instead of sent.
const SPECIAL_CONTRACT_TYPES = ["RCF Contract", "Non-RCF Contract", "RCF - Non RCF Combination"];
const contractTypeKey = (value) => String(value ?? "").toLowerCase().replace(/[\s\-–—]+/g, "");

export function matchSpecialContractType(value) {
  const key = contractTypeKey(value);
  if (!key) return null;
  return SPECIAL_CONTRACT_TYPES.find((option) => contractTypeKey(option) === key) ?? null;
}

const findByLabel = (header, test) => {
  const key = Object.keys(header ?? {}).find(test);
  return key ? String(header[key] ?? "").trim() : "";
};

// Values from SalesforceSync's header object (label → value).
export const readRailwayHeader = (header) => ({
  lob: findByLabel(header, (label) => /^lob$/i.test(String(label).trim())),
  areaServiceManagerEmail: findByLabel(header, isAreaServiceManagerLabel),
  specialContractType: findByLabel(header, isSpecialContractTypeLabel),
});

// Resolves the raw sheet values into Quote fields. Values that cannot be used are left out and
// described in `notes`, so they never make Salesforce reject the rest of the Quote update.
//  - Area Service Manager: Railways only.
//  - Special Contract Type: Railways and Oil & Gas — the org's Opportunity rule makes it
//    mandatory for both LOBs and forbids it for every other.
// Returns { fields, notes } — both empty for any other LOB.
export async function resolveRailwayQuoteFields({ lob, areaServiceManagerEmail, specialContractType }) {
  const fields = {};
  const notes = [];

  if (isRailwayLob(lob) && String(areaServiceManagerEmail ?? "").trim()) {
    const asm = await resolveActiveUserId(areaServiceManagerEmail, "Area Service Manager");
    if (asm.userId) fields.AreaServiceManager__c = asm.userId;
    else notes.push(`Area Service Manager not set: ${asm.reason}`);
  }

  if (requiresSpecialContractType(lob) && String(specialContractType ?? "").trim()) {
    const match = matchSpecialContractType(specialContractType);
    if (match) fields.Special_Contract_Type__c = match;
    else {
      notes.push(
        `Special Contract Type not set: "${specialContractType}" is not one of ${SPECIAL_CONTRACT_TYPES.join(", ")}`
      );
    }
  }

  return { fields, notes };
}
