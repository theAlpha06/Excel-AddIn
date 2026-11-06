import SalesforceService from "../services/SalesforceService";

// The header label on "Final Quote Header & Lines". Matched loosely because Excel labels carry
// stray spacing, dash variants and the odd trailing asterisk.
const KAM_LABEL_PATTERN = /^kam\s*[-–—]\s*account\b/i;

export const isKamLabel = (label) => KAM_LABEL_PATTERN.test(String(label ?? "").trim());

// Value of the KAM - Account cell from a header object keyed by label (SalesforceSync's
// headerData), or "" when the sheet has no such label.
export const findKamEmail = (header) => {
  const key = Object.keys(header ?? {}).find(isKamLabel);
  return key ? String(header[key] ?? "").trim() : "";
};

// Quote.KAM__c is a lookup to User with a mandatory lookup filter: User.UserType__c must equal
// "KAM". Any other user is rejected by Salesforce with FIELD_FILTER_VALIDATION_EXCEPTION, and since
// KAM__c travels in the same Quote update as the other header fields, that rejection would also
// drop those. So the filter is checked here first and a non-KAM user is reported, not sent.
const KAM_USER_TYPE = "KAM";

// Active users whose Email matches (SOQL compares it case-insensitively).
async function findActiveUsersByEmail(email) {
  const escaped = email.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
  const result = await SalesforceService.query(
    `SELECT Id, Name, UserType__c FROM User WHERE Email = '${escaped}' AND IsActive = true`
  );
  return result?.records ?? [];
}

// Any active user with this email — for User lookups without a filter, such as
// Quote.AreaServiceManager__c. Returns { userId, name } or { userId: null, reason }.
export async function resolveActiveUserId(email, fieldLabel = "User") {
  const value = String(email ?? "").trim();
  if (!value) return { userId: null, reason: `No ${fieldLabel} email entered` };

  const users = await findActiveUsersByEmail(value);
  if (users.length === 0) {
    return { userId: null, reason: `No active Salesforce user found with email ${value}` };
  }
  if (users.length > 1) {
    console.warn(`More than one active user has email ${value}; using ${users[0].Name} (${users[0].Id}).`);
  }
  return { userId: users[0].Id, name: users[0].Name };
}

// The email typed in the sheet is matched against an active User's Email. Returns
// { userId, name } when a KAM user is found, { userId: null, reason } otherwise.
export async function resolveKamUserId(email) {
  const value = String(email ?? "").trim();
  if (!value) return { userId: null, reason: "No KAM email entered" };

  const users = await findActiveUsersByEmail(value);

  if (users.length === 0) {
    return { userId: null, reason: `No active Salesforce user found with email ${value}` };
  }

  const kamUsers = users.filter((u) => u.UserType__c === KAM_USER_TYPE);
  if (kamUsers.length === 0) {
    const names = users.map((u) => `${u.Name} (User Type: ${u.UserType__c ?? "blank"})`).join(", ");
    return {
      userId: null,
      reason: `${names} is not a KAM user — KAM must have User Type = ${KAM_USER_TYPE}`,
    };
  }
  if (kamUsers.length > 1) {
    console.warn(
      `More than one active KAM user has email ${value}; using ${kamUsers[0].Name} (${kamUsers[0].Id}).`
    );
  }
  return { userId: kamUsers[0].Id, name: kamUsers[0].Name };
}
