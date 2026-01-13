export const CONFIG = {
  // SharePoint site hosting CID Register / CID Contacts
  siteUrl: 'https://julieandbryce.sharepoint.com',
  // CID Contacts list name
  contactsListName: 'CID Register',
  // Regex pattern to detect an existing CID tag in the subject
  cidTagPattern: /\[CID-\d{4}-\d{4}\]/i,
  // Notification IDs used for info bars
  notifications: {
    noMatch: 'CID_NO_MATCH',
    applied: 'CID_APPLIED'
  },
  // Azure AD App Registration
  // REPLACE THIS WITH YOUR CLIENT ID FROM AZURE PORTAL
  azureClientId: 'f1569983-b73d-4260-840a-e0260d48fa81',
  // Tenant ID (or use 'common' for multi-tenant)
  azureTenantId: 'bd787835-1bb2-4149-a907-9074cfaacb64'
};
