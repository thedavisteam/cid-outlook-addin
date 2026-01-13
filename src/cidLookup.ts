import { CONFIG } from './config';
import { getAccessToken } from './auth';

export interface CidMatch {
  cid: string;
  opportunityName?: string;
  matchedEmails: string[];
}

// MOCK DATA FOR LOCAL TESTING - Only used when running from localhost
const MOCK_DATA: Record<string, CidMatch> = {
  'realty@julieandbryce.com': {
    cid: 'CID-2026-TEST',
    opportunityName: 'Test - Realty Account',
    matchedEmails: ['realty@julieandbryce.com']
  },
  'realty@watervilleaudiology.com': {
    cid: 'CID-2026-0003',
    opportunityName: 'Waterville Audiology - Realty',
    matchedEmails: ['realty@watervilleaudiology.com']
  },
  'jennifer@watervilleaudiology.com': {
    cid: 'CID-2026-0002',
    opportunityName: 'Waterville Audiology',
    matchedEmails: ['jennifer@watervilleaudiology.com']
  },
  'test@example.com': {
    cid: 'CID-2026-0001',
    opportunityName: 'Test Client',
    matchedEmails: ['test@example.com']
  }
};

// Auto-detect environment
function isLocalhost(): boolean {
  return window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
}

function isOutlookContext(): boolean {
  // Check if running inside Office/Outlook add-in context
  try {
    return typeof Office !== 'undefined' && 
           Office.context && 
           Office.context.mailbox && 
           Office.context.mailbox.item !== undefined;
  } catch {
    return false;
  }
}

function shouldUseMockData(): boolean {
  // Use mock data if:
  // 1. Running on localhost (development)
  // 2. Running outside of Outlook (direct browser access to GitHub Pages)
  if (isLocalhost()) return true;
  if (!isOutlookContext()) return true;
  return false;
}

export async function getCidMatches(emails: string[]): Promise<CidMatch[]> {
  const normalized = Array.from(new Set(emails.map((e) => e.trim().toLowerCase()))).filter(Boolean);
  if (!normalized.length) return [];

  // Use mock data for localhost or when not in Outlook context
  if (shouldUseMockData()) {
    const reason = isLocalhost() ? 'localhost' : 'not in Outlook context';
    console.log(`🧪 MOCK MODE (${reason}): Using mock data. Emails:`, normalized);
    const results: CidMatch[] = [];
    for (const email of normalized) {
      const match = MOCK_DATA[email];
      if (match) {
        results.push(match);
      }
    }
    console.log('Mock results:', results);
    return results;
  }

  // Production: Query real SharePoint from within Outlook add-in
  console.log('🚀 PRODUCTION: Querying SharePoint. Emails:', normalized);
  return querySharePoint(normalized);
}

// ============ PRODUCTION CODE BELOW ============

const EMAIL_CHUNK = 10;

function chunk<T>(items: T[], size: number): T[][] {
  const result: T[][] = [];
  for (let i = 0; i < items.length; i += size) {
    result.push(items.slice(i, i + size));
  }
  return result;
}

function encodeValue(value: string): string {
  return value.replace(/'/g, "''");
}

async function querySharePoint(emails: string[]): Promise<CidMatch[]> {
  const batches = chunk(emails, EMAIL_CHUNK);
  const all: CidMatch[] = [];
  for (const batch of batches) {
    const res = await queryChunk(batch);
    all.push(...res);
  }
  return all;
}

async function queryChunk(emails: string[]): Promise<CidMatch[]> {
  if (!CONFIG.siteUrl || CONFIG.siteUrl.startsWith('<')) {
    throw new Error('CONFIG.siteUrl is not configured.');
  }
  
  // Get access token for Microsoft Graph
  const token = await getAccessToken();
  
  if (!token) {
    throw new Error('Failed to get authentication token. Please sign in to Microsoft 365.');
  }

  // Extract site information from URL
  // https://julieandbryce.sharepoint.com -> hostname: julieandbryce.sharepoint.com
  const siteUrl = new URL(CONFIG.siteUrl);
  const hostname = siteUrl.hostname;
  
  // Use Microsoft Graph API instead of SharePoint REST API (avoids CORS issues)
  // Get site ID first
  const siteApiUrl = `https://graph.microsoft.com/v1.0/sites/${hostname}:/`;
  console.log('Getting site info from Graph API:', siteApiUrl);
  
  const siteResponse = await fetch(siteApiUrl, {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json'
    }
  });

  if (!siteResponse.ok) {
    const text = await siteResponse.text();
    console.error('Graph API site error:', siteResponse.status, text);
    throw new Error(`Failed to access SharePoint site via Graph API: ${siteResponse.status}`);
  }

  const siteData = await siteResponse.json();
  const siteId = siteData.id;
  console.log('Site ID:', siteId);

  // Get list ID
  const listApiUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$filter=displayName eq '${CONFIG.contactsListName}'`;
  console.log('Getting list info:', listApiUrl);
  
  const listResponse = await fetch(listApiUrl, {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json'
    }
  });

  if (!listResponse.ok) {
    const text = await listResponse.text();
    console.error('Graph API list error:', listResponse.status, text);
    throw new Error(`Failed to find list '${CONFIG.contactsListName}': ${listResponse.status}`);
  }

  const listData = await listResponse.json();
  if (!listData.value || listData.value.length === 0) {
    throw new Error(`List '${CONFIG.contactsListName}' not found in SharePoint site`);
  }
  
  const listId = listData.value[0].id;
  console.log('List ID:', listId);

  // Query list items with email filter
  // Build filter for emails
  const emailFilters = emails
    .map((e) => `fields/Client_x0020_Email eq '${encodeValue(e)}' or fields/Client_x0020_Email_x0020_2 eq '${encodeValue(e)}'`)
    .join(' or ');
    
  const itemsApiUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=fields&$filter=${emailFilters}`;
  console.log('Querying list items via Graph API');

  const itemsResponse = await fetch(itemsApiUrl, {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json'
    }
  });

  if (!itemsResponse.ok) {
    const text = await itemsResponse.text();
    console.error('Graph API items error:', itemsResponse.status, text);
    if (itemsResponse.status === 401 || itemsResponse.status === 403) {
      throw new Error('Authentication failed. Please ensure you have access to the SharePoint site.');
    }
    throw new Error(`SharePoint query failed: ${itemsResponse.status} - ${text.substring(0, 100)}`);
  }

  const data = await itemsResponse.json();
  console.log('SharePoint response:', data);
  
  const results: CidMatch[] = [];
  const map = new Map<string, CidMatch>();

  for (const item of data.value ?? []) {
    const fields = item.fields;
    const cid = fields?.Title as string | undefined;
    if (!cid) continue;
    const opportunityName = (fields?.KW_x0020_Opportunity_x0020_Name as string | undefined) ?? undefined;
    const email1 = (fields?.Client_x0020_Email as string | undefined)?.toLowerCase();
    const email2 = (fields?.Client_x0020_Email_x0020_2 as string | undefined)?.toLowerCase();
    
    const existing = map.get(cid) ?? { cid, opportunityName, matchedEmails: [] };
    if (email1 && !existing.matchedEmails.includes(email1)) {
      existing.matchedEmails.push(email1);
    }
    if (email2 && !existing.matchedEmails.includes(email2)) {
      existing.matchedEmails.push(email2);
    }
    if (!existing.opportunityName && opportunityName) {
      existing.opportunityName = opportunityName;
    }
    map.set(cid, existing);
  }

  map.forEach((value) => results.push(value));
  return results;
}
