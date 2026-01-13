/// <reference types="office-js" />
import { PublicClientApplication, AccountInfo, AuthenticationResult } from '@azure/msal-browser';
import { CONFIG } from './config';

/**
 * Authentication for SharePoint API calls from Outlook add-in.
 * 
 * Uses MSAL (Microsoft Authentication Library) for Azure AD authentication.
 * This allows the add-in to work from both localhost and GitHub Pages.
 */

let msalInstance: PublicClientApplication | null = null;
let currentAccount: AccountInfo | null = null;

/**
 * Initialize MSAL instance
 */
function getMsalInstance(): PublicClientApplication {
  if (!msalInstance) {
    msalInstance = new PublicClientApplication({
      auth: {
        clientId: CONFIG.azureClientId,
        authority: `https://login.microsoftonline.com/${CONFIG.azureTenantId}`,
        redirectUri: window.location.origin + window.location.pathname,
      },
      cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: false,
      },
    });
  }
  return msalInstance;
}

/**
 * Get an access token for SharePoint/Microsoft Graph API calls.
 */
export async function getAccessToken(): Promise<string | null> {
  // Check if client ID is configured
  if (!CONFIG.azureClientId || CONFIG.azureClientId === 'YOUR_CLIENT_ID_HERE') {
    console.warn('⚠️ Azure Client ID not configured. Using fallback auth.');
    return null;
  }

  try {
    const msal = getMsalInstance();
    
    // Initialize MSAL
    await msal.initialize();

    // Check for existing account
    const accounts = msal.getAllAccounts();
    if (accounts.length > 0) {
      currentAccount = accounts[0];
    }

    // Try silent token acquisition first
    try {
      const tokenRequest = {
        scopes: ['https://graph.microsoft.com/.default'],
        account: currentAccount || undefined,
      };

      const response: AuthenticationResult = await msal.acquireTokenSilent(tokenRequest);
      console.log('✅ Token acquired silently');
      currentAccount = response.account;
      return response.accessToken;
    } catch (error) {
      console.log('Silent token acquisition failed, trying popup...', error);
      
      // Fall back to interactive login
      try {
        const response: AuthenticationResult = await msal.acquireTokenPopup({
          scopes: ['https://graph.microsoft.com/.default'],
        });
        console.log('✅ Token acquired via popup');
        currentAccount = response.account;
        return response.accessToken;
      } catch (popupError) {
        console.error('❌ Token acquisition failed:', popupError);
        throw new Error('Authentication failed. Please sign in to Microsoft 365.');
      }
    }
  } catch (error) {
    console.error('MSAL initialization failed:', error);
    return null;
  }
}

/**
 * Sign out the current user
 */
export async function signOut(): Promise<void> {
  if (msalInstance && currentAccount) {
    await msalInstance.logoutPopup({
      account: currentAccount,
    });
    currentAccount = null;
  }
}
