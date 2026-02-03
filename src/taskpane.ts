/// <reference types="office-js" />
import { getCidMatches } from './cidLookup';
import { CONFIG } from './config';

// VERSION TRACKING - Update this with each build
const APP_VERSION = '3.1.1';
const BUILD_TIMESTAMP = '__BUILD_TIME__'; // Will be replaced by webpack

Office.onReady((info) => {
  console.log(`[CID App v${APP_VERSION}] Taskpane loaded`);
  console.log(`[CID App] Host: ${info.host}, Platform: ${info.platform}`);
  
  // Display debug info
  const debugInfo = document.getElementById('debugInfo');
  if (debugInfo) {
    const platform = info.platform || 'Unknown';
    const host = info.host || 'Unknown';
    debugInfo.innerHTML = `Platform: <b>${platform}</b> | Host: <b>${host}</b> | Build: <b>${BUILD_TIMESTAMP}</b>`;
  }
  
  const lookupBtn = document.getElementById('lookupBtn');
  const applyBtn = document.getElementById('applyBtn');
  const addToRegisterBtn = document.getElementById('addToRegisterBtn');
  const openRegisterListBtn = document.getElementById('openRegisterListBtn');
  const emailInput = document.getElementById('emailInput') as HTMLInputElement;
  const resultDiv = document.getElementById('result');

  function openExternal(url: string) {
    const popup = window.open(url, '_blank', 'noopener,noreferrer');
    if (!popup && resultDiv) {
      showResult('Popup blocked. Allow popups for Outlook/add-ins, then try again.', 'error');
    }
  }

  // Open SharePoint "New item" form for CID Register
  if (addToRegisterBtn) {
    addToRegisterBtn.addEventListener('click', () => {
      openExternal('https://julieandbryce.sharepoint.com/Lists/CID%20Register/NewForm.aspx');
    });
  }

  // Open SharePoint list view for quick search/edit
  if (openRegisterListBtn) {
    openRegisterListBtn.addEventListener('click', () => {
      openExternal('https://julieandbryce.sharepoint.com/Lists/CID%20Register/AllItems.aspx');
    });
  }

  // Manual lookup button
  if (lookupBtn && emailInput && resultDiv) {
    lookupBtn.addEventListener('click', async () => {
      const email = emailInput.value.trim();
      
      if (!email) {
        showResult('Please enter an email address', 'error');
        return;
      }

      showResult('Looking up CID...', 'info');
      
      try {
        const matches = await getCidMatches([email]);
        
        if (matches.length > 0) {
          const match = matches[0];
          const oppName = match.opportunityName ? ` (${match.opportunityName})` : '';
          showResult(`✓ Found: ${match.cid}${oppName}`, 'success');
        } else {
          showResult(`❌ No CID found for ${email}`, 'error');
        }
      } catch (error) {
        showResult(`❌ Error: ${(error as Error).message}`, 'error');
        console.error('Lookup error:', error);
      }
    });
  }

  // Auto-apply button - reads recipients and tags subject
  if (applyBtn && resultDiv) {
    applyBtn.addEventListener('click', async () => {
      showResult('Reading recipients...', 'info');
      
      try {
        // Check if we're in Outlook context
        const inOutlook = isInOutlookContext();
        let emails: string[] = [];
        
        if (inOutlook) {
          // Get recipients from current email
          emails = await getRecipientEmails();
        } else {
          // Fallback: use email from input field for testing outside Outlook
          const testEmail = emailInput?.value.trim();
          if (testEmail) {
            emails = [testEmail.toLowerCase()];
            console.log('🧪 TEST MODE: Using email from input field:', testEmail);
          }
        }
        
        if (emails.length === 0) {
          const msg = inOutlook 
            ? '❌ No recipients found. Add a To or CC recipient first.'
            : '❌ Not in Outlook. Enter an email above to test.';
          showResult(msg, 'error');
          return;
        }

        showResult(`Looking up CID for: ${emails.join(', ')}...`, 'info');
        
        // Look up CID
        const matches = await getCidMatches(emails);
        
        if (matches.length === 0) {
          showResult(`❌ No CID found for recipients`, 'error');
          return;
        }

        const match = matches[0];
        
        if (inOutlook) {
          // Get current subject
          const currentSubject = await getSubject();
          
          // Check if already tagged
          if (CONFIG.cidTagPattern.test(currentSubject)) {
            showResult(`Subject already has a CID tag: ${currentSubject}`, 'info');
            return;
          }

          // Apply tag to END of subject: "Subject - [CID-xxxx-xxxx]"
          const subjectText = currentSubject.trim() || 'New Email';
          const newSubject = `${subjectText} - [${match.cid}]`;
          await setSubject(newSubject);
          
          const oppName = match.opportunityName ? ` (${match.opportunityName})` : '';
          showResult(`✓ Applied: [${match.cid}]${oppName}`, 'success');
        } else {
          // Test mode - just show what would happen
          const oppName = match.opportunityName ? ` (${match.opportunityName})` : '';
          showResult(`🧪 TEST: Would apply [${match.cid}]${oppName} to subject`, 'success');
        }
        
      } catch (error) {
        showResult(`❌ Error: ${(error as Error).message}`, 'error');
        console.error('Apply error:', error);
      }
    });
  }
  
  // Helper to check if running inside Outlook
  function isInOutlookContext(): boolean {
    try {
      return typeof Office !== 'undefined' && 
             Office.context && 
             Office.context.mailbox && 
             Office.context.mailbox.item !== undefined;
    } catch {
      return false;
    }
  }

  function showResult(message: string, type: 'success' | 'error' | 'info') {
    if (resultDiv) {
      resultDiv.textContent = message;
      resultDiv.className = type;
      resultDiv.style.display = 'block';
    }
  }
});

// Helper functions to interact with Outlook
function getRecipientEmails(): Promise<string[]> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) {
      reject(new Error('No mail item available'));
      return;
    }

    const emails: string[] = [];
    let pending = 2;

    const checkDone = () => {
      pending--;
      if (pending === 0) {
        resolve(emails);
      }
    };

    // Get To recipients
    item.to.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        for (const recipient of result.value) {
          if (recipient.emailAddress) {
            emails.push(recipient.emailAddress.toLowerCase());
          }
        }
      }
      checkDone();
    });

    // Get CC recipients
    item.cc.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        for (const recipient of result.value) {
          if (recipient.emailAddress) {
            emails.push(recipient.emailAddress.toLowerCase());
          }
        }
      }
      checkDone();
    });
  });
}

function getSubject(): Promise<string> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) {
      reject(new Error('No mail item available'));
      return;
    }

    item.subject.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || '');
      } else {
        reject(new Error('Failed to get subject'));
      }
    });
  });
}

function setSubject(subject: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) {
      reject(new Error('No mail item available'));
      return;
    }

    item.subject.setAsync(subject, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error('Failed to set subject'));
      }
    });
  });
}
