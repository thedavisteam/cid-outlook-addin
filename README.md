# CID Subject Tagger Outlook Add-in

Automatically appends a CID tag to email subjects when recipients match the CID Contacts SharePoint list. Format: `User Subject - [CID-2026-0002]`. Works on Desktop/Web; mobile provides lookup-only task pane.

## Prerequisites
- Node.js 18+
- Office 365 with Outlook add-ins enabled
- SharePoint site and list:
  - **CID Register** (existing)
  - **CID Contacts** (Email, CID lookup, ContactType, DisplayName). Index Email column.
- Admin rights to deploy add-ins and create Power Automate flows

## Configure
1. Edit `src/config.ts` and set `siteUrl` to your SharePoint site URL.
2. Ensure the list name matches `contactsListName` (default "CID Contacts").
3. Manifest URLs now point to `https://julieandbryce.sharepoint.com/SiteAssets/CID%20Documents/…`; upload the build outputs (events.js, taskpane.html, manifest.xml, assets/) into that folder.

## Install dependencies
```
npm install
```

## Build
```
npm run build
```
Outputs to `dist/` and copies `manifest.xml` + static assets.

## Run locally
1. Serve `dist` at `https://localhost:3000` (for example with `npx http-server dist -S -C cert.pem -K key.pem`).
2. Sideload `manifest.xml` into Outlook (Desktop/Web) via Add-in commands.
3. Compose a message; the add-in will auto-append the CID tag when recipients match.

## Event behaviors
- **OnNewMessageCompose**: Checks recipients + subject on new/reply/forward.
- **OnMessageRecipientsChanged**: Re-checks when To/CC/BCC change.
- **OnMessageSend**: Final validation; soft-blocks if multiple CIDs match.
- **No match**: Shows an info bar only (no tag added).

## Mobile
- Outlook mobile cannot set subject via API. Use the Task Pane to look up CID and copy the tag manually.

## Power Automate (recommended)
- **Dual-write**: Flow that mirrors CID Register entries into CID Contacts on create/update. Map: CID Register `Title` (CID) → CID Contacts `CID` lookup; `Client Email 1` (`Client_x0020_Email`) and `Client Email 2` (`Client_x0020_Email_x0020_2`) each create/update a CID Contacts row (`Email`); `Client Name` (`Client_x0020_Name`) → CID Contacts `DisplayName`; `KW Opportunity Name` stays on the CID item.
- **Incoming mail tagging**: Flow using Graph `PATCH /messages/{id}` to add CID tag for incoming mail when sender matches CID Contacts.
- Existing BCC flow for `CID-20` patterns remains compatible.

## Deployment
- Update `manifest.xml` URLs to your hosted bundle location.
- Deploy via Microsoft 365 Admin Center → Integrated apps for org-wide auto-launch.

## Testing checklist
- Single recipient with CID → subject gains ` - [CID-XXXX-XXXX]`.
- No CID match → info bar only.
- Multiple CIDs → send is blocked with guidance.
- Reply/Forward preserves `RE:/FW:` prefixes while appending CID.
