# M3 Signatures Outlook Add-in

## Overview
The M3 Signatures Outlook Add-in enhances email composition in Outlook Web App by managing email signatures for new, reply, and forward emails. It ensures signatures are applied correctly, validated before sending, and restored if modified. The add-in uses the Office.js API to interact with Outlook, `localStorage` for persisting signature data, an external API to fetch signature templates, and the Microsoft Graph API to retrieve signatures from Sent Items for reply/forward scenarios.

### Features
- **Signature Selection**: Users can select from multiple signatures (Mona, Morgan, Morven, M2, M3) via the ribbon.
- **New Email Handling**: Prompts manual signature selection and stores a temporary signature for restoration if modified.
- **Reply/Forward Auto-Loading**: Automatically applies the signature used in the original email based on `conversationId`, recipients, or subject, using Microsoft Graph API to fetch from Sent Items.
- **Signature Validation**: Ensures the signature is valid and unmodified before sending; restores the original if modified.
- **Error Notifications**: Displays user-friendly notifications for missing or modified signatures.
- **Persistence**: Stores signature data in `localStorage` to track signatures across email threads.

### Architecture
The add-in consists of:
- **manifest.xml**: Defines the add-in's configuration, ribbon actions, and event handlers (version 1.0.0.50).
- **commands.js**: Core logic for signature handling, validation, storage, and Graph API integration.
- **taskpane.js/html**: UI for signature management (optional).
- **LocalStorage**: Stores signature templates (`signature_<key>`) and metadata (`signatureData_<timestamp>`).
- **External API**: Fetches signature templates from `https://m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net`.
- **Microsoft Graph API**: Retrieves email body content from Sent Items for reply/forward signature detection.

## Project Structure
```
.
├── README.md                  # Project documentation
├── assets                     # Source icons and signature images
│   ├── icon-128.png
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   └── signature-16.png
├── babel.config.json          # Babel configuration for transpiling JavaScript
├── certs                      # Development SSL certificates
│   ├── cert.pem
│   └── key.pem
├── dist                       # Build output directory
│   ├── assets                 # Copied icons and images
│   │   ├── icon-128.png
│   │   ├── icon-16.png
│   │   ├── icon-32.png
│   │   ├── icon-64.png
│   │   ├── icon-80.png
│   │   └── signature-16.png
│   ├── commands.html          # Command surface HTML
│   ├── commands.js            # Minified command logic
│   ├── commands.js.LICENSE.txt# License file for dependencies
│   ├── commands.js.map        # Source map for debugging
│   ├── index.html             # Entry point HTML
│   ├── manifest.xml           # Deployable manifest
│   ├── polyfill.js            # Polyfill for browser compatibility
│   ├── polyfill.js.map        # Source map for polyfill
│   ├── taskpane.html          # Taskpane UI HTML
│   └── taskpane.js            # Minified taskpane logic
├── manifest.xml               # Source manifest template
├── package-lock.json          # Dependency lock file
├── package.json               # Project metadata and scripts
├── src                        # Source code directory
│   ├── commands               # Command surface files
│   │   ├── commands.html
│   │   └── commands.js
│   ├── index.html             # Entry point template
│   ├── taskpane               # Taskpane UI files
│   │   └── taskpane.html
│   └── well-known             # Well-known configuration
│       └── microsoft-officeaddins-allowed.json
└── webpack.config.js          # Webpack configuration for building
```
- **`assets`**: Contains icon files for the add-in and a signature image.
- **`certs`**: SSL certificates for local development with HTTPS.
- **`dist`**: Output directory for the built add-in, including minified files and assets.
- **`src`**: Source files for development, including HTML templates and JavaScript logic.
- **`well-known`**: Configuration file for Office Add-in security policies.

## Setup
1. **Prerequisites**:
   - Node.js (v16+)
   - Outlook Web App (https://outlook.office365.com)
   - Azure account for signature API and Microsoft Graph API access (requires registering the add-in in Azure AD and obtaining an API client ID/secret)
   - Access to Microsoft Graph API with permissions: `User.Read`, `Mail.ReadWrite`, `Mail.Read`, `openid`, `profile`

2. **Installation**:
```
git clone https://github.com/mirzailhami/outlook-signature-add-ins
cd outlook-signature-add-ins
npm install
```

3. **Configuration**:
- Set up Azure AD application registration:
  - Register the add-in in Azure AD (https://portal.azure.com).
  - Configure API permissions: `User.Read`, `Mail.ReadWrite`, `Mail.Read`.
- Ensure the external API (`m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net`) is accessible.
- Create `.env` and `.env.production` files with environment variables (e.g., `ASSET_BASE_URL`).

4. **Development**:
- Start the dev server with environment variables:
```
npm run dev
```
- Output is in the `dist` folder, served at `https://localhost:3000`.
- Build for production:
```
npm run build
```
- Production output is in `dist`, deployed to `https://m3emailsignature.z33.web.core.windows.net/`.

5. **Sideloading**:
- Open Outlook Web App.
- Go to **Settings > Manage add-ins**.
- Remove existing add-in (if any).
- Upload `dist/manifest.xml`.
- Verify version `1.0.0.50` in **Manage add-ins**.

6. **Testing**:
- Clear browser cache or use Incognito mode:
- Chrome: DevTools > Application > Clear storage > Clear site data.
- Test new email, reply, and forward scenarios (see Flow below).
- Ensure Graph API authentication works by granting consent during the first run.

## Flow
The add-in handles email composition with the following flows, covering all cases including fixes for signature detection and modification.

### 1. New Email
- **No Signature Detected**:
- On compose, `onNewMessageComposeHandler` detects a new email (`isReplyOrForward: false`).
- Prompts: "Please select an M3 signature from the ribbon."
- Saves initial `signatureData_<timestamp>` with `signature: "none"`.
- **Signature Applied**:
- User selects a signature (e.g., `m2Signature`) from the ribbon.
- `addSignature` applies the signature and stores it in `tempSignature_new`.
- Saves `signatureData_<timestamp>` with `signature: "m2Signature"`.
- **Sending**:
- `validateSignature` checks the signature.
- If valid, clears `tempSignature_new` and saves `signatureData_<timestamp>`.
- **Modified Signature**:
- If the signature is modified, `validateSignatureChanges` detects the mismatch.
- Restores `tempSignature_new` and shows: "Selected M3 signature has been modified. Restoring the original signature."

### 2. Reply/Forward
- **Signature Auto-Loading**:
- `onNewMessageComposeHandler` detects reply/forward (`isReplyOrForward: true`).
- Uses Microsoft Graph API to fetch the latest email from Sent Items matching `conversationId` or `itemId`.
- Applies the matched signature (e.g., `m2Signature`) or prompts for manual selection if none is found.
- Saves updated `signatureData_<timestamp>`.
- **No Signature Detected**:
- If Graph API finds no match (`selectedSignatureKey: null`), prompts: "Please select an M3 signature from the ribbon."
- Saves `signatureData_<timestamp>` with `signature: "none"`.
- **Modified Signature**:
- On send, `validateSignatureChanges` detects modification.
- Restores the original signature using `signatureKey` from Graph API data.
- Shows: "Selected M3 signature has been modified. Restoring the original signature."

### 3. Fixed Cases
- **New Email Restoration**: Fixed signature restoration when modified, using `tempSignature_new`.
- **Reply/Forward Auto-Loading**: Enhanced with Microsoft Graph API to fetch signatures from Sent Items.
- **Signature Validation**: Prevents sending with modified or missing signatures, with proper restoration.
- **Async Reliability**: Made `saveSignatureData` and `restoreSignatureAsync` Promise-based for reliable execution.

## Development Notes
- **Clear Cache**: Always clear browser cache or use Incognito mode to avoid stale `localStorage` or scripts.
- **Logging**: Check console logs for debugging:
- `saveSignatureData`: Confirms storage.
- `detectSignatureKey`: Shows matches or mismatches.
- `validateSignatureChanges`: Tracks validation and restoration.
- `onNewMessageComposeHandler`: Monitors Graph API calls and outcomes.
- **Environment Variables**:
- Use `.env` for development (`npm run dev`).
- Use `.env.production` for production builds (`npm run build`).
- Example `.env`:
```
ASSET_BASE_URL=https://localhost:3000
```
- Example `.env.production`:
```
ASSET_BASE_URL=https://m3emailsignature.z33.web.core.windows.net
```
- **Versioning**: Current `manifest.xml` version is 1.0.0.50.

## Troubleshooting
- **Signature Not Auto-Loading**:
- Check `signatureDataEntries` in logs for `conversationId` or recipient mismatches.
- Verify Microsoft Graph API connectivity and permissions (`Mail.Read`).
- Ensure `localStorage` has `signatureData_<timestamp>` or Graph API returns valid Sent Items.
- **Restoration Fails**:
- Ensure `tempSignature_new` is set for new emails.
- Check `signature_<key>` in `localStorage` for replies or Graph API data integrity.
- **Graph API Errors**:
- Verify Azure AD authentication (client ID, tenant ID, scopes).
- Check network requests in browser DevTools for 401/403 errors; re-authenticate if needed.
- Look for `getGraphAccessToken` failures in logs.
- **Errors**:
- Look for `item.to.getAsync` or `item.subject.getAsync` failures in logs.
- Verify API connectivity to `m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net`.

## Contributing
- Submit pull requests with detailed descriptions and test cases.
- Test all flows (new, reply, forward) before merging, including edge cases (e.g., multiple signatures, no logo).
- Update this README for new features, fixes, or structure changes.
- Use conventional commit messages (e.g., `feat: add signature validation`, `fix: improve regex pattern`).

## License
MIT License. See [LICENSE](LICENSE) for details.