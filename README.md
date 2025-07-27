# M3 Signatures Outlook Add-in

## Overview
The M3 Signatures Outlook Add-in enhances email composition in Outlook by managing email signatures for new, reply, and forward emails. It ensures signatures are applied correctly, validated before sending, and restored if modified. The add-in uses the Office.js API to interact with Outlook, `localStorage` for persisting signature data, an external API to fetch signature templates, and the Microsoft Graph API to retrieve signatures from Sent Items for reply/forward scenarios.

### Features
- **Signature Selection**: Users can select from multiple signatures (Mona, Morgan, Morven, M2, M3) via the ribbon.
- **New Email Handling**: Prompts manual signature selection and stores a temporary signature key (`tempSignature`) for restoration if modified.
- **Reply/Forward Auto-Loading**: Automatically applies the signature used in the original email based on `conversationId` or `itemId`, using Microsoft Graph API to fetch from Sent Items.
- **Signature Validation**: Ensures the signature is valid and unmodified before sending; restores the original if modified.
- **Error Notifications**: Displays user-friendly notifications for missing or modified signatures.
- **Persistence**: Stores signature templates in `localStorage` using keys like `signature_<key>` (e.g., `signature_monaSignature`).

### Architecture
The add-in consists of:
- **manifest.xml**: Defines the add-in's configuration, ribbon actions, and event handlers (version 1.0.0.12).
- **commands.js**: Core logic for signature handling, validation, storage, and Graph API integration.
- **taskpane.js/html**: UI for signature management (optional).
- **LocalStorage**: Stores signature templates (`signature_<key>`) and temporary data (`tempSignature`).
- **External API**: Fetches signature templates from `https://m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net`.
- **Microsoft Graph API**: Retrieves email body content from Sent Items for reply/forward signature detection.


## Project Structure

```

.

├── README.md # Project documentation

├── assets # Source icons and signature images

│ ├── icon-128.png

│ ├── icon-16.png

│ ├── icon-32.png

│ ├── icon-64.png

│ ├── icon-80.png

│ └── signature-16.png

├── babel.config.json # Babel configuration for transpiling JavaScript

├── certs # Development SSL certificates

│ ├── cert.pem

│ └── key.pem

├── dist # Build output directory

│ ├── assets # Copied icons and images

│ │ ├── icon-128.png

│ │ ├── icon-16.png

│ │ ├── icon-32.png

│ │ ├── icon-64.png

│ │ ├── icon-80.png

│ │ └── signature-16.png

│ ├── commands.html # Command surface HTML

│ ├── commands.js # Minified command logic

│ ├── commands.js.LICENSE.txt# License file for dependencies

│ ├── commands.js.map # Source map for debugging

│ ├── index.html # Entry point HTML

│ ├── manifest.xml # Deployable manifest

│ ├── polyfill.js # Polyfill for browser compatibility

│ ├── polyfill.js.map # Source map for polyfill

│ ├── taskpane.html # Taskpane UI HTML

│ └── taskpane.js # Minified taskpane logic

├── manifest.xml # Source manifest template

├── package-lock.json # Dependency lock file

├── package.json # Project metadata and scripts

├── src # Source code directory

│ ├── commands # Command surface files

│ │ ├── commands.html

│ │ └── commands.js

│ ├── index.html # Entry point template

│ ├── taskpane # Taskpane UI files

│ │ └── taskpane.html

│ └── well-known # Well-known configuration

│ └── microsoft-officeaddins-allowed.json

└── webpack.config.js # Webpack configuration for building

```
- **`assets`**: Contains icon files for the add-in and a signature image.
- **`certs`**: SSL certificates for local development with HTTPS.
- **`dist`**: Output directory for the built add-in, including minified files and assets.
- **`src`**: Source files for development, including HTML templates and JavaScript logic.
- **`well-known`**: Configuration file for Office Add-in security policies.

## Setup
1. **Prerequisites**:
   - Node.js (v16+)
   - Outlook Web App[](https://outlook.office365.com)
   - Azure account for signature API and Microsoft Graph API access (requires registering the add-in in Azure AD and obtaining an API client ID/secret)
   - Access to Microsoft Graph API with permissions: `User.Read`, `Mail.ReadWrite`, `Mail.Read`, `openid`, `profile`
   - Azure Storage account access (`m3emailsignature` in resource group `rg-m3wind-vnet-win365-prod-uks01`)

2. **Installation**:
```
git clone https://github.com/mirzailhami/outlook-signature-add-ins
cd outlook-signature-add-ins
npm install
```

3. **Configuration**:
- Set up Azure AD application registration:
  - Register the add-in in Azure AD[](https://portal.azure.com).
  - Configure API permissions: `User.Read`, `Mail.ReadWrite`, `Mail.Read`.
- Ensure the external API (`m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net`) is accessible.
- Create `.env` and `.env.production` files with environment variables (e.g., `ASSET_BASE_URL`).

4. **Development**:
- Start the dev server with environment variables:

```
npm run dev
```
- Output is in `dist`, served at `https://localhost:3000`.
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

6. **Testing**:
- Clear browser cache or use Incognito mode:
  - Chrome: DevTools > Application > Clear storage > Clear site data.
- Test new email, reply, and forward scenarios (see Flow below).
- Ensure Graph API authentication works by granting consent during the first run.

## Flow
The add-in handles email composition with the following flows, reflecting the current implementation.

### 1. New Email
- **No Signature Detected**:
  - `onNewMessageComposeHandler` detects a new email (`isReplyOrForward: false`).
  - Prompts: "Please select an M3 signature from the ribbon."
- **Signature Applied**:
  - User selects a signature (e.g., `m2Signature`) via the ribbon, triggering `addSignature`.
  - `addSignature` fetches the signature template, applies it with a `<!-- signature -->` marker, and stores it in `localStorage` as `signature_m2Signature` and `tempSignature`.
- **Sending**:
  - `validateSignature` extracts the current signature using `extractSignatureForOutlookClassic` (for classic Outlook) or `extractSignature`.
  - If valid (matches the fetched template), allows sending and clears `tempSignature`.
  - If modified, restores the original signature and blocks sending with a notification.
- **Modified Signature**:
  - `validateSignatureChanges` detects mismatches in text or logo, restores the original via `addSignature`, and notifies the user.

### 2. Reply/Forward
- **Signature Auto-Loading**:
  - `onNewMessageComposeHandler` detects reply/forward (`isReplyOrForward: true`).
  - Uses Microsoft Graph API to fetch the original email by `conversationId` or `itemId` from Sent Items.
  - Extracts the signature using `extractSignature`, detects the key with `detectSignatureKey`, and applies it via `addSignature`.
  - Stores the key in `tempSignature` and the template in `signature_<key>`.
- **No Signature Detected**:
  - If no signature is found in the original email, prompts: "Please select an M3 signature from the ribbon."
- **Modified Signature**:
  - On send, `validateSignatureChanges` detects modification, restores the original using `tempSignature`, and notifies the user.

### 3. Fixed Cases
- **Signature Detection**: Improved with `extractSignatureForOutlookClassic` to correctly identify current signatures in reply/forward.
- **Validation**: Ensures no sending with modified signatures, with reliable restoration.
- **Async Reliability**: Uses callbacks for robust async operations.

## Adding a New Signature
To add a new signature (e.g., "M4"), follow these steps:

1. **Update `manifest.xml`**:
   - Add a new `Item` under the `MessageComposeCommandSurface` `ExtensionPoint` in the `DesktopFormFactor`:
     ```xml
     <Item id="msgReadMenuItem6">
       <Label resid="TaskpaneMenu.Label.M4"/>
       <Supertip>
         <Title resid="TaskpaneMenu.Label.M4"/>
         <Description resid="TaskpaneMenu.Tooltip.M4"/>
       </Supertip>
       <Icon xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0">
         <bt:Image size="16" resid="Icon.signature"/>
         <bt:Image size="32" resid="Icon.signature"/>
         <bt:Image size="80" resid="Icon.signature"/>
       </Icon>
       <Action xsi:type="ExecuteFunction">
         <FunctionName>addSignatureM4</FunctionName>
       </Action>
     </Item>
     ```
   - Add corresponding bt:ShortStrings and bt:LongStrings in the Resources section:
		```
		<bt:ShortStrings>
		<bt:String  id="TaskpaneMenu.Label.M4"  DefaultValue="M4"/>
		</bt:ShortStrings>
		<bt:LongStrings>
		<bt:String  id="TaskpaneMenu.Tooltip.M4"  DefaultValue="Insert M4 signature"/>
		</bt:LongStrings>
		``` 
2. **Update `commands.js`**:
	- Add a new function for the signature:
	```
	/**
	 * Adds the M4 signature.
	 * @param {Office.AddinCommands.Event} event - The Outlook event object.
	 */
	function addSignatureM4(event) {
	  addSignature("m4Signature", event, false, () => {});
	}
	Office.actions.associate("addSignatureM4", addSignatureM4);
	```
	- Ensure `detectSignatureKey` recognizes the new signature:
		- Update the `signatureKey` object to include `m4: "m4Signature".`
		- Add a new `companyPatterns` entry for "M4" (e.g., `/M4 Offshore Wind Limited/i`) and a `textChecks` entry if needed.
3. **Update the External API**:
	 -   Contact the API provider to add a new signature template at `https://m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net/api/Signatures/signatures`.
	 - Ensure the API returns the new signature in the `ribbons` endpoint and includes a unique logo URL (e.g., `m4_v1.png`) for detection.
4. **Update taskpane.html**:
  - Add a new `div` with a radio button for the "M4" signature inside the `#signatureOptions` container:
  ```
  <div class="choice-field">
    <label>
        <input type="radio" id="radioM4" name="signatureOption" value="m4Signature">
        M4
    </label>
  </div>
  ```
  - Ensure the new option integrates with the existing save logic, which already handles any `signatureOption` value via the `selectedRadio.value`.
5. **Test the New Signature**:
	- Build the add-in (`npm run build`) and sideload it in Outlook.
	- Select the "M4" option from the ribbon and verify it applies correctly.
  - Open the task pane, select "M4" as the default signature, and confirm it saves to `mobileDefaultSignature`.
	- Test validation and restoration in new, reply, and forward scenarios.

## How to Deploy Production
To deploy the add-in to production, follow these steps:
1. Build the Add-in:
- Run the production build command to generate the distributable files:
```
npm run build
```
- This creates the `dist` folder containing all necessary files (e.g., `manifest.xml`, `commands.js`, `taskpane.html`, etc.).
2. Access Azure Storage:
- Log in to the Azure Portal.
- Navigate to the storage account `m3emailsignature` in the resource group `rg-m3wind-vnet-win365-prod-uks01`.
- Select the Containers section and open the `$web` container, which is configured for static website hosting.
3. Upload Files:
- Use the Azure Portal’s upload feature or an Azure Storage Explorer tool (e.g., Azure Storage Explorer or AzCopy):
  - Drag and drop the entire `dist` folder contents into the `$web` container.
  - Ensure all files (e.g., `manifest.xml`, `commands.js`, `taskpane.html`, assets) are uploaded with the correct file structure.
4. Verify Deployment:
- The add-in will be accessible at `https://m3emailsignature.z33.web.core.windows.net/`.
- Test the deployment by sideloading the `manifest.xml` from this URL in Outlook Web App or updating the `ASSET_BASE_URL` in the `manifest.xml` to point to this location.
5. Update Environment:
- Update the `.env.production` file with the new `ASSET_BASE_URL` (e.g., `https://m3emailsignature.z33.web.core.windows.net`).
- Ensure the `manifest.xml` reflects the production URL for `IconUrl`, `HighResolutionIconUrl`, and other resource paths.
6. Security and Permissions:
- Ensure the `$web` container has public access set to "Container (anonymous read access for containers and blobs)" if needed, or use a Shared Access Signature (SAS) token for secure access.
- Verify Azure AD permissions and API connectivity post-deployment.


## Development Notes

-  **Clear Cache**: Always clear browser cache or use Incognito mode to avoid stale `localStorage` or scripts.

-  **Logging**: Check console logs for debugging:

-  `saveSignatureData`: Confirms storage.

-  `detectSignatureKey`: Shows matches or mismatches.

-  `validateSignatureChanges`: Tracks validation and restoration.

-  `onNewMessageComposeHandler`: Monitors Graph API calls and outcomes.

-  **Environment Variables**:

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

-  **Versioning**: Current `manifest.xml` version is 1.0.0.12.

## Troubleshooting
**Signature Not Auto-Loading**:
- Check `onNewMessageComposeHandler` logs for `fetchMessageById` errors.
- Verify Microsoft Graph API connectivity and permissions (`Mail.Read`).
- Ensure `localStorage` has `signature_<key>` or Graph API returns valid Sent Items.
- Ensure `tempSignature` is set for new emails or replies.
- Check `signature_<key>` in `localStorage` for integrity.
- Verify Azure AD authentication (client ID, tenant ID, scopes).
- Check network requests in browser DevTools for 401/403 errors; re-authenticate if needed.
- Look for `getGraphAccessToken` failures in logs.
- Look for `item.body.getAsync` or `fetchMessageById` failures in logs.
- Verify API connectivity to `m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net`.