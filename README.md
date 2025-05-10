# M3 Signatures Outlook Add-in

## Overview
The M3 Signatures Outlook Add-in enhances email composition in Outlook Web App by managing email signatures for new, reply, and forward emails. It ensures signatures are applied correctly, validated before sending, and restored if modified. The add-in uses the Office.js API to interact with Outlook, `localStorage` for persisting signature data, and an external API to fetch signature templates.

### Features
- **Signature Selection**: Users can select from multiple signatures (Mona, Morgan, Morven, M2, M3) via the ribbon.
- **New Email Handling**: Prompts manual signature selection and stores a temporary signature for restoration if modified.
- **Reply/Forward Auto-Loading**: Automatically applies the signature used in the original email based on `conversationId`, recipients, or subject.
- **Signature Validation**: Ensures the signature is valid and unmodified before sending; restores the original if modified.
- **Error Notifications**: Displays user-friendly notifications for missing or modified signatures.
- **Persistence**: Stores signature data in `localStorage` to track signatures across email threads.

### Architecture
The add-in consists of:
- **manifest.xml**: Defines the add-in's configuration, ribbon actions, and event handlers (version 1.0.0.12).
- **commands.js**: Core logic for signature handling, validation, and storage.
- **taskpane.js/html**: UI for signature management (optional).
- **LocalStorage**: Stores signature templates (`signature_<key>`) and metadata (`signatureData_<timestamp>`).
- **External API**: Fetches signature templates from `https://m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net`.

See the [Architecture Diagram](architecture.mmd) for a visual representation.

## Setup
1. **Prerequisites**:
   - Node.js (v16+)
   - Outlook Web App (https://outlook.office365.com)
   - Azure account for signature API (optional for development)

2. **Installation**:
   ```bash
   git clone <repository-url>
   cd m3-signatures-outlook-addin
   npm install
   ```

3. **Development**:
   - Start the dev server:
     ```bash
     npm run dev-server
     ```
   - Output is in the `dist` folder:
     ```
     assets  commands.html  commands.[contenthash].js  manifest.xml  polyfill.[contenthash].js  taskpane.html  taskpane.[contenthash].js
     ```

4. **Sideloading**:
   - Open Outlook Web App.
   - Go to **Settings > Manage add-ins**.
   - Remove existing add-in (if any).
   - Upload `dist/manifest.xml`.
   - Verify version `1.0.0.12` in **Manage add-ins**.

5. **Testing**:
   - Clear browser cache or use Incognito mode:
     - Chrome: DevTools > Application > Clear storage > Clear site data.
   - Test new email, reply, and forward scenarios (see Flow below).

## Flow
The add-in handles email composition with the following flows, covering all cases including fixes for signature detection and modification.

### 1. New Email
- **No Signature Detected**:
  - On compose, `onNewMessageComposeHandler` detects a new email (`isReplyOrForward: false`).
  - Prompts: "Please select an M3 signature from the ribbon."
  - Saves initial `signatureData_<timestamp>` with `signature: "none"`.
  - Logs:
    ```javascript
    { event: "onNewMessageComposeHandler", status: "New email, requiring manual signature selection" }
    { event: "saveInitialSignatureData", status: "Stored initial signature data", recipients, conversationId, subject }
    ```
- **Signature Applied**:
  - User selects a signature (e.g., `m2Signature`) from the ribbon.
  - `addSignature` applies the signature and stores it in `tempSignature_new`.
  - Saves `signatureData_<timestamp>` with `signature: "m2Signature"`.
  - Logs:
    ```javascript
    { event: "addSignature", signatureKey: "m2Signature", isAutoApplied: false }
    { event: "addSignature", status: "Stored temporary signature for new email" }
    { event: "saveSignatureData", signatureKey: "m2Signature", recipients, conversationId, subject }
    ```
- **Sending**:
  - `validateSignature` checks the signature.
  - If valid, clears `tempSignature_new` and saves `signatureData_<timestamp>`.
  - Logs:
    ```javascript
    { event: "validateSignatureChanges", status: "Matched signature", matchedSignatureKey: "m2Signature" }
    { event: "validateSignatureChanges", status: "Cleared temporary signature for new email" }
    { event: "saveSignatureData", status: "Created new entry", key: "signatureData_..." }
    ```
- **Modified Signature**:
  - If the signature is modified, `validateSignatureChanges` detects the mismatch.
  - Restores `tempSignature_new` and shows: "Selected M3 signature has been modified. Restoring the original signature."
  - Logs:
    ```javascript
    { event: "validateSignatureChanges", status: "Restoring temporary signature for new email" }
    { event: "displayError", message: "Selected M3 signature has been modified. Restoring the original signature." }
    ```

### 2. Reply/Forward
- **Signature Auto-Loading**:
  - `onNewMessageComposeHandler` detects reply/forward (`isReplyOrForward: true`).
  - `getSignatureKeyForRecipients` matches `conversationId`, recipients, or subject in `signatureData_<timestamp>`.
  - Applies the matched signature (e.g., `m2Signature`).
  - Saves updated `signatureData_<timestamp>`.
  - Logs:
    ```javascript
    { event: "checkForReplyOrForward", status: "Reply/forward detected", conversationId: "AAQkAG..." }
    { event: "getSignatureKeyForRecipients", recipients: ["mr.ilhami@gmail.com"], conversationId, currentSubject }
    { event: "getSignatureKeyForRecipients", signatureDataEntries: [{ key, conversationId, signature: "m2Signature", ... }] }
    { event: "getSignatureKeyForRecipients", status: "Found matching signature by conversationId", signatureKey: "m2Signature" }
    { event: "addSignature", signatureKey: "m2Signature", isAutoApplied: true }
    { event: "saveSignatureData", status: "Updated existing entry", key: "signatureData_..." }
    ```
- **No Signature Detected**:
  - If no match is found (`selectedSignatureKey: null`), prompts: "Please select an M3 signature from the ribbon."
  - Saves `signatureData_<timestamp>` with `signature: "none"`.
  - Logs:
    ```javascript
    { event: "getSignatureKeyForRecipients", selectedSignatureKey: null }
    { event: "onNewMessageComposeHandler", status: "No signature found for reply/forward, requiring manual selection" }
    { event: "saveInitialSignatureData", status: "Stored initial signature data" }
    ```
- **Modified Signature** обсуж

  - On send, `validateSignatureChanges` detects modification.
  - Restores the original signature using `signatureKey` from `getSignatureKeyForRecipients`.
  - Shows: "Selected M3 signature has been modified. Restoring the original signature."
  - Logs:
    ```javascript
    { event: "validateSignatureChanges", status: "Signature modified" }
    { event: "displayError", message: "Selected M3 signature has been modified. Restoring the original signature." }
    ```

### 3. Fixed Cases
- **New Email Restoration**: Fixed signature restoration when modified, using `tempSignature_new`.
- **Reply/Forward Auto-Loading**: Fixed by ensuring `saveSignatureData` stores `conversationId` and normalizing recipient emails to lowercase.
- **Signature Validation**: Prevents sending with modified or missing signatures, with proper restoration.
- **Async Reliability**: Made `saveSignatureData` and `restoreSignatureAsync` Promise-based for reliable execution.
- **Debug Logging**: Added `signatureDataEntries` logging to debug mismatches in `getSignatureKeyForRecipients`.

## Development Notes
- **Clear Cache**: Always clear browser cache or use Incognito mode to avoid stale `localStorage` or scripts.
- **Logging**: Check console logs and Sentry for debugging:
  - `saveSignatureData`: Confirms storage.
  - `getSignatureKeyForRecipients`: Shows matches or mismatches.
  - `validateSignatureChanges`: Tracks validation and restoration.
- **Versioning**: Current `manifest.xml` version is 1.0.0.12.

## Troubleshooting
- **Signature Not Auto-Loading**:
  - Check `signatureDataEntries` in logs for `conversationId` or recipient mismatches.
  - Verify `localStorage` has `signatureData_<timestamp>`.
- **Restoration Fails**:
  - Ensure `tempSignature_new` is set for new emails.
  - Check `signature_<key>` in `localStorage` for replies.
- **Errors**:
  - Look for `item.to.getAsync` or `item.subject.getAsync` failures in logs.
  - Verify API connectivity to `m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net`.

## Contributing
- Submit pull requests with detailed descriptions.
- Test all flows (new, reply, forward) before merging.
- Update this README for new features or fixes.

## License
MIT License. See [LICENSE](LICENSE) for details.