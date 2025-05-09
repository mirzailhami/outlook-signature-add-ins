[![Deploy to GitHub Pages](https://github.com/mirzailhami/outlook-signature-add-ins/actions/workflows/deploy.yml/badge.svg)](https://github.com/mirzailhami/outlook-signature-add-ins/actions/workflows/deploy.yml)

# Outlook Signature Add-in

This Outlook add-in enables users to insert and validate M3 email signatures in Outlook. It supports multiple signature templates (Mona, Morgan, Morven, M2, M3) and enforces signature compliance for new emails, replies, and forwards using Smart Alerts.

## Features

- **Signature Insertion**: Select from predefined M3 signature templates via the ribbon menu ("M3 Signatures").
- **Signature Validation**: Prevents sending emails without an M3 signature or with modified signatures, with Smart Alerts for user correction.
- **Reply/Forward Support**: Automatically applies the last used signature for replies and forwards.
- **Cross-Platform**: Works on Outlook web, desktop (Windows, Mac), and mobile (iOS, Android).
- **Event-Based Activation**: Handles signature insertion and validation during compose and send events.

## Prerequisites

- Node.js (v18 or later)
- npm (v8 or later)
- Outlook with Microsoft 365 subscription (web, desktop, or mobile)
- Azure account for deployment
- GitHub account for CI/CD

## Setup

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/mirzailhami/outlook-signature-add-ins.git
   cd outlook-signature-add-ins
   ```

2. **Install Dependencies**:
   ```bash
   npm install
   ```

3. **Start Development Server**:
   ```bash
   npm run dev-server
   ```

4. **Sideload the Add-in**:
   - Open Outlook on the web (Chrome or Edge).
   - Go to **Settings > Manage Integrations > Add from File**.
   - Select `dist/manifest.xml`.
   - Alternatively, use:
     ```bash
     npm run start:web
     ```

## Usage

1. **Insert a Signature**:
   - Open a new email in Outlook.
   - Click the **M3 Signatures** ribbon button.
   - Select a signature (e.g., Mona, Morgan, M3) from the dropdown.
   - The signature is inserted into the email body.

2. **Reply/Forward**:
   - When replying or forwarding, the add-in automatically applies the last used signature from `localStorage`.

3. **Validation**:
   - If an email lacks an M3 signature, a Smart Alert prompts you to apply one ("Apply Signature" button).
   - If the signature is modified, a Smart Alert restores the original and allows sending ("Send Now" button).
   - Use the "Cancel" button to stop sending and edit the email.

## Development

### Project Structure
```
├── assets/               # Icon assets (icon-16.png, icon-32.png, etc.)
├── src/
│   ├── commands/         # Event handlers and ribbon commands
│   │   ├── commands.js   # Signature logic (validation, insertion)
│   │   ├── commands.html # Loads Office.js and commands.js
├── dist/                 # Build output (commands.js, commands.html, manifest.xml, assets)
├── manifest.xml          # Add-in manifest (ribbon commands, events)
├── webpack.config.js     # Webpack configuration
├── babel.config.json     # Babel configuration
├── package.json          # Dependencies and scripts
├── README.md             # Project documentation
```

### Build
```bash
npm run build
```

### Validate
```bash
npm run validate
```

### Test
- Test in Outlook web (Chrome/Edge), desktop (Windows/Mac), and mobile (iOS/Android).
- Verify:
  - Ribbon dropdown ("M3 Signatures") shows all signature options.
  - Signatures are inserted correctly.
  - Smart Alerts appear for missing or modified signatures with correct buttons ("Apply Signature", "Send Now", "Cancel").
  - Reply/forward emails include the last used signature.
- Check browser console logs for errors (e.g., `validateSignature`, `applyDefaultSignature`).

## Deployment

1. **Update manifest.xml**:
   - Replace `https://localhost:3000` with your Azure Static Web Apps URL (e.g., `https://white-grass-0b6dc6e03.6.azurestaticapps.net`).
   - Ensure `AppDomains` includes `https://m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net`.

2. **Deploy to Azure Static Web Apps**:
   - Create a Static Web App in Azure Portal.
   - Add the deployment token to GitHub Secrets (`AZURE_STATIC_WEB_APPS_API_TOKEN`).
   - Push changes to the `main` or `dev` branch to trigger the CI/CD pipeline.

3. **Deploy to Microsoft 365**:
   - Upload `dist/manifest.xml` to Microsoft 365 Admin Center:
     - **Settings > Integrated Apps > Upload Custom Apps**.
   - Assign to users or groups.

4. **Optional: AppSource**:
   - Submit to Microsoft Partner Center for public distribution.

## CI/CD

The repository uses GitHub Actions for CI/CD (`.github/workflows/deploy.yml`). It builds and deploys to Azure Static Web Apps on pushes to `main` or `dev`.

### Workflow Steps
- Checkout code.
- Set up Node.js.
- Install dependencies.
- Build the project.
- Deploy to Azure.
- Validate `manifest.xml`.

### Configuration
- Add `AZURE_STATIC_WEB_APPS_API_TOKEN` to GitHub Secrets.
- Ensure Azure Blob Storage CORS allows `https://white-grass-0b6dc6e03.6.azurestaticapps.net`.

## Troubleshooting

- **Smart Alert Not Showing**:
  - Verify OWA version supports Smart Alerts (`20250411007.09` or later).
  - Check logs for `displayError`:
    ```plaintext
    { event: "displayError", message: "...", restoreSignature: true }
    ```
  - Test manually:
    ```javascript
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "TestError",
      { type: "ErrorMessage", message: "Test error", icon: "Icon.16x16", persistent: true },
      (result) => console.log(result)
    );
    ```

- **Signature Not Applied**:
  - Check `localStorage` for `m3Signature`, `initialSignature`, `lastSentSignature`:
    ```javascript
    console.log(localStorage.getItem("initialSignature"));
    ```
  - Verify `commands.js` logs (`addSignature`, `applyDefaultSignature`).

- **Icons Not Loading**:
  - Ensure `dist/assets/icon-*.png` exists.
  - Check Outlook console for CSP errors.
  - Verify Azure Blob Storage CORS settings.

- **Contact**:
  - Open an issue: https://github.com/mirzailhami/outlook-signature-add-ins/issues
  - Email: support@m3wind.com

## License

MIT License. See [LICENSE](LICENSE) for details.
