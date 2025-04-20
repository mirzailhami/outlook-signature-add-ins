# Outlook Signature Add-in

This is an Outlook add-in that allows users to insert and validate M3 email signatures in Outlook. It supports multiple signature templates (Mona, Morgan, Morven, M2, M3) and ensures signatures are applied correctly in new emails, replies, and forwards.

## Features

- **Signature Insertion**: Select from predefined M3 signature templates via a taskpane or ribbon menu.
- **Signature Validation**: Prevents sending emails without an M3 signature or with modified signatures.
- **Reply/Forward Support**: Automatically applies the last sent signature in reply/forward scenarios.
- **Fluent UI**: Modern, accessible UI using Fluent UI components.
- **Cross-Platform**: Works on Outlook web, desktop (Windows, Mac), and mobile (iOS, Android).

## Prerequisites

- Node.js (v18 or later)
- npm (v8 or later)
- Outlook (web, desktop, or mobile) with Microsoft 365 subscription
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
   - Open Outlook on the web.
   - Go to **Manage Add-ins** > **Add from File**.
   - Upload `manifest.xml`.
   - Alternatively, use:
     ```bash
     npm run start:web
     ```

## Usage

1. **Insert a Signature**:
   - Open a new email in Outlook.
   - Click the **M3 Signatures** ribbon button.
   - Select a signature (e.g., Mona, Morgan) from the menu or taskpane.
   - The signature is inserted into the email body.

2. **Reply/Forward**:
   - When replying or forwarding, the add-in automatically applies the last sent signature.

3. **Validation**:
   - If an email lacks an M3 signature or has a modified signature, a dialog will prompt you to correct it.

## Development

- **Project Structure**:
  ```
  ├── assets/               # Image assets
  ├── src/
  │   ├── commands/         # Event handlers (commands.js)
  │   ├── error/            # Error dialog (error.html, error.js)
  │   ├── taskpane/         # Taskpane UI (App.jsx, HeroList.jsx)
  ├── dist/                 # Build output
  ├── manifest.xml          # Add-in manifest
  ├── webpack.config.js     # Webpack configuration
  ```

- **Build**:
  ```bash
  npm run build
  ```

- **Test**:
  - Test in Outlook web, desktop, and mobile.
  - Check console logs for errors.
  - Verify signature insertion, validation, and reply/forward behavior.

## Deployment

1. **Update manifest.xml**:
   - Replace `https://localhost:3000` with your production URL (e.g., `https://m3wind-addin.azurewebsites.net`).
   - Update image URLs and `AppDomains`.

2. **Deploy to Azure Static Web Apps**:
   - Create a Static Web App in Azure Portal.
   - Add the deployment token to GitHub Secrets (`AZURE_STATIC_WEB_APPS_API_TOKEN`).
   - Push changes to trigger the CI/CD pipeline.

3. **Deploy to Microsoft 365**:
   - Upload `manifest.xml` to Microsoft 365 Admin Center:
     - Settings > Integrated Apps > Upload Custom Apps.
   - Assign to users or groups.

4. **Optional: AppSource**:
   - Submit to Microsoft Partner Center for public distribution.

## CI/CD

The repository uses GitHub Actions for CI/CD. The workflow (`deploy.yml`) builds the add-in and deploys it to Azure Static Web Apps on pushes to the `main` branch.

- **Workflow Steps**:
  - Checkout code.
  - Set up Node.js.
  - Install dependencies.
  - Build the project.
  - Deploy to Azure.
  - Validate `manifest.xml`.

- **Configuration**:
  - Add `AZURE_STATIC_WEB_APPS_API_TOKEN` to GitHub Secrets.
  - Ensure Azure Blob Storage CORS is configured.

## Troubleshooting

- **Images Not Loading**:
  - Verify `dist/assets` contains all images.
  - Check Outlook console logs for CSP errors.
  - Ensure Azure Blob Storage CORS is set.

- **Reply/Forward Issues**:
  - Check `localStorage` for `lastSentSignature` and `initialSignature`.
  - Verify `commands.js` logs in the console.

- **Dialog Not Styled**:
  - Ensure `@fluentui/react-components` is installed (`npm list @fluentui/react-components`).
  - Check for `use-disposable` errors.

- **Contact**:
  - Open an issue on GitHub: https://github.com/mirzailhami/outlook-signature-add-ins/issues
  - Email: support@m3wind.com

## License

MIT License. See [LICENSE](LICENSE) for details.