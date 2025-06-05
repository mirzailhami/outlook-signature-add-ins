/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, Office, Client, logger */

import { createNestablePublicClientApplication } from "@azure/msal-browser";
import "isomorphic-fetch";
import { Client } from "@microsoft/microsoft-graph-client";

let pca = undefined;
let isPCAInitialized = false;

const auth = {
  clientId: "44cb4054-0802-4e2f-8ccb-aba939633fbb",
  authority: "https://login.microsoftonline.com/common",
};

Office.onReady(() => {
  console.log("Office.js is ready");
});

/**
 * Initializes the Public Client Application (PCA) for SSO through NAA.
 * @throws {Error} If PCA initialization fails.
 */
async function initializePCA() {
  if (isPCAInitialized) return;

  try {
    pca = await createNestablePublicClientApplication({ auth });
    isPCAInitialized = true;
    logger.log("info", "initializePCA", { status: "PCA initialized successfully" });
  } catch (error) {
    logger.log("error", "initializePCA", { error: error.message, stack: error.stack });
    throw new Error(`Failed to initialize PCA: ${error.message}`);
  }
}

/**
 * Fetches an access token for Microsoft Graph API.
 * @returns {Promise<string>} The access token.
 * @throws {Error} If token acquisition fails.
 */
async function getGraphAccessToken() {
  await initializePCA();
  const tokenRequest = {
    scopes: ["User.Read", "Mail.ReadWrite", "Mail.Read", "openid", "profile"],
  };

  try {
    logger.log("info", "acquireTokenSilent", { status: "Attempting to acquire token silently" });
    const response = await pca.acquireTokenSilent(tokenRequest);
    logger.log("info", "acquireTokenSilent", { status: "Token acquired silently" });
    return response.accessToken;
  } catch (silentError) {
    logger.log("warn", "acquireTokenSilent", { error: silentError.message, stack: silentError.stack });
    try {
      logger.log("info", "acquireTokenPopup", { status: "Falling back to interactive token acquisition" });
      const response = await pca.acquireTokenPopup(tokenRequest);
      logger.log("info", "acquireTokenPopup", { status: "Token acquired interactively" });
      return response.accessToken;
    } catch (popupError) {
      logger.log("error", "acquireTokenPopup", { popupError: popupError.message, stack: popupError.stack });
      throw new Error(`Failed to acquire access token: ${popupError.message}`);
    }
  }
}

/**
 * Creates a Graph API client with the access token.
 * @returns {Promise<Client>} The initialized Graph API client.
 * @throws {Error} If token acquisition or client initialization fails.
 */
async function createGraphClient() {
  const accessToken = await getGraphAccessToken();
  return Client.init({
    authProvider: (done) => done(null, accessToken),
  });
}

/**
 * Fetches an email by its message ID.
 * @param {string} messageId - The ID of the email to fetch.
 * @returns {Promise<Object>} The email object with subject, body, sentDateTime, and toRecipients.
 * @throws {Error} If the email fetch fails.
 */
async function fetchEmailById(messageId) {
  if (!messageId) {
    throw new Error("Message ID is required to fetch email");
  }

  try {
    const client = await createGraphClient();
    logger.log("info", "fetchEmailById", { status: "Fetching email by ID", messageId });
    const email = await client
      .api(`/me/messages/${messageId}`)
      .select("id,subject,body,sentDateTime,toRecipients")
      .get();

    if (!email) {
      logger.log("warn", "fetchEmailById", { status: "Email not found", messageId });
      throw new Error("Email not found");
    }

    logger.log("info", "fetchEmailById", { status: "Email fetched successfully", emailId: email.id });
    return email;
  } catch (error) {
    logger.log("error", "fetchEmailById", { error: error.message, stack: error.stack, messageId });
    throw new Error(`Failed to fetch email by ID: ${error.message}`);
  }
}

// Export for Webpack bundling
export { fetchEmailById };

// Make functions available globally for Classic Outlook
window.fetchEmailById = fetchEmailById;
