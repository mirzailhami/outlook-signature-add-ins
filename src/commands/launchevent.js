/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, Office */

import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { auth } from "./authconfig";
import { logger } from "./helpers.js";

let pca = undefined;
let isPCAInitialized = false;

Office.onReady(() => {
  console.log("Office.js is ready");
});

/**
 * Initialize the public client application to work with SSO through NAA.
 */
async function initializePCA() {
  if (isPCAInitialized) return;

  try {
    pca = await createNestablePublicClientApplication({ auth });
    isPCAInitialized = true;
    logger.log("info", "initializePCA", { status: "PCA initialized successfully" });
  } catch (error) {
    logger.log("error", "initializePCA", { error: error.message });
    throw error;
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
    logger.log("error", "acquireTokenSilent", { silentError });
    try {
      logger.log("info", "acquireTokenPopup", { status: "Falling back to interactive token acquisition" });
      const response = await pca.acquireTokenPopup(tokenRequest);
      logger.log("info", "acquireTokenPopup", { status: "Token acquired interactively" });
      return response.accessToken;
    } catch (popupError) {
      logger.log("error", "acquireTokenPopup", { popupError });
      throw new Error(`Failed to acquire access token: ${popupError.message}`);
    }
  }
}

export { getGraphAccessToken };
