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

/**
 * Initialize the public client application to work with SSO through NAA.
 */
async function initializePCA() {
  if (isPCAInitialized) return;

  try {
    pca = await createNestablePublicClientApplication({ auth });
    isPCAInitialized = true;
  } catch (error) {
    logger.log(`error`, `initializePCA`, `Error creating pca: ${error}`);
  }
}

/**
 * Fetches an access token for Microsoft Graph API.
 * @returns {Promise<string>} The access token.
 */
async function getGraphAccessToken() {
  await initializePCA();
  const tokenRequest = {
    scopes: ["Mail.Read", "openid", "profile"],
  };

  try {
    const response = await pca.acquireTokenSilent(tokenRequest);
    logger.log(`info`, `getGraphAccessToken`, { response });
    return response.accessToken;
  } catch (silentError) {
    logger.log(`error`, `acquireTokenSilent`, { silentError });
    try {
      const response = await pca.acquireTokenPopup(tokenRequest);
      logger.log(`info`, `acquireTokenPopup`, { response });
      return response.accessToken;
    } catch (popupError) {
      logger.log(`error`, `acquireTokenPopup`, { popupError });
      throw new Error("Failed to acquire access token.");
    }
  }
}

export { getGraphAccessToken };
