/* global Office, console */

/* eslint-disable no-undef */

import { createNestablePublicClientApplication } from "@azure/msal-browser";
import "isomorphic-fetch";
import { Client } from "@microsoft/microsoft-graph-client";

// Robust storage fallback with logging
let storage = typeof localStorage !== "undefined" ? localStorage : {};

function storageSetItem(key, value) {
  const hostName = Office.context.mailbox.diagnostics.hostName;
  displayNotification(
    "Info",
    `storageSetItem: Setting ${key} = ${value}, host: ${hostName}, using ${
      typeof localStorage !== "undefined" ? "localStorage" : "in-memory storage"
    }`
  );
  if (typeof localStorage !== "undefined") {
    localStorage.setItem(key, value);
    displayNotification("Info", `storageSetItem: Using localStorage for ${key} = ${value}`);
  } else {
    storage[key] = value;
    displayNotification("Info", `storageSetItem: Using in-memory for ${key} = ${value}`);
  }
}

function storageGetItem(key) {
  const hostName = Office.context.mailbox.diagnostics.hostName;
  displayNotification(
    "Info",
    `storageGetItem: Getting ${key}, host: ${hostName}, using ${
      typeof localStorage !== "undefined" ? "localStorage" : "in-memory storage"
    }`
  );
  if (typeof localStorage !== "undefined") {
    const value = localStorage.getItem(key);
    displayNotification("Info", `storageGetItem: Using localStorage for ${key} = ${value || "null"}`);
    return value;
  } else {
    const value = storage[key] || null;
    displayNotification("Info", `storageGetItem: Using in-memory for ${key} = ${value || "null"}`);
    return value;
  }
}

function storageRemoveItem(key) {
  const hostName = Office.context.mailbox.diagnostics.hostName;
  displayNotification(
    "Info",
    `storageRemoveItem: Removing ${key}, host: ${hostName}, using ${
      typeof localStorage !== "undefined" ? "localStorage" : "in-memory storage"
    }`
  );
  if (typeof localStorage !== "undefined") {
    localStorage.removeItem(key);
    displayNotification("Info", `storageRemoveItem: Using localStorage for ${key}`);
  } else {
    delete storage[key];
    displayNotification("Info", `storageRemoveItem: Using in-memory for ${key}`);
  }
}

/**
 * Centralized logger for structured logging with timestamps.
 * @type {Object}
 */
const logger = {
  /**
   * Logs a message with a timestamp and structured details.
   * @param {string} level - Log level ("info" or "error").
   * @param {string} event - Event name or identifier.
   * @param {Object} [details={}] - Additional details to log.
   */
  log(level, event, details = {}) {
    const timestamp = new Date().toISOString();
    const logEntry = {
      timestamp,
      level,
      event,
      ...details,
    };
    if (level === "error") {
      console.error(JSON.stringify(logEntry, null, 2));
    } else {
      console.log(JSON.stringify(logEntry, null, 2));
    }
  },
};

/**
 * Manages email signature extraction, normalization, and restoration.
 * @type {Object}
 */
const SignatureManager = {
  /**
   * Extracts a signature from an email body using markers or regex patterns.
   * @param {string|null} body - The email body HTML content.
   * @returns {string|null} The extracted signature, or null if not found.
   */
  extractSignature(body) {
    if (!body) return null;

    const marker = "<!-- signature -->";
    const startIndex = body.indexOf(marker);
    if (startIndex !== -1) {
      const endIndex = body.indexOf("</body>", startIndex);
      const signature = body.slice(startIndex + marker.length, endIndex !== -1 ? endIndex : undefined).trim();
      return signature;
    }

    const regexes = [
      /<div\s+class="Signature"[^>]*>([\s\S]*?)$/is,
      /<div\s+id="Signature"[^>]*>([\s\S]*?)$/is,
      /<table[^>]*>([\s\S]*?)$/is,
    ];
    for (const regex of regexes) {
      const match = body.match(regex);
      if (match) {
        const signature = match[0].trim();
        return signature;
      }
    }

    return null;
  },

  /**
   * Extracts a signature from an email body in Outlook Classic using specific markers or regex.
   * @param {string|null} body - The email body HTML content.
   * @returns {string|null} The extracted signature, or null if not found.
   */
  extractSignatureForOutlookClassic(body) {
    if (!body) return null;

    const marker = "<!-- signature -->";
    const startIndex = body.lastIndexOf(marker);
    if (startIndex !== -1) {
      const endIndex = body.indexOf("</body>", startIndex);
      const signature = body.slice(startIndex + marker.length, endIndex !== -1 ? endIndex : undefined).trim();
      logger.log("info", "extractSignatureForOutlookClassic", { method: "marker", signatureLength: signature.length });
      return signature;
    }

    const regex =
      /<table\s+class=MsoNormalTable[^>]*>([\s\S]*?)(?=(?:<div\s+id="[^"]*appendonsend"|>?\s*<(?:table|hr)\b)|$)/is;
    const match = body.match(regex);
    if (match) {
      const signature = match[1].trim();
      logger.log("info", "extractSignatureForOutlookClassic", { method: "table", signatureLength: signature.length });
      return signature;
    }

    logger.log("info", "extractSignatureForOutlookClassic", { status: "No signature found" });
    return null;
  },

  /**
   * Normalizes a signature by removing HTML tags and standardizing text.
   * @param {string|null} sig - The signature to normalize.
   * @returns {string} The normalized signature, or an empty string if null.
   */
  normalizeSignature(sig) {
    if (!sig) return "";

    let normalized = sig;

    // Manual HTML entity decoding
    const htmlEntities = {
      "&amp;": "&",
      "&lt;": "<",
      "&gt;": ">",
      "&quot;": '"',
      "&nbsp;": " ",
      "&#160;": " ",
      "&#39;": "'",
      "&apos;": "'",
    };
    for (const [entity, char] of Object.entries(htmlEntities)) {
      normalized = normalized.replace(new RegExp(entity, "gi"), char);
    }

    // Remove HTML tags
    normalized = normalized.replace(/<[^>]+>/g, " ");

    // Clean up text
    normalized = normalized
      .replace(/[\r\n]+/g, " ") // Replace newlines with a single space
      .replace(/\s*([.,:;])\s*/g, "$1") // Remove spaces around punctuation
      .replace(/\s+/g, " ") // Collapse multiple spaces into one
      .replace(/\s*:\s*/g, ":") // Remove spaces around colons
      .replace(/\s+(email:)/gi, "$1") // Remove spaces before "email:"
      .trim() // Remove leading/trailing spaces
      .toLowerCase();

    return normalized;
  },

  /**
   * Normalizes an email subject by removing reply/forward prefixes.
   * @param {string|null} subject - The email subject.
   * @returns {string} The normalized subject, or an empty string if null.
   */
  normalizeSubject(subject) {
    if (!subject) return "";
    const normalized = subject
      .replace(/^(re:|fw:|fwd:)\s*/i, "")
      .trim()
      .toLowerCase();
    logger.log("info", "normalizeSubject", { rawLength: subject.length, normalizedLength: normalized.length });
    return normalized;
  },

  /**
   * Checks if an email is a reply or forward using callbacks.
   * @param {Office.MessageCompose} item - The email item.
   * @param {function(boolean, Error|null)} callback - Callback with result and error.
   */
  isReplyOrForward(item, callback) {
    item.getComposeTypeAsync(function (asyncResult) {
      if (asyncResult.status === "succeeded") {
        callback(asyncResult.value.composeType === "newMail" ? false : true, null);
        return;
      }
      callback(false, null);
    });
  },

  /**
   * Restores a signature to the email body using callbacks.
   * @param {Office.MessageCompose} item - The email item.
   * @param {string} signature - The signature to restore.
   * @param {string} signatureKey - The signature key for fallback.
   * @param {function(boolean, Error|null)} callback - Callback with success and error.
   */
  restoreSignature(item, signature, signatureKey, callback) {
    logger.log("info", "restoreSignatureAsync", { signatureKey, cachedSignatureLength: signature?.length });
    if (!signature) {
      signature = storageGetItem(`signature_${signatureKey}`);
      logger.log("info", "restoreSignatureAsync", {
        status: "Falling back to signatureKey",
        fallbackLength: signature?.length,
      });
      if (!signature) {
        logger.log("error", "restoreSignatureAsync", { error: "No signature available" });
        callback(false, new Error("No signature available"));
        return;
      }
    }

    const signatureWithMarker = "<!-- signature -->" + signature.trim();
    item.body.getAsync("html", { asyncContext: { signatureWithMarker, callback } }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        logger.log("error", "restoreSignatureAsync", { error: "Failed to get current body" });
        callback(false, new Error("Failed to get current body"));
        return;
      }

      const currentBody = result.value || "";
      const startIndex = currentBody.indexOf("<!-- signature -->");
      if (startIndex === -1) {
        logger.log("warn", "restoreSignatureAsync", { error: "Signature marker not found, appending instead" });
        item.body.setSignatureAsync(signatureWithMarker, { coercionType: Office.CoercionType.Html }, (asyncResult) => {
          callback(asyncResult.status !== Office.AsyncResultStatus.Failed, asyncResult.error || null);
        });
      } else {
        const endIndex =
          currentBody.indexOf("</body>", startIndex) !== -1
            ? currentBody.indexOf("</body>", startIndex)
            : currentBody.length;
        const newBody = currentBody.substring(0, startIndex) + signatureWithMarker + currentBody.substring(endIndex);
        item.body.setAsync(newBody, { coercionType: Office.CoercionType.Html }, (asyncResult) => {
          callback(asyncResult.status !== Office.AsyncResultStatus.Failed, asyncResult.error || null);
        });
      }
    });
  },
};

/**
 * Detects the signature key based on content keywords and logo URL.
 * @param {string} signatureText - The signature text to analyze.
 * @returns {string|null} The matched signature key, or null if no match.
 */
function detectSignatureKey(signatureText) {
  // Step 1: Logo-based detection
  const logoRegex = /<img[^>]+src=["'](.*?(?:m3signatures\/logo\/([^?"']+))(?:\?[^"']*)?)["'][^>]*>/i;
  const logoMatch = signatureText.match(logoRegex);
  let logoFile = logoMatch ? logoMatch[2] : null;

  if (logoFile) {
    // Extract the prefix (e.g., "m3" from "m3_v1.png")
    const logoPrefixMatch = logoFile.match(/^([a-z0-9]+)(?:_v\d+)?\.png$/i);
    const logoPrefix = logoPrefixMatch ? logoPrefixMatch[1] : null;

    if (logoPrefix) {
      const logoPrefixToKey = {
        morven: "morvenSignature",
        morgan: "morganSignature",
        mona: "monaSignature",
        m2: "m2Signature",
        m3: "m3Signature",
      };
      const keyFromLogo = logoPrefixToKey[logoPrefix.toLowerCase()];
      if (keyFromLogo) {
        logger.log("info", "logoDetection", {
          logoFile,
          logoPrefix,
          keyFromLogo,
        });
        return keyFromLogo;
      }
    } else {
      logger.log("error", "logoDetection", {
        status: "Invalid logo file name format",
        logoFile,
      });
    }
  }

  return null;
}

/**
 * Fetches a signature template from the API and applies user-specific replacements.
 * @param {string} signatureKey - The signature key (e.g., "m3Signature").
 * @param {function(string|null, Error|null)} callback - Callback with the template or error.
 */
function fetchSignature(signatureKey, callback) {
  const signatureIndex = ["monaSignature", "morganSignature", "morvenSignature", "m2Signature", "m3Signature"].indexOf(
    signatureKey
  );
  const initialUrl = "https://m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net/api/Ribbons/ribbons";
  let signatureUrl =
    "https://m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net/api/Signatures/signatures?signatureURL=";

  fetch(initialUrl)
    .then((response) => response.json())
    .then((data) => {
      signatureUrl += data.result[signatureIndex].url;
      fetch(signatureUrl)
        .then((response) => response.json())
        .then((data) => {
          let template = data.result
            .replace("{First name} ", Office.context.mailbox.userProfile.displayName || "")
            .replace("{Last name}", "")
            .replaceAll("{E-mail}", Office.context.mailbox.userProfile.emailAddress || "")
            .replace("{Title}", "")
            .trim();
          callback(template, null);
        })
        .catch((err) => callback(null, err));
    })
    .catch((err) => callback(null, err));
}

/**
 * Appends debug logs to the email body using setSignatureAsync for mobile debugging.
 * @param {Office.MessageCompose} item - The Outlook message item.
 * @param {...(string|any)} args - Variable arguments: label-value pairs or messages.
 * @param {function()} callback - Callback when done.
 */
function appendDebugLogToBody(item, ...args) {
  const timestamp = new Date().toISOString();
  let logContent = `<div style="font-family: Arial, sans-serif; font-size: 12px; color: #333; margin: 10px 0; border-bottom: 1px solid #ccc; padding-bottom: 5px;">`;
  logContent += `<strong>[${timestamp}]</strong><br>`;

  // Process arguments into key-value pairs or messages
  for (let i = 0; i < args.length; i += 2) {
    const label = args[i];
    const value = i + 1 < args.length ? args[i + 1] : "";
    if (typeof label === "string") {
      logContent += `<span style="color: #0055aa;"><strong>${label}:</strong></span> ${
        value !== undefined && value !== null ? JSON.stringify(value) : "undefined"
      }<br>`;
    } else {
      logContent += `<span>${label}</span><br>`;
    }
  }
  logContent += `</div>`;

  // Append the log to the existing body
  item.body.getAsync("html", { asyncContext: logContent }, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const currentBody = result.value || "";
      item.body.setSignatureAsync(
        currentBody + logContent,
        { coercionType: Office.CoercionType.Html },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            // Fallback: Try setting the log alone
            item.body.setSignatureAsync(logContent, { coercionType: Office.CoercionType.Html }, () => {});
          }
        }
      );
    } else {
      // Fallback if getAsync fails
      item.body.setSignatureAsync(logContent, { coercionType: Office.CoercionType.Html }, () => {});
    }
  });
}

/**
 * Completes the event with a signature state and optional notification.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 * @param {string} [notificationType] - Notification type ("Info" or "Error").
 * @param {string} [notificationMessage] - Notification message.
 * @param {boolean} [persistent] - Whether the notification persists.
 */
function completeWithState(event, notificationType, notificationMessage, persistent = false) {
  if (notificationMessage) {
    displayNotification(notificationType, notificationMessage, persistent);
  }
  event.completed();
}

/**
 * Displays a notification in the Outlook UI.
 * @param {string} type - Notification type ("Error" or "Info").
 * @param {string} message - Notification message.
 * @param {boolean} persistent - Whether the notification persists.
 */
function displayNotification(type, message, persistent = false) {
  try {
    const item = Office.context.mailbox.item;
    if (!item) {
      logger.log("error", "displayNotification", { error: "No mailbox item" });
      return;
    }

    const notificationType =
      type === "Error"
        ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
        : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;

    const notification = { type: notificationType, message };
    if (type === "Info") {
      notification.icon = "none";
      notification.persistent = persistent;
    }

    item.notificationMessages.addAsync(`notif_${new Date().getTime()}`, notification, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        logger.log("error", "displayNotification", { error: result.error.message });
      }
    });
  } catch (error) {
    logger.log("error", "displayNotification", { error: error.message });
  }
}

/**
 * Displays an error with a Smart Alert and notification.
 * @param {string} message - Error message.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function displayError(message, event) {
  const markdownMessage = message.includes("modified")
    ? `${message}\n\n**Tip**: Ensure the M3 signature is not edited before sending.`
    : `${message}\n\n**Tip**: Select an M3 signature from the ribbon under "M3 Signatures".`;

  event.completed({
    allowEvent: false,
    errorMessage: message,
    errorMessageMarkdown: markdownMessage,
    cancelLabel: "OK",
  });
}

// Graph API functions from graph.js
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
 * @param {function(Error|null)} callback - Callback with error if initialization fails.
 */
function initializePCA(callback) {
  if (isPCAInitialized) return;

  createNestablePublicClientApplication({ auth }).then(
    (pcaInstance) => {
      pca = pcaInstance;
      isPCAInitialized = true;
      logger.log("info", "initializePCA", { status: "PCA initialized successfully" });
      callback(null);
    },
    (error) => {
      logger.log("error", "initializePCA", { error: error.message, stack: error.stack });
      callback(new Error(`Failed to initialize PCA: ${error.message}`));
    }
  );
}

/**
 * Fetches an access token for Microsoft Graph API.
 * @param {function(string|null, Error|null)} callback - Callback with token or error.
 */
function getGraphAccessToken(callback) {
  initializePCA((initError) => {
    if (initError) {
      callback(null, initError);
      return;
    }

    const tokenRequest = {
      scopes: ["User.Read", "Mail.ReadWrite", "Mail.Read", "openid", "profile"],
    };

    logger.log("info", "acquireTokenSilent", { status: "Attempting to acquire token silently" });
    pca.acquireTokenSilent(tokenRequest).then(
      (silentResponse) => {
        logger.log("info", "acquireTokenSilent", { status: "Token acquired silently" });
        callback(silentResponse.accessToken, null);
      },
      (silentError) => {
        logger.log("warn", "acquireTokenSilent", { error: silentError.message });
        logger.log("info", "acquireTokenPopup", { status: "Falling back to interactive token acquisition" });
        pca.acquireTokenPopup(tokenRequest).then(
          (popupResponse) => {
            logger.log("info", "acquireTokenPopup", { status: "Token acquired interactively" });
            callback(popupResponse.accessToken, null);
          },
          (popupError) => {
            logger.log("error", "acquireTokenPopup", { popupError: popupError.message });
            callback(null, new Error(`Failed to acquire access token: ${popupError.message}`));
          }
        );
      }
    );
  });
}

/**
 * Creates a Graph API client with the access token.
 * @param {function(Client|null, Error|null)} callback - Callback with client or error.
 */
function createGraphClient(callback) {
  getGraphAccessToken((accessToken, error) => {
    if (error || !accessToken) {
      callback(null, error || new Error("No access token available"));
      return;
    }
    try {
      const client = Client.init({
        authProvider: (done) => done(null, accessToken),
      });
      callback(client, null);
    } catch (clientError) {
      callback(null, new Error(`Failed to initialize Graph client: ${clientError.message}`));
    }
  });
}

/**
 * Fetches message by its message ID.
 * @param {string} messageId - The ID of the message to fetch.
 * @param {function(Object|null, Error|null)} callback - Callback with message or error.
 */
function fetchMessageById(messageId, callback) {
  if (!messageId) {
    callback(null, new Error("Message ID is required to fetch message"));
    return;
  }

  createGraphClient((client, error) => {
    if (error || !client) {
      logger.log("error", "fetchMessageById", { error: error?.message, messageId });
      callback(null, error || new Error("Graph client not initialized"));
      return;
    }

    client
      .api(`/me/messages/${messageId}`)
      .select("id,subject,body,sentDateTime,toRecipients")
      .get()
      .then((message) => {
        if (!message) {
          logger.log("warn", "fetchMessageById", { status: "Message not found", messageId });
          callback(null, new Error("Email not found"));
        } else {
          callback(message, null);
        }
      })
      .catch((graphError) => {
        logger.log("error", "fetchMessageById", { error: graphError.message, messageId });
        callback(null, new Error(`Failed to fetch email by ID: ${graphError.message}`));
      });
  });
}

/**
 * Adds a signature to the email and saves it to storage.
 * @param {string} signatureKey - The signature key (e.g., "m3Signature").
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 * @param {boolean} isAutoApplied - Whether the signature is auto-applied.
 * @param {function()} callback - Callback when done.
 */
function addSignature(signatureKey, event, isAutoApplied, callback) {
  try {
    const item = Office.context.mailbox.item;

    storageRemoveItem("tempSignature");
    storageSetItem("tempSignature", signatureKey);
    const cachedSignature = storageGetItem(`signature_${signatureKey}`);

    if (cachedSignature && !isAutoApplied) {
      const signatureWithMarker = "<!-- signature -->" + cachedSignature.trim();
      item.body.setSignatureAsync(signatureWithMarker, { coercionType: Office.CoercionType.Html }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          if (isMobile) {
            appendDebugLogToBody(item, "addSignature Error (Cached)", "Message", asyncResult.error.message);
          }
          if (!isAutoApplied) {
            event.completed();
            callback();
          } else {
            displayError("Failed to set cached signature.", event);
            callback();
          }
          return;
        }
        item.body.getAsync("html", (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            logger.log("debug", "addSignature", {
              bodyContainsMarker: result.value.includes("<!-- signature -->"),
              bodyLength: result.value.length,
            });
          }
          event.completed();
          callback();
        });
      });
    } else {
      fetchSignature(signatureKey, (template, error) => {
        if (error) {
          if (isMobile) {
            appendDebugLogToBody(item, "addSignature Error (Fetch)", "Message", error.message);
          }
          logger.log("error", "addSignature", { error: error.message });
          displayNotification("Error", `Failed to fetch ${signatureKey}.`, true);
          if (!isAutoApplied) {
            event.completed();
            callback();
          } else {
            displayError(`Failed to fetch ${signatureKey}.`, event);
            callback();
          }
          return;
        }

        const signatureWithMarker = "<!-- signature -->" + template.trim();
        item.body.setSignatureAsync(signatureWithMarker, { coercionType: Office.CoercionType.Html }, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            if (isMobile) {
              appendDebugLogToBody(item, "addSignature Error (Set)", "Message", asyncResult.error.message);
            }
            logger.log("error", "addSignature", { error: asyncResult.error.message });
            displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
            if (!isAutoApplied) {
              event.completed();
              callback();
            } else {
              displayError(`Failed to apply ${signatureKey}.`, event);
              callback();
            }
            return;
          }
          item.body.getAsync("html", (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              logger.log("debug", "addSignature", {
                bodyContainsMarker: result.value.includes("<!-- signature -->"),
                bodyLength: result.value.length,
              });
            }
            storageSetItem(`signature_${signatureKey}`, template);
            event.completed();
            callback();
          });
        });
      });
    }
  } catch (error) {
    displayNotification("Error", `addSignature: Exception - ${error.message}`);
    displayError(`Unexpected error occurred during addSignature: ${error.message}`, event);
  }
}

/**
 * Validates the email signature on send.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function validateSignature(event) {
  const hostName = Office.context.mailbox.diagnostics.hostName;
  const isClassicOutlook = hostName === "Outlook";

  const item = Office.context.mailbox.item;
  if (!item) {
    logger.log("error", "validateSignature", { error: "No mailbox item" });
    displayError("No mailbox item available.", event);
    return;
  }

  item.body.getAsync("html", (bodyResult) => {
    if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
      logger.log("error", "validateSignature", { error: "Failed to get body" });
      displayError("Failed to get email body.", event);
      return;
    }

    const body = bodyResult.value;
    const currentSignature = isClassicOutlook
      ? SignatureManager.extractSignatureForOutlookClassic(body)
      : SignatureManager.extractSignature(body);

    displayNotification(
      "Info",
      `validateSignature: currentSignature length: ${currentSignature?.length || "null"}, isClassicOutlook: ${isClassicOutlook}`
    );

    if (!currentSignature) {
      displayError("Email is missing the M3 required signature. Please select an appropriate email signature.", event);
    } else {
      validateSignatureChanges(item, currentSignature, event, isClassicOutlook);
    }
  });
}

/**
 * Validates if the signature has been modified or changed.
 * @param {Office.MessageCompose} item - The email item.
 * @param {string} currentSignature - The current signature in the email body.
 * @param {boolean} isClassicOutlook - Whether the Outlook version is classic.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function validateSignatureChanges(item, currentSignature, event, isClassicOutlook) {
  try {
    if (isClassicOutlook) {
      // Step 1: Detect signature key from current signature
      const originalSignatureKey = detectSignatureKey(currentSignature);
      displayNotification(
        "Info",
        `validateSignatureChanges: Detected originalSignatureKey from current signature: ${originalSignatureKey || "null"}, currentSignature: ${currentSignature.length}`
      );

      if (!originalSignatureKey) {
        displayError("Could not detect M3 signature. Please select a signature from the ribbon.", event);
        event.completed({ allowEvent: false });
        return;
      }

      // Step 2: Fetch the signature and compare
      fetchSignature(originalSignatureKey, (fetchedSignature, error) => {
        if (error || !fetchedSignature) {
          displayNotification(
            "Error",
            `validateSignatureChanges: Failed to fetch ${originalSignatureKey}, error: ${error?.message || "null"}, fetchedSignature: ${fetchedSignature || "null"}`
          );
          displayError("Failed to validate signature. Please reselect.", event);
          event.completed({ allowEvent: false });
          return;
        }

        // Step 2.5: Extract and validate the fetched signature
        const rawMatchedSignature = fetchedSignature;

        // Step 3 & 4: Compare and decide
        const cleanCurrentSignature = SignatureManager.normalizeSignature(currentSignature);
        const cleanFetchedSignature = SignatureManager.normalizeSignature(rawMatchedSignature);

        const logoRegex = /<img[^>]+src=["'](.*?(?:m3signatures\/logo\/[^"']+))["'][^>]*>/i;
        const currentLogoMatch = currentSignature.match(logoRegex);
        let currentLogoUrl = currentLogoMatch ? currentLogoMatch[1].split("?")[0] : null;
        const expectedLogoMatch = rawMatchedSignature.match(logoRegex);
        let expectedLogoUrl = expectedLogoMatch ? expectedLogoMatch[1].split("?")[0] : null;

        const isTextValid = cleanCurrentSignature === cleanFetchedSignature;
        const isLogoValid =
          !expectedLogoUrl || (currentLogoUrl && expectedLogoUrl && currentLogoUrl === expectedLogoUrl);

        displayNotification(
          "Info",
          `currentLogoUrl: ${currentLogoUrl.length || "null"},
          expectedLogoUrl: ${expectedLogoUrl.length || "null"},
          currentSignature: ${currentSignature.length || "null"},
          rawMatchedSignature: ${rawMatchedSignature.length || "null"},
          cleanCurrentSignature: ${cleanCurrentSignature.length || "null"},
          cleanFetchedSignature: ${cleanFetchedSignature.length || "null"}`
        );

        // displayNotification(
        //   "Info",
        //   `validateSignatureChanges:
        //   isTextValid: ${isTextValid},
        //   isLogoValid: ${isLogoValid}`
        // );

        if (isTextValid && isLogoValid) {
          displayNotification("Info", "validateSignatureChanges: Signature valid, allowing send");
          event.completed({ allowEvent: true });
        } else {
          // restore the signature
          displayError(
            "Selected M3 email signature has been modified. M3 email signature is prohibited from modification. The original signature has been restored.",
            event
          );
          const signatureWithMarker = "<!-- signature -->" + rawMatchedSignature.trim();
          item.body.setSignatureAsync(
            signatureWithMarker,
            { coercionType: Office.CoercionType.Html },
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                displayError(`Failed to apply ${signatureKey}.`, event);
                return;
              }
              displayNotification("Info", "validateSignatureChanges: Restore succeeded, displaying modified alert");
              event.completed();
              return;
            }
          );
        }
      });
      return; // Exit early for async handling in Classic Outlook
    }

    // Non-Classic Outlook (OWA, New Outlook, mobile) uses existing storage-based logic
    const originalSignatureKey = storageGetItem("tempSignature");
    const rawMatchedSignature = storageGetItem(`signature_${originalSignatureKey}`);

    const cleanCurrentSignature = SignatureManager.normalizeSignature(currentSignature);
    const cleanCachedSignature = SignatureManager.normalizeSignature(rawMatchedSignature);

    const logoRegex = /<img[^>]+src=["'](.*?(?:m3signatures\/logo\/[^"']+))["'][^>]*>/i;

    const currentLogoMatch = currentSignature.match(logoRegex);
    let currentLogoUrl = currentLogoMatch ? currentLogoMatch[1].split("?")[0] : null;

    const expectedLogoMatch = rawMatchedSignature ? rawMatchedSignature.match(logoRegex) : null;
    let expectedLogoUrl = expectedLogoMatch ? expectedLogoMatch[1].split("?")[0] : null;

    const isTextValid = cleanCurrentSignature === cleanCachedSignature;
    const isLogoValid = !expectedLogoUrl || currentLogoUrl === expectedLogoUrl;

    logger.log("debug", "validateSignatureChanges", {
      rawCurrentSignatureLength: currentSignature.length,
      rawMatchedSignatureLength: rawMatchedSignature ? rawMatchedSignature.length : 0,
      cleanCurrentSignature,
      cleanCachedSignature,
      originalSignatureKey,
      isReplyOrForward,
      currentLogoUrl,
      expectedLogoUrl,
      isTextValid,
      isLogoValid,
    });

    if (isTextValid && isLogoValid) {
      storageRemoveItem("tempSignature");
      event.completed({ allowEvent: true });
    } else {
      SignatureManager.restoreSignature(item, rawMatchedSignature, originalSignatureKey, (restored, error) => {
        if (error || !restored) {
          logger.log("error", "validateSignatureChanges", { error: error?.message || "Restore failed" });
          displayError("Failed to restore the original M3 signature. Please reselect.", event);
        } else {
          logger.log("info", "validateSignatureChanges", { status: "Signature restored successfully" });
          displayError(
            "Selected M3 email signature has been modified. M3 email signature is prohibited from modification. The original signature has been restored.",
            event
          );
        }
        event.completed({ allowEvent: false });
      });
    }
  } catch (error) {
    displayNotification("Error", `validateSignatureChanges: Exception - ${error.message}`);
    logger.log("error", "validateSignatureChanges", { error: error.message, stack: error.stack });
    displayError("An unexpected error occurred during signature validation.", event);
  }
}

/**
 * Handles the new message compose event, applying the appropriate signature for reply/forward or new messages.
 * @param {Object} event - The event object from Office.js.
 */
function onNewMessageComposeHandler(event) {
  isMobile =
    Office.context?.mailbox?.diagnostics?.hostName === "OutlookAndroid" ||
    Office.context?.mailbox?.diagnostics?.hostName === "OutlookIOS";

  isClassicOutlook = Office.context?.mailbox?.diagnostics?.hostName === "Outlook";

  logger.log("info", "Office.onReady", {
    host: Office.context?.mailbox?.diagnostics?.hostName,
    version: Office.context?.mailbox?.diagnostics?.hostVersion,
    isMobile,
    isClassicOutlook,
  });

  const item = Office.context.mailbox.item;

  displayNotification(
    `Info`,
    `Platform: ${Office.context.mailbox.diagnostics.hostName}, Version: ${Office.context.mailbox.diagnostics.hostVersion}`
  );
  SignatureManager.isReplyOrForward(item, (isReplyOrForward, error) => {
    if (error) {
      logger.log("error", "onNewMessageComposeHandler", { error: error.message });
      completeWithState(event, "Error", "Failed to determine reply/forward status.");
      return;
    }

    if (isReplyOrForward) {
      logger.log("info", "onNewMessageComposeHandler", { status: "Processing reply/forward email" });

      let messageId;
      if (isMobile) {
        messageId = item.conversationId;
        processEmailId(messageId, event);
      } else {
        item.getItemIdAsync((itemIdResult) => {
          if (itemIdResult.status !== Office.AsyncResultStatus.Succeeded) {
            logger.log("error", "onNewMessageComposeHandler", { error: itemIdResult.error.message });
            completeWithState(event, "Error", itemIdResult.error.message);
            return;
          }
          messageId = itemIdResult.value;
          logger.log("info", "getItemIdAsync for OWA/Classic", { messageId });
          processEmailId(messageId, event);
        });
      }
    } else {
      if (isMobile) {
        const mobileDefaultSignatureKey = storageGetItem("mobileDefaultSignature");
        if (mobileDefaultSignatureKey) {
          storageRemoveItem("tempSignature");
          storageSetItem("tempSignature", mobileDefaultSignatureKey);
          addSignature(mobileDefaultSignatureKey, event, true, () => {
            completeWithState(event, null, null);
          });
        } else {
          completeWithState(event, "Info", "Please select an M3 signature from the task pane.");
        }
      } else {
        completeWithState(event, "Info", "Please select an M3 signature from the ribbon.");
      }
    }
  });
}

/**
 * Processes the email ID by fetching the email and handling the signature.
 * @param {string} messageId - The ID of the email to process.
 * @param {Office.AddinCommands.Event} event - The event object.
 */
function processEmailId(messageId, event) {
  fetchMessageById(messageId, (message, fetchError) => {
    if (fetchError) {
      logger.log("error", "onNewMessageComposeHandler", { error: fetchError.message });
      completeWithState(event, "Error", fetchError.message);
      return;
    }

    const emailBody = message.body?.content || "";
    const extractedSignature = SignatureManager.extractSignature(emailBody);

    if (!extractedSignature) {
      logger.log("warn", "onNewMessageComposeHandler", { status: "No signature found in email" });
      completeWithState(
        event,
        "Info",
        isMobile
          ? "No signature found in email. Please select an M3 signature from the task pane."
          : "No signature found in email. Please select an M3 signature from the ribbon."
      );
      return;
    }

    logger.log("info", "onNewMessageComposeHandler", {
      status: "Signature extracted from email",
      signatureLength: extractedSignature.length,
    });

    const matchedSignatureKey = detectSignatureKey(extractedSignature);
    if (!matchedSignatureKey) {
      logger.log("warn", "onNewMessageComposeHandler", { status: "Could not detect signature key" });
      completeWithState(
        event,
        "Info",
        isMobile
          ? "Could not detect signature type. Please select an M3 signature from the task pane."
          : "Could not detect signature type. Please select an M3 signature from the ribbon."
      );
      return;
    }

    logger.log("info", "onNewMessageComposeHandler", {
      status: "Detected signature key from content",
      matchedSignatureKey,
    });

    storageRemoveItem("tempSignature");
    storageSetItem("tempSignature", matchedSignatureKey);
    addSignature(matchedSignatureKey, event, true, () => {
      completeWithState(event, null, null);
    });
  });
}

/**
 * Adds the Mona signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureMona(event) {
  addSignature("monaSignature", event, false, () => {});
}

/**
 * Adds the Morgan signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureMorgan(event) {
  addSignature("morganSignature", event, false, () => {});
}

/**
 * Adds the Morven signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureMorven(event) {
  addSignature("morvenSignature", event, false, () => {});
}

/**
 * Adds the M2 signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureM2(event) {
  addSignature("m2Signature", event, false, () => {});
}

/**
 * Adds the M3 signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureM3(event) {
  addSignature("m3Signature", event, false, () => {});
}

Office.actions.associate("addSignatureMona", addSignatureMona);
Office.actions.associate("addSignatureMorgan", addSignatureMorgan);
Office.actions.associate("addSignatureMorven", addSignatureMorven);
Office.actions.associate("addSignatureM2", addSignatureM2);
Office.actions.associate("addSignatureM3", addSignatureM3);
Office.actions.associate("validateSignature", validateSignature);
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);

let isMobile = false;
let isClassicOutlook = false;
