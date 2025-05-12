/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// Detect Android
var isAndroid = navigator.userAgent.toLowerCase().indexOf("android") > -1;
var initialEnvironment = isAndroid ? "mobile" : "desktop";

// Base URL for logging
const logBaseUrl = "https://m3wind-logger-aahvb3ckgmf9e2h2.uksouth-01.azurewebsites.net/api/LogEndpoint";

// Log function to send debug info to server
function log(event, data) {
  const logMessage = {
    event,
    data,
    environment: initialEnvironment,
    timestamp: new Date().toISOString(),
    userAgent: navigator.userAgent,
    url: window.location.href,
  };
  console.log("Add-in Log:", logMessage);

  if (typeof fetch !== "undefined") {
    fetch(logBaseUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(logMessage),
    })
      .then((response) => {
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        console.log("Log Success:", response.status);
      })
      .catch((error) => {
        console.error("Log Error:", error.message);
      });
  } else {
    console.error("Fetch unavailable");
  }
}

// Test fetch function
function testFetch(event) {
  log("testFetch", { status: "Triggered" });
  fetch(logBaseUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ event: "testFetch", data: { status: "Test" } }),
  })
    .then((response) => {
      log("testFetchSuccess", { status: response.status });
    })
    .catch((error) => {
      log("testFetchError", { error: error.message });
    });
  if (event) event.completed();
}

/**
 * Initializes the Outlook add-in and associates event handlers.
 */
Office.onReady((info) => {
  log("commandsInitialized", {
    host: Office.context?.mailbox?.diagnostics?.hostName,
    platform: info?.host,
    version: info?.version,
    isCompose: !!Office.context?.mailbox?.item?.isCompose,
  });

  // Test log to confirm execution
  log("mobileTestLog", { status: "Commands.js loaded" });

  // Network test log
  fetch(logBaseUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ event: "networkTest", data: { status: "Ping" } }),
  })
    .then(() => log("networkTest", { status: "Server reachable" }))
    .catch((error) => log("networkTestError", { error: error.message }));

  // Associate function commands
  Office.actions.associate("addSignatureMona", addSignatureMona);
  Office.actions.associate("addSignatureMorgan", addSignatureMorgan);
  Office.actions.associate("addSignatureMorven", addSignatureMorven);
  Office.actions.associate("addSignatureM2", addSignatureM2);
  Office.actions.associate("addSignatureM3", addSignatureM3);
  Office.actions.associate("applyDefaultSignature", applyDefaultSignature);
  Office.actions.associate("cancelAction", cancelAction);
  Office.actions.associate("validateSignature", validateSignature);
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
});

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
      console.error({ event: "displayNotification", error: "No mailbox item" });
      return;
    }

    const messageId = type === "Error" ? "Err" : "Info";
    const notificationType =
      type === "Error"
        ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
        : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;

    const isPersistent = persistent === true ? true : false;
    console.log({ event: "displayNotification", type, message, isPersistent, status: "Skipped" });
  } catch (error) {
    console.error({ event: "displayNotification", error: error.message });
  }
}

/**
 * Restores the signature using a Promise-based approach.
 * @param {Office.MessageCompose} item - The email item.
 * @param {string} cachedSignature - The signature to restore.
 * @param {string} signatureKey - The signature key.
 * @returns {Promise<boolean>} True if restoration succeeds, false otherwise.
 */
function restoreSignatureAsync(item, cachedSignature, signatureKey) {
  return new Promise((resolve) => {
    console.log({ event: "restoreSignatureAsync", signatureKey, cachedSignatureLength: cachedSignature?.length });
    if (!cachedSignature) {
      console.error({ event: "restoreSignatureAsync", error: "No cached signature", signatureKey });
      resolve(false);
      return;
    }

    item.body.setSignatureAsync(
      "<!-- signature -->" + cachedSignature,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error({ event: "restoreSignatureAsync", error: asyncResult.error.message, signatureKey });
          resolve(false);
        } else {
          console.log({ event: "restoreSignatureAsync", status: "Signature restored", signatureKey });
          resolve(true);
        }
      }
    );
  });
}

/**
 * Displays an error with a Smart Alert and notification.
 * @param {string} message - Error message.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 * @param {boolean} restoreSignature - Whether to restore the original signature.
 * @param {string} signatureKey - The signature key to restore.
 * @param {string} tempSignature - Temporary signature for new emails (optional).
 */
async function displayError(message, event, restoreSignature = false, signatureKey = null, tempSignature = null) {
  console.log({
    event: "displayError",
    message,
    restoreSignature,
    signatureKey,
    tempSignatureLength: tempSignature?.length,
  });

  const markdownMessage = message.includes("modified")
    ? `${message}\n\n**Tip**: Ensure the M3 signature is not edited before sending.`
    : `${message}\n\n**Tip**: Select an M3 signature from the ribbon under "M3 Signatures".`;

  const item = Office.context.mailbox.item;
  if (!item) {
    console.error({ event: "displayError", error: "No mailbox item" });
    displayNotification("Error", message, true);
    event.completed({
      allowEvent: false,
      errorMessage: message,
      errorMessageMarkdown: markdownMessage,
      cancelLabel: "OK",
    });
    return;
  }

  if (restoreSignature) {
    let signatureToRestore = tempSignature;
    if (signatureKey && !signatureToRestore) {
      signatureToRestore = localStorage.getItem(`signature_${signatureKey}`);
    }

    if (!signatureToRestore) {
      console.error({ event: "displayError", error: "No signature to restore", signatureKey });
      displayNotification("Error", message, true);
      event.completed({
        allowEvent: false,
        errorMessage: message,
        errorMessageMarkdown: markdownMessage,
        cancelLabel: "OK",
      });
      return;
    }

    const restored = await restoreSignatureAsync(item, signatureToRestore, signatureKey || "tempSignature");
    if (!restored) {
      console.error({ event: "displayError", error: "Failed to restore signature", signatureKey });
      displayNotification("Error", "Failed to restore signature.", true);
      event.completed({
        allowEvent: false,
        errorMessage: "Failed to restore signature.",
        errorMessageMarkdown: "Failed to restore signature.\n\n**Tip**: Select an M3 signature from the ribbon.",
        cancelLabel: "OK",
      });
      return;
    }

    displayNotification("Error", message, true);
    event.completed({
      allowEvent: false,
      errorMessage: message,
      errorMessageMarkdown: markdownMessage,
      cancelLabel: "OK",
    });
  } else {
    displayNotification("Error", message, true);
    event.completed({
      allowEvent: false,
      errorMessage: message,
      errorMessageMarkdown: markdownMessage,
      cancelLabel: "OK",
    });
  }
}

/**
 * Applies the default M3 signature and allows sending for restored signatures.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function applyDefaultSignature(event) {
  console.log({ event: "applyDefaultSignature" });
  const item = Office.context.mailbox.item;
  item.body.getAsync("html", (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error({ event: "applyDefaultSignature", error: result.error.message });
      displayNotification("Error", "Failed to load email body.", true);
      event.completed();
      return;
    }

    const body = result.value;
    const isClassicOutlook = Office.context.mailbox.diagnostics.hostName === "Outlook";
    const currentSignature = isClassicOutlook ? extractSignatureForOutlookClassic(body) : extractSignature(body);

    getSignatureKeyForRecipients(item).then((signatureKey) => {
      if (!signatureKey) {
        console.log({ event: "applyDefaultSignature", status: "No signature key found, applying default" });
        addSignature("m3Signature", event);
        return;
      }

      const cachedSignature = localStorage.getItem(`signature_${signatureKey}`);
      if (!cachedSignature) {
        console.log({ event: "applyDefaultSignature", status: "No cached signature, applying default" });
        addSignature("m3Signature", event);
        return;
      }

      const cleanCurrentSignature = normalizeSignature(currentSignature);
      const cleanStoredSignature = normalizeSignature(cachedSignature);

      if (cleanCurrentSignature === cleanStoredSignature) {
        console.log({ event: "applyDefaultSignature", status: "Signature matches, allowing send", signatureKey });
        event.completed({ allowEvent: true });
      } else {
        console.log({ event: "applyDefaultSignature", status: "Signature mismatch, applying default" });
        addSignature("m3Signature", event);
      }
    });
  });
}

/**
 * Cancels the Smart Alert action.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function cancelAction(event) {
  console.log({ event: "cancelAction" });
  event.completed({ allowEvent: false });
}

/**
 * Checks if the email is a reply or forward.
 * @param {Office.MessageCompose} item - The email item.
 * @returns {Promise<boolean>} True if reply or forward.
 */
function checkForReplyOrForward(item) {
  return new Promise((resolve, reject) => {
    console.log({ event: "checkForReplyOrForward" });
    if (item.itemType === Office.MailboxEnums.ItemType.Message && item.conversationId) {
      console.log({
        event: "checkForReplyOrForward",
        status: "Reply/forward detected",
        conversationId: item.conversationId,
      });
      resolve(true);
      return;
    }
    if (item.inReplyTo) {
      console.log({ event: "checkForReplyOrForward", status: "Reply detected", inReplyTo: item.inReplyTo });
      resolve(true);
      return;
    }
    item.subject.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const isReplyOrForward =
          result.value.toLowerCase().includes("re:") ||
          result.value.toLowerCase().includes("fw:") ||
          result.value.toLowerCase().includes("fwd:");
        console.log({ event: "checkForReplyOrForward", status: "Subject checked", isReplyOrForward });
        resolve(isReplyOrForward);
      } else {
        console.error({ event: "checkForReplyOrForward", error: result.error.message });
        reject(new Error("Failed to get subject"));
      }
    });
  });
}

/**
 * Validates the email signature on send.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function validateSignature(event) {
  console.log({ event: "validateSignature" });
  try {
    const item = Office.context.mailbox.item;
    isExternalEmail(item)
      .then((isExternal) => {
        console.log({ event: "validateSignature", isExternal });
        checkForReplyOrForward(item)
          .then((isReplyOrForward) => {
            console.log({ event: "validateSignature", isReplyOrForward });
            item.body.getAsync("html", (result) => {
              if (result.status !== Office.AsyncResultStatus.Succeeded) {
                console.error({ event: "validateSignature", error: result.error.message });
                displayError("Failed to load email body.", event);
                return;
              }

              const body = result.value;
              const isClassicOutlook = Office.context.mailbox.diagnostics.hostName === "Outlook";
              const currentSignature = isClassicOutlook
                ? extractSignatureForOutlookClassic(body)
                : extractSignature(body);

              if (!currentSignature) {
                console.log({ event: "validateSignature", status: "No signature found" });
                displayError(
                  "Email is missing the M3 required signature. Please select an appropriate email signature.",
                  event
                );
              } else {
                validateSignatureChanges(item, currentSignature, event, isReplyOrForward);
              }
            });
          })
          .catch((error) => {
            console.error({ event: "validateSignature", error: error.message });
            displayError("Failed to detect reply/forward status.", event);
          });
      })
      .catch((error) => {
        console.error({ event: "validateSignature", error: error.message });
        displayError("Failed to check external email status.", event);
      });
  } catch (error) {
    console.error({ event: "validateSignature", error: error.message });
    displayError("Unexpected error validating signature.", event);
  }
}

/**
 * Validates if the signature has been modified or changed.
 * @param {Office.MessageCompose} item - The email item.
 * @param {string} currentSignature - The current signature in the email body.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 * @param {boolean} isReplyOrForward - Whether the email is a reply/forward.
 */
async function validateSignatureChanges(item, currentSignature, event, isReplyOrForward) {
  try {
    item.body.getAsync("html", (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error({ event: "validateSignatureChanges", error: result.error.message });
        displayError("Failed to load email body.", event);
        return;
      }

      const newBody = result.value;
      const isClassicOutlook = Office.context.mailbox.diagnostics.hostName === "Outlook";
      const newSignature = isClassicOutlook ? extractSignatureForOutlookClassic(newBody) : extractSignature(newBody);

      if (!newSignature) {
        console.log({ event: "validateSignatureChanges", status: "Missing signature" });
        displayError(
          "Email is missing the M3 required signature. Please select an appropriate email signature.",
          event
        );
        return;
      }

      const cleanNewSignature = normalizeSignature(newSignature);
      const signatureKeys = ["monaSignature", "morganSignature", "morvenSignature", "m2Signature", "m3Signature"];
      let matchedSignatureKey = null;

      for (const key of signatureKeys) {
        const cachedSignature = localStorage.getItem(`signature_${key}`);
        if (cachedSignature && cleanNewSignature === normalizeSignature(cachedSignature)) {
          matchedSignatureKey = key;
          console.log({ event: "validateSignatureChanges", status: "Matched signature", matchedSignatureKey });
          break;
        }
      }

      if (matchedSignatureKey) {
        console.log({ event: "validateSignatureChanges", status: "Signature valid", matchedSignatureKey });
        if (!isReplyOrForward) {
          localStorage.removeItem("tempSignature_new");
          console.log({ event: "validateSignatureChanges", status: "Cleared temporary signature for new email" });
        }
        saveSignatureData(item, matchedSignatureKey).then(() => {
          event.completed({ allowEvent: true });
        });
      } else {
        console.log({ event: "validateSignatureChanges", status: "Signature modified" });
        if (isReplyOrForward) {
          getSignatureKeyForRecipients(item).then((signatureKey) => {
            if (signatureKey) {
              displayError(
                "Selected M3 signature has been modified. Restoring the original signature.",
                event,
                true,
                signatureKey
              );
            } else {
              console.log({
                event: "validateSignatureChanges",
                status: "No signatureKey for reply/forward, prompting re-selection",
              });
              displayError(
                "Selected M3 signature has been modified. Please select an appropriate email signature.",
                event,
                false
              );
            }
          });
        } else {
          const tempSignature = localStorage.getItem("tempSignature_new");
          if (tempSignature) {
            console.log({ event: "validateSignatureChanges", status: "Restoring temporary signature for new email" });
            displayError(
              "Selected M3 signature has been modified. Restoring the original signature.",
              event,
              true,
              null,
              tempSignature
            );
          } else {
            console.log({ event: "validateSignatureChanges", status: "No temporary signature found for new email" });
            displayError(
              "Selected M3 signature has been modified. Restoring the original signature.",
              event,
              true,
              null,
              localStorage.getItem(`signature_${signatureKeys[0]}`)
            );
          }
        }
      }
    });
  } catch (error) {
    console.error({ event: "validateSignatureChanges", error: error.message });
    displayError("Unexpected error validating signature changes.", event);
  }
}

/**
 * Extracts the signature from the email body.
 * @param {string} body - The email body HTML.
 * @returns {string|null} The extracted signature or null.
 */
function extractSignature(body) {
  console.log({ event: "extractSignature" });
  const marker = "<!-- signature -->";
  const startIndex = body.indexOf(marker);
  if (startIndex !== -1) {
    return body.slice(startIndex + marker.length).trim();
  }

  const signatureDivRegex = /<div\s+id="Signature">(.*?)<\/table>/s;
  const match = body.match(signatureDivRegex);
  return match ? match[1] : null;
}

/**
 * Extracts the signature for classic Outlook.
 * @param {string} body - The email body HTML.
 * @returns {string|null} The extracted signature or null.
 */
function extractSignatureForOutlookClassic(body) {
  console.log({ event: "extractSignatureForOutlookClassic" });
  const marker = "<!-- signature -->";
  const startIndex = body.indexOf(marker);
  if (startIndex !== -1) {
    return body.slice(startIndex + marker.length).trim();
  }

  const signatureDivRegex = /<table\s+class=MsoNormalTable[^>]*>(.*?)<\/table>/is;
  const match = body.match(signatureDivRegex);
  return match ? match[1] : null;
}

/**
 * Normalizes a signature for comparison.
 * @param {string} sig - The signature HTML.
 * @returns {string} The normalized signature.
 */
function normalizeSignature(sig) {
  if (!sig) return "";
  return sig
    .replace(/<[^>]*>?/gm, "")
    .replace(/\s+/g, "")
    .toLowerCase();
}

/**
 * Normalizes a subject for comparison.
 * @param {string} subject - The email subject.
 * @returns {string} The normalized subject.
 */
function normalizeSubject(subject) {
  if (!subject) return "";
  return subject
    .replace(/^(re:|fw:|fwd:)\s*/i, "")
    .trim()
    .toLowerCase();
}

/**
 * Checks if the email is external.
 * @param {Office.MessageCompose} item - The email item.
 * @returns {Promise<boolean>} True if external.
 */
function isExternalEmail(item) {
  return new Promise((resolve) => {
    console.log({ event: "isExternalEmail" });
    const isClassicOutlook = Office.context.mailbox.diagnostics.hostName === "Outlook";
    resolve(!isClassicOutlook && item.inReplyTo && item.inReplyTo.indexOf("OUTLOOK.COM") === -1);
  });
}

/**
 * Fetches a signature from the API.
 * @param {string} signatureKey - The signature key (e.g., "m3Signature").
 * @param {function} callback - Callback with (template, error).
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
          let template = data.result;
          template = template.replace("{First name} ", Office.context.mailbox.userProfile.displayName || "");
          template = template.replace("{Last name}", "");
          template = template.replaceAll("{E-mail}", Office.context.mailbox.userProfile.emailAddress || "");
          template = template.replace("{Title}", "");
          callback(template, null);
        })
        .catch((err) => callback(null, err));
    })
    .catch((err) => callback(null, err));
}

/**
 * Finds the signature key by matching conversationId, recipient emails, and subject in localStorage.
 * @param {Office.MessageCompose} item - The email item.
 * @returns {Promise<string|null>} The signature key or null if no match.
 */
function getSignatureKeyForRecipients(item) {
  return new Promise((resolve) => {
    item.to.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error({ event: "getSignatureKeyForRecipients", error: result.error.message });
        resolve(null);
        return;
      }

      const recipients = result.value.map((recipient) => recipient.emailAddress.toLowerCase());
      const conversationId = item.conversationId || null;

      item.subject.getAsync((subjectResult) => {
        if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error({ event: "getSignatureKeyForRecipients", error: subjectResult.error.message });
          resolve(null);
          return;
        }

        const currentSubject = normalizeSubject(subjectResult.value);
        console.log({ event: "getSignatureKeyForRecipients", recipients, conversationId, currentSubject });

        const signatureDataEntries = [];
        for (let i = 0; i < localStorage.length; i++) {
          const key = localStorage.key(i);
          if (key.startsWith("signatureData_")) {
            try {
              const data = JSON.parse(localStorage.getItem(key));
              signatureDataEntries.push({ key, data });
            } catch (error) {
              console.error({ event: "getSignatureKeyForRecipients", error: error.message, key });
            }
          }
        }
        console.log({
          event: "getSignatureKeyForRecipients",
          signatureDataEntries: signatureDataEntries.map((entry) => ({
            key: entry.key,
            conversationId: entry.data.conversationId,
            recipients: entry.data.recipients,
            subject: entry.data.subject,
            signature: entry.data.signature,
            timestamp: entry.data.timestamp,
          })),
        });

        signatureDataEntries.sort((a, b) => new Date(b.data.timestamp) - new Date(a.data.timestamp));

        let signatureKey = null;

        if (conversationId) {
          for (const entry of signatureDataEntries) {
            const data = entry.data;
            if (data.conversationId === conversationId && data.signature !== "none") {
              signatureKey = data.signature;
              console.log({
                event: "getSignatureKeyForRecipients",
                status: "Found matching signature by conversationId",
                signatureKey,
                key: entry.key,
                storedSubject: data.subject,
                storedRecipients: data.recipients,
              });
              break;
            }
          }
        }

        if (!signatureKey) {
          for (const entry of signatureDataEntries) {
            const data = entry.data;
            const storedRecipients = data.recipients.map((email) => email.toLowerCase());
            const storedSubject = normalizeSubject(data.subject);
            if (
              recipients.some((recipient) => storedRecipients.includes(recipient)) &&
              storedSubject === currentSubject &&
              data.signature !== "none"
            ) {
              signatureKey = data.signature;
              console.log({
                event: "getSignatureKeyForRecipients",
                status: "Found matching signature by recipients and subject",
                signatureKey,
                key: entry.key,
                storedSubject,
                storedRecipients,
              });
              break;
            }
          }
        }

        if (!signatureKey) {
          for (const entry of signatureDataEntries) {
            const data = entry.data;
            const storedRecipients = data.recipients.map((email) => email.toLowerCase());
            if (recipients.some((recipient) => storedRecipients.includes(recipient)) && data.signature !== "none") {
              signatureKey = data.signature;
              console.log({
                event: "getSignatureKeyForRecipients",
                status: "Found matching signature by recipients only",
                signatureKey,
                key: entry.key,
                storedSubject: data.subject,
                storedRecipients,
              });
              break;
            }
          }
        }

        console.log({ event: "getSignatureKeyForRecipients", selectedSignatureKey: signatureKey });
        resolve(signatureKey);
      });
    });
  });
}

/**
 * Adds a signature to the email and saves it to localStorage.
 * @param {string} signatureKey - The signature key (e.g., "m3Signature").
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 * @param {boolean} isAutoApplied - Whether the signature is auto-applied.
 */
function addSignature(signatureKey, event, isAutoApplied = false) {
  console.log({ event: "addSignature", signatureKey, isAutoApplied });

  try {
    const item = Office.context.mailbox.item;
    displayNotification("Info", `Applying ${signatureKey}...`, false);

    const cachedSignature = localStorage.getItem(`signature_${signatureKey}`);
    if (cachedSignature && !isAutoApplied) {
      item.body.setSignatureAsync(
        "<!-- signature -->" + cachedSignature,
        { coercionType: Office.CoercionType.Html },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error({ event: "addSignature", error: asyncResult.error.message });
            displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
            if (!isAutoApplied) {
              event.completed();
            } else {
              displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
              saveSignatureData(item, "none");
              event.completed();
            }
          } else {
            console.log({ event: "addSignature", status: "Signature applied from cache", signatureKey });
            displayNotification("Info", `${signatureKey} applied.`, false);
            saveSignatureData(item, signatureKey);
            if (!isAutoApplied) {
              localStorage.setItem("tempSignature_new", cachedSignature);
              console.log({ event: "addSignature", status: "Stored temporary signature for new email" });
            }
            event.completed();
          }
        }
      );
    } else {
      fetchSignature(signatureKey, (template, error) => {
        if (error) {
          console.error({ event: "addSignature", error: error.message });
          displayNotification("Error", `Failed to fetch ${signatureKey}.`, true);
          if (!isAutoApplied) {
            event.completed();
          } else {
            displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
            saveSignatureData(item, "none");
            event.completed();
          }
          return;
        }

        item.body.setSignatureAsync(
          "<!-- signature -->" + template,
          { coercionType: Office.CoercionType.Html },
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              console.error({ event: "addSignature", error: asyncResult.error.message });
              displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
              if (!isAutoApplied) {
                event.completed();
              } else {
                displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
                saveSignatureData(item, "none");
                event.completed();
              }
            } else {
              console.log({ event: "addSignature", status: "Signature applied", signatureKey });
              localStorage.setItem(`signature_${signatureKey}`, template);
              displayNotification("Info", `${signatureKey} applied.`, false);
              saveSignatureData(item, signatureKey);
              if (!isAutoApplied) {
                localStorage.setItem("tempSignature_new", template);
                console.log({ event: "addSignature", status: "Stored temporary signature for new email" });
              }
              event.completed();
            }
          }
        );
      });
    }
  } catch (error) {
    console.error({ event: "addSignature", error: error.message });
    displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
    if (!isAutoApplied) {
      event.completed();
    } else {
      displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
      saveSignatureData(item, "none");
      event.completed();
    }
  }
}

/**
 * Saves signature data to localStorage, including subject.
 * @param {Office.MessageCompose} item - The email item.
 * @param {string} signatureKey - The signature key.
 * @returns {Promise<object|null>} The saved data or null if failed.
 */
function saveSignatureData(item, signatureKey) {
  return new Promise((resolve) => {
    item.to.getAsync((result) => {
      let recipients = [];
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        recipients = result.value.map((recipient) => recipient.emailAddress.toLowerCase());
      } else {
        console.error({ event: "saveSignatureData", error: result.error.message });
      }

      const conversationId = item.conversationId || null;

      item.subject.getAsync((subjectResult) => {
        let subject = "";
        if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
          subject = subjectResult.value;
        } else {
          console.error({ event: "saveSignatureData", error: subjectResult.error.message });
        }

        console.log({ event: "saveSignatureData", signatureKey, recipients, conversationId, subject });

        let existingKey = null;
        if (conversationId) {
          for (let i = 0; i < localStorage.length; i++) {
            const key = localStorage.key(i);
            if (key.startsWith("signatureData_")) {
              try {
                const data = JSON.parse(localStorage.getItem(key));
                if (data.conversationId === conversationId) {
                  existingKey = key;
                  break;
                }
              } catch (error) {
                console.error({ event: "saveSignatureData", error: error.message, key });
              }
            }
          }
        }

        const data = {
          recipients,
          signature: signatureKey,
          conversationId,
          subject,
          timestamp: new Date().toISOString(),
        };

        if (existingKey) {
          localStorage.setItem(existingKey, JSON.stringify(data));
          console.log({ event: "saveSignatureData", status: "Updated existing entry", key: existingKey, subject });
        } else {
          const newKey = `signatureData_${Date.now()}`;
          localStorage.setItem(newKey, JSON.stringify(data));
          console.log({ event: "saveSignatureData", status: "Created new entry", key: newKey, subject });
        }
        resolve(data);
      });
    });
  });
}

/**
 * Adds the Mona signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureMona(event) {
  addSignature("monaSignature", event);
}

/**
 * Adds the Morgan signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureMorgan(event) {
  addSignature("morganSignature", event);
}

/**
 * Adds the Morven signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureMorven(event) {
  addSignature("morvenSignature", event);
}

/**
 * Adds the M2 signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureM2(event) {
  addSignature("m2Signature", event);
}

/**
 * Adds the M3 signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureM3(event) {
  addSignature("m3Signature", event);
}

/**
 * Handles new message compose event.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function onNewMessageComposeHandler(event) {
  console.log({ event: "onNewMessageComposeHandler" });

  const item = Office.context.mailbox.item;
  checkForReplyOrForward(item)
    .then((isReplyOrForward) => {
      console.log({ event: "onNewMessageComposeHandler", isReplyOrForward });

      if (isReplyOrForward) {
        getSignatureKeyForRecipients(item)
          .then((signatureKey) => {
            if (signatureKey) {
              console.log({
                event: "onNewMessageComposeHandler",
                status: "Auto-applying signature for reply/forward",
                signatureKey,
              });
              addSignature(signatureKey, event, true);
            } else {
              console.log({
                event: "onNewMessageComposeHandler",
                status: "No signature found for reply/forward, requiring manual selection",
              });
              displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
              saveInitialSignatureData(item);
              event.completed();
            }
          })
          .catch((error) => {
            console.error({ event: "onNewMessageComposeHandler", error: error.message });
            displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
            saveInitialSignatureData(item);
            event.completed();
          });
      } else {
        console.log({ event: "onNewMessageComposeHandler", status: "New email, requiring manual signature selection" });
        displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
        saveInitialSignatureData(item);
        localStorage.removeItem("tempSignature_new");
        console.log({ event: "onNewMessageComposeHandler", status: "Cleared temporary signature for new email" });
        event.completed();
      }
    })
    .catch((error) => {
      console.error({ event: "onNewMessageComposeHandler", error: error.message });
      displayNotification("Error", "Failed to detect reply/forward status.", true);
      saveInitialSignatureData(item);
      event.completed();
    });
}

/**
 * Saves initial signature data with "none" for new or reply/forward emails.
 * @param {Office.MessageCompose} item - The email item.
 */
function saveInitialSignatureData(item) {
  item.to.getAsync((result) => {
    let recipients = [];
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      recipients = result.value.map((recipient) => recipient.emailAddress.toLowerCase());
    } else {
      console.error({ event: "saveInitialSignatureData", error: result.error.message });
    }

    const conversationId = item.conversationId || null;

    item.subject.getAsync((subjectResult) => {
      let subject = "";
      if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
        subject = subjectResult.value;
      } else {
        console.error({ event: "saveInitialSignatureData", error: subjectResult.error.message });
      }

      const data = {
        recipients,
        signature: "none",
        conversationId,
        subject,
        timestamp: new Date().toISOString(),
      };

      const newKey = `signatureData_${Date.now()}`;
      localStorage.setItem(newKey, JSON.stringify(data));
      console.log({
        event: "saveInitialSignatureData",
        status: "Stored initial signature data",
        recipients,
        conversationId,
        subject,
      });
    });
  });
}
