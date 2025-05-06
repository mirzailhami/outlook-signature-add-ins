/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
 * Initializes the Outlook add-in and associates event handlers.
 */
Office.onReady(() => {
  console.log({
    event: "Office.onReady",
    host: Office.context?.mailbox?.diagnostics?.hostName,
    initialSignature: localStorage.getItem("initialSignature")?.slice(0, 50),
  });

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
 * Initializes auto-signature for new or reply/forward emails.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function initializeAutoSignature(event) {
  console.log({ event: "initializeAutoSignature" });
  try {
    const item = Office.context.mailbox.item;
    if (item) {
      applyAutoSignature(event);
    } else {
      console.error({ event: "initializeAutoSignature", error: "No mailbox item" });
      displayNotification("Error", "Failed to detect compose window.", true);
      event?.completed();
    }
  } catch (error) {
    console.error({ event: "initializeAutoSignature", error: error.message });
    displayNotification("Error", "Unexpected error initializing signature.", true);
    event?.completed();
  }
}

/**
 * Attempts to apply a signature with retries.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 * @param {number} attempt - Current attempt number.
 * @param {number} maxAttempts - Maximum retry attempts.
 */
function applyAutoSignature(event, attempt = 1, maxAttempts = 10) {
  console.log({ event: "applyAutoSignature", attempt });
  try {
    const item = Office.context.mailbox.item;

    checkForReplyOrForward(item)
      .then((isReplyOrForward) => {
        if (!isReplyOrForward) {
          console.log({ event: "applyAutoSignature", status: "New email, skipping auto-signature" });
          displayNotification("Info", "No signature applied for new email.", false);
          event?.completed();
          return;
        }

        displayNotification("Info", "Loading M3 signature...", false);
        const lastSignature = localStorage.getItem("initialSignature");
        if (!lastSignature) {
          console.log({ event: "applyAutoSignature", status: "No last signature found" });
          displayNotification("Info", "No previous M3 signature found. Please select a signature.", false);
          event?.completed();
          return;
        }

        item.body.getAsync("html", (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.error({ event: "applyAutoSignature", error: result.error.message });
            displayNotification("Error", "Failed to load email body.", true);
            event?.completed();
            return;
          }

          const body = result.value;
          if (!body.includes("<!-- signature -->")) {
            item.body.setSignatureAsync(
              "<!-- signature -->" + lastSignature,
              { coercionType: Office.CoercionType.Html },
              (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  console.error({ event: "applyAutoSignature", error: asyncResult.error.message });
                  displayNotification("Error", "Failed to apply M3 signature.", true);
                } else {
                  console.log({ event: "applyAutoSignature", status: "Signature applied" });
                  localStorage.setItem("initialSignature", lastSignature);
                  displayNotification("Info", "M3 signature applied.", false);
                }
                event?.completed();
              }
            );
          } else {
            console.log({ event: "applyAutoSignature", status: "Signature already present" });
            displayNotification("Info", "M3 signature already present.", false);
            event?.completed();
          }
        });
      })
      .catch((error) => {
        console.error({ event: "applyAutoSignature", error: error.message });
        displayNotification("Error", "Failed to detect reply/forward status.", true);
        event?.completed();
      });
  } catch (error) {
    console.error({ event: "applyAutoSignature", error: error.message });
    displayNotification("Error", "Unexpected error applying signature.", true);
    event?.completed();
  }
}

/**
 * Displays a notification in the Outlook UI.
 * @param {string} type - Notification type ("Error" or "Info").
 * @param {string} message - Notification message.
 * @param {boolean} persistent - Whether the notification persists.
 */
function displayNotification(type, message, persistent) {
  try {
    const item = Office.context.mailbox.item;
    if (!item) {
      console.error({ event: "displayNotification", error: "No mailbox item" });
      return;
    }

    const messageId = type === "Error" ? "SignatureError" : "SignatureInfo";
    const notificationType =
      type === "Error"
        ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
        : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;

    console.log({ event: "displayNotification", type, message, persistent, messageId });
    item.notificationMessages.replaceAsync(
      messageId,
      {
        type: notificationType,
        message,
        icon: "Icon.16x16",
        persistent,
      },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error({ event: "displayNotification", error: asyncResult.error.message });
        } else {
          console.log({ event: "displayNotification", status: "Success", message });
        }
      }
    );
  } catch (error) {
    console.error({ event: "displayNotification", error: error.message });
  }
}

/**
 * Displays an error with a Smart Alert and notification.
 * @param {string} message - Error message.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 * @param {boolean} restoreSignature - Whether to restore the original signature.
 */
function displayError(message, event, restoreSignature = false) {
  console.log({ event: "displayError", message, restoreSignature });

  const markdownMessage = restoreSignature
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
      cancelLabel: restoreSignature ? "Send Now" : "Apply Signature",
      commandId: "msgComposeApplyDefault",
      action: {
        actionText: "Cancel",
        actionType: "executeFunction",
        commandId: "msgComposeCancelAction",
      },
    });
    return;
  }

  if (restoreSignature) {
    const initialSignature = localStorage.getItem("initialSignature");
    if (!initialSignature) {
      console.error({ event: "displayError", error: "No signature to restore" });
      displayNotification("Error", message, true);
      event.completed({
        allowEvent: false,
        errorMessage: message,
        errorMessageMarkdown: markdownMessage,
        cancelLabel: "Send Now",
        commandId: "msgComposeApplyDefault",
        action: {
          actionText: "Cancel",
          actionType: "executeFunction",
          commandId: "msgComposeCancelAction",
        },
      });
      return;
    }

    item.body.setSignatureAsync(
      "<!-- signature -->" + initialSignature,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error({ event: "displayError", error: asyncResult.error.message });
          displayNotification("Error", "Failed to restore signature.", true);
          event.completed({
            allowEvent: false,
            errorMessage: "Failed to restore signature.",
            errorMessageMarkdown: "Failed to restore signature.\n\n**Tip**: Select an M3 signature from the ribbon.",
            cancelLabel: "Send Now",
            commandId: "msgComposeApplyDefault",
            action: {
              actionText: "Cancel",
              actionType: "executeFunction",
              commandId: "msgComposeCancelAction",
            },
          });
        } else {
          console.log({ event: "displayError", status: "Signature restored" });
          localStorage.setItem("initialSignature", initialSignature);
          displayNotification("Error", message, true);
          event.completed({
            allowEvent: false,
            errorMessage: message,
            errorMessageMarkdown: markdownMessage,
            cancelLabel: "Send Now",
            commandId: "msgComposeApplyDefault",
            action: {
              actionText: "Cancel",
              actionType: "executeFunction",
              commandId: "msgComposeCancelAction",
            },
          });
        }
      }
    );
  } else {
    displayNotification("Error", message, true);
    event.completed({
      allowEvent: false,
      errorMessage: message,
      errorMessageMarkdown: markdownMessage,
      cancelLabel: "Apply Signature",
      commandId: "msgComposeApplyDefault",
      action: {
        actionText: "Cancel",
        actionType: "executeFunction",
        commandId: "msgComposeCancelAction",
      },
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
    const initialSignature = localStorage.getItem("initialSignature");
    const initialSignatureFromStore = isClassicOutlook
      ? extractSignatureFromStoreForOutlookClassic(initialSignature)
      : extractSignatureFromStore(initialSignature);

    if (currentSignature && initialSignatureFromStore) {
      const cleanCurrentSignature = normalizeSignature(currentSignature);
      const cleanStoredSignature = normalizeSignature(initialSignatureFromStore);
      if (cleanCurrentSignature === cleanStoredSignature) {
        console.log({ event: "applyDefaultSignature", status: "Signature restored, allowing send" });
        event.completed({ allowEvent: true });
        return;
      }
    }

    addSignature("m3Signature", 4, event);
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
        status: "Detected via conversationId",
        conversationId: item.conversationId,
      });
      resolve(true);
      return;
    }
    if (item.inReplyTo) {
      console.log({ event: "checkForReplyOrForward", status: "Detected via inReplyTo", inReplyTo: item.inReplyTo });
      resolve(true);
      return;
    }
    item.subject.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const isReplyOrForward =
          result.value.toLowerCase().includes("re:") ||
          result.value.toLowerCase().includes("fw:") ||
          result.value.toLowerCase().includes("fwd:");
        console.log({
          event: "checkForReplyOrForward",
          status: "Detected via subject",
          subject: result.value,
          isReplyOrForward,
        });
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
              const initialSignature = isClassicOutlook
                ? extractSignatureForOutlookClassic(body)
                : extractSignature(body);

              if (isReplyOrForward && !initialSignature) {
                console.log({ event: "validateSignature", status: "No signature in reply/forward" });
                const lastSignature = localStorage.getItem("initialSignature");
                if (lastSignature) {
                  item.body.setSignatureAsync(
                    "<!-- signature -->" + lastSignature,
                    { coercionType: Office.CoercionType.Html },
                    (asyncResult) => {
                      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.error({ event: "validateSignature", error: asyncResult.error.message });
                        displayError("Failed to apply M3 signature.", event);
                      } else {
                        console.log({ event: "validateSignature", status: "Signature set for reply/forward" });
                        event.completed({ allowEvent: true });
                      }
                    }
                  );
                } else {
                  displayError("No M3 signature found for reply/forward. Please select a signature.", event);
                }
              } else if (!initialSignature) {
                displayError(
                  "Email is missing the M3 required signature. Please select an appropriate email signature.",
                  event
                );
              } else {
                validateSignatureChanges(item, initialSignature, event, isReplyOrForward);
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
 * Validates if the signature has been modified.
 * @param {Office.MessageCompose} item - The email item.
 * @param {string} initialSignature - The initial signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 * @param {boolean} isReplyOrForward - Whether the email is a reply/forward.
 */
function validateSignatureChanges(item, initialSignature, event, isReplyOrForward) {
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
      const initialSavedSignature = localStorage.getItem("initialSignature");
      const initialSignatureFromStore = isClassicOutlook
        ? extractSignatureFromStoreForOutlookClassic(initialSavedSignature)
        : extractSignatureFromStore(initialSavedSignature);

      if (!newSignature || !initialSignatureFromStore) {
        console.log({ event: "validateSignatureChanges", status: "Missing signature data" });
        displayError(
          "Email is missing the M3 required signature. Please select an appropriate email signature.",
          event
        );
        return;
      }

      const cleanNewSignature = normalizeSignature(newSignature);
      const cleanStoredSignature = normalizeSignature(initialSignatureFromStore);

      if (cleanNewSignature !== cleanStoredSignature) {
        console.log({ event: "validateSignatureChanges", status: "Signature modified" });
        displayError(
          "Selected M3 signature has been modified. M3 email signatures cannot be modified. Restoring the original signature.",
          event,
          true
        );
      } else {
        console.log({ event: "validateSignatureChanges", status: "Signature unchanged" });
        localStorage.setItem("initialSignature", initialSavedSignature);
        event.completed({ allowEvent: true });
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
 * Extracts the signature from local storage.
 * @param {string} body - The stored signature HTML.
 * @returns {string|null} The extracted signature or null.
 */
function extractSignatureFromStore(body) {
  console.log({ event: "extractSignatureFromStore" });
  if (!body) return null;
  const signatureDivRegex = /<div\s+class="Signature">(.*?)<\/div>/s;
  const match = body.match(signatureDivRegex);
  return match ? match[1] : body;
}

/**
 * Extracts the signature from local storage for classic Outlook.
 * @param {string} body - The stored signature HTML.
 * @returns {string|null} The extracted signature or null.
 */
function extractSignatureFromStoreForOutlookClassic(body) {
  console.log({ event: "extractSignatureFromStoreForOutlookClassic" });
  if (!body) return null;
  const signatureDivRegex = /<table class="MsoNormalTable"[^>]*>(.*?)<\/table>/is;
  const match = body.match(signatureDivRegex);
  return match ? match[1] : body;
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
 * Adds a signature to the email.
 * @param {string} signatureKey - The signature key (e.g., "m3Signature").
 * @param {number} signatureUrlIndex - The index for the signature URL.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignature(signatureKey, signatureUrlIndex, event) {
  console.log({ event: "addSignature", signatureKey });
  try {
    const item = Office.context.mailbox.item;
    displayNotification("Info", `Applying ${signatureKey}...`, false);
    const localTemplate = localStorage.getItem(signatureKey);
    if (localTemplate) {
      item.body.setSignatureAsync(
        "<!-- signature -->" + localTemplate,
        { coercionType: Office.CoercionType.Html },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error({ event: "addSignature", error: asyncResult.error.message });
            displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
            event.completed();
          } else {
            console.log({ event: "addSignature", status: "Signature applied", signatureKey });
            localStorage.setItem("initialSignature", localTemplate);
            localStorage.setItem("lastSentSignature", localTemplate);
            displayNotification("Info", `${signatureKey} applied.`, false);
            event.completed();
          }
        }
      );
    } else {
      const initialUrl = "https://m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net/api/Ribbons/ribbons";
      let signatureUrl =
        "https://m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net/api/Signatures/signatures?signatureURL=";
      fetch(initialUrl)
        .then((response) => response.json())
        .then((data) => {
          signatureUrl += data.result[signatureUrlIndex].url;
          fetch(signatureUrl)
            .then((response) => response.json())
            .then((data) => {
              let template = data.result;
              template = template.replace("{First name} ", Office.context.mailbox.userProfile.displayName || "");
              template = template.replace("{Last name}", "");
              template = template.replaceAll("{E-mail}", Office.context.mailbox.userProfile.emailAddress || "");
              template = template.replace("{Title}", "");
              item.body.setSignatureAsync(
                "<!-- signature -->" + template,
                { coercionType: Office.CoercionType.Html },
                (asyncResult) => {
                  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error({ event: "addSignature", error: asyncResult.error.message });
                    displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
                    event.completed();
                  } else {
                    console.log({ event: "addSignature", status: "Signature applied", signatureKey });
                    localStorage.setItem(signatureKey, template);
                    localStorage.setItem("initialSignature", template);
                    localStorage.setItem("lastSentSignature", template);
                    displayNotification("Info", `${signatureKey} applied.`, false);
                    event.completed();
                  }
                }
              );
            })
            .catch((err) => {
              console.error({ event: "addSignature", error: err.message });
              displayNotification("Error", "Failed to fetch signature.", true);
              event.completed();
            });
        })
        .catch((err) => {
          console.error({ event: "addSignature", error: err.message });
          displayNotification("Error", "Failed to fetch ribbons.", true);
          event.completed();
        });
    }
  } catch (error) {
    console.error({ event: "addSignature", error: error.message });
    displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
    event.completed();
  }
}

/**
 * Adds the Mona signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureMona(event) {
  addSignature("monaSignature", 0, event);
}

/**
 * Adds the Morgan signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureMorgan(event) {
  addSignature("morganSignature", 1, event);
}

/**
 * Adds the Morven signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureMorven(event) {
  addSignature("morvenSignature", 2, event);
}

/**
 * Adds the M2 signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureM2(event) {
  addSignature("m2Signature", 3, event);
}

/**
 * Adds the M3 signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureM3(event) {
  addSignature("m3Signature", 4, event);
}

/**
 * Handles new message compose event.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function onNewMessageComposeHandler(event) {
  console.log({ event: "onNewMessageComposeHandler" });
  // initializeAutoSignature(event);

  // Check whether a default signature is already set.
  const defaultSignature = localStorage.getItem("defaultSignature");
  if (!defaultSignature) {
    // Open the dialog to prompt the user to set their default signature.
    Office.context.ui.displayDialogAsync(
      "https://white-grass-0b6dc6e03.6.azurestaticapps.net/taskpane.html", // URL where the UI for signature selection lives
      { height: 50, width: 30 },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          var dialog = asyncResult.value;
          // Set up an event handler to receive messages from the dialog.
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
            console.log("Dialog message received: " + arg.message);
            // Assume the message is the ID of the default signature selected (e.g., "monaSignature")
            localStorage.setItem("defaultSignature", arg.message);
            // Close the dialog after the default signature is set.
            dialog.close();
            // Now that a default signature is set, if needed, you can insert it immediately:
            insertDefaultSignature(arg.message, event);
          });
        } else {
          console.error("Failed to open dialog: " + asyncResult.error.message);
          event.completed();
        }
      }
    );
  } else {
    // If a default is already set, apply it (if that’s your desired behavior)
    insertDefaultSignature(defaultSignature, event);
  }
}

// Helper function to insert the signature.
function insertDefaultSignature(signatureKey, event) {
  // This function works much like your existing addSignature(..)
  // You may want to reuse your addSignature function here.
  addSignature(signatureKey, 0, event);
}
