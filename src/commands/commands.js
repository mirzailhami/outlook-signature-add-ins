/* global Office */

import { getGraphAccessToken } from "./launchevent.js";
import { Client } from "@microsoft/microsoft-graph-client";
import {
  logger,
  SignatureManager,
  fetchSignature,
  getSignatureKeyForRecipients,
  saveSignatureData,
} from "./helpers.js";

// Mobile needs this initialization
Office.initialize = () => {
  logger.log(`info`, `Office.initialize`, { loaded: true });
};

Office.onReady(() => {
  logger.log("info", "Office.onReady", { host: Office.context?.mailbox?.diagnostics?.hostName });

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
 * Adds a signature to the email and saves it to localStorage.
 * @param {string} signatureKey - The signature key (e.g., "m3Signature").
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 * @param {boolean} isAutoApplied - Whether the signature is auto-applied.
 */
async function addSignature(signatureKey, event, isAutoApplied = false) {
  logger.log("info", "addSignature", { signatureKey, isAutoApplied });

  try {
    const item = Office.context.mailbox.item;
    const cachedSignature = localStorage.getItem(`signature_${signatureKey}`);

    if (cachedSignature && !isAutoApplied) {
      await new Promise((resolve) =>
        item.body.setSignatureAsync(
          "<!-- signature -->" + cachedSignature.trim(),
          { coercionType: Office.CoercionType.Html },
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              logger.log("error", "addSignature", { error: asyncResult.error.message });
              displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
              if (!isAutoApplied) event.completed();
              else {
                displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
                saveSignatureData(item, "none").then(() => event.completed());
              }
            } else {
              saveSignatureData(item, signatureKey).then(() => {
                if (!isAutoApplied) {
                  localStorage.setItem("tempSignature_new", cachedSignature);
                }
                event.completed();
              });
            }
            resolve();
          }
        )
      );
    } else {
      fetchSignature(signatureKey, async (template, error) => {
        if (error) {
          logger.log("error", "addSignature", { error: error.message });
          displayNotification("Error", `Failed to fetch ${signatureKey}.`, true);
          if (!isAutoApplied) event.completed();
          else {
            displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
            saveSignatureData(item, "none").then(() => event.completed());
          }
          return;
        }

        await new Promise((resolve) =>
          item.body.setSignatureAsync(
            "<!-- signature -->" + template.trim(),
            { coercionType: Office.CoercionType.Html },
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                logger.log("error", "addSignature", { error: asyncResult.error.message });
                displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
                if (!isAutoApplied) event.completed();
                else {
                  displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
                  saveSignatureData(item, "none").then(() => event.completed());
                }
              } else {
                localStorage.setItem(`signature_${signatureKey}`, template);
                saveSignatureData(item, signatureKey).then(() => {
                  if (!isAutoApplied) {
                    localStorage.setItem("tempSignature_new", template);
                  }
                  event.completed();
                });
              }
              resolve();
            }
          )
        );
      });
    }
  } catch (error) {
    logger.log("error", "addSignature", { error: error.message });
    displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
    if (!isAutoApplied) event.completed();
    else {
      displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
      saveSignatureData(item, "none").then(() => event.completed());
    }
  }
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
      notification.persistent = false;
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
 * @param {boolean} restoreSignature - Whether to restore the original signature.
 * @param {string} signatureKey - The signature key to restore (optional).
 * @param {string} tempSignature - Temporary signature for new emails (optional).
 */
async function displayError(message, event, restoreSignature = false, signatureKey = null, tempSignature = null) {
  logger.log("info", "displayError", { message, restoreSignature, signatureKey });

  const item = Office.context.mailbox.item;
  if (!item) {
    logger.log("error", "displayError", { error: "No mailbox item" });
    displayNotification("Error", message, true);
    event.completed({
      allowEvent: false,
      errorMessage: message,
      errorMessageMarkdown: `${message}\n\n**Tip**: Select an M3 signature from the ribbon under "M3 Signatures".`,
      cancelLabel: "OK",
    });
    return;
  }

  if (restoreSignature) {
    let signatureToRestore =
      tempSignature || localStorage.getItem("tempSignature_new") || localStorage.getItem("tempSignature_replyForward");
    if (signatureKey && !signatureToRestore) signatureToRestore = localStorage.getItem(`signature_${signatureKey}`);

    if (!signatureToRestore) {
      logger.log("error", "displayError", { error: "No signature to restore", signatureKey });
      displayNotification("Error", `${message} (Failed to restore: No signature available)`, true);
      event.completed({
        allowEvent: false,
        errorMessage: `${message} (Failed to restore: No signature available)`,
        errorMessageMarkdown: `${message} (Failed to restore: No signature available)\n\n**Note**: Failed to restore signature. Please reselect.`,
        cancelLabel: "OK",
      });
      return;
    }

    const restored = await SignatureManager.restoreSignature(item, signatureToRestore, signatureKey || "tempSignature");
    if (!restored) {
      logger.log("error", "displayError", { error: "Restoration failed", signatureKey });
      displayNotification("Error", `${message} (Failed to restore signature)`, true);
      event.completed({
        allowEvent: false,
        errorMessage: `${message} (Failed to restore signature)`,
        errorMessageMarkdown: `${message} (Failed to restore signature)\n\n**Note**: Failed to restore signature. Please reselect.`,
        cancelLabel: "OK",
      });
      return;
    }

    const customMessage =
      "Selected M3 email signature has been modified. M3 email signature is prohibited from modification. The original signature is now restored";
    displayNotification("Error", customMessage, true);
    event.completed({
      allowEvent: false,
      errorMessage: customMessage,
      errorMessageMarkdown: `${customMessage}\n\n**Tip**: Avoid modifying the M3 signature before sending.`,
      cancelLabel: "OK",
    });
  } else {
    displayNotification("Error", message, false);
    event.completed({
      allowEvent: false,
      errorMessage: message,
      errorMessageMarkdown: `${message}\n\n**Tip**: Select an M3 signature from the ribbon under "M3 Signatures".`,
      cancelLabel: "OK",
    });
  }
}

/**
 * Applies the default M3 signature and allows sending for restored signatures.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
async function applyDefaultSignature(event) {
  const item = Office.context.mailbox.item;
  const body = await new Promise((resolve) => item.body.getAsync("html", (result) => resolve(result.value)));
  const isClassicOutlook = Office.context.mailbox.diagnostics.hostName === "Outlook";
  const currentSignature = isClassicOutlook
    ? SignatureManager.extractSignatureForOutlookClassic(body)
    : SignatureManager.extractSignature(body);

  const signatureKey = await getSignatureKeyForRecipients(item);
  if (!signatureKey) {
    await addSignature("m3Signature", event);
    return;
  }

  const cachedSignature = localStorage.getItem(`signature_${signatureKey}`);
  if (!cachedSignature) {
    await addSignature("m3Signature", event);
    return;
  }

  const cleanCurrentSignature = SignatureManager.normalizeSignature(currentSignature);
  const cleanStoredSignature = SignatureManager.normalizeSignature(cachedSignature);

  if (cleanCurrentSignature === cleanStoredSignature) {
    event.completed({ allowEvent: true });
  } else {
    await addSignature("m3Signature", event);
  }
}

/**
 * Cancels the Smart Alert action.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function cancelAction(event) {
  event.completed({ allowEvent: false });
}

/**
 * Validates the email signature on send.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
async function validateSignature(event) {
  try {
    const item = Office.context.mailbox.item;
    if (!item) {
      logger.log("error", "validateSignature", { error: "No mailbox item" });
      displayError("No mailbox item available.", event);
      return;
    }

    const body = await new Promise((resolve) =>
      item.body.getAsync("html", (result) =>
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : null)
      )
    );
    const isClassicOutlook = Office.context.mailbox.diagnostics.hostName === "Outlook";
    const currentSignature = isClassicOutlook
      ? SignatureManager.extractSignatureForOutlookClassic(body)
      : SignatureManager.extractSignature(body);

    if (!currentSignature) {
      displayError("Email is missing the M3 required signature. Please select an appropriate email signature.", event);
    } else {
      const isReplyOrForward = await SignatureManager.isReplyOrForward(item);
      await validateSignatureChanges(item, currentSignature, event, isReplyOrForward);
    }
  } catch (error) {
    logger.log("error", "validateSignature", { error: error.message });
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
    const newBody = await new Promise((resolve) =>
      item.body.getAsync("html", (result) =>
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : null)
      )
    );
    const isClassicOutlook = Office.context.mailbox.diagnostics.hostName === "Outlook";
    const newSignature = isClassicOutlook
      ? SignatureManager.extractSignatureForOutlookClassic(newBody)
      : SignatureManager.extractSignature(newBody);

    if (!newSignature) {
      displayError("Email is missing the M3 required signature. Please select an appropriate email signature.", event);
      return;
    }

    const cleanNewSignature = SignatureManager.normalizeSignature(newSignature);
    const signatureKeys = ["monaSignature", "morganSignature", "morvenSignature", "m2Signature", "m3Signature"];
    let matchedSignatureKey = null;
    let rawMatchedSignature = null;

    for (const key of signatureKeys) {
      const cachedSignature = localStorage.getItem(`signature_${key}`);
      if (cachedSignature) {
        const cleanCachedSignature = SignatureManager.normalizeSignature(cachedSignature);
        if (cleanNewSignature === cleanCachedSignature) {
          matchedSignatureKey = key;
          rawMatchedSignature = cachedSignature;
          break;
        }
      }
    }

    const lastAppliedSignature =
      localStorage.getItem("tempSignature_new") ||
      localStorage.getItem("tempSignature_replyForward") ||
      localStorage.getItem(`signature_${signatureKeys[0]}`);
    const cleanLastAppliedSignature = SignatureManager.normalizeSignature(lastAppliedSignature);

    const logoRegex = /<img[^>]+src=["'](.*?(?:m3signatures\/logo\/[^"']+))["'][^>]*>/i;
    const newLogoMatch = newSignature.match(logoRegex);
    let newLogoUrl = newLogoMatch ? newLogoMatch[1].split("?")[0] : null;
    let expectedLogoUrl = rawMatchedSignature
      ? rawMatchedSignature.match(logoRegex)?.[1].split("?")[0]
      : lastAppliedSignature.match(logoRegex)?.[1].split("?")[0];

    const isTextValid = matchedSignatureKey || cleanNewSignature === cleanLastAppliedSignature;
    const isLogoValid = !expectedLogoUrl || (newLogoUrl && newLogoUrl === expectedLogoUrl);

    if (isTextValid && isLogoValid) {
      if (!isReplyOrForward) localStorage.removeItem("tempSignature_new");
      await saveSignatureData(item, matchedSignatureKey || signatureKeys[0]);
      event.completed({ allowEvent: true });
    } else {
      const tempSignature =
        localStorage.getItem("tempSignature_new") || localStorage.getItem("tempSignature_replyForward");
      const signatureKeyToRestore = matchedSignatureKey || (tempSignature ? null : signatureKeys[0]);
      displayError(
        "Selected M3 email signature has been modified. M3 email signature is prohibited from modification. The original signature is now restored",
        event,
        true,
        signatureKeyToRestore,
        tempSignature || localStorage.getItem(`signature_${signatureKeyToRestore || signatureKeys[0]}`)
      );
    }
  } catch (error) {
    logger.log("error", "validateSignatureChanges", { error: error.message });
    displayError("Unexpected error validating signature changes.", event);
  }
}

/**
 * Handles new message compose event.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
async function onNewMessageComposeHandler(event) {
  const item = Office.context.mailbox.item;
  const isReplyOrForward = await SignatureManager.isReplyOrForward(item);
  const isMobile = Office.context.mailbox.diagnostics.hostName === "OutlookMobile";

  if (isReplyOrForward) {
    const conversationId = item.conversationId;
    if (!conversationId) {
      logger.log("info", "onNewMessageComposeHandler", { status: "No conversationId available" });
      if (isMobile) {
        displayNotification("Info", "Debug: No conversationId available.", false);
      }
      displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
      saveSignatureData(item, "none").then(() => event.completed());
      return;
    }

    try {
      // Log conversationId in notification for mobile debugging
      if (isMobile) {
        displayNotification("Info", `Debug: Raw conversationId = ${conversationId}`, false);
      }

      // Relaxed validation to allow more characters, log the raw value
      const validConversationIdPattern = /^[A-Za-z0-9+/=._-]+$/; // Allow dots and hyphens, common in some IDs
      if (!validConversationIdPattern.test(conversationId)) {
        if (isMobile) {
          displayNotification("Info", `Debug: Invalid conversationId format: ${conversationId}`, false);
        }
        logger.log("warn", "onNewMessageComposeHandler", { status: "Invalid conversationId", conversationId });
        // Fallback to getSignatureKeyForRecipients
        const signatureKey = await getSignatureKeyForRecipients(item);
        if (signatureKey) {
          await addSignature(signatureKey, event, true);
        } else {
          displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
          saveSignatureData(item, "none").then(() => event.completed());
        }
        return;
      }

      const accessToken = await getGraphAccessToken();
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        },
      });

      const response = await client
        .api(`/me/mailFolders/SentItems/messages`)
        .filter(`conversationId eq '${encodeURIComponent(conversationId)}'`)
        .select("body")
        .top(10)
        .get();

      if (response.value && response.value.length > 0) {
        const messages = response.value.sort(
          (a, b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime)
        );

        let extractedSignature = null;
        for (const message of messages) {
          const emailBody = message.body?.content || "";
          extractedSignature = SignatureManager.extractSignature(emailBody);
          if (extractedSignature) {
            logger.log("info", "onNewMessageComposeHandler", {
              status: "Signature extracted from Sent Items",
              signatureLength: extractedSignature.length,
            });
            break;
          }
        }

        if (extractedSignature) {
          localStorage.setItem("tempSignature_replyForward", extractedSignature);

          await new Promise((resolve) =>
            item.body.setSignatureAsync(
              "<!-- signature -->" + extractedSignature.trim(),
              { coercionType: Office.CoercionType.Html },
              (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  logger.log("error", "onNewMessageComposeHandler", { error: asyncResult.error.message });
                  if (isMobile) {
                    displayNotification(
                      "Info",
                      `Debug: Failed to apply signature - ${asyncResult.error.message}`,
                      false
                    );
                  }
                  displayNotification("Error", "Failed to apply your signature from conversation.", true);
                  saveSignatureData(item, "none").then(() => event.completed());
                } else {
                  saveSignatureData(item, "tempSignature_replyForward").then(() => event.completed());
                }
                resolve();
              }
            )
          );
        } else {
          logger.log("info", "onNewMessageComposeHandler", { status: "No signature found in Sent Items" });
          if (isMobile) {
            displayNotification("Info", "Debug: No signature found in Sent Items.", false);
          }
          displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
          saveSignatureData(item, "none").then(() => event.completed());
        }
      } else {
        logger.log("info", "onNewMessageComposeHandler", {
          status: "No messages found in Sent Items for this conversation",
        });
        if (isMobile) {
          displayNotification("Info", "Debug: No messages found in Sent Items.", false);
        }
        displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
        saveSignatureData(item, "none").then(() => event.completed());
      }
    } catch (error) {
      logger.log("error", "onNewMessageComposeHandler", { error: error.message });
      if (isMobile) {
        displayNotification("Info", `Debug: Graph Error - ${error.message}`, false);
      }
      displayNotification("Error", `Failed to fetch signature from Graph: ${error.message}`, true);
      saveSignatureData(item, "none").then(() => event.completed());
    }
  } else {
    logger.log("info", "onNewMessageComposeHandler", { status: "New email, no conversationId" });
    if (isMobile) {
      displayNotification("Info", "Debug: New email, no conversationId.", false);
    }
    displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
    saveSignatureData(item, "none").then(() => {
      localStorage.removeItem("tempSignature_new");
      event.completed();
    });
  }
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

export {
  addSignature,
  addSignatureMona,
  addSignatureMorgan,
  addSignatureMorven,
  addSignatureM2,
  addSignatureM3,
  applyDefaultSignature,
  cancelAction,
  displayError,
  displayNotification,
  validateSignature,
  validateSignatureChanges,
  onNewMessageComposeHandler,
};
