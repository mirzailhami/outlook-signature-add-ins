/* global Office */

import { getGraphAccessToken } from "./launchevent.js";
import "isomorphic-fetch";
import { Client } from "@microsoft/microsoft-graph-client";
import {
  logger,
  SignatureManager,
  fetchSignature,
  getSignatureKeyForRecipients,
  saveSignatureData,
} from "./helpers.js";

// Mobile needs this initialization
Office.initialize = () => {};

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

    // Check if the current signature matches any cached signature
    for (const key of signatureKeys) {
      const cachedSignature = localStorage.getItem(`signature_${key}`);
      if (cachedSignature) {
        const cleanCachedSignature = SignatureManager.normalizeSignature(cachedSignature);
        logger.log("debug", "validateSignatureChanges", {
          key,
          cleanNewSignatureLength: cleanNewSignature.length,
          cleanCachedSignatureLength: cleanCachedSignature.length,
          isMatch: cleanNewSignature === cleanCachedSignature,
        });
        if (cleanNewSignature === cleanCachedSignature) {
          matchedSignatureKey = key;
          rawMatchedSignature = cachedSignature;
          break;
        }
      }
    }

    // Determine the last applied signature based on context
    const lastAppliedSignature = isReplyOrForward
      ? localStorage.getItem("tempSignature_replyForward")
      : localStorage.getItem("tempSignature_new");
    const cleanLastAppliedSignature = SignatureManager.normalizeSignature(lastAppliedSignature);

    // Log the comparison with the last applied signature
    logger.log("debug", "validateSignatureChanges", {
      cleanNewSignature,
      cleanLastAppliedSignature,
      isLastAppliedMatch: cleanNewSignature === cleanLastAppliedSignature,
      isReplyOrForward,
    });

    // Logo validation
    const logoRegex = /<img[^>]+src=["'](.*?(?:m3signatures\/logo\/[^"']+))["'][^>]*>/i;
    const newLogoMatch = newSignature.match(logoRegex);
    let newLogoUrl = newLogoMatch ? newLogoMatch[1].split("?")[0] : null;
    let expectedLogoUrl = null;

    if (rawMatchedSignature) {
      const match = rawMatchedSignature.match(logoRegex);
      expectedLogoUrl = match ? match[1].split("?")[0] : null;
    } else if (lastAppliedSignature) {
      const match = lastAppliedSignature.match(logoRegex);
      expectedLogoUrl = match ? match[1].split("?")[0] : null;
    }

    logger.log("debug", "validateSignatureChanges", {
      newLogoUrl,
      expectedLogoUrl,
      logoMatch: !expectedLogoUrl || newLogoUrl === expectedLogoUrl,
    });

    const isTextValid =
      matchedSignatureKey || (lastAppliedSignature && cleanNewSignature === cleanLastAppliedSignature);
    const isLogoValid = !expectedLogoUrl || !newLogoUrl || newLogoUrl === expectedLogoUrl; // Allow missing logos in new signature

    if (isTextValid && isLogoValid) {
      if (!isReplyOrForward) localStorage.removeItem("tempSignature_new");
      await saveSignatureData(item, matchedSignatureKey || (lastAppliedSignature ? "tempSignature" : signatureKeys[0]));
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
    logger.log("error", "validateSignatureChanges", { error: error.message, stack: error.stack });
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
  const isMobile =
    Office.context.mailbox.diagnostics.hostName === "OutlookAndroid" ||
    Office.context.mailbox.diagnostics.hostName === "OutlookIOS";

  logger.log("info", "onNewMessageComposeHandler", {
    isReplyOrForward,
    isMobile,
    hostName: Office.context.mailbox.diagnostics.hostName,
  });

  // Helper function to save state and complete the event
  const completeWithState = async (signatureKey, notificationType, notificationMessage) => {
    if (notificationMessage) {
      displayNotification(notificationType, notificationMessage, notificationType === "Error");
    }
    await saveSignatureData(item, signatureKey);
    if (signatureKey === "none") {
      localStorage.removeItem("tempSignature_new");
    }
    event.completed();
  };

  // Helper function to wrap fetchSignature in a Promise
  const getSignature = (signatureKey) => {
    return new Promise((resolve, reject) => {
      fetchSignature(signatureKey, (template, error) => {
        if (error) {
          reject(error);
        } else {
          resolve(template);
        }
      });
    });
  };

  if (isReplyOrForward) {
    logger.log("info", "onNewMessageComposeHandler", { status: "Processing reply/forward email" });
    const conversationId = item.conversationId;
    if (!conversationId && !isMobile) {
      logger.log("info", "onNewMessageComposeHandler", { status: "No conversationId available" });
      await completeWithState("none", "Info", "Please select an M3 signature from the ribbon.");
      return;
    }

    try {
      const accessToken = await getGraphAccessToken();
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        },
      });

      // Determine the filter based on platform
      let filterString;
      let response;
      if (!isMobile) {
        // OWA uses conversationId
        filterString = `conversationId eq '${encodeURIComponent(conversationId)}'`;
        response = await client
          .api(`/me/mailFolders/SentItems/messages`)
          .filter(filterString)
          .select("subject,body")
          .top(1)
          .get();
      } else {
        // Mobile uses 'to' email to search
        const toResult = await new Promise((resolve) =>
          Office.context.mailbox.item.to.getAsync((asyncResult) => resolve(asyncResult))
        );
        let recipientEmail = "Unknown";
        if (toResult.status === Office.AsyncResultStatus.Succeeded && toResult.value.length > 0) {
          recipientEmail = toResult.value[0].emailAddress;
          logger.log("info", "onNewMessageComposeHandler", { debug: `Using recipient email: ${recipientEmail}` });
        } else {
          logger.log("error", "onNewMessageComposeHandler", {
            error: "Failed to get 'to'",
            details: toResult.error.message,
          });
          recipientEmail = `Error: ${toResult.error.message}`;
        }

        // Graph API search using raw email address with quotes
        const searchQuery = `"${recipientEmail}"`; // Use quotes to treat as a phrase
        logger.log("debug", "onNewMessageComposeHandler", { searchQuery });
        response = await client
          .api(`/me/mailFolders/SentItems/messages`)
          .search(searchQuery)
          .select("subject,body")
          .top(1)
          .get();
      }

      logger.log("debug", "onNewMessageComposeHandler", { response });

      if (response.value && response.value.length > 0) {
        const message = response.value[0]; // Single message with top(1)
        const emailBody = message.body?.content || "";
        const extractedSignature = SignatureManager.extractSignature(emailBody);
        if (extractedSignature) {
          logger.log("info", "onNewMessageComposeHandler", {
            status: "Signature extracted from Sent Items",
            signatureLength: extractedSignature.length,
          });
          localStorage.setItem("tempSignature_replyForward", extractedSignature);

          await new Promise((resolve) =>
            item.body.setSignatureAsync(
              "<!-- signature -->" + extractedSignature.trim(),
              { coercionType: Office.CoercionType.Html },
              (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  logger.log("error", "onNewMessageComposeHandler", { error: asyncResult.error.message });
                  displayNotification("Error", "Failed to apply your signature from conversation.", true);
                }
                resolve();
              }
            )
          );
          await completeWithState("tempSignature_replyForward", null, null);
        } else {
          logger.log("info", "onNewMessageComposeHandler", { status: "No signature found in Sent Items" });
          await completeWithState(
            "none",
            "Info",
            "No signature found in Sent Items. Please select an M3 signature from the ribbon."
          );
        }
      } else {
        await completeWithState("none", "Info", "Please select an M3 signature from the ribbon.");
      }
    } catch (error) {
      logger.log("error", "onNewMessageComposeHandler", { error: error.message, stack: error.stack });
      await completeWithState("none", "Error", `Failed to fetch signature from Graph: ${error.message}`);
    }
  } else {
    logger.log("info", "onNewMessageComposeHandler", { status: "Processing new email" });
    if (isMobile) {
      const defaultSignatureKey = localStorage.getItem("defaultSignature");
      if (defaultSignatureKey) {
        logger.log("info", "onNewMessageComposeHandler", { status: "Applying default signature", defaultSignatureKey });
        try {
          const signature = await getSignature(defaultSignatureKey);
          if (signature) {
            await new Promise((resolve) =>
              item.body.setSignatureAsync(
                "<!-- signature -->" + signature.trim(),
                { coercionType: Office.CoercionType.Html },
                (asyncResult) => {
                  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    logger.log("error", "onNewMessageComposeHandler", { error: asyncResult.error.message });
                    displayNotification("Error", "Failed to apply default signature.", true);
                  }
                  resolve();
                }
              )
            );
            await completeWithState(defaultSignatureKey, null, null);
          } else {
            logger.log("warn", "onNewMessageComposeHandler", {
              status: "Default signature not found",
              defaultSignatureKey,
            });
            await completeWithState(
              "none",
              "Info",
              "Default signature not found. Please select an M3 signature from the task pane."
            );
          }
        } catch (error) {
          logger.log("error", "onNewMessageComposeHandler", { error: error.message, stack: error.stack });
          await completeWithState("none", "Error", `Failed to fetch default signature: ${error.message}`);
        }
      } else {
        logger.log("info", "onNewMessageComposeHandler", { status: "No default signature set" });
        await completeWithState("none", "Info", "Please select an M3 signature from the task pane.");
      }
    } else {
      await completeWithState("none", "Info", "Please select an M3 signature from the ribbon.");
    }
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
