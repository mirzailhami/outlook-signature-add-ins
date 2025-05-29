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
      const signatureWithMarker = "<!-- signature -->" + cachedSignature.trim();
      await new Promise((resolve) =>
        item.body.setSignatureAsync(signatureWithMarker, { coercionType: Office.CoercionType.Html }, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            logger.log("error", "addSignature", { error: asyncResult.error.message });
            displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
            if (!isAutoApplied) event.completed();
            else {
              displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
              saveSignatureData(item, "none").then(() => event.completed());
            }
          } else {
            // Verify the signature was applied
            item.body.getAsync("html", (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                logger.log("debug", "addSignature", {
                  bodyContainsMarker: result.value.includes("<!-- signature -->"),
                  bodyLength: result.value.length,
                });
              }
              saveSignatureData(item, signatureKey).then(() => {
                if (!isAutoApplied) {
                  localStorage.setItem("tempSignature_new", cachedSignature);
                }
                event.completed();
              });
            });
          }
          resolve();
        })
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

        const signatureWithMarker = "<!-- signature -->" + template.trim();
        await new Promise((resolve) =>
          item.body.setSignatureAsync(
            signatureWithMarker,
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
                // Verify the signature was applied
                item.body.getAsync("html", (result) => {
                  if (result.status === Office.AsyncResultStatus.Succeeded) {
                    logger.log("debug", "addSignature", {
                      bodyContainsMarker: result.value.includes("<!-- signature -->"),
                      bodyLength: result.value.length,
                    });
                  }
                  localStorage.setItem(`signature_${signatureKey}`, template);
                  saveSignatureData(item, signatureKey).then(() => {
                    if (!isAutoApplied) {
                      localStorage.setItem("tempSignature_new", template);
                    }
                    event.completed();
                  });
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
      displayNotification("Error", `${message} (No signature to restore)`, true);
      event.completed({
        allowEvent: false,
        errorMessage: `${message} (No signature to restore)`,
        errorMessageMarkdown: `${message} (No signature to restore)\n\n**Note**: Failed to restore signature. Please reselect.`,
        cancelLabel: "OK",
      });
      return;
    }

    const restored = await SignatureManager.restoreSignature(item, signatureToRestore, signatureKey || "tempSignature");
    if (!restored) {
      logger.log("error", "displayError", { error: "Restoration failed", signatureKey });
      displayNotification("Error", `${message} (Failed to restore)`, true);
      event.completed({
        allowEvent: false,
        errorMessage: `${message} (Failed to restore)`,
        errorMessageMarkdown: `${message} (Failed to restore)\n\n**Note**: Failed to restore signature. Please reselect.`,
        cancelLabel: "OK",
      });
      return;
    }

    await new Promise((resolve) => setTimeout(resolve, 500)); // 500ms delay

    const customMessage =
      "Selected M3 email signature has been modified. M3 email signature is prohibited from modification. The original signature is now restored."; // 88 characters
    displayNotification("Error", customMessage, true);
    event.completed({
      allowEvent: false,
      errorMessage: customMessage,
      errorMessageMarkdown: `${customMessage}\n\n**Tip**: Avoid modifying the M3 signature.`,
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
    let newBody = await new Promise((resolve) =>
      item.body.getAsync("html", (result) =>
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : null)
      )
    );
    const isClassicOutlook = Office.context.mailbox.diagnostics.hostName === "Outlook";
    let newSignature = isClassicOutlook
      ? SignatureManager.extractSignatureForOutlookClassic(newBody)
      : SignatureManager.extractSignature(newBody);

    if (!newSignature) {
      await displayError(
        "Email is missing the M3 required signature. Please select an appropriate email signature.",
        event
      );
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
        console.log({
          event: "validateSignatureChanges",
          signatureKey: key,
          rawCachedSignature: cachedSignature,
          cleanCachedSignature,
        });
        if (cleanNewSignature === cleanCachedSignature) {
          matchedSignatureKey = key;
          rawMatchedSignature = cachedSignature;
          console.log({ event: "validateSignatureChanges", status: "Matched signature", matchedSignatureKey });
          break;
        }
      }
    }

    const lastAppliedSignature =
      localStorage.getItem("tempSignature_new") ||
      localStorage.getItem("tempSignature_replyForward") ||
      localStorage.getItem(`signature_${signatureKeys[0]}`);
    const cleanLastAppliedSignature = SignatureManager.normalizeSignature(lastAppliedSignature);

    // Extract logo URL from new signature
    const logoRegex = /<img[^>]+src=["'](.*?(?:m3signatures\/logo\/[^"']+))["'][^>]*>/i;
    const newLogoMatch = newSignature.match(logoRegex);
    let newLogoUrl = newLogoMatch ? newLogoMatch[1] : null;
    // Remove query parameters for comparison
    if (newLogoUrl) {
      newLogoUrl = newLogoUrl.split("?")[0];
    }
    console.log({ event: "validateSignatureChanges", newLogoUrl });

    // Extract expected logo URL from the matched cached signature
    let expectedLogoUrl = null;
    if (rawMatchedSignature) {
      const expectedLogoMatch = rawMatchedSignature.match(logoRegex);
      expectedLogoUrl = expectedLogoMatch ? expectedLogoMatch[1] : null;
      // Remove query parameters for comparison
      if (expectedLogoUrl) {
        expectedLogoUrl = expectedLogoUrl.split("?")[0];
      }
    } else if (lastAppliedSignature) {
      // Fallback to last applied signature if no match is found
      const lastAppliedLogoMatch = lastAppliedSignature.match(logoRegex);
      expectedLogoUrl = lastAppliedLogoMatch ? lastAppliedLogoMatch[1] : null;
      if (expectedLogoUrl) {
        expectedLogoUrl = expectedLogoUrl.split("?")[0];
      }
    }

    const isTextValid = matchedSignatureKey || cleanNewSignature === cleanLastAppliedSignature;
    const isLogoValid = !expectedLogoUrl || (newLogoUrl && newLogoUrl === expectedLogoUrl);

    logger.log("debug", "validateSignatureChanges", {
      fullBodyLength: newBody?.length,
      rawNewSignature: newSignature,
      rawLastAppliedSignature: lastAppliedSignature,
      cleanNewSignature,
      cleanLastAppliedSignature,
      matchedSignatureKey,
      isReplyOrForward,
      isTextValid,
      isLogoValid,
    });

    if (isTextValid && isLogoValid) {
      if (!isReplyOrForward) localStorage.removeItem("tempSignature_new");
      await saveSignatureData(item, matchedSignatureKey || signatureKeys[0]);
      event.completed({ allowEvent: true });
    } else {
      const tempSignature =
        localStorage.getItem("tempSignature_new") || localStorage.getItem("tempSignature_replyForward");
      const signatureKeyToRestore = matchedSignatureKey || (tempSignature ? null : signatureKeys[0]);
      await displayError(
        "Selected M3 email signature has been modified. M3 email signature is prohibited from modification. The original signature is now restored.",
        event,
        true,
        signatureKeyToRestore,
        tempSignature || localStorage.getItem(`signature_${signatureKeyToRestore || signatureKeys[0]}`)
      );
    }
  } catch (error) {
    logger.log("error", "validateSignatureChanges", { error: error.message });
    await displayError("Unexpected error validating signature changes.", event);
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

    try {
      const accessToken = await getGraphAccessToken();
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        },
      });

      // Get subject and recipient for the query
      const subjectResult = await new Promise((resolve) => item.subject.getAsync((result) => resolve(result)));
      let emailSubject = "Unknown";
      if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
        emailSubject = SignatureManager.normalizeSubject(subjectResult.value);
        logger.log("info", "onNewMessageComposeHandler", { debug: `Using subject: ${emailSubject}` });
      } else {
        logger.log("error", "onNewMessageComposeHandler", {
          error: "Failed to get subject",
          details: subjectResult.error.message,
        });
        emailSubject = `Error: ${subjectResult.error.message}`;
      }

      const toResult = await new Promise((resolve) =>
        Office.context.mailbox.item.to.getAsync((asyncResult) => resolve(asyncResult))
      );
      let recipientEmail = "Unknown";
      if (toResult.status === Office.AsyncResultStatus.Succeeded && toResult.value.length > 0) {
        recipientEmail = toResult.value[0].emailAddress.toLowerCase();
        logger.log("info", "onNewMessageComposeHandler", { debug: `Using recipient email: ${recipientEmail}` });
      } else {
        logger.log("error", "onNewMessageComposeHandler", {
          error: "Failed to get 'to'",
          details: toResult.error.message,
        });
        recipientEmail = `Error: ${toResult.error.message}`;
      }

      // Construct Graph API query with $search for subject
      const searchQuery = `"${emailSubject}"`; // Search for the normalized subject
      logger.log("debug", "onNewMessageComposeHandler", { searchQuery });

      const response = await client
        .api(`/me/mailFolders/SentItems/messages`)
        .search(searchQuery)
        .select("subject,body,sentDateTime,toRecipients,ccRecipients,bccRecipients")
        .top(10) // Fetch multiple results for sorting
        .get();

      logger.log("debug", "onNewMessageComposeHandler", { response });

      if (response.value && response.value.length > 0) {
        // Sort by sentDateTime descending to get the latest email
        const sortedMessages = response.value.sort((a, b) => {
          return new Date(b.sentDateTime) - new Date(a.sentDateTime);
        });

        // Find the first email that matches the recipient
        let matchedMessage = null;
        for (const message of sortedMessages) {
          const recipients = [...(message.toRecipients || []).map((r) => r.emailAddress.address.toLowerCase())];
          if (recipients.includes(recipientEmail)) {
            matchedMessage = message;
            break;
          }
        }

        if (matchedMessage) {
          const emailBody = matchedMessage.body?.content || "";
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
              isMobile
                ? "No signature found in Sent Items. Please select an M3 signature from the task pane."
                : "No signature found in Sent Items. Please select an M3 signature from the ribbon."
            );
          }
        } else {
          logger.log("warn", "onNewMessageComposeHandler", { status: "No matching email found for recipient" });
          await completeWithState(
            "none",
            "Info",
            isMobile
              ? "No matching email found with the recipient. Please select an M3 signature from the task pane."
              : "No matching email found with the recipient. Please select an M3 signature from the ribbon."
          );
        }
      } else {
        await completeWithState(
          "none",
          "Info",
          isMobile
            ? "Please select an M3 signature from the task pane."
            : "Please select an M3 signature from the ribbon."
        );
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
