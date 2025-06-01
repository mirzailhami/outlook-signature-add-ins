/* global Office */

import { getGraphAccessToken } from "./launchevent.js";
import "isomorphic-fetch";
import { Client } from "@microsoft/microsoft-graph-client";
import { logger, SignatureManager, fetchSignature, detectSignatureKey, appendDebugLogToBody } from "./helpers.js";

// Mobile needs this initialization
Office.initialize = () => {};

Office.onReady(() => {
  logger.log("info", "Office.onReady", { host: Office.context?.mailbox?.diagnostics?.hostName });

  Office.actions.associate("addSignatureMona", addSignatureMona);
  Office.actions.associate("addSignatureMorgan", addSignatureMorgan);
  Office.actions.associate("addSignatureMorven", addSignatureMorven);
  Office.actions.associate("addSignatureM2", addSignatureM2);
  Office.actions.associate("addSignatureM3", addSignatureM3);
  Office.actions.associate("validateSignature", validateSignature);
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
});

/**
 * Completes the event with a signature state and optional notification.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 * @param {string} signatureKey - The signature key applied.
 * @param {string} [notificationType] - Notification type ("Info" or "Error").
 * @param {string} [notificationMessage] - Notification message.
 */
async function completeWithState(event, signatureKey, notificationType, notificationMessage) {
  if (notificationMessage) {
    displayNotification(notificationType, notificationMessage, notificationType === "Error");
  }
  event.completed();
}

/**
 * Adds a signature to the email and saves it to localStorage.
 * @param {string} signatureKey - The signature key (e.g., "m3Signature").
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 * @param {Function} completeWithStateFn - The function to complete the event with state.
 * @param {boolean} isAutoApplied - Whether the signature is auto-applied.
 */
async function addSignature(signatureKey, event, completeWithStateFn, isAutoApplied = false) {
  const item = Office.context.mailbox.item;
  const isMobile =
    Office.context.mailbox.diagnostics.hostName === "OutlookAndroid" ||
    Office.context.mailbox.diagnostics.hostName === "OutlookIOS";

  try {
    localStorage.removeItem("tempSignature");
    localStorage.setItem("tempSignature", signatureKey);
    const cachedSignature = localStorage.getItem(`signature_${signatureKey}`);

    if (cachedSignature && !isAutoApplied) {
      const signatureWithMarker = "<!-- signature -->" + cachedSignature.trim();
      await new Promise((resolve, reject) => {
        item.body.setSignatureAsync(signatureWithMarker, { coercionType: Office.CoercionType.Html }, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            if (isMobile) {
              appendDebugLogToBody(item, "addSignature Error (Cached)", "Message", asyncResult.error.message);
            }
            logger.log("error", "addSignature", { error: asyncResult.error.message });
            displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
            if (!isAutoApplied) {
              event.completed();
              resolve();
            } else {
              // Move the completeWithStateFn call outside the callback
              reject(new Error("setSignatureAsync failed for cached signature"));
            }
          } else {
            item.body.getAsync("html", (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                logger.log("debug", "addSignature", {
                  bodyContainsMarker: result.value.includes("<!-- signature -->"),
                  bodyLength: result.value.length,
                });
              }
              event.completed();
              resolve();
            });
          }
        });
      });
      // If rejected due to error and isAutoApplied, handle it here
      throw new Error("setSignatureAsync failed for cached signature");
    } else {
      await new Promise((resolve, reject) => {
        fetchSignature(signatureKey, (template, error) => {
          if (error) {
            if (isMobile) {
              appendDebugLogToBody(item, "addSignature Error (Fetch)", "Message", error.message);
            }
            logger.log("error", "addSignature", { error: error.message });
            displayNotification("Error", `Failed to fetch ${signatureKey}.`, true);
            if (!isAutoApplied) {
              event.completed();
              resolve();
            } else {
              reject(new Error("fetchSignature failed"));
            }
            return;
          }

          const signatureWithMarker = "<!-- signature -->" + template.trim();
          item.body.setSignatureAsync(
            signatureWithMarker,
            { coercionType: Office.CoercionType.Html },
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                if (isMobile) {
                  appendDebugLogToBody(item, "addSignature Error (Set)", "Message", asyncResult.error.message);
                }
                logger.log("error", "addSignature", { error: asyncResult.error.message });
                displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
                if (!isAutoApplied) {
                  event.completed();
                  resolve();
                } else {
                  reject(new Error("setSignatureAsync failed for fetched signature"));
                }
              } else {
                item.body.getAsync("html", (result) => {
                  if (result.status === Office.AsyncResultStatus.Succeeded) {
                    logger.log("debug", "addSignature", {
                      bodyContainsMarker: result.value.includes("<!-- signature -->"),
                      bodyLength: result.value.length,
                    });
                  }
                  localStorage.setItem(`signature_${signatureKey}`, template);
                  event.completed();
                  resolve();
                });
              }
            }
          );
        });
      });
    }
  } catch (error) {
    if (isMobile) {
      await appendDebugLogToBody(item, "addSignature Error (Catch)", "Message", error.message);
    }
    logger.log("error", "addSignature", { error: error.message });
    displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
    if (isAutoApplied) {
      await completeWithStateFn(event, "none", "Info", "Please select an M3 signature from the ribbon.");
    } else {
      event.completed();
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
async function displayError(message, event) {
  logger.log("info", "displayError", { message });

  const markdownMessage = message.includes("modified")
    ? `${message}\n\n**Tip**: Ensure the M3 signature is not edited before sending.`
    : `${message}\n\n**Tip**: Select an M3 signature from the ribbon under "M3 Signatures".`;

  displayNotification("Error", message, true);
  event.completed({
    allowEvent: false,
    errorMessage: message,
    errorMessageMarkdown: markdownMessage,
    cancelLabel: "OK",
  });
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
    const originalSignatureKey = localStorage.getItem("tempSignature");
    const rawMatchedSignature = localStorage.getItem(`signature_${originalSignatureKey}`);

    const cleanCurrentSignature = SignatureManager.normalizeSignature(currentSignature);
    const cleanCachedSignature = SignatureManager.normalizeSignature(rawMatchedSignature);

    const logoRegex = /<img[^>]+src=["'](.*?(?:m3signatures\/logo\/[^"']+))["'][^>]*>/i;
    const currentLogoMatch = currentSignature.match(logoRegex);
    let currentLogoUrl = currentLogoMatch ? currentLogoMatch[1].split("?")[0] : null;

    const expectedLogoMatch = rawMatchedSignature.match(logoRegex);
    let expectedLogoUrl = expectedLogoMatch ? expectedLogoMatch[1].split("?")[0] : null;

    const isTextValid = cleanCurrentSignature === cleanCachedSignature;
    const isLogoValid = !expectedLogoUrl || currentLogoUrl === expectedLogoUrl;

    logger.log("debug", "validateSignatureChanges", {
      rawCurrentSignatureLength: currentSignature.length,
      rawMatchedSignatureLength: rawMatchedSignature.length,
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
      localStorage.removeItem("tempSignature");
      event.completed({ allowEvent: true });
    } else {
      const signatureToRestore = localStorage.getItem(`signature_${originalSignatureKey}`);
      const restored = await SignatureManager.restoreSignature(item, signatureToRestore, originalSignatureKey);
      if (!restored) {
        await displayError("Failed to restore the original M3 signature. Please reselect.", event);
        return;
      }

      await new Promise((resolve) => setTimeout(resolve, 500));
      await displayError(
        "Selected M3 email signature has been modified. M3 email signature is prohibited from modification. The original signature has been restored.",
        event
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

  if (isReplyOrForward) {
    logger.log("info", "onNewMessageComposeHandler", { status: "Processing reply/forward email" });

    try {
      const accessToken = await getGraphAccessToken();
      const client = Client.init({
        authProvider: (done) => done(null, accessToken),
      });

      const subjectResult = await new Promise((resolve) => item.subject.getAsync((result) => resolve(result)));
      let emailSubject = "Unknown";
      if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
        emailSubject = subjectResult.value.trim();
        logger.log("info", "onNewMessageComposeHandler", { debug: emailSubject });
      } else {
        logger.log("error", "onNewMessageComposeHandler", {
          error: "Failed to get subject",
          details: subjectResult.error.message,
        });
      }

      const getItemIdAsync = await new Promise((resolve) =>
        Office.context.mailbox.item.getItemIdAsync((asyncResult) => resolve(asyncResult))
      );

      // console.log(Office.context.mailbox.item.getItemIdAsync);
      console.log(getItemIdAsync.value);

      await appendDebugLogToBody(item, "getItemIdAsync", "itemId", getItemIdAsync.value);

      const toResult = await new Promise((resolve) =>
        Office.context.mailbox.item.to.getAsync((asyncResult) => resolve(asyncResult))
      );
      let recipientEmail = "Unknown";
      console.log(toResult);
      if (toResult.status === Office.AsyncResultStatus.Succeeded && toResult.value.length > 0) {
        recipientEmail = toResult.value[0].emailAddress.toLowerCase();
        logger.log("info", "onNewMessageComposeHandler", { debug: recipientEmail });
      } else {
        logger.log("error", "onNewMessageComposeHandler", {
          error: "Failed to get 'to'",
          details: toResult.error.message,
        });
      }

      const response = await client
        .api("/me/messages")
        .filter(`sentDateTime ge 2023-01-11T07:28:08Z and subject eq '${emailSubject}'`)
        .select("subject,body,sentDateTime,toRecipients")
        .orderby("sentDateTime desc")
        .top(10)
        .get();

      if (response.value && response.value.length > 0) {
        const matchingEmails = response.value.filter((email) =>
          email.toRecipients.some((recipient) => recipient.emailAddress.address.toLowerCase() === recipientEmail)
        );

        if (matchingEmails.length === 0) {
          logger.log("warn", "onNewMessageComposeHandler", {
            status: "No emails matched the recipient in Sent Items",
          });
          await completeWithState(
            event,
            "none",
            "Info",
            isMobile
              ? "No matching email found in Sent Items for this recipient. Please select an M3 signature from the task pane."
              : "No matching email found in Sent Items for this recipient. Please select an M3 signature from the ribbon."
          );
          return;
        }

        let matchedMessage = matchingEmails[0];

        if (matchedMessage) {
          const emailBody = matchedMessage.body?.content || "";
          const extractedSignature = SignatureManager.extractSignature(emailBody);

          if (extractedSignature) {
            logger.log("info", "onNewMessageComposeHandler", {
              status: "Signature extracted from email",
              signatureLength: extractedSignature.length,
            });

            const matchedSignatureKey = detectSignatureKey(extractedSignature);
            if (matchedSignatureKey) {
              logger.log("info", "onNewMessageComposeHandler", {
                status: "Detected signature key from content",
                matchedSignatureKey,
              });
              localStorage.removeItem("tempSignature");
              localStorage.setItem("tempSignature", matchedSignatureKey);
              await addSignature(matchedSignatureKey, event, completeWithState, true);
              await completeWithState(event, matchedSignatureKey, null, null);
            } else {
              logger.log("warn", "onNewMessageComposeHandler", { status: "Could not detect signature key" });
              await completeWithState(
                event,
                "none",
                "Info",
                isMobile
                  ? "Could not detect signature type. Please select an M3 signature from the task pane."
                  : "Could not detect signature type. Please select an M3 signature from the ribbon."
              );
            }
          } else {
            await completeWithState(
              event,
              "none",
              "Info",
              isMobile
                ? "No signature found in email. Please select an M3 signature from the task pane."
                : "No signature found in email. Please select an M3 signature from the ribbon."
            );
          }
        } else {
          logger.log("warn", "onNewMessageComposeHandler", { status: "No matching email found for recipient" });
          await completeWithState(
            event,
            "none",
            "Info",
            isMobile
              ? "No matching email found with the recipient. Please select an M3 signature from the task pane."
              : "No matching email found with the recipient. Please select an M3 signature from the ribbon."
          );
        }
      } else {
        await completeWithState(
          event,
          "none",
          "Info",
          isMobile
            ? "Please select an M3 signature from the task pane."
            : "Please select an M3 signature from the ribbon."
        );
      }
    } catch (error) {
      logger.log("error", "onNewMessageComposeHandler", { error: error.message, stack: error.stack });
      await completeWithState(event, "none", "Error", `Failed to fetch signature from Graph: ${error.message}`);
    }
  } else {
    if (isMobile) {
      const mobileDefaultSignatureKey = localStorage.getItem("mobileDefaultSignature");
      if (mobileDefaultSignatureKey) {
        try {
          localStorage.removeItem("tempSignature");
          localStorage.setItem("tempSignature", mobileDefaultSignatureKey);

          await addSignature(mobileDefaultSignatureKey, event, completeWithState, true);
          await completeWithState(event, mobileDefaultSignatureKey, null, null);
        } catch (error) {
          await appendDebugLogToBody(item, "Error Applying Default Signature", "Message", error.message);
          await completeWithState(event, "none", "Error", `Failed to fetch default signature: ${error.message}`);
        }
      } else {
        await completeWithState(event, "none", "Info", "Please select an M3 signature from the task pane.");
      }
    } else {
      await completeWithState(event, "none", "Info", "Please select an M3 signature from the ribbon.");
    }
  }
}

/**
 * Adds the Mona signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureMona(event) {
  addSignature("monaSignature", event, completeWithState);
}

/**
 * Adds the Morgan signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureMorgan(event) {
  addSignature("morganSignature", event, completeWithState);
}

/**
 * Adds the Morven signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureMorven(event) {
  addSignature("morvenSignature", event, completeWithState);
}

/**
 * Adds the M2 signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureM2(event) {
  addSignature("m2Signature", event, completeWithState);
}

/**
 * Adds the M3 signature.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function addSignatureM3(event) {
  addSignature("m3Signature", event, completeWithState);
}

export {
  addSignature,
  addSignatureMona,
  addSignatureMorgan,
  addSignatureMorven,
  addSignatureM2,
  addSignatureM3,
  displayError,
  displayNotification,
  validateSignature,
  validateSignatureChanges,
  onNewMessageComposeHandler,
  completeWithState,
};
