/* global Office */

import { fetchEmailById } from "./graph.js";
import {
  logger,
  completeWithState,
  SignatureManager,
  fetchSignature,
  detectSignatureKey,
  appendDebugLogToBody,
  displayNotification,
  displayError,
} from "./helpers.js";

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
            // displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
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
      displayError(
        "Email is missing the M3 required signature. Please select an appropriate email signature.",
        event,
        false
      );
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
 * Handles the new message compose event, applying the appropriate signature for reply/forward or new messages.
 * @param {Object} event - The event object from Office.js.
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

  displayNotification("Info", `Platform: ${Office.context.mailbox.diagnostics.hostName}`, true);

  try {
    if (isReplyOrForward) {
      logger.log("info", "onNewMessageComposeHandler", { status: "Processing reply/forward email" });

      let messageId;
      if (isMobile) {
        messageId = Office.context.mailbox.item.conversationId;
      } else {
        const itemIdResult = await new Promise((resolve) => item.getItemIdAsync((asyncResult) => resolve(asyncResult)));
        if (itemIdResult.status !== Office.AsyncResultStatus.Succeeded) {
          throw new Error(`Failed to get item ID: ${itemIdResult.error.message}`);
        }
        messageId = itemIdResult.value;
        logger.log("info", "getItemIdAsync for OWA", { messageId });
      }

      const email = await fetchEmailById(messageId);
      const emailBody = email.body?.content || "";
      const extractedSignature = SignatureManager.extractSignature(emailBody);

      if (!extractedSignature) {
        logger.log("warn", "onNewMessageComposeHandler", { status: "No signature found in email" });
        await completeWithState(
          event,
          "none",
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
        await completeWithState(
          event,
          "none",
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

      localStorage.removeItem("tempSignature");
      localStorage.setItem("tempSignature", matchedSignatureKey);
      await addSignature(matchedSignatureKey, event, completeWithState, true);
      await completeWithState(event, matchedSignatureKey, null, null);
    } else {
      // Handle new message
      if (isMobile) {
        const mobileDefaultSignatureKey = localStorage.getItem("mobileDefaultSignature");
        if (mobileDefaultSignatureKey) {
          localStorage.removeItem("tempSignature");
          localStorage.setItem("tempSignature", mobileDefaultSignatureKey);
          await addSignature(mobileDefaultSignatureKey, event, completeWithState, true);
          await completeWithState(event, mobileDefaultSignatureKey, null, null);
        } else {
          await completeWithState(event, "none", "Info", "Please select an M3 signature from the task pane.");
        }
      } else {
        await completeWithState(event, "none", "Info", "Please select an M3 signature from the ribbon.");
      }
    }
  } catch (error) {
    logger.log("error", "onNewMessageComposeHandler", { error: error.message, stack: error.stack });
    await appendDebugLogToBody(item, "Message", error.message, "Stack", error.stack);
    await completeWithState(event, "none", "Error", `Failed to process compose event: ${error.message}`);
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
