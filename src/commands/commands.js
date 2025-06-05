/* global Office, console, logger, SignatureManager, displayNotification, displayError, completeWithState, fetchSignature, detectSignatureKey, appendDebugLogToBody, fetchEmailById */

// import { completeWithState, displayNotification } from "./helpers";

// Use process.env.ASSET_BASE_URL to construct dynamic URLs
const ASSET_BASE_URL = process.env.ASSET_BASE_URL;

function loadScript(url) {
  return new Promise((resolve, reject) => {
    console.log("Attempting to load script:", url); // Debug log
    var script = document.createElement("script");
    script.type = "text/javascript";
    script.src = url;
    script.onload = () => {
      console.log("Script loaded successfully:", url); // Debug log
      resolve();
    };
    script.onerror = () => {
      console.error("Script failed to load:", url); // Debug log
      reject(new Error("Failed to load: " + url));
    };
    document.head.appendChild(script);
  });
}

// Load scripts dynamically using the base URL
let helpersLoaded = false;
let graphLoaded = false;
let isMobile = false;
let isClassicOutlook = false;

// Load scripts sequentially with error handling
Promise.all([
  loadScript(`${ASSET_BASE_URL}/helpers.js`)
    .then(() => {
      helpersLoaded = true;
    })
    .catch(() => {
      helpersLoaded = false;
    }),
  loadScript(`${ASSET_BASE_URL}/graph.js`)
    .then(() => {
      graphLoaded = true;
    })
    .catch(() => {
      graphLoaded = false;
    }),
])
  .then(() => {
    console.log("All dependencies loaded successfully");
    initializeAddIn();
  })
  .catch((error) => {
    console.error("Dependency loading failed:", error);
    initializeAddIn();
  });

function initializeAddIn() {
  Office.onReady(() => {
    if (!helpersLoaded) {
      console.warn("Helpers.js failed to load; logging and signature management features disabled.");
    }
    if (!graphLoaded) {
      console.warn("Graph.js failed to load; Graph API features disabled.");
    }

    isMobile =
      Office.context.mailbox.diagnostics.hostName === "OutlookAndroid" ||
      Office.context.mailbox.diagnostics.hostName === "OutlookIOS";

    isClassicOutlook = Office.context.mailbox.diagnostics.hostName === "Outlook";

    logger.log("info", "Office.onReady", {
      host: Office.context?.mailbox?.diagnostics?.hostName,
      version: Office.context?.mailbox?.diagnostics?.hostVersion,
      isMobile,
      isClassicOutlook,
    });
    Office.actions.associate("addSignatureMona", addSignatureMona);
    Office.actions.associate("addSignatureMorgan", addSignatureMorgan);
    Office.actions.associate("addSignatureMorven", addSignatureMorven);
    Office.actions.associate("addSignatureM2", addSignatureM2);
    Office.actions.associate("addSignatureM3", addSignatureM3);
  });
}

/**
 * Adds a signature to the email and saves it to localStorage.
 * @param {string} signatureKey - The signature key (e.g., "m3Signature").
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 * @param {boolean} isAutoApplied - Whether the signature is auto-applied.
 */
async function addSignature(signatureKey, event, isAutoApplied = false) {
  const item = Office.context.mailbox.item;

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
            if (!isAutoApplied) {
              event.completed();
              resolve();
            } else {
              // Move the completeWithStateFn call outside the callback
              reject(new Error(asyncResult.error.message));
            }
          }
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
        });
      });
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
    await completeWithState("Error", error.message, true);
  }
}

/**
 * Validates the email signature on send.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
async function validateSignature(event) {
  displayNotification(
    `Info`,
    `Platform: ${Office.context.mailbox.diagnostics.hostName}, Version: ${Office.context.mailbox.diagnostics.hostVersion}`
  );

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
 * Handles the new message compose event, applying the appropriate signature for reply/forward or new messages.
 * @param {Object} event - The event object from Office.js.
 */
async function onNewMessageComposeHandler(event) {
  const item = Office.context.mailbox.item;

  displayNotification(
    `Info`,
    `Platform: ${Office.context.mailbox.diagnostics.hostName}, Version: ${Office.context.mailbox.diagnostics.hostVersion}`
  );
  const isReplyOrForward = await SignatureManager.isReplyOrForward(item);

  logger.log("info", "onNewMessageComposeHandler", {
    isReplyOrForward,
    isMobile,
  });

  try {
    if (isReplyOrForward) {
      logger.log("info", "onNewMessageComposeHandler", { status: "Processing reply/forward email" });

      let messageId;
      if (isMobile) {
        messageId = Office.context.mailbox.item.conversationId;
      } else {
        const itemIdResult = await new Promise((resolve) => item.getItemIdAsync((asyncResult) => resolve(asyncResult)));
        if (itemIdResult.status !== Office.AsyncResultStatus.Succeeded) {
          throw new Error(itemIdResult.error.message);
        }
        messageId = itemIdResult.value;
        logger.log("info", "getItemIdAsync for OWA/Classic", { messageId });
      }

      const email = await fetchEmailById(messageId);
      const emailBody = email.body?.content || "";
      const extractedSignature = SignatureManager.extractSignature(emailBody);

      if (!extractedSignature) {
        logger.log("warn", "onNewMessageComposeHandler", { status: "No signature found in email" });
        await completeWithState(
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
        await completeWithState(
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

      localStorage.removeItem("tempSignature");
      localStorage.setItem("tempSignature", matchedSignatureKey);
      await addSignature(matchedSignatureKey, event, true);
      await completeWithState(event, null, null);
    } else {
      if (isMobile) {
        const mobileDefaultSignatureKey = localStorage.getItem("mobileDefaultSignature");
        if (mobileDefaultSignatureKey) {
          localStorage.removeItem("tempSignature");
          localStorage.setItem("tempSignature", mobileDefaultSignatureKey);
          await addSignature(mobileDefaultSignatureKey, event, true);
          await completeWithState(event, null, null);
        } else {
          await completeWithState(event, "Info", "Please select an M3 signature from the task pane.");
        }
      } else {
        await completeWithState(event, "Info", "Please select an M3 signature from the ribbon.");
      }
    }
  } catch (error) {
    logger.log("error", "onNewMessageComposeHandler", { error: error.message, stack: error.stack });
    if (isMobile) {
      await appendDebugLogToBody(item, "Message", error.message, "Stack", error.stack);
    }
    await completeWithState(event, "Error", `Failed to process compose event: ${error.message}`);
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

Office.actions.associate("validateSignature", validateSignature);
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
