/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

console.log("commands.js loaded");

// Log Office.js availability
console.log("Office:", typeof Office, Office);

Office.onReady(() => {
  console.log("Office.onReady called, host:", Office.context?.mailbox?.diagnostics?.hostName);
  console.log("localStorage initialSignature:", localStorage.getItem("initialSignature"));
  Office.actions.associate("addSignatureMona", addSignatureMona);
  Office.actions.associate("addSignatureMorgan", addSignatureMorgan);
  Office.actions.associate("addSignatureMorven", addSignatureMorven);
  Office.actions.associate("addSignatureM2", addSignatureM2);
  Office.actions.associate("addSignatureM3", addSignatureM3);
  Office.actions.associate("validateSignature", validateSignature);
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
});

function initializeAutoSignature(event) {
  console.log("initializeAutoSignature called");
  try {
    if (Office.context.mailbox.item) {
      console.log("Compose window detected, applying auto-signature");
      applyAutoSignature(event);
    } else {
      console.log("No mailbox item, starting polling");
      startPollingForCompose(event);
    }
  } catch (error) {
    console.error("Error in initializeAutoSignature:", error);
    startPollingForCompose(event);
  }
}

function startPollingForCompose(event) {
  console.log("Starting polling for compose window");
  let attempts = 0;
  const maxAttempts = 20;
  const pollInterval = setInterval(() => {
    attempts++;
    console.log(`Polling attempt ${attempts}`);
    try {
      if (
        Office.context.mailbox.item &&
        Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message
      ) {
        console.log("Compose window detected via polling");
        clearInterval(pollInterval);
        applyAutoSignature(event);
      } else if (attempts >= maxAttempts) {
        console.error("Polling failed: Max attempts reached");
        clearInterval(pollInterval);
        showNotification("Error", "Failed to detect compose window.", true);
        event?.completed();
      }
    } catch (error) {
      console.error("Error during polling:", error);
      if (attempts >= maxAttempts) {
        console.error("Polling failed: Max attempts reached");
        clearInterval(pollInterval);
        showNotification("Error", "Failed to detect compose window.", true);
        event?.completed();
      }
    }
  }, 1000);
}

function applyAutoSignature(event) {
  console.log("applyAutoSignature called");
  tryApplySignatureWithRetry(event);
}

function tryApplySignatureWithRetry(event, attempt = 1, maxAttempts = 10) {
  console.log(`tryApplySignatureWithRetry attempt ${attempt}`);
  try {
    const item = Office.context.mailbox.item;
    if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) {
      console.log("No valid message item, retrying if attempts remain");
      if (attempt < maxAttempts) {
        setTimeout(() => tryApplySignatureWithRetry(event, attempt + 1, maxAttempts), 1000);
      } else {
        console.error("Max attempts reached, no valid item");
        showNotification("Error", "Failed to load M3 signature automatically.", true);
        event?.completed();
      }
      return;
    }

    checkForReplyOrForward(item)
      .then((isReplyOrForward) => {
        console.log("isReplyOrForward:", isReplyOrForward);
        if (!isReplyOrForward) {
          console.log("New email, skipping auto-signature");
          showNotification("Info", "No signature applied for new email.", false);
          event?.completed();
          return;
        }

        showNotification("Info", "Loading M3 signature...", false);
        const lastSignature = localStorage.getItem("initialSignature");
        console.log("Last sent signature:", lastSignature?.slice(0, 200));
        if (lastSignature) {
          item.body.getAsync("html", (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const body = result.value;
              if (!body.includes("<!-- signature -->")) {
                console.log("No signature found, applying last signature");
                item.body.setSignatureAsync(
                  "<!-- signature -->" + lastSignature,
                  { coercionType: Office.CoercionType.Html },
                  (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                      console.error("setSignatureAsync failed:", asyncResult.error.message);
                      showNotification("Error", "Failed to apply M3 signature.", true);
                    } else {
                      console.log("Signature applied successfully");
                      localStorage.setItem("initialSignature", lastSignature);
                      showNotification("Info", "M3 signature applied.", false);
                    }
                    event?.completed();
                  }
                );
              } else {
                console.log("Signature already present");
                showNotification("Info", "M3 signature already present.", false);
                event?.completed();
              }
            } else {
              console.error("Failed to get body:", result.error.message);
              showNotification("Error", "Failed to load email body.", true);
              event?.completed();
            }
          });
        } else {
          console.log("No last signature found");
          showNotification("Info", "No previous M3 signature found. Please select a signature.", false);
          event?.completed();
        }
      })
      .catch((error) => {
        console.error("Error checking reply/forward:", error);
        showNotification("Error", "Failed to detect reply/forward status.", true);
        event?.completed();
      });
  } catch (error) {
    console.error("Unexpected error in tryApplySignatureWithRetry:", error);
    showNotification("Error", "Unexpected error applying signature.", true);
    event?.completed();
  }
}

function showNotification(type, message, persistent) {
  try {
    const item = Office.context.mailbox.item;
    if (!item) {
      console.error("No item for notification");
      return;
    }
    const messageId = type === "Error" ? "SignatureError" : "SignatureInfo";
    console.log(
      `Attempting to show ${type} notification: ${message}, persistent: ${persistent}, messageId: ${messageId}`
    );
    item.notificationMessages.replaceAsync(
      messageId,
      {
        type:
          type === "Error"
            ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
            : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: message,
        icon: "Icon.16x16",
        persistent: persistent,
      },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(`Failed to show ${type} notification:`, asyncResult.error.message);
        } else {
          console.log(`${type} notification shown successfully: ${message}`);
        }
      }
    );
  } catch (error) {
    console.error("Error in showNotification:", error);
  }
}

function showErrorNotification(message, event, restoreSignature = false) {
  console.log("showErrorNotification called:", message, "at:", new Date().toISOString());
  try {
    const item = Office.context.mailbox.item;
    if (!item) {
      console.error("No item for error notification");
      setTimeout(() => {
        console.log("event.completed called:", new Date().toISOString());
        event.completed({ allowEvent: false, errorMessage: message });
      }, 1000);
      return;
    }
    if (restoreSignature) {
      const initialSignature = localStorage.getItem("initialSignature");
      if (initialSignature) {
        item.body.setSignatureAsync(
          "<!-- signature -->" + initialSignature,
          { coercionType: Office.CoercionType.Html },
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              console.error("Failed to restore signature:", asyncResult.error.message);
              showNotification("Error", "Failed to restore signature.", true);
              setTimeout(() => {
                console.log("event.completed called:", new Date().toISOString());
                event.completed({ allowEvent: false, errorMessage: "Failed to restore signature." });
              }, 1000);
            } else {
              console.log("Signature restored in showErrorNotification");
              localStorage.setItem("initialSignature", initialSignature);
              item.body.getAsync("html", (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("Body after restoration:", result.value.slice(0, 200));
                }
              });
              showNotification("Error", message, true);
              setTimeout(() => {
                console.log("event.completed called:", new Date().toISOString());
                event.completed({ allowEvent: false, errorMessage: message });
              }, 1000);
            }
          }
        );
      } else {
        showNotification("Error", message, true);
        setTimeout(() => {
          console.log("event.completed called:", new Date().toISOString());
          event.completed({ allowEvent: false, errorMessage: "No signature to restore." });
        }, 1000);
      }
    } else {
      showNotification("Error", message, true);
      setTimeout(() => {
        console.log("event.completed called:", new Date().toISOString());
        event.completed({ allowEvent: false, errorMessage: message });
      }, 1000);
    }
  } catch (error) {
    console.error("Error in showErrorNotification:", error);
    setTimeout(() => {
      console.log("event.completed called:", new Date().toISOString());
      event.completed({ allowEvent: false, errorMessage: message });
    }, 1000);
  }
}

function checkForOutlookVersion() {
  return Office.context.mailbox.diagnostics.hostName;
}

function checkForReplyOrForward(mailItem) {
  return new Promise((resolve, reject) => {
    console.log("checkForReplyOrForward called, item:", mailItem);
    if (mailItem.itemType === Office.MailboxEnums.ItemType.Message && mailItem.conversationId) {
      console.log("Detected reply/forward via conversationId:", mailItem.conversationId);
      resolve(true);
      return;
    }
    if (mailItem.inReplyTo) {
      console.log("Detected reply/forward via inReplyTo:", mailItem.inReplyTo);
      resolve(true);
      return;
    }
    mailItem.subject.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const isReplyOrForward =
          result.value.toLowerCase().includes("re:") ||
          result.value.toLowerCase().includes("fw:") ||
          result.value.toLowerCase().includes("fwd:");
        console.log("Subject:", result.value, "isReplyOrForward:", isReplyOrForward);
        resolve(isReplyOrForward);
      } else {
        console.error("Failed to get subject:", result.error.message);
        reject(new Error("Failed to get subject"));
      }
    });
  });
}

function validateSignature(event) {
  console.log("validateSignature triggered");
  try {
    const item = Office.context.mailbox.item;
    isExternalEmail(item)
      .then((isExternal) => {
        console.log("Is external email:", isExternal);
        checkForReplyOrForward(item)
          .then((isReplyOrForward) => {
            console.log("validateSignature isReplyOrForward:", isReplyOrForward);
            item.body.getAsync("html", (result) => {
              if (result.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Failed to get body:", result.error.message);
                showErrorNotification("Failed to load email body.", event);
                return;
              }
              const body = result.value;
              console.log("Raw email body:", body.slice(0, 200));
              let initialSignature;
              const hostName = checkForOutlookVersion();
              if (hostName === "Outlook") {
                console.log("Running in classic Outlook");
                initialSignature = extractSignatureForOutlookClassic(body);
              } else {
                console.log("Running in newer Outlook or Outlook on the web");
                initialSignature = extractSignature(body);
              }

              if (isReplyOrForward && !initialSignature) {
                console.log("No signature in reply/forward, applying last signature");
                const lastSignature = localStorage.getItem("initialSignature");
                if (lastSignature) {
                  item.body.setSignatureAsync(
                    "<!-- signature -->" + lastSignature,
                    { coercionType: Office.CoercionType.Html },
                    (asyncResult) => {
                      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.error("Failed to set signature:", asyncResult.error.message);
                        showErrorNotification("Failed to apply M3 signature.", event);
                      } else {
                        console.log("Signature set for reply/forward");
                        event.completed({ allowEvent: true });
                      }
                    }
                  );
                } else {
                  showErrorNotification("No M3 signature found for reply/forward. Please select a signature.", event);
                }
              } else if (!initialSignature) {
                showErrorNotification(
                  "Email is missing the M3 required signature. Please select an appropriate email signature.",
                  event
                );
              } else {
                validateSignatureChanges(item, initialSignature, event, isReplyOrForward);
              }
            });
          })
          .catch((error) => {
            console.error("Error checking reply/forward:", error);
            showErrorNotification("Failed to detect reply/forward status.", event);
          });
      })
      .catch((error) => {
        console.error("Error checking external email:", error);
        showErrorNotification("Failed to check external email status.", event);
      });
  } catch (error) {
    console.error("Error in validateSignature:", error);
    showErrorNotification("Unexpected error validating signature.", event);
  }
}

function validateSignatureChanges(item, initialSignature, event, isReplyOrForward) {
  try {
    item.body.getAsync("html", (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get new body:", result.error.message);
        showErrorNotification("Failed to load email body.", event);
        return;
      }
      const newBody = result.value;
      console.log("Raw new email body:", newBody.slice(0, 200));
      const hostName = checkForOutlookVersion();
      let newSignature;
      if (hostName === "Outlook") {
        console.log("Running in classic Outlook");
        newSignature = extractSignatureForOutlookClassic(newBody);
      } else {
        console.log("Running in newer Outlook or Outlook on the web");
        newSignature = extractSignature(newBody);
      }
      const initialSavedSignature = localStorage.getItem("initialSignature");
      let initialSignatureFromStore;
      if (hostName === "Outlook") {
        initialSignatureFromStore = extractSignatureFromStoreForOutlookClassic(initialSavedSignature);
      } else {
        initialSignatureFromStore = extractSignatureFromStore(initialSavedSignature);
      }

      console.log("Raw initial signature:", initialSignature?.slice(0, 200));
      console.log("Raw new signature:", newSignature?.slice(0, 200));
      console.log("Raw stored signature:", initialSignatureFromStore?.slice(0, 200));

      if (!newSignature || !initialSignatureFromStore) {
        console.log("Missing signature data, treating as missing signature");
        showErrorNotification(
          "Email is missing the M3 required signature. Please select an appropriate email signature.",
          event
        );
        return;
      }

      const cleanNewSignature = normalizeSignature(newSignature);
      const cleanStoredSignature = normalizeSignature(initialSignatureFromStore);

      console.log("Normalized new signature:", cleanNewSignature);
      console.log("Normalized stored signature:", cleanStoredSignature);

      if (cleanNewSignature !== cleanStoredSignature) {
        console.log("Signature modification detected");
        showErrorNotification(
          "Selected M3 signature has been modified. M3 email signatures cannot be modified. Restoring the original signature.",
          event,
          true
        );
      } else {
        console.log("Signature unchanged, allowing send");
        localStorage.setItem("initialSignature", initialSavedSignature);
        event.completed({ allowEvent: true });
      }
    });
  } catch (error) {
    console.error("Error in validateSignatureChanges:", error);
    showErrorNotification("Unexpected error validating signature changes.", event);
  }
}

function extractSignature(body) {
  console.log("extractSignature called, body sample:", body.slice(0, 200));
  const marker = "<!-- signature -->";
  let startIndex = body.indexOf(marker);
  if (startIndex !== -1) {
    const signature = body.slice(startIndex + marker.length).trim();
    console.log("Extracted signature via marker:", signature.slice(0, 200));
    return signature;
  }

  console.log("No signature marker found, attempting regex detection");
  const signatureDivRegex = /<div\s+id="Signature">(.*?)<\/table>/s;
  const match = body.match(signatureDivRegex);
  if (match) {
    console.log("Extracted signature via regex:", match[1].slice(0, 200));
    return match[1];
  }

  console.log("No signature found");
  return null;
}

function extractSignatureForOutlookClassic(body) {
  console.log("extractSignatureForOutlookClassic called, body sample:", body.slice(0, 200));
  const marker = "<!-- signature -->";
  let startIndex = body.indexOf(marker);
  if (startIndex !== -1) {
    const signature = body.slice(startIndex + marker.length).trim();
    console.log("Extracted signature via marker:", signature.slice(0, 200));
    return signature;
  }

  console.log("No signature marker found, attempting regex detection");
  const signatureDivRegex = /<table\s+class=MsoNormalTable[^>]*>(.*?)<\/table>/is;
  const match = body.match(signatureDivRegex);
  if (match) {
    console.log("Extracted signature via regex:", match[1].slice(0, 200));
    return match[1];
  }

  console.log("No signature found");
  return null;
}

function extractSignatureFromStore(body) {
  console.log("extractSignatureFromStore called, body sample:", body?.slice(0, 200));
  if (!body) return null;
  const signatureDivRegex = /<div\s+class="Signature">(.*?)<\/div>/s;
  const match = body.match(signatureDivRegex);
  return match ? match[1] : body;
}

function extractSignatureFromStoreForOutlookClassic(body) {
  console.log("extractSignatureFromStoreForOutlookClassic called, body sample:", body?.slice(0, 200));
  if (!body) return null;
  const signatureDivRegex = /<table class="MsoNormalTable"[^>]*>(.*?)<\/table>/is;
  const match = body.match(signatureDivRegex);
  return match ? match[1] : body;
}

function normalizeSignature(sig) {
  if (!sig) return "";
  // Strip HTML tags and normalize whitespace
  return sig
    .replace(/<[^>]*>?/gm, "")
    .replace(/\s+/g, "")
    .toLowerCase();
}

function isExternalEmail(mailItem) {
  return new Promise((resolve) => {
    const hostName = checkForOutlookVersion();
    if (hostName === "Outlook") {
      resolve(false);
    } else {
      resolve(mailItem.inReplyTo && mailItem.inReplyTo.indexOf("OUTLOOK.COM") === -1);
    }
  });
}

function addSignature(signatureKey, signatureUrlIndex, event) {
  console.log(`addSignature called for ${signatureKey}`);
  try {
    const item = Office.context.mailbox.item;
    showNotification("Info", `Applying ${signatureKey}...`, false);
    const localTemplate = localStorage.getItem(signatureKey);
    if (localTemplate) {
      item.body.setSignatureAsync(
        "<!-- signature -->" + localTemplate,
        { coercionType: Office.CoercionType.Html },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error(`Failed to set ${signatureKey} signature:`, asyncResult.error.message);
            showNotification("Error", `Failed to apply ${signatureKey}.`, true);
            event.completed();
          } else {
            console.log(`${signatureKey} signature applied`);
            localStorage.setItem("initialSignature", localTemplate);
            localStorage.setItem("lastSentSignature", localTemplate);
            showNotification("Info", `${signatureKey} applied.`, false);
            item.body.getAsync("html", (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Body after signature set:", result.value.slice(0, 200));
                console.log("Marker present:", result.value.includes("<!-- signature -->"));
              }
            });
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
                    console.error(`Failed to set ${signatureKey} signature:`, asyncResult.error.message);
                    showNotification("Error", `Failed to apply ${signatureKey}.`, true);
                    event.completed();
                  } else {
                    console.log(`${signatureKey} signature applied`);
                    localStorage.setItem(signatureKey, template);
                    localStorage.setItem("initialSignature", template);
                    localStorage.setItem("lastSentSignature", template);
                    showNotification("Info", `${signatureKey} applied.`, false);
                    item.body.getAsync("html", (result) => {
                      if (result.status === Office.AsyncResultStatus.Succeeded) {
                        console.log("Body after signature set:", result.value.slice(0, 200));
                        console.log("Marker present:", result.value.includes("<!-- signature -->"));
                      }
                    });
                    event.completed();
                  }
                }
              );
            })
            .catch((err) => {
              console.error("Error fetching signature:", err);
              showNotification("Error", "Failed to fetch signature.", true);
              event.completed();
            });
        })
        .catch((err) => {
          console.error("Error fetching ribbons:", err);
          showNotification("Error", "Failed to fetch ribbons.", true);
          event.completed();
        });
    }
  } catch (error) {
    console.error(`Error in addSignature ${signatureKey}:`, error);
    showNotification("Error", `Failed to apply ${signatureKey}.`, true);
    event.completed();
  }
}

function addSignatureMona(event) {
  addSignature("monaSignature", 0, event);
}

function addSignatureMorgan(event) {
  addSignature("morganSignature", 1, event);
}

function addSignatureMorven(event) {
  addSignature("morvenSignature", 2, event);
}

function addSignatureM2(event) {
  addSignature("m2Signature", 3, event);
}

function addSignatureM3(event) {
  addSignature("m3Signature", 4, event);
}

function onNewMessageComposeHandler(event) {
  console.log("onNewMessageComposeHandler called");
  initializeAutoSignature(event);
}
