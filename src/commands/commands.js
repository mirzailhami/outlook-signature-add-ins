/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

console.log("commands.js loaded");

Office.initialize = function () {
  console.log("Office.initialize called as fallback");
};

// Force Office.onReady retry
function initializeWithRetry(attempt = 1, maxAttempts = 5) {
  console.log(`initializeWithRetry attempt ${attempt}`);
  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
      console.log("Office.onReady called, item:", Office.context.mailbox?.item, "host:", Office.context.mailbox.diagnostics.hostName);
      Office.actions.associate("addSignatureMona", addSignatureMona);
      Office.actions.associate("addSignatureMorgan", addSignatureMorgan);
      Office.actions.associate("addSignatureMorven", addSignatureMorven);
      Office.actions.associate("addSignatureM2", addSignatureM2);
      Office.actions.associate("addSignatureM3", addSignatureM3);
      Office.actions.associate("validateSignature", validateSignature);
      Office.actions.associate("onItemChanged", onItemChanged);
      tryApplySignatureWithRetry();
      // Poll for compose mode
      setInterval(() => {
        if (Office.context.mailbox?.item?.itemType === Office.MailboxEnums.ItemType.Message) {
          console.log("Compose mode detected via polling");
          tryApplySignatureWithRetry();
        }
      }, 500);
    } else {
      console.error("Unsupported host:", info);
    }
  }).catch((error) => {
    console.error("Office.onReady failed:", error);
    if (attempt < maxAttempts) {
      setTimeout(() => initializeWithRetry(attempt + 1, maxAttempts), 1000);
    } else {
      console.error("Max Office.onReady retries reached");
    }
  });
}

initializeWithRetry();

function onItemChanged(event) {
  console.log("onItemChanged triggered");
  // tryApplySignatureWithRetry();
  if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message && Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.MessageCompose) {
    tryApplySignatureWithRetry();
  }
  // event.completed();
}

function tryApplySignatureWithRetry(attempt = 1, maxAttempts = 15) {
  console.log(`tryApplySignatureWithRetry attempt ${attempt}`);
  const item = Office.context.mailbox?.item;
  if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) {
    console.log("No valid message item, retrying if attempts remain");
    if (attempt < maxAttempts) {
      setTimeout(() => tryApplySignatureWithRetry(attempt + 1, maxAttempts), 1000);
    } else {
      console.error("Max attempts reached, no valid item");
      showNotification("Error", "Failed to load M3 signature. Please select a signature manually.", true);
    }
    return;
  }

  showNotification("Info", "Loading M3 signature...", false);
  checkForReplyOrForward(item)
    .then((isReplyOrForward) => {
      console.log("isReplyOrForward:", isReplyOrForward);
      if (isReplyOrForward) {
        const lastSignature = localStorage.getItem("lastSentSignature");
        console.log("Last sent signature:", lastSignature);
        if (lastSignature) {
          item.body.getAsync("html", (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const body = result.value;
              console.log("Email body sample:", body.slice(0, 1000));
              const hostName = checkForOutlookVersion();
              const signature = extractSignature(body, hostName);
              if (!signature) {
                console.log("No signature found, applying last signature");
                item.body.setSignatureAsync(
                  "<!-- signature -->" + lastSignature,
                  { coercionType: Office.CoercionType.Html },
                  (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                      console.error("setSignatureAsync failed:", asyncResult.error.name, asyncResult.error.message, asyncResult.error.code);
                      showNotification("Error", "Failed to apply M3 signature.", true);
                    } else {
                      console.log("Signature applied successfully");
                      showNotification("Info", "M3 signature applied.", false);
                    }
                  }
                );
              } else {
                console.log("Signature already present");
                showNotification("Info", "M3 signature already present.", false);
              }
            } else {
              console.error("Failed to get body:", result.error.message);
              showNotification("Error", "Failed to load email body.", true);
            }
          });
        } else {
          console.log("No last signature found");
          showNotification("Error", "No previous M3 signature found. Please select a signature.", true);
        }
      } else {
        console.log("Not a reply/forward");
        showNotification("Info", "No signature applied for new email.", false);
      }
    })
    .catch((error) => {
      console.error("Error checking reply/forward:", error);
      showNotification("Error", "Failed to detect reply/forward status.", true);
    });
}

function showNotification(type, message, persistent) {
  const item = Office.context.mailbox?.item;
  if (!item) {
    console.error("No item for notification");
    return;
  }
  const messageId = type === "Error" ? "SignatureError" : "SignatureInfo";
  item.notificationMessages.replaceAsync(messageId, {
    type: type === "Error" ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: message,
    icon: "Icon.16x16",
    persistent: persistent
  }, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error(`Failed to show ${type} notification:`, asyncResult.error.message);
    } else {
      console.log(`${type} notification shown: ${message}`);
    }
  });
}

function checkForOutlookVersion() {
  return Office.context.mailbox.diagnostics.hostName;
}

function checkForReplyOrForward(mailItem) {
  return new Promise((resolve, reject) => {
    console.log("checkForReplyOrForward called");
    if (mailItem.itemType === Office.MailboxEnums.ItemType.Message && mailItem.conversationId) {
      console.log("Detected reply/forward via conversationId:", mailItem.conversationId);
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
        resolve(true); // Assume reply/forward to avoid blocking
      }
    });
  });
}

function showErrorDialog(message, event, restoreSignature = false) {
  console.log("showErrorDialog called:", message);
  Office.context.ui.displayDialogAsync(
    `${process.env.ASSET_BASE_URL}/error.html?message=${encodeURIComponent(message)}`,
    { height: 20, width: 30, displayInIframe: true },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to open dialog:", asyncResult.error.message);
        event.completed({ allowEvent: false });
        return;
      }
      const dialog = asyncResult.value;
      console.log("Dialog opened successfully");
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        console.log("Dialog message received:", arg.message);
        dialog.close();
        if (restoreSignature) {
          const initialSignature = localStorage.getItem("initialSignature");
          if (initialSignature) {
            Office.context.mailbox?.item.body.setSignatureAsync(
              "<!-- signature -->" + initialSignature,
              { coercionType: Office.CoercionType.Html },
              (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  console.error("Failed to restore signature:", asyncResult.error.message);
                  event.completed({ allowEvent: false });
                } else {
                  console.log("Signature restored in showErrorDialog");
                  localStorage.setItem("lastSentSignature", initialSignature);
                  event.completed({ allowEvent: false });
                }
              }
            );
          } else {
            event.completed({ allowEvent: false });
          }
        } else {
          event.completed({ allowEvent: false });
        }
      });
    }
  );
}

function validateSignature(event) {
  console.log("validateSignature triggered");
  const item = Office.context.mailbox?.item;
  if (!item) {
    console.error("No item available in validateSignature");
    event.completed({ allowEvent: false });
    return;
  }
  isExternalEmail(item)
    .then((isExternal) => {
      console.log("Is external email:", isExternal);
      checkForReplyOrForward(item)
        .then((isReplyOrForward) => {
          console.log("validateSignature isReplyOrForward:", isReplyOrForward);
          item.body.getAsync("html", (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
              console.error("Failed to get body:", result.error.message);
              event.completed({ allowEvent: false });
              return;
            }
            const body = result.value;
            console.log("ValidateSignature body sample:", body.slice(0, 1000));
            const hostName = checkForOutlookVersion();
            const signature = extractSignature(body, hostName);

            if (isReplyOrForward && !signature) {
              console.log("No signature in reply/forward, applying last signature");
              const lastSignature = localStorage.getItem("lastSentSignature");
              if (lastSignature) {
                item.body.setSignatureAsync(
                  "<!-- signature -->" + lastSignature,
                  { coercionType: Office.CoercionType.Html },
                  (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                      console.error("Failed to set signature in validateSignature:", asyncResult.error.message);
                      event.completed({ allowEvent: false });
                    } else {
                      console.log("Signature set for reply/forward in validateSignature");
                      event.completed({ allowEvent: true });
                    }
                  }
                );
              } else {
                showErrorDialog(
                  "No M3 signature found for reply/forward. Please select a signature.",
                  event
                );
              }
            } else if (!signature) {
              showErrorDialog(
                "Email is missing the M3 required signature. Please select an appropriate email signature.",
                event
              );
            } else {
              validateSignatureChanges(item, signature, hostName, event);
            }
          });
        })
        .catch((error) => {
          console.error("Error checking reply/forward in validateSignature:", error);
          event.completed({ allowEvent: false });
        });
    })
    .catch((error) => {
      console.error("Error checking external email:", error);
      event.completed({ allowEvent: false });
    });
}

function validateSignatureChanges(item, initialSignature, hostName, event) {
  item.body.getAsync("html", (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Failed to get new body:", result.error.message);
      event.completed({ allowEvent: false });
      return;
    }
    const newBody = result.value;
    const newSignature = extractSignature(newBody, hostName);
    const initialSavedSignature = localStorage.getItem("initialSignature");
    const storedSignature = extractSignatureFromStore(initialSavedSignature, hostName);

    const cleanNewSignature = newSignature ? newSignature.replace(/<[^>]*>?/gm, "").replace(/\s+/g, "") : "";
    const cleanStoredSignature = storedSignature ? storedSignature.replace(/<[^>]*>?/gm, "").replace(/\s+/g, "") : "";

    console.log("New signature (cleaned):", cleanNewSignature);
    console.log("Stored signature (cleaned):", cleanStoredSignature);

    if (cleanNewSignature !== cleanStoredSignature) {
      console.log("Signature modification detected");
      showErrorDialog(
        "Selected M3 signature has been modified. M3 email signatures cannot be modified. Restoring the original signature.",
        event,
        true
      );
    } else {
      console.log("Signature unchanged");
      localStorage.setItem("lastSentSignature", initialSavedSignature);
      event.completed({ allowEvent: true });
    }
  });
}

function extractSignature(body, hostName) {
  console.log("extractSignature called, hostName:", hostName, "body sample:", body.slice(0, 1000));
  const regex = hostName === "Outlook"
    ? /<table\s+class=["']?MsoNormalTable["']?[^>]*>(.*?)<\/table>/is
    : /<div\s+id=["']?Signature["']?>(.*?)(?:<\/table>|<\/div>)/is;
  const match = body.match(regex);
  console.log("Signature match:", match ? match[1] : null);
  return match ? match[1] : null;
}

function extractSignatureFromStore(body, hostName) {
  if (!body) return null;
  console.log("extractSignatureFromStore called, hostName:", hostName, "body sample:", body.slice(0, 1000));
  const regex = hostName === "Outlook"
    ? /<table\s+class=["']?MsoNormalTable["']?[^>]*>(.*?)<\/table>/is
    : /<div\s+(?:id|class)=["']?Signature["']?>(.*?)(?:<\/table>|<\/div>)/is;
  const match = body.match(regex);
  console.log("Stored signature match:", match ? match[1] : null);
  return match ? match[1] : null;
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
  const item = Office.context.mailbox?.item;
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
          console.log(`${signatureKey} signature set`);
          localStorage.setItem("initialSignature", localTemplate);
          localStorage.setItem("lastSentSignature", localTemplate);
          showNotification("Info", `${signatureKey} applied.`, false);
          event.completed();
        }
      }
    );
  } else {
    const initialUrl = "https://m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net/api/Ribbons/ribbons";
    let signatureUrl = "https://m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net/api/Signatures/signatures?signatureURL=";
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
                  console.log(`${signatureKey} signature set`);
                  localStorage.setItem(signatureKey, template);
                  localStorage.setItem("initialSignature", template);
                  localStorage.setItem("lastSentSignature", template);
                  showNotification("Info", `${signatureKey} applied.`, false);
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