/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(() => {
  console.log("Office.onReady called");
  Office.actions.associate("addSignatureMona", addSignatureMona);
  Office.actions.associate("addSignatureMorgan", addSignatureMorgan);
  Office.actions.associate("addSignatureMorven", addSignatureMorven);
  Office.actions.associate("addSignatureM2", addSignatureM2);
  Office.actions.associate("addSignatureM3", addSignatureM3);
  Office.actions.associate("validateSignature", validateSignature);
  startReplyForwardPolling();
});

function startReplyForwardPolling() {
  console.log("startReplyForwardPolling initiated");
  const interval = setInterval(() => {
    const item = Office.context.mailbox?.item;
    if (item && item.itemType === Office.MailboxEnums.ItemType.Message) {
      checkForReplyOrForward(item)
        .then((isReplyOrForward) => {
          if (isReplyOrForward) {
            console.log("Polling detected reply/forward");
            const lastSignature = localStorage.getItem("lastSentSignature");
            if (lastSignature) {
              item.body.getAsync("html", (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                  const body = result.value;
                  const hostName = checkForOutlookVersion();
                  const signature = extractSignature(body, hostName);
                  if (!signature) {
                    console.log("No signature found, applying last signature");
                    item.body.setSignatureAsync(
                      "<!-- signature -->" + lastSignature,
                      { coercionType: Office.CoercionType.Html },
                      (asyncResult) => {
                        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                          console.error("Polling setSignatureAsync failed:", asyncResult.error.message);
                        } else {
                          console.log("Signature set by polling for reply/forward");
                        }
                      }
                    );
                  }
                }
              });
            }
          }
        })
        .catch((error) => {
          console.error("Polling error:", error);
        });
    }
  }, 1000); // Poll every 1 second
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
        reject(new Error("Failed to get subject"));
      }
    });
  });
}

function showErrorDialog(message, event, restoreSignature = false) {
  console.log("showErrorDialog called:", message);
  Office.context.ui.displayDialogAsync(
    `https://localhost:3000/error.html?message=${encodeURIComponent(message)}`,
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
            Office.context.mailbox.item.body.setSignatureAsync(
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
              event.completed({ allowEvent: false });
              return;
            }
            const body = result.value;
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
  const regex = hostName === "Outlook"
    ? /<table\s+class=MsoNormalTable[^>]*>(.*?)<\/table>/is
    : /<div\s+id="Signature">(.*?)<\/table>/s;
  const match = body.match(regex);
  return match ? match[1] : null;
}

function extractSignatureFromStore(body, hostName) {
  if (!body) return null;
  const regex = hostName === "Outlook"
    ? /<table\s+class="MsoNormalTable"[^>]*>(.*?)<\/table>/is
    : /<div\s+class="Signature">(.*?)<\/div>/s;
  const match = body.match(regex);
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
  const localTemplate = localStorage.getItem(signatureKey);
  if (localTemplate) {
    Office.context.mailbox.item.body.setSignatureAsync(
      "<!-- signature -->" + localTemplate,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(`Failed to set ${signatureKey} signature:`, asyncResult.error.message);
          event.completed();
        } else {
          console.log(`${signatureKey} signature set`);
          localStorage.setItem("initialSignature", localTemplate);
          localStorage.setItem("lastSentSignature", localTemplate);
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
            Office.context.mailbox.item.body.setSignatureAsync(
              "<!-- signature -->" + template,
              { coercionType: Office.CoercionType.Html },
              (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  console.error(`Failed to set ${signatureKey} signature:`, asyncResult.error.message);
                  event.completed();
                } else {
                  console.log(`${signatureKey} signature set`);
                  localStorage.setItem(signatureKey, template);
                  localStorage.setItem("initialSignature", template);
                  localStorage.setItem("lastSentSignature", template);
                  event.completed();
                }
              }
            );
          })
          .catch((err) => {
            console.error("Error fetching signature:", err);
            event.completed();
          });
      })
      .catch((err) => {
        console.error("Error fetching ribbons:", err);
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