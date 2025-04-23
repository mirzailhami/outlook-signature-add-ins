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
  Office.actions.associate("validateSignatureButton", validateSignatureButton);
});

function showNotification(type, message, persistent) {
  try {
    const item = Office.context.mailbox.item;
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
  } catch (error) {
    console.error("Error in showNotification:", error);
  }
}

function checkForOutlookVersion() {
  return Office.context.mailbox.diagnostics.hostName;
}

function showErrorDialog(message, event) {
  console.log("showErrorDialog called:", message);
  try {
    Office.context.ui.displayDialogAsync(
      `${process.env.ASSET_BASE_URL}/error.html?message=${encodeURIComponent(message)}`,
      { height: 20, width: 30, displayInIframe: true },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to open dialog:", asyncResult.error.message);
          event.completed();
          return;
        }
        const dialog = asyncResult.value;
        console.log("Dialog opened successfully");
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          console.log("Dialog message received:", arg.message);
          dialog.close();
          event.completed();
        });
      }
    );
  } catch (error) {
    console.error("Error in showErrorDialog:", error);
    event.completed();
  }
}

function validateSignatureButton(event) {
  console.log("validateSignatureButton triggered");
  try {
    const item = Office.context.mailbox.item;
    item.body.getAsync("html", (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get body:", result.error.message);
        showErrorDialog("Failed to load email body.", event);
        return;
      }
      const body = result.value;
      console.log("Raw email body:", body.slice(0, 200));
      let signature;
      const hostName = checkForOutlookVersion();
      if (hostName === "Outlook") {
        console.log("Running in classic Outlook");
        signature = extractSignatureForOutlookClassic(body);
      } else {
        console.log("Running in newer Outlook or Outlook on the web");
        signature = extractSignature(body);
      }
      const initialSignature = localStorage.getItem("initialSignature");
      let initialSignatureFromStore;
      if (hostName === "Outlook") {
        initialSignatureFromStore = extractSignatureFromStoreForOutlookClassic(initialSignature);
      } else {
        initialSignatureFromStore = extractSignatureFromStore(initialSignature);
      }

      console.log("Raw signature:", signature?.slice(0, 200));
      console.log("Raw stored signature:", initialSignatureFromStore?.slice(0, 200));

      if (!signature || !initialSignatureFromStore) {
        console.log("Missing signature");
        showErrorDialog("Email is missing the M3 required signature. Please select a signature.", event);
        return;
      }

      const cleanSignature = normalizeSignature(signature);
      const cleanStoredSignature = normalizeSignature(initialSignatureFromStore);

      console.log("Normalized signature:", cleanSignature);
      console.log("Normalized stored signature:", cleanStoredSignature);

      if (cleanSignature !== cleanStoredSignature) {
        console.log("Signature modification detected");
        showErrorDialog("Selected M3 signature has been modified. M3 email signatures cannot be modified.", event);
      } else {
        console.log("Signature is valid");
        showNotification("Info", "Signature is valid.", false);
        event.completed();
      }
    });
  } catch (error) {
    console.error("Error in validateSignatureButton:", error);
    showErrorDialog("Unexpected error validating signature.", event);
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
  const signatureDivRegex = /<div\s+class="Signature"[^>]*>.*?<table\s+class="MsoNormalTable"[^>]*>(.*?)(?:<\/table>|<\/div>)/is;
  const match = body.match(signatureDivRegex);
  if (match && match[1].includes("m3.signature@m3wind.com")) {
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
  const signatureDivRegex = /<table\s+class="MsoNormalTable"[^>]*>(.*?)(?:<\/table>|<\/div>)/is;
  const match = body.match(signatureDivRegex);
  if (match && match[1].includes("m3.signature@m3wind.com")) {
    console.log("Extracted signature via regex:", match[1].slice(0, 200));
    return match[1];
  }

  console.log("No signature found");
  return null;
}

function extractSignatureFromStore(body) {
  console.log("extractSignatureFromStore called, body sample:", body?.slice(0, 200));
  if (!body) return null;
  const signatureDivRegex = /<div\s+class="Signature"[^>]*>.*?<table\s+class="MsoNormalTable"[^>]*>(.*?)(?:<\/table>|<\/div>)/is;
  const match = body.match(signatureDivRegex);
  if (match) {
    console.log("Extracted stored signature:", match[1].slice(0, 200));
    return match[1];
  }
  console.log("No table found in stored signature, returning full body");
  return body;
}

function extractSignatureFromStoreForOutlookClassic(body) {
  console.log("extractSignatureFromStoreForOutlookClassic called, body sample:", body?.slice(0, 200));
  if (!body) return null;
  const signatureDivRegex = /<table\s+class="MsoNormalTable"[^>]*>(.*?)(?:<\/table>|<\/div>)/is;
  const match = body.match(signatureDivRegex);
  if (match) {
    console.log("Extracted stored signature:", match[1].slice(0, 200));
    return match[1];
  }
  console.log("No table found in stored signature, returning full body");
  return body;
}

function normalizeSignature(sig) {
  if (!sig) return "";
  return sig
    .replace(/<[^>]*>?/gm, '')
    .replace(/\s+/g, '')
    .toLowerCase()
    .replace(/(mona|m3)$/, '');
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
            console.log(`${signatureKey} signature set`);
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