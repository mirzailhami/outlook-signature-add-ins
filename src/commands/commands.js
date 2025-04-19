/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(() => {
  console.log("office.onready called");
  Office.actions.associate("addSignatureMona", addSignatureMona);
  Office.actions.associate("addSignatureMorgan", addSignatureMorgan);
  Office.actions.associate("addSignatureMorven", addSignatureMorven);
  Office.actions.associate("addSignatureM2", addSignatureM2);
  Office.actions.associate("addSignatureM3", addSignatureM3);
  Office.actions.associate("validateSignature", validateSignature);
  Office.actions.associate("onMessageCompose", onMessageCompose);
});

function onMessageCompose(event) {
  console.log("onMessageCompose called");
  checkForReplyOrForward(Office.context.mailbox.item)
    .then((isReplyOrForward) => {
      if (isReplyOrForward) {
        const lastSignature = localStorage.getItem("initialSignature");
        console.log("Last signature from localStorage:", lastSignature);
        if (lastSignature) {
          Office.context.mailbox.item.body.setSignatureAsync(
            "<!-- signature -->" + lastSignature,
            { coercionType: Office.CoercionType.Html },
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Failed to set signature on compose:", asyncResult.error.message);
              } else {
                console.log("Signature set on reply/forward compose");
              }
              event.completed();
            }
          );
        } else {
          console.log("No previous signature found");
          event.completed();
        }
      } else {
        console.log("Not a reply/forward");
        event.completed();
      }
    })
    .catch((error) => {
      console.error("Error checking reply/forward:", error);
      event.completed();
    });
}

function checkForOutlookVersion() {
  return Office.context.mailbox.diagnostics.hostName;
}

function checkForReplyOrForward(mailItem) {
  return new Promise((resolve, reject) => {
    mailItem.subject.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const isReplyOrForward =
          result.value.includes("Re:") ||
          result.value.includes("Fw:") ||
          result.value.includes("RE:") ||
          result.value.includes("FW:");
        console.log("Subject:", result.value, "isReplyOrForward:", isReplyOrForward);
        resolve(isReplyOrForward);
      } else {
        console.error("Failed to get subject:", result.error.message);
        reject(new Error("Failed to get subject: " + result.error.message));
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
          const senderInitialSavedSignature = localStorage.getItem("initialSignature");
          if (senderInitialSavedSignature) {
            Office.context.mailbox.item.body.setSignatureAsync(
              "<!-- signature -->" + senderInitialSavedSignature,
              { coercionType: Office.CoercionType.Html },
              (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  console.error("Failed to restore signature:", asyncResult.error.message);
                  event.completed({ allowEvent: false });
                } else {
                  console.log("Signature restored");
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
  console.log("validateSignature called");
  const item = Office.context.mailbox.item;
  isExternalEmail(item)
    .then((isExternalEmail) => {
      console.log("external email:", isExternalEmail);
      checkForReplyOrForward(item)
        .then((isReplyOrForward) => {
          item.body.getAsync("html", (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
              console.error("Failed to get body:", result.error.message);
              event.completed({ allowEvent: false });
              return;
            }
            const initialBody = result.value;
            const hostName = checkForOutlookVersion();
            const initialSignature =
              hostName === "Outlook" ? extractSignatureForOutlookClassic(initialBody) : extractSignature(initialBody);

            // Handle reply/forward without MessageCompose
            if (isReplyOrForward && !isExternalEmail && !initialSignature) {
              console.log("No signature in reply/forward, applying last signature");
              const lastSignature = localStorage.getItem("initialSignature");
              if (lastSignature) {
                item.body.setSignatureAsync(
                  "<!-- signature -->" + lastSignature,
                  { coercionType: Office.CoercionType.Html },
                  (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                      console.error("Failed to set signature:", asyncResult.error.message);
                      event.completed({ allowEvent: false });
                    } else {
                      console.log("Signature set for reply/forward");
                      event.completed({ allowEvent: false });
                    }
                  }
                );
              } else {
                showErrorDialog("M3 signature automatically added based on your previous email.", event);
              }
            } else if (!initialSignature) {
              showErrorDialog(
                "Email is missing the M3 required signature. Please select an appropriate email signature for your email.",
                event
              );
            } else {
              validateSignatureChanges(item, initialSignature, hostName, event);
            }
          });
        })
        .catch((error) => {
          console.error("Error checking reply/forward:", error);
          event.completed({ allowEvent: false });
        });
    })
    .catch((error) => {
      console.error("Error checking external email:", error);
      event.completed({ allowEvent: false });
    });
}

function validateSignatureChanges(item, initialSignature, hostName, event) {
  item.body.getAsync("html", (newResult) => {
    if (newResult.status === Office.AsyncResultStatus.Succeeded) {
      const newBody = newResult.value;
      const newSignature =
        hostName === "Outlook" ? extractSignatureForOutlookClassic(newBody) : extractSignature(newBody);
      const initialSavedSignature = localStorage.getItem("initialSignature");
      const initialSignatureFromStore =
        hostName === "Outlook"
          ? extractSignatureFromStoreForOutlookClassic(initialSavedSignature)
          : extractSignatureFromStore(initialSavedSignature);
      const newSignatureForModificationCheck = newSignature
        ? newSignature.replace(/<[^>]*>?/gm, "").replace(/\s+/g, "")
        : "";
      const initialSignatureFromStoreClean = initialSignatureFromStore
        ? initialSignatureFromStore.replace(/<[^>]*>?/gm, "").replace(/\s+/g, "")
        : "";
      console.log("newSignatureForModificationCheck:", newSignatureForModificationCheck);
      console.log("initialSignatureFromStoreClean:", initialSignatureFromStoreClean);
      if (newSignatureForModificationCheck !== initialSignatureFromStoreClean) {
        console.log("Signature modified detected");
        showErrorDialog(
          "Selected M3 signature has been modified. M3 email signature is prohibited from modification. Restoring the original signature.",
          event,
          true
        );
      } else {
        console.log("Signature unchanged");
        event.completed({ allowEvent: true });
      }
    } else {
      console.error("Failed to get new body:", newResult.error.message);
      event.completed({ allowEvent: false });
    }
  });
}

function extractSignature(body) {
  const signatureDivRegex = /<div\s+id="Signature">(.*?)<\/table>/s;
  const match = body.match(signatureDivRegex);
  return match ? match[1] : null;
}

function extractSignatureForOutlookClassic(body) {
  const signatureDivRegex = /<table\s+class=MsoNormalTable[^>]*>(.*?)<\/table>/is;
  const match = body.match(signatureDivRegex);
  return match ? match[1] : null;
}

function extractSignatureFromStore(body) {
  const signatureDivRegex = /<div\s+class="Signature">(.*?)<\/div>/s;
  const match = body.match(signatureDivRegex);
  return match ? match[1] : null;
}

function extractSignatureFromStoreForOutlookClassic(body) {
  const signatureDivRegex = /<table class="MsoNormalTable"[^>]*>(.*?)<\/table>/is;
  const match = body.match(signatureDivRegex);
  return match ? match[1] : null;
}

function isExternalEmail(mailItem) {
  return new Promise((resolve, reject) => {
    const hostName = checkForOutlookVersion();
    if (hostName === "Outlook") {
      resolve(false);
    } else {
      if (mailItem.inReplyTo.indexOf("OUTLOOK.COM") === -1) {
        resolve(true);
      } else {
        resolve(false);
      }
    }
  });
}

function addSignatureMona(event) {
  let localTemplate = localStorage.getItem("monaSignature");
  if (localTemplate) {
    Office.context.mailbox.item.body.setSignatureAsync(
      "<!-- signature -->" + localTemplate,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to set Mona signature:", asyncResult.error.message);
          event.completed();
        } else {
          console.log("Mona signature set");
          localStorage.setItem("initialSignature", localTemplate);
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
        signatureUrl += data.result[0].url;
        fetch(signatureUrl)
          .then((response) => response.json())
          .then((data) => {
            let template = data.result;
            template = template.replace("{First name} ", Office.context.mailbox.userProfile.displayName);
            template = template.replace("{Last name}", "");
            template = template.replaceAll("{E-mail}", Office.context.mailbox.userProfile.emailAddress);
            template = template.replace("{Title}", "");
            Office.context.mailbox.item.body.setSignatureAsync(
              "<!-- signature -->" + template,
              { coercionType: Office.CoercionType.Html },
              (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  console.error("Failed to set Mona signature:", asyncResult.error.message);
                  event.completed();
                } else {
                  console.log("Mona signature set");
                  localStorage.setItem("monaSignature", template);
                  localStorage.setItem("initialSignature", template);
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

function addSignatureMorgan(event) {
  let localTemplate = localStorage.getItem("morganSignature");
  if (localTemplate) {
    Office.context.mailbox.item.body.setSignatureAsync(
      "<!-- signature -->" + localTemplate,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to set Morgan signature:", asyncResult.error.message);
          event.completed();
        } else {
          console.log("Morgan signature set");
          localStorage.setItem("initialSignature", localTemplate);
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
        signatureUrl += data.result[1].url;
        fetch(signatureUrl)
          .then((response) => response.json())
          .then((data) => {
            let template = data.result;
            template = template.replace("{First name} ", Office.context.mailbox.userProfile.displayName);
            template = template.replace("{Last name}", "");
            template = template.replaceAll("{E-mail}", Office.context.mailbox.userProfile.emailAddress);
            template = template.replace("{Title}", "");
            Office.context.mailbox.item.body.setSignatureAsync(
              "<!-- signature -->" + template,
              { coercionType: Office.CoercionType.Html },
              (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  console.error("Failed to set Morgan signature:", asyncResult.error.message);
                  event.completed();
                } else {
                  console.log("Morgan signature set");
                  localStorage.setItem("morganSignature", template);
                  localStorage.setItem("initialSignature", template);
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

function addSignatureMorven(event) {
  let localTemplate = localStorage.getItem("morvenSignature");
  if (localTemplate) {
    Office.context.mailbox.item.body.setSignatureAsync(
      "<!-- signature -->" + localTemplate,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to set Morven signature:", asyncResult.error.message);
          event.completed();
        } else {
          console.log("Morven signature set");
          localStorage.setItem("initialSignature", localTemplate);
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
        signatureUrl += data.result[2].url;
        fetch(signatureUrl)
          .then((response) => response.json())
          .then((data) => {
            let template = data.result;
            template = template.replace("{First name} ", Office.context.mailbox.userProfile.displayName);
            template = template.replace("{Last name}", "");
            template = template.replaceAll("{E-mail}", Office.context.mailbox.userProfile.emailAddress);
            template = template.replace("{Title}", "");
            Office.context.mailbox.item.body.setSignatureAsync(
              "<!-- signature -->" + template,
              { coercionType: Office.CoercionType.Html },
              (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  console.error("Failed to set Morven signature:", asyncResult.error.message);
                  event.completed();
                } else {
                  console.log("Morven signature set");
                  localStorage.setItem("morvenSignature", template);
                  localStorage.setItem("initialSignature", template);
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

function addSignatureM2(event) {
  let localTemplate = localStorage.getItem("m2Signature");
  if (localTemplate) {
    Office.context.mailbox.item.body.setSignatureAsync(
      "<!-- signature -->" + localTemplate,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to set M2 signature:", asyncResult.error.message);
          event.completed();
        } else {
          console.log("M2 signature set");
          localStorage.setItem("initialSignature", localTemplate);
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
        signatureUrl += data.result[3].url;
        fetch(signatureUrl)
          .then((response) => response.json())
          .then((data) => {
            let template = data.result;
            template = template.replace("{First name} ", Office.context.mailbox.userProfile.displayName);
            template = template.replace("{Last name}", "");
            template = template.replaceAll("{E-mail}", Office.context.mailbox.userProfile.emailAddress);
            template = template.replace("{Title}", "");
            Office.context.mailbox.item.body.setSignatureAsync(
              "<!-- signature -->" + template,
              { coercionType: Office.CoercionType.Html },
              (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  console.error("Failed to set M2 signature:", asyncResult.error.message);
                  event.completed();
                } else {
                  console.log("M2 signature set");
                  localStorage.setItem("m2Signature", template);
                  localStorage.setItem("initialSignature", template);
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

function addSignatureM3(event) {
  let localTemplate = localStorage.getItem("m3Signature");
  if (localTemplate) {
    Office.context.mailbox.item.body.setSignatureAsync(
      "<!-- signature -->" + localTemplate,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to set M3 signature:", asyncResult.error.message);
          event.completed();
        } else {
          console.log("M3 signature set");
          localStorage.setItem("initialSignature", localTemplate);
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
        signatureUrl += data.result[4].url;
        fetch(signatureUrl)
          .then((response) => response.json())
          .then((data) => {
            let template = data.result;
            template = template.replace("{First name} ", Office.context.mailbox.userProfile.displayName);
            template = template.replace("{Last name}", "");
            template = template.replaceAll("{E-mail}", Office.context.mailbox.userProfile.emailAddress);
            template = template.replace("{Title}", "");
            Office.context.mailbox.item.body.setSignatureAsync(
              "<!-- signature -->" + template,
              { coercionType: Office.CoercionType.Html },
              (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  console.error("Failed to set M3 signature:", asyncResult.error.message);
                  event.completed();
                } else {
                  console.log("M3 signature set");
                  localStorage.setItem("m3Signature", template);
                  localStorage.setItem("initialSignature", template);
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
