/**
 * Global variables for real-time monitoring
 */
let contentMonitoringInterval = null;
let lastKnownSignature = null;
let isMonitoringActive = false;
let currentItem = null;

/**
 * Initializes the Outlook add-in and associates event handlers.
 */
Office.onReady(() => {
  console.log({ event: "Office.onReady", host: Office.context?.mailbox?.diagnostics?.hostName });

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
 * Core signature management module.
 */
const SignatureManager = {
  /**
   * Extracts the signature from the email body.
   * @param {string} body - The email body HTML.
   * @returns {string|null} The extracted signature or null.
   */
  extractSignature(body) {
    console.log({ event: "extractSignature", bodyLength: body?.length });
    if (!body) return null;

    const marker = "<!-- signature -->";
    const startIndex = body.indexOf(marker);
    if (startIndex !== -1) {
      const endIndex = body.indexOf("</body>", startIndex);
      const signature = body.slice(startIndex + marker.length, endIndex !== -1 ? endIndex : undefined).trim();
      console.log({ event: "extractSignature", method: "marker", signature });
      return signature;
    }

    const regexes = [
      /<div\s+class="Signature"[^>]*>([\s\S]*?)$/is,
      /<div\s+id="Signature"[^>]*>([\s\S]*?)$/is,
      /<table[^>]*>([\s\S]*?)$/is,
    ];
    for (const regex of regexes) {
      const match = body.match(regex);
      if (match) {
        const signature = match[0].trim();
        console.log({ event: "extractSignature", method: regex.source, signature });
        return signature;
      }
    }

    console.log({ event: "extractSignature", status: "No signature found" });
    return null;
  },

  /**
   * Extracts the signature for classic Outlook.
   * @param {string} body - The email body HTML.
   * @returns {string|null} The extracted signature or null.
   */
  extractSignatureForOutlookClassic(body) {
    if (!body) return null;

    const marker = "<!-- signature -->";
    const startIndex = body.indexOf(marker);
    if (startIndex !== -1) {
      const endIndex = body.indexOf("</body>", startIndex);
      const signature = body.slice(startIndex + marker.length, endIndex !== -1 ? endIndex : undefined).trim();
      console.log({ event: "extractSignatureForOutlookClassic", method: "marker", signature });
      return signature;
    }

    const regex = /<table\s+class=MsoNormalTable[^>]*>([\s\S]*?)$/is;
    const match = body.match(regex);
    if (match) {
      const signature = match[0].trim();
      console.log({ event: "extractSignatureForOutlookClassic", method: "table", signature });
      return signature;
    }

    console.log({ event: "extractSignatureForOutlookClassic", status: "No signature found" });
    return null;
  },

  /**
   * Normalizes a signature for comparison by focusing on visible content.
   * @param {string} sig - The signature HTML.
   * @returns {string} The normalized signature.
   */
  normalizeSignature(sig) {
    if (!sig) return "";
    const htmlEntities = { " ": " ", "&": "&", "<": "<", ">": ">", '"': '"' };
    let normalized = sig;
    for (const [entity, char] of Object.entries(htmlEntities)) {
      normalized = normalized.replace(new RegExp(entity, "gi"), char);
    }
    normalized = normalized.replace(/<[^>]+>/g, "");
    const textarea = document.createElement("textarea");
    textarea.innerHTML = normalized;
    normalized = textarea.value
      .replace(/[\r\n]+/g, " ") // Replace newlines with a single space
      .replace(/\s*([.,:;])\s*/g, "$1") // Remove spaces around punctuation (e.g., "attachment. mona" -> "attachment.mona")
      .replace(/\s+/g, " ") // Collapse multiple spaces into one
      .replace(/\s*:\s*/g, ":") // Remove spaces around colons
      .replace(/\s+(email:)/gi, "$1") // Remove spaces before "email:"
      .trim() // Remove leading/trailing spaces
      .toLowerCase();
    console.log({ event: "normalizeSignature", raw: sig, normalized });
    return normalized;
  },

  /**
   * Normalizes a subject for comparison.
   * @param {string} subject - The email subject.
   * @returns {string} The normalized subject.
   */
  normalizeSubject(subject) {
    if (!subject) return "";
    return subject
      .replace(/^(re:|fw:|fwd:)\s*/i, "")
      .trim()
      .toLowerCase();
  },

  /**
   * Enhanced check for reply or forward, including inline detection.
   * @param {Office.MessageCompose} item - The email item.
   * @returns {Promise<boolean>} True if reply or forward.
   */
  async isReplyOrForward(item) {
    console.log({ event: "checkForReplyOrForward" });

    // Check conversation ID (most reliable for inline replies)
    if (item.itemType === Office.MailboxEnums.ItemType.Message && item.conversationId) {
      console.log({
        event: "checkForReplyOrForward",
        status: "Reply/forward detected via conversationId",
        conversationId: item.conversationId,
      });
      return true;
    }

    // Check inReplyTo header
    if (item.inReplyTo) {
      console.log({ event: "checkForReplyOrForward", status: "Reply detected via inReplyTo", inReplyTo: item.inReplyTo });
      return true;
    }

    // Enhanced subject checking for inline replies
    const subject = await new Promise((resolve) => item.subject.getAsync((result) => resolve(result.value || "")));
    const replyPrefixes = ["re:", "fw:", "fwd:", "aw:", "sv:", "vs:", "ref:", "reply:", "forward:"];
    const isReplyOrForward = replyPrefixes.some((prefix) => subject.toLowerCase().trim().startsWith(prefix));

    // Additional check: look for reply/forward indicators in the body
    if (!isReplyOrForward) {
      try {
        const body = await new Promise((resolve) => item.body.getAsync("text", (result) => resolve(result.value || "")));
        const bodyIndicators = ["-----original message-----", "from:", "sent:", "to:", "subject:", "> ", ">>"];
        const hasBodyIndicators = bodyIndicators.some(indicator => body.toLowerCase().includes(indicator));

        if (hasBodyIndicators) {
          console.log({ event: "checkForReplyOrForward", status: "Reply/forward detected via body indicators" });
          return true;
        }
      } catch (error) {
        console.log({ event: "checkForReplyOrForward", error: "Failed to check body indicators", message: error.message });
      }
    }

    console.log({ event: "checkForReplyOrForward", status: "Subject and body checked", isReplyOrForward, subject });
    return isReplyOrForward;
  },

  /**
   * Restores the signature to the email.
   * @param {Office.MessageCompose} item - The email item.
   * @param {string} signature - The signature to restore.
   * @param {string} signatureKey - The signature key.
   * @returns {Promise<boolean>} True if successful.
   */
  async restoreSignature(item, signature, signatureKey) {
    console.log({ event: "restoreSignatureAsync", signatureKey, cachedSignatureLength: signature?.length });
    if (!signature) {
      signature = localStorage.getItem(`signature_${signatureKey}`);
      console.log({
        event: "restoreSignatureAsync",
        status: "Falling back to signatureKey",
        fallbackLength: signature?.length,
      });
      if (!signature) return false;
    }

    const success = await new Promise((resolve) =>
      item.body.setSignatureAsync(
        "<!-- signature -->" + signature.trim(),
        { coercionType: Office.CoercionType.Html },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error({ event: "restoreSignatureAsync", error: asyncResult.error.message, signatureKey });
            resolve(false);
          } else {
            console.log({ event: "restoreSignatureAsync", status: "Signature set", signatureKey });
            resolve(true);
          }
        }
      )
    );

    if (success) {
      const body = await new Promise((resolve) =>
        item.body.getAsync("html", (result) =>
          resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : null)
        )
      );
      if (body) {
        const extracted = this.extractSignature(body);
        console.log({ event: "restoreSignatureAsync", status: "Body refreshed", extractedSignature: extracted });
        const normalizedExtracted = this.normalizeSignature(extracted);
        const normalizedSignature = this.normalizeSignature(signature);
        console.log({
          event: "restoreSignatureAsync",
          normalizedExtracted,
          normalizedSignature,
        });
        return success || (extracted && normalizedExtracted === normalizedSignature);
      }
    }

    console.error({ event: "restoreSignatureAsync", error: "Failed to refresh body" });
    return false;
  },

  /**
   * Starts real-time monitoring of signature modifications.
   * @param {Office.MessageCompose} item - The email item.
   * @param {string} originalSignature - The original signature to monitor.
   */
  startSignatureMonitoring(item, originalSignature) {
    console.log({ event: "startSignatureMonitoring", originalSignatureLength: originalSignature?.length });

    if (isMonitoringActive) {
      this.stopSignatureMonitoring();
    }

    currentItem = item;
    lastKnownSignature = originalSignature;
    isMonitoringActive = true;

    // Monitor every 2 seconds for signature changes
    contentMonitoringInterval = setInterval(async () => {
      try {
        await this.checkSignatureModification();
      } catch (error) {
        console.error({ event: "signatureMonitoring", error: error.message });
      }
    }, 2000);

    console.log({ event: "startSignatureMonitoring", status: "Monitoring started" });
  },

  /**
   * Stops real-time signature monitoring.
   */
  stopSignatureMonitoring() {
    if (contentMonitoringInterval) {
      clearInterval(contentMonitoringInterval);
      contentMonitoringInterval = null;
    }
    isMonitoringActive = false;
    currentItem = null;
    lastKnownSignature = null;
    console.log({ event: "stopSignatureMonitoring", status: "Monitoring stopped" });
  },

  /**
   * Checks for signature modifications and triggers Smart Alert if needed.
   */
  async checkSignatureModification() {
    if (!currentItem || !lastKnownSignature || !isMonitoringActive) {
      return;
    }

    try {
      const body = await new Promise((resolve) =>
        currentItem.body.getAsync("html", (result) =>
          resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : null)
        )
      );

      if (!body) return;

      const isClassicOutlook = Office.context.mailbox.diagnostics.hostName === "Outlook";
      const currentSignature = isClassicOutlook
        ? this.extractSignatureForOutlookClassic(body)
        : this.extractSignature(body);

      if (!currentSignature) {
        console.log({ event: "checkSignatureModification", status: "No signature found in current body" });
        return;
      }

      const normalizedCurrent = this.normalizeSignature(currentSignature);
      const normalizedOriginal = this.normalizeSignature(lastKnownSignature);

      if (normalizedCurrent !== normalizedOriginal) {
        console.log({
          event: "checkSignatureModification",
          status: "Signature modification detected",
          originalLength: normalizedOriginal.length,
          currentLength: normalizedCurrent.length
        });

        await this.showSmartAlert();
        this.stopSignatureMonitoring(); // Stop monitoring after first detection
      }
    } catch (error) {
      console.error({ event: "checkSignatureModification", error: error.message });
    }
  },

  /**
   * Shows a Smart Alert for signature modification.
   */
  async showSmartAlert() {
    try {
      console.log({ event: "showSmartAlert", status: "Displaying Smart Alert for signature modification" });

      // Use Office.js notification system for Smart Alert
      const notification = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "⚠️ M3 Signature Modified: The signature has been changed. Please restore the original signature before sending.",
        icon: "none",
        persistent: true
      };

      const notificationId = `smartAlert_${Date.now()}`;

      currentItem.notificationMessages.addAsync(notificationId, notification, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log({ event: "showSmartAlert", status: "Smart Alert displayed successfully" });

          // Auto-remove the alert after 10 seconds
          setTimeout(() => {
            currentItem.notificationMessages.removeAsync(notificationId, (removeResult) => {
              console.log({ event: "showSmartAlert", status: "Smart Alert auto-removed", success: removeResult.status === Office.AsyncResultStatus.Succeeded });
            });
          }, 10000);
        } else {
          console.error({ event: "showSmartAlert", error: result.error?.message || "Failed to show Smart Alert" });
        }
      });

      // Also try to show a more prominent dialog if available
      if (Office.context.ui && Office.context.ui.displayDialogAsync) {
        this.showSignatureModificationDialog();
      }

    } catch (error) {
      console.error({ event: "showSmartAlert", error: error.message });
    }
  },

  /**
   * Shows a dialog for signature modification warning.
   */
  showSignatureModificationDialog() {
    try {
      const dialogUrl = `${window.location.origin}/taskpane.html?mode=signatureAlert`;

      Office.context.ui.displayDialogAsync(dialogUrl, {
        height: 30,
        width: 50,
        displayInIframe: true
      }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = result.value;
          console.log({ event: "showSignatureModificationDialog", status: "Dialog opened" });

          // Auto-close after 5 seconds
          setTimeout(() => {
            dialog.close();
          }, 5000);
        } else {
          console.log({ event: "showSignatureModificationDialog", error: "Failed to open dialog" });
        }
      });
    } catch (error) {
      console.log({ event: "showSignatureModificationDialog", error: error.message });
    }
  },
};

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
      console.error({ event: "displayNotification", error: "No mailbox item", message });
      return;
    }

    const notificationType =
      type === "Error"
        ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
        : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;

    const notification = {
      type: notificationType,
      message: message,
    };
    if (type === "Info") {
      notification.icon = "none";
      notification.persistent = false;
    }

    console.log({ event: "displayNotification", type, message, persistent });
    item.notificationMessages.addAsync(`notif_${new Date().getTime()}`, notification, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error({
          event: "displayNotification",
          error: result.error.message,
          notification,
          host: Office.context.mailbox.diagnostics.hostName,
        });
      }
    });
  } catch (error) {
    console.error({ event: "displayNotification", error: error.message, message });
  }
}

/**
 * Displays an error with a Smart Alert and notification.
 * @param {string} message - Error message.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 * @param {boolean} restoreSignature - Whether to restore the original signature.
 * @param {string} signatureKey - The signature key to restore.
 * @param {string} tempSignature - Temporary signature for new emails (optional).
 */
async function displayError(message, event, restoreSignature = false, signatureKey = null, tempSignature = null) {
  console.log({
    event: "displayError",
    message,
    restoreSignature,
    signatureKey,
    tempSignatureLength: tempSignature?.length,
  });

  const markdownMessage = message.includes("modified")
    ? `${message}\n\n**Tip**: Ensure the M3 signature is not edited before sending.`
    : `${message}\n\n**Tip**: Select an M3 signature from the ribbon under "M3 Signatures".`;

  const item = Office.context.mailbox.item;
  if (!item) {
    console.error({ event: "displayError", error: "No mailbox item" });
    displayNotification("Error", message, true);
    event.completed({
      allowEvent: false,
      errorMessage: message,
      errorMessageMarkdown: markdownMessage,
      cancelLabel: "OK",
    });
    return;
  }

  if (restoreSignature) {
    let signatureToRestore = tempSignature || localStorage.getItem("tempSignature_new");
    if (signatureKey && !signatureToRestore) {
      signatureToRestore = localStorage.getItem(`signature_${signatureKey}`);
    }

    if (!signatureToRestore) {
      console.error({ event: "displayError", error: "No signature to restore", signatureKey });
      displayNotification("Error", `${message} (Failed to restore: No signature available)`, true);
      event.completed({
        allowEvent: false,
        errorMessage: `${message} (Failed to restore: No signature available)`,
        errorMessageMarkdown: `${markdownMessage}\n**Note**: Failed to restore signature. Please reselect.`,
        cancelLabel: "OK",
      });
      return;
    }

    const restored = await SignatureManager.restoreSignature(item, signatureToRestore, signatureKey || "tempSignature");
    if (!restored) {
      console.error({ event: "displayError", error: "Restoration failed", signatureKey });
      displayNotification("Error", `${message} (Failed to restore signature)`, true);
      event.completed({
        allowEvent: false,
        errorMessage: `${message} (Failed to restore signature)`,
        errorMessageMarkdown: `${markdownMessage}\n**Note**: Failed to restore signature. Please reselect.`,
        cancelLabel: "OK",
      });
      return;
    }

    displayNotification("Error", `${message}`, true);
    event.completed({
      allowEvent: false,
      errorMessage: `${message}`,
      errorMessageMarkdown: `${markdownMessage}`,
      cancelLabel: "OK",
    });
  } else {
    displayNotification("Error", message, false);
    event.completed({
      allowEvent: false,
      errorMessage: message,
      errorMessageMarkdown: markdownMessage,
      cancelLabel: "OK",
    });
  }
}

/**
 * Applies the default M3 signature and allows sending for restored signatures.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
async function applyDefaultSignature(event) {
  console.log({ event: "applyDefaultSignature" });
  const item = Office.context.mailbox.item;
  const body = await new Promise((resolve) => item.body.getAsync("html", (result) => resolve(result.value)));
  const isClassicOutlook = Office.context.mailbox.diagnostics.hostName === "Outlook";
  const currentSignature = isClassicOutlook
    ? SignatureManager.extractSignatureForOutlookClassic(body)
    : SignatureManager.extractSignature(body);

  const signatureKey = await getSignatureKeyForRecipients(item);
  if (!signatureKey) {
    console.log({ event: "applyDefaultSignature", status: "No signature key found, applying default" });
    await addSignature("m3Signature", event);
    return;
  }

  const cachedSignature = localStorage.getItem(`signature_${signatureKey}`);
  if (!cachedSignature) {
    console.log({ event: "applyDefaultSignature", status: "No cached signature, applying default" });
    await addSignature("m3Signature", event);
    return;
  }

  const cleanCurrentSignature = SignatureManager.normalizeSignature(currentSignature);
  const cleanStoredSignature = SignatureManager.normalizeSignature(cachedSignature);

  if (cleanCurrentSignature === cleanStoredSignature) {
    console.log({ event: "applyDefaultSignature", status: "Signature matches, allowing send", signatureKey });
    event.completed({ allowEvent: true });
  } else {
    console.log({ event: "applyDefaultSignature", status: "Signature mismatch, applying default" });
    await addSignature("m3Signature", event);
  }
}

/**
 * Cancels the Smart Alert action.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
function cancelAction(event) {
  console.log({ event: "cancelAction" });
  event.completed({ allowEvent: false });
}

/**
 * Validates the email signature on send.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
async function validateSignature(event) {
  console.log({ event: "validateSignature" });

  // Stop real-time monitoring when validating for send
  SignatureManager.stopSignatureMonitoring();

  try {
    const item = Office.context.mailbox.item;
    if (!item) {
      console.error({ event: "validateSignature", error: "No mailbox item" });
      displayError("No mailbox item available.", event);
      return;
    }

    const isExternal = await isExternalEmail(item);
    console.log({ event: "validateSignature", isExternal });
    const isReplyOrForward = await SignatureManager.isReplyOrForward(item);
    console.log({ event: "validateSignature", isReplyOrForward });
    const body = await new Promise((resolve) => item.body.getAsync("html", (result) => resolve(result.value)));
    const isClassicOutlook = Office.context.mailbox.diagnostics.hostName === "Outlook";
    const currentSignature = isClassicOutlook
      ? SignatureManager.extractSignatureForOutlookClassic(body)
      : SignatureManager.extractSignature(body);

    if (!currentSignature) {
      console.log({ event: "validateSignature", status: "No signature found" });
      displayError("Email is missing the M3 required signature. Please select an appropriate email signature.", event);
    } else {
      await validateSignatureChanges(item, currentSignature, event, isReplyOrForward);
    }
  } catch (error) {
    console.error({ event: "validateSignature", error: error.message });
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
    const newBody = await new Promise((resolve) => item.body.getAsync("html", (result) => resolve(result.value)));
    const isClassicOutlook = Office.context.mailbox.diagnostics.hostName === "Outlook";
    const newSignature = isClassicOutlook
      ? SignatureManager.extractSignatureForOutlookClassic(newBody)
      : SignatureManager.extractSignature(newBody);

    if (!newSignature) {
      console.log({ event: "validateSignatureChanges", status: "Missing signature" });
      displayError("Email is missing the M3 required signature. Please select an appropriate email signature.", event);
      return;
    }

    const cleanNewSignature = SignatureManager.normalizeSignature(newSignature);
    const signatureKeys = ["monaSignature", "morganSignature", "morvenSignature", "m2Signature", "m3Signature"];
    let matchedSignatureKey = null;
    let rawMatchedSignature = null;

    console.log({ event: "validateSignatureChanges", rawNewSignature: newSignature, cleanNewSignature });

    // Check if the current signature matches any stored signature (normalized)
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
      localStorage.getItem("tempSignature_new") || localStorage.getItem(`signature_${signatureKeys[0]}`);
    const cleanLastAppliedSignature = SignatureManager.normalizeSignature(lastAppliedSignature);
    console.log({
      event: "validateSignatureChanges",
      rawLastAppliedSignature: lastAppliedSignature,
      cleanLastAppliedSignature,
    });

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
    console.log({ event: "validateSignatureChanges", expectedLogoUrl });

    // Check if signatures and logos match
    const isTextValid = matchedSignatureKey || cleanNewSignature === cleanLastAppliedSignature;
    const isLogoValid = !expectedLogoUrl || (newLogoUrl && newLogoUrl === expectedLogoUrl);
    console.log({ event: "validateSignatureChanges", isTextValid, isLogoValid });

    if (isTextValid && isLogoValid) {
      console.log({ event: "validateSignatureChanges", status: "Signature and logo valid", matchedSignatureKey });
      if (!isReplyOrForward) {
        localStorage.removeItem("tempSignature_new");
        console.log({ event: "validateSignatureChanges", status: "Cleared temporary signature for new email" });
      }
      await saveSignatureData(item, matchedSignatureKey || signatureKeys[0]);
      event.completed({ allowEvent: true });
    } else {
      console.log({ event: "validateSignatureChanges", status: "Signature or logo modified", matchedSignatureKey });
      if (isReplyOrForward) {
        const signatureKey = await getSignatureKeyForRecipients(item);
        const tempSignature = localStorage.getItem("tempSignature_new");
        if (tempSignature) {
          console.log({ event: "validateSignatureChanges", status: "Restoring temporary signature for reply/forward" });
          displayError(
            "Selected M3 email signature has been modified. M3 email signature is prohibited from modification. The original signature is now restored.",
            event,
            true,
            signatureKey,
            tempSignature
          );
        } else if (signatureKey) {
          console.log({ event: "validateSignatureChanges", status: "Restoring signature from signatureKey" });
          displayError(
            "Selected M3 email signature has been modified. M3 email signature is prohibited from modification. The original signature is now restored.",
            event,
            true,
            signatureKey
          );
        } else {
          console.log({
            event: "validateSignatureChanges",
            status: "No signatureKey or tempSignature for reply/forward, prompting re-selection",
          });
          displayError(
            "Selected M3 signature or logo has been modified. Please select an appropriate email signature.",
            event,
            false
          );
        }
      } else {
        const tempSignature = localStorage.getItem("tempSignature_new");
        if (tempSignature) {
          console.log({ event: "validateSignatureChanges", status: "Restoring temporary signature for new email" });
          displayError(
            "Selected M3 email signature has been modified. M3 email signature is prohibited from modification. The original signature is now restored.",
            event,
            true,
            null,
            tempSignature
          );
        } else {
          console.log({ event: "validateSignatureChanges", status: "Restoring default signature for new email" });
          displayError(
            "Selected M3 email signature has been modified. M3 email signature is prohibited from modification. The original signature is now restored.",
            event,
            true,
            null,
            localStorage.getItem(`signature_${signatureKeys[0]}`)
          );
        }
      }
    }
  } catch (error) {
    console.error({ event: "validateSignatureChanges", error: error.message });
    displayError("Unexpected error validating signature changes.", event);
  }
}

/**
 * Checks if the email is external.
 * @param {Office.MessageCompose} item - The email item.
 * @returns {Promise<boolean>} True if external.
 */
function isExternalEmail(item) {
  return new Promise((resolve) => {
    console.log({ event: "isExternalEmail" });
    const isClassicOutlook = Office.context.mailbox.diagnostics.hostName === "Outlook";
    resolve(!isClassicOutlook && item.inReplyTo && item.inReplyTo.indexOf("OUTLOOK.COM") === -1);
  });
}

/**
 * Fetches a signature from the API.
 * @param {string} signatureKey - The signature key (e.g., "m3Signature").
 * @param {function} callback - Callback with (template, error).
 */
function fetchSignature(signatureKey, callback) {
  const signatureIndex = ["monaSignature", "morganSignature", "morvenSignature", "m2Signature", "m3Signature"].indexOf(
    signatureKey
  );
  const initialUrl = "https://m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net/api/Ribbons/ribbons";
  let signatureUrl =
    "https://m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net/api/Signatures/signatures?signatureURL=";

  fetch(initialUrl)
    .then((response) => response.json())
    .then((data) => {
      signatureUrl += data.result[signatureIndex].url;
      fetch(signatureUrl)
        .then((response) => response.json())
        .then((data) => {
          let template = data.result;
          template = template
            .replace("{First name} ", Office.context.mailbox.userProfile.displayName || "")
            .replace("{Last name}", "")
            .replaceAll("{E-mail}", Office.context.mailbox.userProfile.emailAddress || "")
            .replace("{Title}", "")
            .trim();
          callback(template, null);
        })
        .catch((err) => callback(null, err));
    })
    .catch((err) => callback(null, err));
}

/**
 * Finds the signature key by matching conversationId, recipient emails, and subject in localStorage.
 * @param {Office.MessageCompose} item - The email item.
 * @returns {Promise<string|null>} The signature key or null if no match or signature is "none".
 */
function getSignatureKeyForRecipients(item) {
  return new Promise((resolve) => {
    item.to.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error({ event: "getSignatureKeyForRecipients", error: result.error.message });
        resolve(null);
        return;
      }

      const recipients = result.value.map((recipient) => recipient.emailAddress.toLowerCase());
      const conversationId = item.conversationId || null;

      item.subject.getAsync((subjectResult) => {
        if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error({ event: "getSignatureKeyForRecipients", error: subjectResult.error.message });
          resolve(null);
          return;
        }

        const currentSubject = SignatureManager.normalizeSubject(subjectResult.value);
        console.log({ event: "getSignatureKeyForRecipients", recipients, conversationId, currentSubject });

        const signatureDataEntries = [];
        for (let i = 0; i < localStorage.length; i++) {
          const key = localStorage.key(i);
          if (key.startsWith("signatureData_")) {
            try {
              const data = JSON.parse(localStorage.getItem(key));
              signatureDataEntries.push({ key, data });
            } catch (error) {
              console.error({ event: "getSignatureKeyForRecipients", error: error.message, key });
            }
          }
        }
        console.log({
          event: "getSignatureKeyForRecipients",
          signatureDataEntries: signatureDataEntries.map((entry) => ({
            key: entry.key,
            conversationId: entry.data.conversationId,
            recipients: entry.data.recipients,
            subject: entry.data.subject,
            signature: entry.data.signature,
            timestamp: entry.data.timestamp,
          })),
        });

        signatureDataEntries.sort((a, b) => new Date(b.data.timestamp) - new Date(a.data.timestamp));

        let signatureKey = null;

        if (conversationId) {
          for (const entry of signatureDataEntries) {
            const data = entry.data;
            if (data.conversationId === conversationId && data.signature !== "none") {
              signatureKey = data.signature;
              console.log({
                event: "getSignatureKeyForRecipients",
                status: "Found matching signature by conversationId",
                signatureKey,
                key: entry.key,
                storedSubject: data.subject,
                storedRecipients: data.recipients,
              });
              break;
            }
          }
        }

        if (!signatureKey) {
          for (const entry of signatureDataEntries) {
            const data = entry.data;
            const storedRecipients = data.recipients.map((email) => email.toLowerCase());
            const storedSubject = SignatureManager.normalizeSubject(data.subject);
            if (
              recipients.some((recipient) => storedRecipients.includes(recipient)) &&
              storedSubject === currentSubject &&
              data.signature !== "none"
            ) {
              signatureKey = data.signature;
              console.log({
                event: "getSignatureKeyForRecipients",
                status: "Found matching signature by recipients and subject",
                signatureKey,
                key: entry.key,
                storedSubject,
                storedRecipients,
              });
              break;
            }
          }
        }

        if (signatureKey) {
          console.log({ event: "getSignatureKeyForRecipients", selectedSignatureKey: signatureKey });
          resolve(signatureKey);
        } else {
          console.log({
            event: "getSignatureKeyForRecipients",
            selectedSignatureKey: null,
            status: "No valid signature key",
          });
          resolve(null);
        }
      });
    });
  });
}

/**
 * Adds a signature to the email and saves it to localStorage.
 * @param {string} signatureKey - The signature key (e.g., "m3Signature").
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 * @param {boolean} isAutoApplied - Whether the signature is auto-applied.
 */
async function addSignature(signatureKey, event, isAutoApplied = false) {
  console.log({ event: "addSignature", signatureKey, isAutoApplied });

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
              console.error({ event: "addSignature", error: asyncResult.error.message });
              displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
              if (!isAutoApplied) {
                event.completed();
              } else {
                displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
                saveSignatureData(item, "none");
                event.completed();
              }
            } else {
              console.log({ event: "addSignature", status: "Signature applied from cache", signatureKey });
              saveSignatureData(item, signatureKey);

              // Start real-time monitoring for signature modifications
              SignatureManager.startSignatureMonitoring(item, cachedSignature);

              if (!isAutoApplied) {
                localStorage.setItem("tempSignature_new", cachedSignature);
                console.log({ event: "addSignature", status: "Stored temporary signature for new email" });
              }
              item.body.getAsync("html", (result) => {
                console.log({ event: "addSignature", bodyAfterApply: result.value });
              });
              event.completed();
            }
            resolve();
          }
        )
      );
    } else {
      fetchSignature(signatureKey, async (template, error) => {
        if (error) {
          console.error({ event: "addSignature", error: error.message });
          displayNotification("Error", `Failed to fetch ${signatureKey}.`, true);
          if (!isAutoApplied) {
            event.completed();
          } else {
            displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
            saveSignatureData(item, "none");
            event.completed();
          }
          return;
        }

        await new Promise((resolve) =>
          item.body.setSignatureAsync(
            "<!-- signature -->" + template.trim(),
            { coercionType: Office.CoercionType.Html },
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error({ event: "addSignature", error: asyncResult.error.message });
                displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
                if (!isAutoApplied) {
                  event.completed();
                } else {
                  displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
                  saveSignatureData(item, "none");
                  event.completed();
                }
              } else {
                console.log({ event: "addSignature", status: "Signature applied", signatureKey });
                localStorage.setItem(`signature_${signatureKey}`, template);
                saveSignatureData(item, signatureKey);

                // Start real-time monitoring for signature modifications
                SignatureManager.startSignatureMonitoring(item, template);

                if (!isAutoApplied) {
                  localStorage.setItem("tempSignature_new", template);
                  console.log({ event: "addSignature", status: "Stored temporary signature for new email" });
                }
                item.body.getAsync("html", (result) => {
                  console.log({ event: "addSignature", bodyAfterApply: result.value });
                });
                event.completed();
              }
              resolve();
            }
          )
        );
      });
    }
  } catch (error) {
    console.error({ event: "addSignature", error: error.message });
    displayNotification("Error", `Failed to apply ${signatureKey}.`, true);
    if (!isAutoApplied) {
      event.completed();
    } else {
      displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
      saveSignatureData(item, "none");
      event.completed();
    }
  }
}

/**
 * Saves signature data to localStorage, including subject.
 * @param {Office.MessageCompose} item - The email item.
 * @param {string} signatureKey - The signature key.
 * @returns {Promise<object|null>} The saved data or null if failed.
 */
function saveSignatureData(item, signatureKey) {
  return new Promise((resolve) => {
    item.to.getAsync((result) => {
      let recipients = [];
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        recipients = result.value.map((recipient) => recipient.emailAddress.toLowerCase());
      } else {
        console.error({ event: "saveSignatureData", error: result.error.message });
      }

      const conversationId = item.conversationId || null;

      item.subject.getAsync((subjectResult) => {
        let subject = "";
        if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
          subject = subjectResult.value;
        } else {
          console.error({ event: "saveSignatureData", error: subjectResult.error.message });
        }

        console.log({ event: "saveSignatureData", signatureKey, recipients, conversationId, subject });

        let existingKey = null;
        if (conversationId) {
          for (let i = 0; i < localStorage.length; i++) {
            const key = localStorage.key(i);
            if (key.startsWith("signatureData_")) {
              try {
                const data = JSON.parse(localStorage.getItem(key));
                if (data.conversationId === conversationId) {
                  existingKey = key;
                  break;
                }
              } catch (error) {
                console.error({ event: "saveSignatureData", error: error.message, key });
              }
            }
          }
        }

        const data = {
          recipients,
          signature: signatureKey,
          conversationId,
          subject,
          timestamp: new Date().toISOString(),
        };

        if (existingKey) {
          localStorage.setItem(existingKey, JSON.stringify(data));
          console.log({ event: "saveSignatureData", status: "Updated existing entry", key: existingKey, subject });
        } else {
          const newKey = `signatureData_${Date.now()}`;
          localStorage.setItem(newKey, JSON.stringify(data));
          console.log({ event: "saveSignatureData", status: "Created new entry", key: newKey, subject });
        }
        resolve(data);
      });
    });
  });
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

/**
 * Enhanced handler for new message compose event with improved inline reply/forward support.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
async function onNewMessageComposeHandler(event) {
  console.log({ event: "onNewMessageComposeHandler" });

  const item = Office.context.mailbox.item;
  const isReplyOrForward = await SignatureManager.isReplyOrForward(item);
  console.log({ event: "onNewMessageComposeHandler", isReplyOrForward });

  if (isReplyOrForward) {
    // Enhanced handling for inline replies/forwards
    await handleReplyForwardWithRetry(item, event);
  } else {
    console.log({ event: "onNewMessageComposeHandler", status: "New email, requiring manual signature selection" });
    displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
    await saveInitialSignatureData(item);
    localStorage.removeItem("tempSignature_new");
    console.log({ event: "onNewMessageComposeHandler", status: "Cleared temporary signature for new email" });
    event.completed();
  }
}

/**
 * Handles reply/forward with retry mechanism for inline scenarios.
 * @param {Office.MessageCompose} item - The email item.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 */
async function handleReplyForwardWithRetry(item, event) {
  console.log({ event: "handleReplyForwardWithRetry", status: "Starting enhanced reply/forward handling" });

  let signatureKey = null;
  let retryCount = 0;
  const maxRetries = 3;
  const retryDelay = 1000; // 1 second

  // Retry mechanism for inline replies where context might not be immediately available
  while (retryCount < maxRetries && !signatureKey) {
    signatureKey = await getSignatureKeyForRecipients(item);

    if (!signatureKey && retryCount < maxRetries - 1) {
      console.log({
        event: "handleReplyForwardWithRetry",
        status: "No signature found, retrying",
        attempt: retryCount + 1,
        maxRetries
      });

      // Wait before retry
      await new Promise(resolve => setTimeout(resolve, retryDelay));
      retryCount++;
    } else {
      break;
    }
  }

  if (signatureKey) {
    console.log({
      event: "handleReplyForwardWithRetry",
      status: "Auto-applying signature for reply/forward",
      signatureKey,
      attempts: retryCount + 1
    });

    // Apply signature with enhanced inline support
    await addSignatureWithInlineSupport(signatureKey, event, true, item);
  } else {
    console.log({
      event: "handleReplyForwardWithRetry",
      status: "No valid signature found after retries, requiring manual selection",
      attempts: retryCount + 1
    });
    displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
    await saveInitialSignatureData(item);
    event.completed();
  }
}

/**
 * Enhanced signature application with inline reply/forward support.
 * @param {string} signatureKey - The signature key.
 * @param {Office.AddinCommands.Event} event - The Outlook event object.
 * @param {boolean} isAutoApplied - Whether the signature is auto-applied.
 * @param {Office.MessageCompose} item - The email item.
 */
async function addSignatureWithInlineSupport(signatureKey, event, isAutoApplied, item) {
  console.log({ event: "addSignatureWithInlineSupport", signatureKey, isAutoApplied });

  try {
    const cachedSignature = localStorage.getItem(`signature_${signatureKey}`);
    let signatureToApply = cachedSignature;

    // If no cached signature, fetch it
    if (!signatureToApply) {
      signatureToApply = await new Promise((resolve, reject) => {
        fetchSignature(signatureKey, (template, error) => {
          if (error) {
            reject(error);
          } else {
            localStorage.setItem(`signature_${signatureKey}`, template);
            resolve(template);
          }
        });
      });
    }

    if (!signatureToApply) {
      throw new Error("No signature template available");
    }

    // Enhanced signature application with retry for inline scenarios
    const success = await applySignatureWithRetry(item, signatureToApply, signatureKey);

    if (success) {
      console.log({ event: "addSignatureWithInlineSupport", status: "Signature applied successfully", signatureKey });
      await saveSignatureData(item, signatureKey);

      // Start real-time monitoring for signature modifications
      SignatureManager.startSignatureMonitoring(item, signatureToApply);

      if (!isAutoApplied) {
        localStorage.setItem("tempSignature_new", signatureToApply);
        console.log({ event: "addSignatureWithInlineSupport", status: "Stored temporary signature for new email" });
      }

      event.completed();
    } else {
      throw new Error("Failed to apply signature after retries");
    }

  } catch (error) {
    console.error({ event: "addSignatureWithInlineSupport", error: error.message });
    displayNotification("Error", `Failed to apply ${signatureKey}.`, true);

    if (!isAutoApplied) {
      event.completed();
    } else {
      displayNotification("Info", "Please select an M3 signature from the ribbon.", false);
      await saveSignatureData(item, "none");
      event.completed();
    }
  }
}

/**
 * Applies signature with retry mechanism for inline scenarios.
 * @param {Office.MessageCompose} item - The email item.
 * @param {string} signature - The signature to apply.
 * @param {string} signatureKey - The signature key for logging.
 * @returns {Promise<boolean>} True if successful.
 */
async function applySignatureWithRetry(item, signature, signatureKey) {
  const maxRetries = 3;
  const retryDelay = 500; // 500ms

  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      const success = await new Promise((resolve) => {
        item.body.setSignatureAsync(
          "<!-- signature -->" + signature.trim(),
          { coercionType: Office.CoercionType.Html },
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              console.log({
                event: "applySignatureWithRetry",
                attempt: attempt + 1,
                error: asyncResult.error.message,
                signatureKey
              });
              resolve(false);
            } else {
              console.log({
                event: "applySignatureWithRetry",
                attempt: attempt + 1,
                status: "Success",
                signatureKey
              });
              resolve(true);
            }
          }
        );
      });

      if (success) {
        return true;
      }

      // Wait before retry
      if (attempt < maxRetries - 1) {
        await new Promise(resolve => setTimeout(resolve, retryDelay));
      }

    } catch (error) {
      console.error({
        event: "applySignatureWithRetry",
        attempt: attempt + 1,
        error: error.message,
        signatureKey
      });

      if (attempt < maxRetries - 1) {
        await new Promise(resolve => setTimeout(resolve, retryDelay));
      }
    }
  }

  console.error({
    event: "applySignatureWithRetry",
    status: "Failed after all retries",
    maxRetries,
    signatureKey
  });
  return false;
}

/**
 * Saves initial signature data with "none" for new or reply/forward emails.
 * @param {Office.MessageCompose} item - The email item.
 */
function saveInitialSignatureData(item) {
  item.to.getAsync((result) => {
    let recipients = [];
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      recipients = result.value.map((recipient) => recipient.emailAddress.toLowerCase());
    } else {
      console.error({ event: "saveInitialSignatureData", error: result.error.message });
    }

    const conversationId = item.conversationId || null;

    item.subject.getAsync((subjectResult) => {
      let subject = "";
      if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
        subject = subjectResult.value;
      } else {
        console.error({ event: "saveInitialSignatureData", error: subjectResult.error.message });
      }

      const data = {
        recipients,
        signature: "none",
        conversationId,
        subject,
        timestamp: new Date().toISOString(),
      };

      const newKey = `signatureData_${Date.now()}`;
      localStorage.setItem(newKey, JSON.stringify(data));
      console.log({
        event: "saveInitialSignatureData",
        status: "Stored initial signature data",
        recipients,
        conversationId,
        subject,
      });
    });
  });
}
