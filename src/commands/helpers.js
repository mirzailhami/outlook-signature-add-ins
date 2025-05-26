import { DateTime } from "luxon";

/**
 * Centralized logger for structured logging with timestamps.
 * @type {Object}
 */
const logger = {
  /**
   * Logs a message with a timestamp and structured details.
   * @param {string} level - Log level ("info" or "error").
   * @param {string} event - Event name or identifier.
   * @param {Object} [details={}] - Additional details to log.
   */
  log(level, event, details = {}) {
    const timestamp = DateTime.now().toISO({ suppressMilliseconds: true });
    const logEntry = {
      timestamp,
      level,
      event,
      ...details,
    };
    if (level === "error") {
      console.error(JSON.stringify(logEntry, null, 2));
    } else {
      console.log(JSON.stringify(logEntry, null, 2));
    }
  },
};

/**
 * Manages email signature extraction, normalization, and restoration.
 * @type {Object}
 */
const SignatureManager = {
  /**
   * Extracts a signature from an email body using markers or regex patterns.
   * @param {string|null} body - The email body HTML content.
   * @returns {string|null} The extracted signature, or null if not found.
   */
  extractSignature(body) {
    logger.log("info", "extractSignature", { bodyLength: body?.length });
    if (!body) return null;

    const marker = "<!-- signature -->";
    const startIndex = body.indexOf(marker);
    if (startIndex !== -1) {
      const endIndex = body.indexOf("</body>", startIndex);
      const signature = body.slice(startIndex + marker.length, endIndex !== -1 ? endIndex : undefined).trim();
      logger.log("info", "extractSignature", { method: "marker", signatureLength: signature.length });
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
        logger.log("info", "extractSignature", { method: regex.source, signatureLength: signature.length });
        return signature;
      }
    }

    logger.log("info", "extractSignature", { status: "No signature found" });
    return null;
  },

  /**
   * Extracts a signature from an email body in Outlook Classic using specific markers or regex.
   * @param {string|null} body - The email body HTML content.
   * @returns {string|null} The extracted signature, or null if not found.
   */
  extractSignatureForOutlookClassic(body) {
    if (!body) return null;

    const marker = "<!-- signature -->";
    const startIndex = body.indexOf(marker);
    if (startIndex !== -1) {
      const endIndex = body.indexOf("</body>", startIndex);
      const signature = body.slice(startIndex + marker.length, endIndex !== -1 ? endIndex : undefined).trim();
      logger.log("info", "extractSignatureForOutlookClassic", { method: "marker", signatureLength: signature.length });
      return signature;
    }

    const regex = /<table\s+class=MsoNormalTable[^>]*>([\s\S]*?)$/is;
    const match = body.match(regex);
    if (match) {
      const signature = match[0].trim();
      logger.log("info", "extractSignatureForOutlookClassic", { method: "table", signatureLength: signature.length });
      return signature;
    }

    logger.log("info", "extractSignatureForOutlookClassic", { status: "No signature found" });
    return null;
  },

  /**
   * Normalizes a signature by removing HTML tags and standardizing text.
   * @param {string|null} sig - The signature to normalize.
   * @returns {string} The normalized signature, or an empty string if null.
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
      .replace(/[\r\n]+/g, " ")
      .replace(/\s*([.,:;])\s*/g, "$1")
      .replace(/\s+/g, " ")
      .replace(/\s*:\s*/g, ":")
      .replace(/\s+(email:)/gi, "$1")
      .trim()
      .toLowerCase();
    logger.log("info", "normalizeSignature", { rawLength: sig.length, normalizedLength: normalized.length });
    return normalized;
  },

  /**
   * Normalizes an email subject by removing reply/forward prefixes.
   * @param {string|null} subject - The email subject.
   * @returns {string} The normalized subject, or an empty string if null.
   */
  normalizeSubject(subject) {
    if (!subject) return "";
    const normalized = subject
      .replace(/^(re:|fw:|fwd:)\s*/i, "")
      .trim()
      .toLowerCase();
    logger.log("info", "normalizeSubject", { rawLength: subject.length, normalizedLength: normalized.length });
    return normalized;
  },

  /**
   * Checks if an email is a reply or forward.
   * @async
   * @param {Office.MessageCompose} item - The email item.
   * @returns {Promise<boolean>} True if the email is a reply or forward, false otherwise.
   */
  async isReplyOrForward(item) {
    if (item.itemType === Office.MailboxEnums.ItemType.Message && item.conversationId) {
      logger.log("info", "checkForReplyOrForward", {
        status: "Reply/forward detected",
        conversationId: item.conversationId,
      });
      return true;
    }
    if (item.inReplyTo) {
      logger.log("info", "checkForReplyOrForward", { status: "Reply detected", inReplyTo: item.inReplyTo });
      return true;
    }
    const subject = await new Promise((resolve) => item.subject.getAsync((result) => resolve(result.value || "")));
    const isReplyOrForward = ["re:", "fw:", "fwd:"].some((prefix) => subject.toLowerCase().includes(prefix));
    logger.log("info", "checkForReplyOrForward", { status: "Subject checked", isReplyOrForward });
    return isReplyOrForward;
  },

  /**
   * Restores a signature to the email body.
   * @async
   * @param {Office.MessageCompose} item - The email item.
   * @param {string} signature - The signature to restore.
   * @param {string} signatureKey - The signature key for fallback.
   * @returns {Promise<boolean>} True if the signature was restored successfully, false otherwise.
   */
  async restoreSignature(item, signature, signatureKey) {
    logger.log("info", "restoreSignatureAsync", { signatureKey, cachedSignatureLength: signature?.length });
    if (!signature) {
      signature = localStorage.getItem(`signature_${signatureKey}`);
      logger.log("info", "restoreSignatureAsync", {
        status: "Falling back to signatureKey",
        fallbackLength: signature?.length,
      });
      if (!signature) return false;
    }

    const success = await new Promise((resolve) =>
      item.body.setSignatureAsync(
        "<!-- signature -->" + signature.trim(),
        { coercionType: Office.CoercionType.Html },
        (asyncResult) => (asyncResult.status === Office.AsyncResultStatus.Failed ? resolve(false) : resolve(true))
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
        logger.log("info", "restoreSignatureAsync", {
          status: "Body refreshed",
          extractedSignatureLength: extracted?.length,
        });
        const normalizedExtracted = this.normalizeSignature(extracted);
        const normalizedSignature = this.normalizeSignature(signature);
        return success || (extracted && normalizedExtracted === normalizedSignature);
      }
    }

    logger.log("error", "restoreSignatureAsync", { error: "Failed to refresh body" });
    return false;
  },
};

/**
 * Fetches a signature template from the API and applies user-specific replacements.
 * @async
 * @param {string} signatureKey - The signature key (e.g., "m3Signature").
 * @param {function(string|null, Error|null): void} callback - Callback with the template or error.
 */
async function fetchSignature(signatureKey, callback) {
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
          let template = data.result
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
 * @async
 * @param {Office.MessageCompose} item - The email item.
 * @returns {Promise<string|null>} The signature key, or null if no match or signature is "none".
 */
async function getSignatureKeyForRecipients(item) {
  return new Promise((resolve) => {
    item.to.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        logger.log("error", "getSignatureKeyForRecipients", { error: result.error.message });
        resolve(null);
        return;
      }

      const recipients = result.value.map((recipient) => recipient.emailAddress.toLowerCase());
      const conversationId = item.conversationId || null;

      item.subject.getAsync((subjectResult) => {
        if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
          logger.log("error", "getSignatureKeyForRecipients", { error: subjectResult.error.message });
          resolve(null);
          return;
        }

        const currentSubject = SignatureManager.normalizeSubject(subjectResult.value);
        logger.log("info", "getSignatureKeyForRecipients", { recipients, conversationId, currentSubject });

        const signatureDataEntries = [];
        for (let i = 0; i < localStorage.length; i++) {
          const key = localStorage.key(i);
          if (key.startsWith("signatureData_")) {
            try {
              const data = JSON.parse(localStorage.getItem(key));
              signatureDataEntries.push({ key, data });
            } catch (error) {
              logger.log("error", "getSignatureKeyForRecipients", { error: error.message, key });
            }
          }
        }

        signatureDataEntries.sort((a, b) => new Date(b.data.timestamp) - new Date(a.data.timestamp));
        let signatureKey = null;

        if (conversationId) {
          for (const entry of signatureDataEntries) {
            if (entry.data.conversationId === conversationId && entry.data.signature !== "none") {
              signatureKey = entry.data.signature;
              break;
            }
          }
        }

        if (!signatureKey) {
          for (const entry of signatureDataEntries) {
            const storedRecipients = entry.data.recipients.map((email) => email.toLowerCase());
            const storedSubject = SignatureManager.normalizeSubject(entry.data.subject);
            if (
              recipients.some((recipient) => storedRecipients.includes(recipient)) &&
              storedSubject === currentSubject &&
              entry.data.signature !== "none"
            ) {
              signatureKey = entry.data.signature;
              break;
            }
          }
        }

        resolve(signatureKey);
      });
    });
  });
}

/**
 * Saves signature data to localStorage, including recipients, signature, and subject.
 * @async
 * @param {Office.MessageCompose} item - The email item.
 * @param {string} signatureKey - The signature key.
 * @returns {Promise<Object|null>} The saved data, or null if failed.
 */
async function saveSignatureData(item, signatureKey) {
  return new Promise((resolve) => {
    item.to.getAsync((result) => {
      const recipients =
        result.status === Office.AsyncResultStatus.Succeeded
          ? result.value.map((r) => r.emailAddress.toLowerCase())
          : [];
      if (result.status !== Office.AsyncResultStatus.Succeeded)
        logger.log("error", "saveSignatureData", { error: result.error.message });

      const conversationId = item.conversationId || null;

      item.subject.getAsync((subjectResult) => {
        const subject = subjectResult.status === Office.AsyncResultStatus.Succeeded ? subjectResult.value : "";
        if (subjectResult.status !== Office.AsyncResultStatus.Succeeded)
          logger.log("error", "saveSignatureData", { error: subjectResult.error.message });

        logger.log("info", "saveSignatureData", { signatureKey, recipients, conversationId, subject });

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
                logger.log("error", "saveSignatureData", { error: error.message, key });
              }
            }
          }
        }

        const data = {
          recipients,
          signature: signatureKey,
          conversationId,
          subject,
          timestamp: DateTime.now().toISO(),
        };

        if (existingKey) {
          localStorage.setItem(existingKey, JSON.stringify(data));
          logger.log("info", "saveSignatureData", { status: "Updated existing entry", key: existingKey, subject });
        } else {
          const newKey = `signatureData_${DateTime.now().toMillis()}`;
          localStorage.setItem(newKey, JSON.stringify(data));
          logger.log("info", "saveSignatureData", { status: "Created new entry", key: newKey, subject });
        }
        resolve(data);
      });
    });
  });
}

/**
 * Saves initial signature data with "none" for new or reply/forward emails.
 * @async
 * @param {Office.MessageCompose} item - The email item.
 * @returns {Promise<void>}
 */
async function saveInitialSignatureData(item) {
  item.to.getAsync((result) => {
    const recipients =
      result.status === Office.AsyncResultStatus.Succeeded ? result.value.map((r) => r.emailAddress.toLowerCase()) : [];
    if (result.status !== Office.AsyncResultStatus.Succeeded)
      logger.log("error", "saveInitialSignatureData", { error: result.error.message });

    const conversationId = item.conversationId || null;

    item.subject.getAsync((subjectResult) => {
      const subject = subjectResult.status === Office.AsyncResultStatus.Succeeded ? subjectResult.value : "";
      if (subjectResult.status !== Office.AsyncResultStatus.Succeeded)
        logger.log("error", "saveInitialSignatureData", { error: subjectResult.error.message });

      const data = { recipients, signature: "none", conversationId, subject, timestamp: DateTime.now().toISO() };
      const newKey = `signatureData_${DateTime.now().toMillis()}`;
      localStorage.setItem(newKey, JSON.stringify(data));
      logger.log("info", "saveInitialSignatureData", {
        status: "Stored initial signature data",
        recipients,
        conversationId,
        subject,
      });
    });
  });
}

/**
 * Checks if the email is external based on host and reply details.
 * @async
 * @param {Office.MessageCompose} item - The email item.
 * @returns {Promise<boolean>} True if the email is external, false otherwise.
 */
function isExternalEmail(item) {
  return new Promise((resolve) =>
    resolve(
      Office.context.mailbox.diagnostics.hostName !== "Outlook" &&
        item.inReplyTo &&
        item.inReplyTo.indexOf("OUTLOOK.COM") === -1
    )
  );
}

export {
  logger,
  SignatureManager,
  fetchSignature,
  getSignatureKeyForRecipients,
  saveSignatureData,
  saveInitialSignatureData,
  isExternalEmail,
};
