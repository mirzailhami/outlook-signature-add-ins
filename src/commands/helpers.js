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
    if (!body) return null;

    const marker = "<!-- signature -->";
    const startIndex = body.indexOf(marker);
    if (startIndex !== -1) {
      const endIndex = body.indexOf("</body>", startIndex);
      const signature = body.slice(startIndex + marker.length, endIndex !== -1 ? endIndex : undefined).trim();
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
        return signature;
      }
    }

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
    const startIndex = body.lastIndexOf(marker);
    if (startIndex !== -1) {
      const endIndex = body.indexOf("</body>", startIndex);
      const signature = body.slice(startIndex + marker.length, endIndex !== -1 ? endIndex : undefined).trim();
      logger.log("info", "extractSignatureForOutlookClassic", { method: "marker", signatureLength: signature.length });
      return signature;
    }

    const regex =
      /<table\s+class=MsoNormalTable[^>]*>([\s\S]*?)(?=(?:<div\s+id="[^"]*appendonsend"|>?\s*<(?:table|hr)\b)|$)/is;
    const match = body.match(regex);
    if (match) {
      const signature = match[1].trim();
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

    // Decode HTML entities first
    const textarea = document.createElement("textarea");
    textarea.innerHTML = sig;
    let normalized = textarea.value;

    // Replace HTML entities
    const htmlEntities = { " ": " ", "&": "&", "<": "<", ">": ">", '"': '"' };
    for (const [entity, char] of Object.entries(htmlEntities)) {
      normalized = normalized.replace(new RegExp(entity, "gi"), char);
    }
    // Remove HTML tags
    normalized = normalized.replace(/<[^>]+>/g, " ");
    // Clean up text
    normalized = normalized
      .replace(/[\r\n]+/g, " ") // Replace newlines with a single space
      .replace(/\s*([.,:;])\s*/g, "$1") // Remove spaces around punctuation
      .replace(/\s+/g, " ") // Collapse multiple spaces into one
      .replace(/\s*:\s*/g, ":") // Remove spaces around colons
      .replace(/\s+(email:)/gi, "$1") // Remove spaces before "email:"
      .trim() // Remove leading/trailing spaces
      .toLowerCase();
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
    // Check 1: inReplyTo (reliable indicator of a reply)
    if (item.inReplyTo) {
      return true;
    }

    // Check 2: Subject prefix
    const subjectResult = await new Promise((resolve) => item.subject.getAsync((result) => resolve(result)));
    let subject = "";
    if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
      subject = subjectResult.value || "";
    }
    const hasReplyOrForwardPrefix = ["re:", "fw:", "fwd:"].some((prefix) => subject.toLowerCase().includes(prefix));

    // Check 3: conversationId (only if subject indicates reply/forward)
    if (item.itemType === Office.MailboxEnums.ItemType.Message && item.conversationId && hasReplyOrForwardPrefix) {
      return true;
    }

    return false;
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
      if (!signature) {
        logger.log("error", "restoreSignatureAsync", { error: "No signature available" });
        return false;
      }
    }

    const signatureWithMarker = "<!-- signature -->" + signature.trim();
    let success = false;

    const currentBody = await new Promise((resolve) =>
      item.body.getAsync("html", (result) =>
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : null)
      )
    );
    if (!currentBody) {
      logger.log("error", "restoreSignatureAsync", { error: "Failed to get current body" });
      return false;
    }

    const startIndex = currentBody.indexOf("<!-- signature -->");
    if (startIndex === -1) {
      logger.log("warn", "restoreSignatureAsync", { error: "Signature marker not found, appending instead" });
      success = await new Promise((resolve) =>
        item.body.setSignatureAsync(signatureWithMarker, { coercionType: Office.CoercionType.Html }, (asyncResult) =>
          resolve(asyncResult.status !== Office.AsyncResultStatus.Failed)
        )
      );
    } else {
      const endIndex =
        currentBody.indexOf("</body>", startIndex) !== -1
          ? currentBody.indexOf("</body>", startIndex)
          : currentBody.length;
      const newBody = currentBody.substring(0, startIndex) + signatureWithMarker + currentBody.substring(endIndex);
      success = await new Promise((resolve) =>
        item.body.setAsync(newBody, { coercionType: Office.CoercionType.Html }, (asyncResult) =>
          resolve(asyncResult.status !== Office.AsyncResultStatus.Failed)
        )
      );
    }

    return success;
  },
};

/**
 * Detects the signature key based on content keywords and logo URL.
 * @param {string} signatureText - The signature text to analyze.
 * @returns {string|null} The matched signature key, or null if no match.
 */
function detectSignatureKey(signatureText) {
  const signatureKeyMapping = {
    "mona offshore wind limited": "monaSignature",
    "morgan offshore wind limited": "morganSignature",
    "morven offshore wind limited": "morvenSignature",
    "m2 offshore wind limited": "m2Signature",
    "m3 offshore wind limited": "m3Signature",
  };

  // Extract logo URL to determine the signature
  const logoRegex = /<img[^>]+src=["'](.*?(?:m3signatures\/logo\/([^"']+)))["'][^>]*>/i;
  const logoMatch = signatureText.match(logoRegex);
  let logoFile = logoMatch ? logoMatch[2] : null;
  if (logoFile) {
    const logoToKey = {
      "morven_v1.png": "morvenSignature",
      "morgan_v1.png": "morganSignature",
      "mona_v1.png": "monaSignature",
      "m2_v1.png": "m2Signature",
      "m3_v1.png": "m3Signature",
    };
    const keyFromLogo = logoToKey[logoFile];
    if (keyFromLogo) return keyFromLogo; // Prioritize logo if found
  }

  // Fallback to text-based detection, return the first match or null
  for (const [keyword, key] of Object.entries(signatureKeyMapping)) {
    if (signatureText.toLowerCase().includes(keyword)) {
      return key; // Return the first match instead of the last
    }
  }
  return null;
}

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
 * Appends debug logs to the email body using setSignatureAsync for mobile debugging.
 * @param {Office.MessageCompose} item - The Outlook message item.
 * @param {...(string|any)} args - Variable arguments: label-value pairs or messages (e.g., "label", value, "label2", value2).
 */
async function appendDebugLogToBody(item, ...args) {
  const timestamp = new Date().toISOString();
  let logContent = `<div style="font-family: Arial, sans-serif; font-size: 12px; color: #333; margin: 10px 0; border-bottom: 1px solid #ccc; padding-bottom: 5px;">`;
  logContent += `<strong>[${timestamp}]</strong><br>`;

  // Process arguments into key-value pairs or messages
  for (let i = 0; i < args.length; i += 2) {
    const label = args[i];
    const value = i + 1 < args.length ? args[i + 1] : "";
    if (typeof label === "string") {
      logContent += `<span style="color: #0055aa;"><strong>${label}:</strong></span> ${
        value !== undefined && value !== null ? JSON.stringify(value) : "undefined"
      }<br>`;
    } else {
      logContent += `<span>${label}</span><br>`;
    }
  }
  logContent += `</div>`;

  // Append the log to the existing body
  return new Promise((resolve) => {
    item.body.getAsync("html", { asyncContext: logContent }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const currentBody = result.value || "";
        item.body.setSignatureAsync(
          currentBody + logContent,
          { coercionType: Office.CoercionType.Html },
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              // Fallback: Try setting the log alone if append fails
              item.body.setSignatureAsync(logContent, { coercionType: Office.CoercionType.Html }, () => resolve());
            } else {
              resolve();
            }
          }
        );
      } else {
        // Fallback if getAsync fails
        item.body.setSignatureAsync(logContent, { coercionType: Office.CoercionType.Html }, () => resolve());
      }
    });
  });
}

export { logger, SignatureManager, fetchSignature, detectSignatureKey, appendDebugLogToBody };
