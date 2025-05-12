// Detect Android
var isAndroid = navigator.userAgent.toLowerCase().indexOf("android") > -1;
var initialEnvironment = isAndroid ? "mobile" : "desktop";

// Initialize Sentry
// Sentry.onLoad(function () {
//   Sentry.init({
//     dsn: "https://9cb6398daefb0df54d63e4da9ff3e7a3@o4509305237864448.ingest.us.sentry.io/4509305244680192",
//     tracesSampleRate: 1.0,
//     environment: initialEnvironment,
//     release: "m3-signatures@1.0.0.13",
//   });
//   Sentry.configureScope(function (scope) {
//     scope.setTag("context", "taskpane");
//     scope.setTag("userAgent", navigator.userAgent);
//   });
//   Sentry.captureMessage("Task pane initialized", "info");
//   console.log({ event: "taskPaneInitialized", environment: initialEnvironment });
// });

// Detect commands.js
if (typeof Office !== "undefined" && typeof Office.actions !== "undefined") {
  // Sentry.captureMessage("commands.js detected in taskpane", "warning");
  console.warn({ event: "commandsJsDetected", status: "Unexpected in taskpane" });
}

// Initialize task pane
function initializeTaskPane() {
  // Load saved settings
  var defaultSignature = localStorage.getItem("defaultSignature");
  if (defaultSignature) {
    var selectedRadio = document.querySelector('input[value="' + defaultSignature + '"]');
    if (selectedRadio) {
      selectedRadio.checked = true;
      // Sentry.captureMessage("Loaded default signature: " + defaultSignature, "info");
      console.log({ event: "loadDefaultSignature", signatureKey: defaultSignature });
    }
  }

  // Choice field click handlers
  document.querySelectorAll(".choice-field").forEach(function (field) {
    field.addEventListener("click", function (e) {
      var radio = field.querySelector('input[type="radio"]');
      if (radio && e.target !== radio) {
        radio.checked = true;
        // Sentry.captureMessage("Selected signature: " + radio.value, "info");
        console.log({ event: "selectSignature", signatureKey: radio.value });
      }
    });
  });

  // Save settings handler
  document.getElementById("saveButton").addEventListener("click", function () {
    var selectedRadio = document.querySelector('input[name="signatureOption"]:checked');
    if (selectedRadio) {
      var signatureKey = selectedRadio.value;
      localStorage.setItem("defaultSignature", signatureKey);
      // Sentry.captureMessage("Saved default signature: " + signatureKey, "info");
      console.log({ event: "saveDefaultSignature", signatureKey: signatureKey });
      alert("Signature saved: " + signatureKey);
    } else {
      Sentry.captureMessage("No signature selected", "warning");
      console.log({ event: "saveDefaultSignatureError", message: "No signature selected" });
      alert("Please select a signature option");
    }
  });

  // Test error button
  var testButton = document.getElementById("test-error");
  if (testButton) {
    testButton.addEventListener("click", function () {
      // Sentry.captureMessage("Test error button clicked", "info");
      console.log({ event: "testErrorButton", status: "Clicked" });
      throw new Error("This is a test error");
    });
  }
}

// Initialize immediately
initializeTaskPane();
alert("a");
