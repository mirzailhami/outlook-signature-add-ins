import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
  
 

/* global document, Office, module, require */

const title = "E-mail signatures";

const rootElement = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady((info) => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <App title={title} />
    </FluentProvider>
  );

  
    // Associate the function with the onMessageSend event
    //Office.actions.associate("onMessageSend", onMessageSendHandler);
  
 
    // // Disable the default client signature
    // Office.context.mailbox.item.disableClientSignatureAsync((asyncResult) => {
    //   if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    //     console.log("Default client signature disabled.");

    //     // Add your custom signature
    //     Office.context.mailbox.item.body.setSignatureAsync(
    //     "shrawan1",
         
    //       { coercionType: Office.CoercionType.Html },
    //       (asyncResult) => {
    //         if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    //           console.error("Failed to add custom signature:", asyncResult.error.message);
    //         } else {
    //           console.log("Custom signature added successfully.");
    //         }
    //       }
    //     );
    //   } else {
    //     console.error("Failed to disable client signature:", asyncResult.error.message);
    //   }
    // });
 
});

 

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}


  // Register the ItemSend event handler when the component mounts
  // Office.context.mailbox.item.addHandlerAsync(
  //   Office.EventType.ItemSend,
  //   onMessageSendHandler,
  //   (result) => {
  //     if (result.status === Office.AsyncResultStatus.Succeeded) {
  //       console.log('ItemSend event handler registered successfully.');
  //     } else {
  //       console.error('Failed to register ItemSend event handler:', result.error);
  //     }
  //   });

    

////////////////////
// Handler to validate the email body before sending
// const onMessageSendHandler = async (event) => {
//   // Get the email body
//   Office.context.mailbox.item.body.getAsync(
//     'text',
//     (result) => {
//       if (result.status === Office.AsyncResultStatus.Succeeded) {
//         const emailBody = result.value.trim();

//         // Validation rules
//         if (!emailBody) {
//           // Block the email from being sent
//           event.completed({ allowEvent: false, errorMessage: 'Email body cannot be empty.' });
//           return;
//         }

//         if (emailBody.length > 1000) {
//           // Block the email from being sent
//           event.completed({ allowEvent: false, errorMessage: 'Email body cannot exceed 1000 characters.' });
//           return;
//         }

//         // If validation passes, allow the email to be sent
//         event.completed({ allowEvent: true });
//       } else {
//         console.error('Failed to get email body:', result.error);
//         // Block the email from being sent
//         event.completed({ allowEvent: false, errorMessage: 'Failed to validate email body.' });
//       }
//     }
//   );
// };
