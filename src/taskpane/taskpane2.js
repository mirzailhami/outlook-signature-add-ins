/* global Office console */
import * as React from "react";
 

export async function insertText(template) {

  try {
    Office.actions.associate("validateSignature", function (event) {
      event.completed({ allowEvent: true });
    });
    template = template.replace('{First name} ', Office.context.mailbox.userProfile.displayName);
    template = template.replace('{Last name}', "");
    template = template.replaceAll('{E-mail}', Office.context.mailbox.userProfile.emailAddress);
    template = template.replace('{Title}', "Wipro Support");
   // document.getElementsByClassName("Signature").contentEditable = false;
  //  await getDesignation(Office.context.mailbox.userProfile.emailAddress, template);

    Office.context.mailbox.item?.body.setSignatureAsync(
      template,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          throw asyncResult.error.message;
        }
      }
    );
    //////////////////////////
     
     
      
        // // Disable the default client signature
        // Office.context.mailbox.item.disableClientSignatureAsync((asyncResult) => {
        //   if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        //     console.log("Default client signature disabled.");
    
        //     // Add your custom signature
        //     Office.context.mailbox.item.body.setSignatureAsync(
        //       template,
             
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
     
    

  
     
      //   // Check if the current item is a new email
      //   const item = Office.context.mailbox.item;
      //   if (item.itemType === Office.MailboxEnums.ItemType.Message && item.displayReplyAll === undefined) {
      //     // This is a new email (not a reply or forward)
      //     console.log("New email detected.");
    
      //     // Associate the function with the onMessageSend event
      //  //   Office.actions.associate("onMessageSend", onMessageSendHandler);
      //   }
     
   
    
    
    //////////////////////////////////////
  } catch (error) {
    console.log("Error: " + error);
  }

  async function getDesignation(email, newTemplate) {
    let accessToken = "";
    const graphUrl = `https://graph.microsoft.com/v1.0/users/${email}?$select=jobTitle`;
    Office.context.mailbox.item.saveAsync((result) => {
      Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          accessToken = result.value;
          console.log("Access Token:", accessToken);

          fetch(graphUrl, {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              "Content-Type": "application/json",
            },
          })
            .then((response) => response.json())
            .then((data) => {
              const jobTitle = data.jobTitle || "Wipro Support"; //only for testing 
              console.log("Job Title:", jobTitle);
              //   template=template.replace('{Title}',jobTitle);
            })
            .catch((error) => {
              console.error("Error fetching job title:", error);
            });
        } else {
          console.error("Failed to get access token:", result.error);
        }
      });
    });


  }
}


// Office.actions.associate("validateSignature", function (event) {
//   const message = Office.context.mailbox.item;
//   console.log("at line 110.............");
  
//   // Use .then() instead of async/await
//   message.body.getAsync(Office.CoercionType.Text).then(function (body) {
//     const required = ["Name:", "Title:", "Company:", "Email:", "Phone:"];
//     const isValid = required.every(field => body.value.includes(field));

//     if (!isValid) {
//       event.completed({ allowEvent: false });
//       // Show a subtle notification (not a dialog)
//       Office.context.mailbox.item.notificationMessages.addAsync("signatureError", {
//         type: "errorMessage",
//         message: "Signature validation failed. Check your signature fields.",
//       });
//     } else {
//       event.completed({ allowEvent: true });
//     }
//   }).catch(function (error) {
//     console.error("Error:", error);
//     event.completed({ allowEvent: true }); // Fail open (allow send)
//   });

//  //event.completed({ allowEvent: true }); // Always allow sending
// });


function validateSignature(event) {
  mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);
}
function validationFunction(event) {
  message.body.getAsync(Office.CoercionType.Text).then(function (body) {
    const required = ["Name:", "Title:", "Company:", "Email:", "Phone:"];
    const isValid = required.every(field => body.value.includes(field));

    if (!isValid) {
      event.completed({ allowEvent: false });
      // Show a subtle notification (not a dialog)
      Office.context.mailbox.item.notificationMessages.addAsync("signatureError", {
        type: "errorMessage",
        message: "Signature validation failed. Check your signature fields.",
      });
    } else {
      event.completed({ allowEvent: true });
    }
  }).catch(function (error) {
    console.error("Error:", error);
    event.completed({ allowEvent: true }); // Fail open (allow send)
  });
}
// Check if the body contains a specific set of blocked words. If it contains the blocked words, block email from being sent. Otherwise allows sending.
// <param name="asyncResult">MessageSend event passed from the calling function.</param>
function checkBodyOnlyOnSendCallBack(asyncResult) {
  var listOfBlockedWords = new Array("blockedword", "blockedword1", "blockedword2");
  var wordExpression = listOfBlockedWords.join('|');

  // \b to perform a "whole words only" search using a regular expression in the form of \bword\b.
  // i to perform case-insensitive search.
  var regexCheck = new RegExp('\\b(' + wordExpression + ')\\b', 'i');
  var checkBody = regexCheck.test(asyncResult.value);

  if (checkBody) {
      mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked words have been found in the body of this email. Please remove them.' });
      // Block send.
      asyncResult.asyncContext.completed({ allowEvent: false });
  }
  else {

      // Allow send.
      asyncResult.asyncContext.completed({ allowEvent: true });
  }
}