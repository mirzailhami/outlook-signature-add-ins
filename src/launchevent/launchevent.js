/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// Add start-up logic code here, if any.
Office.onReady((info) => {
  console.log(Office.context?.mailbox?.diagnostics);
  console.log(info);
  console.log(Office.context);

  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
});

function onNewMessageComposeHandler(event) {
  const item = Office.context.mailbox.item;
  const signatureIcon =
    "iVBORw0KGgoAAAANSUhEUgAAACcAAAAnCAMAAAC7faEHAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAzUExURQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKMFRskAAAAQdFJOUwAQIDBAUGBwgI+fr7/P3+8jGoKKAAAACXBIWXMAAA7DAAAOwwHHb6hkAAABT0lEQVQ4T7XT2ZalIAwF0DAJhMH+/6+tJOQqot6X6joPiouNBo3w9/Hd6+hrYnUt6vhLcjEAJevVW0zJxABSlcunhERpjY+UKoNN5+ZgDGu2onNz0OngjP2FM1VdyBW1LtvGeYrBLs7U5I1PTXZt+zifcS3Icw2GcS3vxRY3Vn/iqx31hUyTnV515kdTfbaNhZLI30AceqDiIo4tyKEmJpKdP5M4um+nUwfDWxAXdzqMNKQ14jLdL5ntXzxcRF440mhS6yu882Kxa30RZcUIjTCJg7lscsR4VsMjfX9Q0Vuv/Wd3YosD1J4LuSRtaL7bzXGN1wx2cytUdncDuhA3fu6HPTiCvpQUIjZ3sCcHVbvLtbNTHlysx2w9/s27m9gEb+7CTri6hR1wcTf2gVf3wBRe3CMbcHYvTODkXhnD0+178K/pZ9+n/C1ru/2HAPwAo7YM1X4+tLMAAAAASUVORK5CYII=";

  // Get the sender's account information.
  item.from.getAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log(result.error.message);
      event.completed();
      return;
    }

    // Create a signature based on the sender's information.
    const name = result.value.displayName;
    const options = { asyncContext: name, isInline: true };
    item.addFileAttachmentFromBase64Async(signatureIcon, "signatureIcon.png", options, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log(result.error.message);
        event.completed();
        return;
      }

      // Add the created signature to the message.
      const signature = "<img src='cid:signatureIcon.png'>" + result.asyncContext;
      item.body.setSignatureAsync(signature, { coercionType: Office.CoercionType.Html }, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log(result.error.message);
          event.completed();
          return;
        }

        // Show a notification when the signature is added to the message.
        const notification = {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: "Company signature added.",
          icon: "none",
          persistent: false,
        };
        item.notificationMessages.addAsync("signature_notification", notification, (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.log(result.error.message);
            event.completed();
            return;
          }

          event.completed();
        });
      });
    });
  });
}
