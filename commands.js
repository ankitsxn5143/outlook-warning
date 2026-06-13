/*
  commands.js - Recipient Warning Add-in (minimal version)
*/

const FLAGGED_ADDRESSES = ["ankit@bwesglobal.com"];

Office.initialize = function () {};

function onItemSend(event) {
  try {
    const item = Office.context.mailbox.item;

    item.to.getAsync(function (toResult) {
      item.cc.getAsync(function (ccResult) {
        item.bcc.getAsync(function (bccResult) {

          var allEmails = [];

          if (toResult.status === Office.AsyncResultStatus.Succeeded) {
            toResult.value.forEach(function(r) { allEmails.push(r.emailAddress.toLowerCase()); });
          }
          if (ccResult.status === Office.AsyncResultStatus.Succeeded) {
            ccResult.value.forEach(function(r) { allEmails.push(r.emailAddress.toLowerCase()); });
          }
          if (bccResult.status === Office.AsyncResultStatus.Succeeded) {
            bccResult.value.forEach(function(r) { allEmails.push(r.emailAddress.toLowerCase()); });
          }

          var found = false;
          for (var i = 0; i < allEmails.length; i++) {
            if (FLAGGED_ADDRESSES.indexOf(allEmails[i]) !== -1) {
              found = true;
              break;
            }
          }

          if (found) {
            item.notificationMessages.addAsync("warning", {
              type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
              message: "Send blocked: This email contains a flagged recipient. Please review before sending."
            }, function() {
              event.completed({ allowEvent: false });
            });
          } else {
            event.completed({ allowEvent: true });
          }

        });
      });
    });

  } catch (e) {
    event.completed({ allowEvent: true });
  }
}
