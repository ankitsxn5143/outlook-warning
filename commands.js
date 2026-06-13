/*
  commands.js - Recipient Warning Add-in (Debug Version)
*/

const FLAGGED_ADDRESSES = ["ankit@bwesglobal.com"];

Office.initialize = function () {};

function onItemSend(event) {
  const item = Office.context.mailbox.item;

  // Log every step to console so we can see exactly where it fails
  console.log("[RecipientWarning] onItemSend fired");
  console.log("[RecipientWarning] item:", item);
  console.log("[RecipientWarning] item.to:", item.to);
  console.log("[RecipientWarning] item.cc:", item.cc);
  console.log("[RecipientWarning] item.bcc:", item.bcc);

  item.to.getAsync(function(result) {
    console.log("[RecipientWarning] to.getAsync result:", result.status, result.value, result.error);

    var toEmails = [];
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      toEmails = result.value.map(function(r) { return r.emailAddress.toLowerCase(); });
    }

    console.log("[RecipientWarning] To emails:", toEmails);

    var flagged = toEmails.filter(function(e) {
      return FLAGGED_ADDRESSES.indexOf(e) !== -1;
    });

    console.log("[RecipientWarning] Flagged:", flagged);

    if (flagged.length > 0) {
      console.log("[RecipientWarning] Blocking send");
      item.notificationMessages.addAsync("warning", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: "Send blocked: flagged recipient detected (" + flagged.join(", ") + ")."
      }, function(r) {
        console.log("[RecipientWarning] notification added:", r.status);
        event.completed({ allowEvent: false });
      });
    } else {
      console.log("[RecipientWarning] Allowing send");
      event.completed({ allowEvent: true });
    }
  });
}
