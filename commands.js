/*
  commands.js - Recipient Warning Add-in
*/

const FLAGGED_ADDRESSES = ["ankit@bwesglobal.com"];

Office.initialize = function () {
  console.log("[RecipientWarning] Office.initialize called");
};

Office.onReady(function() {
  console.log("[RecipientWarning] Office.onReady called");
});

function onItemSend(event) {
  console.log("[RecipientWarning] onItemSend fired");
  const item = Office.context.mailbox.item;

  item.to.getAsync(function(result) {
    console.log("[RecipientWarning] to result:", result.status);
    var toEmails = [];
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      toEmails = result.value.map(function(r) { return r.emailAddress.toLowerCase(); });
    }

    var found = toEmails.some(function(e) {
      return FLAGGED_ADDRESSES.indexOf(e) !== -1;
    });

    if (found) {
      item.notificationMessages.addAsync("warning", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: "Send blocked: flagged recipient detected."
      }, function() {
        event.completed({ allowEvent: false });
      });
    } else {
      event.completed({ allowEvent: true });
    }
  });
}

// Explicitly register with Office runtime
if (typeof Office.actions !== "undefined") {
  Office.actions.associate("onItemSend", onItemSend);
}

// Also expose globally
if (typeof window !== "undefined") {
  window.onItemSend = onItemSend;
}
