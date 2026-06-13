var FLAGGED_ADDRESSES = ["performancedesk@bwesglobal.com"];

console.log("[RecipientWarning] commands.js loaded");

Office.onReady(function() {
  console.log("[RecipientWarning] Office.onReady called");
});

function onMessageSendHandler(event) {
  console.log("[RecipientWarning] onMessageSendHandler fired!");
  var item = Office.context.mailbox.item;

  item.to.getAsync(function(result) {
    var found = false;
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      for (var i = 0; i < result.value.length; i++) {
        var email = result.value[i].emailAddress.toLowerCase();
        console.log("[RecipientWarning] checking:", email);
        if (FLAGGED_ADDRESSES.indexOf(email) !== -1) {
          found = true;
          break;
        }
      }
    }

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

console.log("[RecipientWarning] onMessageSendHandler defined:", typeof onMessageSendHandler);
