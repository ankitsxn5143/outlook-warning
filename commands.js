var FLAGGED_ADDRESSES = ["ankit@bwesglobal.com"];

console.log("[RecipientWarning] commands.js loaded");

var onItemSend = function(event) {
  console.log("[RecipientWarning] onItemSend fired");
  var item = Office.context.mailbox.item;
  item.to.getAsync(function(result) {
    console.log("[RecipientWarning] to.getAsync result:", result.status);
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
    console.log("[RecipientWarning] found flagged:", found);
    if (found) {
      item.notificationMessages.addAsync("warning", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: "Send blocked: flagged recipient detected."
      }, function() { 
        console.log("[RecipientWarning] send blocked");
        event.completed({ allowEvent: false }); 
      });
    } else {
      console.log("[RecipientWarning] allowing send");
      event.completed({ allowEvent: true });
    }
  });
};

console.log("[RecipientWarning] onItemSend registered:", typeof onItemSend);
