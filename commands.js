/*
  commands.js
  Handles the OnMessageSend Smart Alert event.
  Checks To, CC, BCC for flagged addresses and shows a warning dialog.
*/

// ─── CONFIGURATION ────────────────────────────────────────────────────────────
const FLAGGED_ADDRESSES = [
  "ankit@bwesglobal.com",
  // "another@example.com",
];

const ADDIN_BASE_URL = "https://ankitsxn5143.github.io/outlook-warning";
// ──────────────────────────────────────────────────────────────────────────────

Office.onReady(() => {});

async function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;

    const [toRecips, ccRecips, bccRecips] = await Promise.all([
      getRecipients(item.to),
      getRecipients(item.cc),
      getRecipients(item.bcc),
    ]);

    const allRecipients = [...toRecips, ...ccRecips, ...bccRecips];
    const flagged = allRecipients.filter(addr =>
      FLAGGED_ADDRESSES.includes(addr.toLowerCase())
    );

    if (flagged.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    const flaggedParam = encodeURIComponent(flagged.join(", "));
    const dialogUrl = `${ADDIN_BASE_URL}/warning.html?flagged=${flaggedParam}`;

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 38, width: 35, displayInIframe: true },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          event.completed({ allowEvent: false });
          return;
        }

        const dialog = asyncResult.value;

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (messageArg) => {
            dialog.close();
            let userChoice = { action: "no" };
            try { userChoice = JSON.parse(messageArg.message); } catch (e) {}

            if (userChoice.action === "yes") {
              event.completed({ allowEvent: true });
            } else {
              event.completed({ allowEvent: false });
            }
          }
        );

        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          () => {
            dialog.close();
            event.completed({ allowEvent: false });
          }
        );
      }
    );
  } catch (err) {
    console.error("onMessageSendHandler error:", err);
    event.completed({ allowEvent: true });
  }
}

function getRecipients(recipientField) {
  return new Promise((resolve) => {
    if (!recipientField) { resolve([]); return; }
    recipientField.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve(asyncResult.value.map(r => r.emailAddress || "").filter(Boolean));
      } else {
        resolve([]);
      }
    });
  });
}

// Required for event-based activation
if (Office.actions) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
