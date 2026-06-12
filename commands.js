/*
  commands.js
  -----------
  This file handles the OnMessageSend Outlook event.
  It checks To, CC, and BCC for any flagged email addresses.
  If found, it opens a warning dialog. The user's Yes/No
  choice either allows the send or cancels it.
*/

// ─── CONFIGURATION ────────────────────────────────────────────────────────────
// Add any email addresses you want to flag (all lowercase).
const FLAGGED_ADDRESSES = [
  "ankit@bwesglobal.com",
  // "another@example.com",  // ← add more here
];

// The public URL where your add-in is hosted (GitHub Pages or your server).
// Update this to match your deployed URL.
const ADDIN_BASE_URL = "https://ankitsxn5143.github.io/outlook-warning";
// ──────────────────────────────────────────────────────────────────────────────

Office.onReady(() => {
  // No-op: event-based activation uses the exported function below.
});

/**
 * Called by Outlook when the user clicks Send.
 * @param {Office.AddinCommands.Event} event
 */
async function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;

    // Collect all recipients from To, CC, BCC
    const [toRecips, ccRecips, bccRecips] = await Promise.all([
      getRecipients(item.to),
      getRecipients(item.cc),
      getRecipients(item.bcc),
    ]);

    const allRecipients = [...toRecips, ...ccRecips, ...bccRecips];

    // Find any that match our flagged list
    const flagged = allRecipients.filter(addr =>
      FLAGGED_ADDRESSES.includes(addr.toLowerCase())
    );

    if (flagged.length === 0) {
      // No flagged addresses — allow send immediately
      event.completed({ allowEvent: true });
      return;
    }

    // Open warning dialog
    const flaggedParam = encodeURIComponent(flagged.join(", "));
    const dialogUrl = `${ADDIN_BASE_URL}/warning.html?flagged=${flaggedParam}`;

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 38, width: 35, displayInIframe: true },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          // If dialog can't open, default to blocking the send for safety
          event.completed({ allowEvent: false });
          return;
        }

        const dialog = asyncResult.value;

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (messageArg) => {
            dialog.close();

            let userChoice = { action: "no" };
            try {
              userChoice = JSON.parse(messageArg.message);
            } catch (e) { /* default to no */ }

            if (userChoice.action === "yes") {
              event.completed({ allowEvent: true });   // ✅ Send
            } else {
              event.completed({ allowEvent: false });  // 🚫 Cancel
            }
          }
        );

        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          (errorArg) => {
            // Dialog was closed without a choice (e.g. X button) → cancel
            dialog.close();
            event.completed({ allowEvent: false });
          }
        );
      }
    );
  } catch (err) {
    console.error("onMessageSendHandler error:", err);
    event.completed({ allowEvent: true }); // fail open so user isn't stuck
  }
}

/**
 * Helper: wraps item.to/cc/bcc.getAsync in a Promise
 * and returns an array of email address strings.
 */
function getRecipients(recipientField) {
  return new Promise((resolve) => {
    if (!recipientField) {
      resolve([]);
      return;
    }
    recipientField.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const emails = asyncResult.value.map(r => r.emailAddress || "");
        resolve(emails.filter(Boolean));
      } else {
        resolve([]);
      }
    });
  });
}

// Make the handler globally accessible for the manifest
// (required for event-based activation)
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
