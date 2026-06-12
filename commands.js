/*
  commands.js - Recipient Warning Add-in
  Event-based activation handler for OnMessageSend (Smart Alerts)
*/

// ─── CONFIGURATION ────────────────────────────────────────────────────────────
const FLAGGED_ADDRESSES = [
  "ankit@bwesglobal.com",
  // "another@example.com",  // add more here
];
const ADDIN_BASE_URL = "https://ankitsxn5143.github.io/outlook-warning";
// ──────────────────────────────────────────────────────────────────────────────

function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  // Collect To, CC, BCC
  const getRecips = (field) => new Promise((resolve) => {
    if (!field) return resolve([]);
    field.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value.map(r => (r.emailAddress || "").toLowerCase()).filter(Boolean));
      } else {
        resolve([]);
      }
    });
  });

  Promise.all([getRecips(item.to), getRecips(item.cc), getRecips(item.bcc)])
    .then(([to, cc, bcc]) => {
      const all = [...to, ...cc, ...bcc];
      const flagged = all.filter(addr => FLAGGED_ADDRESSES.map(f => f.toLowerCase()).includes(addr));

      if (flagged.length === 0) {
        event.completed({ allowEvent: true });
        return;
      }

      const flaggedParam = encodeURIComponent(flagged.join(", "));
      const dialogUrl = `${ADDIN_BASE_URL}/warning.html?flagged=${flaggedParam}`;

      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 40, width: 40, displayInIframe: false },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Dialog failed to open:", asyncResult.error);
            event.completed({ allowEvent: false });
            return;
          }

          const dialog = asyncResult.value;

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (msg) => {
            dialog.close();
            let choice = { action: "no" };
            try { choice = JSON.parse(msg.message); } catch (e) {}
            event.completed({ allowEvent: choice.action === "yes" });
          });

          dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
            dialog.close();
            event.completed({ allowEvent: false });
          });
        }
      );
    })
    .catch((err) => {
      console.error("Error in onMessageSendHandler:", err);
      event.completed({ allowEvent: true });
    });
}

// CRITICAL: Must be called at top level for event-based activation
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
