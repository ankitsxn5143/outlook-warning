/*
  commands.js - Recipient Warning Add-in
  ItemSend handler — compatible with Microsoft 365 Business Basic
*/

// ─── CONFIGURATION ────────────────────────────────────────────────────────────
const FLAGGED_ADDRESSES = [
  "ankit@bwesglobal.com",
  // "another@example.com",
];
// ──────────────────────────────────────────────────────────────────────────────

Office.initialize = function () {};

function onItemSend(event) {
  const item = Office.context.mailbox.item;
  let completed = false;

  function finish(allow) {
    if (completed) return;
    completed = true;
    if (!allow) {
      item.notificationMessages.addAsync("flagged", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: "⚠️ Send blocked: This email contains a flagged recipient. Please review your recipients before sending.",
      });
    }
    event.completed({ allowEvent: allow });
  }

  // Safety timeout — must complete within 5 seconds or Outlook shows the warning
  const timeout = setTimeout(() => finish(true), 4500);

  function getRecips(field) {
    return new Promise((resolve) => {
      if (!field) return resolve([]);
      field.getAsync({ asyncContext: null }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value.map(r => (r.emailAddress || "").toLowerCase()));
        } else {
          resolve([]);
        }
      });
    });
  }

  Promise.all([
    getRecips(item.to),
    getRecips(item.cc),
    getRecips(item.bcc),
  ]).then(([to, cc, bcc]) => {
    clearTimeout(timeout);
    const all = [...to, ...cc, ...bcc];
    const flaggedLower = FLAGGED_ADDRESSES.map(f => f.toLowerCase());
    const found = all.some(addr => flaggedLower.includes(addr));
    finish(!found); // block if found, allow if not
  }).catch(() => {
    clearTimeout(timeout);
    finish(true); // fail open
  });
}
