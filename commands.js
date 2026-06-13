/*
  commands.js - Recipient Warning Add-in
  Uses synchronous EWS approach to check recipients faster.
*/

// ─── CONFIGURATION ────────────────────────────────────────────────────────────
const FLAGGED_ADDRESSES = [
  "ankit@bwesglobal.com",
];
// ──────────────────────────────────────────────────────────────────────────────

Office.initialize = function () {};

function onItemSend(event) {
  const item = Office.context.mailbox.item;

  // Use makeEwsRequestAsync to get recipients directly — much faster than getAsync
  const ewsRequest =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
    '  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <t:RequestServerVersion Version="Exchange2013"/>' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '          <t:FieldURI FieldURI="message:ToRecipients"/>' +
    '          <t:FieldURI FieldURI="message:CcRecipients"/>' +
    '          <t:FieldURI FieldURI="message:BccRecipients"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds>' +
    '        <t:ItemId Id="' + item.itemId + '"/>' +
    '      </ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

  // Fallback: if itemId not available (new unsaved draft), use getAsync
  if (!item.itemId) {
    checkWithGetAsync(event, item);
    return;
  }

  Office.context.mailbox.makeEwsRequestAsync(ewsRequest, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      checkWithGetAsync(event, item);
      return;
    }

    try {
      const xml = result.value;
      const emails = [];
      const matches = xml.match(/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g);
      if (matches) {
        matches.forEach(e => emails.push(e.toLowerCase()));
      }

      const flaggedLower = FLAGGED_ADDRESSES.map(f => f.toLowerCase());
      const found = emails.some(addr => flaggedLower.includes(addr));

      if (found) {
        item.notificationMessages.addAsync("flagged", {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: "⚠️ Send blocked: This email contains a flagged recipient (" +
            FLAGGED_ADDRESSES.filter(f => emails.includes(f.toLowerCase())).join(", ") +
            "). Please review before sending.",
        });
        event.completed({ allowEvent: false });
      } else {
        event.completed({ allowEvent: true });
      }
    } catch (e) {
      event.completed({ allowEvent: true });
    }
  });
}

function checkWithGetAsync(event, item) {
  let done = false;

  function finish(allow) {
    if (done) return;
    done = true;
    event.completed({ allowEvent: allow });
  }

  setTimeout(() => finish(true), 4000);

  function getRecips(field) {
    return new Promise((resolve) => {
      if (!field) return resolve([]);
      field.getAsync((r) => {
        resolve(r.status === Office.AsyncResultStatus.Succeeded
          ? r.value.map(x => (x.emailAddress || "").toLowerCase())
          : []);
      });
    });
  }

  Promise.all([getRecips(item.to), getRecips(item.cc), getRecips(item.bcc)])
    .then(([to, cc, bcc]) => {
      const all = [...to, ...cc, ...bcc];
      const flaggedLower = FLAGGED_ADDRESSES.map(f => f.toLowerCase());
      const found = all.some(addr => flaggedLower.includes(addr));
      if (found) {
        item.notificationMessages.addAsync("flagged", {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: "⚠️ Send blocked: Flagged recipient detected. Please review before sending.",
        });
      }
      finish(!found);
    })
    .catch(() => finish(true));
}
