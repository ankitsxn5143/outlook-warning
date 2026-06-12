Office.onReady(() => {});

const TARGET = "ankit@bwesglobal.com";

async function onMessageSend(event) {

    const item = Office.context.mailbox.item;

    const getRecipients = (field) =>
        new Promise(resolve =>
            field.getAsync(r => resolve(r.value || []))
        );

    const recipients = [
        ...(await getRecipients(item.to)),
        ...(await getRecipients(item.cc)),
        ...(await getRecipients(item.bcc))
    ];

    const found = recipients.some(
        r => r.emailAddress &&
        r.emailAddress.toLowerCase() === TARGET
    );

    if (found) {
        event.completed({
            allowEvent: false,
            errorMessage:
            "TEST WARNING: ankit@bwesglobal.com is a recipient."
        });
    } else {
        event.completed({ allowEvent: true });
    }
}

Office.actions.associate(
    "onMessageSend",
    onMessageSend
);