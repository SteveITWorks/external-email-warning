Office.initialize = () => {
  // No UI initialization needed for OnMessageSend
};

Office.actions.associate("onMessageSend", onMessageSend);

async function onMessageSend(event) {
  try {
    const item = Office.context.mailbox.item;

    item.getAllInternetRecipientsAsync(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const recipients = result.value;
        const internalDomain = "@itworks.co.nz"; // 🔁 Change this to your actual domain

        const hasExternal = recipients.some(email =>
          email.toLowerCase().includes("@") &&
          !email.toLowerCase().endsWith(internalDomain)
        );

        if (hasExternal) {
          Office.context.ui.displayDialogAsync(
            "https://steveitworks.github.io/external-email-warning/confirm.html",
            { height: 30, width: 30 },
            result => {
              const dlg = result.value;
              dlg.addEventHandler(Office.EventType.DialogMessageReceived, arg => {
                dlg.close();
                event.completed({ allowEvent: arg.message === "send" });
              });
              dlg.addEventHandler(Office.EventType.DialogEventReceived, () => {
                dlg.close();
                event.completed({ allowEvent: false });
              });
            }
          );
        } else {
          event.completed({ allowEvent: true });
        }
      } else {
        console.error("Could not retrieve recipients.");
        event.completed({ allowEvent: true });
      }
    });
  } catch (e) {
    console.error("Error in onMessageSend:", e);
    event.completed({ allowEvent: true });
  }
}
