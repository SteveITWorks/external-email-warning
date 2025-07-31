Office.initialize = () => {};
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

async function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;
    const recipients = (item.to || []).concat(item.cc || []);
    const internalDomain = "@yourcompany.com".toLowerCase(); // CHANGE THIS

    const hasExternal = recipients.some(r => {
      const addr = (r.emailAddress || "").toLowerCase();
      return addr.includes("@") && !addr.endsWith(internalDomain);
    });

    if (hasExternal) {
      Office.context.ui.displayDialogAsync(
        "https://SteveITWorks.github.io/external-email-warning/confirm.html",
        { height: 30, width: 30 },
        result => {
          const dlg = result.value;
          dlg.addEventHandler(Office.EventType.DialogMessageReceived, arg => {
            dlg.close();
            event.completed({ allowEvent: arg.message === "send" });
          });
        }
      );
    } else {
      event.completed({ allowEvent: true });
    }
  } catch (e) {
    console.error("Error in on-send handler:", e);
    event.completed({ allowEvent: true });
  }
}
