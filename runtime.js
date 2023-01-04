function onMessageSendHandler(event) {
  event.completed({ allowEvent: false, errorMessage: "testing again" });
}

//Office.actions.associate("checkSignature", onMessageSendHandler);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
