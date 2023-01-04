function onMessageSendHandler(event) {
  event.completed({ allowEvent: false, errorMessage: "test" });
}

//Office.actions.associate("checkSignature", onMessageSendHandler);
Office.actions.associate("checkSignature", onMessageSendHandler);
