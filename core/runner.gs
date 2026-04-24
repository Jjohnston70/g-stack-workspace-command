function runCommand(action, payload) {

  switch(action) {

    case "DRIVE_CREATE_FOLDER":
      return DriveModule.createFolder(payload.name);

    case "DRIVE_LIST":
      return DriveModule.listFiles();

    case "GMAIL_CREATE_LABEL":
      return GmailModule.createLabel(payload.name);

    case "GMAIL_LIST":
      return GmailModule.listLabels();

    default:
      throw new Error("Unknown action: " + action);
  }
}