function getMediator() {
  return Mediator;
  //return new Mediator(SpreadsheetApp.getActiveSpreadsheet());
}
function setupMenu(prefix: string = "MediatorLib") {
  return Mediator.setupMenu(prefix);
}

function processEmails() {
  return Mediator.processEmails();
}