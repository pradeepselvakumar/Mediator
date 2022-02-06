class Mediator {
  
  //private _activeSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  private static get activeSpreadsheet() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  
  
  static setupMenu(prefix: string) {
    const menuEntries = [
      { name: "Process emails", functionName: `${prefix}.processEmails` }
    ];

    Mediator.activeSpreadsheet.addMenu("Mediator", menuEntries);
  }
  
  static processEmails() {
    let afterDate = Mediator.getDate();
    let threads = GmailApp.search(`from:(medium daily digest) after:${afterDate}`);
  
  //let d = threads[0].getLastMessageDate();

    // Logger.log(threads[0].getMessages()[0].getBody());
  }

  /**
 * Returns the date of the last processed email. If none are available, returns
 * the date from a month ago.
 * Uses the active worksheet in the active spreadsheet to look for previously processed emails
 */
  static getDate() {
    let worksheet = Mediator.activeSpreadsheet.getActiveSheet();

    let date = worksheet.getRange("B2").getValues()[0][1];
    Logger.log(date);
  }
}
