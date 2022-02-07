class Mediator {

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
    const afterDate = Mediator.getDate();
    const query = `from:(medium daily digest) after:${afterDate.toLocaleDateString("en-US")}`;
    
    const threads = GmailApp.search(query);
  
    Logger.log(`Retrieved ${threads.length} threads matching "${query}"`);
  
    if (threads.length == 0) {
      return;
    }

    const modifiedLinks = threads.map(thread => Mediator.getModifiedLinks(thread));

    const activeSheet = Mediator.activeSpreadsheet.getActiveSheet();
    const dataRange = activeSheet.getDataRange();
    const existingData = dataRange.getValues();
    const header = existingData.splice(0, 1);

    const modifiedData = header.concat(...modifiedLinks).concat(existingData);

    dataRange.clearContent()
    activeSheet.getRange(1, 1, modifiedData.length, modifiedData[0].length).setValues(modifiedData)
  }

  static getModifiedLinks(thread: GoogleAppsScript.Gmail.GmailThread): Array<[string, Date]> {
    const urls = Mediator.processMessage(thread.getMessages()[0]);
    const date = thread.getLastMessageDate() as Date;

    return urls.map(url => [url, date]);
  }

  /**
 * Returns the date of the last processed email. If none are available, returns
 * the date from a month ago.
 * Uses the active worksheet in the active spreadsheet to look for previously processed emails
 */
  static getDate(): Date {
    let worksheet = Mediator.activeSpreadsheet.getActiveSheet();

    let date = worksheet.getRange("B2").getValues()[0][0] as Date;

    if (!date) {
      date = new Date();
      date.setDate(date.getDate() - 7);
    } else {
      date.setDate(date.getDate() + 1);
    }

    return date;
  }

  static processMessage(msg: GoogleAppsScript.Gmail.GmailMessage): Array<string> {
    const msgBody = msg.getBody();
    const regex = /<a href="(https:\/\/medium\.com\/@\w+\/[^\?]*)\?/g;
    
    let urls = []
    let match;
    do {
      match = regex.exec(msgBody);
      if (match) {
        urls.push(match[1]);
      }
    } while (match)
    
    // some of the links are duplicated, so dedup
    return urls.filter((val, index, self) => {return self.indexOf(val) === index});
  }
}