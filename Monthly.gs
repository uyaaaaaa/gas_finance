class Monthly{
  /**
   * Constructor.
   */
  constructor() {
    this.templateName = "テンプレ";
    this.utils = new utils();
    this.date_col = 1;
    this.sum_table_label_row = 41;
  }

  /**
   * Create new sheet of next month.
   */
  createNewSheet() {
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    const template = activeSheet.getSheetByName(this.templateName);
    const newSheet = template.copyTo(activeSheet);
    
    const sheetName = this.getNextMonth();
    const nextMonth = (new Date().getMonth()) + 2;

    newSheet.showSheet();
    newSheet.setName(sheetName);
    newSheet.getRange(this.date_col, COLUMNS["a"]).setValue(sheetName);
    newSheet.getRange(this.sum_table_label_row, COLUMNS["h"]).setValue(`【${nextMonth}月合算】`);
  }

  /**
   * Hide the sheet of previous month.
   */
  hidePrevious() {
    const sheetName = this.getPreviousMonth();

    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).hideSheet();
  }

  /**
   * Set protection to the sheet of current month.
   */
  protectCurrent() {
    const sheetName = this.getCurrentMonth();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    const protectionThis = sheet.protect();
    protectionThis.setDescription("fixed.");
  }

  /**
   * @param {null|string} format
   * @return {string}
   */
  getPreviousMonth(format = MONTH_FORMAT) {
    return this.utils.getMonth(MONTH["prev"], format);
  }

  /**
   * @param {null|string} format
   * @return {string}
   */
  getCurrentMonth(format = MONTH_FORMAT) {
    return this.utils.getMonth(MONTH["current"], format);
  }

  /**
   * @param {null|string} format
   * @return {string}
   */
  getNextMonth(format = MONTH_FORMAT) {
    return this.utils.getMonth(MONTH["next"], format);
  }
}