/**
 * This class must be roaded before chlid class.
 */
class Sheet {
  /**
   * Constructor.
   */
  constructor() {
    this.sheet;
    // this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CF管理");
    this.utils = new utils();
    this.header_r = 1;
  }

  /**
   * @return {number}
   */
  getLastRow() {
    return this.sheet.getLastRow();
  }

  /**
   * @param {number} row
   * @param {number} col
   * @return {string}
   */
  getCellValue(row, col) {
    return this.sheet.getRange(row, col).getValue();
  }

  /**
   * @param {number} row
   * @param {number} col
   * @param {number} numRows
   * @param {number} numColumns
   * @return {array}
   */
  getValues(row, col, numRows, numColumns) {
    return this.sheet.getRange(row, col, numRows, numColumns).getValues();
  }

  /**
   * @param {number} row
   * @param {number} col
   * @param {string} val
   */
  setCellValue(row, col, val) {
    this.sheet.getRange(row, col).setValue(val);
  }

  /**
   * @param {number} row
   * @param {number} col
   * @param {number} numRows
   * @param {number} numColumns
   * @param {array} values
   */
  setValues(row, col, numRows, numColumns, values) {
    this.sheet.getRange(row, col, numRows, numColumns).setValues(values);
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