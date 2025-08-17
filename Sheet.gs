/**
 * This class must be roaded before chlid class.
 */
class Sheet {
  /**
   * Constructor.
   */
  constructor() {
    this.sheet;
    this.utils = new utils();
    this.header_r = 1;
    this.default_format = "yyyy/MM";
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
   * Get the row number in column [col] that matches [val].
   * @param {string} val
   * @param {number} col
   * @return {number}
   */
  findRow(val, col) {
    const values = this.getValues(this.header_r, col, this.getLastRow(), 1);

    let arr_values = [];

    for(var i = 0; i < values.length; i++){
      arr_values.push(values[i][0]);
    }

    return arr_values.indexOf(val) + 1;
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
  getPreviousMonth(format = this.default_format) {
    return this.utils.getFormattedMonth(-1, format);
  }

  /**
   * @param {null|string} format
   * @return {string}
   */
  getCurrentMonth(format = this.default_format) {
    return this.utils.getFormattedMonth(0, format);
  }

  /**
   * @param {null|string} format
   * @return {string}
   */
  getNextMonth(format = this.default_format) {
    return this.utils.getFormattedMonth(1, format);
  }
}