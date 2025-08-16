class utils {
  /**
   * Get the month in which the difference is calculated from the current date in [format].
   * @param {number} diff
   * @param {string} format
   * @return {string}
   */
  getMonth(diff, format = MONTH_FORMAT) {
    const today = new Date();
    today.setMonth(today.getMonth() + diff);

    return Utilities.formatDate(today, "JST", format);
  }

  /**
   * @param {SpreadsheetApp.Sheet} sheet
   * @param {number} lastRow
   * @param {null|number} col
   * @return {number|false}
   */
  getRowNumOfCurrentMonth(sheet, lastRow, col = 1) {
    const currentMonth = this.getMonth(MONTH["current"]);
    
    for(let i = 1; i <= lastRow; i++) {
      let target = sheet.getRange(i, col).getValue();

      if(target == currentMonth){ 
        return i;
      }
    }

    return false;
  }

  /**
   * @param {string} val 
   * @param {null|number} date
   * @return {Date}
   */
  convertYYYYMMtoDate(val, date = 1) {
    const regex = /^\d{4}\/\d{2}$/;

    if (!regex.test(val)) {
      throw new Error("Invalid format: val is not 'YYYY/MM' format.")
    }

    const parts = val.split('/');
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;

    return new Date(year, month, date);
  }
}