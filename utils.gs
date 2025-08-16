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

    return Utilities.formatDate(today, TIMEZONE, format);
  }

  /**
   * @param {number} str
   * @param {number} end
   * @return {number}
   */
  getCellDiff(str, end) {
    return end - str + 1;
  }
}