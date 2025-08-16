class CashFlow{
  /**
   * Constructor.
   */
  constructor() {
    this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CF管理");;
    this.utils = new utils();
    this.header_r = 1;
    this.diff_last_and_latest = 4;
  }

  /**
   * Set asset amount of current month.
   */
  setCurrentAsset() {
    const cur_month_r = this.getRowNumOfCurrentMonth();
    const amount = this.getCellValue(cur_month_r, COLUMNS["n"]);

    this.setCellValue(cur_month_r, COLUMNS["j"], amount);
  }

  /**
   * Insert next month row and set values.
   */
  addNextRow() {
    const cur_month_r = this.getRowNumOfCurrentMonth();
    const next_month_r = cur_month_r + 1;

    this.sheet.insertRows(next_month_r);

    const values = [
      [
        this.utils.getMonth(MONTH["next"]),
        `=IFERROR(INDIRECT("'"&$A${next_month_r}&"'!$P$42"), 0)`,
        `=IFERROR(INDIRECT("'"&$A${next_month_r}&"'!$P$45"), 0)`,
        `=IFERROR(INDIRECT("'"&$A${next_month_r}&"'!$P$50"), 0)`,
        `=IFERROR(INDIRECT("'"&$A${next_month_r}&"'!$P$55"), 0)`,
        `=IFERROR(INDIRECT("'"&$A${next_month_r}&"'!$P$57"), 0)`,
        `=IFERROR(INDIRECT("'"&$A${next_month_r}&"'!$P$61"), 0)`,
        `=IFERROR(INDIRECT("'"&$A${next_month_r}&"'!$I$54"), 0)`,
        `=H${next_month_r}-G${next_month_r}`,
        null,
        `=IF(J${next_month_r}-J${cur_month_r}>0,J${next_month_r}-J${cur_month_r},0)`,
      ],
    ];

    this.setValues(next_month_r, COLUMNS["a"], values.length, COLUMNS["k"], values)
  }

  /**
   * Move the summary table down [count] row(s).
   * @param {number} count
   */
  moveSummaryTable(count) {
    const str_c = COLUMNS["m"];
    const str_r = this.findRow("【資産管理】", str_c);
    const table_length = 12;
    const table_width = this.getColDiff(str_c, COLUMNS["n"]);

    const values = this.sheet.getRange(str_r, str_c, table_length, table_width);
    const range = this.sheet.getRange((str_r + count), str_c, table_length, table_width);

    values.moveTo(range);
  }

  /**
   * Update Reference month based on setting of diff month.
   */
  changeReferenceMonth() {
    const target_r = this.getLastRow() - 1;
    const diffMonth = this.getCellValue(target_r, COLUMNS["e"]);
    const refMonth = this.utils.getMonth(1 - diffMonth);
    
    this.setCellValue(target_r, COLUMNS["b"], refMonth);
  }

  /**
   * Update display range of asset table.
   */
  changeGraphRange() {
    const startMonth = this.getCellValue((this.getLastRow() - 1), COLUMNS["b"]);
    const endMonth = this.getCellValue(this.getLastRow(), COLUMNS["b"]);
  
    let str_r = this.findRow(startMonth, COLUMNS["a"]);
    let end_r = this.findRow(endMonth, COLUMNS["a"])

    if (end_r === 0) {
      end_r = this.getRowNumOfCurrentMonth();
      // update diff value
      const actualDiffMonth = this.getMonthDiff(this.getCellValue(str_r, COLUMNS["a"]), this.getCellValue(end_r, COLUMNS["a"]));
      this.setCellValue((this.getLastRow() - 1), COLUMNS["e"], actualDiffMonth);
    }

    // display all rows once.
    this.sheet.unhideRow(this.sheet.getRange(this.header_r, COLUMNS["a"], this.getLastRow(), 1));

    const range1rows = str_r - 2;

    // hide the range before start month
    if (range1rows > 0) {
      const range1 = this.sheet.getRange((this.header_r + 1), COLUMNS["a"], range1rows, 1);
      this.sheet.hideRow(range1);
    }

    const range2rows = this.getLatestRow() - end_r;

    // hide the range after end month
    if (range2rows > 0) {
      const range2 = this.sheet.getRange((end_r + 1), COLUMNS["a"], range2rows, 1);
      this.sheet.hideRow(range2);
    }
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
   * @return {number}
   */
  getLatestRow() {
    return this.getLastRow() - this.diff_last_and_latest;
  }

  /**
   * @param {number} str_c
   * @param {number} end_c
   * @return {number}
   */
  getColDiff(str_c, end_c) {
    return end_c - str_c + 1;
  }

  /**
   * @return {number}
   */
  getRowNumOfCurrentMonth() {
    const currentMonth = this.utils.getMonth(MONTH["current"]);

    return this.findRow(currentMonth, COLUMNS["a"]);
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
   * @param {string} startMonth ex."2023/01"
   * @param {string} endMonth   ex."2024/05"
   * @return {number}
   */
  getMonthDiff(startMonth, endMonth) {
    const startDate = new Date(startMonth + '/01');
    const endDate = new Date(endMonth + '/01');

    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      throw new Error("Invalid format: 'yyyy/mm' is required.");
    }

    const years = endDate.getFullYear() - startDate.getFullYear();
    const months = endDate.getMonth() - startDate.getMonth();
    
    return Math.abs(years * 12 + months) + 1;
  }
}