class CashFlow{
  /**
   * Constructor.
   */
  constructor() {
    this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CF管理");;
    this.utils = new utils();
    this.header_r = 1;
    this.diff_last_and_latest = 4;
    this.asset_table_length = 12;
    this.asset_table_width = 2;
  }

  /**
   * Set asset amount of current month.
   */
  setCurrentAsset() {
    const rowNum = this.utils.getRowNumOfCurrentMonth(this.sheet, this.getLastRow());
    // const rotNum = this.findRow(this.utils.getMonth(MONTH["current"]), COLUMNS["a"]);
    const amount = this.getCellValue(rowNum, COLUMNS["n"]);

    this.setCellValue(rowNum, COLUMNS["j"], amount);
  }

  /**
   * Insert next month row and set values.
   */
  addNextRow() {
    const latestRow = this.getLatestRow();
    const newRow = latestRow + 1;

    this.sheet.insertRows(newRow);

    const values = [
      [
        this.utils.getMonth(MONTH["next"]),
        `=IFERROR(INDIRECT("'"&$A${newRow}&"'!$P$42"), 0)`,
        `=IFERROR(INDIRECT("'"&$A${newRow}&"'!$P$45"), 0)`,
        `=IFERROR(INDIRECT("'"&$A${newRow}&"'!$P$50"), 0)`,
        `=IFERROR(INDIRECT("'"&$A${newRow}&"'!$P$55"), 0)`,
        `=IFERROR(INDIRECT("'"&$A${newRow}&"'!$P$57"), 0)`,
        `=IFERROR(INDIRECT("'"&$A${newRow}&"'!$P$61"), 0)`,
        `=IFERROR(INDIRECT("'"&$A${newRow}&"'!$I$54"), 0)`,
        `=H${newRow}-G${newRow}`,
        null,
        `=IF(J${newRow}-J${latestRow}>0,J${newRow}-J${latestRow},0)`,
      ],
    ];

    this.setValues(newRow, COLUMNS["a"], values.length, COLUMNS["k"], values)
  }

  /**
   * Move the summary table down [count] row(s).
   * @param {number} count
   */
  moveSummaryTable(count) {
    const targetRow = this.sheet.getRange(this.header_r, COLUMNS["m"]).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
    const values = this.sheet.getRange(targetRow, COLUMNS["m"], this.asset_table_length, this.asset_table_width);
    const targetRange = this.sheet.getRange((targetRow + count), COLUMNS["m"], this.asset_table_length, this.asset_table_width);

    values.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL);
    this.sheet.getRange(targetRow, COLUMNS["m"], count, this.asset_table_width).clear();
  }

  /**
   * Update Reference month based on setting of diff month.
   */
  changeReferenceMonth() {
    const targetRow = this.getLastRow() - 1;
    const diffMonth = this.getCellValue(targetRow, COLUMNS["e"]);
    const referenceMonth = this.utils.getMonth(1 - diffMonth);
    
    this.setCellValue(targetRow, COLUMNS["b"], referenceMonth);
  }

  /**
   * Update display range of asset table.
   */
  changeGraphRange() {
    const lastRow = this.getLastRow(); // 終了月を指定するセルの行(最終行)
    const targetRow = lastRow - 1;     // 開始月を指定するセルの行

    let startMonth = this.getCellValue(targetRow, COLUMNS["b"]);
    let endMonth = this.getCellValue(lastRow, COLUMNS["b"]);

    // endMonthが現在日付より大きい場合、endMonthを強制的に現在月にする
    const end = this.utils.convertYYYYMMtoDate(endMonth);

    if (end > new Date()) {
      endMonth = this.utils.getMonth(MONTH["current"]);
    }
  
    const str_row = this.findRow(startMonth, COLUMNS["a"]);
    const end_row = this.findRow(endMonth, COLUMNS["a"]);

    // console.log(startMonth + ": line " + str_row);
    // console.log(endMonth + ": line " + end_row);

    // グラフの非表示を解除
    this.sheet.unhideRow(this.sheet.getRange(this.header_r, COLUMNS["a"], lastRow, 1));

    // 取得した開始月・終了月の範囲以外の行を隠す
    const range1rows = str_row - 2;

    if (range1rows > 0) {
      const range1 = this.sheet.getRange((this.header_r + 1), COLUMNS["a"], range1rows, 1);
      this.sheet.hideRow(range1);
    }

    const range2rows = this.getLatestRow() - end_row;

    if (range2rows > 0) {
      const range2 = this.sheet.getRange((end_row + 1), COLUMNS["a"], range2rows, 1);
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
}