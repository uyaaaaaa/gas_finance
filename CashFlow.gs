class CashFlow extends Sheet {
  /**
   * Constructor.
   */
  constructor() {
    super();
    this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CF管理");
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
        this.getNextMonth(),
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
    const table_width = this.utils.getCellDiff(str_c, COLUMNS["t"]);

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
    const refMonth = this.utils.getFormattedMonth(1 - diffMonth, this.default_format);
    
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
      const diff = this.utils.getCellDiff(str_r, end_r);
      this.setCellValue((this.getLastRow() - 1), COLUMNS["e"], diff);
    }

    // display all rows once.
    this.sheet.unhideRow(this.sheet.getRange(this.header_r, COLUMNS["a"], this.getLastRow()));

    const range1rows = str_r - 2;

    // hide the range before start month
    if (range1rows > 0) {
      const range1 = this.sheet.getRange((this.header_r + 1), COLUMNS["a"], range1rows);
      this.sheet.hideRow(range1);
    }

    const range2rows = this.getRowNumOfCurrentMonth() - end_r;

    // hide the range after end month
    if (range2rows > 0) {
      const range2 = this.sheet.getRange((end_r + 1), COLUMNS["a"], range2rows);
      this.sheet.hideRow(range2);
    }
  }

  /**
   * @return {number}
   */
  getRowNumOfCurrentMonth() {
    const currentMonth = this.getCurrentMonth();

    return this.findRow(currentMonth, COLUMNS["a"]);
  }
}