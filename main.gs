/**
 * Create custom menu.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GAS実行')
    .addItem('シートを再表示', 'showHiddenSheet')
    .addToUi();
}

/**
 * Show hidden sheet.
 */
function showHiddenSheet() {
}

/**
 * Run on the 25th of every month (by trigger).
 */
function createNewSheet() {
  new Monthly().createNewSheet();
}

/**
 * Change the display range of asset table.
 * NOTE: DO NOT DELETE THIS METHOD. This method is used by GUI button of SpreadSheet.
 */
function changeGraphRange() {
  new CashFlow().changeGraphRange();
}

/**
 * Run every day at pm 11:55 (by trigger).
 */
function endOfMonth() {
  setEndTrigger();

  const monthly = new Monthly();
  const cashflow = new CashFlow();

  monthly.hidePrevious();
  monthly.protectCurrent();
  cashflow.setCurrentAsset();
  cashflow.addNextRow();
  cashflow.moveSummaryTable(1);
} 

/**
 * Run every day at am 1:00 ~ 2:00 (by trigger).
 */
function startOfMonth() {
  setStartTrigger();
  
  const cashflow = new CashFlow();
  cashflow.changeReferenceMonth();
  cashflow.changeGraphRange();
}

/* ------------------------------------- */

const setEndTrigger = () => {
  const nextMonth = getNextMonthNum();
  const lastDayOfNextMonth = getLastDayOfNextMonth();

  new Trigger("endOfMonth")
    .setMonth(nextMonth)
    .setDate(lastDayOfNextMonth)
    .setHours(23)
    .setMinutes(00)
    .update();
}

const setStartTrigger = () => {
  const nextMonth = getNextMonthNum();

  new Trigger("startOfMonth")
    .setMonth(nextMonth)
    .setDate(1)
    .setHours(0)
    .setMinutes(10)
    .update();
}

const getNextMonthNum = () => {
  return (new utils).getFormattedMonth(1, "M");
};

const getLastDayOfNextMonth = () => {
  const today = new Date();

  const nextMonthFirstDay = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  const lastDayOfNextMonth = new Date(nextMonthFirstDay.getFullYear(), nextMonthFirstDay.getMonth() + 1, 0);
  
  return lastDayOfNextMonth.getDate();
}