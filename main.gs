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

  new Monthly().hidePrevious();
  new Monthly().protectCurrent();
  new CashFlow().setCurrentAsset();
  new CashFlow().addNextRow();
  new CashFlow().moveSummaryTable(1);
} 

/**
 * Run every day at am 1:00 ~ 2:00 (by trigger).
 */
function startOfMonth() {
  setStartTrigger();
  
  new CashFlow().changeReferenceMonth();
  new CashFlow().changeGraphRange();
}

/* ------------------------------------- */

function setEndTrigger() {
  const nextMonth = this.getNextMonthNum();
  const lastDayOfNextMonth = this.getLastDayOfNextMonth();

  new Trigger("endOfMonth")
    .setMonth(nextMonth)
    .setDate(lastDayOfNextMonth)
    .setHours(23)
    .setMinutes(00)
    .update();
}

function setStartTrigger() {
  const nextMonth = this.getNextMonthNum();

  new Trigger("startOfMonth")
    .setMonth(nextMonth)
    .setDate(1)
    .setHours(0)
    .setMinutes(10)
    .update();
}

function getNextMonthNum() {
  const today = new Date();
  const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);

  return nextMonth.getMonth() + 1;
}

function getLastDayOfNextMonth() {
  const today = new Date();
  
  const currentMonth = today.getMonth();
  const currentYear = today.getFullYear();
  
  const nextMonthFirstDay = new Date(currentYear, currentMonth + 1, 1);
  const lastDayOfNextMonth = new Date(nextMonthFirstDay.getFullYear(), nextMonthFirstDay.getMonth() + 1, 0);
  
  return lastDayOfNextMonth.getDate();
}