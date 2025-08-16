/**
 * Run on the 25th of every month (by trigger).
 */
function createNewSheet() {
  new Monthly().createNewSheet();
}

/**
 * Run every day at pm 10:00 ~ 11:00 (by trigger).
 */
function endOfMonth() {
  const date = new Date();
  const today = date.getDate();
  date.setDate(today + 1);
  const nextDay = date.getDate();
  
  if (nextDay !== 1) {
    return;
  }

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
  const date = new Date();
  const today = date.getDate();

  if (today !== 1) {
    return;
  }
  
  new CashFlow().changeReferenceMonth();
  new CashFlow().changeGraphRange();
}

/**
 * Change the display range of asset table.
 * NOTE: DO NOT DELETE THIS METHOD. This method is used by GUI button of SpreadSheet.
 */
function changeGraphRange() {
  new CashFlow().changeGraphRange();
}
