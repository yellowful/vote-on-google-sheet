function setNewDay() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("吃飯提議");
  var dateCell = sheet.getRange("A2");
  var optionsRange = sheet.getRange("B2:H2");

  // 設定 A2 儲存格為今天日期
  var today = new Date();
  dateCell.setValue(today);

  // 清空 B2 到 H2 的選項
  optionsRange.clearContent();
  pickRandomOption();
}

function pickRandomOption() {
  var proposalsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('吃飯提議');
  var optionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('選項');
  var selectedOptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('選過暫存');

  if (!selectedOptionsSheet) {
    // 如果「選過暫存」工作表不存在，則創建它並加入欄位標題
    selectedOptionsSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('選過暫存');
    selectedOptionsSheet.appendRow(['選過的餐廳']);
  }

  var lastOptionRow = optionsSheet.getLastRow();
  var lastSelectedRow = selectedOptionsSheet.getLastRow();

  // 檢查選過暫存表格是否有餐廳記錄
  var selectedOptions = selectedOptionsSheet.getRange('A2:A' + lastSelectedRow).getValues();
  var selectedOptionValues = selectedOptions.flat(); // 將二維陣列轉換為一維陣列

  // 如果所有選項都已經選過，則清空「選過暫存」表格
  if (selectedOptionValues.length === lastOptionRow - 1) {
    selectedOptionsSheet.getRange('A2:A' + lastSelectedRow).clearContent();
    selectedOptionValues = [];
  }

  // 隨機選擇一個餐廳，確保不在已選擇的餐廳中
  var randomIndex;
  var selectedOption;
  do {
    randomIndex = Math.floor(Math.random() * (lastOptionRow - 1)) + 2;
    selectedOption = optionsSheet.getRange('A' + randomIndex).getValue();
  } while (selectedOptionValues.indexOf(selectedOption) !== -1);

  // 更新吃飯提議的I欄
  proposalsSheet.getRange('J2').setValue(selectedOption);

  // 將選過的餐廳記錄在選過暫存表格中
  selectedOptionsSheet.getRange(lastSelectedRow + 1, 1).setValue(selectedOption);
}
