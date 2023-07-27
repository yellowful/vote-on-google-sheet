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

  var lastOptionRow = optionsSheet.getLastRow();
  
  // 隨機選擇一個餐廳
  var randomIndex = Math.floor(Math.random() * (lastOptionRow - 1)) + 2;
  var selectedOption = optionsSheet.getRange('A' + randomIndex).getValue();
  
  // 更新吃飯提議的I欄
  proposalsSheet.getRange('J2').setValue(selectedOption);
}
