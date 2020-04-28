var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = ss.getSheets();

var statusSheetName = "Картриджи";
var statusSheet = ss.getSheetByName(statusSheetName);
var statusColumn = 4;

var logSheetName = "История";
var logSheet = ss.getSheetByName(logSheetName);

// Записывает на отдельный лист строку при изменении статуса
function onEdit(e){
  
  var range = e.range;
  
  // Проверяем, является ли это колонкой статуса
  if(range.getSheet().getSheetName() !== statusSheetName 
    && range.getColumn() !== statusColumn) {
      return;
  }
  
  var rangeRowNumber = range.getRow();
  var rangeRow = statusSheet.getRange(rangeRowNumber, 1, 1, statusSheet.getMaxColumns())
  
  // ставит сегодняшнюю дату
  statusSheet.getRange(rangeRowNumber, statusColumn + 1).setValue(new Date());
 
  // записывает в историю
  var logLastRowNumber = getLastCell(logSheet, "A:C");
  logSheet.insertRowAfter(logLastRowNumber);
  rangeRow.copyTo(logSheet.getRange(logLastRowNumber + 1, 1));
  
  // чистит исходный диапозон
  rangeRow.getCell(1, 7).setValue("FALSE");
  
}

// Ищет и возвращает последнюю ячейку в данном диапозоне.
function getLastCell(sheet, rangeA1) {
  var values = sheet.getRange(rangeA1).getValues();
  
  var maxIndex = values.reduce(function(maxIndex, row, index) {
    return row[0] + row[1] + row[2] === "" ? maxIndex : index;
  }, 0);
  
  return maxIndex + 1;
}