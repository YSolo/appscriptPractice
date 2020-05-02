/** @OnlyCurrentDoc */

function test() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var ui = SpreadsheetApp.getUi();
//  var data = sheet.getRange(2, 1, sheet.getMaxRows(), sheet.getMaxColumns());

//  var operationsSheet = getSheetById(1402392418);
//  operationsSheet.activate();

}

function registerToBeFilled() {
  var sheet = getSheetById(1402392418);
  sheet.activate();
}

/**
 * If cartridge number is selected - inputs new unique number
 * If elsewhere - redirects to status page and alerts user
 */
function addOldCartrige() {
  var ui = SpreadsheetApp.getUi();
  
  var activeRange = SpreadsheetApp.getActiveRange();
  var statusSheet = getSheetById(448257713);
  
  if(activeRange.getSheet().getSheetId() == statusSheet.getSheetId() 
    && activeRange.getColumn() === 1
    && activeRange.getRow() !== 1) {
    activeRange.setValue(getUniqueId());
  } else {
    ui.alert("Чтобы присвоить номер, \n"
      + "- картридж нужно найти на странице Картриджи,\n"
      + "- поставить курсор в соответствующую ему ячейку столбца 'номер'\n"
      + "- и снова вызвать функцию 'Добавить картридж'");
    statusSheet.activate();
    return;
  }
}
