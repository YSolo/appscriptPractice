/** @OnlyCurrentDoc */

function test() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var ui = SpreadsheetApp.getUi();
//  var data = sheet.getRange(2, 1, sheet.getMaxRows(), sheet.getMaxColumns());

//  var operationsSheet = getSheetById(1402392418);
//  operationsSheet.activate();

  var activeRange = ss.getActiveRange();
  ss.toast(getLastFullRow(activeRange));

}

/**
 * If not on register sheet - moves user there
 * If on register sheet - exports all entered cartriges to history
 * and clears the regiter sheet
 */
function registerToBeFIlled() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var registerSheetId = 1402392418
  var registerSheet = getSheetById(registerSheetId);
  
  var activeSheetId = SpreadsheetApp.getActiveSheet().getSheetId();
  
  if(registerSheetId === activeSheetId) {
    var filledRows = getLastFullRow(registerSheet.getRange('A5:A'));
    var lastColumn = registerSheet.getLastColumn();

    var enteredData = registerSheet.getRange(5, 1, filledRows, lastColumn).getValues();
    var base = registerSheet.getRange('B2').getValue();
    
    // причесать данные
    
    // вызвать функцию логирования.
    
  
  
  
  
  

  } else {
    registerSheet.activate();
  }
}

/**
 * If cartridge number is selected - inputs new unique number
 * If elsewhere - redirects to status page and alerts user
 */
function addOldCartrige() {
  var ui = SpreadsheetApp.getUi();
  
  var statusSheetId = 448257713;
  var activeRange = SpreadsheetApp.getActiveRange();
  var statusSheet = getSheetById(statusSheetId);
  
  if(activeRange.getSheet().getSheetId() == statusSheetId
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
