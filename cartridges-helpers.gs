/** @OnlyCurrentDoc */

/** 
 * Takes Sheet ID and returns the sheet
 * @param {number} sheetId - ID of the sheet to return
 * @return Sheet with given ID
 */
function getSheetById(sheetId) {
  return SpreadsheetApp.getActive().getSheets().filter(
    function(s) {return s.getSheetId() === sheetId;}
  )[0];
}

/**
 * Generates and returns unique cartridge number
 * Format: YMMDDnn
 * @return newID
 */
function getUniqueId() {
  var date = Utilities.formatDate(new Date(), "GMT+3", "yyMMdd").split('').splice(1).join('');
  var newId = date + "01";
  var statusSheet = getSheetById(448257713);
  var currentIds = statusSheet.getRange('A2:A').getValues().map(function(row) {return row[0]});
  
  while(currentIds.includes(newId)) {
    var counter = Number(newId.substr(-2));
    counter++;
    if (String(counter).length < 2) counter = '0' + counter;
    newId = date + counter;
  }
  
  return newId;
}

/**
 * Returns number of last non-empty row in the given column range
 * @param {Range} range, which is a colunm (ex: sheet.getRange('A:A'))
 * @return {integer} numOfRow
 */
function getLastFullRow(range) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var values = range.getValues();

  
  for(var row = values.length; row > 0; row --) {
    if(values[row-1][0] != '') return row;
  }
}
