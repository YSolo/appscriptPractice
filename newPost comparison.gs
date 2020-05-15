function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Смарт кнопки')
    .addItem('Сравнить', 'populateComparisonResult')
    .addToUi();
}

function populateComparisonResult() {
  var ui = SpreadsheetApp.getUi();

  // Receive data from data sheet
  var tables = getData()
  
  // split target table and rest of data
  var target = tables.splice(0, 1)[0];
  
  // get necesary target range
  var targetData = SpreadsheetApp.openByUrl(target.url).getRange(target.range).getValues();
  
  // clean targetData
  targetData = targetData.filter(function(row) {
    return Number.isInteger(row[0]);
  });
  
  // get numbers from all the tables
  var confirmedNumbers = [];
  
  for (var table of tables) {
     var tableSheet = SpreadsheetApp.openByUrl(table.url);
     if (table.sheet) {
       tableSheet = tableSheet.getSheetByName(table.sheet);
     }
     
     var data = tableSheet.getRange(table.range).getValues();
     
     var numbers = data.filter(function(row) {return Number.isInteger(row[table.col - 1])});
     numbers = numbers.map(function(row) {return row[table.col - 1]});
     
     confirmedNumbers = confirmedNumbers.concat(numbers);
  }
  
  var filteredTargetData = targetData.filter(function(row) {
    return (confirmedNumbers.indexOf(row[target.col - 1]) === -1);
  });
  
  
  // TODO - Populate into a sheet
  var resultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Результат');
  resultSheet.clear();
  var resultRange = resultSheet.getRange(1, 1, filteredTargetData.length, filteredTargetData[0].length);
  resultRange.setValues(filteredTargetData);
  ui.alert(filteredTargetData);
}

function getData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Данные');
  
  var givenData = sheet.getDataRange().getValues().splice(1);
  
  return givenData.map(function(row) {
    
    return {
      name: row[0],
      url: row[1],
      sheet: row[2],
      range: row[3],
      col: row[4],
    }
  
  });
}
