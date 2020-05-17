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
     
     var numbers = data.filter(function(row) {return Number.isInteger(+row[table.col - 1]) && row[table.col - 1] != ""});
     
     numbers = numbers.map(function(row) {return row[table.col - 1]});

     confirmedNumbers = confirmedNumbers.concat(numbers);
  }
  
  // Split to two separate rows
  var present = [];
  var notPresent = [];
  
  for (var row of targetData) {
  
    if(confirmedNumbers.indexOf(row[target.col - 1]) == -1) {
      notPresent.push(row);
    } else {
      present.push(row);
    }
  }
  
  // Clear and populate sheet with data
  var resultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Результат');
  resultSheet.clear();
  
  try {
    resultSheet.getRange('A1').setValue('Не найденные декларации:').setFontWeight(800);
    var notPresentRange = resultSheet.getRange(2, 1, notPresent.length, notPresent[0].length);
    notPresentRange.setValues(notPresent);
  } catch (error) {
    Logger.log(error);
  }
  
  try {
  resultSheet.getRange(notPresent.length + 3, 1).setValue('Найденные декларации:').setFontWeight(800);
  var presentRange = resultSheet.getRange(notPresent.length + 4, 1, present.length, present[0].length);
  presentRange.setValues(present);
  } catch (error) {
    Logger.log(error);
  }
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
