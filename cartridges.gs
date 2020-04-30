var ui = SpreadsheetApp.getUi();

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = ss.getSheets();

var statusSheetName = "Картриджи";
var statusSheet = ss.getSheetByName(statusSheetName);
var statusColumn = 4;

var logSheetName = "История";
var logSheet = ss.getSheetByName(logSheetName);

function onOpen() {

  ui.createMenu('Смарт кнопки')
      .addItem('Новый картридж', 'addCartridge')
      .addSeparator()
      .addItem('Генерировать отчет', 'generateReport')
      .addToUi();
}

// Записывает на отдельный лист строку при изменении статуса
function onEdit(e){
  
  var range = e.range;
  
  // Проверяем, является ли это колонкой статуса
  if(range.getSheet().getSheetName() !== statusSheetName
    || range.getColumn() !== statusColumn) {
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
  rangeRow.getCell(1, 6).setValue("FALSE");
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


function generateReport() {
  var contents = logSheet.getRange("A:I").getValues();
  var sheet = ss.getActiveSheet();
  
  var dateFrom = sheet.getRange('B1').getValue();
  var dateTo = sheet.getRange('B2').getValue();
  
  // Очищаем текущий отчет
  sheet.deleteRows(6, sheet.getLastRow()-5);

  // Список заголовков таблицы
  sheet.appendRow(["Дата", "Отделение", "Оборудование", "Услуга", "Стоимость"]);

  var rowCount = 7;

  for (var row of contents) {  
    // Если не в диапозоне дат - не берем в расчет
    if (row[4] < dateFrom || row[4] > dateTo) continue;
  
    // Если не "получен от подрядчика" - не берем в расчет
    if (row[3] !== 'получен от подрядчика') continue;
  
    // Список услуг
    var services = {
      'заправка': row[5], // заправка
      'восстановление': row[6], // восстановление
      'ремонт': row[7], // ремонт
      'новый': row[0] === '_новый' // новый
    };
    

    
    for (var service in services) {
       if(services[service] === true) {
        // Список искомых значений
        data = [
          row[4], // Дата
          row[2], // Отделение
          row[1], // Оборудование
          service, // Услуга
          '=iferror(vlookup(C' + rowCount + ' & ", " & D' + rowCount + '; \'Прайс\'!C:D; 2; false); "цена не найдена")', // Стоимость
        ];
    
        sheet.appendRow(data);
        rowCount ++;
      }
    }
  }
  
  sheet.appendRow(["Итого:",,,,"=sum(E7:E" + sheet.getLastRow() + ")"]);
  
}

function addCartridge() {
  var data = [
    "",
    ui.prompt('Введите модель картриджа. Например: Brother TN-1075').getResponseText(),
    ui.prompt('Введите отделение, на которое направлен картридж').getResponseText(),
    'получен от подрядчика',
    new Date(),
  ];
  
  statusSheet.appendRow(data);
  data[0] = '_новый';
  logSheet.appendRow(data);
  
  // ui.alert(Session.getActiveUser().getEmail());
}
