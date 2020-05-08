/** @OnlyCurrentDoc */

/**
 * If not on register sheet - moves user there
 * If on register sheet - exports all entered cartriges to history
 * and clears the regiter sheet
 */
function registerToBeFIlled() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var registerSheetId = 1402392418
  var registerSheet = getSheetById(registerSheetId);
  var cartNumCol = registerSheet.getRange('A5:A');
  
  var activeSheetId = SpreadsheetApp.getActiveSheet().getSheetId();
  
  if(registerSheetId === activeSheetId) {
    var filledRows = getLastFullRow(cartNumCol);
    if (!filledRows) {ss.toast("введите номера картриджей в первую колонку"); return};
    
    var lastColumn = registerSheet.getLastColumn();

    var enteredData = registerSheet.getRange(5, 1, filledRows, lastColumn).getValues();
    
    var base = registerSheet.getRange('B2').getValue();
    if (!base) {ss.toast("введите базу отправки/получения"); return};
    
    // Добавляем в массив недостающие для истории данные
    var logData = enteredData.map(function(row) {
      row.splice(3, 0, base);
      row.splice(4, 1, "Получен на заправку");
      row.splice(6, 0, new Date());
      return row;
    });
    
    // Записываем в историю
    log(logData);
  
    // Очищаем данные таблицы
    registerSheet.getRange('A5:A').clear();
    registerSheet.getRange('E5:F').clear();
    registerSheet.getRange('B2').clear();

    } else {
      registerSheet.activate();
      ss.toast("Введите номера картриджей");
  }
}

/**
 * If not on register sheet - moves user to it
 * If doesn't have base notifies user and returns
 * If has no values - adds cartridges from Status sheet
 * If has values - logs and clens the sheet
 */
function sendToBeFilled() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var registerSheetId = 400518586
  var registerSheet = getSheetById(registerSheetId);
  var base = registerSheet.getRange('B2').getValue();
  var cartNumCol = registerSheet.getRange('A5:A');
  
  // Если не на том листе или не заполнена база - премещает пользователя на лист
  var activeSheetId = SpreadsheetApp.getActiveSheet().getSheetId();
  if (registerSheetId !== activeSheetId || !base) {
      registerSheet.activate();
      ss.toast('выберите базу отправки и вызовите функцию снова');
      return;
  }
  
  // Если нет картриджей в списке отправки - наполняем согласно базы и статуса
  var filledRows = getLastFullRow(cartNumCol);
  if (!filledRows) {
    var statusSheet = getSheetById(448257713);
    var carts = statusSheet.getRange('A2:G').getValues();
    
    var filteredCarts = carts.filter(function(row) {
      return (row[4] === "Получен на заправку" && 
        row[3] === base);
    }).map(function(row) {
      row = row.splice(0,3);
      row.length = 7;
      return row;
    });

    if(filteredCarts.length > 1) registerSheet.insertRows(5, filteredCarts.length - 1);
    registerSheet.getRange(5, 1, filteredCarts.length, registerSheet.getLastColumn()).setValues(filteredCarts);
    
    ss.toast('Проверьте правильность данных, когда будет готово - распечатайте и вызовите функцию снова');
  }
  
  // Теперь, когда картриджи введены спрашиваем, распечатана ли ведомость?
  if (filledRows) {
    var printed = ui.alert('Ведомость распечатана?', ui.ButtonSet.YES_NO);
    if (printed == "NO") return;
    
    // Если данные есть и ведомость распечатана - логируем и отчищаем
    var rawData = registerSheet.getRange('A5:G').getValues();
    var data = rawData.map(function(row) {
      row.length = 8;
      row[3] = base;
      row[4] = "Отправлен на заправку";
      row[6] = new Date();
      return row;
    });
    
    log(data);
    registerSheet.getRange('B2').clearContent();
    registerSheet.getRange('A5:G5').clearContent();
    var lastRow = registerSheet.getLastRow();
    if (lastRow >5) registerSheet.deleteRows(6, lastRow - 5);
  }
}



/**
 * If not on register sheet - moves user to it
 * If doesn't have base notifies user and returns
 * If has no values - adds cartridges from Status sheet
 * If has values - logs and clens the sheet
 */
function registerFilled() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var registerSheetId = 1735472467;
  var registerSheet = getSheetById(registerSheetId);
  var base = registerSheet.getRange('B2').getValue();
  var cartNumCol = registerSheet.getRange('A5:A');
  
  // Если не на том листе или не заполнена база - премещает пользователя на лист
  var activeSheetId = SpreadsheetApp.getActiveSheet().getSheetId();
  if (registerSheetId !== activeSheetId || !base) {
      registerSheet.activate();
      ss.toast('выберите базу отправки и вызовите функцию снова');
      return;
  }
  
  // Если нет картриджей в списке приема - наполняем согласно базы и статуса
  var filledRows = getLastFullRow(cartNumCol);
  if (!filledRows) {
    var statusSheet = getSheetById(448257713);
    var carts = statusSheet.getRange('A2:G').getValues();
    
    var filteredCarts = carts.filter(function(row) {
      return (row[4] === "Отправлен на заправку" && 
        row[3] === base);
    }).map(function(row) {
      row = row.splice(0,3);
      row.length = 7;
      row[3] = 'FALSE';
      row[4] = 'FALSE';
      row[5] = 'FALSE';
      return row;
    });

    if(filteredCarts.length > 1) registerSheet.insertRows(5, filteredCarts.length - 1);
    registerSheet.getRange(5, 1, filteredCarts.length, registerSheet.getLastColumn()).setValues(filteredCarts);
    
    ss.toast('Отметьте все операции и, когда будет готово - вызовите функцию снова');
  }
  
  if (filledRows) {
  
    var printed = ui.alert('Ведомость заполнена?', ui.ButtonSet.YES_NO);
    if (printed == "NO") return;
    
    // Если данные есть и ведомость заполнена - логируем и отчищаем
    var rawData = registerSheet.getRange('A5:G').getValues();
    var data = [];
    var date = new Date();
    for (row of rawData) {
      var statuses = [];
      if (row[5] === true) statuses.push("списан");
      statuses.push("Получен из заправки");
      if (row[3] === true) statuses.push("заправка");
      if (row[4] === true) statuses.push("восстановление");
   
      for (status of statuses) {
        data.push([
          row[0],
          row[1],
          row[2],
          base,
          status,
          ,
          date,
          row[6]        
      ])}
    }
    
    log(data);
    registerSheet.getRange('B2').clearContent();
    registerSheet.getRange('A5:G5').clearContent();
    var lastRow = registerSheet.getLastRow();
    if (lastRow >5) registerSheet.deleteRows(6, lastRow - 5);
  
  }
}

/**
 * If not on register sheet - moves user to it
 * If doesn't have base notifies user and returns
 * If has no values - adds cartridges from Status sheet
 * If has values - logs and clens the sheet
 */
function sendToClient() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var registerSheetId = 1554849670
  var registerSheet = getSheetById(registerSheetId);
  var base = registerSheet.getRange('B2').getValue();
  var cartNumCol = registerSheet.getRange('A5:A');
  
  // Если не на том листе или не заполнена база - премещает пользователя на лист
  var activeSheetId = SpreadsheetApp.getActiveSheet().getSheetId();
  if (registerSheetId !== activeSheetId || !base) {
      registerSheet.activate();
      ss.toast('выберите базу отправки и вызовите функцию снова');
      return;
  }
  
  // Если нет картриджей в списке отправки - наполняем согласно базы и статуса
  var filledRows = getLastFullRow(cartNumCol);
  if (!filledRows) {
    var statusSheet = getSheetById(448257713);
    var carts = statusSheet.getRange('A2:G').getValues();
    
    var filteredCarts = carts.filter(function(row) {
      return (row[4] === "Получен из заправки" && 
        row[3] === base);
    }).map(function(row) {
      row.splice(3,3);
      return row;
    });

    if(filteredCarts.length > 1) registerSheet.insertRows(5, filteredCarts.length - 1);
    registerSheet.getRange(5, 1, filteredCarts.length, registerSheet.getLastColumn()).setValues(filteredCarts);
    
    ss.toast('Проверьте правильность данных, когда будет готово - распечатайте и вызовите функцию снова');
  }
  
  // Теперь, когда картриджи введены спрашиваем, распечатана ли ведомость?
  if (filledRows) {
    var printed = ui.alert('Ведомость распечатана?', ui.ButtonSet.YES_NO);
    if (printed == "NO") return;
    
    // Если данные есть и ведомость распечатана - логируем и отчищаем
    var rawData = registerSheet.getRange('A5:G').getValues();
    var data = rawData.map(function(row) {
      row.length = 8;
      row[7] = row[3];
      row[3] = base;
      row[4] = "Отправлен на отделение";
      row[6] = new Date();
      return row;
    });
    
    log(data);
    registerSheet.getRange('B2').clearContent();
    registerSheet.getRange('A5:D5').clearContent();
    var lastRow = registerSheet.getLastRow();
    if (lastRow >5) registerSheet.deleteRows(6, lastRow - 5);
  }
}

/**
 * Prompts user to enter division and cartridge name
 * appends new one to the list in status page
 * sorst status page
 */
function addNewCartridge() {
  var ui = SpreadsheetApp.getUi();
  var statusSheetId = 448257713;
  var statusSheet = getSheetById(statusSheetId);
  
  var model = ui.prompt("Введите модель нового картриджа \nнапирмер: 'TN-22'").getResponseText();
  var division = ui.prompt("Введите адресс отделения, \nнапример: 'Малиновского, 70' \nили 'Николаев - Московская, 39'").getResponseText();
  var Id = getUniqueId();
  
  statusSheet.insertRows(2,1);
  
  statusSheet.getRange('A2:G2').setValues([[
    Id, 
    model, 
    division,
    '=iferror(vlookup(A2;\'История\'!A:G;4;FALSE); "-")',
    '=iferror(vlookup(A2;\'История\'!A:G;5;FALSE); "-")',
    '=iferror(vlookup(A2;\'История\'!A:G;7;FALSE); "-")',
    '=iferror(vlookup(A2;\'История\'!A:H;8;FALSE); "-")',
    ]]);
  
  log([[Id, model, division, "-", "новый", "", new Date(), "закуплен новый"]]);
  
  statusSheet.getRange('A2:G').sort(3);
}

/**
 * If cartridge number is selected - inputs new unique number
 * If elsewhere - redirects to status page and alerts user
 */
function generateNumber() {
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

function addOldCartridge() {
  var ui = SpreadsheetApp.getUi();
  var statusSheetId = 448257713;
  var statusSheet = getSheetById(statusSheetId);
  
  var model = ui.prompt("Введите модель нового картриджа \nнапирмер: 'TN-22'").getResponseText();
  var division = ui.prompt("Введите адресс отделения, \nнапример: 'Малиновского, 70' \nили 'Николаев - Московская, 39'").getResponseText();
  var Id = getUniqueId();
  
  statusSheet.insertRows(2,1);
  
  statusSheet.getRange('A2:G2').setValues([[
    Id, 
    model, 
    division,
    '=iferror(vlookup(A2;\'История\'!A:G;4;FALSE); "-")',
    '=iferror(vlookup(A2;\'История\'!A:G;5;FALSE); "-")',
    '=iferror(vlookup(A2;\'История\'!A:G;7;FALSE); "-")',
    '=iferror(vlookup(A2;\'История\'!A:H;8;FALSE); "-")',
    ]]);
  
  log([[Id, model, division, "-", "", "", new Date(), "присвоили номер старому картриджу"]]);
  
  statusSheet.getRange('A2:G').sort(3);
}

function generateReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var reportSheetId = 494447641;
  var reportSheet = getSheetById(reportSheetId);
  
  reportSheet.activate();
  
  var startDate = reportSheet.getRange('B1').getValue();
  var endDate = reportSheet.getRange('B2').getValue();
  var cartNumCol = reportSheet.getRange('A6:A');
  
  var filledRows = getLastFullRow(cartNumCol);
  if (!filledRows) {
    var logSheet = getSheetById(743789502);
    var carts = logSheet.getRange('A2:H').getValues();
    
    var filteredCarts = carts.filter(function(row) {
      return ((row[4] === "заправка" || row[4] === "восстановление") && (row[6] <= endDate && row[6] >= startDate));
    }).map(function(row) {
      var newRow = [];
      newRow[0] = row[6];
      newRow[1] = row[2];
      newRow[2] = row[1];
      newRow[3] = row[4];
      newRow[4] = 'formula';
      
      return newRow;
    });

    reportSheet.insertRows(6, filteredCarts.length - 1);
    reportSheet.getRange(6, 1, filteredCarts.length, reportSheet.getLastColumn()).setValues(filteredCarts);
  }
}

function logNoNumber() {
  var ui = SpreadsheetApp.getUi();

  var model = ui.prompt("Введите модель нового картриджа \nнапирмер: 'TN-22'").getResponseText();
  var division = ui.prompt("Введите адресс отделения, \nнапример: 'Малиновского, 70' \nили 'Николаев - Московская, 39'").getResponseText();
  
  var fill = ui.alert("Была заправка?", ui.ButtonSet.YES_NO);
  if (fill) log([["неизвестен", model, division, "-", "заправка", "", new Date(), "работа подрядчика на отделении"]]);
  
  var recover = ui.alert('Было восстановление?', ui.ButtonSet.YES_NO);
  if (recover) log([["неизвестен", model, division, "-", "восстановление", "", new Date(), "работа подрядчика на отделении"]]);
}
