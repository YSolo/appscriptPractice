function onOpen() {
  
  var ui = SpreadsheetApp.getUi();

ui.createMenu('Смарт кнопки')
    .addItem('Выгрузить файлы', 'listFilesInFolder')
    .addItem('Создать новое отделение', 'createNew')
    .addItem('Создать новый месяц', 'createNewMonths')
    .addItem('Предоставить доступ', 'giveAccess')
    .addItem('test','test')
    .addToUi();
}

function listFilesInFolder() {
// Функция очищает лист и наполняет его данными, указанными в data
  
  var folder = DriveApp.getFolderById("10KWbx1hopJq2sd-EdDvukGX5acp752bE");
  var contents = folder.getFiles();

  var file, data, sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();

  // Список заголовков таблицы
  sheet.appendRow(["Отделение","Изменен","Доступ", "Права", "Ссылка", "Листы", "Владелец"]);

  while (contents.hasNext()) {

    file = contents.next();
    
   
    // Список искомых значений
    data = [
      '=HYPERLINK("' + file.getUrl() + '";"' + file.getName() + '")',
      file.getLastUpdated(),
      file.getSharingAccess(),
      file.getSharingPermission(),
      file.getUrl(),
      SpreadsheetApp.open(file).getSheets().map(function(s) {return s.getSheetName()}).join(', '),
      file.getOwner().getEmail(),
    ];

    sheet.appendRow(data);
    
  }
};



function createNew() {
  // Функция создает новую таблицу на основаниии Шаблона и добавляет название и ссылку в содержание
  
  var source = SpreadsheetApp.getActiveSpreadsheet();
  // выбираем шаблон
  var sheet = source.getSheetByName('ШАБЛОН');
  
  // спрашивает у пользователя имя
  var fileName = SpreadsheetApp.getUi().prompt("Введите полное название отделения").getResponseText()
  if (!fileName) return;
  
  // создает копию документа в папке
  var destinationFolder = DriveApp.getFolderById("10KWbx1hopJq2sd-EdDvukGX5acp752bE");  
  var newFile = DriveApp.getFileById(source.getId()).makeCopy(fileName, destinationFolder);
  var newFileOpened = SpreadsheetApp.open(newFile)
  
  // удаляет все, кроме шаблона
  newFileOpened.deleteSheet(newFileOpened.getSheets()[3]);
  newFileOpened.deleteSheet(newFileOpened.getSheets()[0]);
  newFileOpened.deleteSheet(newFileOpened.getSheets()[0]);
  
  // и переименовывает лист
  newFileOpened.renameActiveSheet('Начало учета');
  
  // Добавляет название отделения и ссылку на файл в конец содержания
  var indexSheet = source.getSheetByName('Содержание')
  var values = indexSheet.getRange("A:A").getValues();
  var maxIndex = values.reduce(function(maxIndex, row, index) {
    return row[0] === "" ? maxIndex : index;
  }, 0);
  var cell = indexSheet.setActiveRange(indexSheet.getRange(maxIndex + 2, 1));
  cell.setValue('=HYPERLINK("' + newFileOpened.getUrl() + '";"' + fileName + '")');
  
  // Установка прав доступа
  newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

}

function createNewMonth(file, template, newMonthName) {
  // Открываем таблицу отделения
  var ss = SpreadsheetApp.openById(file.getId());
  
  // Записываем список названий листов
  var sheetNames = ss.getSheets().map(function(s) {return s.getSheetName()});
  
  
  // Проверяем, есть ли текущий месяц. Если нет - ничего не делаем
  if ( sheetNames.includes(newMonthName) || sheetNames.includes("ШАБЛОН") ) return;
  
  

  // Копируем шаблон 
  var newMonthSheet = template.copyTo(ss);
  // Делаем шаблон первым листом
  ss.setActiveSheet(newMonthSheet);
  ss.moveActiveSheet(0);
  // Переименовываем шаблон
  newMonthSheet.setName(newMonthName);
  // Копируем значения из прошлого месяца в шаблон
  ss.getSheets()[1].getRange('A:F').copyTo(newMonthSheet.getRange('A:F'));
}

function createNewMonths() {
  var ui = SpreadsheetApp.getUi();
  var folder = DriveApp.getFolderById("10KWbx1hopJq2sd-EdDvukGX5acp752bE");
  var contents = folder.getFiles();
  var template = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ШАБЛОН');
  var newMonthName = ui.prompt("Введите название листа нового месяца (например: 04.2020)").getResponseText();
  var file;

  while (contents.hasNext()) {
    file = contents.next();
    createNewMonth(file, template, newMonthName);
    Logger.log("done with: " + file.getName());
  }
}

function giveAccess() {
  var folder = DriveApp.getFolderById("10KWbx1hopJq2sd-EdDvukGX5acp752bE");
  var contents = folder.getFiles();

  var file, protections, data, sheet = SpreadsheetApp.getActiveSheet();
  

  var ui = SpreadsheetApp.getUi();
  while (contents.hasNext()) {
  
    file = contents.next();

    protections = SpreadsheetApp.open(file).getProtections(SpreadsheetApp.ProtectionType.RANGE)
    for (protection of protections) {
      try {
        protection.addEditor("rashod1.smartlab@gmail.com");
      } catch (e) {
        ui.alert(file.getName() + " " + e);
      }
    }
  }
}

function test() {

  var file = DriveApp.getFileById("1C-mQuPcZYXU5w_YC_qTeiGr6gwgMdeVx0UvTEnLCcH0");
  
  var ss = SpreadsheetApp.open(file)
  var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE)
  ss.toast(protections)
  for (protection of protections) {
    protection.removeEditors(protection.getEditors());

    protection.addEditor("rashod1.smartlab@gmail.com");
    
  }
  
}
