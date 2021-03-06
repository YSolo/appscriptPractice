function onOpen() {
  
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Смарт кнопки')
      .addItem('Сделать запись', 'makeBooking')
      .addSeparator()
      .addItem('Добавить врача', 'newDoctor')
      .addItem('test', 'test')
      .addToUi();
}

function newDoctor() {
  
  var ui = SpreadsheetApp.getUi();
  var source = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = source.getSheetByName("Данные");
  var informationSheet = source.getSheetByName("Информация");
  
  // ---ПОЛУЧАЕМ ДАННЫЕ ОТ ПОЛЬЗОВАТЕЛЯ---
  // запрашиваем у пользователя ФИО. Например: Иванов Иван Иванович
  var fullName = ui.prompt('ФИО Врача полностью:').getResponseText();
  
  // делаем список из ФИО для проверки
  var nameList = fullName.split(" ");
  
  // если нет трех слов - выкидываем сообщение об ошибке
  if (nameList.length !== 3) {
    ui.alert('Ожидается ФИО в формате "Иванов Иван Иванович"');
    return;
  };
  
  // сохраняем короткую форму ФИО. Например: Иванов И.И.
  var shortName = nameList[0]+" "+nameList[1][0]+"."+nameList[2][0]+"."
  
  
  // ---СОЗДАНИЕ ЛИСТА ВРАЧА---  
  // копируем четвертый лист как шаблон
  var newDoctorSheet = source.getSheets()[3].copyTo(source);

  newDoctorSheet.setName("Врач:"+shortName);
  
  // фокус на лист
  source.setActiveSheet(newDoctorSheet);
  
  // двигаем лист на четвертую позицию
  source.moveActiveSheet(4)
  
  // подписываем именем врача и чистим лист
  newDoctorSheet.getRange('C1').setValue(fullName);
  newDoctorSheet.getRange('D13:F19').clear({contentsOnly: true});
  newDoctorSheet.getRange('D10:E10').clear({contentsOnly: true});
  newDoctorSheet.getRange('B4').clear({contentsOnly: true});
  
  // ---ОБНОВЛЯЕМ ИНФОРМАЦИЮ В СПИСКЕ ВРАЧЕЙ---
  var $doctors = dataSheet.getRange('A:A');
  var $nextDoctor = dataSheet.getRange(getLastCellInRange($doctors)+1, 1);
  $nextDoctor.setValue('=hyperlink("#gid=' 
                        + newDoctorSheet.getSheetId() 
                        + '";"' 
                        + shortName
                        + '")'
  );
 
  // ---СОЗДАЕМ ВРАЧА В ПАПКЕ ВРАЧЕЙ И ДОБАВЛЯЕМ ССЫЛКУ НА ДОКУМЕНТ---
  var doctorFileUrl = createDoctorJournal(shortName);
  dataSheet.getRange($nextDoctor.getRow(), 2).setValue(doctorFileUrl);
  newDoctorSheet.getRange('F1').setValue(doctorFileUrl);
  
  // меняем имя врача в файле
  SpreadsheetApp.openByUrl(doctorFileUrl).getRange('B1').setValue(shortName);

  // ---ВСТАВЛЯЕМ ИМЯ ВРАЧА В ТАБЛИЦУ ИНФОРМАЦИИ---
  var lastRow = informationSheet.getLastRow();
  
  var fromRange = informationSheet.getRange(lastRow-1, 1, 2, informationSheet.getMaxColumns());
  var toRange = informationSheet.getRange(lastRow+1, 1, 2, informationSheet.getMaxColumns());
  
  fromRange.copyTo(toRange);
  
  informationSheet.hideRows(lastRow+1);
  
  informationSheet.getRange(lastRow+2, 1).setValue(shortName);
 
}


// принимает область и возвращает номер последней строки
function getLastCellInRange(lookupRange) {

  var values = lookupRange.getValues();
  var last = values.filter(String).length;
  
  return last;
  
}

function makeBooking() {
  var ui = SpreadsheetApp.getUi();
  var source = SpreadsheetApp.getActiveSpreadsheet();
  var journalSheet = source.getSheetByName('Журнал');
  
  var timeRange = source.getActiveRange();
  if (timeRange.getNumColumns() > 2) {
    ui.alert("1 или две ячейки должны быть выбраны в графике")
    return;
  };
  
  var sheet = source.getActiveSheet();
  
  var doctor = sheet.getRange(timeRange.getRowIndex(), 1).getValue();
  var date = sheet.getRange('B1').getValue();
  var place = sheet.getRange(timeRange.getRowIndex(), 5).getValue();
  
  var patient = ui.prompt('ФИО Пациента').getResponseText();
  var tel = ui.prompt('Номер телефона').getResponseText();
  var age = sheet.getRange('B4').getValue() || ui.prompt("Укажите возраст").getResponseText();
  var service = sheet.getRange('B3').getValue() || ui.prompt("Укажите вид исследования").getResponseText();
  
  var nextJournalRow = journalSheet.getLastRow()+1;
  journalSheet.getRange(nextJournalRow, 1).setValue(patient);
  journalSheet.getRange(nextJournalRow, 2).setValue(age);
  journalSheet.getRange(nextJournalRow, 3).setValue(tel);
  journalSheet.getRange(nextJournalRow, 4).setValue(service);
  
  journalSheet.getRange(nextJournalRow, 5).setValue(date);
  journalSheet.getRange(nextJournalRow, 6).setValue(doctor);
  journalSheet.getRange(nextJournalRow, 7).setValue("=E"+nextJournalRow+"&F"+nextJournalRow+"&round(J" + nextJournalRow + "*1440)");
  journalSheet.getRange(nextJournalRow, 8).setValue("=E"+nextJournalRow+"&F"+nextJournalRow+"&round(K" + nextJournalRow + "*1440)");
  journalSheet.getRange(nextJournalRow, 9).setValue(place);

  journalSheet.getRange(nextJournalRow, 10, 1, timeRange.getNumColumns()).setValues(timeRange.getValues());
  
}

function createDoctorJournal(doctorName) {
  var doctorsFolder = DriveApp.getFolderById("16JKW6GhvMGyu2FdIMtfCDpfmkJQBjtMp");
  var template = DriveApp.getFileById("1cSYOvNMF9VWudix7L-4lfJNe1p9s0Hy7S2NfADOAOv4");
  
  // скопировать шаблон и переименовать
  var doctorFile = template.makeCopy(doctorName, doctorsFolder);
  
  // устанавливаем права доступа на просмотр
  doctorFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  // возвращаем ссылку для врача
  return doctorFile.getUrl();
  
}

function goHome() {

  SpreadsheetApp.getActive().getSheetByName('Информация').activate();

}

function test() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const infoSheet = ss.getSheetByName("Информация");
  
  // Input fields
  const $date = infoSheet.getRange('B1');
  const $age = infoSheet.getRange('B4');
  const $analysis = infoSheet.getRange('B6:B9');
  
  // get values
  let date = $date.getValue();
  const age = $age.getValue() || 18;
  const analysis = $analysis.getValues().filter(function(row) {return row[0]}).map(function(row) {return row[0]});

  // get doctors sheets
  const doctorsSheets = ss.getSheets().filter(function(sheet) {return sheet.getSheetName().startsWith("Врач:")});
  
  // get doctors data
  const doctorsData = doctorsSheets.map(function(sheet) {
    return {
      name: sheet.getRange('C1').getValue(),
      minAge: sheet.getRange('D10').getValue(),
      maxAge: sheet.getRange('E10').getValue(),
      analysis: sheet.getRange('H2:H').getValues().filter(function(row) {return row[0]}).map(function(row) {return row[0]}),
      schedule: sheet.getRange('B13:F19').getValues().filter(function(row) {return row[2]})
    }
  });
  
  // Filter doctors who are qualified
  const filteredDoctorsData = doctorsData.filter(function(doctor) {
    if (age < doctor.minAge || age > doctor.maxAge) return false;

    for (const a of analysis) {
      if(!doctor.analysis.includes(a)) {
        return false
      }
    }
    
    return true;
  });
  
  // If no doctors for the job - let user know
  if(!filteredDoctorsData.length) ui.alert('В этот период нет врачей, которые делают выбранный набор исследований');
  
  const journalData = ss.getSheetByName("Журнал").getRange("A2:K").getValues();
  
  // Accumulate printable data for 5 days
  const printableData = [];
  
  const timeIntervals = [];
  for (let hours = 8; hours < 19; hours++) {
    for (let minutes = 0; minutes < 60; minutes += 5) {
      timeIntervals.push([hours,minutes]);
    }
  }
  
  for (let i = 0; i < 5; i++) {
    const day = date.getDay();
    
    for (let doctor of filteredDoctorsData) {
      if(doctor.schedule.map(function(row) {return row[0]})
        .includes(day)) {
          let schedule = doctor.schedule.find(function(row) {return row[0] == day});
        
          let visualSchedule = [];

          // Iterate over 5 min intervals starting from 8
          let scheduleStartTime = schedule[2];
          let scheduleEndTime = schedule[3];
  
          // Set schedule marks
          for (let interval of timeIntervals) {
          
            if (scheduleStartTime.getHours() <= interval[0] 
              && interval[0] < scheduleEndTime.getHours()) {
                
              mark = "O";   
            } else {
              mark = "X";
            }
            
            visualSchedule.push(mark);
          }
          
          // Set doctor Appointements marks
          const doctorAppointments = journalData.filter(function(row) {return (row[5] == doctor.name && row[4].getDate() == date.getDate())});
          
          if (doctorAppointments.length) {
            for(let app of doctorAppointments) {
              const startIndex = timeIntervals.findIndex(function(stamp) {return stamp[0] == app[7].getHours() && stamp[1] == app[7].getMinutes()});
              const intervals = app[8] / 5;
              const visual = ["A", ..."P".repeat(intervals - 1).split('')];
              
              visualSchedule.splice(startIndex, intervals, ...visual);
              
            }
          }
          
          printableData.push([new Date(date), doctor.name, schedule[4], ...visualSchedule]);
      }
    }
       
    // increment one day
    date.setDate(date.getDate() + 1);
  }
  
  // remove frozen pointer to be able to delete rows
  infoSheet.setFrozenRows(0);
  
  try {
    infoSheet.deleteRows(14, infoSheet.getMaxRows() - 13);
  } catch(e) {
  }
  
  
  infoSheet.insertRowsAfter(13, printableData.length);
  infoSheet.getRange(14, 1, printableData.length, 135).setValues(printableData).setBorder(false, false, true, true, true, true, 'white', SpreadsheetApp.BorderStyle.SOLID);
  infoSheet.getRange("A14:C").clearFormat();
  
  // Set conditional format rules
  const rules = [];
  const xRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("X")
    .setFontColor("white")
    .setRanges([infoSheet.getRange('D14:EG' + infoSheet.getMaxRows())])
    .build();
  rules.push(xRule);
  
  const oRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("O")
    .setBackground("#b6d7a8")
    .setFontColor("#b6d7a8")
    .setRanges([infoSheet.getRange('D14:EG' + infoSheet.getMaxRows())])
    .build();
  rules.push(oRule);
    
  const aRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("A")
    .setBackground("salmon")
    .setFontColor("salmon")
    .setRanges([infoSheet.getRange('D14:EG' + infoSheet.getMaxRows())])
    .build();
  rules.push(aRule);
  
  const pRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("P")
    .setBackground("lightsalmon")
    .setFontColor("lightsalmon")
    .setRanges([infoSheet.getRange('D14:EG' + infoSheet.getMaxRows())])
    .build();
  rules.push(pRule);  
  
  infoSheet.setConditionalFormatRules(rules);
  
  infoSheet.setFrozenRows(13);
  

}
