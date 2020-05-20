function onOpen() {
  
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Смарт кнопки')
      .addItem('Создать новую закупку', 'newRequest')
      .addItem('Создать содержание', 'createIndex')
      .addItem('Обновить содержание', 'updateIndex')
      .addToUi();
 
}

// =======================
// function to create the index
function createIndex() {
  
  // Get all the different sheet IDs
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  var namesArray = sheetNamesIds(sheets);
  
  var indexSheetNames = namesArray[0];
  var indexSheetIds = namesArray[1];
  
  // check if sheet called sheet called already exists
  // if no index sheet exists, create one
  if (ss.getSheetByName('содержание') == null) {
    
    var indexSheet = ss.insertSheet('содержание',0);
    
  }
  // if sheet called index does exist, prompt user for a different name or option to cancel
  else {
    
    var indexNewName = Browser.inputBox('The name Index is already being used, please choose a different name:', 'Please choose another name', Browser.Buttons.OK_CANCEL);
    
    if (indexNewName != 'cancel') {
      var indexSheet = ss.insertSheet(indexNewName,0);
    }
    else {
      Browser.msgBox('No index sheet created');
    }
    
  }
  
  // add sheet title, sheet names and hyperlink formulas
  if (indexSheet) {
    
    printIndex(indexSheet,indexSheetNames,indexSheetIds);

  }
    
}


// =======================
// function to update the index, assumes index is the first sheet in the workbook
function updateIndex() {
  
  // Get all the different sheet IDs
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var indexSheet = sheets[0];
  
  var namesArray = sheetNamesIds(sheets);
  
  var indexSheetNames = namesArray[0];
  var indexSheetIds = namesArray[1];
  
  printIndex(indexSheet,indexSheetNames,indexSheetIds);
}


// =======================
// function to copy a template
function newRequest() {

  
  var name = "labnol";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('шаблон').copyTo(ss);
  
  /* Before cloning the sheet, delete any previous copy */
  var old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old); // or old.setName(new Name);
  
  /* Make up the name of the sheet */
  var current = new Date();
  var year = current.getFullYear().toString().substr(-2);
  var month = ((current.getMonth()+1).toString().length > 1) ? '' : '0';
  month += (current.getMonth() + 1).toString();
  var requestName = "" + year + month + "01";
  
  counter = 1
  while (ss.getSheetByName(requestName)) {
    counter +=1
    if (counter <=9) {
      requestName = requestName.substring(0, requestName.length - 2) + "0" + counter;
  } else {
      requestName = requestName.substring(0, requestName.length - 2) + counter;
    }
  }

  
  SpreadsheetApp.flush(); // Utilities.sleep(2000);
  sheet.setName(requestName);

  /* Make the new sheet active */
  ss.setActiveSheet(sheet);
  
  /* Move sheet to first position */
  ss.moveActiveSheet(4)

  /* Fill in the values in cells */
  ss.getRange('E1').setValue(current);
  ss.getRange('E2').setValue(requestName);
  
  /* Create folder */
  var parentFolder=DriveApp.getFolderById("1m3OCnwO0mtjzi1_WHO-U-zco8cmh76Sm");
  var newFolder=parentFolder.createFolder(requestName);

}


// function to print out the index
function printIndex(sheet,names,formulas) {
  
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('содержание').getRange("A:B").clearContent();  
  
//  sheet.clearContents();
  
  sheet.getRange(1,2).setValue('Содержание').setFontWeight('bold');
//  sheet.getRange(3,1,names.length,1).setValues(names);
  sheet.getRange(3,2,formulas.length,1).setFormulas(formulas);
  
}


// function to create array of sheet names and sheet ids
function sheetNamesIds(sheets) {
  
  var indexSheetNames = [];
  var indexSheetIds = [];
  
  // create array of sheet names and sheet gids
  sheets.forEach(function(sheet){
    indexSheetNames.push([sheet.getSheetName()]);
    indexSheetIds.push(['=hyperlink("#gid=' 
                        + sheet.getSheetId() 
                        + '";"' 
                        + sheet.getSheetName() 
                        + '")']);
    
  });
  
  return [indexSheetNames, indexSheetIds];
  
}


function onEdit(e) 
{
  var editRange = { // D5
    top : 5,
    bottom : 5,
    left : 4,
    right : 4
  };

  // Exit if we're out of range
  var thisRow = e.range.getRow();
  if (thisRow < editRange.top || thisRow > editRange.bottom) return;

  var thisCol = e.range.getColumn();
  if (thisCol < editRange.left || thisCol > editRange.right) return;

  // We're in range; timestamp the edit
  var ss = e.range.getSheet();
  var state = ss.getRange("D5").getValue();
  
  switch (state) {
      
    case '6. Распределено':
    ss.setTabColor("000");      // black
    break;

    case '5. Получено':
    ss.setTabColor("008000");   // green
    break;

    case '4. Ждем доставки':
    ss.setTabColor("FFD700");      // gold
    break;      
      
    case '3. Ждем оплаты':
    ss.setTabColor("808000");   // olive
    break;
      
    case '2. На утверждении':
    ss.setTabColor("f00");      // red
    break;
      
    case '1. Ждем счет':
    ss.setTabColor("00f");      // blue
    break;      
      
  default:
    ss.setTabColor("fff");
  }

} 