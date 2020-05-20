/** @OnlyCurrentDoc */

/**
 * Создаем меню с кнопками в верхней панели документа
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Actions')
    .addItem('Add order', 'takeOrder')
    .addToUi();
}

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

function registerOrder(shop, item, qty, price) {
  ordersSheet = getSheetById(2129613957);
  
  ordersSheet.insertRowBefore(2);
  var entry = ordersSheet.getRange('A2:E2');
  entry.setValues([[shop, item, qty, price, new Date()]])
}

function takeOrder() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  var message = ui.prompt("Input client's order").getResponseText().split(' ');
  
  var shop = message[0];
  var itemId = message[1];
  var qty = message[2];
  
  var productsSheet = getSheetById(1892121132);
  var products = productsSheet.getDataRange().getValues().slice(1);
  
  var product = products.filter(function(row) {return row[0] == itemId})[0];
  var response = ui.alert(product[1] + " (" + qty + product[2] + ") at " + product[4] + "NAD each \n Your total is: " + qty * product[4] + "NAD", ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    registerOrder(shop, itemId, qty, product[4]);
    ss.toast("your order is registered successfully");
  }
  
  
}
