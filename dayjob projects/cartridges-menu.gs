/** @OnlyCurrentDoc */

/**
 * Создаем меню с кнопками в верхней панели документа
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Смарт кнопки')
      .addSubMenu(
        ui.createMenu('Добавить картридж')
          .addItem('Новый', 'addNewCartrige')
          .addItem('Старый', 'addOldCartrige')
      )
      .addSubMenu(
        ui.createMenu('Перемещения')
          .addItem('Поступление на заправку','registerToBeFIlled')
          .addItem('Отправка на заправку', 'sendToBeFilled')
          .addItem('Поступление после заправки', 'registerFilled')
          .addItem('Отправка на отделение', 'sendToClient')
      )
      .addItem('Отчет', 'generateReport')
      .addToUi();
}
