/** @OnlyCurrentDoc */

/**
 * Создаем меню с кнопками в верхней панели документа
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Смарт кнопки')
      .addItem('Тест', 'test')
      .addSeparator()
      .addSubMenu(
        ui.createMenu('Перемещения')
          .addItem('Поступление на заправку','registerToBeFilled')
       )
      .addSubMenu(
        ui.createMenu('Добавить картридж')
          .addItem('Новый', 'addNewCartrige')
          .addItem('Старый', 'addOldCartrige')
      )
      .addToUi();
}
