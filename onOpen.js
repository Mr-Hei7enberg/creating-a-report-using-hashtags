function onOpen(e) {
  SpreadsheetApp.getUi()
  .createMenu("Отчет")
  .addItem("Отчет по кол-ву квартир", "otchet_kvartir")
  .addItem("Отчет по обходам", "otchet_obhod")
  .addToUi();
}
