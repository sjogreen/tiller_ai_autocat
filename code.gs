function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Tiller AI AutoCat')
      .addItem('Run AutoCat', 'categorizeUncategorizedTransactions')
      .addToUi();
}
