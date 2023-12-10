function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Tools')
      .addItem('Run AI Cleanup', 'categorizeUncategorizedTransactions')
      .addToUi();
}
