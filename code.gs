function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Tiller AI AutoCat')
      .addItem('Run AutoCat', 'categorizeUncategorizedTransactions')
      .addItem('Search Transactions (Active Cell)', 'searchFromActiveCell')
      .addToUi();

  // Hook for menus that exist only in your own copy of the sheet.  Apps Script
  // allows one onOpen per project, so local additions have to come in through
  // here.  Define addLocalMenus(ui) in a file of your own.
  if (typeof addLocalMenus === 'function') {
    addLocalMenus(ui);
  }
}
