function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Tiller AI AutoCat')
      .addItem('Run AutoCat', 'categorizeUncategorizedTransactions')
      .addItem('Search Transactions (Active Cell)', 'searchFromActiveCell');

  // Hook for menu items that exist only in your own copy of the sheet.  Define
  // addLocalMenuItems(menu) in a file of your own; Apps Script only allows one
  // onOpen per project, so local additions have to come in through here.
  if (typeof addLocalMenuItems === 'function') {
    addLocalMenuItems(menu);
  }

  menu.addToUi();
}
