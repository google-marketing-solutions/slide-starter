/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  loadConfiguration();
  const spreadsheet = SpreadsheetApp.getActive();
  const menuItems = [
    {
      name: 'Generate deck',
      functionName: 'createDeckFromDatasources',
    },
  ];
  spreadsheet.addMenu('Katalyst', menuItems);
}