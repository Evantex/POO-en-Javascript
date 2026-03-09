function Borrartodo() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getCurrentCell().offset(-30, -2, 6, 7).activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getCurrentCell().offset(8, 0, 6, 7).activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getCurrentCell().offset(8, 0, 6, 7).activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};

function BorrarTodo() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRangeList([spreadsheet.getCurrentCell().offset(-27, -1, 2, 1).getA1Notation(),
  spreadsheet.getCurrentCell().offset(-27, 0, 6, 6).getA1Notation(),
  spreadsheet.getCurrentCell().offset(-19, -1, 2, 1).getA1Notation(),
  spreadsheet.getCurrentCell().offset(-19, 0, 6, 6).getA1Notation(),
  spreadsheet.getCurrentCell().offset(-11, -1, 2, 1).getA1Notation(),
  spreadsheet.getCurrentCell().offset(-11, 0, 6, 6).getA1Notation()]).activate()
  .clear({contentsOnly: true, skipFilteredRows: true});
};