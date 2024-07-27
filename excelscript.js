function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Menu')
        .addItem('Calculate Time Differences', 'main')
        .addToUi();
  }
  
  function main() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getDataRange();
    var values = range.getValues();
  
    if (values.length < 2) return;
  
    var datetimeColumn = 1; 
    var minColumn = 2; 
    for (var i = 1; i < values.length; i++) {
      var currentDateTime = new Date(values[i][datetimeColumn]);
      var previousDateTime = new Date(values[i-1][datetimeColumn]);
      
      var diffMinutes = Math.round((currentDateTime - previousDateTime) / (1000 * 60));
      
      values[i-1][minColumn] = diffMinutes;
    }
  
    values[values.length - 1][minColumn] = "";
  
    range.setValues(values);
  }
  
  function onEdit(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
  
    if (range.getRow() != sheet.getLastRow() || range.getColumn() != 2) return;
  
    main();
  }