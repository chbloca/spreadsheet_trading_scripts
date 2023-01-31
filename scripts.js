function SaveTradeResults() {
  var spreadsheet = SpreadsheetApp.getActive();
  var calculatorSheet = spreadsheet.getSheetByName('Calculator');

  let dateTime = new Date().toLocaleDateString();
  var base = calculatorSheet.getRange('B1').getValue();
  var quote = calculatorSheet.getRange('B2').getValue();
  var capital = calculatorSheet.getRange('B4').getValue();
  var risk = calculatorSheet.getRange('B5').getValue();
  var direction = calculatorSheet.getRange('B6').getValue();
  var pnl = calculatorSheet.getRange('F7').getValue();
  
  var diarySheet = spreadsheet.getSheetByName('Diary');
  var lastRow = diarySheet.getLastRow();

  var myArray = [dateTime, base + "/" + quote, capital, risk, direction, pnl];

  var i = 0;
  myArray.forEach(function (value, _) {
    diarySheet.getRange(lastRow + 1, i + 1).setValue(value)
    i++;
  });

  diarySheet.getRange(lastRow + 1, i + 1).setFormula("=F" + (lastRow + 1) + "+" + "C" + (lastRow + 1));
  i++;

  diarySheet.getRange(lastRow + 1, i + 1).setFormula("=(C" + (lastRow + 1) + "+" + "F" + (lastRow + 1) + ")" + "/" + "C" + (lastRow + 1) + "-1");
  i++;

  diarySheet.getRange(lastRow + 1, i + 1).setFormula("=text(" + "H" + (lastRow + 1) + "/" + "D" + (lastRow + 1) + ";\"0.00\")&\"R\"");
  i++;
};