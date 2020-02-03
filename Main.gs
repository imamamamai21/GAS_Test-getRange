var MY_SHEET_ID = '1U7i--WWg6q4XOjNSYCiqmqpsxScIDuzlWGVqFYI8vF0';

function myFunction() {
  /*var sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('シート1');
  var a1Data = sheet.getRange('A1').getValue();
  Logger.log(a1Data)*/
  
  var index = testSheet.getIndex();
  var a1Data = testSheet.values[1][index.a];
  Logger.log(a1Data);
}


function edit() {
  var key = testSheet.getRowKey('a'); // return 'B'
  var lastRow = testSheet.sheet.getRange(key + ':' + key).getValues().filter(String).length + 1; // return 15
  testSheet.sheet.getRange(key + lastRow).setValue('A列の最後行に書き換え');
  // testSheet.values[lastRow - 1][testSheet.getIndex().a]) = 'A列の最後行に書き換え'になっている
  Logger.log('hoge')
}