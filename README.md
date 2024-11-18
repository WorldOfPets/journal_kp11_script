# journal_kp11_script

```JavaScript 
function myFunction() {
  //const sourceSpreadsheetId = '1ES2gXbEXSrndX7ceulT2LE9DLh_ZBKYzmj7Odc6gfhg'; // Replace with the ID of your source spreadsheet
  const arraySourceSpreadsheetId = ['1ES2gXbEXSrndX7ceulT2LE9DLh_ZBKYzmj7Odc6gfhg', '1pFmtRvYWdkGzenPIzMeoWhFriaJcTixtER8-ijW5heA',  '1-XBkwjXQwBbI3U-FrIwz5yzOfagpGhzE8Yvoqd9N4Gw']

  const sourceRangeFIO = 'Посещаемость!B1:B32'; // Adjust the range as needed
  const sourceRangeData = 'Посещаемость!AV1:AY32';

  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  
  targetSheet.clear();
  // Get the source data
  for(var j = 0; j < arraySourceSpreadsheetId.length; j++){
    const sourceSpreadsheet = SpreadsheetApp.openById(arraySourceSpreadsheetId[j]);
    const sourceData = sourceSpreadsheet.getRange(sourceRangeFIO).getValues();
    const sourceData2 = sourceSpreadsheet.getRange(sourceRangeData).getValues();
    const listName = sourceSpreadsheet.getName()
    targetSheet.getRange((j * 32)+1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
    targetSheet.getRange((j * 32)+1, sourceData[0].length + 1, sourceData2.length, sourceData2[0].length).setValues(sourceData2)
    var sheets = SpreadsheetApp.openById(arraySourceSpreadsheetId[j]).getSheets()
    for (var i = 2; i < sheets.length; i++) {

    var sheet = sheets[i];
    var id = findCellIdByValue(sheet, "средний балл")
    var column_id = id.slice(0, -1);
    Logger.log("Reading data from sheet: " + sheet.getName());
    Logger.log("Reading data from sheet: " + id.slice(0, -1));
    var values = sheet.getRange(column_id+"7:"+column_id+"37").getValues();
    Logger.log(values)
    targetSheet.getRange((j * 32)+1, i+4, 1, 1).setValue(sheet.getName());
    targetSheet.getRange((j * 32)+2, i+4, values.length, values[0].length).setValues(values);
    targetSheet.getRange((j * 32)+1, 1, 1, 1).setValue(listName)
    // Get the data range (all data in the sheet)

    }
  }
}
function findCellIdByValue(sheet, targetValue){
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange(); // Change this if you want to limit the range
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] === targetValue) {
        // Found the value, return the address
        var cell = sheet.getRange(i + 1, j + 1); // +1 because arrays are 0-indexed
        Logger.log('Cell Address: ' + cell.getA1Notation());
        return cell.getA1Notation(); // Return cell address in A1 notation
      }
    }
  }
  Logger.log('Value not found');
  return 'Value not found';
}
```
