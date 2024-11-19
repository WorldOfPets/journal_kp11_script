# journal_kp11_script

```JavaScript 
function myFunction() {
  /*
  ИСИП-11
  ИСИП-12
  ИСИП-13
  ИСИП-14в
  КСИК-11
  КСИК-12
  С-11
  С-12

  ИСИП-26
  ИСИП-27
  ИСИП-28
  КСИК-22
  С-24
  С-25

  ИСИП-32
  ИСИП-33
  ИСИП-35
  КСИК-31
  С-32
  С-33

  ИСИП-41
  КСИК-42
  С-41
  
   */
  const arraySourceSpreadsheetId = [
    '1ES2gXbEXSrndX7ceulT2LE9DLh_ZBKYzmj7Odc6gfhg', 
    '1pFmtRvYWdkGzenPIzMeoWhFriaJcTixtER8-ijW5heA',  
    '1-XBkwjXQwBbI3U-FrIwz5yzOfagpGhzE8Yvoqd9N4Gw',
    '1zgVJuiVR89yvtaJN4WU-0Lh3s2EfHN-LZ65u-hCnW-I',
    '1D_vAwzZzXs44Ca8uVT51UYRzQatrD6M4FNQOoq_2Un0',
    '1VBqd3pKPs0JeebwvNRWWufbMfNsuLG9TLjNEdj8SHwU',
    '1OaAfw17hXTM5J7-GPXsMgCuEENkul7sonD1E1aqzhMw',
    '1BgLJz6cRZeBSJ8hFN3Z9W6MP9LUiP980e4nHbO8d-7k',

    '1eyS6nTum0lyT3o9oZlLnvJj7KjuwEg_dW4N0UGndI8M',
    '1IbdiiVlgn5poYsgF3Ey0ewONscf3VVznr7fqfpoHzjc',
    '1szCjm07I8UMIHd3InYy9fOxQJkDPDcovUladbqU4VCs',
    '1ocONoa29t_z35A_dZfb4-dcYfbTMtl59989Nu0cxzIw',
    '1GaF_NRKjqIhWDAMRCbT0QccJs4uZP9K8T1nLnTewhSI',
    '12dWCC__Ta4QBWlkAULEs03by3IQKQxf_azsaN3mR1RQ',

    '12TWiXkZu-3TPYKX3WNUaZoVLf978Oy6preHee7Vm4cQ',
    '1I-S7oD0lDjiCQkBLBikwoZhK8lCfEV2WoPQBZBETU_4',
    '1Zf2mOYRrX6UQOrHFFEmKG3j5StmjokZiCBb55klqoIA',
    '1UQmpHaPcmyCKNMCWmU2Ji7iOHq7HJ3mzqo1PV9psfxQ',
    '1MDOSc7TrMMpeOpwW8TlN_w52m6w2QPJb5OiyjKKI4hQ',
    '1uNkXQz5182M6NSH_Bhyrw5KlrCqqWY06VenHu36kMq4',

    '16NMbYS2CqtpClbrvcqKNvQwZg3DUseeypsm0hTFZejs',
    '15kB3jpwWz9RNm8ri5FagdIfmqUh4UEUoKFTEYgeHiZE',
    '13c0M0OSw7y-SILa-kBs8SP-C9QBzIFfGUEDJWLNMz3o',

    ]

  const sourceRangeFIO = 'Посещаемость!B1:B32'; // Adjust the range as needed
  const sourceRangeData = 'Посещаемость!AV1:AY32';

  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  
  targetSheet.clear();
  // Get the source data
  for(var j = 0; j < arraySourceSpreadsheetId.length; j++){
    try{
      const sourceSpreadsheet = SpreadsheetApp.openById(arraySourceSpreadsheetId[j]);
      const sourceData = sourceSpreadsheet.getRange(sourceRangeFIO).getValues();
      const sourceData2 = sourceSpreadsheet.getRange(sourceRangeData).getValues();
      const listName = sourceSpreadsheet.getName()
      targetSheet.getRange((j * 32)+1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
      targetSheet.getRange((j * 32)+1, sourceData[0].length + 1, sourceData2.length, sourceData2[0].length).setValues(sourceData2)
      var sheets = SpreadsheetApp.openById(arraySourceSpreadsheetId[j]).getSheets()
      for (var i = 2; i < sheets.length; i++) {
        try{
          var sheet = sheets[i];
        var id = findCellIdByValue(sheet, "средний балл")
        var column_id = id.slice(0, -1);
        //Logger.log("Reading data from sheet: " + sheet.getName());
        //Logger.log("Reading data from sheet: " + id.slice(0, -1));
        var values = sheet.getRange(column_id+"7:"+column_id+"37").getValues();
        //Logger.log(values)
        targetSheet.getRange((j * 32)+1, i+4, 1, 1).setValue(sheet.getName());
        targetSheet.getRange((j * 32)+2, i+4, values.length, values[0].length).setValues(values);
        targetSheet.getRange((j * 32)+1, 1, 1, 1).setValue(listName)
        }catch(err){
          Logger.log(err)
        }
      }
    }catch(err){
      Logger.log(err)
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
        //Logger.log('Cell Address: ' + cell.getA1Notation());
        return cell.getA1Notation(); // Return cell address in A1 notation
      }
    }
  }
  Logger.log('Value not found');
  return 'Value not found';
}

```
