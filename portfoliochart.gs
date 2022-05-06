function fiveDayChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("PvD (Auto)");
  var value = sheet.getRange("A1").getValue();
  var range = sheet.getRange("I1:I").getValues();

  var lastRow = null;                                 
  var i = 0;
  while(i < range.length) {
    if (lastRow = 289){                                                
      sheet.getRange("I2").deleteCells(SpreadsheetApp.Dimension.ROWS)
      break;
    } else {
      if (range[i][0] !== "") {
        i++;
      } else {
        lastRow = i + 1;
        break;
      }
    }
  }
  sheet.getRange(`I${lastRow}`).setValue(value)
}

//One-month chart
function oneMonthChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("PvD (Auto)");
  var value = sheet.getRange("A1").getValue();
  var range = sheet.getRange("L1:L").getValues();

  var lastRow = null;
  var i = 0;
  while(i < range.length) {
    if (lastRow = 373){                                                
      sheet.getRange("L2").deleteCells(SpreadsheetApp.Dimension.ROWS) 
      break;
    } else {
      if (range[i][0] !== "") {
        i++;
      } else {
        lastRow = i + 1;
        break;
      }
    }
  }
  sheet.getRange(`L${lastRow}`).setValue(value)
}

//6-month chart
function sixMonthChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("PvD (Auto)");
  var value = sheet.getRange("A1").getValue();
  var range = sheet.getRange("O1:O").getValues();

  var lastRow = null;
  var i = 0;
  while(i < range.length) {
    if (lastRow = 544){                                                
      sheet.getRange("O2").deleteCells(SpreadsheetApp.Dimension.ROWS) 
      break;
    } else {
      if (range[i][0] !== "") {
        i++;
      } else {
        lastRow = i + 1;
        break;
      }
    }
  }
  sheet.getRange(`O${lastRow}`).setValue(value)
}

//Annual Chart
function oneYearChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();      
  var sheet = ss.getSheetByName("PvD (Auto)");
  var value = sheet.getRange("A1").getValue();
  var range = sheet.getRange("R1:R").getValues();

  var lastRow = null;
  var i = 0;
  while(i < range.length) {
    if (lastRow = 733){                                                
      sheet.getRange("R2").deleteCells(SpreadsheetApp.Dimension.ROWS) 
      break;
    } else {
      if (range[i][0] !== "") {
        i++;
      } else {
        lastRow = i + 1;
        break;
      }
    }
  }
  sheet.getRange(`R${lastRow}`).setValue(value)
}


//Daily chart
function dailyChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("PvD (Auto)");
  var value = sheet.getRange("A1").getValue();
  var range = sheet.getRange("F1:F").getValues();

  var lastRow = null;
  var i = 0;
  while(i < range.length) {
    if (range[i][0] !== "") {
      i++;
    } else {
      lastRow = i + 1;
      break;                                              
    }
  }
  sheet.getRange(`F${lastRow}`).setValue(value)
}

function dayChartDelete() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("PvD (Auto)");
  var range = sheet.getRange("F2:F");

  range.clear();
}

