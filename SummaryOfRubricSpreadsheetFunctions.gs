//TODO: Javadoc
//https://developers.google.com/apps-script/reference/spreadsheet/

function sheetName(idx) {
  if (!idx)
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  else {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var idx = parseInt(idx);
    if (isNaN(idx) || idx < 1 || sheets.length < idx)
      throw "Invalid parameter (it should be a number from 0 to "+sheets.length+")";
    return sheets[idx-1].getName();
  }
}

function getMarkFromSheet(idx, row, col) {
  if (!idx)
    throw "Invalid parameter (it should be a number from 0 to "+sheets.length+")";
  else {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var idx = parseInt(idx);
    if (isNaN(idx) || idx < 1 || sheets.length < idx)
      throw "Invalid parameter (it should be a number from 0 to "+sheets.length+")";
    var values = sheets[idx-1].getRange(row, col, 1, 1).getValues();
    return values[0][0];
  }
}

function getFeedbackFromSheet(idx, questionRowStart, questionRowEnd, questionColStart, questionColEnd) {
  if (!idx)
    throw "Invalid parameter (it should be a number from 0 to "+sheets.length+")";
  else {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var idx = parseInt(idx);
    if (isNaN(idx) || idx < 1 || sheets.length < idx)
      throw "Invalid parameter (it should be a number from 0 to "+sheets.length+")";
    
    var feedbackString = "";
    //Adding 1 since these values are based on indicies 1 to n.
    //So if we had n columns, it would be (n - 1) + 1 = n different indicies
    var numberOfRows = questionRowEnd - questionRowStart + 1;
    var numberOfCols = questionColEnd - questionColStart + 1;
    var range = sheets[idx-1].getRange(questionRowStart, questionColStart, numberOfRows, numberOfCols);
    var values = range.getValues();
    
    for (var row in values) {
      if(values[row][questionColEnd-1] != "") {
        feedbackString += values[row][questionColStart-1] + ": " + values[row][questionColEnd-1] + "<br>\n";
      }
    }
    //Probably a less gross way of doing this, but I can't be bothered for something of this scale
    //Cutting off the last '\n' from feeback
    feedbackString = feedbackString.slice(0,feedbackString.length-1);
    return feedbackString;
  }
}

function getLabGradeFromStudentName(labSheetNum, nameCellValue, nameCol, markCol) {
  if (!labSheetNum)
    throw "Invalid parameter (it should be a number from 0 to "+sheets.length+")";
  else {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var idx = parseInt(labSheetNum);
    if (isNaN(idx) || idx < 1 || sheets.length < idx)
      throw "Invalid parameter (it should be a number from 0 to "+sheets.length+")";
    var values = sheets[idx-1].getDataRange().getValues();
    for (var row in values) {
      if(values[row][nameCol] == nameCellValue) {
        return values[row][markCol];
      }
    }
      return "";
  }
}

//TODO: Update cells that use these functions asynchronously
//https://developers.google.com/web/fundamentals/primers/async-functions
//Potentially just an observer?

//TODO: Research populating a spreadsheet async as new tabs are created?
