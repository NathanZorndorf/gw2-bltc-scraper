function getSheetDetails() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Explicitly define the sheet name here
  var sheet = spreadsheet.getSheetByName("scraper-results");
  
  // Gets the range for the first row of data (row 2)
  var dataRow = sheet.getRange(2, 1, 1, sheet.getLastColumn());
  
  // Get all values, formulas, and display values in one call for efficiency
  var values = dataRow.getValues()[0];
  var formulas = dataRow.getFormulas()[0];
  var displayValues = dataRow.getDisplayValues()[0];
  
  // Get the headers from the first row for context
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var output = [];
  // Add headers for the new sheet
  output.push(["Column Header", "Cell Reference", "Data Type", "Format", "Value", "Formula String"]); 

  for (var j = 0; j < values.length; j++) {
    var cellValue = values[j];
    var cellFormula = formulas[j];
    var cellDisplayValue = displayValues[j];
    
    // Determine data type
    var dataType = "";
    if (cellFormula) {
      dataType = "Formula";
    } else if (typeof cellValue === 'string') {
      dataType = "String";
    } else if (typeof cellValue === 'number') {
      dataType = "Number";
    } else if (cellValue instanceof Date) {
      dataType = "Date";
    } else {
      dataType = "Other";
    }
    
    // Get A1 notation and add all data to the output array
    var cellRef = sheet.getRange(2, j + 1).getA1Notation();
    var header = headers[j];
    
    output.push([
      header,
      cellRef,
      dataType,
      cellDisplayValue,
      cellValue,
      cellFormula ? "'" + cellFormula : "" // Prepend with an apostrophe to force string format
    ]);
  }

  // Create a new sheet to write the results to
  var newSheet = spreadsheet.insertSheet("Sheet Details");
  newSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  
  SpreadsheetApp.getUi().alert("Done!", "Sheet details have been written to the new sheet titled 'Sheet Details'.", SpreadsheetApp.getUi().ButtonSet.OK);
}