function onOpen() {
  createOrUpdateIndexAndCover();
}

function createOrUpdateIndexAndCover() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = ss.getSheets().map(sheet => sheet.getName());
  var indexSheetName = 'Index';
  var coverSheetName = 'Cover';
  
  // Create "Index" sheet if it doesn't exist
  var indexSheet = ss.getSheetByName(indexSheetName);
  if (!indexSheet) {
    indexSheet = ss.insertSheet(indexSheetName);
  }
  
  // Get the last row for the existing values
  var lastRow = indexSheet.getLastRow();
  
  // Store existing values for both Column A (Obsolete/Current) and Column C (About)
  var existingStatusValues = indexSheet.getRange(2, 1, lastRow - 1).getValues(); // Column A (starting from row 2)
  var existingAboutValues = indexSheet.getRange(2, 3, lastRow - 1).getValues();  // Column C (starting from row 2)
  
  // Only clear the content of Column B (Hyperlinks) from row 2 downwards
  indexSheet.getRange(2, 2, lastRow - 1).clearContent();  // Clear Column B only
  
  // Delete columns D onwards if they exist
  var lastColumn = indexSheet.getMaxColumns();
  if (lastColumn > 3) {
    indexSheet.deleteColumns(4, lastColumn - 3);
  }
  
  // Set up headers for Index tab
  indexSheet.getRange('A1').setValue('Obsolete/Current');
  indexSheet.getRange('B1').setValue('Hyperlinks');
  indexSheet.getRange('C1').setValue('About');
  
  var row = 2; // Start at row 2 to skip headers
  
  // Populate the "Index" tab with sheet names and hyperlinks
  sheetNames.forEach(function(sheetName, sheetIndex) {
    if (sheetName !== indexSheetName && sheetName !== coverSheetName) {
      var sheet = ss.getSheetByName(sheetName);
      var sheetUrl = ss.getUrl() + '#gid=' + sheet.getSheetId();
      
      // Set hyperlinks in column B (this column is always refreshed)
      indexSheet.getRange(row, 2).setFormula('=HYPERLINK("' + sheetUrl + '","' + sheetName + '")');
      
      // Retain the existing value in Column A (Obsolete/Current)
      if (existingStatusValues[row - 2] && existingStatusValues[row - 2][0] !== "") {
        indexSheet.getRange(row, 1).setValue(existingStatusValues[row - 2][0]);
      } else {
        setDropdownWithColor(indexSheet, row);  // Set dropdown for new rows
      }
      
      // Retain the existing value in Column C (About)
      if (existingAboutValues[row - 2] && existingAboutValues[row - 2][0] !== "") {
        indexSheet.getRange(row, 3).setValue(existingAboutValues[row - 2][0]);
      }
      
      row++;
    }
  });
  
  // Remove unused rows beyond the last populated one
  lastRow = indexSheet.getLastRow();
  var maxRows = indexSheet.getMaxRows();
  if (maxRows > lastRow) {
    indexSheet.deleteRows(lastRow + 1, maxRows - lastRow); // Delete all unused rows from lastRow+1 onwards
  }
  
  // Style the "Index" tab with a black background
  indexSheet.setTabColor('black');
  var headerRange = indexSheet.getRange(1, 1, 1, 3);
  headerRange.setBackground('black')
             .setFontColor('white')
             .setFontWeight('bold')
             .setHorizontalAlignment('center');
  
  // Adjust column widths
  indexSheet.autoResizeColumn(1);
  indexSheet.autoResizeColumn(2);
  indexSheet.autoResizeColumn(3);
  
  // Create "Cover" tab if it doesn't exist
  var coverSheet = ss.getSheetByName(coverSheetName);
  if (!coverSheet) {
    coverSheet = ss.insertSheet(coverSheetName);
  } else {
    coverSheet.clear(); // Clear previous content
  }
  
  // Style the "Cover" tab as black background
  coverSheet.setTabColor('black');
  
  // Populate "Cover" tab with records where "Index" tab Column A is "Current"
  var column = 1;
  var recordStart = 1;
  
  while (recordStart <= lastRow) {
    var formula = '=IFERROR(FILTER(Index!B' + recordStart + ':B' + (recordStart + 24) + ', Index!A' + recordStart + ':A' + (recordStart + 24) + '="Current"), "")';
    coverSheet.getRange(1, column).setFormula(formula);
    
    column++;
    recordStart += 25;
  }
  
  // Adjust column widths in Cover tab
  for (var i = 1; i <= column; i++) {
    coverSheet.autoResizeColumn(i);
  }
}

// Function to set dropdown with colors for "Obsolete" and "Current" in the Index tab
function setDropdownWithColor(sheet, row) {
  var range = sheet.getRange(row, 1);
  var rule = SpreadsheetApp.newDataValidation()
                           .requireValueInList(['Obsolete', 'Current'])
                           .setAllowInvalid(false)
                           .build();
  
  range.setDataValidation(rule);
  
  // Apply conditional formatting for the dropdown options
  var conditionalFormatRules = sheet.getConditionalFormatRules();
  
  // Light grey for "Obsolete" (Light Grey 2)
  var obsoleteRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Obsolete')
    .setBackground('#D9D9D9') // Light Grey 2
    .setRanges([range])
    .build();
  
  // Light green for "Current" (Light Green 3)
  var currentRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Current')
    .setBackground('#C6EFCE') // Light Green 3
    .setRanges([range])
    .build();
  
  conditionalFormatRules.push(obsoleteRule);
  conditionalFormatRules.push(currentRule);
  
  sheet.setConditionalFormatRules(conditionalFormatRules);
}
