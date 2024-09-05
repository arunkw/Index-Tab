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
  
  // Store existing "About" column values to retain them
  var existingAboutValues = indexSheet.getRange(2, 3, indexSheet.getLastRow() - 1, 1).getValues();
  
  // Clear the "Index" sheet content
  indexSheet.clear();
  
  // Delete columns D onwards
  var lastColumn = indexSheet.getMaxColumns();
  if (lastColumn > 3) {
    indexSheet.deleteColumns(4, lastColumn - 3);
  }
  
  // Set up headers
  indexSheet.getRange('A1').setValue('Obsolete/Current');
  indexSheet.getRange('B1').setValue('Hyperlinks');
  indexSheet.getRange('C1').setValue('About');
  
  // Define options for dropdown
  var options = ['Obsolete', 'Current'];
  var dropdownRange = indexSheet.getRange(2, 1, sheetNames.length - 1, 1);
  var rule = SpreadsheetApp.newDataValidation()
                          .requireValueInList(options)
                          .build();
  dropdownRange.setDataValidation(rule);
  
  var row = 2;
  
  // Populate the "Index" tab with sheet names and hyperlinks
  sheetNames.forEach(function(sheetName, index) {
    if (sheetName !== indexSheetName && sheetName !== coverSheetName) {
      var sheet = ss.getSheetByName(sheetName);
      var sheetUrl = ss.getUrl() + '#gid=' + sheet.getSheetId();
      
      // Set dropdown in column A and sheet names as hyperlinks in column B
      indexSheet.getRange(row, 1).setValue('Current');
      indexSheet.getRange(row, 2).setFormula('=HYPERLINK("' + sheetUrl + '","' + sheetName + '")');
      
      // If the "About" column already has a value, retain it, otherwise leave it blank
      if (existingAboutValues[index] && existingAboutValues[index][0] !== "") {
        indexSheet.getRange(row, 3).setValue(existingAboutValues[index][0]);
      }
      
      row++;
    }
  });
  
  // Remove unused rows beyond the last populated one
  var lastRow = indexSheet.getLastRow();
  var maxRows = indexSheet.getMaxRows();
  if (maxRows > lastRow) {
    indexSheet.deleteRows(lastRow + 1, maxRows - lastRow); // Delete all unused rows from lastRow+1 onwards
  }
  
  // Style the "Index" tab as black background
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
  
  // Populate "Cover" tab with 25 records per column from the "Index" tab, using the formula
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
