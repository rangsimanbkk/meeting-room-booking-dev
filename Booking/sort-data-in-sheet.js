// This function will be triggered by a time-driven trigger (e.g., every day at a specific time) to sort the data in the specified Google Sheet.
function SortForm(e) {
  const sheetId = "xxxxxxxxxxxxxxxxxxxxxxx"; // Replace with your Google Sheet ID
  const sheetName = "xxxxxxxxxxxxxxxxxxxxxxx"; // Replace with your Google Sheet name
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);

  // Sort the range (excluding header)
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow > 1) { // Only sort if there is more than just the header
    // Uncomment the line below if you want to add a delay before sorting, for example, to ensure that all data is updated before sorting
    //Utilities.sleep(30000); //Wait 30s before sorting, you can adjust the time as needed
    sheet.getRange(2, 1, lastRow - 1, lastColumn).sort(1);  //(startRow, startColumn, numRows, numColumns).sort(columnToSortBy)
  }
}