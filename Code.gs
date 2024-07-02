function LinkChecker() {
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LinkChecker');
var data = sheet.getRange('B2:B' + sheet.getLastRow()).getValues();

  // Get values from the Assets sheet's range C2:C
  var assetsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assets');
  var assetsData = assetsSheet.getRange('C2:C' + assetsSheet.getLastRow()).getValues();
  
  // Flatten the assetsData array to a single-dimensional array for easier comparison
  var assetsValues = assetsData.flat();
  
  for (var i = 0; i < data.length; i++) {
    var url = data[i][0];
    var responseCode, responseBody;

    try {
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      responseCode = response.getResponseCode();
      responseBody = response.getContentText();
    } catch (e) {
      responseCode = -1;
      responseBody = "";
    }

    // Check if responseCode is 200 and responseBody contains any full text match from the assetsValues array
    if (responseCode === 200 && assetsValues.some(value => new RegExp('\\b' + escapeRegExp(value) + '\\b').test(responseBody))) {
      sheet.getRange('C' + (i + 2)).setValue('Removed');
    } else {
      sheet.getRange('C' + (i + 2)).setValue('Active');
    }
  }
}

// Function to escape special characters in a string to use in a regular expression
function escapeRegExp(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}