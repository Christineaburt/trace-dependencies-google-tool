function extractDataPointsFromDoc() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var tables = body.getTables();
  var regex = /\$\d{1,3}\b(?!\.\d+)/g;
  var matches = [];

  for (var i = 0; i < tables.length; i++) {
    var table = tables[i];
    for (var row = 0; row < table.getNumRows(); row++) {
      for (var col = 0; col < table.getRow(row).getNumCells(); col++) {
        var cellText = table.getRow(row).getCell(col).getText();
        var cellMatches = cellText.match(regex);
        if (cellMatches) {
          matches = matches.concat(cellMatches);
        }
      }
    }
  }

  Logger.log("Doc matches found: " + JSON.stringify(matches));
  return matches;
}

function extractDataPointsFromSheet() {
  var sheetId = '1fj9HtMXh0T_maFuPNdO0mhESg9yn-qGBk-doh1EBrIY';
  var sheetName = 'wordsmith_main_done_states';

  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);
  var data = sheet.getDataRange().getDisplayValues();
  var regex = /\b\d{1,3}\b(?!\.\d+)/g;
  var matches = {};

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === 'Florida') {
      for (var j = 0; j < data[i].length; j++) {
        var cellText = data[i][j].toString();
        var cellMatches = cellText.match(regex);
        if (cellMatches) {
          cellMatches.forEach(function(match) {
            var formattedMatch = '$' + match;
            var cellRef = sheet.getRange(i + 1, j + 1).getA1Notation();
            matches[formattedMatch] = cellRef;
          });
        }
      }
    }
  }

  Logger.log("Sheet matches found: " + JSON.stringify(matches, null, 2));
  return matches;
}

function matchDataPoints(docData, sheetData) {
  var matches = [];

  if (Array.isArray(docData) && typeof sheetData === 'object') {
    docData.forEach(function(point) {
      if (sheetData[point]) {
        matches.push({
          docPoint: point,
          sheetReference: sheetData[point]
        });
      } else {
        Logger.log("No match found in Sheet for Doc value: " + point);
      }
    });

    Logger.log("Final matched entries: " + JSON.stringify(matches, null, 2));
  } else {
    Logger.log("docData is not an array or sheetData is not an object.");
  }

  return matches;
}

function main() {
  var docData = extractDataPointsFromDoc();
  var sheetData = extractDataPointsFromSheet();
  var matches = matchDataPoints(docData, sheetData);

  matches.forEach(function(match) {
    Logger.log(`Match: ${match.docPoint} -> ${match.sheetReference}`);
  });

  Logger.log("✅ Done.");
}

// Run the main function
main();
