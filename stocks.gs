function onEdit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  // Only run for edits in Transactions tab
  if (sheet.getName() !== "Transactions") return;

  // Get the edited row and column
  var row = range.getRow();
  var column = range.getColumn();

  // Only process if editing row 3 or below (skip header)
  if (row > 2 && column <= 8) { // Columns A-H
    // Ensure Total Cost formula is applied
    var totalCostRange = sheet.getRange("E2:E1000");
    totalCostRange.setFormula('=IF(AND(C2<>"",D2<>""),C2*D2,"")');

    // Sort data by Date (Column A, ascending), excluding header
    var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 2, 8); // A2:H
    dataRange.sort({column: 1, ascending: true});

    // Assign Lot ID if type = BUY
    var values = dataRange.getValues();
    if(values[row -2][1] === "BUY" && values[row -2][5] === '') {
      var nextLotId = getNextLotId(values, row);
      sheet.getRange(row, 6).setValue(nextLotId);
    }

    // Set font to red for Sell type
    var typeRange = sheet.getRange("B2:B1000");
    var typeValues = typeRange.getValues();
    for (var i = 0; i < typeValues.length; i++) { // Rows
      var value = typeValues[i][0];
      if(value === "SELL") {
        sheet.getRange(i + 2, 2).setFontColor("red");
      } else {
        sheet.getRange(i + 2, 2).clearFormat();
      }
    }
  }

  // Trigger updates to Holdings and Tax Report
  updateAll();
}

// Find the last row above the current one with "BUY" in the type column
// values is an array, so row and col starts at 0 unlike getRange!
function getNextLotId(values, row) {
  var typeColumn = 1;
  var lotIdColumn = 5;
  var nextSeq = 1;  // Default to 1 if no previous "BUY" found
  
  for (var i = row - 3; i >= 0; i--) {  // Ignore current row with -3
    if (values[i][typeColumn] === "BUY") {
      var seqValue = values[i][lotIdColumn];  // Get sequential value from that row
      //SpreadsheetApp.getUi().alert("seqValue: " + seqValue);
      if (!isNaN(seqValue) && seqValue > 0) {
        nextSeq = seqValue + 1;
      }
      break;  // Stop at the first (highest row) match above the current row
    }
  }

  return nextSeq;
}


// Combined function to update Holdings and Tax Report
function updateAll() {
  updateHoldings();
  updateTaxReport();
}

function updateHoldings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var transactions = ss.getSheetByName("Transactions").getDataRange().getValues();
  var holdings = ss.getSheetByName("Holdings");
  
  // Process buys and sells
  var newHoldings = {};
  for (var i = 1; i < transactions.length; i++) {
    var row = transactions[i];
    var type = row[1];
    var origQty = row[2];
    var costPerShare = row[3];
    var lotId = row[5];
    var purchaseDate = row[0];

    if (type == "BUY") {
      if (!newHoldings[lotId]) {
        newHoldings[lotId] = { lotId: lotId, origQty: 0, currQty: 0, purchaseDate: purchaseDate, costPerShare: costPerShare };
      }
      newHoldings[lotId].origQty += origQty;
      newHoldings[lotId].currQty += origQty;
    } else if (type == "SELL" && lotId) {
      if (newHoldings[lotId]) {
        newHoldings[lotId].currQty -= origQty;
      } else {
        SpreadsheetApp.getUi().alert("Lot ID not found");
      }
    }
  }

  // Clear Holdings tab (except header and summary row)
  var numRows = holdings.getLastRow() - 2; // Subtract 2 for header rows
  if (numRows > 0) {
    holdings.getRange(3, 1, numRows, 6).clearContent();
  }
  // holdings.getRange(3, 1, holdings.getLastRow() - 2, 6).clearContent();
  
  // Write data starting at row 3
  var data = Object.values(newHoldings).filter(h => h.origQty > 0);
  for (var i = 0; i < data.length; i++) {
    holdings.getRange(i + 3, 1, 1, 6).setValues([
      [data[i].lotId, data[i].origQty, data[i].currQty, data[i].purchaseDate, data[i].costPerShare, 
       data[i].origQty * data[i].costPerShare]
    ]);
  }

  // Write summary row
  holdings.getRange(2, 1, 1, 6).setValues([[
    "Summary", "=SUM(B3:B100)", "=SUM(C3:C100)", "", "=IF(B2>0, F2/B2, 0)", "=SUM(F3:F100)"
  ]]);
}

function updateTaxReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var transactionsSheet = ss.getSheetByName("Transactions");
  var holdingsSheet = ss.getSheetByName("Holdings");
  var taxReportSheet = ss.getSheetByName("Tax Report");

  // Get data from Transactions and Holdings
  var transactions = transactionsSheet.getDataRange().getValues();
  var holdings = holdingsSheet.getDataRange().getValues();

  // Create a map of holdings for quick lookup (Lot ID -> Cost/Share, Purchase Date)
  var holdingsMap = {};
  for (var i = 2; i < holdings.length; i++) { // Start at row 3 to skip summary row
    var lotId = holdings[i][0]; // Lot ID (Column A)
    var purchaseDate = holdings[i][3]; // Purchase Date (Column C)
    var costPerShare = holdings[i][4]; // Cost/Share (Column D)
    holdingsMap[lotId] = { purchaseDate: purchaseDate, costPerShare: costPerShare };
  }

  // Prepare Tax Report data (include header)
  var taxReportData = [[
    "Sell Date", "Lot ID", "Shares Sold", "UNit Price", "Proceeds", 
    "Cost Basis", "Gain/Loss", "Holding Period", "Short-Term G/L", 
    "Long-Term G/L", "Taxable G/L"
  ]];

  // Process SELL transactions
  for (var i = 1; i < transactions.length; i++) {
    var row = transactions[i];
    var type = row[1]; // Type (Column B)
    if (type == "SELL") {
      var sellDate = row[0]; // Date (Column A)
      var sharesSold = row[2]; // Shares (Column C)
      var salePricePerShare = row[3]; // Price/Share (Column D)
      var proceeds = row[4]; // Total Cost (Column E)
      var lotId = row[5]; // Lot ID (Column F)
      var fee = row[6] || 0; // Brokerage Fee (Column G)

      // Adjust proceeds for fees
      proceeds = proceeds - fee;

      // Get cost basis and purchase date from holdings
      var holding = holdingsMap[lotId];
      if (holding) {
        var costBasis = sharesSold * holding.costPerShare;
        var gainLoss = proceeds - costBasis;
        var holdingPeriod = Math.round((new Date(sellDate) - new Date(holding.purchaseDate)) / (1000 * 60 * 60 * 24)); // Days

        // Calculate Taxable Gain/Loss (50% discount for long-term gains)
        var taxableGainLoss = (holdingPeriod > 365 && gainLoss > 0) ? gainLoss * 0.5 : gainLoss;

        // Set Short-Term or Long-Term Gain/Loss
        var shortTermGainLoss = holdingPeriod <= 365 ? gainLoss : 0;
        var longTermGainLoss = holdingPeriod > 365 ? gainLoss : 0;

        // Add to Tax Report data
        taxReportData.push([
          sellDate, lotId, sharesSold, salePricePerShare, proceeds, 
          costBasis, gainLoss, holdingPeriod, shortTermGainLoss, 
          longTermGainLoss, taxableGainLoss
        ]);
      }
    }
  }

  // Clear Tax Report tab (except header and summary row)
  var numRows = taxReportSheet.getLastRow() - 2; // Subtract 2 for header rows
  if (numRows > 0) {
    taxReportSheet.getRange(3, 1, numRows, 11).clearContent();
  }
  // taxReportSheet.getRange(3, 1, taxReportSheet.getLastRow() - 2, 11).clearContent();

  // Write data starting at row 3
  if (taxReportData.length > 1) {
    taxReportSheet.getRange(3, 1, taxReportData.length - 1, 11).setValues(taxReportData.slice(1));
  }

  // Write summary row
  taxReportSheet.getRange(2, 1, 1, 11).setValues([[
    "Summary", "", "=SUM(C3:C100)", "", "=SUM(E3:E100)", "=SUM(F3:F100)", 
    "=SUM(G3:G100)", "", "=SUM(I3:I100)", "=SUM(J3:J100)", "=SUM(K3:K100)"
  ]]);
}