// TEST CLASP PUSH
function createCustomMenu() {
var ui = SpreadsheetApp.getUi();
  ui.createMenu('Start Process')
    .addItem('Set Balance of Cash Advance', 'setBalance')
    .addItem('Replenish Balance', 'replenishBalance')
    .addItem('Reset SCA Process', 'resetSCA')
    .addSeparator()
      .addItem('Check Authorization and Scripts', 'authorize')
    .addToUi();
}

//sheets
const ws = SpreadsheetApp.getActiveSpreadsheet()
const ws_lb = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DV Logbook");
const ws_opex = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RCDisb Operating Expenses");
const ws_mbap = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RCDisb Medical and Burial Assistance");
const ws_rcdisb = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Report of Cash Disbursements");
const ws_cashdr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cash Disbursement Register");
const dumpsheet = ws.getSheetByName("Information Sheet");
const records_hidden = ws.getSheetByName("RECORDS_HIDDEN");
const ws_url = ("https://docs.google.com/spreadsheets/d/13jvehMdKhF7gcCy3B8zTdooYcgkTgFzzRIZR554TmA4/edit#gid=1388670075");

//lastrow_indicators
const lastrow_lb = ws_lb.getLastRow();
const lastrow_rcdisb = ws_rcdisb.getLastRow();
const lastrow_cashdr = ws_cashdr.getLastRow();
const lastRowFinal = dumpsheet.getLastRow();

//folders
const folder_DVOPEX = DriveApp.getFolderById("1lS60oPJpYgArZq1u8tSPa-8VOzvQA8Md");
const folder_ORSOPEX = DriveApp.getFolderById("1rGr0Py26QQgNG7xV7QhW1mXyYjUDuYDa");
const folder_DVMBAP = DriveApp.getFolderById("1La8N-8xmzlUntaOuWBMcM7hTstm4Cll9");
const folder_ORSMBAP = DriveApp.getFolderById("14t_t6X8ZVZ25NXLlKL2u-CU5vRI9WjGH");

//other
const date = Utilities.formatDate(new Date(), "UTC+8", "MM/dd/yyyy");
var dateObject2 = new Date(date);
var formattedDate2 = Utilities.formatDate(dateObject2, "GMT+8", "yyyy-dd-MM");
const opexDvSingle = dumpsheet.getRange("C16").getValue();
const opexDvReplenish = dumpsheet.getRange("C17").getValue();
const opexORSReplenish = dumpsheet.getRange("C18").getValue();
const mbapDvSingle = dumpsheet.getRange("C19").getValue();
const mbapDvReplenish = dumpsheet.getRange("C20").getValue();
const mbapORSReplenish = dumpsheet.getRange("C21").getValue();
const resetCounter = dumpsheet.getRange("C28").getValue();

function test() {
}

function authorize(){
  var ui = SpreadsheetApp.getUi();
  ui.alert("✅ Authorization and Scripts are Ready!", "Welcome []! The Automated Sheet is ready to use.", ui.ButtonSet.OK)
}

function mainFunction() {
    setControlNumber(function() {
      sortSCA(function() {
        opexOrMBAP(function() {
            newSortCashDR(function() {
        });
      });
    });
  });
}


function setBalance() {
  var ui = SpreadsheetApp.getUi();

  var response1 = ui.prompt(
    'Cash Advance Amount for Operating Expenses',
    'Please enter balance:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response1.getSelectedButton() != ui.Button.OK) {
    ui.alert('⚠️ Operation cancelled.');
    return;
  }
  var cashAdvanceOPEX = response1.getResponseText();

  var response2 = ui.prompt(
    'Cash Advance Amount for Medical And Burial Assistance',
    'Please enter balance :',
    ui.ButtonSet.OK_CANCEL
  );
  if (response2.getSelectedButton() != ui.Button.OK) {
    ui.alert('⚠️ Operation cancelled.');
    return;
  }
  var cashAdvanceMBAP = response2.getResponseText();

  // Update Dump Sheet with selected process and cash advance balance
  dumpsheet.getRange("C3").setValue(cashAdvanceOPEX);
  dumpsheet.getRange("C8").setValue(cashAdvanceMBAP);

  // Show confirmation message
  ui.alert('Cash Advance Balance for Operating Expenses set to ₱' + cashAdvanceOPEX + '\nCash Advance Amount for Medical and Burial Assistance set to ₱' + cashAdvanceMBAP);
  setSheetNumber();
}

function replenishBalance(){
  var ui = SpreadsheetApp.getUi();
  
  var opexBalance = dumpsheet.getRange("C3").getValue();
  var mbapBalance = dumpsheet.getRange("C8").getValue();

  var response1 = ui.prompt(
    'Replenishment Amount for Operating Expenses',
    'Please enter balance:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response1.getSelectedButton() != ui.Button.OK) {
    ui.alert('⚠️ Operation cancelled.');
    return;
  }
  var cashAdvanceOPEX = response1.getResponseText();

  var response2 = ui.prompt(
    'Replenishment Amount for Medical And Burial Assistance',
    'Please enter balance:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response2.getSelectedButton() != ui.Button.OK) {
    ui.alert('⚠️ Operation cancelled.');
    return;
  }
  var cashAdvanceMBAP = response2.getResponseText();

  // Convert values to numbers
  opexBalance = parseFloat(opexBalance) || 0;
  mbapBalance = parseFloat(mbapBalance) || 0;
  cashAdvanceOPEX = parseFloat(cashAdvanceOPEX) || 0;
  cashAdvanceMBAP = parseFloat(cashAdvanceMBAP) || 0;

  // Perform addition
  var finalOpex = opexBalance + cashAdvanceOPEX;
  var finalMbap = mbapBalance + cashAdvanceMBAP;
  var replenishmentTotal = cashAdvanceOPEX + cashAdvanceMBAP;

  Logger.log(finalOpex);
  Logger.log(finalMbap);

  // Update Dump Sheet with new balance values
  dumpsheet.getRange("C3").setValue(finalOpex);
  dumpsheet.getRange("C8").setValue(finalMbap);

  // Show confirmation message
  ui.alert('Replenishment Complete!\n\nCurrent Operating Expenses is ₱' + finalOpex.toFixed(2) + '\nCurrent Medical and Burial Assistance is ₱' + finalMbap.toFixed(2));
  replenishControlNumber(replenishmentTotal);
}

function replenishControlNumber(replenishmentTotal){
var dateObject = new Date(date);
var formattedDate = Utilities.formatDate(dateObject, "GMT+8", "yyyy-dd-MM");
var replenishmentCount = dumpsheet.getRange("C26").getValue();

var controlNumberPrefix;
var controlNumberSuffix;

  controlNumberPrefix = "CEB REPLENISHMENT-" + formattedDate + "-";
  controlNumberSuffix = padNumber(replenishmentCount);
  dumpsheet.getRange("C26").setValue(parseInt(replenishmentCount) + 1);

var controlNumber = (controlNumberPrefix + controlNumberSuffix);
Logger.log(controlNumber)

replenishEntry(replenishmentTotal, controlNumber)
}

function replenishEntry(replenishmentTotal, controlNumber) {
  const rowOffset_cashdr = 15
  const count_cashdr = ws_cashdr.getRange("C13:C71").getDisplayValues().flat().filter(String).length;
  const lastrow_cashdr = count_cashdr + rowOffset_cashdr;
  const rowOffset_balance = 14
  const count_balance = ws_cashdr.getRange("C13:C71").getDisplayValues().flat().filter(String).length;
  const lastrow_balance = count_balance + rowOffset_balance;
  // Get the latest last row before inserting the new entry
  const lastrow_lb = ws_lb.getLastRow();
  
  // Define the values to be set in each column
  var formattedDate = Utilities.formatDate(new Date(), "GMT+8", "yyyy-dd-MM");
  var payee = "SOUTHERN MINDANAO SATELLITE OFFICE";
  var sca = "Replenishment";

  const balanceCashDR = ws_cashdr.getRange(lastrow_balance + 1, 7, 1, 1).getValue();
  Logger.log(balanceCashDR);
  const finalbalance = (balanceCashDR + replenishmentTotal);
  
  // Define the columns where values will be set
  var values = [formattedDate, controlNumber, payee, " ", replenishmentTotal, finalbalance, " ", " ", " ", " ", sca, replenishmentTotal];
  
  // Set values in the row
  var range = ws_cashdr.getRange(lastrow_cashdr + 1 , 2, 1, 12);
  range.setValues([values]);
  addRowCashDR();
}

// function replenishSortCashDR(){
//   var values = ws_lb.getRange(lastrow_lb, 1, 1, 10).getValues()[0];
//   const rowOffset_cashdr = 15
//   const count_cashdr = ws_cashdr.getRange("C13:C71").getDisplayValues().flat().filter(String).length;
//   const lastrow_cashdr = count_cashdr + rowOffset_cashdr;
//   const rowOffset_balance = 14
//   const count_balance = ws_cashdr.getRange("C13:C71").getDisplayValues().flat().filter(String).length;
//   const lastrow_balance = count_balance + rowOffset_balance;

//   const balanceCashDR = ws_cashdr.getRange(lastrow_balance + 1, 7, 1, 1).getValue();
//   Logger.log(balanceCashDR);
//   const finalbalance = (balanceCashDR + values[5]);

//   ws_cashdr.getRange(lastrow_cashdr + 1 , 2, 1, 12).setValues([[values[0], values[1], values[2], " ", values[5], finalbalance, " ", " ", " ", " ", values[8], values[5]]]);
// }

function setSheetNumber() {
  var cashAdvanceBalance = dumpsheet.getRange("C13").getValue();
  var date = Utilities.formatDate(new Date(), "GMT+8", "MM/dd/yyyy");

      //Update OPEX Sheet
      ws_opex.getRange("B5").setValue("Period Covered: " + date);
      ws_opex.getRange("H8").setValue(resetCounter);
      ws_opex.getRange("H9").setValue(resetCounter);

      //Update MBAP Sheet
      ws_mbap.getRange("B5").setValue("Period Covered: " + date);
      ws_mbap.getRange("H8").setValue(resetCounter);
      ws_mbap.getRange("H9").setValue(resetCounter);

      //Update RCDisb Sheet
      ws_rcdisb.getRange("B5").setValue("Period Covered: " + date);
      ws_rcdisb.getRange("H8").setValue(resetCounter);
      ws_rcdisb.getRange("H9").setValue(resetCounter);

      // Update Cash Disbursement sheet
      ws_cashdr.getRange("D16").setValue(["Cash Advance"]);
      ws_cashdr.getRange("E16").setValue(cashAdvanceBalance);
      ws_cashdr.getRange("G16").setValue(cashAdvanceBalance);
      ws_cashdr.getRange("B6").setValue(["Sub-Office/District/Division: Davao Satellite Office"]);
      ws_cashdr.getRange("B7").setValue(["Municipality/City/Province: Davao City"]);
      ws_cashdr.getRange("J5").setValue(["Name of Accountable Officer: SHERMALYN I. MAMARIL"]);
    }

function setControlNumber(callback, trigger) {
  var values = ws_lb.getRange(lastrow_lb, 1, 1, 10).getValues()[0];
  var controlNumberColumn = ws_lb.getRange(lastrow_lb, 2)
  var dateObject = new Date(date);
  var formattedDate = Utilities.formatDate(dateObject, "GMT+8", "yyyy-dd-MM");
  const documentNumber = dumpsheet.getRange("C27").getValue()
  var entries = dumpsheet.getRange("C27").getValue();
  var entriesOPEX = dumpsheet.getRange("C24").getValue();
  var entriesMBAP = dumpsheet.getRange("C25").getValue();
  var trigger = values[7];
  
  var controlNumberPrefix;
  var controlNumberSuffix;
  
  if (trigger == "Operating Expenses") {
    controlNumberPrefix = "CEB OPEX-" + formattedDate + "-";
    controlNumberSuffix = padNumber(documentNumber);
    dumpsheet.getRange("C16").setValue(parseInt(documentNumber) + 1);
    dumpsheet.getRange("C27").setValue(parseInt(entries) + 1);
    dumpsheet.getRange("C24").setValue(parseInt(entriesOPEX) + 1);

  } else if (trigger == "Medical and Burial Assistance") {
    controlNumberPrefix = "CEB MBAP-" + formattedDate + "-";
    controlNumberSuffix = padNumber(documentNumber);
    dumpsheet.getRange("C19").setValue(parseInt(documentNumber) + 1);
    dumpsheet.getRange("C27").setValue(parseInt(entries) + 1);
    dumpsheet.getRange("C25").setValue(parseInt(entriesMBAP) + 1);
  }
  
  controlNumberColumn.setValue(controlNumberPrefix + controlNumberSuffix);
  callback();
}

function padNumber(number) {
  var paddedNumber = number.toString();
  var padLength = 5 - paddedNumber.length;
  for (var i = 0; i < padLength; i++) {
    paddedNumber = "0" + paddedNumber;
  }
  return paddedNumber;
}

function sortSCA(callback){
  var values = ws_lb.getRange(lastrow_lb, 1, 1, 10).getValues()[0];
  var trigger = values[7];

  const rowOffset_rcdisb = 12
  const count_rcdisb = ws_rcdisb.getRange("I13:I49").getValues().flat().filter(String).length;
  const lastrow_rcdisb = count_rcdisb + rowOffset_rcdisb;
  
  if (trigger == "Operating Expenses"){
    const rowOffset_opex = 12
    const count_opex = ws_opex.getRange("I13:I49").getValues().flat().filter(String).length;
    const lastrow_opex = count_opex + rowOffset_opex;
    //opex
    ws_opex.getRange(lastrow_opex + 1 , 2, 1, 8).setValues([[values[0], values[1], "", "", values[2], "", values[4], values[5]]]);
    ws_rcdisb.getRange(lastrow_rcdisb + 1 , 2, 1, 8).setValues([[values[0], values[1], "", "", values[2], "", values[4], values[5]]]);

    generateDVforOPEX();

  } else if(trigger == "Medical and Burial Assistance"){
    const rowOffset_mbap = 12
    const count_mbap = ws_mbap.getRange("I13:I49").getValues().flat().filter(String).length;
    const lastrow_mbap = count_mbap + rowOffset_mbap;
    //mbap
    ws_mbap.getRange(lastrow_mbap + 1 , 2, 1, 8).setValues([[values[0], values[1], "", "", values[2], "", values[4], values[5]]]);
    ws_rcdisb.getRange(lastrow_rcdisb + 1 , 2, 1, 8).setValues([[values[0], values[1], "", "", values[2], "", values[4], values[5]]]);

    generateDVforMBAP();
  }
  callback();
}

function opexOrMBAP(callback) {
  // Fetch the values from the DV Logbook
  var values = ws_lb.getRange(lastrow_lb, 1, 1, 10).getValues()[0];
  var trigger = values[7];

  // Retrieve the current row numbers from "records_hidden"
  var rowNum_Opex = records_hidden.getRange("B2").getValue();
  var rowNum_MBAP = records_hidden.getRange("B3").getValue();
  var rowNum_RCDisb = records_hidden.getRange("B4").getValue();
  Logger.log("Row Numbers - Opex: " + rowNum_Opex + ", MBAP: " + rowNum_MBAP + ", RCDisb: " + rowNum_RCDisb);

  var range, targetSheet, rowNumTarget;

  // Determine which sheet to use based on the trigger
  if (trigger === "Operating Expenses") {
    range = ws_opex.getRange("B13:I49");
    targetSheet = ws_opex;
    rowNumTarget = rowNum_Opex;
  } else if (trigger === "Medical and Burial Assistance") {
    range = ws_mbap.getRange("B13:I49");
    targetSheet = ws_mbap;
    rowNumTarget = rowNum_MBAP;
  } else {
    Logger.log("Trigger does not match any expected values.");
    return; // Exit the function if the trigger doesn't match
  }

  // Check if the specified row has a value and insert a row if necessary
  if (rowNumTarget > 0) {
    var rowValues = targetSheet.getRange(rowNumTarget, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    Logger.log("Row " + rowNumTarget + " Values: " + rowValues);

    var hasValue = rowValues.some(cell => cell !== "");

    if (hasValue) {
      var lastRow = targetSheet.getLastRow(); // Get the last row number in the sheet
      
      if (lastRow >= rowNumTarget) {
        targetSheet.insertRowAfter(rowNumTarget); // Insert after the specified row
        Logger.log("Inserted row after row " + rowNumTarget + " in " + targetSheet.getName());
        
        // Update the row number in "records_hidden"
        if (trigger === "Operating Expenses") {
          records_hidden.getRange("B2").setValue(rowNum_Opex + 1);
        } else if (trigger === "Medical and Burial Assistance") {
          records_hidden.getRange("B3").setValue(rowNum_MBAP + 1);
        }
      }
    }
  } else {
    Logger.log("Invalid row number for target sheet.");
  }

  // Handle the RCDisb sheet
  if (ws_rcdisb) {
    if (rowNum_RCDisb > 0) {
      var rowValues_rcdisb = ws_rcdisb.getRange(rowNum_RCDisb, 1, 1, ws_rcdisb.getLastColumn()).getValues()[0];
      Logger.log("Row " + rowNum_RCDisb + " Values in ws_rcdisb: " + rowValues_rcdisb);

      var hasValue_rcdisb = rowValues_rcdisb.some(cell => cell !== "");

      if (hasValue_rcdisb) {
        var lastRow_rcdisb = ws_rcdisb.getLastRow();
        
        if (lastRow_rcdisb >= rowNum_RCDisb) {
          ws_rcdisb.insertRowAfter(rowNum_RCDisb); // Insert after the specified row
          Logger.log("Inserted row after row " + rowNum_RCDisb + " in ws_rcdisb");
          
          // Update the row number in "records_hidden"
          records_hidden.getRange("B4").setValue(rowNum_RCDisb + 1);
        }
      }
    } else {
      Logger.log("Invalid row number for RCDisb.");
    }
  }
  
  // Append data to ws_rcdisb from the source sheet
  var sourceValues = range.getValues();
  Logger.log("Source Values: " + JSON.stringify(sourceValues));

  var hasValue = sourceValues.some(row => row.some(cell => cell !== ""));
  
  if (hasValue) {
    var lastRow_rcdisb = ws_rcdisb.getLastRow();
    if (lastRow_rcdisb < 1) lastRow_rcdisb = 1; // Default to row 1 if lastRow_rcdisb is invalid

    var targetRange = ws_rcdisb.getRange(lastRow_rcdisb + 1, 1, sourceValues.length, sourceValues[0].length);
    targetRange.setValues(sourceValues);
    Logger.log("Appended data to ws_rcdisb at row: " + (lastRow_rcdisb + 1));
  } else {
    Logger.log("No values found in the source range.");
  }

  if (typeof callback === 'function') {
    callback();
  }
}

// function addRowRCDisbOPEX() {
//   var rangeValues = ws_opex.getRange("B47:I47").getValues();
//   Logger.log("Range Values: " + rangeValues); // Log the values of the range
  
//   // Check if any cell in the range contains a value
//   var hasValue = false;
//   for (var i = 0; i < rangeValues[0].length; i++) {
//     if (rangeValues[0][i] != "") {
//       hasValue = true;
//       break;
//     }
//   }
  
//   // If any cell contains a value, insert a row after row 21
//   if (hasValue) {
//     var currentRowNumber = records_hidden.getRange("B2").getValue();
//     var newRowNumber = parseInt(currentRowNumber) + 1;

//     // Copy formatting and contents from the existing row (rowNumber) to the new row (newRowNumber)
//     var rangeToCopy = ws_opex.getRange("B" + currentRowNumber + ":I" + currentRowNumber);
//     var targetRange = ws_opex.getRange("B" + newRowNumber + ":I" + newRowNumber);
//     rangeToCopy.copyTo(targetRange, { formatOnly: false, contentsOnly: false });

//     // Insert the new row after the current row
//     ws_opex.insertRowAfter(currentRowNumber);
//     Logger.log("Inserted row after: " + currentRowNumber);
    
//     // Update the value in the control sheet
//     records_hidden.getRange("B2").setValue(newRowNumber);
//     Logger.log("Control Sheet updated: New Row Number - " + newRowNumber);
//   }
// }

// function addRowRCDisbMBAP() {
//   var rangeValues = ws_mbap.getRange("B47:I47").getValues();
//   Logger.log("Range Values: " + rangeValues); // Log the values of the range
  
//   // Check if any cell in the range contains a value
//   var hasValue = false;
//   for (var i = 0; i < rangeValues[0].length; i++) {
//     if (rangeValues[0][i] != "") {
//       hasValue = true;
//       break;
//     }
//   }
  
//   // If any cell contains a value, insert a row after row 21
//   if (hasValue) {
//     var currentRowNumber = records_hidden.getRange("B3").getValue();
//     var newRowNumber = parseInt(currentRowNumber) + 1;

//     // Copy formatting and contents from the existing row (rowNumber) to the new row (newRowNumber)
//     var rangeToCopy = ws_mbap.getRange("B" + currentRowNumber + ":I" + currentRowNumber);
//     var targetRange = ws_mbap.getRange("B" + newRowNumber + ":I" + newRowNumber);
//     rangeToCopy.copyTo(targetRange, { formatOnly: false, contentsOnly: false });

//     // Insert the new row after the current row
//     ws_mbap.insertRowAfter(currentRowNumber);
//     Logger.log("Inserted row after: " + currentRowNumber);
    
//     // Update the value in the control sheet
//     records_hidden.getRange("B3").setValue(newRowNumber);
//     Logger.log("Control Sheet updated: New Row Number - " + newRowNumber);
//   }
// }

function newSortCashDR(){
  var values = ws_lb.getRange(lastrow_lb, 1, 1, 10).getValues()[0];
  var trigger = values[8];
  const rowOffset_cashdr = 15
  const count_cashdr = ws_cashdr.getRange("C13:C71").getDisplayValues().flat().filter(String).length;
  const lastrow_cashdr = count_cashdr + rowOffset_cashdr;
  const rowOffset_balance = 14
  const count_balance = ws_cashdr.getRange("C13:C71").getDisplayValues().flat().filter(String).length;
  const lastrow_balance = count_balance + rowOffset_balance;

  const balanceCashDR = ws_cashdr.getRange(lastrow_balance + 1, 7, 1, 1).getValue();
  Logger.log(balanceCashDR);
  const finalbalance = (balanceCashDR - values[5]);

  if (trigger == "5021499000 - Medical/Burial"){
    ws_cashdr.getRange(lastrow_cashdr + 1 , 2, 1, 7).setValues([[values[0], values[1], values[2], " ", values[5], finalbalance, values[5]]]);

  } else if(trigger == "5020301002 - Supplies and Materials"){
    ws_cashdr.getRange(lastrow_cashdr + 1 , 2, 1, 8).setValues([[values[0], values[1], values[2], " ", values[5], finalbalance, " ",values[5]]]);

  } else if(trigger == "5020401000 - Utility Expenses"){
    ws_cashdr.getRange(lastrow_cashdr + 1 , 2, 1, 9).setValues([[values[0], values[1], values[2], " ", values[5], finalbalance, " ", " ",values[5]]]);

  } else {
    ws_cashdr.getRange(lastrow_cashdr + 1 , 2, 1, 12).setValues([[values[0], values[1], values[2], " ", values[5], finalbalance, " ", " ", " ", " ", trigger, values[5]]]);
  }
  // Call the callback function to indicate that newSortCashDR() has completed
  addRowCashDR();
}

function addRowCashDR() {
  const dataRange = ws_cashdr.getRange("B16:B41"); // Define the range of rows for data
  const lastDataRowNumber = dataRange.getValues().filter(String).length + 16; // Offset by 15 for the starting row
  
  // Check if the last data row is at or beyond row 21
  if (lastDataRowNumber >= 21) {
    // Insert the new row after the last filled row
    ws_cashdr.insertRowAfter(lastDataRowNumber);
    Logger.log("Inserted row after: " + lastDataRowNumber);
  }

}

// function addRowCashDR(callback) {
//   var rangeValues = ws_cashdr.getRange("B21:D21").getValues();
//   Logger.log("Range Values: " + rangeValues); // Log the values of the range
  
//   // Check if any cell in the range contains a value
//   var hasValue = false;
//   for (var i = 0; i < rangeValues[0].length; i++) {
//     if (rangeValues[0][i] != "") {
//       hasValue = true;
//       break;
//     }
//   }
  
//   // If any cell contains a value, insert a row after row 21
//   if (hasValue) {
//     var currentRowNumber = records_hidden.getRange("B1").getValue();
//     var newRowNumber = parseInt(currentRowNumber) + 1;

//     // Copy formatting and contents from the existing row (rowNumber) to the new row (newRowNumber)
//     var rangeToCopy = ws_cashdr.getRange("B" + currentRowNumber + ":M" + currentRowNumber);
//     var targetRange = ws_cashdr.getRange("B" + newRowNumber + ":M" + newRowNumber);
//     rangeToCopy.copyTo(targetRange, { formatOnly: false, contentsOnly: false });

//     // Insert the new row after the current row
//     ws_cashdr.insertRowAfter(currentRowNumber);
//     Logger.log("Inserted row after: " + currentRowNumber);
    
//     // Update the value in the control sheet
//     records_hidden.getRange("B1").setValue(newRowNumber);
//     Logger.log("Control Sheet updated: New Row Number - " + newRowNumber);
//   }
  
//   callback();
// }

function generateDVforOPEX(){
  //DV Details
  const file = DriveApp.getFileById("177007s6iwDH1waf-2gHHPvm58jhdPuC3ZtzOCUkXJjs");
  var controlNumberColumn = ws_lb.getRange(lastrow_lb, 2).getValue();
  const values = ws_lb.getRange(lastrow_lb, 1, 1, 10).getValues()[0];
  const totalsAmount = values[5]
    var formattedValue = Utilities.formatString('₱%s', totalsAmount.toLocaleString('en-PH', {minimumFractionDigits: 2}));
    Logger.log(formattedValue);
  const folderSCA = DriveApp.getFolderById("1A2BjRwN9gD1GdwS_kZHEubiLOGfinhC7");
  const particularsFinal = `PAYMENT FOR ${values[4]} IN THE AMOUNT OF ${formattedValue} THIS ${date}.`;

  //Create a Copy of TEMPLATE File and Exchange Values to Placeholders.
  const copyFile = file.makeCopy(`DV ${controlNumberColumn}`, folderSCA);
  const doc = DocumentApp.openById(copyFile.getId());
  const body = doc.getBody(); 

  // Define a list of placeholders and corresponding values
  var placeholders = ['#DATE#','#PARTICULARS#', '#TOTAL#', '#DVNUMBER#', '#ADDRESS#', '#PAYEE#'];
  var dvValues = [date, particularsFinal, formattedValue, values[1], values[3], values[2]];

  // Replace all the placeholders in one batch operation
  for (var i = 0; i < placeholders.length; i++) {
    body.replaceText(placeholders[i], dvValues[i]);
  }
  doc.saveAndClose(); 

  // Update the spreadsheet with the Google Doc link
    var googleDocLink = doc.getUrl();
    Logger.log(googleDocLink);
    ws_lb.getRange(lastrow_lb, 10).setValue(googleDocLink);
}

function generateDVforMBAP(){
  const file = DriveApp.getFileById("177007s6iwDH1waf-2gHHPvm58jhdPuC3ZtzOCUkXJjs");
  const dvLogbookSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DV Logbook");
  var controlNumberColumn = ws_lb.getRange(lastrow_lb, 2).getValue();
  const values = dvLogbookSheet.getRange(lastrow_lb, 1, 1, 10).getValues()[0];
  const totalsAmount = values[5]
    var formattedValue = Utilities.formatString('₱%s', totalsAmount.toLocaleString('en-PH', {minimumFractionDigits: 2}));
    Logger.log(formattedValue);
  const folderSCA = DriveApp.getFolderById("1rrl3hKn-zUsxzwMM2S5vhYoMLTaGF4Zm");

  const particularsFinal = `Payment for Medical / Burial Assistance given to ${values[2]} OF ${values[3]} in the amount of ${formattedValue} this ${date}`;

  //Create a Copy of TEMPLATE File and Exchange Values to Placeholders.
  const copyFile = file.makeCopy(`DV ${controlNumberColumn}`, folderSCA);
  const doc = DocumentApp.openById(copyFile.getId());
  const body = doc.getBody(); 

  // Define a list of placeholders and corresponding values
  var placeholders = ['#DATE#', '#PARTICULARS#', '#TOTAL#', '#DVNUMBER#', '#ADDRESS#', '#PAYEE#'];
  var dvValues = [date, particularsFinal, formattedValue, values[1], values[3], values[2]];

  // Replace all the placeholders in one batch operation
  for (var i = 0; i < placeholders.length; i++) {
    body.replaceText(placeholders[i], dvValues[i]);
  }
  doc.saveAndClose(); 
  var googleDocLink = doc.getUrl();
  Logger.log(googleDocLink);
  ws_lb.getRange(lastrow_lb, 10).setValue(googleDocLink);
}

function resetSCA() {
  try {
    // Define constants
    const sheetNameOPEX = 'RCDisb Operating Expenses';
    const sheetNameMBAP = 'RCDisb Medical and Burial Assistance';
    const sheetNameRCDisb = 'Report of Cash Disbursements';
    const sheetNameCASHDR = 'Cash Disbursement Register';
    const destFolderId = "1Dz79vfQImfNFTNye5LSQlXzBVq6pUiiC";
    const fileId = "15gmsOoqYWD_zOY9FaeTbmBNdhe4LHMv6yKfVKyFQKZE";
    const copyFileName = `${date} SCA_Spreadsheet-0${resetCounter}`;

    Logger.log(`Starting resetSCA process at ${new Date()}`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const destFolder = DriveApp.getFolderById(destFolderId);

    // Make a copy of the file
    Logger.log(`Making a copy of the file ${fileId} to folder ${destFolderId} with name ${copyFileName}`);
    DriveApp.getFileById(fileId).makeCopy(copyFileName, destFolder);
    Logger.log(`File copy created successfully.`);

    // Delete rows and columns if needed
    Logger.log(`Deleting rows and columns`);
    deleteRowsandColumns(); 
    Logger.log(`Rows and columns deleted.`);

    // Delete sheets
    Logger.log(`Deleting sheets: ${[sheetNameOPEX, sheetNameMBAP, sheetNameRCDisb, sheetNameCASHDR].join(', ')}`);
    deleteSheets([sheetNameOPEX, sheetNameMBAP, sheetNameRCDisb, sheetNameCASHDR]);

    // Copy sheets
    Logger.log(`Copying sheets.`);
    copySheets([
      { copy: "templateOpEx", name: sheetNameOPEX, index: 2 },
      { copy: "templateMBAP", name: sheetNameMBAP, index: 3 },
      { copy: "templateRCDisb", name: sheetNameRCDisb, index: 4 },
      { copy: "templateCashDR", name: sheetNameCASHDR, index: 5 }
    ]);
    Logger.log(`Sheets copied.`);

    // Wait before proceeding
    Logger.log(`Waiting for 2 seconds.`);
    Utilities.sleep(2000);
    Logger.log(`Wait completed.`);

    // Delete file (ensure deleteFile function is implemented)
    Logger.log(`Deleting file.`);
    deleteFile();
    Logger.log(`File deleted.`);

    // Clear content and set formulas
    Logger.log(`Clearing content and setting formulas in 'Dumpsheet'.`);
    dumpsheet.getRangeList(['C3', 'C8']).clearContent();
    dumpsheet.getRange("C4").setFormula(`='${sheetNameOPEX}'!I50`);
    dumpsheet.getRange("C9").setFormula(`='${sheetNameMBAP}'!I50`);
    dumpsheet.getRange("C5").setFormula("=MINUS(C3, C4)");
    dumpsheet.getRange("C10").setFormula("=MINUS(C8, C9)");
    dumpsheet.getRange("C12").setFormula("=SUM(C4, C9)");
    dumpsheet.getRange("C13").setFormula("=SUM(C5, C10)");
    dumpsheet.getRange("C28").setValue(resetCounter + 1);

    Logger.log(`Updating 'Records Hidden' sheet.`);
    records_hidden.getRange("B1").setValue(21);
    records_hidden.getRange("B2").setValue(47);
    records_hidden.getRange("B3").setValue(47);
    records_hidden.getRange("B4").setValue(47);
    Logger.log(`'Records Hidden' sheet updated.`);

    Logger.log(`resetSCA process completed at ${new Date()}`);

  } catch (error) {
    Logger.log(`Error in resetSCA: ${error.message}`);
  }
}

function deleteSheets(sheetNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(`Starting to delete sheets: ${sheetNames.join(', ')}`);
  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      Logger.log(`Deleting sheet: ${name}`);
      ss.deleteSheet(sheet);
    } else {
      Logger.log(`Sheet not found: ${name}`);
    }
  });
  Logger.log(`Sheets deletion completed.`);
}

function copySheets(sheets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(`Starting to copy sheets.`);
  sheets.forEach(({ copy, name, index }) => {
    const templateSheet = ss.getSheetByName(copy);
    if (templateSheet) {
      Logger.log(`Copying sheet: ${copy} to new sheet: ${name} at index ${index}`);
      const newSheet = ss.insertSheet(name, index, { template: templateSheet });
      
      // Ensure the newly copied sheet is visible
      newSheet.showSheet();
      Logger.log(`The newly copied sheet ${name} is now visible.`);
    } else {
      Logger.log(`Template sheet not found: ${copy}`);
    }
  });
  Logger.log(`Sheets copying completed.`);
}

// function resetSCA() {
//   // array of promises
// const promises = [
//   generateDVReplenishmentMBAP(),
//   generateORSReplenishmentMBAP(),
//   generateDVReplenishmentOPEX(),
//   generateORSReplenishmentOPEX()
// ];

// // wait for all promises to resolve
// Promise.all(promises)
//   .then(() => {
//     // all documents generated
//     // do something else if needed
//   })
//   .catch((error) => {
//     // handle errors if needed
//   });
  
//   var sheetNameOPEX = ('RCDisb Operating Expenses');
//   var sheetNameMBAP = ('RCDisb Medical and Burial Assistance');
//   var sheetNameRCDisb = ('Report of Cash Disbursements')
//   var sheetNameCASHDR = ('Cash Disbursement Register');
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var destFolder = DriveApp.getFolderById("1Dz79vfQImfNFTNye5LSQlXzBVq6pUiiC"); 
//   DriveApp.getFileById("15gmsOoqYWD_zOY9FaeTbmBNdhe4LHMv6yKfVKyFQKZE").makeCopy(date +" SCA_Spreadsheet-0" + resetCounter, destFolder);

//   deleteRowsandColumns(); 

//   //reset sheet count
//   var resetsheetCountOpex = dumpsheet.getRange("G2:G4").setValue(1);

//   //delete sheets
//   const deleteSheetNames = [
//   sheetNameOPEX, sheetNameMBAP, sheetNameRCDisb, sheetNameCASHDR
// ];

// //copy sheets
// const copySheetNames = [
//   { copy: "templateOpEx", name: "RCDisb Operating Expenses", index: 2 },
//   { copy: "templateMBAP", name: "RCDisb Medical and Burial Assistance", index: 3 },
//   { copy: "templateRCDisb", name: "Report of Cash Disbursements", index: 4 },
//   { copy: "templateCashDR", name: "Cash Disbursement Register", index: 5 }
// ];

//   // Delete sheets.
//   deleteSheetNames.forEach(name => {
//     const sheet = ss.getSheetByName(name);
//     if (!sheet) return;
//     ss.deleteSheet(sheet);
//   });

//   // Copy sheets.
//   copySheetNames.forEach(({ copy, name, index }) => {
//     const sheet = ss.getSheetByName(copy);
//     if (!sheet) return;
//     ss.insertSheet(name, index, { template: sheet });
//   });

//   Utilities.sleep(2000)
//   deleteFile();
//   //dumpsheet clear content
//   dumpsheet.getRangeList(['C3', 'C8']).clearContent();
//   dumpsheet.getRange("C4").setFormula("='" + sheetNameOPEX + "'" + "!I50")
//   dumpsheet.getRange("C9").setFormula("='" + sheetNameMBAP + "'" + "!I50")
//   dumpsheet.getRange("C5").setFormula("=MINUS(C3, C4)")
//   dumpsheet.getRange("C10").setFormula("=MINUS(C8, C9)")
//   dumpsheet.getRange("D12").setFormula("=SUM(C5, C10)")
//   dumpsheet.getRange("C27").setValue(resetCounter + 1);
//   records_hidden.getRange("B1").setValue(parseInt(22));
//   records_hidden.getRange("B2").setValue(parseInt(48));
//   records_hidden.getRange("B3").setValue(parseInt(48));
// }

function deleteRowsandColumns(){
  var placeholderRow = 1; // The row number of the placeholder row
  
  // Delete all rows below the placeholder row
  var numRowsToDelete = ws_lb.getLastRow() - placeholderRow;
  if (numRowsToDelete > 0) {
    ws_lb.deleteRows(placeholderRow + 1, numRowsToDelete);
  }

  // Clear data in all cells of the rows after the placeholder row
  var numRowsToClear = ws_lb.getLastRow() - placeholderRow;
  if (numRowsToClear > 0) {
    var rangeToClear = ws_lb.getRange(placeholderRow + 1, 1, numRowsToClear, 10);
    rangeToClear.clearContent();
    rangeToClear.clearFormat();
  }
  Logger.log(`deleteRowsandColumns function is not yet implemented.`);
}

//REPLENISH FUND OPEX
function generateDVReplenishmentOPEX() {
  var file = DriveApp.getFileById("1sHvYV3WUqbum3rhHzRI99epiDnAd6QMdo-xlXPJuUro");
  var values = ws_lb.getRange(lastrow_lb, 1, 1, 10).getValues()[0];
  var dateStarted = ws_lb.getRange("A2").getValue();
  var dateObject = new Date(dateStarted);
  var formattedDate = Utilities.formatDate(dateObject, "GMT+8", "MM dd yyyy HH:mm:ss");
  var totalsAmount = dumpsheet.getRange("C4").getValue();
    var formattedValue = Utilities.formatString('₱%s', totalsAmount.toLocaleString('en-PH', {minimumFractionDigits: 2}));
    Logger.log(formattedValue);

  var particularsFinal = `Replenishment of Special Advance (OpEx) of the Office of the Vice President for the period of ${formattedDate} to ${date} hereto attached in the amount of ${formattedValue}.`;

  var fileName = (`CEB OPEX DVR-${formattedDate2}-0${opexDvReplenish}`);
  var copyFile = file.makeCopy(fileName, folder_DVOPEX);
  var doc = DocumentApp.openById(copyFile.getId());
  var body = doc.getBody(); 
  // Define a list of placeholders and corresponding values
  var placeholders = ['#DATE#', '#PARTICULARS#', '#TOTAL#', '#DVNUMBER#'];
  var dvValues = [date, particularsFinal, formattedValue, fileName];

  // Replace all the placeholders in one batch operation
  for (var i = 0; i < placeholders.length; i++) {
    body.replaceText(placeholders[i], dvValues[i]);
  }
  doc.saveAndClose(); 
  dumpsheet.getRange("C17").setValue(parseInt(opexDvReplenish) + 1);
}

function generateORSReplenishmentOPEX() {
  var file = DriveApp.getFileById("1ObOMa41ZKc93DXhZRVB59xaIBS-yssyBywLkzGw-__4");
  var dateStarted = ws_lb.getRange("A2").getValue();
  var dateObject = new Date(dateStarted);
  var formattedDate = Utilities.formatDate(dateObject, "GMT+8", "MM dd yyyy HH:mm:ss");
  var totalsAmount = dumpsheet.getRange("C4").getValue();
    var formattedValue = Utilities.formatString('₱%s', totalsAmount.toLocaleString('en-PH', {minimumFractionDigits: 2}));
    Logger.log(formattedValue);

  var particularsFinal = `To obligate Replenishment of Special Advance (OpEx) of the Office of the Vice President for the period of ${formattedDate} to ${date} hereto attached in the amount of ${formattedValue}.`;

  var copyFile = file.makeCopy(`CEB OPEX ORSR-${formattedDate2}-0${opexORSReplenish}`, folder_ORSOPEX);
  var doc = DocumentApp.openById(copyFile.getId());
  var body = doc.getBody(); 
  // Define a list of placeholders and corresponding values
  var placeholders = ['#DATE#', '#PARTICULARS#', '#TOTAL#'];
  var dvValues = [date, particularsFinal, formattedValue];

  // Replace all the placeholders in one batch operation
  for (var i = 0; i < placeholders.length; i++) {
    body.replaceText(placeholders[i], dvValues[i]);
  }
  doc.saveAndClose(); 
  dumpsheet.getRange("C18").setValue(parseInt(opexORSReplenish) + 1);
}
//REPLENISH FUNDS OPEX

//REPLENISH FUNDS MBAP
function generateDVReplenishmentMBAP() {
  var file = DriveApp.getFileById("1sHvYV3WUqbum3rhHzRI99epiDnAd6QMdo-xlXPJuUro");
  var values = ws_lb.getRange(lastrow_lb, 1, 1, 10).getValues()[0];
  var dateStarted = ws_lb.getRange("A2").getValue();
  var dateObject = new Date(dateStarted);
  var formattedDate = Utilities.formatDate(dateObject, "GMT+8", "MM dd yyyy HH:mm:ss EEE");
  var totalsAmount = dumpsheet.getRange("C9").getValue();
    var formattedValue = Utilities.formatString('₱%s', totalsAmount.toLocaleString('en-PH', {minimumFractionDigits: 2}));
    Logger.log(formattedValue);

  var particularsFinal = `To Replenish various expenses incurred in the Office of the Vice President for the period of ${formattedDate} to ${date} hereto attached in the amount of ${formattedValue}.`;

  var fileName = (`CEB MBAP DVR-${formattedDate2}-0${mbapDvReplenish}`);
  var copyFile = file.makeCopy(fileName, folder_DVMBAP);
  var doc = DocumentApp.openById(copyFile.getId());
  var body = doc.getBody(); 
  // Define a list of placeholders and corresponding values
  var placeholders = ['#DATE#', '#PARTICULARS#', '#TOTAL#', '#DVNUMBER#'];
  var dvValues = [date, particularsFinal, formattedValue, fileName];

  // Replace all the placeholders in one batch operation
  for (var i = 0; i < placeholders.length; i++) {
    body.replaceText(placeholders[i], dvValues[i]);
  }
  doc.saveAndClose(); 
  dumpsheet.getRange("C20").setValue(parseInt(mbapDvReplenish) + 1);
}

function generateORSReplenishmentMBAP(){
  var file = DriveApp.getFileById("1ObOMa41ZKc93DXhZRVB59xaIBS-yssyBywLkzGw-__4");
  var dateStarted = ws_lb.getRange("A2").getValue();
  var dateObject = new Date(dateStarted);
  var formattedDate = Utilities.formatDate(dateObject, "GMT+8", "MM dd yyyy HH:mm:ss EEE");
  var totalsAmount = dumpsheet.getRange("C9").getValue();
    var formattedValue = Utilities.formatString('₱%s', totalsAmount.toLocaleString('en-PH', {minimumFractionDigits: 2}));
    Logger.log(formattedValue);

  var particularsFinal = `To Replenish various expenses incurred in the Office of the Vice President for the period of ${formattedDate} to ${date} hereto attached in the amount of ${formattedValue}.`;

  var copyFile = file.makeCopy(`CEB MBAP ORSR-${formattedDate2}-0${mbapORSReplenish}`, folder_ORSMBAP);
  var doc = DocumentApp.openById(copyFile.getId());
  var body = doc.getBody(); 
  // Define a list of placeholders and corresponding values
  var placeholders = ['#DATE#', '#PARTICULARS#', '#TOTAL#'];
  var dvValues = [date, particularsFinal, formattedValue];

  // Replace all the placeholders in one batch operation
  for (var i = 0; i < placeholders.length; i++) {
    body.replaceText(placeholders[i], dvValues[i]);
  }
  doc.saveAndClose(); 
  dumpsheet.getRange("C21").setValue(parseInt(mbapORSReplenish) + 1);
}
//REPLENISH FUNDS MBAP
function deleteFile(myFileName) {
var allFiles, idToDLET, myFolder, rtrnFromDLET, thisFile;

  myFileName = ("Copy of OVP DV & ORS FORM")
  myFolder = DriveApp.getFolderById('10XwDmMlITwGguZZ6jm_HNAGKV041aJmR');

  allFiles = myFolder.getFilesByName(myFileName);

  while (allFiles.hasNext()) {//If there is another element in the iterator
    thisFile = allFiles.next();
    idToDLET = thisFile.getId();
    //Logger.log('idToDLET: ' + idToDLET);
    Logger.log(thisFile)
    Logger.log(idToDLET);

     DriveApp.getFileById(idToDLET).setTrashed(true);
    DriveApp.getFileById
  };
  Logger.log(`deleteFile function is not yet implemented.`);
};

function inputRow(e){
  const authInfo = ScriptApp.getAuthorizationInfo(e.authMode);
  SpreadsheetApp.getUi().alert(authInfo.getAuthorizationStatus());
  if (authInfo.getAuthorizationStatus() == 'REQUIRED') {
    SpreadsheetApp.getUi.alert('Please authenticate this script to run here: ${authInfo.getAuthorizationUrl()}');
  }
  GmailApp.sendEmail('ovpdavaoit@gmail.com', "Hello", "Someone Accessed your spreadsheet.");
}