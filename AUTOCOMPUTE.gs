function onOpen() {
  //SpreadsheetApp.getUi()
    //.createMenu(' Autocompute')
    //.addItem('Calculate All Commissions', 'processPayroll')
    //.addToUi();
  SpreadsheetApp.getUi().createMenu('Invite Count')
    .addItem('Save New Baseline (After Update)', 'saveBaseline')
    .addItem('Generate 3PM/9PM Report', 'generateDailyReport')
    .addToUi();
}

//Add a checkbox in the appropriate tab first
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const cellAddress = range.getA1Notation();
  const newValue = e.value;

  // Change the cells based on the sheet
  const targetCell = "E3"; 
  const targetSheet = "AUTOCOMPUTE";

  if (sheet.getName() === targetSheet && cellAddress === targetCell && newValue === "TRUE") {
    
    //SpreadsheetApp.getUi().alert("Checkbox 'Button' clicked!");
    
    range.setValue(false); 

    processPayroll(); 
  }
}

function processPayroll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Change "Sheet1" to the exact name of your tab
  var sheet = ss.getSheetByName("AUTOCOMPUTE"); 
  
  
  
  // Cell where name to start the calc
  var calculatorInputCell = "C3"; 
  
  // Cell where FINAL commission result appears
  var calculatorOutputCell = "E38";
  
  // Column letter where list of Host Names
  var namesColumn = "M";
  
  // Column letter for Result
  var destinationColumn = "N";
  
  // Row number where list of names starts
  var startRow = 3;
  


  var ui = SpreadsheetApp.getUi();
  
  // Get the list of names
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    ui.alert("No names found in the list.");
    return;
  }
  
  // Get all names in a batch
  var rangeHeight = lastRow - startRow + 1;
  var names = sheet.getRange(namesColumn + startRow + ":" + namesColumn + lastRow).getValues();
  
  // Prepare an array to hold the results
  var results = [];
  
  // Loop through every name
  for (var i = 0; i < names.length; i++) {
    var hostName = names[i][0];
    
    if (hostName === "") {
      results.push([""]); // If name is empty, result is empty
      continue;
    }
    
    // Paste Name into Calculator Input
    sheet.getRange(calculatorInputCell).setValue(hostName);
    
    // Force sheet to update formulas immediately
    SpreadsheetApp.flush(); 
    
    // Get result from the Calculator Output
    var calculatedValue = sheet.getRange(calculatorOutputCell).getValue();
    
    // Add to results list
    results.push([calculatedValue]);
  }
  
  // Paste all results back into the Destination Column at once
  sheet.getRange(destinationColumn + startRow + ":" + destinationColumn + lastRow).setValues(results);
  
  // Optional: Clear the calculator input cell when done
  sheet.getRange(calculatorInputCell).clearContent();
  
  ui.alert("Everything's good!");
}
