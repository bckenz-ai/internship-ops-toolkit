function generateTierReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("GUESTLIST");
  var targetSheet = ss.getSheetByName("GUESTLIST "); 
  
  if (!targetSheet) {
    targetSheet = ss.insertSheet("GUESTLIST ");
  } 

  // SETUP HEADERS & FREEZE ROW 1
  var headers = [
    "#", 
    "Timestamp", 
    "Name of Guest/s\nFormat: Last Name, First Name", 
    "GENDER", 
    "Email Address", 
    "Phone Number", 
    "IG Handle of Guest", 
    "Kindly input the Instagram username without \"@\" of the host who invited you to this event!\n(Example: rycaella_)\nIf none, kindly input PARADIMES as your host.", 
    "TICKET TYPE", 
    "Enter the last (6) digits of your GCash Transaction Reference No.", 
    "Upload your proof of payment below.\n\nNote: Cropped or edited photos will not be accepted. Please upload the full and clear screenshot or image of your payment.", 
    "Do you agree to the terms and conditions stated above?", 
    "Payment Verified", 
    "₱0.00" 
  ];
  
  targetSheet.getRange("A1:N1")
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#d9d9d9")
    .setWrap(true)
    .setVerticalAlignment("middle");
  targetSheet.setFrozenRows(1);
  
  // CLEAR PREVIOUS DATA & OLD CHECKBOXES
  targetSheet.getRange("A2:N").clearContent(); 
  targetSheet.getRange("A2:N").setFontWeight("normal").setBackground(null);
  
  // Wipes out rogue checkboxes from Column M without hurting conditional formatting
  targetSheet.getRange("M2:M").clearDataValidations(); 

  // DEFINE TIX
  var tiers = [
    { label: "EARLY BIRD PHASE 1 - ₱400", filterTerm: "EARLY BIRD PHASE 1 - ₱400" },
    { label: "EARLY BIRD PHASE 2 - ₱500", filterTerm: "EARLY BIRD PHASE 2 - ₱500" },
    { label: "REGULAR - ₱600",               filterTerm: "REGULAR - ₱600" } 
  ];

  // FETCH DATA
  var lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) return; 
  
  var data = sourceSheet.getRange(2, 2, lastRow - 1, 13).getValues(); 

  var outputArray = [];
  var emptyRow = ["", "", "", "", "", "", "", "", "", "", "", "", "", ""]; 
  var headerRowsToBold = []; 
  var dataRowsToCheckbox = []; // NEW FIX: We will track exactly where to put checkboxes

  // BUILD REPORT
  tiers.forEach(function(tier, index) {
    
    var filteredData = data.filter(function(row) {
      return row[7] === tier.filterTerm; 
    });

    var count = filteredData.length;

    var headerRow = [count, tier.label, "", "", "", "", "", "", "", "", "", "", "", ""];
    outputArray.push(headerRow);
    headerRowsToBold.push(outputArray.length + 1); 

    if (count > 0) {
      for (var i = 0; i < count; i++) {
        var sequenceNum = i + 1;
        var newRow = [sequenceNum].concat(filteredData[i]);
        outputArray.push(newRow);
        
        // NEW FIX: Remember this row number so we can put a checkbox here later
        dataRowsToCheckbox.push(outputArray.length + 1);
      }
    } else {
      outputArray.push(emptyRow); 
    }

    if (index < tiers.length - 1) {
      outputArray.push(emptyRow);
      outputArray.push(emptyRow);
    }
  });

  // PASTE RESULTS & APPLY FORMATTING
  if (outputArray.length > 0) {
    targetSheet.getRange(2, 1, outputArray.length, 14).setValues(outputArray);
    
    // Bold the headers
    headerRowsToBold.forEach(function(rowNum) {
      targetSheet.getRange(rowNum, 1, 1, 2).setFontWeight("bold").setBackground("#f3f3f3");
    });

    // NEW FIX: Automatically insert clean checkboxes ONLY on guest rows!
    var checkboxRanges = dataRowsToCheckbox.map(function(rowNum) {
      return "M" + rowNum;
    });
    
    if (checkboxRanges.length > 0) {
      targetSheet.getRangeList(checkboxRanges).insertCheckboxes();
    }
  }
}

function onEdit(e) {
  if (!e) return; 
  
  var editedSheet = e.source.getActiveSheet();
  var sheetName = editedSheet.getName();
  
  // SCENARIO 1: New entry added to raw (raw???) GUESTLIST
  if (sheetName === "GUESTLIST") {
    generateTierReport(); // Updates the generated report
  }
  
  // SCENARIO 2: Tick a box or enter payment in the generated GUEST LIST
  if (sheetName === "GUEST LIST") {
    var row = e.range.getRow();
    var col = e.range.getColumn();
    
    // Check if the edit happened in Col M (13) or Col N (14), and from Row 2 downwards
    if (row >= 2 && (col === 13 || col === 14)) {
      
      // CONFIGURATION: Which column holds the Guest's UNIQUE NAME?
      // 2 = Col B, 3 = Col C, 4 = Col D
      var nameColumnIndex = 3; 
      
      var targetSheet = editedSheet;
      var guestName = targetSheet.getRange(row, nameColumnIndex).getValue();
      var updatedValue = e.range.getValue(); // The new checkbox state or payment amount
      
      if (!guestName || guestName === "") return; // If blank row, do nothing
      
      // Go to the raw database and find this person
      var sourceSheet = e.source.getSheetByName("GUESTLIST");
      var sourceData = sourceSheet.getDataRange().getValues();
      
      for (var i = 1; i < sourceData.length; i++) {
        // Look for the matching name in the raw data (Arrays are 0-indexed, so Col C is index 2)
        if (sourceData[i][nameColumnIndex - 1] === guestName) {
          
          // Found them! Update the exact same column (M or N) in the raw database
          sourceSheet.getRange(i + 1, col).setValue(updatedValue);
          
          // Show a quick popup so you know the database was updated
          // e.source.toast("Updated " + guestName + " in the raw database.", "Sync Complete", 3);
          break;
        }
      }
    }
  }
}
