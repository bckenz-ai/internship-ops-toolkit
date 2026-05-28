// MENU
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Invite Count')
      .addItem('Save New Baseline (After Update)', 'saveBaseline')
      .addItem('Generate 3PM/9PM Report', 'generateDailyReport')
      .addToUi();
}

// SAVE BASELINE (Run this AFTER update is posted)
function saveBaseline() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("AUTOCOUNTER"); // Live Sheet
  var targetSheet = ss.getSheetByName("BASELINE"); // Hidden Storage Tab
  
  if (!targetSheet) {
    targetSheet = ss.insertSheet("BASELINE");
    targetSheet.hideSheet();
  }
  
  targetSheet.clear();
  
  // Save Timestamp
  var timestamp = new Date();
  targetSheet.getRange(1, 1).setValue("LAST_SAVED");
  targetSheet.getRange(1, 2).setValue(timestamp);
  
  // GET DYNAMIC DATA (Starts at Row 5)
  var lastRow = sourceSheet.getLastRow();
  if (lastRow < 5) {
    SpreadsheetApp.getUi().alert("Not enough data found (Row 5+ needed).");
    return;
  }

  // Get range from D5 down to N (lastRow)
  var names = sourceSheet.getRange(5, 4, lastRow-4, 1).getValues(); // Col D
  var users = sourceSheet.getRange(5, 5, lastRow-4, 1).getValues(); // Col E (Unique ID)
  
  // UPDATED: Now grabbing from Column 13 (M) instead of 14 (N)
  var totals = sourceSheet.getRange(5, 13, lastRow-4, 1).getValues(); // Col M
  
  // Headers for Baseline
  targetSheet.getRange(2, 1).setValue("USERNAME"); // Username as Key
  targetSheet.getRange(2, 2).setValue("NAME");
  targetSheet.getRange(2, 3).setValue("SCORE");
  
  // Save to Baseline
  targetSheet.getRange(3, 1, users.length, 1).setValues(users);
  targetSheet.getRange(3, 2, names.length, 1).setValues(names);
  targetSheet.getRange(3, 3, totals.length, 1).setValues(totals);
  
  SpreadsheetApp.getUi().alert('Baseline Saved. Next report will compare against this moment.');
}

// GENERATE REPORT
function generateDailyReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var liveSheet = ss.getSheetByName("AUTOCOUNTER");
  var baseSheet = ss.getSheetByName("BASELINE");
  
  if (!baseSheet) {
    SpreadsheetApp.getUi().alert('No baseline found. Run "Save New Baseline" first.');
    return;
  }
  
  // Get Last Saved Time
  var lastSaved = baseSheet.getRange(1, 2).getValue();
  var timeStr = Utilities.formatDate(lastSaved, Session.getScriptTimeZone(), "MMM dd, h:mm a");
  
  // Get Maps
  // UPDATED: Changed scoreCol from 14 to 13 (Column M)
  var liveData = getMap(liveSheet, 5, 13, 4, 5); 
  var baseData = getMap(baseSheet, 1, 3, 2, 3); 
  
  // EXCLUDED
  var excludedHosts = ["PARADIMES", "MANILA MI VIDA"];

  // Create TWO lists
  var eligible = [];
  var ineligible = [];
  var totalInvites = 0;
  
  for (var key in liveData) {
    var hostObj = liveData[key]; 
    var current = hostObj.score;
    var name = hostObj.name;
    
    // If name is in the blocked list, skip
    if (name && excludedHosts.indexOf(name.toUpperCase()) > -1) {
      continue; 
    }

    totalInvites += current;

    var previous = 0;
    if (baseData[key]) {
      previous = baseData[key].score;
    }
    
    var diff = current - previous;
    var item = {name: name, current: current, diff: diff};

    // SPLIT LOGIC: Threshold is 5
    if (current >= 5) {
      eligible.push(item);
    } else {
      ineligible.push(item);
    }
  }
  
  // Sorting Function (High to Low)
  var sortFunc = function(a, b) { 
    return b.current - a.current; 
  };
  
  eligible.sort(sortFunc);
  ineligible.sort(sortFunc);
  
  // Get Current Time
  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM dd, h:mm a");

  // Header
  var part1 = "INVITE COUNT AS OF (" + nowStr + ")\n";
  part1 += "(Compared to: " + timeStr + ")\n\n";
  part1 += "ELIGIBLE HOSTS (5+)\n";
  
  // PART TOTAL
  var partTotal = "💰 TOTAL INVITES: " + totalInvites + "\n";
  
  // List
  var part2 = "";
  if (eligible.length === 0) part2 += "(None yet!)\n";
  
  eligible.forEach(function(item) {
    var symbol = "";
    if (item.diff >= 5) symbol = " 🔥🔥"; 
    
    // CLEANER LOOK: Only show (+X) if diff > 0
    var diffDisplay = "";
    if (item.diff > 0) {
      diffDisplay = " (+" + item.diff + ")";
    }
    
    part2 += "- " + item.name + " - " + item.current + diffDisplay + symbol + "\n";
  });

  // INELIGIBLE HOSTS
  part2 += "\nINELIGIBLE HOSTS (< 5)\n";
  if (ineligible.length === 0) part2 += "(Everyone is eligible)\n";

  ineligible.forEach(function(item) {
    var diffDisplay = "";
    if (item.diff > 0) {
      diffDisplay = " (+" + item.diff + ")";
    }
    part2 += "- " + item.name + " - " + item.current + diffDisplay + "\n";
  });

  // BONUS
  part2 += "\nTOP HOST BONUS\n";
  part2 += "* ₱1,000 bonus\n";
  part2 += "* Free +1 for the event\n";
  part2 += "* 5 Access Granted\n";
  part2 += "* Sponsored product (if applicable)";
  
  // HTML OUTPUT WITH TOGGLE
  var initialMsg = part1 + partTotal + part2;
  
  var htmlContent = 
    '<div style="font-family: sans-serif; padding: 5px;">' +
      // TOGGLE CHECKBOX
      '<div style="margin-bottom: 10px;">' +
        '<label style="cursor: pointer; font-weight: bold; font-size: 14px; color: #333;">' +
          '<input type="checkbox" id="toggleTotal" checked onclick="updateReport()" style="transform: scale(1.2); margin-right: 8px;">' +
          'Show Total Invites' +
        '</label>' +
      '</div>' +
      
      '<textarea id="report" style="width:100%; height:300px; padding: 5px; border: 1px solid #ccc; border-radius: 4px; font-family: monospace;">' + initialMsg + '</textarea>' +
      '<br><br>' +
      '<button onclick="copyToClipboard()" style="background-color: #8B0000; color: white; padding: 12px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; width: 100%; font-weight: bold;">' +
        'COPY' +
      '</button>' +
      '<p id="success" style="color: green; display: none; text-align: center; margin-top: 10px; font-weight: bold;">Copied successfully!</p>' +
    '</div>' +
    
    '<script>' +
      // INJECT VARIABLES FROM APPS SCRIPT
      'var p1 = ' + JSON.stringify(part1) + ';' +
      'var pTotal = ' + JSON.stringify(partTotal) + ';' +
      'var p2 = ' + JSON.stringify(part2) + ';' +
      
      'function updateReport() {' +
        'var isChecked = document.getElementById("toggleTotal").checked;' +
        'var box = document.getElementById("report");' +
        // Combine parts based on checkbox
        'box.value = p1 + (isChecked ? pTotal : "") + p2;' +
      '}' +
      
      'function copyToClipboard() {' +
        'var copyText = document.getElementById("report");' +
        'copyText.select();' +
        'document.execCommand("copy");' +
        'var successMsg = document.getElementById("success");' +
        'successMsg.style.display = "block";' +
        'setTimeout(function() { successMsg.style.display = "none"; }, 2000);' +
      '}' +
    '</script>';

  var htmlOutput = HtmlService
    .createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(560); 
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Generate Report');
}

// Returns Object
function getMap(sheet, userCol, scoreCol, nameCol, startRow) {
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return {}; 
  
  // Get all data range
  var data = sheet.getRange(startRow, 1, lastRow-(startRow-1), sheet.getLastColumn()).getValues();
  var map = {};
  
  for (var i = 0; i < data.length; i++) {
    // Arrays are 0-indexed, so Column 1 = Index 0
    var user = data[i][userCol-1]; 
    var score = data[i][scoreCol-1];
    var name = data[i][nameCol-1];
    
    // Must have a username and not be a header/empty line
    if (user && typeof user === 'string' && user != "USERNAME" && user != "") { 
      user = user.trim().toLowerCase(); // Normalize username
      
      map[user] = {
        score: Number(score || 0),
        name: name
      };
    }
  }
  return map;
}
