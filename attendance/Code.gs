// Set script timezone to IST (Asia/Kolkata)
// This ensures all time-based triggers run according to IST
const TIMEZONE = 'Asia/Kolkata';

// Send daily attendance reminder ONLY to players who haven't updated for tomorrow
function sendAttendanceReminder() {
  var players = getPlayersList();
  if (players.length === 0) return;
  var link = "page link";
  var subject = "Badminton Club: Please update your attendance for tomorrow";

  // Get tomorrow's date in IST
  var now = new Date();
  var tomorrow = new Date(now.getTime());
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowStr = Utilities.formatDate(tomorrow, TIMEZONE, 'yyyy-MM-dd');

  // Get attendance sheet data
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Attendance');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var dateIdx = headers.indexOf("Date");
  var nameIdx = headers.indexOf("Name");
  var emailIdx = headers.indexOf("Email");

  // Find players who have NOT updated for tomorrow
  var updatedEmails = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][dateIdx] === tomorrowStr && data[i][emailIdx]) {
      updatedEmails.push(String(data[i][emailIdx]).trim().toLowerCase());
    }
  }

  players.forEach(player => {
    if (player.email && player.name) {
      var playerEmail = String(player.email).trim().toLowerCase();
      if (!updatedEmails.includes(playerEmail)) {
        var body = `Dear ${player.name},\n\nPlease update your attendance for tomorrow using the following link:\n${link}\n\nThank you!`;
        MailApp.sendEmail({
          to: player.email,
          subject: subject,
          body: body
        });
      }
    }
  });
}
// Configuration - Update these with your actual IDs/URLs
const SHEET_ID = '1yoqVoIMuFlcgm0gvNmu7fOXwyt1tHlyi-GVXn5xMweI';

// Main function to set up daily entries
function dailySetup() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Attendance');
  const players = getPlayersList();
  // Get today's date in IST
  const now = new Date();
  const today = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd');
  
  // Check if today's date already exists
  const dates = sheet.getRange("A2:A").getValues().flat();
  if (dates.includes(today)) {
    Logger.log(`Date ${today} already exists in sheet. Skipping.`);
    return;
  }
  
  // Clean up tracking properties for yesterday (since we're setting up for new day)
  cleanupOldTrackingProperties();
  
  // Create new rows for today
  // players is array of {name, email}
  // Always use formatted date string for each row
  // Set 'Attended' to match 'Availability' (default 'Yes' if available, else 'No')
  // If you want to allow marking 'Yes' for availability, set both to 'Yes' by default
  const rows = players.map(p => [today, p.name, p.email, "Yes", "", "Yes", 0, 0, ""]);
  const lastRow = sheet.getLastRow();
  // Update headers if needed
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!headers.includes("Email")) {
    sheet.insertColumnAfter(2);
    sheet.getRange(1, 3).setValue("Email");
  }
  sheet.getRange(lastRow + 1, 1, rows.length, 9).setValues(rows);

  Logger.log(`Added ${rows.length} rows for date: ${today}`);
}

// Helper function to cleanup old tracking properties
function cleanupOldTrackingProperties() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var allProperties = scriptProperties.getProperties();
  var now = new Date();
  var yesterday = new Date(now.getTime());
  yesterday.setDate(yesterday.getDate() - 1);
  var yesterdayStr = Utilities.formatDate(yesterday, TIMEZONE, 'yyyy-MM-dd');
  
  // Remove tracking properties for yesterday and older
  Object.keys(allProperties).forEach(function(key) {
    if (key.startsWith('lastCount_') || key.startsWith('celebrationSent_') || key.startsWith('lastNames_')) {
      var dateInKey = key.split('_')[1];
      if (dateInKey && dateInKey <= yesterdayStr) {
        scriptProperties.deleteProperty(key);
        Logger.log('Cleaned up old property: ' + key);
      }
    }
  });
}

// Function to check warnings and inactivity
function checkWarningsAndInactivity() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Attendance');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Get column indices
  const dateIdx = headers.indexOf("Date");
  const nameIdx = headers.indexOf("Name");
  const emailIdx = headers.indexOf("Email");
  const availIdx = headers.indexOf("Availability");
  const reasonIdx = headers.indexOf("Reason");
  const attendIdx = headers.indexOf("Attended");
  const warnIdx = headers.indexOf("Warning Count");
  const missedIdx = headers.indexOf("Missed Days");
  const autoRemoveIdx = headers.indexOf("AutoRemove");
  
  // Get today's date in IST
  const now = new Date();
  const today = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd');
  const playerHistory = {};
  
  // Process all rows (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = row[dateIdx];
    const name = row[nameIdx];
    const email = row[emailIdx];
    const avail = row[availIdx];
    const attended = row[attendIdx];
    const reason = row[reasonIdx];
    
    if (!playerHistory[name]) {
      playerHistory[name] = [];
    }
    
    playerHistory[name].push({
      date: date,
      avail: avail,
      attended: attended,
      reason: reason,
      rowIndex: i + 1
    });
    
    // Check for warning: Availability = 'Yes' but Attended = 'No'
    if (date === today && avail === 'Yes' && attended === 'No') {
      const currentWarn = row[warnIdx] || 0;
      const newWarn = currentWarn + 1;
      sheet.getRange(i + 1, warnIdx + 1).setValue(newWarn);
      
      if (email) {
        MailApp.sendEmail({
          to: email,
          subject: `Badminton Club: Warning Received`,
          body: `Dear ${name},\n\nYou have received a warning for marking 'Yes' for availability but not attending today. Total warnings: ${newWarn}.\n\nPlease ensure you attend if you mark 'Yes'.`
        });
      }
      
      if (newWarn >= 5) {
        sheet.getRange(i + 1, autoRemoveIdx + 1).setValue("Remove");
        Logger.log(`Player ${name} marked for removal due to 5 warnings`);
        MailApp.sendEmail({
          to: "ajithantonnie17@gmail.com",
          subject: `Badminton Club: Player Removal Alert`,
          body: `Player ${name} has received 5 warnings (marked 'Yes' but did not attend) and should be removed.`
        });
      }
    }
  }
// At 10:30pm IST, send email to all players with names of those who marked 'Yes' for tomorrow
function sendAvailabilitySummaryEmail() {
  var players = getPlayersList();
  if (players.length === 0) return;
  var now = new Date();
  var tomorrow = new Date(now.getTime());
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowStr = Utilities.formatDate(tomorrow, TIMEZONE, 'yyyy-MM-dd');
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Attendance');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var dateIdx = headers.indexOf("Date");
  var nameIdx = headers.indexOf("Name");
  var availIdx = headers.indexOf("Availability");
  var yesNames = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][dateIdx] === tomorrowStr && data[i][availIdx] === "Yes") {
      yesNames.push(data[i][nameIdx]);
    }
  }
  
  var subject = "Badminton Club: Tomorrow's Confirmed Players (" + yesNames.length + " players)";
  var body = "Final update for tomorrow's badminton session:\n\n";
  body += "Total confirmed players: " + yesNames.length + "\n\n";
  
  if (yesNames.length > 0) {
    body += "Players who have marked 'Yes' for tomorrow:\n";
    yesNames.forEach(function(name, index) {
      body += (index + 1) + ". " + name + "\n";
    });
  } else {
    body += "No players have confirmed for tomorrow.\n";
  }
  
  body += "\n---\nThis is the final confirmation for tomorrow's session at 10:30 PM IST.";
  
  players.forEach(player => {
    if (player.email) {
      MailApp.sendEmail({
        to: player.email,
        subject: subject,
        body: body
      });
    }
  });
}
// When 4 players have marked 'Yes' for tomorrow, send celebratory email to those 4
// Also detect when someone opts out after 4 confirmed (reducing to 3 or less)
function sendCelebrationIfFourAvailable() {
  var now = new Date();
  var tomorrow = new Date(now.getTime());
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowStr = Utilities.formatDate(tomorrow, TIMEZONE, 'yyyy-MM-dd');
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Attendance');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var dateIdx = headers.indexOf("Date");
  var nameIdx = headers.indexOf("Name");
  var emailIdx = headers.indexOf("Email");
  var availIdx = headers.indexOf("Availability");
  var yesPlayers = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][dateIdx] === tomorrowStr && data[i][availIdx] === "Yes") {
      yesPlayers.push({ name: data[i][nameIdx], email: data[i][emailIdx] });
    }
  }
  
  // Use Script Properties to track state
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastCountKey = 'lastCount_' + tomorrowStr;
  var celebrationSentKey = 'celebrationSent_' + tomorrowStr;
  var lastNamesKey = 'lastNames_' + tomorrowStr;
  
  var lastCount = parseInt(scriptProperties.getProperty(lastCountKey) || '0');
  var celebrationSent = scriptProperties.getProperty(celebrationSentKey) === 'true';
  var lastNamesStr = scriptProperties.getProperty(lastNamesKey) || '';
  var lastNames = lastNamesStr ? lastNamesStr.split(',') : [];
  
  var currentCount = yesPlayers.length;
  var currentNames = yesPlayers.map(p => p.name);
  
  // Case 1: Exactly 4 players confirmed (send celebration if not sent yet)
  if (currentCount === 4 && !celebrationSent) {
    var names = currentNames.join(", ");
    var subject = "🎉 Hurray! 4 members have confirmed for tomorrow";
    var body = `Congratulations! Game on tomorrow! 🏸\n\nThe following 4 members have marked 'Yes' for tomorrow:\n`;
    currentNames.forEach(function(name, index) {
      body += (index + 1) + ". " + name + "\n";
    });
    body += "\nSee you on the court!";
    
    yesPlayers.forEach(p => {
      if (p.email) {
        MailApp.sendEmail({
          to: p.email,
          subject: subject,
          body: body
        });
      }
    });
    
    // Mark celebration as sent
    scriptProperties.setProperty(celebrationSentKey, 'true');
    scriptProperties.setProperty(lastCountKey, '4');
    scriptProperties.setProperty(lastNamesKey, currentNames.join(','));
    
    Logger.log('Celebration email sent to 4 confirmed players');
  }
  // Case 2: Count dropped from 4 to 3 or less (someone opted out)
  else if (lastCount >= 4 && currentCount === 3) {
    // Find who opted out
    var optedOut = lastNames.filter(name => !currentNames.includes(name));
    
    if (optedOut.length > 0) {
      var subject = "⚠️ Update: Player opted out - Only 3 confirmed now";
      var body = `Update on tomorrow's badminton session:\n\n`;
      body += optedOut[0] + " has opted out.\n\n";
      body += "Currently confirmed players (3):\n";
      currentNames.forEach(function(name, index) {
        body += (index + 1) + ". " + name + "\n";
      });
      body += "\nWe need at least 4 players. Please confirm if you can join!";
      
      // Send to all players (not just the 3)
      var allPlayers = getPlayersList();
      allPlayers.forEach(player => {
        if (player.email) {
          MailApp.sendEmail({
            to: player.email,
            subject: subject,
            body: body
          });
        }
      });
      
      Logger.log('Opt-out notification sent: ' + optedOut[0] + ' dropped out, now only 3 players');
    }
    
    // Update tracking
    scriptProperties.setProperty(lastCountKey, currentCount.toString());
    scriptProperties.setProperty(lastNamesKey, currentNames.join(','));
  }
  // Case 3: Update tracking for any other count change
  else if (currentCount !== lastCount) {
    scriptProperties.setProperty(lastCountKey, currentCount.toString());
    scriptProperties.setProperty(lastNamesKey, currentNames.join(','));
  }
}
  
  // Update Missed Days for each player for the current month
  Object.keys(playerHistory).forEach(name => {
    const records = playerHistory[name];
    let missedDays = 0;
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();

    // Count missed days in current month (for removal: only if attended is 'No')
    records.forEach(record => {
      const recordDate = new Date(record.date);
      if (
        recordDate.getMonth() === currentMonth &&
        recordDate.getFullYear() === currentYear
      ) {
        // Missed for removal: only if attended is 'No'
        if (record.attended === "No") {
          missedDays++;
        }
      }
    });

    // Update Missed Days in sheet for today's record
    // Find today's record for this player
    const todayStr = now.toISOString().split('T')[0];
    const todayRecord = records.find(r => r.date === todayStr);
    if (todayRecord) {
      sheet.getRange(todayRecord.rowIndex, missedIdx + 1).setValue(missedDays);
      // Extended Leave: More than 15 days absent in a month (even with reason) = Removal
      if (missedDays > 15) {
        sheet.getRange(todayRecord.rowIndex, autoRemoveIdx + 1).setValue("Remove");
        Logger.log(`Player ${name} marked for removal due to more than 15 missed days in a month`);
        MailApp.sendEmail({
          to: "ajithantonnie17@gmail.com",
          subject: `Badminton Club: Player Removal Alert`,
          body: `Player ${name} has more than 15 missed days in this month and should be removed.`
        });
      }
    }

    // Check for 10 "No" without valid reason in the current month
    let noWithoutReasonCount = 0;
    records.forEach(record => {
      const recordDate = new Date(record.date);
      if (
        recordDate.getMonth() === currentMonth &&
        recordDate.getFullYear() === currentYear &&
        record.avail === "No" && (!record.reason || record.reason.trim() === "")
      ) {
        noWithoutReasonCount++;
      }
    });
    if (noWithoutReasonCount >= 10) {
      // Find the latest record for this player in the current month
      const latestRecord = records
        .filter(record => {
          const recordDate = new Date(record.date);
          return recordDate.getMonth() === currentMonth && recordDate.getFullYear() === currentYear;
        })
        .sort((a, b) => new Date(b.date) - new Date(a.date))[0];
      if (latestRecord) {
        sheet.getRange(latestRecord.rowIndex, autoRemoveIdx + 1).setValue("Remove");
        Logger.log(`Player ${name} marked for removal due to 10 'No' without valid reason in a month`);
        MailApp.sendEmail({
          to: "ajithantonnie17@gmail.com",
          subject: `Badminton Club: Player Removal Alert`,
          body: `Player ${name} has marked 'No' 10 times without valid reason in this month and should be removed.`
        });
      }
    }
  });
}

// Web app to serve data to HTML site
function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'getPlayers') {
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      players: getPlayersList()
    }))
    .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === 'getHosts') {
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      hosts: getHostsList()
    }))
    .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify(getAttendanceData()))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const action = e.parameter.action;
  let result = {};
  
  try {
    if (action === 'submitPlayer') {
      result = submitPlayerAttendance(
        e.parameter.playerName,
        e.parameter.playerEmail,
        e.parameter.availability,
        e.parameter.reason,
        e.parameter.playerPassword
      );
    } else if (action === 'submitHost') {
      result = submitHostAttendance(
        e.parameter.hostName,
        e.parameter.hostEmail,
        e.parameter.hostPasswordHash,
        e.parameter.playerName,
        e.parameter.playerEmail,
        e.parameter.attended
      );
    } else if (action === 'authenticateAdmin') {
      result = authenticateAdmin(
        e.parameter.adminName,
        e.parameter.adminPasswordHash
      );
    } else if (action === 'addPlayer') {
      result = addPlayer(
        e.parameter.adminName,
        e.parameter.adminPasswordHash,
        e.parameter.newPlayerName,
        e.parameter.newPlayerEmail,
        e.parameter.newPlayerPassword
      );
    } else if (action === 'removePlayer') {
      result = removePlayer(
        e.parameter.adminName,
        e.parameter.adminPasswordHash,
        e.parameter.playerToRemove,
        e.parameter.playerToRemoveEmail
      );
    } else if (action === 'addHost') {
      result = addHost(
        e.parameter.adminName,
        e.parameter.adminPasswordHash,
        e.parameter.newHostName,
        e.parameter.newHostEmail,
        e.parameter.newHostRole,
        e.parameter.newHostPasswordHash
      );
    } else if (action === 'removeHost') {
      result = removeHost(
        e.parameter.adminName,
        e.parameter.adminPasswordHash,
        e.parameter.hostToRemove,
        e.parameter.hostToRemoveEmail
      );
    } else if (action === 'setupFirstHost') {
      result = setupFirstHost(
        e.parameter.firstName,
        e.parameter.firstEmail,
        e.parameter.firstPasswordHash
      );
    } else if (action === 'getData') {
      result = getAttendanceData();
    } else {
      result = { success: false, message: 'Invalid action' };
    }
  } catch (error) {
    result = { success: false, message: error.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// Helper function to submit player attendance
function submitPlayerAttendance(playerName, playerEmail, availability, reason, playerPassword) {
  // Check cutoff time (10:30 PM IST)
  const now = new Date();
  const hour = parseInt(Utilities.formatDate(now, TIMEZONE, 'HH'), 10);
  const minute = parseInt(Utilities.formatDate(now, TIMEZONE, 'mm'), 10);
  if (hour >= 22 && minute >= 30) {
    return { success: false, message: 'Form submission is closed. Please submit before 10:30 PM IST.' };
  }
  
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Attendance');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const dateIdx = headers.indexOf("Date");
  const nameIdx = headers.indexOf("Name");
  const emailIdx = headers.indexOf("Email");
  const availIdx = headers.indexOf("Availability");
  const reasonIdx = headers.indexOf("Reason");

  // Validate reason field: if availability is "No", reason is required
  if (availability === 'No' && (!reason || reason.trim() === '')) {
    return { success: false, message: 'Please provide a reason when marking availability as "No".' };
  }
  
  // If availability is "Yes", clear the reason field
  let finalReason = availability === 'Yes' ? '' : (reason || '');

  // Get tomorrow's date in IST
  const tomorrow = new Date(now.getTime());
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, TIMEZONE, 'yyyy-MM-dd');

  // Get player from Players sheet for password verification
  const players = getPlayersList();
  const normalizedInputName = playerName ? playerName.trim().toLowerCase() : '';
  const normalizedInputEmail = playerEmail ? playerEmail.trim().toLowerCase() : '';
  
  let playerObj = null;
  for (let p of players) {
    let pName = p.name ? p.name.trim().toLowerCase() : '';
    let pEmail = p.email ? p.email.trim().toLowerCase() : '';
    if (pName === normalizedInputName && pEmail === normalizedInputEmail) {
      playerObj = p;
      break;
    }
  }
  
  if (!playerObj) {
    // List available player names/emails for debugging
    let availablePlayers = players.map(p => `${p.name} <${p.email}>`).join(', ');
    return { success: false, message: `Invalid player name or email. Available: ${availablePlayers}` };
  }
  
  const playerPasswordHash = playerObj.password;
  
  // Verify password (frontend sends hashed password)
  if (!playerPassword) {
    return { success: false, message: 'Player password is required.' };
  }
  
  // Debug logging for password mismatch
  if (playerPassword !== playerPasswordHash) {
    Logger.log('Password mismatch for player: ' + playerName);
    Logger.log('Received hash length: ' + playerPassword.length);
    Logger.log('Stored hash length: ' + playerPasswordHash.length);
    Logger.log('Received hash (first 20 chars): ' + playerPassword.substring(0, 20));
    Logger.log('Stored hash (first 20 chars): ' + playerPasswordHash.substring(0, 20));
    return { success: false, message: 'Invalid password. Debug: Received=' + playerPassword.length + ' chars, Stored=' + playerPasswordHash.length + ' chars' };
  }

  // Check if entry already exists (compare name and email separately)
  for (let i = 1; i < data.length; i++) {
    const sheetName = data[i][nameIdx] ? String(data[i][nameIdx]).trim() : '';
    const sheetEmail = data[i][emailIdx] ? String(data[i][emailIdx]).trim().toLowerCase() : '';
    const sheetDate = data[i][dateIdx];
    
    // Format the sheet date to ensure proper comparison
    let formattedSheetDate = '';
    try {
      if (sheetDate instanceof Date) {
        formattedSheetDate = Utilities.formatDate(sheetDate, TIMEZONE, 'yyyy-MM-dd');
      } else {
        formattedSheetDate = String(sheetDate).trim();
      }
    } catch (e) {
      formattedSheetDate = String(sheetDate).trim();
    }
    
    // Check if this is the same player and same date
    if (formattedSheetDate === tomorrowStr && 
        sheetName.toLowerCase() === playerName.trim().toLowerCase() && 
        sheetEmail === playerEmail.trim().toLowerCase()) {
      // Update existing entry (use finalReason which is empty for "Yes")
      sheet.getRange(i + 1, availIdx + 1).setValue(availability);
      sheet.getRange(i + 1, reasonIdx + 1).setValue(finalReason);
      // Reset Attended to empty when availability is updated (host/scheduled function will set it)
      const attendIdx = headers.indexOf("Attended");
      if (attendIdx >= 0) {
        sheet.getRange(i + 1, attendIdx + 1).setValue('');
      }
      Logger.log(`Updated availability for ${playerName} (${playerEmail}) for ${tomorrowStr}: ${availability}`);
      
      // Trigger celebration check after player update
      try {
        sendCelebrationIfFourAvailable();
      } catch (e) {
        Logger.log('Warning: Could not trigger celebration check: ' + e.toString());
      }
      
      return { success: true, message: 'Availability updated successfully' };
    }
  }

  // Create new entry (use finalReason which is empty for "Yes")
  const newRow = [
    tomorrowStr,
    playerName,
    playerEmail,
    availability,
    finalReason,
    '', // Attended left empty for host/scheduled function
    0,  // Warning Count
    0,  // Missed Days
    ''  // AutoRemove
  ];

  sheet.appendRow(newRow);
  
  // Trigger celebration check after player submission
  try {
    sendCelebrationIfFourAvailable();
  } catch (e) {
    Logger.log('Warning: Could not trigger celebration check: ' + e.toString());
  }
  
  return { success: true, message: 'Availability submitted successfully' };
}

// Helper function to submit host attendance
function submitHostAttendance(hostName, hostEmail, hostPasswordHash, playerName, playerEmail, attended) {
  // Verify host authentication (using hash, same as host management)
  const hosts = getHostsList();
  let host = null;
  if (hostEmail) {
    host = hosts.find(h => h.name === hostName && h.email && h.email.trim().toLowerCase() === hostEmail.trim().toLowerCase() && h.password === hostPasswordHash);
  } else {
    host = hosts.find(h => h.name === hostName && h.password === hostPasswordHash);
  }
  if (!host) {
    return { success: false, message: 'Invalid host credentials. Please ensure you enter the correct password used during host setup.' };
  }
  
  // Check time restriction (before 10 AM IST)
  const now = new Date();

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Attendance');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const dateIdx = headers.indexOf("Date");
  const nameIdx = headers.indexOf("Name");
  const emailIdx = headers.indexOf("Email");
  const attendIdx = headers.indexOf("Attended");
  const availIdx = headers.indexOf("Availability");
  const warnIdx = headers.indexOf("Warning Count");
  const autoRemoveIdx = headers.indexOf("AutoRemove");

  // Get today's date in IST
  const today = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd');

  // Find the record for player (match both name and email, ignore case and extra spaces)
  const targetName = playerName ? String(playerName).trim() : '';
  const targetEmail = playerEmail ? String(playerEmail).trim().toLowerCase() : '';
  let found = false;
  for (let i = 1; i < data.length; i++) {
    let rawDate = data[i][dateIdx];
    let formattedRowDate = '';
    // Try to format as yyyy-MM-dd regardless of type
    try {
      formattedRowDate = Utilities.formatDate(new Date(rawDate), TIMEZONE, 'yyyy-MM-dd');
    } catch (e) {
      formattedRowDate = String(rawDate).trim();
    }
    const rowName = data[i][nameIdx] ? String(data[i][nameIdx]).trim() : '';
    const rowEmail = data[i][emailIdx] ? String(data[i][emailIdx]).trim().toLowerCase() : '';
    Logger.log(`Row ${i+1}: date="${formattedRowDate}", name="${rowName}", email="${rowEmail}" | Target: date="${today}", name="${targetName}", email="${targetEmail}"`);
    if (
      formattedRowDate === today &&
      rowName.toLowerCase() === targetName.toLowerCase() &&
      rowEmail === targetEmail
    ) {
      found = true;
      // Update attendance
      sheet.getRange(i + 1, attendIdx + 1).setValue(attended);

      // Check for warning condition: ONLY when player marked "Yes" but host updates to "No"
      let currentWarnings = data[i][warnIdx] || 0;
      let newWarnings = currentWarnings;
      let warningGiven = false;
      
      // Warning condition: Player marked Yes, host updates to No (did not attend despite confirming)
      if (data[i][availIdx] === "Yes" && attended === "No") {
        newWarnings = currentWarnings + 1;
        sheet.getRange(i + 1, warnIdx + 1).setValue(newWarnings);
        warningGiven = true;
        Logger.log(`Warning given: Player ${rowName} marked 'Yes' but did not attend. Warning count: ${newWarnings}`);
      }
      
      // Send warning email if warning was given
      if (warningGiven && rowEmail) {
        MailApp.sendEmail({
          to: rowEmail,
          subject: `Badminton Club: Warning Received`,
          body: `Dear ${rowName},\n\nYou have received a warning for marking 'Yes' for availability but not attending. Total warnings: ${newWarnings}.\n\nPlease ensure you attend if you mark 'Yes', or mark 'No' with a reason if you cannot attend.`
        });
      }
      
      // Check for removal after 5 warnings
      if (newWarnings >= 5) {
        sheet.getRange(i + 1, autoRemoveIdx + 1).setValue("Remove");
        Logger.log(`Player ${rowName} marked for removal due to 5 warnings.`);
        if (rowEmail) {
          MailApp.sendEmail({
            to: rowEmail,
            subject: `Badminton Club: Account Removal Notice`,
            body: `Dear ${rowName},\n\nYou have received 5 warnings for not attending despite marking 'Yes'. Your account has been marked for removal. Please contact the host if you believe this is an error.`
          });
        }
      }

      Logger.log(`Attendance updated for player: ${rowName} (${rowEmail}) on ${formattedRowDate}`);
      return { success: true, message: `Attendance updated successfully by ${host.name} (${host.role})` };
    }
  }
  if (!found) {
    Logger.log('Record not found for today. Searched for: date="' + today + '", name="' + targetName + '", email="' + targetEmail + '"');
  }
  return { success: false, message: 'Record not found for today' };
}

// Authentication functions
function authenticateAdmin(adminName, adminPassword) {
  const hosts = getHostsList();
  const admin = hosts.find(h => h.name === adminName && h.password === adminPassword && h.isAdmin);
  if (admin) {
    return { success: true, message: `Welcome ${admin.name}!`, admin: admin };
  } else {
    return { success: false, message: 'Invalid admin credentials' };
  }
}

function setupFirstHost(firstName, firstEmail, firstPasswordHash) {
  if (!firstName || !firstEmail || !firstPasswordHash) {
    return { success: false, message: 'Name, email, and password are required.' };
  }
  // Simple email validation
  var emailRegex = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
  if (!emailRegex.test(firstEmail)) {
    return { success: false, message: 'Please enter a valid email address.' };
  }
  const players = getPlayersList();
  const hosts = getHostsList();

  // Check if system is already initialized
  if (players.length > 0 || hosts.length > 0) {
    return { success: false, message: 'System already initialized' };
  }

  // Add to players sheet with email and password (frontend already hashed it)
  const playersSheet = getPlayersSheet();
  playersSheet.appendRow([firstName, firstEmail, firstPasswordHash]);

  // Add to hosts sheet (store hashed password and email)
  const hostsSheet = getHostsSheet();
  hostsSheet.appendRow([firstName, firstPasswordHash, 'Host', true, firstEmail]);

  return {
    success: true,
    message: `Welcome ${firstName}! You are now the main host.`,
    host: { name: firstName, role: 'Host', isAdmin: true, email: firstEmail }
  };
}

function addPlayer(adminName, adminPassword, newPlayerName, newPlayerEmail, newPlayerPassword) {
  const admin = authenticateAdmin(adminName, adminPassword);
  if (!admin.success) {
    return admin;
  }
  if (!newPlayerName || !newPlayerEmail) {
    return { success: false, message: 'Player name and email are required.' };
  }
  var emailRegex = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
  if (!emailRegex.test(newPlayerEmail)) {
    return { success: false, message: 'Please enter a valid email address.' };
  }
  // Require password for new player (frontend sends hashed password)
  if (!newPlayerPassword) {
    return { success: false, message: 'Player password is required.' };
  }
  // Password is already hashed by frontend - use it directly
  var hashedPassword = newPlayerPassword;
  const players = getPlayersList();
  if (players.find(p => p.name === newPlayerName && p.email === newPlayerEmail)) {
    return { success: false, message: 'Player already exists' };
  }
  const playersSheet = getPlayersSheet();
  playersSheet.appendRow([newPlayerName, newPlayerEmail, hashedPassword]);
  return { success: true, message: `Player "${newPlayerName}" added successfully` };
}

function removePlayer(adminName, adminPassword, playerToRemove, playerToRemoveEmail) {
  const admin = authenticateAdmin(adminName, adminPassword);
  if (!admin.success) {
    return admin;
  }
  
  const targetName = String(playerToRemove).trim().toLowerCase();
  const targetEmail = String(playerToRemoveEmail).trim().toLowerCase();
  const adminNameLower = String(adminName).trim().toLowerCase();
  
  // Prevent self-removal: Check if trying to remove themselves
  if (targetName === adminNameLower) {
    return { success: false, message: 'You cannot remove yourself from the system.' };
  }
  
  Logger.log('Attempting to remove player: name="' + targetName + '", email="' + targetEmail + '"');
  
  // Step 1: Check if player is also a host/cohost and remove them from hosts first
  const hostsSheet = getHostsSheet();
  const hostsData = hostsSheet.getDataRange().getValues();
  let hostRemoved = false;
  let hostRole = "";
  
  // Iterate backwards through hosts to find matching host
  for (let i = hostsData.length - 1; i >= 1; i--) {
    const sheetName = String(hostsData[i][0]).trim().toLowerCase();
    const sheetEmail = String(hostsData[i][4]).trim().toLowerCase();
    
    if (sheetName === targetName && sheetEmail === targetEmail) {
      hostRole = hostsData[i][2];
      Logger.log('Player is also a ' + hostRole + ' - removing from hosts first');
      hostsSheet.deleteRow(i + 1);
      hostRemoved = true;
      break;
    }
  }
  
  // Step 2: Remove player from players sheet
  const playersSheet = getPlayersSheet();
  const playersData = playersSheet.getDataRange().getValues();
  let playerRemoved = false;
  
  // Iterate backwards to avoid index shifting when deleting rows
  for (let i = playersData.length - 1; i >= 1; i--) {
    const sheetName = String(playersData[i][0]).trim().toLowerCase();
    const sheetEmail = String(playersData[i][1]).trim().toLowerCase();
    
    Logger.log('Row ' + (i+1) + ': name="' + sheetName + '", email="' + sheetEmail + '"');
    
    if (sheetName === targetName && sheetEmail === targetEmail) {
      Logger.log('Deleting row ' + (i+1) + ' for player: name="' + sheetName + '", email="' + sheetEmail + '"');
      playersSheet.deleteRow(i + 1);
      playerRemoved = true;
      break;
    }
  }
  
  // Return appropriate message based on what was removed
  if (hostRemoved && playerRemoved) {
    return { 
      success: true, 
      message: `${hostRole} "${playerToRemove}" removed from both hosts and players successfully` 
    };
  } else if (playerRemoved) {
    return { 
      success: true, 
      message: `Player "${playerToRemove}" removed successfully` 
    };
  } else if (hostRemoved) {
    return { 
      success: true, 
      message: `${hostRole} "${playerToRemove}" removed from hosts (was not found in players)` 
    };
  } else {
    return { 
      success: false, 
      message: 'Player not found in either players or hosts list' 
    };
  }
}

function addHost(adminName, adminPassword, newHostName, newHostEmail, newHostRole, newHostPassword) {
  const admin = authenticateAdmin(adminName, adminPassword);
  if (!admin.success) {
    return admin;
  }
  
  const adminNameLower = String(adminName).trim().toLowerCase();
  const newHostNameLower = String(newHostName).trim().toLowerCase();
  
  // Get current admin's details to check if they're a Co-Host
  const hosts = getHostsList();
  const currentAdminHost = hosts.find(h => 
    h.name.trim().toLowerCase() === adminNameLower
  );
  
  // Get the player's existing password from Players sheet
  const players = getPlayersList();
  const player = players.find(p => 
    p.name.trim().toLowerCase() === newHostNameLower && 
    p.email.trim().toLowerCase() === newHostEmail.trim().toLowerCase()
  );
  
  if (!player) {
    return { success: false, message: 'Player not found. Only existing players can be promoted to Host/Co-Host.' };
  }
  
  if (!player.password) {
    return { success: false, message: 'Player password not found. Cannot promote to host.' };
  }
  
  // Check if already a host - if so, update their role instead of adding new
  const existingHost = hosts.find(h => 
    h.name.trim().toLowerCase() === newHostNameLower && 
    h.email.trim().toLowerCase() === newHostEmail.trim().toLowerCase()
  );
  
  const hostsSheet = getHostsSheet();
  const isAdmin = newHostRole === 'Host';
  
  if (existingHost) {
    // Prevent demoting yourself or changing your own role
    if (newHostNameLower === adminNameLower) {
      return { success: false, message: 'You cannot change your own role. Only another host can modify your privileges.' };
    }
    
    // Prevent Co-Hosts from changing other hosts' roles
    if (currentAdminHost && currentAdminHost.role === 'Co-Host') {
      return { success: false, message: 'Co-Hosts cannot change the role of other hosts. Only a Host can modify host privileges.' };
    }
    
    // Update existing host's role
    const hostsData = hostsSheet.getDataRange().getValues();
    for (let i = 1; i < hostsData.length; i++) {
      const sheetName = String(hostsData[i][0]).trim().toLowerCase();
      const sheetEmail = String(hostsData[i][4]).trim().toLowerCase();
      
      if (sheetName === newHostNameLower && sheetEmail === newHostEmail.trim().toLowerCase()) {
        // Update Role (column 3) and IsAdmin (column 4)
        hostsSheet.getRange(i + 1, 3).setValue(newHostRole); // Role
        hostsSheet.getRange(i + 1, 4).setValue(isAdmin); // IsAdmin
        return { success: true, message: `${newHostRole} "${newHostName}" role updated successfully.` };
      }
    }
  }
  
  // Add new host if not found
  hostsSheet.appendRow([newHostName, player.password, newHostRole, isAdmin, newHostEmail]);
  
  return { success: true, message: `${newHostRole} "${newHostName}" added successfully. They can use their existing player password.` };
}

function removeHost(adminName, adminPassword, hostToRemove, hostToRemoveEmail) {
  const admin = authenticateAdmin(adminName, adminPassword);
  if (!admin.success) {
    return admin;
  }
  
  const targetName = String(hostToRemove).trim().toLowerCase();
  const adminNameLower = String(adminName).trim().toLowerCase();
  
  // Prevent self-removal: Compare names (case-insensitive)
  if (targetName === adminNameLower) {
    return { success: false, message: 'You cannot remove yourself from hosts. Only another host can remove you.' };
  }
  
  // Get current admin's details to check if they're a Co-Host
  const hosts = getHostsList();
  const currentAdminHost = hosts.find(h => 
    h.name.trim().toLowerCase() === adminNameLower
  );
  
  // Prevent Co-Hosts from removing any hosts
  if (currentAdminHost && currentAdminHost.role === 'Co-Host') {
    return { success: false, message: 'Co-Hosts cannot remove other hosts. Only a Host can remove hosts or co-hosts.' };
  }
  
  const hostsSheet = getHostsSheet();
  const data = hostsSheet.getDataRange().getValues();
  let removed = false;
  let removedRole = "";
  
  const targetEmail = String(hostToRemoveEmail).trim().toLowerCase();
  
  // IMPORTANT: Iterate backwards to avoid index shifting when deleting rows
  for (let i = data.length - 1; i >= 1; i--) {
    const sheetName = String(data[i][0]).trim().toLowerCase();
    const sheetEmail = String(data[i][4]).trim().toLowerCase();
    
    if (sheetName === targetName && sheetEmail === targetEmail) {
      removedRole = data[i][2];
      hostsSheet.deleteRow(i + 1);
      removed = true;
      break;
    }
  }
  
  if (removed) {
    return { success: true, message: `${removedRole} "${hostToRemove}" removed successfully` };
  }
  return { success: false, message: 'Host not found' };
}

// Helper function to get attendance data
function getAttendanceData() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Attendance');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { success: true, data: [] };
  }
  
  const headers = data[0];
  const dateIdx = headers.indexOf("Date");
  const nameIdx = headers.indexOf("Name");
  const availIdx = headers.indexOf("Availability");
  const reasonIdx = headers.indexOf("Reason");
  const attendIdx = headers.indexOf("Attended");
  const warnIdx = headers.indexOf("Warning Count");
  const autoRemoveIdx = headers.indexOf("AutoRemove");
  
  const result = [];
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    result.push({
      date: row[dateIdx],
      player: row[nameIdx],
      availability: row[availIdx],
      reason: row[reasonIdx],
      attended: row[attendIdx],
      warnings: row[warnIdx] || 0,
      status: row[autoRemoveIdx] || 'Active'
    });
  }
  
  return { success: true, data: result };
}

// Helper function to get players list
function getPlayersList() {
  const playersSheet = getPlayersSheet();
  const lastRow = playersSheet.getLastRow();
  
  // If only header row exists, return empty array
  if (lastRow <= 1) {
    return [];
  }
  
  const data = playersSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return data.map(row => ({ name: row[0], email: row[1], password: row[2] })).filter(p => p.name && p.name.toString().trim() !== '');
}

// Helper function to get hosts list
function getHostsList() {
  const hostsSheet = getHostsSheet();
  const lastRow = hostsSheet.getLastRow();
  
  // If only header row exists, return empty array
  if (lastRow <= 1) {
    return [];
  }
  
  const data = hostsSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  return data.map(row => ({
    name: row[0],
    password: row[1],
    role: row[2],
    isAdmin: row[3] === true || row[3] === 'TRUE',
    email: row[4]
  })).filter(host => host.name && host.name.toString().trim() !== '');
}

// Helper functions to get/create sheets
function getPlayersSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Players');
  if (!sheet) {
    sheet = ss.insertSheet('Players');
  sheet.getRange(1, 1, 1, 3).setValues([['Name', 'Email', 'Password']]);
  } else {
    // Ensure Email and Password columns exist
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes('Email')) {
      sheet.insertColumnAfter(1);
      sheet.getRange(1, 2).setValue('Email');
    }
    if (!headers.includes('Password')) {
      sheet.insertColumnAfter(2);
      sheet.getRange(1, 3).setValue('Password');
    }
  }
  return sheet;
}

function getHostsSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Hosts');
  if (!sheet) {
    sheet = ss.insertSheet('Hosts');
    sheet.getRange(1, 1, 1, 5).setValues([['Name', 'Password', 'Role', 'IsAdmin', 'Email']]);
  } else {
    // Ensure Email column exists
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes('Email')) {
      sheet.insertColumnAfter(4);
      sheet.getRange(1, 5).setValue('Email');
    }
  }
  return sheet;
}

// Function to set up triggers (run this once manually)
function setupTriggers() {
  // Delete existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Set script timezone to IST
  // Note: Also set this in Project Settings -> General Settings -> Time zone: (GMT+05:30) India Standard Time
  
  // All triggers scheduled in IST (Asia/Kolkata)
  // The .inTimezone() method ensures triggers fire according to IST
  
  // dailySetup for next day at 8am IST
  ScriptApp.newTrigger('dailySetup')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone(TIMEZONE)
    .create();

  // checkWarningsAndInactivity at 11pm IST
  ScriptApp.newTrigger('checkWarningsAndInactivity')
    .timeBased()
    .atHour(23)
    .everyDays(1)
    .inTimezone(TIMEZONE)
    .create();

  // sendAttendanceReminder at 9pm IST
  ScriptApp.newTrigger('sendAttendanceReminder')
    .timeBased()
    .atHour(21)
    .everyDays(1)
    .inTimezone(TIMEZONE)
    .create();

  // sendAvailabilitySummaryEmail at 10:30pm IST
  ScriptApp.newTrigger('sendAvailabilitySummaryEmail')
    .timeBased()
    .atHour(22)
    .everyDays(1)
    .nearMinute(30)
    .inTimezone(TIMEZONE)
    .create();

  // markMissingPlayersAsNo at 10:30pm IST
  ScriptApp.newTrigger('markMissingPlayersAsNo')
    .timeBased()
    .atHour(22)
    .everyDays(1)
    .nearMinute(30)
    .inTimezone(TIMEZONE)
    .create();

  // sendCelebrationIfFourAvailable: single time-based trigger every 10 minutes
  // This runs continuously to check for 4 players or opt-outs
  ScriptApp.newTrigger('sendCelebrationIfFourAvailable')
    .timeBased()
    .everyMinutes(10)
    .create();

  // autoSetAttendedToAvailability at 12:01am IST
  ScriptApp.newTrigger('autoSetAttendedToAvailability')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .nearMinute(1)
    .inTimezone(TIMEZONE)
    .create();

  Logger.log('All triggers set up successfully with IST timezone');
}

// At 10:30pm IST, mark missing players as 'No' for tomorrow
function markMissingPlayersAsNo() {
  var players = getPlayersList();
  if (players.length === 0) return;
  var now = new Date();
  var tomorrow = new Date(now.getTime());
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowStr = Utilities.formatDate(tomorrow, TIMEZONE, 'yyyy-MM-dd');
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Attendance');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var dateIdx = headers.indexOf("Date");
  var nameIdx = headers.indexOf("Name");
  var emailIdx = headers.indexOf("Email");

  // Find emails already updated for tomorrow
  var updatedEmails = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][dateIdx] === tomorrowStr && data[i][emailIdx]) {
      updatedEmails.push(String(data[i][emailIdx]).trim().toLowerCase());
    }
  }

  // For each player not updated, add a row with 'No' for tomorrow
  players.forEach(function(player) {
    if (player.email && player.name) {
      var playerEmail = String(player.email).trim().toLowerCase();
      if (!updatedEmails.includes(playerEmail)) {
        var newRow = [
          tomorrowStr,
          player.name,
          player.email,
          "No",
          "", // Reason
          "", // Attended
          0,   // Warning Count
          0,   // Missed Days
          ""  // AutoRemove
        ];
        sheet.appendRow(newRow);
      }
    }
  });
}

// Function to initialize sheets with headers (run this once manually)
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // Create/setup Attendance sheet
  let attendanceSheet = ss.getSheetByName('Attendance');
  if (!attendanceSheet) {
    attendanceSheet = ss.insertSheet('Attendance');
  }
  
  // Set headers if not already set or if missing columns
  const requiredHeaders = ['Date', 'Name', 'Email', 'Availability', 'Reason', 'Attended', 'Warning Count', 'Missed Days', 'AutoRemove'];
  const lastCol = attendanceSheet.getLastColumn();
  
  if (lastCol === 0) {
    // Sheet is empty, set headers directly
    attendanceSheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
  } else {
    // Check existing headers
    const headers = attendanceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    let needsUpdate = false;
    for (let i = 0; i < requiredHeaders.length; i++) {
      if (headers[i] !== requiredHeaders[i]) {
        needsUpdate = true;
        break;
      }
    }
    if (needsUpdate) {
      attendanceSheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    }
  }
  
  // Create Players sheet
  getPlayersSheet();
  
  // Create Hosts sheet
  getHostsSheet();
  
  Logger.log('Sheets initialized successfully');
}

function autoSetAttendedToAvailability() {
  // This function runs at 12:01 AM every day
  // It processes YESTERDAY's session records (the session that just ended)
  // Example: Runs on Nov 6 at 12:01 AM → Processes Nov 5's session
  
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Attendance');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dateIdx = headers.indexOf("Date");
  const nameIdx = headers.indexOf("Name");
  const emailIdx = headers.indexOf("Email");
  const availIdx = headers.indexOf("Availability");
  const attendIdx = headers.indexOf("Attended");

  // Get yesterday's date in IST (the session day that just ended at midnight)
  // If today is Nov 6 at 12:01 AM, we process Nov 5's records
  const now = new Date();
  const yesterday = new Date(now.getTime());
  yesterday.setDate(yesterday.getDate() - 1);
  const yesterdayStr = Utilities.formatDate(yesterday, TIMEZONE, 'yyyy-MM-dd');
  
  Logger.log(`Running autoSetAttendedToAvailability for session date: ${yesterdayStr}`);

  let updatedCount = 0;
  for (let i = 1; i < data.length; i++) {
    // Format the date from sheet to ensure proper comparison
    let rowDate = data[i][dateIdx];
    let formattedRowDate = '';
    try {
      if (rowDate instanceof Date) {
        formattedRowDate = Utilities.formatDate(rowDate, TIMEZONE, 'yyyy-MM-dd');
      } else {
        formattedRowDate = String(rowDate).trim();
      }
    } catch (e) {
      formattedRowDate = String(rowDate).trim();
    }
    
    // Check if this is yesterday's record (the session that just ended)
    if (formattedRowDate === yesterdayStr) {
      const attended = data[i][attendIdx];
      const availability = data[i][availIdx];
      const playerName = data[i][nameIdx];
      
      // If host didn't update Attended field, set it to match Availability
      if (!attended || attended === '' || attended === null) {
        const attendedValue = availability === 'Yes' ? 'Yes' : 'No';
        sheet.getRange(i + 1, attendIdx + 1).setValue(attendedValue);
        updatedCount++;
        Logger.log(`Set Attended="${attendedValue}" for ${playerName} on ${yesterdayStr} (Availability was "${availability}")`);
      }
    }
  }
  
  Logger.log(`Auto-set complete: Updated ${updatedCount} records for session on ${yesterdayStr}`);
}
