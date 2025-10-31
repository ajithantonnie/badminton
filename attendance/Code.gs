// Send daily attendance reminder to all players
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
  var tomorrowStr = Utilities.formatDate(tomorrow, 'Asia/Kolkata', 'yyyy-MM-dd');

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
  const today = Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM-dd');
  
  // Check if today's date already exists
  const dates = sheet.getRange("A2:A").getValues().flat();
  if (dates.includes(today)) {
    Logger.log(`Date ${today} already exists in sheet. Skipping.`);
    return;
  }
  
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

// Function to check warnings and inactivity
function checkWarningsAndInactivity() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Attendance');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Get column indices
  const dateIdx = headers.indexOf("Date");
  const nameIdx = headers.indexOf("Name");
  const availIdx = headers.indexOf("Availability");
  const reasonIdx = headers.indexOf("Reason");
  const attendIdx = headers.indexOf("Attended");
  const warnIdx = headers.indexOf("Warning Count");
  const missedIdx = headers.indexOf("Missed Days");
  const autoRemoveIdx = headers.indexOf("AutoRemove");
  
  // Get today's date in IST
  const now = new Date();
  const today = Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM-dd');
  const playerHistory = {};
  
  // Process all rows (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = row[dateIdx];
    const name = row[nameIdx];
    const email = row[emailIdx];
    const avail = row[availIdx];
    const reason = row[reasonIdx];
    const attended = row[attendIdx];
    const warnings = row[warnIdx] || 0;
    // Initialize player history if needed
    if (!playerHistory[name]) {
      playerHistory[name] = [];
    }
    // Add to player history
    playerHistory[name].push({
      date: date,
      avail: avail,
      reason: reason,
      attended: attended,
      rowIndex: i + 1 // +1 because we skipped header
    });
    // Check for today's records and update warnings
    if (date === today && avail === "Yes" && attended === "No") {
      const newWarn = warnings + 1;
      sheet.getRange(i + 1, warnIdx + 1).setValue(newWarn);
      // Send warning notification to player
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
  var tomorrowStr = Utilities.formatDate(tomorrow, 'Asia/Kolkata', 'yyyy-MM-dd');
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
  var subject = "Badminton Club: Tomorrow's Available Players";
  var body = "Players who have marked 'Yes' for tomorrow's availability:\n\n" + yesNames.join("\n");
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
function sendCelebrationIfFourAvailable() {
  var now = new Date();
  var tomorrow = new Date(now.getTime());
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowStr = Utilities.formatDate(tomorrow, 'Asia/Kolkata', 'yyyy-MM-dd');
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
  if (yesPlayers.length === 4) {
    var names = yesPlayers.map(p => p.name).join(", ");
    var subject = "Hurray! 4 members have marked for availability tomorrow";
    var body = `Congratulations! The following 4 members have marked 'Yes' for tomorrow: ${names}`;
    yesPlayers.forEach(p => {
      if (p.email) {
        MailApp.sendEmail({
          to: p.email,
          subject: subject,
          body: body
        });
      }
    });
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
        e.parameter.availability,
        e.parameter.reason
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
        e.parameter.newPlayerEmail
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
function submitPlayerAttendance(playerName, availability, reason) {
  // Check cutoff time (10:30 PM IST)
  const now = new Date();
  const hour = parseInt(Utilities.formatDate(now, 'Asia/Kolkata', 'HH'), 10);
  const minute = parseInt(Utilities.formatDate(now, 'Asia/Kolkata', 'mm'), 10);
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

  // Get tomorrow's date in IST
  const tomorrow = new Date(now.getTime());
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, 'Asia/Kolkata', 'yyyy-MM-dd');

  // Get player email from Players sheet
  const players = getPlayersList();
  // Find player by name and email, case-insensitive and trimmed
  let playerObj = null;
  let normalizedInputName = playerName ? playerName.trim().toLowerCase() : '';
  // If reason is actually the email (from frontend), use it
  let normalizedInputEmail = reason && reason.includes('@') ? reason.trim().toLowerCase() : null;
  for (let p of players) {
    let pName = p.name ? p.name.trim().toLowerCase() : '';
    let pEmail = p.email ? p.email.trim().toLowerCase() : '';
    if (normalizedInputEmail) {
      if (pName === normalizedInputName && pEmail === normalizedInputEmail) {
        playerObj = p;
        break;
      }
    } else {
      if (pName === normalizedInputName) {
        playerObj = p;
        break;
      }
    }
  }
  if (!playerObj) {
    // List available player names/emails for debugging
    let availablePlayers = players.map(p => `${p.name} <${p.email}>`).join(', ');
    return { success: false, message: `Invalid player name or email. Available: ${availablePlayers}` };
  }
  const playerEmail = playerObj.email;

  // Check if entry already exists (compare name and email separately)
  for (let i = 1; i < data.length; i++) {
    const sheetName = data[i][nameIdx] ? String(data[i][nameIdx]).trim() : '';
    const sheetEmail = data[i][emailIdx] ? String(data[i][emailIdx]).trim().toLowerCase() : '';
    if (data[i][dateIdx] === tomorrowStr && sheetName === playerName.trim() && sheetEmail === playerEmail.trim().toLowerCase()) {
      // Update existing entry
      sheet.getRange(i + 1, availIdx + 1).setValue(availability);
      sheet.getRange(i + 1, reasonIdx + 1).setValue(reason || '');
      // Do NOT set Attended here; host or scheduled function will set it
      return { success: true, message: 'Availability updated successfully' };
    }
  }

  // Create new entry
  const newRow = [
    tomorrowStr,
    playerName,
    playerEmail,
    availability,
    reason || '',
    '', // Attended left empty for host/scheduled function
    0,  // Warning Count
    0,  // Missed Days
    ''  // AutoRemove
  ];

  sheet.appendRow(newRow);
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
  const today = Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM-dd');

  // Find the record for player (match both name and email, ignore case and extra spaces)
  const targetName = playerName ? String(playerName).trim() : '';
  const targetEmail = playerEmail ? String(playerEmail).trim().toLowerCase() : '';
  let found = false;
  for (let i = 1; i < data.length; i++) {
    let rawDate = data[i][dateIdx];
    let formattedRowDate = '';
    // Try to format as yyyy-MM-dd regardless of type
    try {
      formattedRowDate = Utilities.formatDate(new Date(rawDate), 'Asia/Kolkata', 'yyyy-MM-dd');
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

      // Check for warning condition
      let currentWarnings = data[i][warnIdx] || 0;
      let newWarnings = currentWarnings;
      let warningGiven = false;
      // Case 1: Player marked Yes, host updates to No (did not attend)
      if (data[i][availIdx] === "Yes" && attended === "No") {
        newWarnings = currentWarnings + 1;
        sheet.getRange(i + 1, warnIdx + 1).setValue(newWarnings);
        warningGiven = true;
      }
      // Case 2: Player forgot to mark No with reason, host updates to Yes
      if (data[i][availIdx] === "" && attended === "Yes") {
        newWarnings = currentWarnings + 1;
        sheet.getRange(i + 1, warnIdx + 1).setValue(newWarnings);
        Logger.log(`Warning: Player ${rowName} did not mark No with reason, host updated to Yes.`);
        warningGiven = true;
      }
      // Send warning email if warning was given
      if (warningGiven && rowEmail) {
        MailApp.sendEmail({
          to: rowEmail,
          subject: `Badminton Club: Warning Received`,
          body: `Dear ${rowName},\n\nYou have received a warning for attendance update by host. Total warnings: ${newWarnings}.\n\nPlease ensure you mark your attendance correctly and attend if you mark 'Yes'.`
        });
      }
      if (newWarnings >= 5) {
        sheet.getRange(i + 1, autoRemoveIdx + 1).setValue("Remove");
        Logger.log(`Player ${rowName} marked for removal due to 5 warnings (host update after missed No with reason).`);
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

function setupFirstHost(firstName, firstEmail, firstPassword) {
  if (!firstName || !firstEmail || !firstPassword) {
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

  // Add to players sheet with email
  const playersSheet = getPlayersSheet();
  playersSheet.appendRow([firstName, firstEmail]);

  // Add to hosts sheet (store hash and email)
  const hostsSheet = getHostsSheet();
  hostsSheet.appendRow([firstName, firstPassword, 'Host', true, firstEmail]);

  return {
    success: true,
    message: `Welcome ${firstName}! You are now the main host.`,
    host: { name: firstName, role: 'Host', isAdmin: true, email: firstEmail }
  };
}

function addPlayer(adminName, adminPassword, newPlayerName, newPlayerEmail) {
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
  const players = getPlayersList().map(p => p.name);
  if (players.includes(newPlayerName) && players.includes(newPlayerEmail)) {
    return { success: false, message: 'Player already exists' };
  }
  const playersSheet = getPlayersSheet();
  playersSheet.appendRow([newPlayerName, newPlayerEmail]);
  return { success: true, message: `Player "${newPlayerName}" added successfully` };
}

function removePlayer(adminName, adminPassword, playerToRemove, playerToRemoveEmail) {
  const admin = authenticateAdmin(adminName, adminPassword);
  if (!admin.success) {
    return admin;
  }
  
  const targetName = String(playerToRemove).trim().toLowerCase();
  const targetEmail = String(playerToRemoveEmail).trim().toLowerCase();
  
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
  const hosts = getHostsList();
  if (hosts.find(h => h.name === newHostName && h.email === newHostEmail)) {
    return { success: false, message: 'Host already exists' };
  }
  const hostsSheet = getHostsSheet();
  const isAdmin = newHostRole === 'Host';
  hostsSheet.appendRow([newHostName, newHostPassword, newHostRole, isAdmin, newHostEmail]);
  return { success: true, message: `${newHostRole} "${newHostName}" added successfully` };
}

function removeHost(adminName, adminPassword, hostToRemove, hostToRemoveEmail) {
  const admin = authenticateAdmin(adminName, adminPassword);
  if (!admin.success) {
    return admin;
  }
  if (hostToRemove === adminName) {
    return { success: false, message: 'You cannot remove yourself' };
  }
  const hostsSheet = getHostsSheet();
  const data = hostsSheet.getDataRange().getValues();
  let removed = false;
  let removedRole = "";
  
  const targetName = String(hostToRemove).trim().toLowerCase();
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
  
  const data = playersSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  return data.map(row => ({ name: row[0], email: row[1] })).filter(p => p.name && p.name.toString().trim() !== '');
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
    sheet.getRange(1, 1, 1, 2).setValues([['Name', 'Email']]);
  } else {
    // Ensure Email column exists
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes('Email')) {
      sheet.insertColumnAfter(1);
      sheet.getRange(1, 2).setValue('Email');
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
  
  // All triggers scheduled in IST (Asia/Kolkata)
  // dailySetup for next day at 8am IST
  ScriptApp.newTrigger('dailySetup')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  // checkWarningsAndInactivity at 11pm IST
  ScriptApp.newTrigger('checkWarningsAndInactivity')
    .timeBased()
    .atHour(23)
    .everyDays(1)
    .create();

  // sendAttendanceReminder at 9pm IST
  ScriptApp.newTrigger('sendAttendanceReminder')
    .timeBased()
    .atHour(21)
    .everyDays(1)
    .create();

  // sendAvailabilitySummary at 10:30pm IST
  ScriptApp.newTrigger('sendAvailabilitySummary')
    .timeBased()
    .atHour(22)
    .everyDays(1)
    .nearMinute(30)
    .create();

  // sendCelebrationIfFourAvailable: single time-based trigger every 10 minutes
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
    .create();


  Logger.log('All triggers set up successfully');
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
  const headers = attendanceSheet.getRange(1, 1, 1, attendanceSheet.getLastColumn()).getValues()[0];
  const requiredHeaders = ['Date', 'Name', 'Email', 'Availability', 'Reason', 'Attended', 'Warning Count', 'Missed Days', 'AutoRemove'];
  let needsUpdate = false;
  for (let i = 0; i < requiredHeaders.length; i++) {
    if (headers[i] !== requiredHeaders[i]) {
      needsUpdate = true;
      break;
    }
  }
  if (needsUpdate) {
    attendanceSheet.getRange(1, 1, 1, requiredHeaders.length).setValues([
      requiredHeaders
    ]);
  }
  
  // Create Players sheet
  getPlayersSheet();
  
  // Create Hosts sheet
  getHostsSheet();
  
  Logger.log('Sheets initialized successfully');
}

function autoSetAttendedToAvailability() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Attendance');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dateIdx = headers.indexOf("Date");
  const nameIdx = headers.indexOf("Name");
  const emailIdx = headers.indexOf("Email");
  const availIdx = headers.indexOf("Availability");
  const attendIdx = headers.indexOf("Attended");

  // Get yesterday's date in IST (host update window ends at midnight)
  const now = new Date();
  const istNow = new Date(now.getTime() + (5.5 * 60 * 60 * 1000));
  istNow.setDate(istNow.getDate() - 1);
  const yesterdayStr = Utilities.formatDate(istNow, 'Asia/Kolkata', 'yyyy-MM-dd');

  for (let i = 1; i < data.length; i++) {
    if (data[i][dateIdx] === yesterdayStr) {
      const attended = data[i][attendIdx];
      const availability = data[i][availIdx];
      if (!attended || attended === '') {
        // Set Attended to match Availability
        sheet.getRange(i + 1, attendIdx + 1).setValue(availability === 'Yes' ? 'Yes' : 'No');
      }
    }
  }
}
