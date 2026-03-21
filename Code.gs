// ============================================================
// JETGRID INDIA — GOOGLE APPS SCRIPT BACKEND
// ============================================================
// Paste this entire file into Google Apps Script (Extensions → Apps Script)
// Then: Deploy → New deployment → Web app → Execute as Me → Anyone can access
// Copy the web app URL → paste into JETGRID_ENDPOINT in your HTML
// ============================================================

// ===== CONFIG — UPDATE THESE =====
var CONFIG = {
  // Google Sheet ID — get from the URL: docs.google.com/spreadsheets/d/THIS_PART/edit
  SHEET_ID: 'PASTE_YOUR_SHEET_ID_HERE',
  
  // Sheet tab name
  SHEET_NAME: 'Leads',
  
  // Email to receive lead alerts (use mihir.desai.aero@gmail.com until charters@jetgrid.in is set up)
  NOTIFY_EMAIL: 'mihir.desai.aero@gmail.com',
  
  // CallMeBot WhatsApp API key
  // To get: WhatsApp +34 644 66 84 11 → "I allow callmebot to send me messages" → they reply with your key
  CALLMEBOT_PHONE: '919579339605',
  CALLMEBOT_API_KEY: 'PASTE_YOUR_CALLMEBOT_KEY_HERE',
  
  // WhatsApp alerts on/off (set to false until you get CallMeBot key)
  WHATSAPP_ENABLED: false
};

// ===== SHEET COLUMNS (in order) =====
var COLUMNS = [
  'Timestamp',      // A
  'Name',           // B
  'Phone',          // C
  'Email',          // D
  'From',           // E
  'To',             // F
  'Date',           // G
  'Passengers',     // H
  'Enquiry Type',   // I
  'Message',        // J
  'WhatsApp Opt-in',// K
  'UTM Source',     // L
  'UTM Medium',     // M
  'Page URL',       // N
  'Referrer',       // O
  'Status',         // P
  'Notes'           // Q
];

// ============================================================
// MAIN HANDLER — receives POST from the website form
// ============================================================
function doPost(e) {
  try {
    var data;
    
    // Parse incoming data
    if (e.postData && e.postData.type === 'application/json') {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      data = e.parameter;
    } else {
      throw new Error('No data received');
    }
    
    // Get or create sheet
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME);
      // Add headers
      sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
      sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    // Format timestamp for IST
    var now = new Date();
    var ist = Utilities.formatDate(now, 'Asia/Kolkata', 'dd-MMM-yyyy hh:mm:ss a');
    
    // Build row
    var row = [
      ist,                                    // Timestamp
      data.name || '',                        // Name
      data.phone || '',                       // Phone
      data.email || '',                       // Email
      data.from || '',                        // From
      data.to || '',                          // To
      data.date || '',                        // Date
      data.passengers || '',                  // Passengers
      data.enquiry_type || '',                // Enquiry Type
      data.message || '',                     // Message
      data.whatsapp_optin || 'no',            // WhatsApp Opt-in
      data.utm_source || '',                  // UTM Source
      data.utm_medium || '',                  // UTM Medium
      data.page_url || '',                    // Page URL
      data.referrer || '',                    // Referrer
      'New',                                  // Status
      ''                                      // Notes
    ];
    
    // Append to sheet
    sheet.appendRow(row);
    
    // Color-code the Status cell
    var lastRow = sheet.getLastRow();
    colorStatusCell(sheet, lastRow);
    
    // Auto-resize columns on first few leads
    if (lastRow <= 5) {
      sheet.autoResizeColumns(1, COLUMNS.length);
    }
    
    // Send email alert
    sendEmailAlert(data, ist);
    
    // Send WhatsApp alert
    if (CONFIG.WHATSAPP_ENABLED) {
      sendWhatsAppAlert(data, ist);
    }
    
    // Log success
    logEvent('SUCCESS', 'Lead captured: ' + data.name + ' (' + data.enquiry_type + ')');
    
    // Return success
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success', message: 'Lead captured' })
    ).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    logEvent('ERROR', err.toString());
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// Also handle GET (for testing the endpoint in browser)
function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ status: 'ok', message: 'JetGrid backend is live. Use POST to submit leads.' })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// EMAIL ALERT
// ============================================================
function sendEmailAlert(data, timestamp) {
  var route = '';
  if (data.from && data.to) {
    route = data.from + ' → ' + data.to;
  } else if (data.from) {
    route = 'From ' + data.from;
  } else if (data.to) {
    route = 'To ' + data.to;
  }
  
  var subject = '✈️ New ' + (data.enquiry_type || 'Charter') + ' Lead — ' + data.name;
  
  var body = '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
  body += 'JETGRID — NEW LEAD\n';
  body += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n';
  body += 'Name:          ' + data.name + '\n';
  body += 'Phone:         ' + data.phone + '\n';
  body += 'Email:         ' + data.email + '\n';
  body += 'Type:          ' + (data.enquiry_type || 'Not specified') + '\n';
  if (route) body += 'Route:         ' + route + '\n';
  if (data.date) body += 'Date:          ' + data.date + '\n';
  if (data.passengers) body += 'Passengers:    ' + data.passengers + '\n';
  if (data.message) body += 'Message:       ' + data.message + '\n';
  body += 'WhatsApp OK:   ' + (data.whatsapp_optin || 'no') + '\n';
  body += '\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
  body += 'Received:      ' + timestamp + ' IST\n';
  if (data.utm_source) body += 'UTM Source:    ' + data.utm_source + '\n';
  if (data.utm_medium) body += 'UTM Medium:    ' + data.utm_medium + '\n';
  body += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
  
  try {
    MailApp.sendEmail({
      to: CONFIG.NOTIFY_EMAIL,
      subject: subject,
      body: body
    });
  } catch (err) {
    logEvent('EMAIL_ERROR', err.toString());
  }
}

// ============================================================
// WHATSAPP ALERT (via CallMeBot)
// ============================================================
function sendWhatsAppAlert(data, timestamp) {
  var route = (data.from && data.to) ? (data.from + ' → ' + data.to) : 'Not specified';
  
  var msg = '✈️ *JETGRID NEW LEAD*\n\n';
  msg += '*' + data.name + '*\n';
  msg += '📞 ' + data.phone + '\n';
  msg += '📧 ' + data.email + '\n';
  msg += '🏷️ ' + (data.enquiry_type || 'Charter') + '\n';
  if (data.from || data.to) msg += '🛫 ' + route + '\n';
  if (data.date) msg += '📅 ' + data.date + '\n';
  if (data.passengers) msg += '👥 ' + data.passengers + ' pax\n';
  msg += '\n⏰ ' + timestamp;
  
  var url = 'https://api.callmebot.com/whatsapp.php'
    + '?phone=' + CONFIG.CALLMEBOT_PHONE
    + '&text=' + encodeURIComponent(msg)
    + '&apikey=' + CONFIG.CALLMEBOT_API_KEY;
  
  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      logEvent('WHATSAPP_ERROR', 'HTTP ' + response.getResponseCode());
    }
  } catch (err) {
    logEvent('WHATSAPP_ERROR', err.toString());
  }
}

// ============================================================
// STATUS COLOR-CODING
// ============================================================
function colorStatusCell(sheet, row) {
  var statusCell = sheet.getRange(row, COLUMNS.indexOf('Status') + 1);
  var status = statusCell.getValue();
  
  var colors = {
    'New':               { bg: '#FADBD8', fg: '#922B21' },
    'Contacted':         { bg: '#FCF3CF', fg: '#7D6608' },
    'Quoted':            { bg: '#D4E6F1', fg: '#1B4F72' },
    'Negotiating':       { bg: '#FAE5D3', fg: '#A04000' },
    'Confirmed':         { bg: '#D5F5E3', fg: '#196F3D' },
    'Flew':              { bg: '#196F3D', fg: '#FFFFFF' },
    'Lost':              { bg: '#D5D8DC', fg: '#616A6B' },
    'Jet Card Interest': { bg: '#E8DAEF', fg: '#6C3483' }
  };
  
  var c = colors[status];
  if (c) {
    statusCell.setBackground(c.bg).setFontColor(c.fg).setFontWeight('bold');
  }
}

// Re-color all status cells (run manually if you batch-update statuses)
function recolorAllStatuses() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = sheet.getLastRow();
  for (var i = 2; i <= lastRow; i++) {
    colorStatusCell(sheet, i);
  }
}

// ============================================================
// DAILY SUMMARY EMAIL — set up a daily trigger for this
// ============================================================
function sendDailySummary() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  
  // Count by status
  var counts = {};
  var todayLeads = [];
  var today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MMM-yyyy');
  
  for (var i = 1; i < data.length; i++) {
    var status = data[i][COLUMNS.indexOf('Status')] || 'Unknown';
    counts[status] = (counts[status] || 0) + 1;
    
    // Check if lead came in today
    if (String(data[i][0]).indexOf(today) === 0) {
      todayLeads.push(data[i][1] + ' (' + data[i][8] + ')');
    }
  }
  
  var total = data.length - 1; // minus header
  
  var body = '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
  body += 'JETGRID — DAILY LEAD SUMMARY\n';
  body += today + '\n';
  body += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n';
  body += 'Total leads: ' + total + '\n\n';
  
  body += 'BY STATUS:\n';
  var statusOrder = ['New', 'Contacted', 'Quoted', 'Negotiating', 'Confirmed', 'Flew', 'Lost', 'Jet Card Interest'];
  statusOrder.forEach(function(s) {
    if (counts[s]) body += '  ' + s + ': ' + counts[s] + '\n';
  });
  
  if (todayLeads.length > 0) {
    body += '\nTODAY\'S LEADS (' + todayLeads.length + '):\n';
    todayLeads.forEach(function(l) { body += '  • ' + l + '\n'; });
  } else {
    body += '\nNo new leads today.\n';
  }
  
  body += '\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
  
  MailApp.sendEmail({
    to: CONFIG.NOTIFY_EMAIL,
    subject: '📊 JetGrid Daily Summary — ' + today + ' (' + (todayLeads.length) + ' new)',
    body: body
  });
}

// ============================================================
// SETUP — Run this once to create headers and daily trigger
// ============================================================
function initialSetup() {
  // Create sheet with headers if it doesn't exist
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }
  
  // Set headers
  sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
  sheet.getRange(1, 1, 1, COLUMNS.length)
    .setFontWeight('bold')
    .setBackground('#0A1628')
    .setFontColor('#F4F1EA');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, COLUMNS.length);
  
  // Set up daily summary trigger at 9 PM IST
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'sendDailySummary') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  ScriptApp.newTrigger('sendDailySummary')
    .timeBased()
    .atHour(21) // 9 PM
    .everyDays(1)
    .inTimezone('Asia/Kolkata')
    .create();
  
  Logger.log('✅ Setup complete. Sheet headers created, daily summary trigger set for 9 PM IST.');
}

// ============================================================
// ERROR LOGGING
// ============================================================
function logEvent(type, message) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var logSheet = ss.getSheetByName('Logs');
    if (!logSheet) {
      logSheet = ss.insertSheet('Logs');
      logSheet.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Type', 'Message']]);
      logSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    }
    var ist = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MMM-yyyy hh:mm:ss a');
    logSheet.appendRow([ist, type, message]);
  } catch (e) {
    Logger.log('Log error: ' + e.toString());
  }
}
