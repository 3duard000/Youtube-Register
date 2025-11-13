/**
 * Sunday Service Registration System - ENHANCED VERSION
 * Smart Sunday detection, improved dashboards, automated updates
 * 
 * SETUP: Run setupAllSheets() once, then run setupAutomatedTriggers()
 */

// ==================== CONFIGURATION ====================

// COMMUNITY ROUTING - Each community has its own Google Sheet
const COMMUNITY_SHEETS = {
  'Belvedere': '1U9VS9kEH7DGD_gbwqo1bzVTfgHgL-WS3mwQlysoNGJE',
  'New Jersey': '1U9VS9kEH7DGD_gbwqo1bzVTfgHgL-WS3mwQlysoNGJE'  // Change this to different sheet ID if needed
};

// Central tracking sheet (optional - tracks all communities)
const CENTRAL_SHEET_ID = '1Z7JrY-7US55iUnNmyXB3m5hKFGoWpR0bGMW1XXRgWnQ';

// Sheet tab names (same for all community sheets)
const SHEET_NAME = 'Registrations';
const MONTHLY_SHEET_NAME = 'Current Month';
const YEARLY_SHEET_NAME = 'Yearly Summary';
const SUNDAY_SHEET_NAME = 'Sunday Breakdown';

// Live stream link
const LIVE_STREAM_LINK = 'https://www.youtube.com/live/xU2alDtuXkw';

// Time after which we switch to next Sunday (in 24-hour format)
const SERVICE_END_HOUR = 14; // 2:00 PM

// ==================== WEB APP FUNCTIONS ====================

/**
 * Serves the HTML web app
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sunday Service Registration')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://www.gstatic.com/images/branding/product/1x/forms_48dp.png');
}

/**
 * Returns the list of available communities
 */
function getCommunities() {
  return Object.keys(COMMUNITY_SHEETS);
}

/**
 * Returns the live stream link to the web app
 */
function getLiveStreamLink() {
  return LIVE_STREAM_LINK;
}

/**
 * Calculates the appropriate Sunday date for registration
 * If it's Sunday before SERVICE_END_HOUR, show today
 * Otherwise, show next Sunday
 */
function getNextSunday() {
  var now = new Date();
  var dayOfWeek = now.getDay(); // 0 = Sunday, 6 = Saturday
  var currentHour = now.getHours();
  
  var targetDate = new Date(now);
  
  if (dayOfWeek === 0) {
    // It's Sunday
    if (currentHour < SERVICE_END_HOUR) {
      // Before service end time - register for TODAY
      Logger.log('It\'s Sunday before ' + SERVICE_END_HOUR + ':00, showing today');
    } else {
      // After service end time - register for NEXT Sunday
      Logger.log('It\'s Sunday after ' + SERVICE_END_HOUR + ':00, showing next Sunday');
      targetDate.setDate(now.getDate() + 7);
    }
  } else {
    // Not Sunday - calculate next Sunday
    var daysUntilSunday = (7 - dayOfWeek) % 7;
    if (daysUntilSunday === 0) daysUntilSunday = 7;
    targetDate.setDate(now.getDate() + daysUntilSunday);
    Logger.log('Today is not Sunday, showing upcoming Sunday');
  }
  
  var options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  
  return {
    formatted: targetDate.toLocaleDateString('en-US', options),
    date: targetDate.toISOString(),
    dateOnly: Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    isToday: dayOfWeek === 0 && currentHour < SERVICE_END_HOUR
  };
}

// ==================== REGISTRATION FUNCTIONS ====================

/**
 * Saves registration data to community-specific Google Sheet and sends confirmation email
 * AUTOMATICALLY UPDATES ALL TABS
 */
function saveRegistration(formData) {
  try {
    Logger.log('========================================');
    Logger.log('STARTING REGISTRATION PROCESS');
    Logger.log('========================================');
    Logger.log('Form data received: ' + JSON.stringify(formData));
    
    // Get the community
    var community = formData.community || 'Belvedere';
    Logger.log('Community: ' + community);
    
    // Get the correct sheet ID for this community
    var sheetId = COMMUNITY_SHEETS[community];
    if (!sheetId) {
      throw new Error('Invalid community: ' + community);
    }
    
    Logger.log('Routing to sheet ID: ' + sheetId);
    
    var sheet = getOrCreateSheet(sheetId);
    var timestamp = new Date();
    var emails = [];
    
    Logger.log('Processing ' + formData.registrants.length + ' registrant(s)');
    
    // Add each person in the registration
    for (var i = 0; i < formData.registrants.length; i++) {
      var person = formData.registrants[i];
      Logger.log('Adding person ' + (i + 1) + ': ' + person.firstName + ' ' + person.lastName + ' (' + person.email + ')');
      
      // IMPORTANT: Column order must match headers
      // Columns: Timestamp, Community, First Name, Last Name, Email, Type, Session, Sunday Date
      sheet.appendRow([
        timestamp,              // Column A: Timestamp
        community,              // Column B: Community
        person.firstName,       // Column C: First Name
        person.lastName,        // Column D: Last Name
        person.email,           // Column E: Email
        person.type,            // Column F: Type
        formData.sessionInfo || 'Sunday Service',  // Column G: Session
        formData.sundayDate || ''  // Column H: Sunday Date
      ]);
      
      // Collect unique emails for sending confirmations
      if (person.email && emails.indexOf(person.email) === -1) {
        emails.push(person.email);
      }
    }
    
    Logger.log('✓ All registrants added to ' + community + ' sheet');
    
    // Sort by timestamp (newest first)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var range = sheet.getRange(2, 1, lastRow - 1, 8);
      range.sort({column: 1, ascending: false});
      Logger.log('✓ Data sorted by timestamp');
    }
    
    // AUTOMATICALLY UPDATE ALL TABS FOR THIS COMMUNITY
    Logger.log('Updating dashboard tabs for ' + community + '...');
    updateMonthlySummary(sheetId);
    Logger.log('✓ Current Month tab updated');
    
    updateYearlySummary(sheetId);
    Logger.log('✓ Yearly Summary tab updated');
    
    updateSundayBreakdown(sheetId);
    Logger.log('✓ Sunday Breakdown tab updated');
    
    // Send confirmation emails
    Logger.log('Sending confirmation emails to ' + emails.length + ' recipient(s)...');
    var emailsSent = 0;
    var emailErrors = [];
    
    for (var i = 0; i < emails.length; i++) {
      var email = emails[i];
      Logger.log('Attempting to send email to: ' + email);
      
      try {
        if (sendConfirmationEmail(email, formData.registrants, formData.sundayDate, community)) {
          emailsSent++;
          Logger.log('✓ Email sent to: ' + email);
        } else {
          emailErrors.push(email);
          Logger.log('✗ Failed to send email to: ' + email);
        }
      } catch (emailError) {
        emailErrors.push(email);
        Logger.log('✗ Error sending email to ' + email + ': ' + emailError.toString());
      }
    }
    
    Logger.log('========================================');
    Logger.log('REGISTRATION COMPLETE');
    Logger.log('Community: ' + community);
    Logger.log('Total registrants: ' + formData.registrants.length);
    Logger.log('Emails sent: ' + emailsSent + '/' + emails.length);
    if (emailErrors.length > 0) {
      Logger.log('Email errors for: ' + emailErrors.join(', '));
    }
    Logger.log('========================================');
    
    return {
      success: true,
      message: 'Registration successful for ' + community + '! ' + (emailsSent > 0 ? 'Check your email for confirmation.' : ''),
      count: formData.registrants.length,
      community: community,
      liveStreamLink: LIVE_STREAM_LINK,
      emailsSent: emailsSent,
      emailErrors: emailErrors.length
    };
    
  } catch (error) {
    Logger.log('========================================');
    Logger.log('✗ REGISTRATION ERROR');
    Logger.log('Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    Logger.log('========================================');
    
    return {
      success: false,
      message: 'There was an error processing your registration. Please try again.',
      error: error.toString()
    };
  }
}

/**
 * Sends confirmation email to registrant
 */
function sendConfirmationEmail(email, registrants, sundayDate, community) {
  try {
    // Find all people with this email
    var people = [];
    for (var i = 0; i < registrants.length; i++) {
      if (registrants[i].email === email) {
        people.push(registrants[i]);
      }
    }
    
    if (people.length === 0) {
      Logger.log('No registrants found for email: ' + email);
      return false;
    }
    
    var names = [];
    for (var i = 0; i < people.length; i++) {
      names.push(people[i].firstName + ' ' + people[i].lastName);
    }
    var namesString = names.join(', ');
    
    var types = [];
    for (var i = 0; i < people.length; i++) {
      types.push(people[i].type);
    }
    var typesString = types.join(', ');
    
    var subject = '✅ Sunday Service Registration Confirmed - ' + community + ' - ' + sundayDate;
    
    var htmlBody = 
      '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">' +
      '<div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 30px; text-align: center; border-radius: 10px 10px 0 0;">' +
      '<h1 style="color: white; margin: 0;">🙏 Registration Confirmed!</h1>' +
      '</div>' +
      '<div style="background: #f8f9fa; padding: 30px; border-radius: 0 0 10px 10px;">' +
      '<p style="font-size: 16px; color: #333;">Dear ' + people[0].firstName + ',</p>' +
      '<p style="font-size: 16px; color: #333;">Thank you for registering for our Sunday Service!</p>' +
      '<div style="background: white; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #667eea;">' +
      '<h3 style="color: #667eea; margin-top: 0;">📅 Service Details</h3>' +
      '<p style="margin: 10px 0;"><strong>Community:</strong> ' + community + '</p>' +
      '<p style="margin: 10px 0;"><strong>Date:</strong> ' + sundayDate + '</p>' +
      '<p style="margin: 10px 0;"><strong>Registered:</strong> ' + namesString + '</p>' +
      '<p style="margin: 10px 0;"><strong>Type:</strong> ' + typesString + '</p>' +
      '</div>' +
      '<div style="text-align: center; margin: 30px 0;">' +
      '<a href="' + LIVE_STREAM_LINK + '" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px 30px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">📺 Join Live Stream</a>' +
      '</div>' +
      '<p style="font-size: 14px; color: #666;">We look forward to worshiping with you this Sunday!</p>' +
      '<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">' +
      '<p style="font-size: 12px; color: #999; text-align: center;">This is an automated confirmation email for your Sunday Service registration at ' + community + '.</p>' +
      '</div>' +
      '</div>';
    
    var plainBody = 
      'Sunday Service Registration Confirmed\n\n' +
      'Dear ' + people[0].firstName + ',\n\n' +
      'Thank you for registering for our Sunday Service!\n\n' +
      'Service Details:\n' +
      'Community: ' + community + '\n' +
      'Date: ' + sundayDate + '\n' +
      'Registered: ' + namesString + '\n' +
      'Type: ' + typesString + '\n\n' +
      'Live Stream Link: ' + LIVE_STREAM_LINK + '\n\n' +
      'We look forward to worshiping with you this Sunday!\n\n' +
      'This is an automated confirmation email for ' + community + '.';
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      body: plainBody
    });
    
    return true;
    
  } catch (error) {
    Logger.log('Email sending error: ' + error.toString());
    return false;
  }
}

// ==================== SETUP FUNCTIONS ====================

/**
 * MAIN SETUP FUNCTION - RUN THIS FOR EACH COMMUNITY
 * Creates all 4 tabs with proper formatting
 * 
 * Usage: setupAllSheets('Belvedere') or setupAllSheets('New Jersey')
 */
function setupAllSheets(communityName) {
  // If no community specified, show selection dialog
  if (!communityName) {
    var ui = SpreadsheetApp.getUi();
    var communities = Object.keys(COMMUNITY_SHEETS);
    var communityList = communities.join('\\n');
    
    var response = ui.prompt(
      'Select Community',
      'Enter the community name to set up:\\n\\n' + communityList + '\\n\\nCommunity:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      communityName = response.getResponseText().trim();
    } else {
      Logger.log('Setup cancelled by user');
      return { success: false, message: 'Setup cancelled' };
    }
  }
  
  // Validate community
  if (!COMMUNITY_SHEETS[communityName]) {
    Logger.log('Invalid community: ' + communityName);
    Browser.msgBox(
      'Invalid Community',
      'Community "' + communityName + '" not found. Valid communities: ' + Object.keys(COMMUNITY_SHEETS).join(', '),
      Browser.Buttons.OK
    );
    return { success: false, message: 'Invalid community' };
  }
  
  var sheetId = COMMUNITY_SHEETS[communityName];
  
  Logger.log('╔═══════════════════════════════════════╗');
  Logger.log('║  SUNDAY SERVICE REGISTRATION SETUP     ║');
  Logger.log('║  Community: ' + communityName.padEnd(25) + '║');
  Logger.log('╚═══════════════════════════════════════╝');
  
  try {
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    Logger.log('✓ Connected to spreadsheet: ' + spreadsheet.getName());
    Logger.log('');
    
    // STEP 1: Setup main Registrations sheet
    Logger.log('STEP 1: Setting up Registrations tab...');
    Logger.log('----------------------------------------');
    
    var mainSheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (mainSheet) {
      Logger.log('⚠ Registrations tab already exists');
      var response = Browser.msgBox(
        'Registrations Tab Exists',
        'The Registrations tab already exists for ' + communityName + '. Do you want to recreate it? This will DELETE all existing data!',
        Browser.Buttons.YES_NO
      );
      
      if (response === 'yes') {
        spreadsheet.deleteSheet(mainSheet);
        Logger.log('✓ Old Registrations tab deleted');
        mainSheet = spreadsheet.insertSheet(SHEET_NAME);
        formatMainSheet(mainSheet);
        Logger.log('✓ New Registrations tab created');
      } else {
        Logger.log('ℹ Keeping existing Registrations tab');
      }
    } else {
      mainSheet = spreadsheet.insertSheet(SHEET_NAME);
      formatMainSheet(mainSheet);
      Logger.log('✓ Registrations tab created');
    }
    
    Logger.log('');
    
    // STEP 2: Setup Current Month tab
    Logger.log('STEP 2: Setting up Current Month tab...');
    Logger.log('----------------------------------------');
    
    var monthlySheet = spreadsheet.getSheetByName(MONTHLY_SHEET_NAME);
    if (monthlySheet) {
      spreadsheet.deleteSheet(monthlySheet);
      Logger.log('✓ Old Current Month tab deleted');
    }
    
    monthlySheet = spreadsheet.insertSheet(MONTHLY_SHEET_NAME);
    Logger.log('✓ Current Month tab created');
    updateMonthlySummary(sheetId);
    Logger.log('✓ Current Month tab populated');
    
    Logger.log('');
    
    // STEP 3: Setup Yearly Summary tab
    Logger.log('STEP 3: Setting up Yearly Summary tab...');
    Logger.log('----------------------------------------');
    
    var yearlySheet = spreadsheet.getSheetByName(YEARLY_SHEET_NAME);
    if (yearlySheet) {
      spreadsheet.deleteSheet(yearlySheet);
      Logger.log('✓ Old Yearly Summary tab deleted');
    }
    
    yearlySheet = spreadsheet.insertSheet(YEARLY_SHEET_NAME);
    Logger.log('✓ Yearly Summary tab created');
    updateYearlySummary(sheetId);
    Logger.log('✓ Yearly Summary tab populated');
    
    Logger.log('');
    
    // STEP 4: Setup Sunday Breakdown tab
    Logger.log('STEP 4: Setting up Sunday Breakdown tab...');
    Logger.log('----------------------------------------');
    
    var sundaySheet = spreadsheet.getSheetByName(SUNDAY_SHEET_NAME);
    if (sundaySheet) {
      spreadsheet.deleteSheet(sundaySheet);
      Logger.log('✓ Old Sunday Breakdown tab deleted');
    }
    
    sundaySheet = spreadsheet.insertSheet(SUNDAY_SHEET_NAME);
    Logger.log('✓ Sunday Breakdown tab created');
    updateSundayBreakdown(sheetId);
    Logger.log('✓ Sunday Breakdown tab populated');
    
    Logger.log('');
    Logger.log('╔═══════════════════════════════════════╗');
    Logger.log('║           SETUP COMPLETE! ✓            ║');
    Logger.log('║  Community: ' + communityName.padEnd(25) + '║');
    Logger.log('╚═══════════════════════════════════════╝');
    Logger.log('');
    Logger.log('Sheet is ready at:');
    Logger.log(spreadsheet.getUrl());
    Logger.log('');
    Logger.log('All 4 tabs created:');
    Logger.log('  1. Registrations (main data)');
    Logger.log('  2. Current Month (auto-updating)');
    Logger.log('  3. Yearly Summary (auto-updating)');
    Logger.log('  4. Sunday Breakdown (auto-updating)');
    Logger.log('');
    
    Browser.msgBox(
      'Setup Complete!',
      'All tabs created successfully for ' + communityName + '\\n\\n' +
      '✓ Registrations tab\\n' +
      '✓ Current Month dashboard\\n' +
      '✓ Yearly Summary dashboard\\n' +
      '✓ Sunday Breakdown dashboard\\n\\n' +
      'Sheet URL: ' + spreadsheet.getUrl(),
      Browser.Buttons.OK
    );
    
    return {
      success: true,
      message: 'Setup completed successfully for ' + communityName,
      url: spreadsheet.getUrl()
    };
    
  } catch (error) {
    Logger.log('');
    Logger.log('╔═══════════════════════════════════════╗');
    Logger.log('║         SETUP FAILED! ✗                ║');
    Logger.log('╚═══════════════════════════════════════╝');
    Logger.log('');
    Logger.log('Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    
    Browser.msgBox(
      'Setup Failed',
      'There was an error during setup for ' + communityName + ': ' + error.toString(),
      Browser.Buttons.OK
    );
    
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * Formats the main Registrations sheet
 */
function formatMainSheet(sheet) {
  Logger.log('  → Adding headers...');
  
  var headers = [
    'Timestamp',      // Column A
    'Community',      // Column B
    'First Name',     // Column C
    'Last Name',      // Column D
    'Email',          // Column E
    'Type',           // Column F
    'Session',        // Column G
    'Sunday Date'     // Column H
  ];
  
  sheet.appendRow(headers);
  
  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold')
            .setFontSize(12)
            .setBackground('#667eea')
            .setFontColor('#ffffff')
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');
  
  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1);
  
  // Set column widths
  sheet.setColumnWidth(1, 180);  // Timestamp
  sheet.setColumnWidth(2, 130);  // Community
  sheet.setColumnWidth(3, 120);  // First Name
  sheet.setColumnWidth(4, 120);  // Last Name
  sheet.setColumnWidth(5, 250);  // Email
  sheet.setColumnWidth(6, 100);  // Type
  sheet.setColumnWidth(7, 180);  // Session
  sheet.setColumnWidth(8, 180);  // Sunday Date
  
  // Add border to headers
  headerRange.setBorder(
    true, true, true, true, true, true,
    '#ffffff',
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  // Data validation for Type column
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Member', 'Guest'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 6, 1000, 1).setDataValidation(typeRule);
  
  // Apply row banding
  var dataRange = sheet.getRange(2, 1, 1000, 8);
  var banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  banding.setHeaderRowColor('#667eea')
         .setFirstRowColor('#f3f4ff')
         .setSecondRowColor('#ffffff');
  
  // Center align Type column
  sheet.getRange(2, 6, 1000, 1).setHorizontalAlignment('center');
  
  // Format timestamp column
  sheet.getRange(2, 1, 1000, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  
  // Add summary section
  addMainSummarySection(sheet);
  
  Logger.log('  → Main sheet formatting complete');
}

/**
 * Adds summary section to main sheet
 */
function addMainSummarySection(sheet) {
  var summaryStartCol = 10; // Column J (moved because of Community column)
  
  sheet.getRange(1, summaryStartCol).setValue('REGISTRATION SUMMARY');
  sheet.getRange(1, summaryStartCol, 1, 2).merge()
       .setFontWeight('bold')
       .setFontSize(12)
       .setBackground('#764ba2')
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  // Summary formulas - adjusted for Community column (Email is now column E)
  var summaryData = [
    ['Total Registrations', '=COUNTA(C2:C)'],
    ['Total Members', '=COUNTIF(F2:F,"Member")'],
    ['Total Guests', '=COUNTIF(F2:F,"Guest")'],
    ['Unique Emails', '=COUNTA(UNIQUE(FILTER(E2:E,E2:E<>"")))'],
    ['Last Registration', '=IF(COUNTA(A2:A)>0,TEXT(MAX(A2:A),"yyyy-mm-dd hh:mm"),"No data")'],
    ['', ''],
    ['🔄 Refresh Dashboard', 'Click to update all tabs']
  ];
  
  for (var i = 0; i < summaryData.length; i++) {
    var rowNum = 3 + i;
    sheet.getRange(rowNum, summaryStartCol).setValue(summaryData[i][0])
         .setFontWeight('bold')
         .setBackground('#f3f4ff');
    
    if (i < summaryData.length - 1) {
      sheet.getRange(rowNum, summaryStartCol + 1).setFormula(summaryData[i][1])
           .setHorizontalAlignment('center')
           .setBackground('#ffffff');
    } else {
      // Add button-like styling for refresh
      sheet.getRange(rowNum, summaryStartCol, 1, 2).merge()
           .setBackground('#28a745')
           .setFontColor('#ffffff')
           .setHorizontalAlignment('center')
           .setFontWeight('bold');
    }
  }
  
  sheet.setColumnWidth(summaryStartCol, 180);
  sheet.setColumnWidth(summaryStartCol + 1, 150);
  
  sheet.getRange(1, summaryStartCol, summaryData.length + 2, 2).setBorder(
    true, true, true, true, true, true,
    '#cccccc',
    SpreadsheetApp.BorderStyle.SOLID
  );
}

/**
 * Gets or creates the main sheet for a given sheet ID
 */
function getOrCreateSheet(sheetId) {
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    Logger.log('Main sheet not found, creating it...');
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    formatMainSheet(sheet);
  }
  
  return sheet;
}

// ==================== AUTO-UPDATE FUNCTIONS ====================

/**
 * Updates Current Month tab
 */
function updateMonthlySummary() {
  try {
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    var mainSheet = spreadsheet.getSheetByName(SHEET_NAME);
    var monthlySheet = spreadsheet.getSheetByName(MONTHLY_SHEET_NAME);
    
    if (!monthlySheet) {
      monthlySheet = spreadsheet.insertSheet(MONTHLY_SHEET_NAME);
    } else {
      monthlySheet.clear();
    }
    
    var now = new Date();
    var currentMonth = now.getMonth();
    var currentYear = now.getFullYear();
    
    // Title
    monthlySheet.getRange(1, 1, 1, 7).merge()
                .setValue('REGISTRATIONS FOR ' + getMonthName(currentMonth) + ' ' + currentYear)
                .setFontWeight('bold')
                .setFontSize(14)
                .setBackground('#667eea')
                .setFontColor('#ffffff')
                .setHorizontalAlignment('center');
    
    monthlySheet.setRowHeight(1, 50);
    
    // Headers
    var headers = ['Timestamp', 'First Name', 'Last Name', 'Email', 'Type', 'Session', 'Sunday Date'];
    monthlySheet.getRange(2, 1, 1, headers.length).setValues([headers])
                .setFontWeight('bold')
                .setBackground('#764ba2')
                .setFontColor('#ffffff')
                .setHorizontalAlignment('center');
    
    monthlySheet.setRowHeight(2, 40);
    monthlySheet.setFrozenRows(2);
    
    // Get data for current month
    var mainData = mainSheet.getDataRange().getValues();
    var monthlyData = [];
    
    for (var i = 1; i < mainData.length; i++) {
      var timestamp = new Date(mainData[i][0]);
      if (timestamp.getMonth() === currentMonth && timestamp.getFullYear() === currentYear) {
        monthlyData.push(mainData[i]);
      }
    }
    
    // Add data if exists
    if (monthlyData.length > 0) {
      monthlySheet.getRange(3, 1, monthlyData.length, 7).setValues(monthlyData);
      
      var dataRange = monthlySheet.getRange(3, 1, monthlyData.length, 7);
      var banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
      banding.setFirstRowColor('#f3f4ff').setSecondRowColor('#ffffff');
    }
    
    // Set column widths
    monthlySheet.setColumnWidth(1, 180);
    monthlySheet.setColumnWidth(2, 120);
    monthlySheet.setColumnWidth(3, 120);
    monthlySheet.setColumnWidth(4, 250);
    monthlySheet.setColumnWidth(5, 100);
    monthlySheet.setColumnWidth(6, 180);
    monthlySheet.setColumnWidth(7, 180);
    
    // Monthly summary
    var summaryStartRow = monthlyData.length + 4;
    monthlySheet.getRange(summaryStartRow, 1, 1, 2).merge()
                .setValue('MONTHLY SUMMARY')
                .setFontWeight('bold')
                .setFontSize(12)
                .setBackground('#764ba2')
                .setFontColor('#ffffff')
                .setHorizontalAlignment('center');
    
    var summaryData = [
      ['Total Registrations:', monthlyData.length],
      ['Total Members:', countByType(monthlyData, 'Member')],
      ['Total Guests:', countByType(monthlyData, 'Guest')],
      ['Unique Emails:', getUniqueEmails(monthlyData).length]
    ];
    
    for (var i = 0; i < summaryData.length; i++) {
      monthlySheet.getRange(summaryStartRow + i + 1, 1).setValue(summaryData[i][0])
                  .setFontWeight('bold')
                  .setBackground('#f3f4ff');
      monthlySheet.getRange(summaryStartRow + i + 1, 2).setValue(summaryData[i][1])
                  .setHorizontalAlignment('center')
                  .setBackground('#ffffff');
    }
    
  } catch (error) {
    Logger.log('Error updating monthly summary: ' + error.toString());
  }
}

/**
 * Updates Yearly Summary tab
 */
function updateYearlySummary() {
  try {
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    var mainSheet = spreadsheet.getSheetByName(SHEET_NAME);
    var yearlySheet = spreadsheet.getSheetByName(YEARLY_SHEET_NAME);
    
    if (!yearlySheet) {
      yearlySheet = spreadsheet.insertSheet(YEARLY_SHEET_NAME);
    } else {
      yearlySheet.clear();
    }
    
    var currentYear = new Date().getFullYear();
    
    // Title
    yearlySheet.getRange(1, 1, 1, 5).merge()
               .setValue('YEARLY REGISTRATION SUMMARY - ' + currentYear)
               .setFontWeight('bold')
               .setFontSize(14)
               .setBackground('#667eea')
               .setFontColor('#ffffff')
               .setHorizontalAlignment('center');
    
    yearlySheet.setRowHeight(1, 50);
    
    // Headers
    var headers = ['Month', 'Total Registrations', 'Members', 'Guests', 'Unique Emails'];
    yearlySheet.getRange(3, 1, 1, headers.length).setValues([headers])
               .setFontWeight('bold')
               .setBackground('#764ba2')
               .setFontColor('#ffffff')
               .setHorizontalAlignment('center');
    
    yearlySheet.setRowHeight(3, 40);
    
    // Get all data
    var mainData = mainSheet.getDataRange().getValues();
    
    // Calculate monthly stats
    var monthlyStats = [];
    for (var month = 0; month < 12; month++) {
      var monthData = [];
      
      for (var i = 1; i < mainData.length; i++) {
        var timestamp = new Date(mainData[i][0]);
        if (timestamp.getMonth() === month && timestamp.getFullYear() === currentYear) {
          monthData.push(mainData[i]);
        }
      }
      
      if (monthData.length > 0 || month <= new Date().getMonth()) {
        monthlyStats.push([
          getMonthName(month),
          monthData.length,
          countByType(monthData, 'Member'),
          countByType(monthData, 'Guest'),
          getUniqueEmails(monthData).length
        ]);
      }
    }
    
    // Add monthly data
    if (monthlyStats.length > 0) {
      yearlySheet.getRange(4, 1, monthlyStats.length, 5).setValues(monthlyStats);
      
      for (var i = 0; i < monthlyStats.length; i++) {
        var bgColor = i % 2 === 0 ? '#f3f4ff' : '#ffffff';
        yearlySheet.getRange(4 + i, 1, 1, 5).setBackground(bgColor);
      }
    }
    
    // Set column widths
    yearlySheet.setColumnWidth(1, 120);
    yearlySheet.setColumnWidth(2, 150);
    yearlySheet.setColumnWidth(3, 120);
    yearlySheet.setColumnWidth(4, 120);
    yearlySheet.setColumnWidth(5, 150);
    
    // Yearly totals
    var totalRow = 4 + monthlyStats.length + 1;
    yearlySheet.getRange(totalRow, 1).setValue('YEARLY TOTAL')
               .setFontWeight('bold')
               .setBackground('#764ba2')
               .setFontColor('#ffffff');
    
    var yearData = mainData.filter(function(row, index) {
      if (index === 0) return false;
      var timestamp = new Date(row[0]);
      return timestamp.getFullYear() === currentYear;
    });
    
    yearlySheet.getRange(totalRow, 2).setValue(yearData.length)
               .setFontWeight('bold')
               .setBackground('#f3f4ff')
               .setHorizontalAlignment('center');
    yearlySheet.getRange(totalRow, 3).setValue(countByType(yearData, 'Member'))
               .setFontWeight('bold')
               .setBackground('#f3f4ff')
               .setHorizontalAlignment('center');
    yearlySheet.getRange(totalRow, 4).setValue(countByType(yearData, 'Guest'))
               .setFontWeight('bold')
               .setBackground('#f3f4ff')
               .setHorizontalAlignment('center');
    yearlySheet.getRange(totalRow, 5).setValue(getUniqueEmails(yearData).length)
               .setFontWeight('bold')
               .setBackground('#f3f4ff')
               .setHorizontalAlignment('center');
    
  } catch (error) {
    Logger.log('Error updating yearly summary: ' + error.toString());
  }
}

/**
 * NEW: Updates Sunday Breakdown tab - Shows registrations grouped by Sunday dates
 */
function updateSundayBreakdown() {
  try {
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    var mainSheet = spreadsheet.getSheetByName(SHEET_NAME);
    var sundaySheet = spreadsheet.getSheetByName(SUNDAY_SHEET_NAME);
    
    if (!sundaySheet) {
      sundaySheet = spreadsheet.insertSheet(SUNDAY_SHEET_NAME);
    } else {
      sundaySheet.clear();
    }
    
    // Title
    sundaySheet.getRange(1, 1, 1, 6).merge()
               .setValue('📊 SUNDAY SERVICE REGISTRATION BREAKDOWN')
               .setFontWeight('bold')
               .setFontSize(14)
               .setBackground('#667eea')
               .setFontColor('#ffffff')
               .setHorizontalAlignment('center');
    
    sundaySheet.setRowHeight(1, 50);
    
    // Last updated timestamp
    sundaySheet.getRange(2, 1, 1, 6).merge()
               .setValue('Last Updated: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'))
               .setFontSize(10)
               .setFontColor('#666666')
               .setHorizontalAlignment('center');
    
    // Headers
    var headers = ['Sunday Date', 'Total Registrations', 'Members', 'Guests', 'Unique Emails', 'Details'];
    sundaySheet.getRange(4, 1, 1, headers.length).setValues([headers])
               .setFontWeight('bold')
               .setBackground('#764ba2')
               .setFontColor('#ffffff')
               .setHorizontalAlignment('center');
    
    sundaySheet.setRowHeight(4, 40);
    sundaySheet.setFrozenRows(4);
    
    // Get all data
    var mainData = mainSheet.getDataRange().getValues();
    
    // Group by Sunday date
    var sundayGroups = {};
    
    for (var i = 1; i < mainData.length; i++) {
      var sundayDate = mainData[i][6]; // Column G: Sunday Date
      
      if (sundayDate) {
        if (!sundayGroups[sundayDate]) {
          sundayGroups[sundayDate] = [];
        }
        sundayGroups[sundayDate].push(mainData[i]);
      }
    }
    
    // Convert to sorted array
    var sundayDates = Object.keys(sundayGroups).sort().reverse(); // Most recent first
    
    var sundayData = [];
    for (var i = 0; i < sundayDates.length; i++) {
      var date = sundayDates[i];
      var dateData = sundayGroups[date];
      
      sundayData.push([
        date,
        dateData.length,
        countByType(dateData, 'Member'),
        countByType(dateData, 'Guest'),
        getUniqueEmails(dateData).length,
        'View Details →'
      ]);
    }
    
    // Add data
    if (sundayData.length > 0) {
      sundaySheet.getRange(5, 1, sundayData.length, 6).setValues(sundayData);
      
      // Format data rows
      for (var i = 0; i < sundayData.length; i++) {
        var rowNum = 5 + i;
        var bgColor = i % 2 === 0 ? '#f3f4ff' : '#ffffff';
        
        sundaySheet.getRange(rowNum, 1, 1, 6).setBackground(bgColor);
        
        // Make Sunday date bold
        sundaySheet.getRange(rowNum, 1).setFontWeight('bold');
        
        // Center align numbers
        sundaySheet.getRange(rowNum, 2, 1, 4).setHorizontalAlignment('center');
        
        // Style details link
        sundaySheet.getRange(rowNum, 6)
                   .setFontColor('#667eea')
                   .setFontWeight('bold')
                   .setHorizontalAlignment('center');
      }
    } else {
      sundaySheet.getRange(5, 1, 1, 6).merge()
                 .setValue('No registrations yet')
                 .setHorizontalAlignment('center')
                 .setFontStyle('italic')
                 .setFontColor('#999999');
    }
    
    // Set column widths
    sundaySheet.setColumnWidth(1, 200); // Sunday Date
    sundaySheet.setColumnWidth(2, 150); // Total
    sundaySheet.setColumnWidth(3, 120); // Members
    sundaySheet.setColumnWidth(4, 120); // Guests
    sundaySheet.setColumnWidth(5, 150); // Emails
    sundaySheet.setColumnWidth(6, 150); // Details
    
    // Add summary section
    var summaryRow = 5 + sundayData.length + 2;
    
    sundaySheet.getRange(summaryRow, 1, 1, 3).merge()
               .setValue('📈 OVERALL STATISTICS')
               .setFontWeight('bold')
               .setFontSize(12)
               .setBackground('#764ba2')
               .setFontColor('#ffffff')
               .setHorizontalAlignment('center');
    
    var totalRegistrations = 0;
    var totalMembers = 0;
    var totalGuests = 0;
    
    for (var i = 0; i < sundayData.length; i++) {
      totalRegistrations += sundayData[i][1];
      totalMembers += sundayData[i][2];
      totalGuests += sundayData[i][3];
    }
    
    var overallStats = [
      ['Total Sundays', sundayDates.length, ''],
      ['Total Registrations', totalRegistrations, ''],
      ['Total Members', totalMembers, ''],
      ['Total Guests', totalGuests, ''],
      ['Average per Sunday', sundayDates.length > 0 ? Math.round(totalRegistrations / sundayDates.length) : 0, 'people']
    ];
    
    for (var i = 0; i < overallStats.length; i++) {
      sundaySheet.getRange(summaryRow + i + 1, 1).setValue(overallStats[i][0])
                 .setFontWeight('bold')
                 .setBackground('#f3f4ff');
      sundaySheet.getRange(summaryRow + i + 1, 2).setValue(overallStats[i][1])
                 .setHorizontalAlignment('center')
                 .setBackground('#ffffff')
                 .setFontWeight('bold')
                 .setFontSize(14);
      sundaySheet.getRange(summaryRow + i + 1, 3).setValue(overallStats[i][2])
                 .setFontSize(10)
                 .setFontColor('#666666')
                 .setBackground('#ffffff');
    }
    
    sundaySheet.getRange(summaryRow, 1, overallStats.length + 1, 3).setBorder(
      true, true, true, true, true, true,
      '#cccccc',
      SpreadsheetApp.BorderStyle.SOLID
    );
    
  } catch (error) {
    Logger.log('Error updating Sunday breakdown: ' + error.toString());
  }
}

// ==================== AUTOMATED TRIGGERS ====================

/**
 * Sets up automated triggers for dashboard updates
 * Run this function ONCE after setupAllSheets()
 */
function setupAutomatedTriggers() {
  try {
    // Delete existing triggers first
    deleteExistingTriggers();
    
    // Create time-based trigger - runs every 5 hours
    ScriptApp.newTrigger('updateAllSummaries')
        .timeBased()
        .everyHours(5)
        .create();
    
    Logger.log('✓ Automated trigger created: Updates every 5 hours');
    
    Browser.msgBox(
      'Triggers Setup Complete!',
      'Automated updates configured:\\n\\n' +
      '✓ Dashboard updates every 5 hours\\n\\n' +
      'You can also manually update by running updateAllSummaries()\\n\\n' +
      'To view or delete triggers, go to:\\n' +
      'Extensions > Apps Script > Triggers',
      Browser.Buttons.OK
    );
    
    return { success: true };
    
  } catch (error) {
    Logger.log('Error setting up triggers: ' + error.toString());
    Browser.msgBox(
      'Trigger Setup Failed',
      'Error: ' + error.toString(),
      Browser.Buttons.OK
    );
    return { success: false, error: error.toString() };
  }
}

/**
 * Deletes all existing triggers for this project
 */
function deleteExistingTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log('✓ Deleted ' + triggers.length + ' existing trigger(s)');
}

/**
 * View all current triggers
 */
function viewTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  
  Logger.log('╔═══════════════════════════════════════╗');
  Logger.log('║        CURRENT TRIGGERS                ║');
  Logger.log('╚═══════════════════════════════════════╝');
  
  if (triggers.length === 0) {
    Logger.log('No triggers found.');
    Logger.log('Run setupAutomatedTriggers() to create them.');
  } else {
    for (var i = 0; i < triggers.length; i++) {
      var trigger = triggers[i];
      Logger.log('Trigger ' + (i + 1) + ':');
      Logger.log('  Function: ' + trigger.getHandlerFunction());
      Logger.log('  Trigger Source: ' + trigger.getTriggerSource());
      Logger.log('  Event Type: ' + trigger.getEventType());
      Logger.log('');
    }
  }
  
  Browser.msgBox(
    'Current Triggers',
    'Found ' + triggers.length + ' trigger(s).\\n\\n' +
    'Check the Execution log for details.\\n\\n' +
    'To manage triggers, go to:\\n' +
    'Extensions > Apps Script > Triggers',
    Browser.Buttons.OK
  );
}

// ==================== HELPER FUNCTIONS ====================

function getMonthName(monthIndex) {
  var months = ['January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'];
  return months[monthIndex];
}

function countByType(data, type) {
  var count = 0;
  for (var i = 0; i < data.length; i++) {
    if (data[i][4] === type) { // Column E (index 4) is Type
      count++;
    }
  }
  return count;
}

function getUniqueEmails(data) {
  var emails = [];
  for (var i = 0; i < data.length; i++) {
    var email = data[i][3]; // Column D (index 3) is Email
    if (email && emails.indexOf(email) === -1) {
      emails.push(email);
    }
  }
  return emails;
}

// ==================== MANUAL UPDATE & MAINTENANCE ====================

/**
 * Manual refresh of all dashboard tabs
 * Can be run anytime to update all dashboards
 */
function updateAllSummaries() {
  Logger.log('╔═══════════════════════════════════════╗');
  Logger.log('║    UPDATING ALL DASHBOARD TABS         ║');
  Logger.log('╚═══════════════════════════════════════╝');
  
  try {
    Logger.log('Updating Current Month...');
    updateMonthlySummary();
    Logger.log('✓ Current Month updated');
    
    Logger.log('Updating Yearly Summary...');
    updateYearlySummary();
    Logger.log('✓ Yearly Summary updated');
    
    Logger.log('Updating Sunday Breakdown...');
    updateSundayBreakdown();
    Logger.log('✓ Sunday Breakdown updated');
    
    Logger.log('');
    Logger.log('✓ All dashboards updated successfully!');
    Logger.log('Updated at: ' + new Date().toLocaleString());
    
    // Show confirmation in UI if called manually
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      '✓ Dashboard Update Complete',
      'All dashboard tabs have been refreshed:\\n\\n' +
      '• Current Month\\n' +
      '• Yearly Summary\\n' +
      '• Sunday Breakdown\\n\\n' +
      'Updated at: ' + new Date().toLocaleString(),
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('✗ Error updating dashboards: ' + error.toString());
    
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      'Update Failed',
      'Error: ' + error.toString(),
      ui.ButtonSet.OK
    );
  }
}

/**
 * Creates a custom menu in Google Sheets for easy access
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Sunday Service')
      .addItem('🔄 Refresh All Dashboards', 'updateAllSummaries')
      .addSeparator()
      .addItem('⚙️ Setup All Sheets', 'setupAllSheets')
      .addItem('⏰ Setup Auto-Updates (5 hours)', 'setupAutomatedTriggers')
      .addSeparator()
      .addItem('👁️ View Configuration', 'viewConfiguration')
      .addItem('🔔 View Triggers', 'viewTriggers')
      .addSeparator()
      .addItem('🧪 Test Registration', 'testRegistration')
      .addToUi();
}

/**
 * Test function
 */
function testRegistration() {
  Logger.log('Running test registration...');
  
  var testData = {
    registrants: [
      {
        firstName: 'Test',
        lastName: 'User',
        email: Session.getActiveUser().getEmail(),
        type: 'Member'
      }
    ],
    sessionInfo: 'Sunday Service - Test',
    sundayDate: getNextSunday().formatted
  };
  
  var result = saveRegistration(testData);
  Logger.log('Test result: ' + JSON.stringify(result));
  
  if (result.success) {
    Browser.msgBox(
      'Test Successful!',
      'Registration test completed.\\n\\n' +
      'Check your email: ' + Session.getActiveUser().getEmail() + '\\n\\n' +
      'Check your Google Sheet for the test data.',
      Browser.Buttons.OK
    );
  } else {
    Browser.msgBox(
      'Test Failed',
      'Error: ' + result.error,
      Browser.Buttons.OK
    );
  }
}

/**
 * View configuration
 */
function viewConfiguration() {
  var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  
  Logger.log('╔═══════════════════════════════════════╗');
  Logger.log('║        CURRENT CONFIGURATION           ║');
  Logger.log('╚═══════════════════════════════════════╝');
  Logger.log('Sheet ID: ' + SHEET_ID);
  Logger.log('Sheet URL: ' + spreadsheet.getUrl());
  Logger.log('Main Tab: ' + SHEET_NAME);
  Logger.log('Monthly Tab: ' + MONTHLY_SHEET_NAME);
  Logger.log('Yearly Tab: ' + YEARLY_SHEET_NAME);
  Logger.log('Sunday Tab: ' + SUNDAY_SHEET_NAME);
  Logger.log('Live Stream: ' + LIVE_STREAM_LINK);
  Logger.log('Service End Hour: ' + SERVICE_END_HOUR + ':00');
  Logger.log('');
  
  var sundayInfo = getNextSunday();
  Logger.log('Current Sunday Logic:');
  Logger.log('  Next/Current Sunday: ' + sundayInfo.formatted);
  Logger.log('  Is Today: ' + sundayInfo.isToday);
  Logger.log('  Date Only: ' + sundayInfo.dateOnly);
  
  Browser.msgBox(
    'Configuration',
    'Configuration details logged to Execution log.\\n\\n' +
    'Sheet URL: ' + spreadsheet.getUrl() + '\\n\\n' +
    'Next Sunday: ' + sundayInfo.formatted,
    Browser.Buttons.OK
  );
}
