/**
 * Sunday Service Registration System - FIXED VERSION
 * All tabs auto-update, email column properly created, automatic email sending
 * 
 * SETUP: Run setupAllSheets() once, then deploy as web app
 */

// ==================== CONFIGURATION ====================
const SHEET_ID = '1Z7JrY-7US55iUnNmyXB3m5hKFGoWpR0bGMW1XXRgWnQ';
const SHEET_NAME = 'Registrations';
const MONTHLY_SHEET_NAME = 'Current Month';
const YEARLY_SHEET_NAME = 'Yearly Summary';
const LIVE_STREAM_LINK = 'https://www.youtube.com/live/xU2alDtuXkw';

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
 * Returns the live stream link to the web app
 */
function getLiveStreamLink() {
  return LIVE_STREAM_LINK;
}

/**
 * Calculates the next Sunday date
 */
function getNextSunday() {
  var today = new Date();
  var dayOfWeek = today.getDay();
  var daysUntilSunday = (7 - dayOfWeek) % 7;
  
  if (daysUntilSunday === 0) {
    daysUntilSunday = 0;
  }
  
  var nextSunday = new Date(today);
  nextSunday.setDate(today.getDate() + daysUntilSunday);
  
  var options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  return {
    formatted: nextSunday.toLocaleDateString('en-US', options),
    date: nextSunday
  };
}

// ==================== REGISTRATION FUNCTIONS ====================

/**
 * Saves registration data to Google Sheet and sends confirmation email
 * AUTOMATICALLY UPDATES ALL TABS
 */
function saveRegistration(formData) {
  try {
    Logger.log('========================================');
    Logger.log('STARTING REGISTRATION PROCESS');
    Logger.log('========================================');
    Logger.log('Form data received: ' + JSON.stringify(formData));
    
    var sheet = getOrCreateSheet();
    var timestamp = new Date();
    var emails = [];
    
    Logger.log('Processing ' + formData.registrants.length + ' registrant(s)');
    
    // Add each person in the registration
    for (var i = 0; i < formData.registrants.length; i++) {
      var person = formData.registrants[i];
      Logger.log('Adding person ' + (i + 1) + ': ' + person.firstName + ' ' + person.lastName + ' (' + person.email + ')');
      
      // IMPORTANT: Column order must match headers
      // Columns: Timestamp, First Name, Last Name, Email, Type, Session, Sunday Date
      sheet.appendRow([
        timestamp,              // Column A: Timestamp
        person.firstName,       // Column B: First Name
        person.lastName,        // Column C: Last Name
        person.email,           // Column D: Email
        person.type,            // Column E: Type
        formData.sessionInfo || 'Sunday Service',  // Column F: Session
        formData.sundayDate || ''  // Column G: Sunday Date
      ]);
      
      // Collect unique emails for sending confirmations
      if (person.email && emails.indexOf(person.email) === -1) {
        emails.push(person.email);
      }
    }
    
    Logger.log('✓ All registrants added to sheet');
    
    // Sort by timestamp (newest first)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var range = sheet.getRange(2, 1, lastRow - 1, 7);
      range.sort({column: 1, ascending: false});
      Logger.log('✓ Data sorted by timestamp');
    }
    
    // AUTOMATICALLY UPDATE ALL TABS
    Logger.log('Updating Current Month tab...');
    updateMonthlySummary();
    Logger.log('✓ Current Month tab updated');
    
    Logger.log('Updating Yearly Summary tab...');
    updateYearlySummary();
    Logger.log('✓ Yearly Summary tab updated');
    
    // Send confirmation emails
    Logger.log('Sending confirmation emails to ' + emails.length + ' recipient(s)...');
    var emailsSent = 0;
    var emailErrors = [];
    
    for (var i = 0; i < emails.length; i++) {
      var email = emails[i];
      Logger.log('Attempting to send email to: ' + email);
      
      try {
        if (sendConfirmationEmail(email, formData.registrants, formData.sundayDate)) {
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
    Logger.log('Total registrants: ' + formData.registrants.length);
    Logger.log('Emails sent: ' + emailsSent + '/' + emails.length);
    if (emailErrors.length > 0) {
      Logger.log('Email errors for: ' + emailErrors.join(', '));
    }
    Logger.log('========================================');
    
    return {
      success: true,
      message: 'Registration successful! ' + (emailsSent > 0 ? 'Check your email for confirmation.' : ''),
      count: formData.registrants.length,
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
function sendConfirmationEmail(email, registrants, sundayDate) {
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
    
    var subject = '✅ Sunday Service Registration Confirmed - ' + sundayDate;
    
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
      '<p style="margin: 10px 0;"><strong>Date:</strong> ' + sundayDate + '</p>' +
      '<p style="margin: 10px 0;"><strong>Registered:</strong> ' + namesString + '</p>' +
      '<p style="margin: 10px 0;"><strong>Type:</strong> ' + typesString + '</p>' +
      '</div>' +
      '<div style="text-align: center; margin: 30px 0;">' +
      '<a href="' + LIVE_STREAM_LINK + '" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px 30px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">📺 Join Live Stream</a>' +
      '</div>' +
      '<p style="font-size: 14px; color: #666;">We look forward to worshiping with you this Sunday!</p>' +
      '<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">' +
      '<p style="font-size: 12px; color: #999; text-align: center;">This is an automated confirmation email for your Sunday Service registration.</p>' +
      '</div>' +
      '</div>';
    
    var plainBody = 
      'Sunday Service Registration Confirmed\n\n' +
      'Dear ' + people[0].firstName + ',\n\n' +
      'Thank you for registering for our Sunday Service!\n\n' +
      'Service Details:\n' +
      'Date: ' + sundayDate + '\n' +
      'Registered: ' + namesString + '\n' +
      'Type: ' + typesString + '\n\n' +
      'Live Stream Link: ' + LIVE_STREAM_LINK + '\n\n' +
      'We look forward to worshiping with you this Sunday!\n\n' +
      'This is an automated confirmation email.';
    
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
 * MAIN SETUP FUNCTION - RUN THIS FIRST
 * Creates all 3 tabs with proper formatting including EMAIL COLUMN
 */
function setupAllSheets() {
  Logger.log('╔════════════════════════════════════════╗');
  Logger.log('║  SUNDAY SERVICE REGISTRATION SETUP     ║');
  Logger.log('╚════════════════════════════════════════╝');
  
  try {
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
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
        'The Registrations tab already exists. Do you want to recreate it? This will DELETE all existing data!',
        Browser.Buttons.YES_NO
      );
      
      if (response === 'yes') {
        spreadsheet.deleteSheet(mainSheet);
        Logger.log('✓ Old Registrations tab deleted');
        mainSheet = spreadsheet.insertSheet(SHEET_NAME);
        formatMainSheet(mainSheet);
        Logger.log('✓ New Registrations tab created with EMAIL column');
      } else {
        Logger.log('ℹ Keeping existing Registrations tab');
      }
    } else {
      mainSheet = spreadsheet.insertSheet(SHEET_NAME);
      formatMainSheet(mainSheet);
      Logger.log('✓ Registrations tab created with EMAIL column');
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
    updateMonthlySummary();
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
    updateYearlySummary();
    Logger.log('✓ Yearly Summary tab populated');
    
    Logger.log('');
    Logger.log('╔════════════════════════════════════════╗');
    Logger.log('║           SETUP COMPLETE! ✓            ║');
    Logger.log('╚════════════════════════════════════════╝');
    Logger.log('');
    Logger.log('Your Google Sheet is ready at:');
    Logger.log(spreadsheet.getUrl());
    Logger.log('');
    Logger.log('All 3 tabs created:');
    Logger.log('  1. Registrations (with Email in column D)');
    Logger.log('  2. Current Month (auto-updating)');
    Logger.log('  3. Yearly Summary (auto-updating)');
    Logger.log('');
    Logger.log('Next steps:');
    Logger.log('  1. Deploy your web app');
    Logger.log('  2. Test with testRegistration() function');
    Logger.log('  3. Embed form on your website');
    
    Browser.msgBox(
      'Setup Complete!',
      'All tabs have been created successfully.\\n\\n' +
      '✓ Registrations tab (with Email column)\\n' +
      '✓ Current Month dashboard\\n' +
      '✓ Yearly Summary dashboard\\n\\n' +
      'Check the Execution log for details.',
      Browser.Buttons.OK
    );
    
    return {
      success: true,
      message: 'Setup completed successfully',
      url: spreadsheet.getUrl()
    };
    
  } catch (error) {
    Logger.log('');
    Logger.log('╔════════════════════════════════════════╗');
    Logger.log('║         SETUP FAILED! ✗                ║');
    Logger.log('╚════════════════════════════════════════╝');
    Logger.log('');
    Logger.log('Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    
    Browser.msgBox(
      'Setup Failed',
      'There was an error during setup: ' + error.toString(),
      Browser.Buttons.OK
    );
    
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * Formats the main Registrations sheet with PROPER EMAIL COLUMN
 */
function formatMainSheet(sheet) {
  Logger.log('  → Adding headers with Email column...');
  
  // CRITICAL: These columns must match the saveRegistration function
  var headers = [
    'Timestamp',      // Column A
    'First Name',     // Column B
    'Last Name',      // Column C
    'Email',          // Column D ← EMAIL COLUMN
    'Type',           // Column E
    'Session',        // Column F
    'Sunday Date'     // Column G
  ];
  
  sheet.appendRow(headers);
  Logger.log('  → Headers added: ' + headers.join(', '));
  
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
  sheet.setColumnWidth(2, 120);  // First Name
  sheet.setColumnWidth(3, 120);  // Last Name
  sheet.setColumnWidth(4, 250);  // Email ← WIDER FOR EMAILS
  sheet.setColumnWidth(5, 100);  // Type
  sheet.setColumnWidth(6, 180);  // Session
  sheet.setColumnWidth(7, 180);  // Sunday Date
  
  Logger.log('  → Column widths set (Email column is 250px)');
  
  // Add border to headers
  headerRange.setBorder(
    true, true, true, true, true, true,
    '#ffffff',
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  // Data validation for Type column (column E)
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Member', 'Guest'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 5, 1000, 1).setDataValidation(typeRule);
  
  Logger.log('  → Data validation added for Type column');
  
  // Apply row banding
  var dataRange = sheet.getRange(2, 1, 1000, 7);
  var banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  banding.setHeaderRowColor('#667eea')
         .setFirstRowColor('#f3f4ff')
         .setSecondRowColor('#ffffff');
  
  Logger.log('  → Alternating row colors applied');
  
  // Center align Type column
  sheet.getRange(2, 5, 1000, 1).setHorizontalAlignment('center');
  
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
  var summaryStartCol = 9; // Column I
  
  sheet.getRange(1, summaryStartCol).setValue('REGISTRATION SUMMARY');
  sheet.getRange(1, summaryStartCol, 1, 2).merge()
       .setFontWeight('bold')
       .setFontSize(12)
       .setBackground('#764ba2')
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  // Summary formulas - Column D is Email
  var summaryData = [
    ['Total Registrations', '=COUNTA(B2:B)'],
    ['Total Members', '=COUNTIF(E2:E,"Member")'],
    ['Total Guests', '=COUNTIF(E2:E,"Guest")'],
    ['Unique Emails', '=COUNTA(UNIQUE(FILTER(D2:D,D2:D<>"")))'],
    ['Last Registration', '=IF(COUNTA(A2:A)>0,TEXT(MAX(A2:A),"yyyy-mm-dd hh:mm"),"No data")']
  ];
  
  for (var i = 0; i < summaryData.length; i++) {
    var rowNum = 3 + i;
    sheet.getRange(rowNum, summaryStartCol).setValue(summaryData[i][0])
         .setFontWeight('bold')
         .setBackground('#f3f4ff');
    sheet.getRange(rowNum, summaryStartCol + 1).setFormula(summaryData[i][1])
         .setHorizontalAlignment('center')
         .setBackground('#ffffff');
  }
  
  sheet.setColumnWidth(summaryStartCol, 180);
  sheet.setColumnWidth(summaryStartCol + 1, 150);
  
  sheet.getRange(1, summaryStartCol, 7, 2).setBorder(
    true, true, true, true, true, true,
    '#cccccc',
    SpreadsheetApp.BorderStyle.SOLID
  );
  
  Logger.log('  → Summary section added');
}

/**
 * Gets or creates the main sheet
 */
function getOrCreateSheet() {
  var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    Logger.log('Main sheet not found, creating it...');
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    formatMainSheet(sheet);
  }
  
  // Verify email column exists
  var headers = sheet.getRange(1, 1, 1, 7).getValues()[0];
  if (headers[3] !== 'Email') {
    Logger.log('WARNING: Email column not found in position D!');
    Logger.log('Current headers: ' + headers.join(', '));
    Logger.log('Please run setupAllSheets() to fix the sheet structure');
  }
  
  return sheet;
}

// ==================== AUTO-UPDATE FUNCTIONS ====================

/**
 * Updates Current Month tab - CALLED AUTOMATICALLY AFTER EACH REGISTRATION
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
    
    // Headers WITH EMAIL
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
    monthlySheet.setColumnWidth(4, 250); // Email
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
 * Updates Yearly Summary tab - CALLED AUTOMATICALLY AFTER EACH REGISTRATION
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

// ==================== TEST & MAINTENANCE ====================

/**
 * Test function - sends test data and email
 */
function testRegistration() {
  Logger.log('Running test registration...');
  
  var testData = {
    registrants: [
      {
        firstName: 'Test',
        lastName: 'User',
        email: Session.getActiveUser().getEmail(), // Your email
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
 * Manual refresh of all tabs
 */
function updateAllSummaries() {
  Logger.log('Manually updating all summaries...');
  updateMonthlySummary();
  updateYearlySummary();
  Logger.log('All summaries updated!');
  
  Browser.msgBox(
    'Update Complete',
    'All dashboard tabs have been refreshed.',
    Browser.Buttons.OK
  );
}

/**
 * View configuration
 */
function viewConfiguration() {
  var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  
  Logger.log('╔════════════════════════════════════════╗');
  Logger.log('║        CURRENT CONFIGURATION           ║');
  Logger.log('╚════════════════════════════════════════╝');
  Logger.log('Sheet ID: ' + SHEET_ID);
  Logger.log('Sheet URL: ' + spreadsheet.getUrl());
  Logger.log('Main Tab: ' + SHEET_NAME);
  Logger.log('Monthly Tab: ' + MONTHLY_SHEET_NAME);
  Logger.log('Yearly Tab: ' + YEARLY_SHEET_NAME);
  Logger.log('Live Stream: ' + LIVE_STREAM_LINK);
  Logger.log('');
  
  var mainSheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (mainSheet) {
    var headers = mainSheet.getRange(1, 1, 1, 7).getValues()[0];
    Logger.log('Column Headers:');
    for (var i = 0; i < headers.length; i++) {
      Logger.log('  Column ' + String.fromCharCode(65 + i) + ': ' + headers[i]);
    }
    Logger.log('');
    Logger.log('✓ Email column: ' + (headers[3] === 'Email' ? 'Correctly positioned in column D' : 'MISSING OR WRONG POSITION!'));
  }
  
  Browser.msgBox(
    'Configuration',
    'Configuration details logged to Execution log.\\n\\n' +
    'Sheet URL: ' + spreadsheet.getUrl(),
    Browser.Buttons.OK
  );
}
