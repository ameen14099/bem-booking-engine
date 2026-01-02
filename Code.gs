// ============================================
// BEM - BOOKING ENGINE MACHINE
// Code.gs - Main Backend
// ============================================

const SHEET_ID = '12EvZajKN_2a7ATo3M2p9ZS5agHos3pP8ulp42UWJ7PA';

const SHEETS = {
  SETTINGS: 'Settings',
  LEADS: 'Leads',
  RESEARCH: 'Research_Queue',
  CONTRACTS: 'Contracts',
  EOD: 'EOD_Metrics',
  USERS: 'Users'
};

// Status values that trigger Location changes
const PARKING_LOT_STATUSES = ['Dead', 'Lost', 'Slow Cooker', 'Canceled'];
const BOOKED_STATUSES = ['Booked'];

// ============================================
// WEB APP ENTRY POINT
// ============================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('BEM - Booking Engine Machine')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// SETTINGS - DYNAMIC DROPDOWNS
// ============================================

function getDropdownOptions() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SETTINGS);
  const data = sheet.getDataRange().getValues();
  
  const options = {};
  const headers = data[0];
  
  headers.forEach((header, colIndex) => {
    options[header] = [];
    for (let row = 1; row < data.length; row++) {
      const value = data[row][colIndex];
      if (value && value.toString().trim() !== '') {
        options[header].push(value.toString().trim());
      }
    }
  });
  
  return options;
}

// ============================================
// USER MANAGEMENT
// ============================================

function getUsers() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEETS.USERS);
    
    if (!sheet) {
      Logger.log('Users sheet not found');
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      Logger.log('No user data rows');
      return [];
    }
    
    var headers = data[0];
    var users = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Skip empty rows
      if (!row[1]) continue; // Column B is Name
      
      var user = {};
      for (var j = 0; j < headers.length; j++) {
        var value = row[j];
        // Convert Date objects to strings
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        user[headers[j]] = value;
      }
      
      // Only include active users
      if (user.Name && (user.Active === 'Yes' || user.Active === true)) {
        users.push(user);
      }
    }
    
    Logger.log('Returning ' + users.length + ' users');
    return users;
    
  } catch (e) {
    Logger.log('Error in getUsers: ' + e.toString());
    return [];
  }
}

function getUserByName(name) {
  const users = getUsers();
  return users.find(u => u.Name.toLowerCase() === name.toLowerCase()) || null;
}

// ============================================
// DASHBOARD STATS
// ============================================

function getDashboardStats(currentUser) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const leadsSheet = ss.getSheetByName(SHEETS.LEADS);
  const data = leadsSheet.getDataRange().getValues();
  
  if (data.length < 2) {
    return {
      totalActive: 0,
      hotLeads: 0,
      warmLeads: 0,
      bookedCount: 0,
      parkingLotCount: 0,
      storageCount: 0,
      pipelineValue: '$0',
      contractsReview: 0,
      todayFollowUps: 0,
      myLeadsCount: 0,
      dhillCount: 0
    };
  }
  
  const headers = data[0];
  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i);
  
  let totalActive = 0;
  let hotLeads = 0;
  let warmLeads = 0;
  let bookedCount = 0;
  let parkingLotCount = 0;
  let storageCount = 0;
  let pipelineValue = 0;
  let contractsReview = 0;
  let todayFollowUps = 0;
  let myLeadsCount = 0;
  let dhillCount = 0;
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const location = row[colIndex['Location']];
    const status = row[colIndex['Lead_Status']];
    const quote = row[colIndex['Active_Quote']];
    const mockContract = row[colIndex['Create_Mock_Contract']];
    const followupDate = row[colIndex['Date_Next_Followup']];
    const owner = row[colIndex['Lead_Owner']];
    const dhillReview = row[colIndex['Dhill_Call_List']];
    
    // Count by location
    if (location === 'Active') {
      totalActive++;
      
      // Count by status
      if (status === 'Hot' || status === 'Super Hot' || status === 'Contract Lead') {
        hotLeads++;
      } else if (status === 'Warm') {
        warmLeads++;
      }
      
      // Pipeline value (only active leads)
      if (quote) {
        const numericQuote = parseFloat(String(quote).replace(/[$,]/g, ''));
        if (!isNaN(numericQuote)) {
          pipelineValue += numericQuote;
        }
      }
      
      // Contracts to review
      if (mockContract === 'Yes') {
        contractsReview++;
      }
      
      // Today's follow ups
      if (followupDate instanceof Date) {
        const followup = new Date(followupDate);
        followup.setHours(0, 0, 0, 0);
        if (followup.getTime() === today.getTime()) {
          todayFollowUps++;
        }
      }
      
      // My leads count
      if (currentUser && owner === currentUser) {
        myLeadsCount++;
      }
      
      // Dhill review count
      if (dhillReview === 'Yes') {
        dhillCount++;
      }
    } else if (location === 'Booked') {
      bookedCount++;
    } else if (location === 'Parking Lot') {
      parkingLotCount++;
    } else if (location === 'Storage') {
      storageCount++;
    }
  }
  
  return {
    totalActive,
    hotLeads,
    warmLeads,
    bookedCount,
    parkingLotCount,
    storageCount,
    pipelineValue: '$' + pipelineValue.toLocaleString(),
    contractsReview,
    todayFollowUps,
    myLeadsCount,
    dhillCount
  };
}

// ============================================
// GET LEADS (with filtering)
// ============================================

function getLeads(view, currentUser) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.LEADS);
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) return [];
  
  const headers = data[0];
  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i);
  
  const leads = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Skip empty rows
    if (!row[colIndex['Contact_Person']] && !row[colIndex['Email']] && !row[colIndex['Company_Name']]) {
      continue;
    }
    
    const location = row[colIndex['Location']] || 'Active';
    const status = row[colIndex['Lead_Status']];
    const owner = row[colIndex['Lead_Owner']];
    const dhillReview = row[colIndex['Dhill_Call_List']];
    const mockContract = row[colIndex['Create_Mock_Contract']];
    const followupDate = row[colIndex['Date_Next_Followup']];
    
    // Apply view filters
    let include = false;
    
    switch (view) {
      case 'all':
        include = (location === 'Active');
        break;
      case 'my':
        include = (location === 'Active' && owner === currentUser);
        break;
      case 'hot':
        include = (location === 'Active' && ['Hot', 'Super Hot', 'Contract Lead'].includes(status));
        break;
      case 'contracts':
        include = (location === 'Active' && mockContract === 'Yes');
        break;
      case 'followups':
        if (location === 'Active' && followupDate instanceof Date) {
          const followup = new Date(followupDate);
          followup.setHours(0, 0, 0, 0);
          include = (followup.getTime() === today.getTime());
        }
        break;
      case 'booked':
        include = (location === 'Booked');
        break;
      case 'parking':
        include = (location === 'Parking Lot');
        break;
      case 'dhill':
        include = (location === 'Active' && dhillReview === 'Yes');
        break;
      case 'storage':
        include = (location === 'Storage');
        break;
      default:
        include = (location === 'Active');
    }
    
    if (include) {
      const lead = { rowIndex: i + 1 };
      headers.forEach((header, index) => {
        let value = row[index];
        if (value instanceof Date) {
          value = formatDate(value);
        }
        lead[header] = value;
      });
      leads.push(lead);
    }
  }
  
  // Sort by Priority_Score descending, then by Date_Last_Touch
  leads.sort((a, b) => {
    const scoreA = parseFloat(a.Priority_Score) || 0;
    const scoreB = parseFloat(b.Priority_Score) || 0;
    if (scoreB !== scoreA) return scoreB - scoreA;
    
    const dateA = a.Date_Last_Touch ? new Date(a.Date_Last_Touch) : new Date(0);
    const dateB = b.Date_Last_Touch ? new Date(b.Date_Last_Touch) : new Date(0);
    return dateB - dateA;
  });
  
  return leads;
}

// ============================================
// GET SINGLE LEAD
// ============================================

function getLeadByRow(rowIndex) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.LEADS);
  const data = sheet.getDataRange().getValues();
  
  if (rowIndex < 2 || rowIndex > data.length) return null;
  
  const headers = data[0];
  const row = data[rowIndex - 1];
  
  const lead = { rowIndex };
  headers.forEach((header, index) => {
    let value = row[index];
    if (value instanceof Date) {
      value = formatDateForInput(value);
    }
    lead[header] = value;
  });
  
  return lead;
}

// ============================================
// CREATE NEW LEAD
// ============================================

function createLead(leadData, currentUser) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.LEADS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Generate Record_ID
  const lastRow = sheet.getLastRow();
  const newId = 'BEM-' + String(lastRow).padStart(5, '0');
  
  // Prepare row data
  const rowData = headers.map(header => {
    switch (header) {
      case 'Record_ID':
        return newId;
      case 'Location':
        return leadData.Location || 'Active';
      case 'Date_Last_Touch':
        return new Date();
      case 'Created_Date':
        return new Date();
      case 'Created_By':
        return currentUser;
      case 'Last_Modified_By':
        return currentUser;
      case 'Priority_Score':
        return calculatePriorityScore(leadData);
      default:
        return leadData[header] || '';
    }
  });
  
  sheet.appendRow(rowData);
  
  return { success: true, recordId: newId, rowIndex: lastRow + 1 };
}

// ============================================
// UPDATE LEAD
// ============================================

function updateLead(rowIndex, updates, currentUser) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.LEADS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i + 1);
  
  // Get current data for comparison
  const currentData = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const oldStatus = currentData[colIndex['Lead_Status'] - 1];
  const newStatus = updates.Lead_Status || oldStatus;
  
  // Update each field
  Object.keys(updates).forEach(field => {
    if (colIndex[field]) {
      let value = updates[field];
      
      // Convert date strings to Date objects
      if (['Date_of_Event', 'Date_Next_Followup'].includes(field) && value) {
        value = new Date(value);
      }
      
      sheet.getRange(rowIndex, colIndex[field]).setValue(value);
    }
  });
  
  // Auto-update timestamp and modifier
  sheet.getRange(rowIndex, colIndex['Date_Last_Touch']).setValue(new Date());
  sheet.getRange(rowIndex, colIndex['Last_Modified_By']).setValue(currentUser);
  
  // Recalculate priority score
  const updatedData = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const leadForScore = {};
  headers.forEach((h, i) => leadForScore[h] = updatedData[i]);
  const newScore = calculatePriorityScore(leadForScore);
  sheet.getRange(rowIndex, colIndex['Priority_Score']).setValue(newScore);
  
  // Handle status-based Location changes
  let moved = false;
  let newLocation = null;
  
  if (newStatus !== oldStatus) {
    if (PARKING_LOT_STATUSES.includes(newStatus)) {
      sheet.getRange(rowIndex, colIndex['Location']).setValue('Parking Lot');
      moved = true;
      newLocation = 'Parking Lot';
    } else if (BOOKED_STATUSES.includes(newStatus)) {
      sheet.getRange(rowIndex, colIndex['Location']).setValue('Booked');
      moved = true;
      newLocation = 'Booked';
    } else if (PARKING_LOT_STATUSES.includes(oldStatus) || BOOKED_STATUSES.includes(oldStatus)) {
      // Moving back to Active from Parking Lot or Booked
      sheet.getRange(rowIndex, colIndex['Location']).setValue('Active');
      moved = true;
      newLocation = 'Active';
    }
    
    // Super Hot alert
    if (newStatus === 'Super Hot') {
      sendSuperHotAlert(leadForScore, currentUser);
    }
    
    // Track status upgrade for EOD metrics
    if (['Hot', 'Warm', 'Contract Lead'].includes(newStatus) && 
        !['Hot', 'Warm', 'Contract Lead', 'Super Hot'].includes(oldStatus)) {
      incrementMetric(currentUser, 'Leads_Upgraded', 1);
    }
  }
  
  return { 
    success: true, 
    moved, 
    newLocation,
    message: moved ? 'Lead moved to ' + newLocation : 'Lead updated successfully'
  };
}

// ============================================
// PRIORITY SCORE CALCULATION
// ============================================

function calculatePriorityScore(lead) {
  let score = 0;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // Event date urgency (closer = higher, max 200 points)
  if (lead.Date_of_Event) {
    const eventDate = new Date(lead.Date_of_Event);
    eventDate.setHours(0, 0, 0, 0);
    const daysUntil = Math.floor((eventDate - today) / (1000 * 60 * 60 * 24));
    if (daysUntil >= 0 && daysUntil <= 90) {
      score += Math.max(0, 200 - (daysUntil * 2));
    }
  }
  
  // Staleness (older last touch = higher priority to follow up, max 50 points)
  if (lead.Date_Last_Touch) {
    const lastTouch = new Date(lead.Date_Last_Touch);
    lastTouch.setHours(0, 0, 0, 0);
    const daysSince = Math.floor((today - lastTouch) / (1000 * 60 * 60 * 24));
    score += Math.min(daysSince * 3, 50);
  }
  
  // Status weight
  const statusWeights = {
    'Super Hot': 100,
    'Contract Lead': 90,
    'Hot': 80,
    'Warm': 50,
    'Cold': 10
  };
  score += statusWeights[lead.Lead_Status] || 0;
  
  // Quote value boost (max 30 points)
  if (lead.Active_Quote) {
    const quote = parseFloat(String(lead.Active_Quote).replace(/[$,]/g, ''));
    if (!isNaN(quote)) {
      score += Math.min(quote / 1000, 30);
    }
  }
  
  // Follow-up due today = top priority
  if (lead.Date_Next_Followup) {
    const followup = new Date(lead.Date_Next_Followup);
    followup.setHours(0, 0, 0, 0);
    if (followup.getTime() === today.getTime()) {
      score += 200;
    }
  }
  
  return Math.round(score);
}

// ============================================
// SUPER HOT ALERT - WITH HTML TEMPLATE
// ============================================
// Replace your existing sendSuperHotAlert function with this one

function sendSuperHotAlert(lead, markedBy) {
  try {
    // ⚠️ TESTING: Using your email. Change to dhillmagic@gmail.com for production
    const recipient = 'elamin.elneel@gmail.com';
    // const recipient = 'dhillmagic@gmail.com'; // PRODUCTION - Uncomment when ready
    
    const contactName = lead.Contact_Person || lead.Company_Name || 'New Lead';
    const subject = '🔥 SUPER HOT LEAD: ' + contactName;
    
    // Format the quote nicely
    let quoteDisplay = 'Not quoted yet';
    if (lead.Active_Quote) {
      const quoteNum = parseFloat(String(lead.Active_Quote).replace(/[$,]/g, ''));
      if (!isNaN(quoteNum) && quoteNum > 0) {
        quoteDisplay = '$' + quoteNum.toLocaleString('en-US', {minimumFractionDigits: 0, maximumFractionDigits: 0});
      }
    }
    
    // Format event date
    let eventDateDisplay = 'TBD';
    if (lead.Date_of_Event) {
      try {
        const d = new Date(lead.Date_of_Event);
        if (!isNaN(d.getTime())) {
          const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
          const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
          eventDateDisplay = days[d.getDay()] + ', ' + months[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
        }
      } catch (e) {
        eventDateDisplay = String(lead.Date_of_Event);
      }
    }
    
    // Current timestamp
    const now = new Date();
    const timeStamp = now.toLocaleString('en-US', { 
      weekday: 'short', 
      month: 'short', 
      day: 'numeric', 
      hour: 'numeric', 
      minute: '2-digit',
      hour12: true 
    });
    
    // Build HTML email
    const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif; background-color: #f4f4f5;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f4f4f5; padding: 40px 20px;">
    <tr>
      <td align="center">
        <table width="600" cellpadding="0" cellspacing="0" style="background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
          
          <!-- Header -->
          <tr>
            <td style="background: linear-gradient(135deg, #f97316 0%, #ea580c 100%); padding: 30px 40px; text-align: center;">
              <div style="font-size: 48px; margin-bottom: 10px;">🔥</div>
              <h1 style="margin: 0; color: #ffffff; font-size: 28px; font-weight: 700; letter-spacing: -0.5px;">
                SUPER HOT LEAD
              </h1>
              <p style="margin: 8px 0 0 0; color: rgba(255,255,255,0.9); font-size: 14px;">
                Requires immediate attention
              </p>
            </td>
          </tr>
          
          <!-- Contact Info Card -->
          <tr>
            <td style="padding: 30px 40px 20px 40px;">
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #fafafa; border-radius: 8px; border: 1px solid #e4e4e7;">
                <tr>
                  <td style="padding: 20px;">
                    <h2 style="margin: 0 0 15px 0; color: #18181b; font-size: 22px; font-weight: 600;">
                      ${escapeHtml(contactName)}
                    </h2>
                    
                    <!-- Contact Details Grid -->
                    <table width="100%" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="50%" style="padding: 8px 0; vertical-align: top;">
                          <div style="color: #71717a; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">Company</div>
                          <div style="color: #18181b; font-size: 15px; font-weight: 500;">${escapeHtml(lead.Company_Name || 'N/A')}</div>
                        </td>
                        <td width="50%" style="padding: 8px 0; vertical-align: top;">
                          <div style="color: #71717a; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">Event</div>
                          <div style="color: #18181b; font-size: 15px; font-weight: 500;">${escapeHtml(lead.Event_Name || 'N/A')}</div>
                        </td>
                      </tr>
                      <tr>
                        <td width="50%" style="padding: 8px 0; vertical-align: top;">
                          <div style="color: #71717a; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">Phone</div>
                          <div style="color: #18181b; font-size: 15px; font-weight: 500;">
                            <a href="tel:${escapeHtml(lead.Phone_Number || '')}" style="color: #6366f1; text-decoration: none;">${escapeHtml(lead.Phone_Number || 'N/A')}</a>
                          </div>
                        </td>
                        <td width="50%" style="padding: 8px 0; vertical-align: top;">
                          <div style="color: #71717a; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">Email</div>
                          <div style="color: #18181b; font-size: 15px; font-weight: 500;">
                            <a href="mailto:${escapeHtml(lead.Email || '')}" style="color: #6366f1; text-decoration: none;">${escapeHtml(lead.Email || 'N/A')}</a>
                          </div>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          
          <!-- Key Metrics -->
          <tr>
            <td style="padding: 0 40px 30px 40px;">
              <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <!-- Quote Amount -->
                  <td width="50%" style="padding-right: 10px;">
                    <div style="background-color: #dcfce7; border-radius: 8px; padding: 20px; text-align: center;">
                      <div style="color: #166534; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px;">Quote Amount</div>
                      <div style="color: #15803d; font-size: 28px; font-weight: 700;">${quoteDisplay}</div>
                    </div>
                  </td>
                  <!-- Event Date -->
                  <td width="50%" style="padding-left: 10px;">
                    <div style="background-color: #e0e7ff; border-radius: 8px; padding: 20px; text-align: center;">
                      <div style="color: #3730a3; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px;">Event Date</div>
                      <div style="color: #4338ca; font-size: 16px; font-weight: 600;">${eventDateDisplay}</div>
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          
          <!-- Notes Section (if any) -->
          ${lead.Additional_Notes ? `
          <tr>
            <td style="padding: 0 40px 30px 40px;">
              <div style="background-color: #fffbeb; border-left: 4px solid #f59e0b; padding: 15px 20px; border-radius: 0 8px 8px 0;">
                <div style="color: #92400e; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px;">Notes</div>
                <div style="color: #78350f; font-size: 14px; line-height: 1.5;">${escapeHtml(lead.Additional_Notes)}</div>
              </div>
            </td>
          </tr>
          ` : ''}
          
          <!-- Footer -->
          <tr>
            <td style="background-color: #f4f4f5; padding: 20px 40px; border-top: 1px solid #e4e4e7;">
              <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td style="color: #71717a; font-size: 13px;">
                    Marked by <strong style="color: #18181b;">${escapeHtml(markedBy)}</strong>
                  </td>
                  <td align="right" style="color: #a1a1aa; font-size: 12px;">
                    ${timeStamp}
                  </td>
                </tr>
              </table>
              <div style="margin-top: 15px; padding-top: 15px; border-top: 1px solid #e4e4e7; text-align: center;">
                <span style="color: #a1a1aa; font-size: 11px;">Sent from BEM - Booking Engine Machine</span>
              </div>
            </td>
          </tr>
          
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;

    // Plain text fallback
    const plainBody = 
      '🔥 SUPER HOT LEAD!\n\n' +
      'Contact: ' + contactName + '\n' +
      'Company: ' + (lead.Company_Name || 'N/A') + '\n' +
      'Event: ' + (lead.Event_Name || 'N/A') + '\n' +
      'Date: ' + eventDateDisplay + '\n' +
      'Phone: ' + (lead.Phone_Number || 'N/A') + '\n' +
      'Email: ' + (lead.Email || 'N/A') + '\n' +
      'Quote: ' + quoteDisplay + '\n\n' +
      'Marked by: ' + markedBy + '\n' +
      'Time: ' + timeStamp;
    
    // Send email with HTML
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody
    });
    
    Logger.log('Super Hot alert sent to ' + recipient + ' for lead: ' + contactName);
    return { success: true };
    
  } catch (e) {
    console.error('Failed to send Super Hot alert:', e);
    return { success: false, error: e.toString() };
  }
}

// Helper function to escape HTML
function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ============================================
// TEST FUNCTION - Run this to test the email
// ============================================
function testSuperHotAlert() {
  var testLead = {
    Contact_Person: 'John Smith',
    Company_Name: 'Acme Corporation',
    Event_Name: 'Annual Holiday Gala 2026',
    Date_of_Event: new Date('2026-03-15'),
    Phone_Number: '(555) 123-4567',
    Email: 'john.smith@acme.com',
    Active_Quote: '$4,500',
    Additional_Notes: 'CEO loved the demo video. Ready to book ASAP. Needs contract by Friday.'
  };
  
  var result = sendSuperHotAlert(testLead, 'Elamin');
  Logger.log('Test result: ' + JSON.stringify(result));
}

// Helper function to escape HTML
function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ============================================
// TEST FUNCTION - Run this to test the email
// ============================================
function testSuperHotAlert() {
  var testLead = {
    Record_ID: 'BEM-00142',
    Contact_Person: 'John Smith',
    Company_Name: 'Acme Corporation',
    Event_Name: 'Annual Holiday Gala 2026',
    Date_of_Event: new Date('2026-03-15'),
    Phone_Number: '(555) 123-4567',
    Email: 'john.smith@acme.com',
    Active_Quote: '$4,500',
    Additional_Notes: 'CEO loved the demo video. Ready to book ASAP. Needs contract by Friday.'
  };
  
  var result = sendSuperHotAlert(testLead, 'Elamin');
  Logger.log('Test result: ' + JSON.stringify(result));
}

// Helper function to escape HTML
function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ============================================
// TEST FUNCTION - Run this to test the email
// ============================================
function testSuperHotAlert() {
  var testLead = {
    Contact_Person: 'John Smith',
    Company_Name: 'Acme Corporation',
    Event_Name: 'Annual Holiday Gala 2026',
    Date_of_Event: new Date('2026-03-15'),
    Phone_Number: '(555) 123-4567',
    Email: 'john.smith@acme.com',
    Active_Quote: '$4,500',
    Additional_Notes: 'CEO loved the demo video. Ready to book ASAP. Needs contract by Friday.'
  };
  
  var result = sendSuperHotAlert(testLead, 'Elamin');
  Logger.log('Test result: ' + JSON.stringify(result));
}

// ============================================
// EMAIL GENERATION
// ============================================

function generateEmailLink(lead) {
  const to = lead.Email || '';
  const subject = encodeURIComponent('Dewayne Hill - Entertainment Proposal for ' + (lead.Event_Name || 'Your Event'));
  const body = encodeURIComponent('Hi ' + (lead.Contact_Person || 'there') + ',\n\n' +
    'Thank you for your interest in having Dewayne Hill perform at ' + (lead.Event_Name || 'your upcoming event') + '!\n\n' +
    'I wanted to follow up regarding your event' + (lead.Date_of_Event ? ' on ' + lead.Date_of_Event : '') + '.\n\n' +
    (lead.Active_Quote ? 'As discussed, the investment for this performance would be ' + lead.Active_Quote + '.\n\n' : '') +
    'Please let me know if you have any questions or if you\'re ready to move forward!\n\n' +
    'Best regards,\n' +
    'Dewayne Hill\n' +
    'America\'s Funniest Comedy Magician\n' +
    '855-306-2442\n' +
    'dhillmagic@gmail.com');
  
  return 'https://mail.google.com/mail/?view=cm&fs=1&to=' + to + '&su=' + subject + '&body=' + body;
}

// ============================================
// CONTRACT GENERATION
// ============================================

const CONTRACT_CONFIG = {
  CONTRACTS_FOLDER_ID: '1RB0Z5oCOZOV8JqgLn6BkPjiOr2JGdsa5',
  HEADER_IMAGE_ID: '1Pfm5l7OV-fidz9SMp32k5xh67ZizHYGM',
  SIGNATURE_IMAGE_ID: '15uuLVcldri2JhuvPF7LZtDzwJwoMPZcy',
  W9_LINK: 'https://drive.google.com/file/d/18o49tfzNCP1MxqgbsWuaRQVEDMSc0pCJ/view'
};

var IMAGE_CACHE = {};

function getImageAsBase64(fileId) {
  if (IMAGE_CACHE[fileId]) {
    return IMAGE_CACHE[fileId];
  }
  
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var contentType = blob.getContentType();
    var base64 = Utilities.base64Encode(blob.getBytes());
    var dataUri = 'data:' + contentType + ';base64,' + base64;
    IMAGE_CACHE[fileId] = dataUri;
    return dataUri;
  } catch (e) {
    console.error('Error getting image:', e);
    return '';
  }
}

function generateContract(rowIndex, currentUser) {
  try {
    const lead = getLeadByRow(rowIndex);
    if (!lead) {
      return { success: false, error: 'Lead not found' };
    }
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const contractsSheet = ss.getSheetByName(SHEETS.CONTRACTS);
    
    const timestamp = new Date();
    const invoiceNumber = 'INV' + Utilities.formatDate(timestamp, 'America/New_York', 'yyMMddHHmmssSSS');
    
    const headerImageData = getImageAsBase64(CONTRACT_CONFIG.HEADER_IMAGE_ID);
    const signatureImageData = getImageAsBase64(CONTRACT_CONFIG.SIGNATURE_IMAGE_ID);
    
    const contractData = {
      invoiceNumber: invoiceNumber,
      eventName: lead.Event_Name || lead.Company_Name || '',
      companyName: lead.Company_Name || lead.Event_Name || '',
      contactPerson: lead.Contact_Person || '',
      email: lead.Email || '',
      phone: lead.Phone_Number || '',
      dateOfEvent: formatEventDateLong(lead.Date_of_Event),
      timeZone: lead.Time_Zone || 'EST',
      venueCity: lead.Venue_City || lead.Company_City || '',
      venueState: lead.Venue_State || lead.Company_State || '',
      activeQuote: formatCurrency(lead.Active_Quote),
      headerImage: headerImageData,
      signatureImage: signatureImageData
    };
    
    contractData.location = contractData.venueCity ? 
      contractData.venueCity + ',' + contractData.venueState : '';
    contractData.lodgingBy = '- ' + contractData.eventName + ' - ' + contractData.contactPerson;
    contractData.contactDetails = contractData.contactPerson + ' - ' + contractData.email + ' - ' + contractData.phone;
    
    var filename;
    if (contractData.eventName && lead.Date_of_Event) {
      filename = contractData.eventName + ' - ' + lead.Date_of_Event;
    } else if (contractData.eventName) {
      filename = contractData.eventName + ' - ' + invoiceNumber;
    } else {
      filename = invoiceNumber;
    }
    
    const html = generateContractHTML(contractData);
    const blob = Utilities.newBlob(html, 'text/html', filename + '.html');
    const pdf = blob.getAs('application/pdf').setName(filename + '.pdf');
    
    const folder = DriveApp.getFolderById(CONTRACT_CONFIG.CONTRACTS_FOLDER_ID);
    const file = folder.createFile(pdf);
    
    const contractId = 'CON-' + String(contractsSheet.getLastRow()).padStart(5, '0');
    
    const contractRow = [
      contractId,
      invoiceNumber,
      lead.Record_ID,
      contractData.contactPerson,
      contractData.eventName,
      lead.Date_of_Event,
      lead.Active_Quote,
      new Date(),
      currentUser,
      filename + '.pdf',
      file.getId(),
      file.getUrl(),
      'Draft',
      'No',
      '',
      'No',
      ''
    ];
    contractsSheet.appendRow(contractRow);
    
    const leadsSheet = ss.getSheetByName(SHEETS.LEADS);
    const headers = leadsSheet.getRange(1, 1, 1, leadsSheet.getLastColumn()).getValues()[0];
    const contractIdCol = headers.indexOf('Contract_ID') + 1;
    if (contractIdCol > 0) {
      leadsSheet.getRange(rowIndex, contractIdCol).setValue(contractId);
    }
    
    // Track contract generation for EOD metrics
    incrementMetric(currentUser, 'Contracts_Generated', 1);
    
    return {
      success: true,
      contractId: contractId,
      invoiceNumber: invoiceNumber,
      filename: filename + '.pdf',
      url: file.getUrl(),
      message: 'Contract generated successfully'
    };
    
  } catch (error) {
    console.error('Contract generation error:', error);
    return { success: false, error: error.toString() };
  }
}

function generateContractHTML(data) {
  return '<!DOCTYPE html>\
<html>\
<head>\
  <meta charset="UTF-8">\
  <style>\
    @page { size: letter; margin: 0.4in; }\
    * { margin: 0; padding: 0; box-sizing: border-box; }\
    body { font-family: Helvetica, Arial, sans-serif; font-size: 9pt; color: #000; line-height: 1.3; }\
    .page { page-break-after: always; }\
    .page:last-child { page-break-after: avoid; }\
    .red-bar { height: 3px; background: #B91C1C; margin-bottom: 8px; }\
    .header-section { display: table; width: 100%; margin-bottom: 8px; }\
    .header-left { display: table-cell; width: 65%; vertical-align: middle; }\
    .header-right { display: table-cell; width: 35%; vertical-align: middle; text-align: right; font-size: 8pt; }\
    .header-img { max-width: 340px; height: auto; }\
    .divider { height: 1px; background: #E5E7EB; margin: 8px 0; }\
    .invoice-table { width: 100%; border-collapse: collapse; margin-bottom: 10px; }\
    .invoice-table td { padding: 3px 0; font-size: 8pt; vertical-align: middle; }\
    .invoice-table .label { width: 45%; }\
    .invoice-table .value { width: 55%; text-align: right; }\
    .invoice-table .value.bold { font-weight: bold; }\
    .invoice-table tr.border-bottom td { border-bottom: 1px solid #E5E7EB; }\
    .section-header { color: #B91C1C; font-weight: bold; font-size: 9pt; padding: 5px 0; margin: 10px 0 6px 0; text-decoration: underline; }\
    .bullet-list { margin: 0; padding-left: 20px; font-size: 7.5pt; }\
    .bullet-list li { margin-bottom: 2px; }\
    .w9-box { background: #F9FAFB; padding: 5px 10px; margin: 6px 0; display: table; width: 100%; font-size: 7.5pt; }\
    .w9-box .left { display: table-cell; width: 70%; }\
    .w9-box .right { display: table-cell; width: 30%; text-align: right; }\
    .w9-box a { color: #1D4ED8; text-decoration: underline; }\
    .cancel-box { background: #F9FAFB; border-left: 3px solid #B91C1C; padding: 8px 10px; margin: 6px 0; font-size: 7pt; text-align: justify; }\
    .subhead { font-weight: bold; font-size: 8pt; margin: 8px 0 4px 0; }\
    .body-text { font-size: 7.5pt; text-align: justify; margin-bottom: 4px; }\
    .deposit-box { background: #F9FAFB; padding: 12px 15px; margin: 10px 0; font-size: 8.5pt; }\
    .deposit-box .option { margin-bottom: 10px; }\
    .deposit-box .address { margin-left: 20px; margin-top: 8px; }\
    .deposit-box .address .name { font-weight: bold; font-size: 9pt; }\
    .signature-section { margin-top: 20px; margin-left: 10px; }\
    .signature-img { max-width: 180px; height: auto; }\
    .footer { margin-top: 20px; padding-top: 6px; border-top: 1px solid #E5E7EB; text-align: center; font-size: 7pt; color: #9CA3AF; }\
    .checkbox { display: inline-block; width: 12px; height: 12px; border: 1px solid #000; margin-right: 8px; vertical-align: middle; }\
  </style>\
</head>\
<body>\
\
<div class="page">\
  <div class="red-bar"></div>\
  <div class="header-section">\
    <div class="header-left">\
      <img src="' + data.headerImage + '" class="header-img" alt="Dewayne Hill">\
    </div>\
    <div class="header-right">\
      <strong>Invoice/Contract/Agreement</strong><br>\
      <strong>Vendor:</strong> Dewayne Hill<br>\
      dhillmagic@gmail.com<br>\
      855-306-2442<br>\
      <strong>Cell:</strong> 813-482-7888<br>\
      <strong>Tax ID:</strong> 235-xx-xxxxx\
    </div>\
  </div>\
  <div class="divider"></div>\
  <table class="invoice-table">\
    <tr class="border-bottom"><td class="label">Invoice Number:</td><td class="value bold">' + data.invoiceNumber + '</td></tr>\
    <tr><td class="label">Name of Company/Client/Event:</td><td class="value bold">' + data.eventName + '</td></tr>\
    <tr class="border-bottom"><td class="label">Date of Event:</td><td class="value bold">' + data.dateOfEvent + '</td></tr>\
    <tr><td class="label">Report Time for Dewayne</td><td class="value">06:00 pm ' + data.timeZone + '</td></tr>\
    <tr class="border-bottom"><td class="label">Address of Appearance:</td><td class="value">' + data.location + '</td></tr>\
    <tr><td class="label">Total Appearance Amount</td><td class="value bold">' + data.activeQuote + '</td></tr>\
    <tr><td class="label">Deposit Due, please see cancellation policy, 50%</td><td class="value bold">50% of Agreed Fee</td></tr>\
    <tr><td class="label">Balance Left after deposit is paid - due the day/night of performance</td><td class="value bold">Final 50% due day of event</td></tr>\
    <tr><td class="label">The lodging will be paid and taken care of by:</td><td class="value">' + data.lodgingBy + '</td></tr>\
    <tr><td class="label">Contact Details</td><td class="value">' + data.contactDetails + '</td></tr>\
    <tr><td class="label"></td><td class="value">' + data.eventName + '</td></tr>\
  </table>\
  <div class="section-header">Presenter/Performer Terms</div>\
  <ul class="bullet-list">\
    <li>Presenter will arrive early/or on time the day performing</li>\
    <li>Performer will perform stage/in person, that will be non offensive and comply with any HR policies.</li>\
    <li>Performer prefers to be paid by check, no personal checks unless previous cleared. Make Check payable to: <strong>Dewayne Hill</strong></li>\
  </ul>\
  <div class="w9-box">\
    <div class="left">Here is A signed W-9 for your records and accouting- PLEASE NOTICE THIS Doc is included</div>\
    <div class="right"><a href="' + CONTRACT_CONFIG.W9_LINK + '">W-9 link here</a></div>\
  </div>\
  <div class="section-header">Client/Account Terms</div>\
  <ul class="bullet-list">\
    <li>Client agrees to supply a copy of photos and video footage, should it be taken.</li>\
    <li>Payment for the show(s) must be made ON the same evening as show.</li>\
    <li>All transactions/performance fees are to remain confidential</li>\
  </ul>\
  <div class="cancel-box">\
    <strong>New Cancellation Policy, because the COVID-19 situation.</strong> Once signing this agreement, client/account is agrees to pay 50% of the agreed performance fee, regardless of reason for cancellation such as Force of Nature, Storms, Loss of Power, Lose of Venue or building collapsing or any other form considered weather or nature related, or state of emergency, issued by Local goverment OR Federal Government, If the first party (Dewayne Hill), failed to report due to being late where performing is no longer an option due to time constraints or due to health condition, the First Party will find a suitable replacement for the same or less amount of money or the refund of the deposits will be issued in the event of the First Party not being able to perform and will deduct the cost of any airfare that may be acquired. <strong>If cancellation occurs 31 day OR closer to the event - 100% of the agreed fee is required and owed.</strong> If this needs to be amended or discussed we are happy to do so. 50% deposit is required and to be sent effective immediately, either by Check or Credit Card payment.\
  </div>\
  <ul class="bullet-list">\
    <li>Dewayne Hill releases rights and asks that any informational material (press releases, advertisements, etc.) include his name and his performance times.</li>\
    <li>IF client decided to pay with Credit Card - a 2.9% fee maybe added to the total.</li>\
    <li>All paperwork related to the event needs to be signed and filled out prior to appearance.</li>\
  </ul>\
  <div class="subhead">Promotional Materials Approval</div>\
  <p class="body-text">Any promotional materials featuring Dewayne Hill, including but not limited to flyers, social media posts, digital ads, or event listings, must be approved in writing by Dewayne Hill prior to publication or distribution. The client is not authorized to create or distribute any promotional content using Dewayne\'s name, image, likeness, or performance details without prior written consent.</p>\
  <p class="body-text"><strong>Client understands and agrees that Dewayne Hill may have a booking assistant OR marketing person taking photos/or videos during the event.</strong></p>\
</div>\
\
<div class="page">\
  <div class="red-bar"></div>\
  <div class="section-header">Rider/AV Requirements</div>\
  <ul class="bullet-list">\
    <li>If performing for more than 60 people, Dewayne will need a sound system with a microphone, in a self contained microphone stand.</li>\
    <li>If Dewayne needs to supply his own sound system and micstand, arrangements need to be discussed and an additional fee my apply.</li>\
    <li>No podium microphones, the entertainer cannot perform behind a podium.</li>\
    <li>A standard microphone contained in a mic-stand center stage to be put in place</li>\
    <li><strong>Client Will supply a Free-Standing Mic-Stand</strong></li>\
  </ul>\
  <div class="section-header">Deposit, Please indicate</div>\
  <div class="deposit-box">\
    <div class="option"><span class="checkbox"></span> We will pay deposit with credit card, Please send the secure link</div>\
    <div class="option">\
      <span class="checkbox"></span> We will pay deposit wth a check:<br>\
      <span style="font-size: 7.5pt; color: #6B7280; margin-left: 20px;">If paying wth Check - all mail goes to the Florida address for tax purposes:</span>\
      <div class="address">\
        <div class="name">Dewayne Hill</div>\
        <div>13913 Lazy Oak Drive</div>\
        <div>Tampa, FL 33613</div>\
      </div>\
    </div>\
  </div>\
  <div class="signature-section">\
    <img src="' + data.signatureImage + '" class="signature-img" alt="Signature">\
  </div>\
  <div class="footer">Contract ' + data.invoiceNumber + ' &bull; Dewayne Hill Entertainment &bull; dhillmagic@gmail.com</div>\
</div>\
\
</body>\
</html>';
}

function formatEventDateLong(dateValue) {
  if (!dateValue) return 'To Be Determined';
  
  try {
    var date;
    if (dateValue instanceof Date) {
      date = dateValue;
    } else {
      date = new Date(dateValue);
    }
    
    if (isNaN(date.getTime())) return String(dateValue);
    
    var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    var months = ['January', 'February', 'March', 'April', 'May', 'June', 
                  'July', 'August', 'September', 'October', 'November', 'December'];
    
    return days[date.getDay()] + ', ' + months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
  } catch (e) {
    return String(dateValue);
  }
}

function formatCurrency(value) {
  if (!value) return '$0.00';
  
  var num = parseFloat(String(value).replace(/[$,]/g, ''));
  if (isNaN(num)) return '$0.00';
  
  return '$' + num.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function generateSingleContract(rowIndex, currentUser) {
  try {
    return generateContract(rowIndex, currentUser);
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================
// BULK ACTIONS
// ============================================

function bulkUpdateStatus(rowIndexes, newStatus, currentUser) {
  const results = [];
  
  rowIndexes.forEach(rowIndex => {
    const result = updateLead(rowIndex, { Lead_Status: newStatus }, currentUser);
    results.push({ rowIndex: rowIndex, success: result.success, error: result.error });
  });
  
  return results;
}

function bulkGenerateContracts(rowIndexes, currentUser) {
  const results = [];
  
  for (let i = 0; i < rowIndexes.length; i++) {
    const rowIndex = rowIndexes[i];
    try {
      const result = generateContract(rowIndex, currentUser);
      results.push({ 
        rowIndex: rowIndex, 
        index: i + 1,
        total: rowIndexes.length,
        success: result.success,
        error: result.error,
        url: result.url,
        filename: result.filename
      });
    } catch (e) {
      results.push({ 
        rowIndex: rowIndex, 
        index: i + 1,
        total: rowIndexes.length,
        success: false, 
        error: e.toString() 
      });
    }
    
    // Small delay between contracts to prevent rate limiting
    if (i < rowIndexes.length - 1) {
      Utilities.sleep(500);
    }
  }
  
  return results;
}

// ============================================
// RESEARCH QUEUE
// ============================================

function getResearchQueue() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.RESEARCH);
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) return [];
  
  const headers = data[0];
  const items = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    
    const item = { rowIndex: i + 1 };
    headers.forEach((header, index) => {
      let value = row[index];
      if (value instanceof Date) {
        value = formatDate(value);
      }
      item[header] = value;
    });
    
    if (item.Research_Status === 'Pending' || item.Research_Status === 'In Progress') {
      items.push(item);
    }
  }
  
  return items;
}

function addToResearchQueue(data, currentUser) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.RESEARCH);
  
  const newId = 'RES-' + String(sheet.getLastRow()).padStart(5, '0');
  
  const rowData = [
    newId,
    new Date(),
    currentUser,
    data.Looking_For || '',
    data.City_Location || '',
    data.Event_Type || '',
    data.Date_Needed || '',
    data.Guest_Count || '',
    data.Message || '',
    data.Budget_Range || '',
    'Pending',
    '', '', '', '', '', '', 'No', ''
  ];
  
  sheet.appendRow(rowData);
  
  return { success: true, researchId: newId };
}

function updateResearchItem(rowIndex, data, currentUser) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.RESEARCH);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i + 1);
  
  if (data.Found_Name) sheet.getRange(rowIndex, colIndex['Found_Name']).setValue(data.Found_Name);
  if (data.Found_Email) sheet.getRange(rowIndex, colIndex['Found_Email']).setValue(data.Found_Email);
  if (data.Found_Phone) sheet.getRange(rowIndex, colIndex['Found_Phone']).setValue(data.Found_Phone);
  if (data.Found_Company) sheet.getRange(rowIndex, colIndex['Found_Company']).setValue(data.Found_Company);
  
  sheet.getRange(rowIndex, colIndex['Researched_By']).setValue(currentUser);
  sheet.getRange(rowIndex, colIndex['Researched_Date']).setValue(new Date());
  sheet.getRange(rowIndex, colIndex['Research_Status']).setValue(data.Research_Status || 'Found');
  
  return { success: true };
}

function moveResearchToLeads(rowIndex, currentUser) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const researchSheet = ss.getSheetByName(SHEETS.RESEARCH);
  const headers = researchSheet.getRange(1, 1, 1, researchSheet.getLastColumn()).getValues()[0];
  const data = researchSheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  
  const research = {};
  headers.forEach((h, i) => research[h] = data[i]);
  
  const leadData = {
    Contact_Person: research.Found_Name,
    Email: research.Found_Email,
    Phone_Number: research.Found_Phone,
    Company_Name: research.Found_Company,
    Lead_Source: 'Bark.com',
    Lead_Status: 'Cold',
    Location: 'Active',
    Lead_Owner: currentUser,
    Category: research.Event_Type,
    Additional_Notes: 'From Bark Research:\n' + (research.Message || '') + '\nGuests: ' + (research.Guest_Count || 'N/A') + '\nBudget: ' + (research.Budget_Range || 'N/A')
  };
  
  const result = createLead(leadData, currentUser);
  
  if (result.success) {
    const colIndex = {};
    headers.forEach((h, i) => colIndex[h] = i + 1);
    researchSheet.getRange(rowIndex, colIndex['Moved_to_Leads']).setValue('Yes');
    researchSheet.getRange(rowIndex, colIndex['Lead_Record_ID']).setValue(result.recordId);
    researchSheet.getRange(rowIndex, colIndex['Research_Status']).setValue('Found');
  }
  
  return result;
}

// ============================================
// EOD METRICS & TEAM PERFORMANCE
// ============================================

function getTodayMetrics(currentUser) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEETS.EOD);
  
  if (!sheet) {
    return { Calls_Made: 0, Conversations: 0, Followups_Sent: 0, Leads_Upgraded: 0, Contracts_Generated: 0 };
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var colIndex = {};
  headers.forEach(function(h, i) { colIndex[h] = i; });
  
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  for (var i = data.length - 1; i >= 1; i--) {
    var rowDate = new Date(data[i][colIndex['Date']]);
    rowDate.setHours(0, 0, 0, 0);
    
    if (rowDate.getTime() === today.getTime() && data[i][colIndex['User']] === currentUser) {
      return {
        Calls_Made: data[i][colIndex['Calls_Made']] || 0,
        Conversations: data[i][colIndex['Conversations']] || 0,
        Followups_Sent: data[i][colIndex['Followups_Sent']] || 0,
        Leads_Upgraded: data[i][colIndex['Leads_Upgraded']] || 0,
        Contracts_Generated: data[i][colIndex['Contracts_Generated']] || 0,
        Training_Time: data[i][colIndex['Training_Time']] || '',
        Roleplay_Time: data[i][colIndex['Roleplay_Time']] || '',
        Call_Score_Self: data[i][colIndex['Call_Score_Self']] || '',
        Roadblocks: data[i][colIndex['Roadblocks']] || '',
        Special_Notes: data[i][colIndex['Special_Notes']] || ''
      };
    }
  }
  
  return { Calls_Made: 0, Conversations: 0, Followups_Sent: 0, Leads_Upgraded: 0, Contracts_Generated: 0 };
}

function incrementMetric(currentUser, metricName, amount) {
  if (!amount) amount = 1;
  
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEETS.EOD);
  
  if (!sheet) return { success: false, error: 'EOD sheet not found' };
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  var colIndex = {};
  headers.forEach(function(h, i) { colIndex[h] = i + 1; });
  
  var rowIndex = -1;
  for (var i = data.length - 1; i >= 1; i--) {
    var rowDate = new Date(data[i][0]);
    rowDate.setHours(0, 0, 0, 0);
    
    if (rowDate.getTime() === today.getTime() && data[i][1] === currentUser) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    var newRow = [];
    for (var j = 0; j < headers.length; j++) {
      var h = headers[j];
      if (h === 'Date') newRow.push(today);
      else if (h === 'User') newRow.push(currentUser);
      else if (h === metricName) newRow.push(amount);
      else newRow.push(0);
    }
    sheet.appendRow(newRow);
  } else {
    var currentValue = sheet.getRange(rowIndex, colIndex[metricName]).getValue() || 0;
    sheet.getRange(rowIndex, colIndex[metricName]).setValue(currentValue + amount);
  }
  
  return { success: true };
}

function submitEODReport(data, currentUser) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEETS.EOD);
  
  if (!sheet) return { success: false, error: 'EOD_Metrics sheet not found' };
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var colIndex = {};
  headers.forEach(function(h, i) { colIndex[h] = i + 1; });
  
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  var allData = sheet.getDataRange().getValues();
  var rowIndex = -1;
  
  for (var i = allData.length - 1; i >= 1; i--) {
    var rowDate = new Date(allData[i][0]);
    rowDate.setHours(0, 0, 0, 0);
    
    if (rowDate.getTime() === today.getTime() && allData[i][1] === currentUser) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    var newRow = [];
    for (var j = 0; j < headers.length; j++) {
      var h = headers[j];
      if (h === 'Date') newRow.push(today);
      else if (h === 'User') newRow.push(currentUser);
      else if (h === 'Submitted_Time') newRow.push(new Date());
      else if (data[h] !== undefined) newRow.push(data[h]);
      else newRow.push(0);
    }
    sheet.appendRow(newRow);
  } else {
    var fieldsToUpdate = ['Training_Time', 'Roleplay_Time', 'Roadblocks', 'Special_Notes', 'Call_Score_Self'];
    for (var k = 0; k < fieldsToUpdate.length; k++) {
      var field = fieldsToUpdate[k];
      if (data[field] !== undefined && colIndex[field]) {
        sheet.getRange(rowIndex, colIndex[field]).setValue(data[field]);
      }
    }
    if (colIndex['Submitted_Time']) {
      sheet.getRange(rowIndex, colIndex['Submitted_Time']).setValue(new Date());
    }
  }
  
  return { success: true, message: 'EOD Report submitted!' };
}

function getTeamPerformance(period) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEETS.EOD);
  
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var colIndex = {};
  headers.forEach(function(h, i) { colIndex[h] = i; });
  
  var now = new Date();
  var startDate = new Date();
  
  if (period === 'week') {
    startDate.setDate(now.getDate() - 7);
  } else if (period === 'month') {
    startDate.setDate(now.getDate() - 30);
  } else {
    startDate.setHours(0, 0, 0, 0);
  }
  
  var userStats = {};
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowDate = new Date(row[colIndex['Date']]);
    
    if (rowDate >= startDate) {
      var user = row[colIndex['User']];
      if (!user) continue;
      
      if (!userStats[user]) {
        userStats[user] = {
          name: user,
          calls: 0,
          conversations: 0,
          followups: 0,
          upgrades: 0,
          contracts: 0,
          avgScore: 0,
          scoreCount: 0,
          daysWorked: 0
        };
      }
      
      userStats[user].calls += Number(row[colIndex['Calls_Made']]) || 0;
      userStats[user].conversations += Number(row[colIndex['Conversations']]) || 0;
      userStats[user].followups += Number(row[colIndex['Followups_Sent']]) || 0;
      userStats[user].upgrades += Number(row[colIndex['Leads_Upgraded']]) || 0;
      userStats[user].contracts += Number(row[colIndex['Contracts_Generated']]) || 0;
      
      var score = Number(row[colIndex['Call_Score_Self']]);
      if (score > 0) {
        userStats[user].avgScore += score;
        userStats[user].scoreCount++;
      }
      userStats[user].daysWorked++;
    }
  }
  
  var result = [];
  for (var userName in userStats) {
    var u = userStats[userName];
    if (u.scoreCount > 0) {
      u.avgScore = Math.round(u.avgScore / u.scoreCount * 10) / 10;
    }
    result.push(u);
  }
  
  result.sort(function(a, b) { return b.calls - a.calls; });
  
  return result;
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
}

function formatDateForInput(date) {
  if (!date) return '';
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return year + '-' + month + '-' + day;
}

// ============================================
// SEARCH
// ============================================

function searchLeads(query, currentUser) {
  const leads = getLeads('all', currentUser);
  const q = query.toLowerCase();
  
  return leads.filter(function(lead) {
    return (
      (lead.Contact_Person && lead.Contact_Person.toLowerCase().indexOf(q) !== -1) ||
      (lead.Email && lead.Email.toLowerCase().indexOf(q) !== -1) ||
      (lead.Phone_Number && lead.Phone_Number.toLowerCase().indexOf(q) !== -1) ||
      (lead.Company_Name && lead.Company_Name.toLowerCase().indexOf(q) !== -1) ||
      (lead.Event_Name && lead.Event_Name.toLowerCase().indexOf(q) !== -1) ||
      (lead.Record_ID && lead.Record_ID.toLowerCase().indexOf(q) !== -1)
    );
  });
}

// ============================================
// NEXT LEAD IN QUEUE
// ============================================

function getNextLead(currentRowIndex, view, currentUser) {
  const leads = getLeads(view, currentUser);
  
  var currentIndex = -1;
  for (var i = 0; i < leads.length; i++) {
    if (leads[i].rowIndex === currentRowIndex) {
      currentIndex = i;
      break;
    }
  }
  
  if (currentIndex === -1 || currentIndex >= leads.length - 1) {
    return leads.length > 0 ? leads[0] : null;
  }
  
  return leads[currentIndex + 1];
}

// ============================================
// TEST FUNCTIONS
// ============================================

function testGetUsers() {
  var users = getUsers();
  Logger.log('Users found: ' + JSON.stringify(users));
  Logger.log('Count: ' + users.length);
  return users;
}

function testGenerateContract() {
  var result = generateContract(2, 'Dewayne');
  Logger.log(result);
}

// ============================================
// DIALPAD INTEGRATION
// ============================================

const DIALPAD_API_KEY = 'xJBc9HX9pDQcVLYL4WGEgQZ3DgNExsqwdfZscX7Vfr9Z5xM7jc6mUYqSXLAH8e95d5gHApKWkmTRwfL5ABfvQyDPHsq6pNdtJGh2'; // Replace with actual key
const DIALPAD_BASE_URL = 'https://dialpad.com/api/v2';

// Map Dialpad user emails to BEM user names
const DIALPAD_USER_MAP = {
  'dhillmagic@gmail.com': 'Dewayne',
  'christine.pacpro@gmail.com': 'Christine',
  'dhill2@dewaynehillshows.com': 'Dewayne 2',
  'jan.borja13@gmail.com': 'Jan'
};

/**
 * Get call stats from Dialpad for a date range
 */
function getDialpadCallStats(daysAgo) {
  daysAgo = daysAgo || 7;
  
  try {
    var url = DIALPAD_BASE_URL + '/stats';
    var payload = {
      'days_ago_start': 0,
      'days_ago_end': daysAgo,
      'export_type': 'records',
      'stat_type': 'calls',
      'timezone': 'America/New_York'
    };
    
    var options = {
      'method': 'post',
      'headers': {
        'Authorization': 'Bearer ' + DIALPAD_API_KEY,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    
    return { success: true, data: result };
  } catch (e) {
    Logger.log('Error getting Dialpad stats: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Get transcript for a specific call
 */
function getDialpadTranscript(callId) {
  try {
    var url = DIALPAD_BASE_URL + '/transcripts/' + callId;
    
    var options = {
      'method': 'get',
      'headers': {
        'Authorization': 'Bearer ' + DIALPAD_API_KEY
      },
      'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    
    return { success: true, data: result };
  } catch (e) {
    Logger.log('Error getting transcript for call ' + callId + ': ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Get list of Dialpad users
 */
function getDialpadUsers() {
  try {
    var url = DIALPAD_BASE_URL + '/users';
    
    var options = {
      'method': 'get',
      'headers': {
        'Authorization': 'Bearer ' + DIALPAD_API_KEY
      },
      'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    
    return { success: true, data: result };
  } catch (e) {
    Logger.log('Error getting Dialpad users: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Get recent calls for all users
 */
function getDialpadRecentCalls(daysAgo) {
  daysAgo = daysAgo || 7;
  
  try {
    // First get the stats export
    var statsResult = getDialpadCallStats(daysAgo);
    if (!statsResult.success) {
      return statsResult;
    }
    
    // The stats API returns a job ID, we need to poll for results
    // For now, return the initial response
    return statsResult;
  } catch (e) {
    Logger.log('Error getting recent calls: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Sync Dialpad calls to BEM
 * Pulls calls, gets transcripts, grades with OpenAI, saves to sheet
 */
function syncDialpadCalls(daysAgo) {
  daysAgo = daysAgo || 1;
  
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    
    // Get or create Dialpad_Calls sheet
    var callsSheet = ss.getSheetByName('Dialpad_Calls');
    if (!callsSheet) {
      callsSheet = ss.insertSheet('Dialpad_Calls');
      callsSheet.appendRow([
        'Call_ID', 'Date', 'Time', 'User', 'User_Email', 'Direction', 
        'Contact_Phone', 'Contact_Name', 'Duration_Sec', 'Duration_Min',
        'Dialpad_Summary', 'Dialpad_Action_Items', 'Dialpad_Purpose',
        'Transcript', 'AI_Score_Pct', 'AI_Breakdown', 'AI_Strengths', 
        'AI_Improvements', 'AI_Coaching_Tip', 'AI_Summary',
        'Lead_Record_ID', 'Lead_Name', 'Synced_At'
      ]);
      // Format header row
      callsSheet.getRange(1, 1, 1, 23).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
    }
    
    // Get existing call IDs to avoid duplicates
    var existingData = callsSheet.getDataRange().getValues();
    var existingCallIds = {};
    for (var i = 1; i < existingData.length; i++) {
      existingCallIds[existingData[i][0]] = true;
    }
    
    // Get leads for phone matching
    var leadsSheet = ss.getSheetByName(SHEETS.LEADS);
    var leadsData = leadsSheet.getDataRange().getValues();
    var leadsHeaders = leadsData[0];
    var phoneColIdx = leadsHeaders.indexOf('Phone_Number');
    var contactColIdx = leadsHeaders.indexOf('Contact_Person');
    var recordIdColIdx = leadsHeaders.indexOf('Record_ID');
    
    var phoneToLead = {};
    for (var i = 1; i < leadsData.length; i++) {
      var phone = normalizePhone(leadsData[i][phoneColIdx]);
      if (phone) {
        phoneToLead[phone] = {
          recordId: leadsData[i][recordIdColIdx],
          name: leadsData[i][contactColIdx]
        };
      }
    }
    
    // Fetch calls from Dialpad
    var startDate = new Date();
    startDate.setDate(startDate.getDate() - daysAgo);
    var timestamp = Math.floor(startDate.getTime());
    
    var url = DIALPAD_BASE_URL + '/call?started_after=' + timestamp + '&limit=50';
    var options = {
      'method': 'get',
      'headers': { 'Authorization': 'Bearer ' + DIALPAD_API_KEY },
      'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var callsData = JSON.parse(response.getContentText());
    
    if (!callsData.items || callsData.items.length === 0) {
      return { success: true, message: 'No calls found in the last ' + daysAgo + ' days', synced: 0 };
    }
    
    var syncedCount = 0;
    var skippedCount = 0;
    var gradedCount = 0;
    var newRows = [];
    
    for (var i = 0; i < callsData.items.length; i++) {
      var call = callsData.items[i];
      
      // Skip if already synced
      if (existingCallIds[call.call_id]) {
        skippedCount++;
        continue;
      }
      
      // Skip very short calls (< 30 seconds)
      var durationSec = Math.round((call.duration || 0) / 1000);
      if (durationSec < 30) {
        skippedCount++;
        continue;
      }
      
      var callDate = new Date(parseInt(call.date_started));
      var contactPhone = normalizePhone(call.external_number);
      var matchedLead = phoneToLead[contactPhone] || null;
      
      // Get AI Recap from Dialpad
      var dialpadSummary = '';
      var dialpadActionItems = '';
      var dialpadPurpose = '';
      
      try {
        var recapUrl = DIALPAD_BASE_URL + '/call/' + call.call_id + '/ai_recap';
        var recapResponse = UrlFetchApp.fetch(recapUrl, options);
        if (recapResponse.getResponseCode() === 200) {
          var recap = JSON.parse(recapResponse.getContentText());
          dialpadSummary = recap.summary ? recap.summary.content : '';
          dialpadActionItems = recap.action_items ? recap.action_items.map(function(a) { return a.content; }).join('; ') : '';
          dialpadPurpose = recap.purposes ? recap.purposes.map(function(p) { return p.content; }).join(', ') : '';
        }
      } catch (e) {
        Logger.log('Error getting AI recap for call ' + call.call_id + ': ' + e.toString());
      }
      
      // Get transcript from Dialpad
      var transcript = '';
      try {
        var transcriptUrl = DIALPAD_BASE_URL + '/transcripts/' + call.call_id;
        var transcriptResponse = UrlFetchApp.fetch(transcriptUrl, options);
        if (transcriptResponse.getResponseCode() === 200) {
          var transcriptData = JSON.parse(transcriptResponse.getContentText());
          if (transcriptData.lines) {
            transcript = transcriptData.lines
              .filter(function(line) { return line.type === 'transcript'; })
              .map(function(line) { return line.name + ': ' + line.content; })
              .join('\n');
          }
        }
      } catch (e) {
        Logger.log('Error getting transcript for call ' + call.call_id + ': ' + e.toString());
      }
      
      // Grade with OpenAI (only for calls > 60 seconds with transcript)
      var aiScore = '';
      var aiBreakdown = '';
      var aiStrengths = '';
      var aiImprovements = '';
      var aiCoachingTip = '';
      var aiSummary = '';
      
      if (transcript.length > 100 && durationSec > 60) {
        try {
          var gradeResult = gradeCallTranscript(
            transcript, 
            call.contact ? call.contact.name : '',
            matchedLead ? matchedLead.name : '',
            dialpadSummary
          );
          
          if (gradeResult.success && gradeResult.grade) {
            var g = gradeResult.grade;
            aiScore = g.overall_score + '%';
            aiBreakdown = JSON.stringify(g.breakdown);
            aiStrengths = (g.strengths || []).join('; ');
            aiImprovements = (g.improvements || []).join('; ');
            aiCoachingTip = g.coaching_tip || '';
            aiSummary = g.summary || '';
            gradedCount++;
          }
        } catch (e) {
          Logger.log('Error grading call ' + call.call_id + ': ' + e.toString());
        }
        
        // Rate limit - wait between OpenAI calls
        Utilities.sleep(500);
      }
      
      // Prepare row data
      var row = [
        call.call_id,
        Utilities.formatDate(callDate, 'America/New_York', 'yyyy-MM-dd'),
        Utilities.formatDate(callDate, 'America/New_York', 'HH:mm:ss'),
        call.target ? call.target.name : '',
        call.target ? call.target.email : '',
        call.direction || '',
        call.external_number || '',
        call.contact ? call.contact.name : '',
        durationSec,
        Math.round(durationSec / 60 * 10) / 10, // Duration in minutes
        dialpadSummary,
        dialpadActionItems,
        dialpadPurpose,
        transcript.substring(0, 5000), // Limit transcript length
        aiScore,
        aiBreakdown,
        aiStrengths,
        aiImprovements,
        aiCoachingTip,
        aiSummary,
        matchedLead ? matchedLead.recordId : '',
        matchedLead ? matchedLead.name : '',
        new Date().toISOString()
      ];
      
      newRows.push(row);
      syncedCount++;
      
      // Rate limit for Dialpad API
      Utilities.sleep(200);
    }
    
    // Write all new rows at once
    if (newRows.length > 0) {
      callsSheet.getRange(callsSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }
    
    return { 
      success: true, 
      synced: syncedCount,
      skipped: skippedCount,
      graded: gradedCount,
      message: 'Synced ' + syncedCount + ' calls, graded ' + gradedCount + ' with AI, skipped ' + skippedCount + ' (duplicates or too short)'
    };
    
  } catch (e) {
    Logger.log('Error syncing Dialpad calls: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Normalize phone number for matching
 */
function normalizePhone(phone) {
  if (!phone) return '';
  return String(phone).replace(/[^0-9]/g, '').slice(-10);
}

/**
 * Grade a call transcript using AI (OpenAI GPT)
 * Returns percentage score (0-100%) with detailed breakdown and comments
 */
function gradeCallTranscript(transcript, contactName, eventName, dialpadSummary) {
  if (!transcript || transcript.length < 50) {
    return { success: false, error: 'Transcript too short to grade' };
  }
  
  try {
    // Use OpenAI API for grading
    var OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    
    if (!OPENAI_API_KEY) {
      return { success: false, error: 'OpenAI API key not configured. Add OPENAI_API_KEY to Script Properties.' };
    }
    
    var prompt = `You are a sales call quality analyst for Dewayne Hill, "America's Funniest Comedy Magician" who books entertainment for corporate events, parties, and galas.

Grade this sales call on a scale of 0-100% based on these criteria:

1. **Opening & Rapport** (0-20%): Professional greeting, building connection, energy level
2. **Discovery & Qualifying** (0-20%): Asked about event date, type, budget, guest count, decision maker
3. **Value Proposition & Pitch** (0-20%): Explained services clearly, highlighted unique value, addressed needs
4. **Objection Handling** (0-20%): Handled concerns professionally, provided solutions, stayed positive
5. **Close & Next Steps** (0-20%): Clear call-to-action, set follow-up, got commitment

Contact: ${contactName || 'Unknown'}
Event: ${eventName || 'Unknown'}
${dialpadSummary ? 'Call Summary: ' + dialpadSummary : ''}

TRANSCRIPT:
${transcript.substring(0, 6000)}

Respond ONLY in this exact JSON format, no other text:
{
  "overall_score": <number 0-100>,
  "breakdown": {
    "opening_rapport": <0-20>,
    "discovery_qualifying": <0-20>,
    "value_proposition": <0-20>,
    "objection_handling": <0-20>,
    "close_next_steps": <0-20>
  },
  "strengths": ["specific strength 1", "specific strength 2", "specific strength 3"],
  "improvements": ["specific improvement 1", "specific improvement 2", "specific improvement 3"],
  "key_moments": ["notable moment or quote from the call"],
  "coaching_tip": "One actionable tip for the agent to improve",
  "summary": "2-3 sentence summary of call quality and outcome"
}`;

    var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      'method': 'post',
      'headers': {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + OPENAI_API_KEY
      },
      'payload': JSON.stringify({
        'model': 'gpt-4o-mini',
        'messages': [
          { 'role': 'system', 'content': 'You are a sales call quality analyst. Always respond with valid JSON only, no markdown or extra text.' },
          { 'role': 'user', 'content': prompt }
        ],
        'max_tokens': 1500,
        'temperature': 0.3
      }),
      'muteHttpExceptions': true
    });
    
    var result = JSON.parse(response.getContentText());
    
    if (result.choices && result.choices[0] && result.choices[0].message) {
      var gradeText = result.choices[0].message.content;
      // Extract JSON from response
      var jsonMatch = gradeText.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        var grade = JSON.parse(jsonMatch[0]);
        return { success: true, grade: grade };
      }
    }
    
    // Log error response for debugging
    if (result.error) {
      Logger.log('OpenAI API error: ' + JSON.stringify(result.error));
      return { success: false, error: result.error.message || 'OpenAI API error' };
    }
    
    return { success: false, error: 'Could not parse AI response' };
  } catch (e) {
    Logger.log('Error grading transcript: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Get call history with transcripts and grades for a user
 */
function getUserCallHistory(userName, daysAgo) {
  daysAgo = daysAgo || 7;
  
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var callsSheet = ss.getSheetByName('Dialpad_Calls');
    
    if (!callsSheet) {
      return { success: true, calls: [] };
    }
    
    var data = callsSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, calls: [] };
    }
    
    var headers = data[0];
    var calls = [];
    
    var cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysAgo);
    cutoffDate.setHours(0, 0, 0, 0); // Start of day
    
    // Find column indices
    var dateIdx = headers.indexOf('Date');
    var userIdx = headers.indexOf('User');
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var callDateValue = row[dateIdx];
      
      // Handle both Date objects and strings
      var callDate;
      if (callDateValue instanceof Date) {
        callDate = callDateValue;
      } else {
        callDate = new Date(callDateValue);
      }
      
      var callUser = row[userIdx];
      
      // Check if date is valid and within range
      if (!isNaN(callDate.getTime()) && callDate >= cutoffDate) {
        if (!userName || callUser === userName) {
          var call = {};
          headers.forEach(function(h, idx) {
            call[h] = row[idx];
          });
          // Map to frontend expected names
          call.Date = row[dateIdx];
          call.User = row[userIdx];
          call.Direction = row[headers.indexOf('Direction')];
          call.Contact_Phone = row[headers.indexOf('Contact_Phone')];
          call.Contact_Name = row[headers.indexOf('Contact_Name')];
          call.Duration_Sec = row[headers.indexOf('Duration_Sec')];
          call.AI_Score = row[headers.indexOf('AI_Score_Pct')] || '';
          call.AI_Summary = row[headers.indexOf('AI_Summary')] || '';
          calls.push(call);
        }
      }
    }
    
    return { success: true, calls: calls };
  } catch (e) {
    Logger.log('Error getting call history: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Get team call performance summary
 */
function getTeamCallPerformance(period) {
  period = period || 'week';
  
  var daysAgo = period === 'today' ? 1 : period === 'week' ? 7 : 30;
  
  try {
    var callHistory = getUserCallHistory(null, daysAgo);
    if (!callHistory.success) {
      return callHistory;
    }
    
    // Aggregate by user
    var userStats = {};
    
    callHistory.calls.forEach(function(call) {
      var user = call.User || 'Unknown';
      if (!userStats[user]) {
        userStats[user] = {
          name: user,
          totalCalls: 0,
          totalDuration: 0,
          avgScore: 0,
          scores: [],
          conversations: 0
        };
      }
      
      userStats[user].totalCalls++;
      userStats[user].totalDuration += (call.Duration_Sec || 0);
      
      if (call.Duration_Sec > 60) {
        userStats[user].conversations++;
      }
      
      if (call.AI_Score) {
        userStats[user].scores.push(call.AI_Score);
      }
    });
    
    // Calculate averages
    var result = [];
    for (var user in userStats) {
      var stats = userStats[user];
      if (stats.scores.length > 0) {
        var sum = stats.scores.reduce(function(a, b) { return a + b; }, 0);
        stats.avgScore = sum / stats.scores.length;
      }
      delete stats.scores;
      result.push(stats);
    }
    
    // Sort by total calls descending
    result.sort(function(a, b) { return b.totalCalls - a.totalCalls; });
    
    return { success: true, data: result };
  } catch (e) {
    Logger.log('Error getting team call performance: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Initiate a call via Dialpad (click-to-call)
 */
function initiateDialpadCall(phoneNumber, userId) {
  try {
    var url = DIALPAD_BASE_URL + '/users/' + userId + '/initiate_call';
    
    var options = {
      'method': 'post',
      'headers': {
        'Authorization': 'Bearer ' + DIALPAD_API_KEY,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify({
        'phone_number': phoneNumber
      }),
      'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    
    return { success: true, data: result };
  } catch (e) {
    Logger.log('Error initiating Dialpad call: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Test Dialpad API connection
 */
function testDialpadConnection() {
  try {
    var result = getDialpadUsers();
    Logger.log('Dialpad connection test: ' + JSON.stringify(result));
    return result;
  } catch (e) {
    Logger.log('Dialpad connection test failed: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Test getting call list from Dialpad
 */
function testGetCallList() {
  try {
    var url = DIALPAD_BASE_URL + '/call';
    
    // Get calls from last 7 days - use Unix timestamp in milliseconds
    var sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
    var timestamp = Math.floor(sevenDaysAgo.getTime()); // Unix timestamp in ms
    
    var params = {
      'started_after': timestamp,
      'limit': 50
    };
    
    var queryString = Object.keys(params).map(function(key) {
      return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
    }).join('&');
    
    var options = {
      'method': 'get',
      'headers': {
        'Authorization': 'Bearer ' + DIALPAD_API_KEY
      },
      'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(url + '?' + queryString, options);
    var responseCode = response.getResponseCode();
    var result = JSON.parse(response.getContentText());
    
    Logger.log('Response Code: ' + responseCode);
    Logger.log('Call List Result: ' + JSON.stringify(result).substring(0, 3000));
    
    if (result.items) {
      Logger.log('Number of calls found: ' + result.items.length);
    }
    
    return { success: responseCode === 200, code: responseCode, data: result };
  } catch (e) {
    Logger.log('Error getting call list: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Test getting call stats
 */
function testGetCallStats() {
  var result = getDialpadCallStats(7);
  Logger.log('Call Stats Result: ' + JSON.stringify(result));
  return result;
}

/**
 * Test getting AI Recap for a call
 */
function testGetAIRecap() {
  // First get a recent call
  var url = DIALPAD_BASE_URL + '/call';
  var sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
  var timestamp = Math.floor(sevenDaysAgo.getTime());
  
  var options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + DIALPAD_API_KEY
    },
    'muteHttpExceptions': true
  };
  
  var response = UrlFetchApp.fetch(url + '?started_after=' + timestamp + '&limit=10', options);
  var calls = JSON.parse(response.getContentText());
  
  if (!calls.items || calls.items.length === 0) {
    Logger.log('No calls found');
    return;
  }
  
  // Find a call with duration > 60 seconds (actual conversation)
  var callWithDuration = calls.items.find(function(c) { return c.duration > 60000; });
  
  if (!callWithDuration) {
    Logger.log('No calls with duration > 60s found');
    callWithDuration = calls.items[0];
  }
  
  var callId = callWithDuration.call_id;
  Logger.log('Testing AI Recap for call: ' + callId);
  Logger.log('Call details: ' + callWithDuration.target.name + ' -> ' + callWithDuration.contact.name + ' (' + Math.round(callWithDuration.duration/1000) + 's)');
  
  // Get AI Recap
  var recapUrl = DIALPAD_BASE_URL + '/call/' + callId + '/ai_recap';
  var recapResponse = UrlFetchApp.fetch(recapUrl, options);
  var recapCode = recapResponse.getResponseCode();
  var recapResult = recapResponse.getContentText();
  
  Logger.log('AI Recap Response Code: ' + recapCode);
  Logger.log('AI Recap Result: ' + recapResult);
  
  // Also try getting transcript
  var transcriptUrl = DIALPAD_BASE_URL + '/transcripts/' + callId;
  var transcriptResponse = UrlFetchApp.fetch(transcriptUrl, options);
  var transcriptCode = transcriptResponse.getResponseCode();
  var transcriptResult = transcriptResponse.getContentText();
  
  Logger.log('Transcript Response Code: ' + transcriptCode);
  Logger.log('Transcript Result: ' + transcriptResult.substring(0, 1500));
}

/**
 * Test OpenAI grading with a real transcript
 */
function testOpenAIGrading() {
  // Get a call with transcript
  var url = DIALPAD_BASE_URL + '/call';
  var sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
  var timestamp = Math.floor(sevenDaysAgo.getTime());
  
  var options = {
    'method': 'get',
    'headers': { 'Authorization': 'Bearer ' + DIALPAD_API_KEY },
    'muteHttpExceptions': true
  };
  
  var response = UrlFetchApp.fetch(url + '?started_after=' + timestamp + '&limit=20', options);
  var calls = JSON.parse(response.getContentText());
  
  // Find a call with good duration
  var call = calls.items.find(function(c) { return c.duration > 60000; });
  if (!call) {
    Logger.log('No suitable call found');
    return;
  }
  
  Logger.log('Testing grading for call: ' + call.call_id);
  Logger.log('Agent: ' + call.target.name + ', Contact: ' + call.contact.name + ', Duration: ' + Math.round(call.duration/1000) + 's');
  
  // Get transcript
  var transcriptUrl = DIALPAD_BASE_URL + '/transcripts/' + call.call_id;
  var transcriptResponse = UrlFetchApp.fetch(transcriptUrl, options);
  var transcriptData = JSON.parse(transcriptResponse.getContentText());
  
  var transcript = transcriptData.lines
    .filter(function(line) { return line.type === 'transcript'; })
    .map(function(line) { return line.name + ': ' + line.content; })
    .join('\n');
  
  Logger.log('Transcript length: ' + transcript.length + ' chars');
  Logger.log('Transcript preview: ' + transcript.substring(0, 500));
  
  // Get Dialpad summary
  var recapUrl = DIALPAD_BASE_URL + '/call/' + call.call_id + '/ai_recap';
  var recapResponse = UrlFetchApp.fetch(recapUrl, options);
  var recap = JSON.parse(recapResponse.getContentText());
  var dialpadSummary = recap.summary ? recap.summary.content : '';
  
  // Grade with OpenAI
  Logger.log('Sending to OpenAI for grading...');
  var gradeResult = gradeCallTranscript(transcript, call.contact.name, '', dialpadSummary);
  
  Logger.log('Grade Result: ' + JSON.stringify(gradeResult, null, 2));
  
  if (gradeResult.success) {
    var g = gradeResult.grade;
    Logger.log('=== GRADE SUMMARY ===');
    Logger.log('Overall Score: ' + g.overall_score + '%');
    Logger.log('Breakdown: ' + JSON.stringify(g.breakdown));
    Logger.log('Strengths: ' + (g.strengths || []).join(', '));
    Logger.log('Improvements: ' + (g.improvements || []).join(', '));
    Logger.log('Coaching Tip: ' + g.coaching_tip);
  }
}

function testSyncDebug() {
  var daysAgo = 7;
  var startDate = new Date();
  startDate.setDate(startDate.getDate() - daysAgo);
  var timestamp = Math.floor(startDate.getTime());
  
  Logger.log('Timestamp: ' + timestamp);
  
  var url = DIALPAD_BASE_URL + '/call?started_after=' + timestamp + '&limit=50';
  Logger.log('URL: ' + url);
  
  var options = {
    'method': 'get',
    'headers': { 'Authorization': 'Bearer ' + DIALPAD_API_KEY },
    'muteHttpExceptions': true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  
  Logger.log('Response code: ' + response.getResponseCode());
  Logger.log('Response body: ' + response.getContentText().substring(0, 500));
}

function testSync() {
  var result = syncDialpadCalls(7);
  Logger.log(JSON.stringify(result));
}

function testGetAllCalls() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var callsSheet = ss.getSheetByName('Dialpad_Calls');
  var data = callsSheet.getDataRange().getValues();
  
  Logger.log('Total rows: ' + data.length);
  Logger.log('Headers: ' + JSON.stringify(data[0]));
  
  if (data.length > 1) {
    Logger.log('Row 2 Date value: ' + data[1][1]);
    Logger.log('Row 2 Date type: ' + typeof data[1][1]);
  }
}

function testGetCallHistory() {
  var result = getUserCallHistory(null, 7);
  Logger.log('Success: ' + result.success);
  Logger.log('Calls count: ' + (result.calls ? result.calls.length : 0));
  if (result.calls && result.calls.length > 0) {
    Logger.log('First call: ' + JSON.stringify(result.calls[0]));
  }
  if (result.error) {
    Logger.log('Error: ' + result.error);
  }
}

function testDirectCall() {
  var result = getUserCallHistory('', 7);
  Logger.log('Result: ' + JSON.stringify(result).substring(0, 500));
}

function testCallHistoryDirect() {
  var result = getUserCallHistory('', 7);
  Logger.log('Type: ' + typeof result);
  Logger.log('Is null: ' + (result === null));
  Logger.log('Result: ' + JSON.stringify(result).substring(0, 300));
}

function getCallHistoryForWeb(userName, daysAgo) {
  var result = getUserCallHistory(userName, daysAgo);
  // Convert to JSON and back to ensure all Date objects are serialized
  return JSON.parse(JSON.stringify(result));
}

// ============================================
// getAnalyticsData and getDailyTrends functions
// ============================================

function getAnalyticsData(period) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var leadsSheet = ss.getSheetByName('Leads');
    var eodSheet = ss.getSheetByName('EOD_Metrics');
    var callsSheet = ss.getSheetByName('Dialpad_Calls');
    
    // Calculate date range based on period
    var now = new Date();
    var periodStart = new Date();
    var prevPeriodStart = new Date();
    var prevPeriodEnd = new Date();
    
    if (period === 'today') {
      periodStart.setHours(0, 0, 0, 0);
      prevPeriodStart = new Date(periodStart);
      prevPeriodStart.setDate(prevPeriodStart.getDate() - 1);
      prevPeriodEnd = new Date(periodStart);
    } else if (period === 'week') {
      periodStart.setDate(now.getDate() - 7);
      prevPeriodStart = new Date(periodStart);
      prevPeriodStart.setDate(prevPeriodStart.getDate() - 7);
      prevPeriodEnd = new Date(periodStart);
    } else if (period === 'month') {
      periodStart.setDate(now.getDate() - 30);
      prevPeriodStart = new Date(periodStart);
      prevPeriodStart.setDate(prevPeriodStart.getDate() - 30);
      prevPeriodEnd = new Date(periodStart);
    } else if (period === 'quarter') {
      periodStart.setDate(now.getDate() - 90);
      prevPeriodStart = new Date(periodStart);
      prevPeriodStart.setDate(prevPeriodStart.getDate() - 90);
      prevPeriodEnd = new Date(periodStart);
    }
    
    // Get leads data
    var leadsData = leadsSheet.getDataRange().getValues();
    var leadsHeaders = leadsData[0];
    var leads = [];
    
    var colIdx = {};
    leadsHeaders.forEach(function(h, i) { colIdx[h] = i; });
    
    for (var i = 1; i < leadsData.length; i++) {
      var row = leadsData[i];
      leads.push({
        rowIndex: i + 1,
        Location: row[colIdx['Location']] || '',
        Lead_Status: row[colIdx['Lead_Status']] || '',
        Lead_Source: row[colIdx['Lead_Source']] || '',
        Lead_Owner: row[colIdx['Lead_Owner']] || '',
        Active_Quote: parseFloat(String(row[colIdx['Active_Quote']] || '0').replace(/[^0-9.-]/g, '')) || 0,
        Event_Name: row[colIdx['Event_Name']] || '',
        Contact_Person: row[colIdx['Contact_Person']] || '',
        Date_Last_Touch: row[colIdx['Date_Last_Touch']] || ''
      });
    }
    
    // Calculate KPIs
    var activeLeads = leads.filter(function(l) { return l.Location === 'Active'; });
    var hotLeads = activeLeads.filter(function(l) { 
      var s = (l.Lead_Status || '').toLowerCase();
      return s.indexOf('hot') >= 0 || s.indexOf('super') >= 0 || s === 'contract lead';
    });
    var bookedLeads = leads.filter(function(l) { return l.Location === 'Booked'; });
    
    var pipelineValue = activeLeads.reduce(function(sum, l) { return sum + l.Active_Quote; }, 0);
    var hotValue = hotLeads.reduce(function(sum, l) { return sum + l.Active_Quote; }, 0);
    var bookedValue = bookedLeads.reduce(function(sum, l) { return sum + l.Active_Quote; }, 0);
    
    // Conversion rate: $ Booked / ($ Pipeline + $ Booked)
    var totalValue = pipelineValue + bookedValue;
    var conversionRate = totalValue > 0 ? Math.round((bookedValue / totalValue) * 100) : 0;
    
    // Sources by $ value
    var sourceMap = {};
    activeLeads.forEach(function(l) {
      var src = l.Lead_Source || 'Unknown';
      if (!sourceMap[src]) sourceMap[src] = 0;
      sourceMap[src] += l.Active_Quote;
    });
    
    var sourceData = Object.keys(sourceMap).map(function(name) {
      return { name: name, value: sourceMap[name] };
    }).sort(function(a, b) { return b.value - a.value; }).slice(0, 8);
    
    // Funnel
    var funnel = {
      total: leads.length,
      totalValue: leads.reduce(function(s, l) { return s + l.Active_Quote; }, 0),
      active: activeLeads.length,
      activeValue: pipelineValue,
      hot: hotLeads.length,
      hotValue: hotValue,
      booked: bookedLeads.length,
      bookedValue: bookedValue
    };
    
    // Team leaderboard - aggregate by owner
    var teamMap = {};
    
    // Get booked value per owner
    bookedLeads.forEach(function(l) {
      var owner = l.Lead_Owner || 'Unknown';
      if (!teamMap[owner]) teamMap[owner] = { name: owner, bookedValue: 0, bookedCount: 0, pipelineValue: 0, calls: 0, conversations: 0, aiScore: 0, points: 0 };
      teamMap[owner].bookedValue += l.Active_Quote;
      teamMap[owner].bookedCount++;
    });
    
    // Get pipeline value per owner
    activeLeads.forEach(function(l) {
      var owner = l.Lead_Owner || 'Unknown';
      if (!teamMap[owner]) teamMap[owner] = { name: owner, bookedValue: 0, bookedCount: 0, pipelineValue: 0, calls: 0, conversations: 0, aiScore: 0, points: 0 };
      teamMap[owner].pipelineValue += l.Active_Quote;
    });
    
    // Get EOD metrics
    if (eodSheet) {
      var eodData = eodSheet.getDataRange().getValues();
      var eodHeaders = eodData[0];
      var eodColIdx = {};
      eodHeaders.forEach(function(h, i) { eodColIdx[h] = i; });
      
      for (var i = 1; i < eodData.length; i++) {
        var row = eodData[i];
        var dateVal = row[eodColIdx['Date']];
        if (dateVal && new Date(dateVal) >= periodStart) {
          var user = row[eodColIdx['User']] || '';
          if (user && teamMap[user]) {
            teamMap[user].calls += parseInt(row[eodColIdx['Calls_Made']] || 0);
            teamMap[user].conversations += parseInt(row[eodColIdx['Conversations']] || 0);
          }
        }
      }
    }
    
    // Get AI scores from Dialpad_Calls
    if (callsSheet) {
      var callsData = callsSheet.getDataRange().getValues();
      var callsHeaders = callsData[0];
      var callsColIdx = {};
      callsHeaders.forEach(function(h, i) { callsColIdx[h] = i; });
      
      var userScores = {};
      for (var i = 1; i < callsData.length; i++) {
        var row = callsData[i];
        var user = row[callsColIdx['User']] || '';
        var score = parseFloat(String(row[callsColIdx['AI_Score']] || '').replace('%', '')) || 0;
        if (score > 0 && score < 1) score = score * 100;
        
        if (user && score > 0) {
          if (!userScores[user]) userScores[user] = [];
          userScores[user].push(score);
        }
      }
      
      // Calculate average AI score per user
      Object.keys(userScores).forEach(function(user) {
        if (teamMap[user]) {
          var scores = userScores[user];
          teamMap[user].aiScore = Math.round(scores.reduce(function(a, b) { return a + b; }, 0) / scores.length);
        }
      });
    }
    
    // Calculate points
    Object.keys(teamMap).forEach(function(user) {
      var m = teamMap[user];
      m.points = Math.round(
        (m.bookedValue / 100) +  // $100 = 1 point
        (m.bookedCount * 25) +   // Each deal = 25 points
        (m.calls * 1) +          // Each call = 1 point
        (m.conversations * 3) +  // Each convo = 3 points
        (m.aiScore * 0.5)        // AI score bonus
      );
    });
    
    var teamLeaderboard = Object.values(teamMap).sort(function(a, b) { return b.points - a.points; });
    
    // Recent bookings
    var recentBookings = bookedLeads.slice(0, 5).map(function(l) {
      return {
        event: l.Event_Name,
        contact: l.Contact_Person,
        owner: l.Lead_Owner,
        value: l.Active_Quote,
        date: l.Date_Last_Touch
      };
    });
    
    return {
      kpis: {
        pipeline: { value: pipelineValue, count: activeLeads.length },
        hotLeads: { value: hotValue, count: hotLeads.length },
        booked: { value: bookedValue, count: bookedLeads.length, change: 0 },
        conversionRate: { value: conversionRate }
      },
      funnel: funnel,
      sourceData: sourceData,
      teamLeaderboard: teamLeaderboard,
      recentBookings: recentBookings
    };
    
  } catch (e) {
    Logger.log('Error in getAnalyticsData: ' + e.toString());
    return null;
  }
}

// ============================================
//getDailyTrends function
// ============================================

function getDailyTrends(period) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var eodSheet = ss.getSheetByName('EOD_Metrics');
    var leadsSheet = ss.getSheetByName('Leads');
    
    // Determine number of days
    var days = 7;
    if (period === 'today') days = 1;
    else if (period === 'week') days = 7;
    else if (period === 'month') days = 30;
    else if (period === 'quarter') days = 90;
    else if (typeof period === 'number') days = period;
    
    var trends = [];
    var now = new Date();
    
    // Initialize all days
    for (var i = days - 1; i >= 0; i--) {
      var d = new Date(now);
      d.setDate(d.getDate() - i);
      d.setHours(0, 0, 0, 0);
      trends.push({
        date: d.toISOString().split('T')[0],
        calls: 0,
        conversations: 0,
        quoted: 0,
        booked: 0
      });
    }
    
    // Get EOD data for calls and conversations
    if (eodSheet) {
      var data = eodSheet.getDataRange().getValues();
      var headers = data[0];
      var colIdx = {};
      headers.forEach(function(h, i) { colIdx[h] = i; });
      
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var dateVal = row[colIdx['Date']];
        if (!dateVal) continue;
        
        var rowDate = new Date(dateVal);
        rowDate.setHours(0, 0, 0, 0);
        var dateStr = rowDate.toISOString().split('T')[0];
        
        // Find matching trend day
        for (var j = 0; j < trends.length; j++) {
          if (trends[j].date === dateStr) {
            trends[j].calls += parseInt(row[colIdx['Calls_Made']] || 0);
            trends[j].conversations += parseInt(row[colIdx['Conversations']] || 0);
            break;
          }
        }
      }
    }
    
    // Get Leads data for $ quoted and deals booked per day
    if (leadsSheet) {
      var leadsData = leadsSheet.getDataRange().getValues();
      var leadsHeaders = leadsData[0];
      var leadsColIdx = {};
      leadsHeaders.forEach(function(h, i) { leadsColIdx[h] = i; });
      
      for (var i = 1; i < leadsData.length; i++) {
        var row = leadsData[i];
        var location = row[leadsColIdx['Location']] || '';
        var lastTouch = row[leadsColIdx['Date_Last_Touch']];
        var quote = parseFloat(String(row[leadsColIdx['Active_Quote']] || '0').replace(/[^0-9.-]/g, '')) || 0;
        var status = row[leadsColIdx['Lead_Status']] || '';
        
        if (!lastTouch) continue;
        
        var touchDate = new Date(lastTouch);
        touchDate.setHours(0, 0, 0, 0);
        var dateStr = touchDate.toISOString().split('T')[0];
        
        // Find matching trend day
        for (var j = 0; j < trends.length; j++) {
          if (trends[j].date === dateStr) {
            // Count $ quoted (any lead touched with a quote value)
            if (quote > 0 && (status === 'Hot' || status === 'Super Hot' || status === 'Warm' || status === 'Contract Lead')) {
              trends[j].quoted += quote;
            }
            
            // Count bookings (leads that moved to Booked on this day)
            if (location === 'Booked') {
              trends[j].booked += 1;
            }
            break;
          }
        }
      }
    }
    
    return trends;
    
  } catch (e) {
    Logger.log('Error in getDailyTrends: ' + e.toString());
    return [];
  }
}

function getDailyTrendsForWeb(period) {
  var result = getDailyTrends(period);
  return JSON.parse(JSON.stringify(result));
}

function testAnalytics() {
  var result = getAnalyticsData('week');
  Logger.log(JSON.stringify(result));
}

function getAnalyticsDataForWeb(period) {
  var result = getAnalyticsData(period);
  return JSON.parse(JSON.stringify(result));
}

// ============================================
// AUTO-SYNC DIALPAD CALLS (For Time-Based Trigger)
// ============================================

function autoSyncDialpadCalls() {
  try {
    var result = syncDialpadCalls(1); // Sync last 1 day
    Logger.log('Auto-sync completed: ' + JSON.stringify(result));
    return result;
  } catch (e) {
    Logger.log('Auto-sync failed: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}
