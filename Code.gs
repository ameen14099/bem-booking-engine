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

  Logger.log('Settings sheet headers: ' + JSON.stringify(headers));

  headers.forEach((header, colIndex) => {
    if (header && header.toString().trim() !== '') {
      const headerName = header.toString().trim();
      options[headerName] = [];
      for (let row = 1; row < data.length; row++) {
        const value = data[row][colIndex];
        if (value && value.toString().trim() !== '') {
          options[headerName].push(value.toString().trim());
        }
      }
      Logger.log('Column ' + headerName + ' has ' + options[headerName].length + ' values');
    }
  });

  Logger.log('Month_of_Event values: ' + JSON.stringify(options['Month_of_Event']));

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
      
      // Include users with a name - treat missing/empty Active as active
      // Only exclude if explicitly set to 'No' or false
      if (user.Name) {
        var isActive = user.Active === undefined || user.Active === '' || user.Active === 'Yes' || user.Active === true;
        if (isActive || user.Active !== 'No') {
          users.push(user);
        }
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
      clientsCount: 0,
      parkingLotCount: 0,
      storageCount: 0,
      pipelineValue: '$0',
      contractsReview: 0,
      todayFollowUps: 0,
      myLeadsCount: 0,
      dhillCount: 0,
      hotLeadsPercentage: 0,
      avgCallsMade: 0,
      avgConversations: 0,
      overdueFollowups: 0,
      noFollowupSet: 0,
      staleLeads: 0
    };
  }

  const headers = data[0];
  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i);

  let totalActive = 0;
  let hotLeads = 0;
  let warmLeads = 0;
  let bookedCount = 0;
  let clientsCount = 0;
  let parkingLotCount = 0;
  let storageCount = 0;
  let pipelineValue = 0;
  let contractsReview = 0;
  let todayFollowUps = 0;
  let myLeadsCount = 0;
  let dhillCount = 0;
  let overdueFollowups = 0;
  let noFollowupSet = 0;
  let staleLeads = 0;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const sevenDaysAgo = new Date(today);
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const location = row[colIndex['Location']];
    const status = row[colIndex['Lead_Status']];
    const quote = row[colIndex['Active_Quote']];
    const contractId = row[colIndex['Contract_ID']];
    const followupDate = row[colIndex['Date_Next_Followup']];
    const owner = row[colIndex['Lead_Owner']];
    const dhillReview = row[colIndex['Dhill_Call_List']];
    const lastTouch = row[colIndex['Date_Last_Touch']];

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

      // Contracts to review (leads that have a contract generated)
      if (contractId && String(contractId).trim() !== '') {
        contractsReview++;
      }

      // Today's follow ups
      if (followupDate instanceof Date) {
        const followup = new Date(followupDate);
        followup.setHours(0, 0, 0, 0);
        if (followup.getTime() === today.getTime()) {
          todayFollowUps++;
        }
        // SF3: Overdue follow-ups (Date_Next_Followup < TODAY)
        if (followup.getTime() < today.getTime()) {
          overdueFollowups++;
        }
      } else {
        // SF4: No follow-up set
        noFollowupSet++;
      }

      // SF5: Stale leads (Date_Last_Touch > 7 days ago)
      if (lastTouch instanceof Date) {
        const lastTouchDate = new Date(lastTouch);
        lastTouchDate.setHours(0, 0, 0, 0);
        if (lastTouchDate.getTime() < sevenDaysAgo.getTime()) {
          staleLeads++;
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
      // Check if it's a past client (event date in the past)
      const eventDate = row[colIndex['Date_of_Event']];
      if (eventDate instanceof Date) {
        const event = new Date(eventDate);
        event.setHours(0, 0, 0, 0);
        if (event.getTime() < today.getTime()) {
          clientsCount++;
        }
      }
    } else if (location === 'Clients') {
      clientsCount++;
    } else if (location === 'Parking Lot') {
      parkingLotCount++;
    } else if (location === 'Storage') {
      storageCount++;
    }
  }

  // D7: Hot leads percentage = Math.round((hotLeads / totalActive) * 100)
  const hotLeadsPercentage = totalActive > 0 ? Math.round((hotLeads / totalActive) * 100) : 0;

  // D8 & D9: Calculate avg calls and conversations from EOD_Metrics
  const eodAverages = getEODAverages();

  // Calculate values for unified dashboard
  var hotValue = 0;
  var warmValue = 0;
  var overdueValue = 0;
  var staleValue = 0;
  var noFollowupValue = 0;
  var hotGoingCold = 0;
  var bookedMTD = 0;
  var bookedCountMTD = 0;
  var leadSources = {};

  // New funnel data (Cold, Warm, Hot, Contract Lead, Booked)
  var coldCount = 0;
  var coldValue = 0;
  var warmCount = 0;
  var warmFunnelValue = 0;
  var hotCount = 0;
  var hotFunnelValue = 0;
  var contractCount = 0;
  var contractValue = 0;
  var bookedFunnelCount = 0;
  var bookedFunnelValue = 0;

  // Get month start for MTD calculations
  var monthStart = new Date(today.getFullYear(), today.getMonth(), 1);

  // Second pass for additional metrics
  for (var j = 1; j < data.length; j++) {
    var r = data[j];
    var loc = r[colIndex['Location']];
    var stat = r[colIndex['Lead_Status']];
    var qt = r[colIndex['Active_Quote']];
    var lastT = r[colIndex['Date_Last_Touch']];
    var followD = r[colIndex['Date_Next_Followup']];
    var src = r[colIndex['Lead_Source']];
    var bookedDate = r[colIndex['Date_Booked']];

    var quoteNum = 0;
    if (qt) {
      quoteNum = parseFloat(String(qt).replace(/[$,]/g, '')) || 0;
    }

    if (loc === 'Active') {
      // Hot value
      if (stat === 'Hot' || stat === 'Super Hot' || stat === 'Contract Lead') {
        hotValue += quoteNum;
        // Hot going cold (hot leads with last touch > 5 days)
        if (lastT instanceof Date) {
          var lastTouchD = new Date(lastT);
          var fiveDaysAgo = new Date(today);
          fiveDaysAgo.setDate(fiveDaysAgo.getDate() - 5);
          if (lastTouchD < fiveDaysAgo) {
            hotGoingCold++;
          }
        }
      } else if (stat === 'Warm') {
        warmValue += quoteNum;
      }

      // Overdue value
      if (followD instanceof Date) {
        var followDate = new Date(followD);
        followDate.setHours(0, 0, 0, 0);
        if (followDate < today) {
          overdueValue += quoteNum;
        }
      } else {
        noFollowupValue += quoteNum;
      }

      // Stale value
      if (lastT instanceof Date) {
        var ltd = new Date(lastT);
        ltd.setHours(0, 0, 0, 0);
        if (ltd < sevenDaysAgo) {
          staleValue += quoteNum;
        }
      }

      // Funnel stages based on Lead_Status (Cold, Warm, Hot, Contract Lead)
      if (stat === 'Cold' || stat === 'New' || stat === 'Not Contacted') {
        coldCount++;
        coldValue += quoteNum;
      } else if (stat === 'Warm') {
        warmCount++;
        warmFunnelValue += quoteNum;
      } else if (stat === 'Hot' || stat === 'Super Hot') {
        hotCount++;
        hotFunnelValue += quoteNum;
      } else if (stat === 'Contract Lead') {
        contractCount++;
        contractValue += quoteNum;
      }

      // Lead sources
      if (src) {
        leadSources[src] = (leadSources[src] || 0) + 1;
      }
    }

    // Booked leads (for funnel)
    if (loc === 'Booked') {
      bookedFunnelCount++;
      bookedFunnelValue += quoteNum;

      // Booked MTD
      if (bookedDate instanceof Date) {
        var bd = new Date(bookedDate);
        if (bd >= monthStart) {
          bookedMTD += quoteNum;
          bookedCountMTD++;
        }
      }
    }
  }

  // Convert lead sources to array
  var sourcesArray = [];
  for (var srcName in leadSources) {
    sourcesArray.push({ name: srcName, count: leadSources[srcName] });
  }
  sourcesArray.sort(function(a, b) { return b.count - a.count; });

  return {
    totalActive,
    hotLeads,
    warmLeads,
    bookedCount,
    clientsCount,
    parkingLotCount,
    storageCount,
    pipelineValue: '$' + pipelineValue.toLocaleString(),
    contractsReview,
    todayFollowUps,
    myLeadsCount,
    dhillCount,
    hotLeadsPercentage,
    avgCallsMade: eodAverages.avgCalls,
    avgConversations: eodAverages.avgConversations,
    overdueFollowups,
    noFollowupSet,
    staleLeads,
    // Unified dashboard fields
    hotValue: '$' + hotValue.toLocaleString(),
    warmValue: '$' + warmValue.toLocaleString(),
    overdueValue: '$' + overdueValue.toLocaleString(),
    staleValue: '$' + staleValue.toLocaleString(),
    noFollowupValue: '$' + noFollowupValue.toLocaleString(),
    hotGoingCold: hotGoingCold,
    bookedMTD: '$' + bookedMTD.toLocaleString(),
    bookedCountMTD: bookedCountMTD,
    // Total Booked (all time) - used by Total Booked card
    totalBookedValue: '$' + bookedFunnelValue.toLocaleString(),
    // Funnel data (Cold, Warm, Hot, Contract Lead, Booked)
    coldCount: coldCount,
    coldValue: coldValue,
    warmCount: warmCount,
    warmFunnelValue: warmFunnelValue,
    hotCount: hotCount,
    hotFunnelValue: hotFunnelValue,
    contractCount: contractCount,
    contractValue: contractValue,
    bookedFunnelCount: bookedFunnelCount,
    bookedFunnelValue: bookedFunnelValue,
    leadSources: sourcesArray
  };
}

// ============================================
// EOD AVERAGES (D8, D9, E6)
// ============================================

function getEODAverages() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.EOD);

    if (!sheet) {
      return { avgCalls: 0, avgConversations: 0 };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { avgCalls: 0, avgConversations: 0 };
    }

    const headers = data[0];
    const colIndex = {};
    headers.forEach((h, i) => colIndex[h] = i);

    let totalCalls = 0;
    let totalConversations = 0;
    let rowCount = 0;

    // Get last 30 days of data for averages
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    thirtyDaysAgo.setHours(0, 0, 0, 0);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dateVal = row[colIndex['Date']];

      if (dateVal instanceof Date && dateVal >= thirtyDaysAgo) {
        const calls = parseInt(row[colIndex['Calls_Made']]) || 0;
        const convos = parseInt(row[colIndex['Conversations']]) || 0;

        if (calls > 0 || convos > 0) {
          totalCalls += calls;
          totalConversations += convos;
          rowCount++;
        }
      }
    }

    return {
      avgCalls: rowCount > 0 ? Math.round(totalCalls / rowCount) : 0,
      avgConversations: rowCount > 0 ? Math.round(totalConversations / rowCount) : 0
    };
  } catch (e) {
    Logger.log('Error in getEODAverages: ' + e.toString());
    return { avgCalls: 0, avgConversations: 0 };
  }
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
    const contractId = row[colIndex['Contract_ID']];
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
        // Show leads that have a contract generated (Contract_ID exists)
        include = (location === 'Active' && contractId && String(contractId).trim() !== '');
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
      case 'clients':
        // Clients = past booked clients (event date is in the past) or Location = 'Clients'
        if (location === 'Clients') {
          include = true;
        } else if (location === 'Booked') {
          const eventDateRaw = row[colIndex['Date_of_Event']];
          // Handle both Date objects and string dates
          let eventDate = null;
          if (eventDateRaw instanceof Date) {
            eventDate = new Date(eventDateRaw);
          } else if (eventDateRaw && typeof eventDateRaw === 'string') {
            eventDate = new Date(eventDateRaw);
          }
          if (eventDate && !isNaN(eventDate.getTime())) {
            eventDate.setHours(0, 0, 0, 0);
            include = (eventDate.getTime() < today.getTime());
          }
        }
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
      // SF3: Overdue follow-ups filter
      case 'overdue':
        if (location === 'Active' && followupDate instanceof Date) {
          const followup = new Date(followupDate);
          followup.setHours(0, 0, 0, 0);
          include = (followup.getTime() < today.getTime());
        }
        break;
      // SF4: No follow-up set filter
      case 'nofollowup':
        include = (location === 'Active' && !followupDate);
        break;
      // SF5: Stale leads filter (last touch > 7 days ago)
      case 'stale':
        if (location === 'Active') {
          const lastTouch = row[colIndex['Date_Last_Touch']];
          if (lastTouch instanceof Date) {
            const sevenDaysAgo = new Date(today);
            sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
            const lastTouchDate = new Date(lastTouch);
            lastTouchDate.setHours(0, 0, 0, 0);
            include = (lastTouchDate.getTime() < sevenDaysAgo.getTime());
          }
        }
        break;
      default:
        include = (location === 'Active');
    }

    if (include) {
      const lead = { rowIndex: i + 1 };
      headers.forEach((header, index) => {
        let value = row[index];
        if (value instanceof Date) {
          // Use formatDateForDisplay for event dates to prevent timezone shift
          if (header === 'Date_of_Event' || header === 'Event_Date') {
            value = formatDateForDisplay(value);
          } else {
            value = formatDate(value);
          }
        }
        lead[header] = value;
      });

      // FL1: Add days since last touched
      const lastTouchVal = row[colIndex['Date_Last_Touch']];
      if (lastTouchVal instanceof Date) {
        const lastTouch = new Date(lastTouchVal);
        lastTouch.setHours(0, 0, 0, 0);
        lead.daysSinceTouch = Math.floor((today - lastTouch) / (1000 * 60 * 60 * 24));
      } else {
        lead.daysSinceTouch = null;
      }

      // FL2: Follow-up overdue flag
      if (followupDate instanceof Date) {
        const followup = new Date(followupDate);
        followup.setHours(0, 0, 0, 0);
        lead.followupOverdue = followup.getTime() < today.getTime();
      } else {
        lead.followupOverdue = false;
      }

      // FL3: No follow-up set flag
      lead.noFollowupSet = !followupDate;

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
// BULK MOVE STORAGE → ACTIVE (A4)
// ============================================

function bulkMoveStorageToActive(rowIndexes, currentUser) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.LEADS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i + 1);

  const results = [];
  const now = new Date();

  rowIndexes.forEach(rowIndex => {
    try {
      // Update Location to Active
      sheet.getRange(rowIndex, colIndex['Location']).setValue('Active');
      sheet.getRange(rowIndex, colIndex['Date_Last_Touch']).setValue(now);
      sheet.getRange(rowIndex, colIndex['Last_Modified_By']).setValue(currentUser);

      results.push({ rowIndex: rowIndex, success: true });
    } catch (e) {
      results.push({ rowIndex: rowIndex, success: false, error: e.toString() });
    }
  });

  return {
    success: true,
    moved: results.filter(r => r.success).length,
    failed: results.filter(r => !r.success).length,
    results: results
  };
}

// ============================================
// ADVANCED SEARCH WITH MULTI-FILTER (N5)
// ============================================

function advancedSearchLeads(filters, currentUser) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.LEADS);
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) return [];

  const headers = data[0];
  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i);

  const leads = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Skip empty rows
    if (!row[colIndex['Contact_Person']] && !row[colIndex['Email']] && !row[colIndex['Company_Name']]) {
      continue;
    }

    const location = row[colIndex['Location']] || 'Active';

    // Only search active leads by default
    if (location !== 'Active' && !filters.includeAll) continue;

    let match = true;

    // Filter by event month
    if (filters.eventMonth && match) {
      const monthOfEvent = row[colIndex['Month_of_Event']] || '';
      if (monthOfEvent.toLowerCase() !== filters.eventMonth.toLowerCase()) {
        match = false;
      }
    }

    // Filter by state (company or venue)
    if (filters.state && match) {
      const companyState = (row[colIndex['Company_State']] || '').toLowerCase();
      const venueState = (row[colIndex['Venue_State']] || '').toLowerCase();
      const filterState = filters.state.toLowerCase();
      if (companyState !== filterState && venueState !== filterState) {
        match = false;
      }
    }

    // Filter by category
    if (filters.category && match) {
      const category = row[colIndex['Category']] || '';
      if (category.toLowerCase() !== filters.category.toLowerCase()) {
        match = false;
      }
    }

    // Fix #13: Filter by What Did We Pitch
    if (filters.whatPitched && match) {
      const whatPitched = row[colIndex['What_Pitched']] || '';
      if (whatPitched.toLowerCase() !== filters.whatPitched.toLowerCase()) {
        match = false;
      }
    }

    // Fix #13: Filter by Event Date Range
    if ((filters.eventDateFrom || filters.eventDateTo) && match) {
      const eventDate = row[colIndex['Date_of_Event']];
      if (eventDate instanceof Date) {
        const eventTime = eventDate.getTime();
        if (filters.eventDateFrom) {
          const fromDate = new Date(filters.eventDateFrom);
          fromDate.setHours(0, 0, 0, 0);
          if (eventTime < fromDate.getTime()) match = false;
        }
        if (filters.eventDateTo && match) {
          const toDate = new Date(filters.eventDateTo);
          toDate.setHours(23, 59, 59, 999);
          if (eventTime > toDate.getTime()) match = false;
        }
      } else {
        match = false;
      }
    }

    // Fix #13: Filter by Follow-up Date Range
    if ((filters.followupFrom || filters.followupTo) && match) {
      const followupDate = row[colIndex['Date_Next_Followup']];
      if (followupDate instanceof Date) {
        const followupTime = followupDate.getTime();
        if (filters.followupFrom) {
          const fromDate = new Date(filters.followupFrom);
          fromDate.setHours(0, 0, 0, 0);
          if (followupTime < fromDate.getTime()) match = false;
        }
        if (filters.followupTo && match) {
          const toDate = new Date(filters.followupTo);
          toDate.setHours(23, 59, 59, 999);
          if (followupTime > toDate.getTime()) match = false;
        }
      } else {
        match = false;
      }
    }

    // Filter by quote range
    if ((filters.quoteMin || filters.quoteMax) && match) {
      const quote = parseFloat(String(row[colIndex['Active_Quote']] || '0').replace(/[^0-9.-]/g, '')) || 0;
      if (filters.quoteMin && quote < filters.quoteMin) match = false;
      if (filters.quoteMax && quote > filters.quoteMax) match = false;
    }

    // Filter by status
    if (filters.status && match) {
      const status = row[colIndex['Lead_Status']] || '';
      if (status.toLowerCase() !== filters.status.toLowerCase()) {
        match = false;
      }
    }

    // Filter by owner
    if (filters.owner && match) {
      const owner = row[colIndex['Lead_Owner']] || '';
      if (owner.toLowerCase() !== filters.owner.toLowerCase()) {
        match = false;
      }
    }

    // Filter by source
    if (filters.source && match) {
      const source = row[colIndex['Lead_Source']] || '';
      if (source.toLowerCase() !== filters.source.toLowerCase()) {
        match = false;
      }
    }

    // Text search
    if (filters.searchText && match) {
      const searchText = filters.searchText.toLowerCase();
      const searchFields = [
        row[colIndex['Contact_Person']],
        row[colIndex['Email']],
        row[colIndex['Phone_Number']],
        row[colIndex['Company_Name']],
        row[colIndex['Event_Name']],
        row[colIndex['Record_ID']]
      ].join(' ').toLowerCase();

      if (searchFields.indexOf(searchText) === -1) {
        match = false;
      }
    }

    if (match) {
      const lead = { rowIndex: i + 1 };
      headers.forEach((header, index) => {
        let value = row[index];
        if (value instanceof Date) {
          value = formatDate(value);
        }
        lead[header] = value;
      });

      // FL1: Add days since last touched
      if (row[colIndex['Date_Last_Touch']] instanceof Date) {
        const lastTouch = new Date(row[colIndex['Date_Last_Touch']]);
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        lastTouch.setHours(0, 0, 0, 0);
        lead.daysSinceTouch = Math.floor((today - lastTouch) / (1000 * 60 * 60 * 24));
      } else {
        lead.daysSinceTouch = null;
      }

      // FL2: Follow-up overdue flag
      if (row[colIndex['Date_Next_Followup']] instanceof Date) {
        const followup = new Date(row[colIndex['Date_Next_Followup']]);
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        followup.setHours(0, 0, 0, 0);
        lead.followupOverdue = followup.getTime() < today.getTime();
      } else {
        lead.followupOverdue = false;
      }

      // FL3: No follow-up set flag
      lead.noFollowupSet = !row[colIndex['Date_Next_Followup']];

      leads.push(lead);
    }
  }

  // Sort by Priority_Score descending
  leads.sort((a, b) => {
    const scoreA = parseFloat(a.Priority_Score) || 0;
    const scoreB = parseFloat(b.Priority_Score) || 0;
    return scoreB - scoreA;
  });

  return leads;
}

// ============================================
// DUPLICATE DETECTION (N2)
// ============================================

function checkForDuplicates(email, phone, excludeRecordId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.LEADS);
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) return { hasDuplicate: false, duplicates: [] };

  const headers = data[0];
  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i);

  const normalizedEmail = (email || '').toLowerCase().trim();
  const normalizedPhone = normalizePhone(phone);

  const duplicates = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowRecordId = row[colIndex['Record_ID']];

    // Skip the current lead being edited (exclude from duplicate check)
    if (excludeRecordId && rowRecordId === excludeRecordId) {
      continue;
    }

    const rowEmail = (row[colIndex['Email']] || '').toLowerCase().trim();
    const rowPhone = normalizePhone(row[colIndex['Phone_Number']]);

    let isDuplicate = false;
    let matchType = [];

    // Check email match
    if (normalizedEmail && rowEmail && normalizedEmail === rowEmail) {
      isDuplicate = true;
      matchType.push('EMAIL');
    }

    // Check phone match
    if (normalizedPhone && rowPhone && normalizedPhone === rowPhone) {
      isDuplicate = true;
      matchType.push('PHONE');
    }

    if (isDuplicate) {
      duplicates.push({
        rowIndex: i + 1,
        recordId: rowRecordId,
        contactPerson: row[colIndex['Contact_Person']],
        email: row[colIndex['Email']],
        phone: row[colIndex['Phone_Number']],
        status: row[colIndex['Lead_Status']],
        location: row[colIndex['Location']],
        matchType: matchType.join(', ')
      });
    }
  }

  return {
    hasDuplicate: duplicates.length > 0,
    duplicates: duplicates
  };
}

// ============================================
// LEAD VALIDATION (VL1, VL2)
// ============================================

function validateLeadForStatus(leadData, newStatus) {
  const errors = [];

  // VL1: Hot/Super Hot/Contract Lead requires phone, event date, followup date
  if (['Hot', 'Super Hot', 'Contract Lead'].includes(newStatus)) {
    if (!leadData.Phone_Number || leadData.Phone_Number.trim() === '') {
      errors.push('Phone number is required for ' + newStatus + ' status');
    }
    if (!leadData.Date_of_Event) {
      errors.push('Event date is required for ' + newStatus + ' status');
    }
    if (!leadData.Date_Next_Followup) {
      errors.push('Follow-up date is required for ' + newStatus + ' status');
    }
  }

  // VL2: Booked requires quote, city, state, contact, email
  if (newStatus === 'Booked') {
    if (!leadData.Active_Quote || parseFloat(String(leadData.Active_Quote).replace(/[^0-9.-]/g, '')) <= 0) {
      errors.push('Quote amount is required for Booked status');
    }
    if (!leadData.Company_City && !leadData.Venue_City) {
      errors.push('City is required for Booked status');
    }
    if (!leadData.Company_State && !leadData.Venue_State) {
      errors.push('State is required for Booked status');
    }
    if (!leadData.Contact_Person || leadData.Contact_Person.trim() === '') {
      errors.push('Contact person is required for Booked status');
    }
    if (!leadData.Email || leadData.Email.trim() === '') {
      errors.push('Email is required for Booked status');
    }
  }

  return {
    isValid: errors.length === 0,
    errors: errors
  };
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

/**
 * Get client by email for rebooking
 */
function getClientByEmail(email) {
  if (!email) return null;

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.LEADS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const colIndex = {};
  headers.forEach((header, index) => {
    colIndex[header] = index;
  });

  // Find the most recent lead with this email
  let latestLead = null;
  let latestDate = new Date(0);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[colIndex['Email']] === email) {
      const created = new Date(row[colIndex['Created_Date']] || 0);
      if (created > latestDate) {
        latestDate = created;
        const lead = {};
        headers.forEach((header, index) => {
          lead[header] = row[index];
        });
        latestLead = lead;
      }
    }
  }

  return latestLead;
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

  // Check if auto-assign is enabled and no owner specified
  let assignedOwner = leadData.Lead_Owner;
  if (!assignedOwner || assignedOwner === '') {
    try {
      const settings = getAutoAssignSettings();
      if (settings.enabled && settings.onCreate) {
        const agent = autoAssignLead({ Record_ID: newId, Lead_Source: leadData.Lead_Source }, settings.method);
        if (agent) {
          assignedOwner = agent.name;
        }
      }
    } catch (e) {
      Logger.log('Auto-assign error: ' + e.toString());
    }
  }

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
      case 'Lead_Owner':
        return assignedOwner || leadData.Lead_Owner || '';
      case 'Priority_Score':
        return calculatePriorityScore(leadData);
      default:
        return leadData[header] || '';
    }
  });

  sheet.appendRow(rowData);

  // Get user info for logging
  const user = getUserWithId(currentUser);
  const userId = user ? user.User_ID : '';

  // Log activity
  logActivity({
    leadId: newId,
    userId: userId,
    userName: currentUser,
    type: 'Created',
    title: 'Lead Created',
    description: 'Source: ' + (leadData.Lead_Source || 'Unknown')
  });

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
  const oldOwner = currentData[colIndex['Lead_Owner'] - 1];
  const recordId = currentData[colIndex['Record_ID'] - 1];
  const newStatus = updates.Lead_Status || oldStatus;
  const newOwner = updates.Lead_Owner || oldOwner;

  // Get user info for logging
  const user = getUserWithId(currentUser);
  const userId = user ? user.User_ID : '';

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

      // Log Super Hot activity
      logActivity({
        leadId: recordId,
        userId: userId,
        userName: currentUser,
        type: 'Super_Hot',
        title: 'Marked as Super Hot',
        description: 'Lead escalated to Super Hot status',
        oldValue: oldStatus,
        newValue: newStatus
      });
    }

    // Log status change activity
    logActivity({
      leadId: recordId,
      userId: userId,
      userName: currentUser,
      type: 'Status_Change',
      title: 'Status: ' + oldStatus + ' → ' + newStatus,
      description: '',
      oldValue: oldStatus,
      newValue: newStatus
    });

    // Track status upgrade for EOD metrics
    if (['Hot', 'Warm', 'Contract Lead'].includes(newStatus) &&
        !['Hot', 'Warm', 'Contract Lead', 'Super Hot'].includes(oldStatus)) {
      incrementMetric(currentUser, 'Leads_Upgraded', 1);
    }
  }

  // Log owner change if changed
  if (newOwner !== oldOwner && updates.Lead_Owner) {
    logActivity({
      leadId: recordId,
      userId: userId,
      userName: currentUser,
      type: 'Owner_Change',
      title: 'Owner: ' + (oldOwner || 'None') + ' → ' + newOwner,
      oldValue: oldOwner || 'None',
      newValue: newOwner
    });
  }

  // Log location change if moved
  if (moved) {
    logActivity({
      leadId: recordId,
      userId: userId,
      userName: currentUser,
      type: 'Location_Move',
      title: 'Moved to ' + newLocation,
      description: 'Status changed to ' + newStatus,
      newValue: newLocation
    });
  }

  // Log edit activity ONLY for fields that actually changed (excluding status/owner which are logged separately)
  const actualChanges = [];
  Object.keys(updates).forEach(function(field) {
    if (field === 'Lead_Status' || field === 'Lead_Owner') return; // Already logged separately
    if (!colIndex[field]) return; // Skip fields not in sheet

    const oldVal = currentData[colIndex[field] - 1];
    const newVal = updates[field];

    // Normalize values for comparison (handle dates, empty strings, nulls)
    const oldNorm = normalizeValueForComparison(oldVal);
    const newNorm = normalizeValueForComparison(newVal);

    if (oldNorm !== newNorm) {
      actualChanges.push({
        field: field,
        oldValue: oldNorm || '(empty)',
        newValue: newNorm || '(empty)'
      });
    }
  });

  if (actualChanges.length > 0) {
    let description;
    if (actualChanges.length <= 5) {
      // Show detailed changes
      description = actualChanges.map(function(c) {
        return formatFieldName(c.field) + ' (' + truncateValue(c.oldValue) + ' → ' + truncateValue(c.newValue) + ')';
      }).join(', ');
    } else {
      // Too many changes - summarize
      description = actualChanges.length + ' fields updated: ' + actualChanges.map(function(c) {
        return formatFieldName(c.field);
      }).join(', ');
    }

    logActivity({
      leadId: recordId,
      userId: userId,
      userName: currentUser,
      type: 'Edited',
      title: 'Lead Updated',
      description: description
    });
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
  Logger.log('=== generateContract CALLED ===');
  Logger.log('Row Index: ' + rowIndex);
  Logger.log('Current User: ' + currentUser);

  try {
    // Step 1: Get Lead Data
    Logger.log('Step 1: Getting lead data...');
    const lead = getLeadByRow(rowIndex);
    if (!lead) {
      Logger.log('ERROR: Lead not found for row ' + rowIndex);
      return { success: false, error: 'Lead not found for row ' + rowIndex };
    }
    Logger.log('Lead found: ' + JSON.stringify({
      Record_ID: lead.Record_ID,
      Contact_Person: lead.Contact_Person,
      Event_Name: lead.Event_Name,
      Date_of_Event: lead.Date_of_Event,
      Active_Quote: lead.Active_Quote
    }));

    // Step 2: Open spreadsheet and contracts sheet
    Logger.log('Step 2: Opening spreadsheet and contracts sheet...');
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const contractsSheet = ss.getSheetByName(SHEETS.CONTRACTS);
    if (!contractsSheet) {
      Logger.log('ERROR: Contracts sheet not found');
      return { success: false, error: 'Contracts sheet not found' };
    }
    Logger.log('Contracts sheet found, last row: ' + contractsSheet.getLastRow());

    // Step 3: Generate invoice number
    const timestamp = new Date();
    const invoiceNumber = 'INV' + Utilities.formatDate(timestamp, 'America/New_York', 'yyMMddHHmmssSSS');
    Logger.log('Step 3: Invoice number generated: ' + invoiceNumber);

    // Step 4: Get images for PDF
    Logger.log('Step 4: Getting images...');
    Logger.log('Header Image ID: ' + CONTRACT_CONFIG.HEADER_IMAGE_ID);
    Logger.log('Signature Image ID: ' + CONTRACT_CONFIG.SIGNATURE_IMAGE_ID);
    const headerImageData = getImageAsBase64(CONTRACT_CONFIG.HEADER_IMAGE_ID);
    const signatureImageData = getImageAsBase64(CONTRACT_CONFIG.SIGNATURE_IMAGE_ID);
    Logger.log('Images loaded - Header: ' + (headerImageData ? 'OK' : 'FAILED') + ', Signature: ' + (signatureImageData ? 'OK' : 'FAILED'));
    
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
    
    // Step 5: Build contract data
    Logger.log('Step 5: Building contract data...');
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
    Logger.log('Filename: ' + filename);

    // Step 6: Generate HTML and PDF
    Logger.log('Step 6: Generating HTML and PDF...');
    const html = generateContractHTML(contractData);
    Logger.log('HTML generated, length: ' + html.length + ' chars');

    const blob = Utilities.newBlob(html, 'text/html', filename + '.html');
    Logger.log('HTML blob created');

    const pdf = blob.getAs('application/pdf').setName(filename + '.pdf');
    Logger.log('PDF created: ' + filename + '.pdf');

    // Step 7: Save to Drive
    Logger.log('Step 7: Saving to Google Drive...');
    Logger.log('Folder ID: ' + CONTRACT_CONFIG.CONTRACTS_FOLDER_ID);
    const folder = DriveApp.getFolderById(CONTRACT_CONFIG.CONTRACTS_FOLDER_ID);
    Logger.log('Folder found: ' + folder.getName());

    const file = folder.createFile(pdf);
    Logger.log('File created in Drive - ID: ' + file.getId() + ', URL: ' + file.getUrl());

    // Step 8: Log to Contracts sheet
    Logger.log('Step 8: Logging to Contracts sheet...');
    const contractId = 'CON-' + String(contractsSheet.getLastRow()).padStart(5, '0');
    Logger.log('Contract ID: ' + contractId);

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
    Logger.log('Contract row data: ' + JSON.stringify(contractRow));
    contractsSheet.appendRow(contractRow);
    Logger.log('Row appended to Contracts sheet');

    // Step 9: Update lead with Contract ID
    Logger.log('Step 9: Updating lead with Contract ID...');
    const leadsSheet = ss.getSheetByName(SHEETS.LEADS);
    const headers = leadsSheet.getRange(1, 1, 1, leadsSheet.getLastColumn()).getValues()[0];
    const contractIdCol = headers.indexOf('Contract_ID') + 1;
    if (contractIdCol > 0) {
      leadsSheet.getRange(rowIndex, contractIdCol).setValue(contractId);
      Logger.log('Lead updated with Contract ID at column ' + contractIdCol);
    } else {
      Logger.log('WARNING: Contract_ID column not found in Leads sheet');
    }

    // Step 9B: Update lead status to "Contract Lead"
    Logger.log('Step 9B: Updating lead status to Contract Lead...');
    const statusCol = headers.indexOf('Lead_Status') + 1;
    if (statusCol > 0) {
      leadsSheet.getRange(rowIndex, statusCol).setValue('Contract Lead');
      Logger.log('Lead status updated to Contract Lead at column ' + statusCol);
    } else {
      Logger.log('WARNING: Lead_Status column not found in Leads sheet');
    }

    // Step 9C: Update Date_Last_Touch
    Logger.log('Step 9C: Updating Date_Last_Touch...');
    const lastTouchCol = headers.indexOf('Date_Last_Touch') + 1;
    if (lastTouchCol > 0) {
      leadsSheet.getRange(rowIndex, lastTouchCol).setValue(new Date());
      Logger.log('Date_Last_Touch updated at column ' + lastTouchCol);
    } else {
      Logger.log('WARNING: Date_Last_Touch column not found in Leads sheet');
    }

    // Step 10: Track metrics
    Logger.log('Step 10: Tracking metrics...');
    incrementMetric(currentUser, 'Contracts_Generated', 1);

    // Step 11: Log activity
    Logger.log('Step 11: Logging activity...');
    const user = getUserWithId(currentUser);
    const userId = user ? user.User_ID : '';
    logActivity({
      leadId: lead.Record_ID,
      userId: userId,
      userName: currentUser,
      type: 'Contract',
      title: 'Contract Generated',
      description: filename + '.pdf',
      relatedId: contractId
    });

    Logger.log('=== CONTRACT GENERATION COMPLETE ===');
    Logger.log('Contract ID: ' + contractId);
    Logger.log('File URL: ' + file.getUrl());

    return {
      success: true,
      contractId: contractId,
      invoiceNumber: invoiceNumber,
      filename: filename + '.pdf',
      url: file.getUrl(),
      message: 'Contract generated successfully',
      statusUpdated: true,
      newStatus: 'Contract Lead'
    };

  } catch (error) {
    Logger.log('=== CONTRACT GENERATION ERROR ===');
    Logger.log('Error: ' + error.toString());
    Logger.log('Error stack: ' + (error.stack || 'No stack trace'));
    console.error('Contract generation error:', error);
    return {
      success: false,
      error: error.toString(),
      stack: error.stack || 'No stack trace'
    };
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
  <div class="client-signature-section" style="margin-top: 30px; padding-top: 20px; border-top: 2px solid #E5E7EB;">\
    <p style="font-weight: bold; font-size: 10pt; margin-bottom: 20px;">Client Signature:</p>\
    <p style="margin-top: 30px; font-size: 11pt;">X _____________________________________________________________</p>\
    <div style="margin-top: 15px; display: table; width: 100%;">\
      <div style="display: table-cell; width: 65%;">\
        <p style="font-size: 8pt; color: #666;">Client Name: _________________________________</p>\
      </div>\
      <div style="display: table-cell; width: 35%;">\
        <p style="font-size: 8pt; color: #666;">Date: ______________</p>\
      </div>\
    </div>\
  </div>\
  <div class="dewayne-signature-closing" style="margin-top: 50px; padding-top: 30px; border-top: 1px solid #eee; text-align: center;">\
    <img src="' + data.signatureImage + '" style="height: 70px; margin-bottom: 8px;" alt="Dewayne Hill Signature">\
    <p style="font-weight: bold; font-size: 10pt; margin: 5px 0;">Dewayne Hill</p>\
    <p style="font-size: 8pt; color: #666; margin: 0;">America\'s Funniest Comedy Magician</p>\
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

  // Get headers to map columns dynamically
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i);

  // Build row array based on headers
  const rowData = new Array(headers.length).fill('');

  // Set values using column mapping
  if (colIndex['Research_ID'] !== undefined) rowData[colIndex['Research_ID']] = newId;
  if (colIndex['Created_Date'] !== undefined) rowData[colIndex['Created_Date']] = new Date();
  if (colIndex['Created_By'] !== undefined) rowData[colIndex['Created_By']] = currentUser;
  if (colIndex['Looking_For'] !== undefined) rowData[colIndex['Looking_For']] = data.Looking_For || '';
  if (colIndex['City_Location'] !== undefined) rowData[colIndex['City_Location']] = data.City_Location || '';
  if (colIndex['Event_Type'] !== undefined) rowData[colIndex['Event_Type']] = data.Event_Type || '';
  if (colIndex['Date_Needed'] !== undefined) rowData[colIndex['Date_Needed']] = data.Date_Needed || '';
  if (colIndex['Guest_Count'] !== undefined) rowData[colIndex['Guest_Count']] = data.Guest_Count || '';
  if (colIndex['Message'] !== undefined) rowData[colIndex['Message']] = data.Message || '';
  if (colIndex['Budget_Range'] !== undefined) rowData[colIndex['Budget_Range']] = data.Budget_Range || '';
  if (colIndex['Research_Status'] !== undefined) rowData[colIndex['Research_Status']] = 'Pending';

  // Add Bark fields
  if (colIndex['Bark_Name'] !== undefined) rowData[colIndex['Bark_Name']] = data.Bark_Name || '';
  if (colIndex['Bark_Email_Partial'] !== undefined) rowData[colIndex['Bark_Email_Partial']] = data.Bark_Email_Partial || '';
  if (colIndex['Bark_Phone_Partial'] !== undefined) rowData[colIndex['Bark_Phone_Partial']] = data.Bark_Phone_Partial || '';

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

/**
 * Fix #10, #11, #12: Submit daily check-in with new fields and email to Dewayne
 */
function submitCheckin(data, currentUser) {
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

  // Find existing row for today
  for (var i = allData.length - 1; i >= 1; i--) {
    var rowDate = new Date(allData[i][0]);
    rowDate.setHours(0, 0, 0, 0);

    if (rowDate.getTime() === today.getTime() && allData[i][1] === currentUser) {
      rowIndex = i + 1;
      break;
    }
  }

  // Map new field names to column names
  var fieldMapping = {
    'callScore': 'Call_Score_Self',
    'cardoneTraining': 'Cardone_Training',
    'roleplayTime': 'Roleplay_Time',
    'barkLeadsFound': 'Bark_Leads_Found',
    'closePhrases': 'Close_Phrases_Used',
    'roadblocks': 'Roadblocks',
    'wins': 'Special_Notes'
  };

  if (rowIndex === -1) {
    // Create new row
    var newRow = [];
    for (var j = 0; j < headers.length; j++) {
      var h = headers[j];
      if (h === 'Date') newRow.push(today);
      else if (h === 'User') newRow.push(currentUser);
      else if (h === 'Submitted_Time') newRow.push(new Date());
      else {
        // Check if this header maps to our data
        var found = false;
        for (var key in fieldMapping) {
          if (fieldMapping[key] === h && data[key] !== undefined) {
            newRow.push(data[key]);
            found = true;
            break;
          }
        }
        if (!found) newRow.push('');
      }
    }
    sheet.appendRow(newRow);
  } else {
    // Update existing row
    for (var field in fieldMapping) {
      var colName = fieldMapping[field];
      if (data[field] !== undefined && colIndex[colName]) {
        sheet.getRange(rowIndex, colIndex[colName]).setValue(data[field]);
      }
    }
    if (colIndex['Submitted_Time']) {
      sheet.getRange(rowIndex, colIndex['Submitted_Time']).setValue(new Date());
    }
  }

  // Fix #12: Send email to Dewayne
  try {
    sendCheckinEmailToDewayne(data, currentUser);
  } catch (e) {
    Logger.log('Error sending email: ' + e.toString());
    // Don't fail the whole submission if email fails
  }

  return { success: true, message: 'Check-in submitted!' };
}

/**
 * Fix #12: Send professional HTML EOD email breakdown to Dewayne
 */
function sendCheckinEmailToDewayne(data, workerName) {
  var recipient = 'dhillmagic@gmail.com';
  var date = Utilities.formatDate(new Date(), 'America/New_York', 'EEEE, MMMM d, yyyy');
  var shortDate = Utilities.formatDate(new Date(), 'America/New_York', 'MM/dd');
  var subject = 'EOD Report | ' + workerName + ' | ' + shortDate;

  // Get today's activity metrics for this user
  var metrics = getTodayMetrics(workerName);

  // Calculate score color
  var scoreVal = parseInt(data.callScore) || 0;
  var scoreColor = scoreVal <= 5 ? '#ef4444' : scoreVal <= 10 ? '#eab308' : '#22c55e';
  var scoreLabel = scoreVal <= 5 ? 'Needs Improvement' : scoreVal <= 10 ? 'Average' : 'Excellent';

  // Build HTML email
  var htmlBody = '<!DOCTYPE html>' +
    '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>' +
    '<body style="margin:0;padding:0;background:#f4f4f5;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;">' +
    '<div style="max-width:600px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 6px rgba(0,0,0,0.1);">' +

    // Header
    '<div style="background:linear-gradient(135deg,#6366f1,#8b5cf6);padding:32px 24px;text-align:center;">' +
    '<h1 style="margin:0;color:#fff;font-size:24px;font-weight:700;">End of Day Report</h1>' +
    '<p style="margin:8px 0 0;color:rgba(255,255,255,0.9);font-size:14px;">' + date + '</p>' +
    '</div>' +

    // Worker Info
    '<div style="padding:24px;border-bottom:1px solid #e5e7eb;">' +
    '<div style="display:flex;align-items:center;">' +
    '<div style="width:48px;height:48px;background:linear-gradient(135deg,#6366f1,#a855f7);border-radius:12px;display:inline-flex;align-items:center;justify-content:center;margin-right:16px;">' +
    '<span style="color:#fff;font-size:18px;font-weight:700;">' + workerName.charAt(0) + '</span>' +
    '</div>' +
    '<div><div style="font-size:18px;font-weight:600;color:#111;">' + workerName + '</div>' +
    '<div style="font-size:13px;color:#6b7280;">Daily Check-in Submitted</div></div></div>' +
    '</div>' +

    // Performance Metrics
    '<div style="padding:24px;border-bottom:1px solid #e5e7eb;">' +
    '<h2 style="margin:0 0 16px;font-size:14px;color:#6b7280;text-transform:uppercase;letter-spacing:0.5px;">Today\'s Performance</h2>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr>' +
    '<td style="padding:12px;text-align:center;background:#f9fafb;border-radius:8px 0 0 8px;">' +
    '<div style="font-size:28px;font-weight:700;color:#6366f1;">' + (metrics.Calls_Made || 0) + '</div>' +
    '<div style="font-size:11px;color:#6b7280;text-transform:uppercase;">Calls</div></td>' +
    '<td style="padding:12px;text-align:center;background:#f9fafb;">' +
    '<div style="font-size:28px;font-weight:700;color:#22c55e;">' + (metrics.Conversations || 0) + '</div>' +
    '<div style="font-size:11px;color:#6b7280;text-transform:uppercase;">Convos</div></td>' +
    '<td style="padding:12px;text-align:center;background:#f9fafb;">' +
    '<div style="font-size:28px;font-weight:700;color:#eab308;">' + (metrics.Followups_Sent || 0) + '</div>' +
    '<div style="font-size:11px;color:#6b7280;text-transform:uppercase;">Follow-ups</div></td>' +
    '<td style="padding:12px;text-align:center;background:#f9fafb;">' +
    '<div style="font-size:28px;font-weight:700;color:#f97316;">' + (metrics.Leads_Upgraded || 0) + '</div>' +
    '<div style="font-size:11px;color:#6b7280;text-transform:uppercase;">Upgrades</div></td>' +
    '<td style="padding:12px;text-align:center;background:#f9fafb;border-radius:0 8px 8px 0;">' +
    '<div style="font-size:28px;font-weight:700;color:#8b5cf6;">' + (metrics.Contracts_Generated || 0) + '</div>' +
    '<div style="font-size:11px;color:#6b7280;text-transform:uppercase;">Contracts</div></td>' +
    '</tr></table></div>' +

    // Self Assessment Score
    '<div style="padding:24px;border-bottom:1px solid #e5e7eb;text-align:center;">' +
    '<h2 style="margin:0 0 16px;font-size:14px;color:#6b7280;text-transform:uppercase;letter-spacing:0.5px;">Self Assessment</h2>' +
    '<div style="display:inline-block;padding:20px 40px;background:#f9fafb;border-radius:16px;">' +
    '<div style="font-size:48px;font-weight:700;color:' + scoreColor + ';">' + (data.callScore || '-') + '<span style="font-size:24px;color:#9ca3af;">/15</span></div>' +
    '<div style="font-size:13px;color:#6b7280;margin-top:4px;">' + scoreLabel + '</div></div>' +
    '</div>' +

    // Training Section
    '<div style="padding:24px;border-bottom:1px solid #e5e7eb;">' +
    '<h2 style="margin:0 0 16px;font-size:14px;color:#6b7280;text-transform:uppercase;letter-spacing:0.5px;">Training & Development</h2>' +
    '<table style="width:100%;">' +
    '<tr><td style="padding:8px 0;color:#374151;">Cardone U Training</td><td style="padding:8px 0;text-align:right;font-weight:600;color:#111;">' + (data.cardoneTraining || '0') + ' min</td></tr>' +
    '<tr><td style="padding:8px 0;color:#374151;">Role Play Practice</td><td style="padding:8px 0;text-align:right;font-weight:600;color:#111;">' + (data.roleplayTime || '0') + ' min</td></tr>' +
    '<tr><td style="padding:8px 0;color:#374151;">Bark Leads Found</td><td style="padding:8px 0;text-align:right;font-weight:600;color:#111;">' + (data.barkLeadsFound || '0') + '</td></tr>' +
    '</table></div>' +

    // Notes Section
    '<div style="padding:24px;">' +
    '<h2 style="margin:0 0 16px;font-size:14px;color:#6b7280;text-transform:uppercase;letter-spacing:0.5px;">Daily Notes</h2>' +

    (data.closePhrases ? '<div style="margin-bottom:16px;"><div style="font-size:12px;color:#6b7280;margin-bottom:4px;">Closing Phrases Used</div>' +
    '<div style="padding:12px;background:#f0fdf4;border-radius:8px;border-left:3px solid #22c55e;color:#166534;">' + data.closePhrases + '</div></div>' : '') +

    (data.roadblocks ? '<div style="margin-bottom:16px;"><div style="font-size:12px;color:#6b7280;margin-bottom:4px;">Roadblocks</div>' +
    '<div style="padding:12px;background:#fef2f2;border-radius:8px;border-left:3px solid #ef4444;color:#991b1b;">' + data.roadblocks + '</div></div>' : '') +

    (data.wins ? '<div><div style="font-size:12px;color:#6b7280;margin-bottom:4px;">Wins & Notable Leads</div>' +
    '<div style="padding:12px;background:#eff6ff;border-radius:8px;border-left:3px solid #3b82f6;color:#1e40af;">' + data.wins + '</div></div>' : '') +

    (!data.closePhrases && !data.roadblocks && !data.wins ? '<div style="padding:16px;background:#f9fafb;border-radius:8px;text-align:center;color:#6b7280;">No additional notes provided</div>' : '') +

    '</div>' +

    // Footer
    '<div style="padding:16px 24px;background:#f9fafb;text-align:center;">' +
    '<p style="margin:0;font-size:12px;color:#9ca3af;">Sent from BEM Lead Management System</p>' +
    '</div>' +

    '</div></body></html>';

  // Plain text fallback
  var plainBody = 'EOD Report - ' + workerName + ' - ' + date + '\n\n' +
    'PERFORMANCE: Calls: ' + (metrics.Calls_Made || 0) + ' | Convos: ' + (metrics.Conversations || 0) +
    ' | Follow-ups: ' + (metrics.Followups_Sent || 0) + ' | Upgrades: ' + (metrics.Leads_Upgraded || 0) +
    ' | Contracts: ' + (metrics.Contracts_Generated || 0) + '\n\n' +
    'SELF SCORE: ' + (data.callScore || '-') + '/15\n\n' +
    'TRAINING: Cardone U: ' + (data.cardoneTraining || '0') + ' min | Role Play: ' + (data.roleplayTime || '0') + ' min\n' +
    'Bark Leads: ' + (data.barkLeadsFound || 0) + '\n\n' +
    'NOTES:\n' +
    'Phrases: ' + (data.closePhrases || 'None') + '\n' +
    'Roadblocks: ' + (data.roadblocks || 'None') + '\n' +
    'Wins: ' + (data.wins || 'None');

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody
  });
}

/**
 * Get weekly score average for user
 */
function getWeeklyScoreAverage(userName) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEETS.EOD);

  if (!sheet) return 'N/A';

  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var colIndex = {};
  headers.forEach(function(h, i) { colIndex[h] = i; });

  var today = new Date();
  var weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);

  var scores = [];
  for (var i = 1; i < data.length; i++) {
    var rowDate = new Date(data[i][colIndex['Date'] || 0]);
    var user = data[i][colIndex['User'] || 1];
    var score = data[i][colIndex['Call_Score_Self']];

    if (user === userName && rowDate >= weekAgo && score) {
      var numScore = parseInt(score);
      if (!isNaN(numScore)) {
        scores.push(numScore);
      }
    }
  }

  if (scores.length === 0) return 'N/A';

  var sum = scores.reduce(function(a, b) { return a + b; }, 0);
  return (sum / scores.length).toFixed(1);
}

function getTeamPerformance(period, dateFrom, dateTo) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var eodSheet = ss.getSheetByName(SHEETS.EOD);
  var leadsSheet = ss.getSheetByName(SHEETS.LEADS);

  // Determine date range
  var now = new Date();
  var startDate = new Date();
  var endDate = new Date();

  if (period === 'custom' && dateFrom && dateTo) {
    startDate = new Date(dateFrom);
    endDate = new Date(dateTo);
    endDate.setHours(23, 59, 59, 999);
  } else if (period === 'today') {
    startDate.setHours(0, 0, 0, 0);
    endDate = now;
  } else if (period === 'week') {
    startDate.setDate(now.getDate() - 7);
    startDate.setHours(0, 0, 0, 0);
    endDate = now;
  } else if (period === 'month') {
    startDate.setDate(now.getDate() - 30);
    startDate.setHours(0, 0, 0, 0);
    endDate = now;
  } else {
    startDate.setHours(0, 0, 0, 0);
    endDate = now;
  }

  var userStats = {};

  // Get activity data from EOD sheet
  if (eodSheet) {
    var eodData = eodSheet.getDataRange().getValues();
    var eodHeaders = eodData[0];
    var eodColIndex = {};
    eodHeaders.forEach(function(h, i) { eodColIndex[h] = i; });

    for (var i = 1; i < eodData.length; i++) {
      var row = eodData[i];
      var rowDate = new Date(row[eodColIndex['Date']]);

      if (rowDate >= startDate && rowDate <= endDate) {
        var user = row[eodColIndex['User']];
        if (!user) continue;

        if (!userStats[user]) {
          userStats[user] = {
            name: user,
            calls: 0,
            conversations: 0,
            emails: 0,
            upgrades: 0,
            deals: 0,
            bookedValue: 0,
            pipelineValue: 0,
            score: 0
          };
        }

        userStats[user].calls += Number(row[eodColIndex['Calls_Made']]) || 0;
        userStats[user].conversations += Number(row[eodColIndex['Conversations']]) || 0;
        userStats[user].emails += Number(row[eodColIndex['Followups_Sent']]) || 0;
        userStats[user].upgrades += Number(row[eodColIndex['Leads_Upgraded']]) || 0;
      }
    }
  }

  // Get deals, booked value, and pipeline from Leads sheet
  if (leadsSheet) {
    var leadsData = leadsSheet.getDataRange().getValues();
    var leadHeaders = leadsData[0];
    var leadColIndex = {};
    leadHeaders.forEach(function(h, i) { leadColIndex[h] = i; });

    for (var j = 1; j < leadsData.length; j++) {
      var lead = leadsData[j];
      var owner = lead[leadColIndex['Lead_Owner']];
      var location = lead[leadColIndex['Location']];
      var bookedDate = lead[leadColIndex['Date_Booked']];
      var quote = parseFloat(String(lead[leadColIndex['Active_Quote']] || '0').replace(/[^0-9.-]/g, '')) || 0;

      if (!owner) continue;

      // Initialize user if not exists
      if (!userStats[owner]) {
        userStats[owner] = {
          name: owner,
          calls: 0,
          conversations: 0,
          emails: 0,
          upgrades: 0,
          deals: 0,
          bookedValue: 0,
          pipelineValue: 0,
          score: 0
        };
      }

      // Count booked deals in period
      if (location === 'Booked' && bookedDate instanceof Date && bookedDate >= startDate && bookedDate <= endDate) {
        userStats[owner].deals++;
        userStats[owner].bookedValue += quote;
      }

      // Pipeline: leads not yet booked (Hot, Warm, Follow-up, etc.)
      if (location !== 'Booked' && location !== 'Lost' && location !== 'Parking Lot' && location !== 'Storage') {
        userStats[owner].pipelineValue += quote;
      }
    }
  }

  // Get AI Quality Scores from Dialpad_Calls sheet
  var callsSheet = ss.getSheetByName('Dialpad_Calls');
  var userAiScores = {};

  if (callsSheet) {
    var callsData = callsSheet.getDataRange().getValues();
    var callsHeaders = callsData[0];
    var callsColIdx = {};
    callsHeaders.forEach(function(h, i) { callsColIdx[h] = i; });

    for (var k = 1; k < callsData.length; k++) {
      var callRow = callsData[k];
      var callDate = new Date(callRow[callsColIdx['Date']]);

      // Only include calls within the date range
      if (callDate >= startDate && callDate <= endDate) {
        var callUser = callRow[callsColIdx['User']] || '';
        var aiScoreRaw = callRow[callsColIdx['AI_Score_Pct']] || '';

        // Parse AI score (could be "85%" or just "85")
        var aiScoreNum = parseFloat(String(aiScoreRaw).replace('%', '')) || 0;
        if (aiScoreNum > 0 && aiScoreNum < 1) aiScoreNum = aiScoreNum * 100;

        if (callUser && aiScoreNum > 0) {
          if (!userAiScores[callUser]) userAiScores[callUser] = [];
          userAiScores[callUser].push(aiScoreNum);
        }
      }
    }
  }

  // Get user roles to filter out admins from team performance
  var usersSheet = ss.getSheetByName(SHEETS.USERS);
  var adminUsers = {};
  if (usersSheet) {
    var usersData = usersSheet.getDataRange().getValues();
    var usersHeaders = usersData[0];
    var userColIdx = {};
    usersHeaders.forEach(function(h, i) { userColIdx[h] = i; });

    for (var m = 1; m < usersData.length; m++) {
      var userRow = usersData[m];
      var uName = userRow[userColIdx['Name']] || '';
      var uRole = userRow[userColIdx['Role']] || '';
      var showInDash = userRow[userColIdx['Show_In_Dashboard']];

      // Mark user as admin if Role is Admin OR Show_In_Dashboard is 'No'
      if (uRole === 'Admin' || showInDash === 'No' || showInDash === false) {
        adminUsers[uName] = true;
      }
    }
  }

  // Calculate score and quality score, then build result array
  var result = [];
  for (var userName in userStats) {
    // Skip admin users - they should not appear in team performance
    if (adminUsers[userName]) {
      continue;
    }

    var u = userStats[userName];
    // Score formula: calls*1 + conversations*3 + emails*1 + upgrades*5 + deals*20 + bookedValue/100
    u.score = Math.round(
      (u.calls * 1) +
      (u.conversations * 3) +
      (u.emails * 1) +
      (u.upgrades * 5) +
      (u.deals * 20) +
      (u.bookedValue / 100)
    );

    // Quality Score: Pull from OpenAI grading in Dialpad_Calls
    // Calculate average AI score for this user
    if (userAiScores[userName] && userAiScores[userName].length > 0) {
      var scores = userAiScores[userName];
      var sum = scores.reduce(function(a, b) { return a + b; }, 0);
      u.qualityScore = Math.round(sum / scores.length);
    } else {
      // No AI scores available - return null to show "N/A"
      u.qualityScore = null;
    }

    result.push(u);
  }

  // Sort by score descending
  result.sort(function(a, b) { return b.score - a.score; });

  return result;
}

// ============================================
// UPCOMING SHOWS (Dashboard)
// ============================================

function getUpcomingShows(currentUser) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEETS.LEADS);
  var data = sheet.getDataRange().getValues();

  if (data.length < 2) return [];

  var headers = data[0];
  var colIndex = {};
  headers.forEach(function(h, i) { colIndex[h] = i; });

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var thirtyDaysLater = new Date(today);
  thirtyDaysLater.setDate(thirtyDaysLater.getDate() + 30);

  Logger.log('=== UPCOMING SHOWS DEBUG ===');
  Logger.log('Today: ' + today);
  Logger.log('30 days later: ' + thirtyDaysLater);

  var shows = [];
  var bookedCount = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var location = row[colIndex['Location']];
    var eventDateRaw = row[colIndex['Date_of_Event']];
    var eventName = row[colIndex['Event_Name']] || 'Unknown';

    // Only process Booked leads
    if (location !== 'Booked') continue;

    bookedCount++;
    Logger.log('--- Booked Lead #' + bookedCount + ': ' + eventName + ' ---');
    Logger.log('  Date_of_Event raw value: ' + eventDateRaw);
    Logger.log('  Type: ' + typeof eventDateRaw);
    Logger.log('  Is Date object: ' + (eventDateRaw instanceof Date));

    if (!eventDateRaw) {
      Logger.log('  SKIPPED: No event date');
      continue;
    }

    // Parse date - handle multiple formats
    var ed = parseFlexibleDate(eventDateRaw);

    Logger.log('  Parsed date: ' + ed);
    Logger.log('  Is valid: ' + (ed && !isNaN(ed.getTime())));

    // Skip if we couldn't parse a valid date
    if (!ed || isNaN(ed.getTime())) {
      Logger.log('  SKIPPED: Could not parse date');
      continue;
    }

    ed.setHours(0, 0, 0, 0);

    var inRange = (ed >= today && ed <= thirtyDaysLater);
    Logger.log('  Date in range (today to +30 days): ' + inRange);
    Logger.log('  ed >= today: ' + (ed >= today) + ' | ed <= thirtyDaysLater: ' + (ed <= thirtyDaysLater));

    if (inRange) {
      shows.push({
        eventDate: formatDate(ed),
        eventTimestamp: ed.getTime(),
        Event_Name: eventName,
        Company_Name: row[colIndex['Company_Name']] || '',
        Venue_City: row[colIndex['Venue_City']] || '',
        Company_City: row[colIndex['Company_City']] || '',
        Active_Quote: row[colIndex['Active_Quote']] || 0
      });
      Logger.log('  ADDED to upcoming shows!');
    }
  }

  Logger.log('=== SUMMARY ===');
  Logger.log('Total booked leads found: ' + bookedCount);
  Logger.log('Upcoming shows (in next 30 days): ' + shows.length);

  shows.sort(function(a, b) {
    return a.eventTimestamp - b.eventTimestamp;
  });

  Logger.log('Returning shows: ' + JSON.stringify(shows));
  return shows;
}

// Helper function to parse dates - DD/MM/YYYY format
function parseFlexibleDate(dateValue) {
  if (!dateValue) return null;

  Logger.log('  parseFlexibleDate input: ' + dateValue + ' (type: ' + typeof dateValue + ')');

  // Already a Date object (check multiple ways due to Apps Script quirks)
  if (dateValue instanceof Date) {
    Logger.log('  -> Detected as Date object via instanceof');
    return new Date(dateValue.getTime());
  }

  // Check if it's a Date-like object with getTime method
  if (typeof dateValue === 'object' && dateValue !== null && typeof dateValue.getTime === 'function') {
    Logger.log('  -> Detected as Date-like object with getTime()');
    return new Date(dateValue.getTime());
  }

  // Number (Excel serial date)
  if (typeof dateValue === 'number') {
    Logger.log('  -> Detected as number (Excel serial date)');
    var excelEpoch = new Date(1899, 11, 30);
    var msPerDay = 24 * 60 * 60 * 1000;
    return new Date(excelEpoch.getTime() + dateValue * msPerDay);
  }

  // Convert to string for parsing
  var str = String(dateValue).trim();
  if (str === '' || str === 'undefined' || str === 'null') return null;

  Logger.log('  -> Parsing as string: "' + str + '"');

  // DD/MM/YYYY format (e.g., "19/01/2026")
  if (str.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
    var parts = str.split('/');
    var day = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10) - 1; // JS months are 0-indexed
    var year = parseInt(parts[2], 10);
    Logger.log('  -> Parsed DD/MM/YYYY: day=' + day + ', month=' + (month+1) + ', year=' + year);
    return new Date(year, month, day);
  }

  // YYYY-MM-DD format (ISO)
  if (str.match(/^\d{4}-\d{1,2}-\d{1,2}/)) {
    var parts = str.split(/[-T]/);
    var year = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10) - 1;
    var day = parseInt(parts[2], 10);
    Logger.log('  -> Parsed YYYY-MM-DD: day=' + day + ', month=' + (month+1) + ', year=' + year);
    return new Date(year, month, day);
  }

  // Fallback: let JavaScript try to parse it
  var fallback = new Date(str);
  if (!isNaN(fallback.getTime())) {
    Logger.log('  -> Fallback parsing succeeded: ' + fallback);
    return fallback;
  }

  Logger.log('  -> FAILED to parse date');
  return null;
}

// TEST FUNCTION - Run this directly in Apps Script to debug
function testUpcomingShows() {
  Logger.log('========== TEST UPCOMING SHOWS ==========');
  var result = getUpcomingShows('');
  Logger.log('========== RESULT ==========');
  Logger.log('Shows returned: ' + result.length);
  result.forEach(function(show, i) {
    Logger.log((i+1) + '. ' + show.Event_Name + ' - ' + show.eventDate + ' - $' + show.Active_Quote);
  });
  return result;
}

// ============================================
// BOOKED STATS (Performance Section)
// ============================================

function getBookedStats(period) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var eodSheet = ss.getSheetByName(SHEETS.EOD);
  var leadsSheet = ss.getSheetByName(SHEETS.LEADS);

  var now = new Date();
  var startDate = new Date();

  if (period === 'week') {
    startDate.setDate(now.getDate() - 7);
  } else if (period === 'month') {
    startDate.setDate(now.getDate() - 30);
  } else if (period === 'quarter') {
    startDate.setDate(now.getDate() - 90);
  } else if (period === 'year') {
    startDate.setDate(now.getDate() - 365);
  } else {
    startDate.setHours(0, 0, 0, 0);
  }

  var result = {
    calls: 0,
    conversations: 0,
    emails: 0,
    deals: 0,
    revenue: 0
  };

  // Get calls and conversations from EOD
  if (eodSheet) {
    var eodData = eodSheet.getDataRange().getValues();
    var eodHeaders = eodData[0];
    var eodColIndex = {};
    eodHeaders.forEach(function(h, i) { eodColIndex[h] = i; });

    for (var i = 1; i < eodData.length; i++) {
      var row = eodData[i];
      var rowDate = new Date(row[eodColIndex['Date']]);

      if (rowDate >= startDate) {
        result.calls += Number(row[eodColIndex['Calls_Made']]) || 0;
        result.conversations += Number(row[eodColIndex['Conversations']]) || 0;
        result.emails += Number(row[eodColIndex['Followups_Sent']]) || 0;
      }
    }
  }

  // Get deals and revenue from Leads (booked in period)
  if (leadsSheet) {
    var leadsData = leadsSheet.getDataRange().getValues();
    var leadHeaders = leadsData[0];
    var leadColIndex = {};
    leadHeaders.forEach(function(h, i) { leadColIndex[h] = i; });

    for (var j = 1; j < leadsData.length; j++) {
      var lead = leadsData[j];
      var location = lead[leadColIndex['Location']];
      var bookedDate = lead[leadColIndex['Date_Booked']];
      var quote = lead[leadColIndex['Active_Quote']];

      if (location === 'Booked' && bookedDate instanceof Date && bookedDate >= startDate) {
        result.deals++;
        if (quote) {
          result.revenue += parseFloat(String(quote).replace(/[$,]/g, '')) || 0;
        }
      }
    }
  }

  return result;
}

// ============================================
// PERFORMANCE STATS (Dashboard Performance Section)
// ============================================

function getPerformanceStats(period, dateFrom, dateTo) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var eodSheet = ss.getSheetByName(SHEETS.EOD);
  var leadsSheet = ss.getSheetByName(SHEETS.LEADS);

  var now = new Date();
  var startDate = new Date();
  var endDate = new Date();

  if (period === 'custom' && dateFrom && dateTo) {
    startDate = new Date(dateFrom);
    endDate = new Date(dateTo);
    endDate.setHours(23, 59, 59, 999);
  } else if (period === 'today') {
    startDate.setHours(0, 0, 0, 0);
    endDate = now;
  } else if (period === 'week') {
    startDate.setDate(now.getDate() - 7);
    startDate.setHours(0, 0, 0, 0);
    endDate = now;
  } else if (period === 'month') {
    startDate.setDate(now.getDate() - 30);
    startDate.setHours(0, 0, 0, 0);
    endDate = now;
  } else {
    startDate.setHours(0, 0, 0, 0);
    endDate = now;
  }

  var result = {
    calls: 0,
    conversations: 0,
    emails: 0,
    deals: 0,
    revenue: 0
  };

  // Get calls, conversations, and emails from EOD
  if (eodSheet) {
    var eodData = eodSheet.getDataRange().getValues();
    var eodHeaders = eodData[0];
    var eodColIndex = {};
    eodHeaders.forEach(function(h, i) { eodColIndex[h] = i; });

    for (var i = 1; i < eodData.length; i++) {
      var row = eodData[i];
      var rowDate = new Date(row[eodColIndex['Date']]);

      if (rowDate >= startDate && rowDate <= endDate) {
        result.calls += Number(row[eodColIndex['Calls_Made']]) || 0;
        result.conversations += Number(row[eodColIndex['Conversations']]) || 0;
        result.emails += Number(row[eodColIndex['Followups_Sent']]) || 0;
      }
    }
  }

  // Get deals and revenue from Leads (booked in period)
  if (leadsSheet) {
    var leadsData = leadsSheet.getDataRange().getValues();
    var leadHeaders = leadsData[0];
    var leadColIndex = {};
    leadHeaders.forEach(function(h, i) { leadColIndex[h] = i; });

    for (var j = 1; j < leadsData.length; j++) {
      var lead = leadsData[j];
      var location = lead[leadColIndex['Location']];
      var bookedDate = lead[leadColIndex['Date_Booked']];
      var quote = lead[leadColIndex['Active_Quote']];

      if (location === 'Booked' && bookedDate instanceof Date && bookedDate >= startDate && bookedDate <= endDate) {
        result.deals++;
        if (quote) {
          result.revenue += parseFloat(String(quote).replace(/[$,]/g, '')) || 0;
        }
      }
    }
  }

  return result;
}

// ============================================
// DAILY TRENDS (Dashboard Chart)
// ============================================

function getDailyTrends(period) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEETS.EOD);

  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIndex = {};
  headers.forEach(function(h, i) { colIndex[h] = i; });

  var now = new Date();
  var daysToShow = 7;

  if (period === 'week') {
    daysToShow = 7;
  } else if (period === 'month') {
    daysToShow = 30;
  } else if (period === 'quarter') {
    daysToShow = 12; // Show weeks instead
  } else if (period === 'year') {
    daysToShow = 12; // Show months
  }

  var dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  var monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  // For quarter/year, aggregate differently
  if (period === 'quarter' || period === 'year') {
    return getAggregatedTrends(data, colIndex, period);
  }

  // Daily trends for week/month
  var trends = [];

  for (var d = daysToShow - 1; d >= 0; d--) {
    var targetDate = new Date(now);
    targetDate.setDate(now.getDate() - d);
    targetDate.setHours(0, 0, 0, 0);

    var dayData = { label: dayNames[targetDate.getDay()], calls: 0, convos: 0 };

    for (var i = 1; i < data.length; i++) {
      var rowDate = new Date(data[i][colIndex['Date']]);
      rowDate.setHours(0, 0, 0, 0);

      if (rowDate.getTime() === targetDate.getTime()) {
        dayData.calls += Number(data[i][colIndex['Calls_Made']]) || 0;
        dayData.convos += Number(data[i][colIndex['Conversations']]) || 0;
      }
    }

    trends.push(dayData);
  }

  return trends;
}

function getAggregatedTrends(data, colIndex, period) {
  var now = new Date();
  var trends = [];
  var monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  if (period === 'quarter') {
    // Last 12 weeks
    for (var w = 11; w >= 0; w--) {
      var weekStart = new Date(now);
      weekStart.setDate(now.getDate() - (w * 7) - now.getDay());
      weekStart.setHours(0, 0, 0, 0);

      var weekEnd = new Date(weekStart);
      weekEnd.setDate(weekStart.getDate() + 7);

      var weekData = { label: 'W' + (12 - w), calls: 0, convos: 0 };

      for (var i = 1; i < data.length; i++) {
        var rowDate = new Date(data[i][colIndex['Date']]);
        if (rowDate >= weekStart && rowDate < weekEnd) {
          weekData.calls += Number(data[i][colIndex['Calls_Made']]) || 0;
          weekData.convos += Number(data[i][colIndex['Conversations']]) || 0;
        }
      }

      trends.push(weekData);
    }
  } else {
    // Last 12 months
    for (var m = 11; m >= 0; m--) {
      var monthStart = new Date(now.getFullYear(), now.getMonth() - m, 1);
      var monthEnd = new Date(now.getFullYear(), now.getMonth() - m + 1, 0);

      var monthData = { label: monthNames[monthStart.getMonth()], calls: 0, convos: 0 };

      for (var j = 1; j < data.length; j++) {
        var rowDate = new Date(data[j][colIndex['Date']]);
        if (rowDate >= monthStart && rowDate <= monthEnd) {
          monthData.calls += Number(data[j][colIndex['Calls_Made']]) || 0;
          monthData.convos += Number(data[j][colIndex['Conversations']]) || 0;
        }
      }

      trends.push(monthData);
    }
  }

  return trends;
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

/**
 * Format date for display - handles timezone issues by formatting on server side
 * This prevents dates from being off by 1 day due to timezone conversion
 */
function formatDateForDisplay(dateValue) {
  if (!dateValue) return '-';

  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  // If it's already a string in YYYY-MM-DD format, parse manually to avoid timezone shift
  if (typeof dateValue === 'string' && dateValue.match(/^\d{4}-\d{2}-\d{2}/)) {
    const parts = dateValue.split('-');
    const year = parts[0];
    const month = parseInt(parts[1]) - 1;
    const day = parseInt(parts[2]);
    return months[month] + ' ' + day + ', ' + year;
  }

  // If it's a Date object from the spreadsheet, use Utilities.formatDate with script timezone
  if (dateValue instanceof Date) {
    if (isNaN(dateValue.getTime())) return '-';
    return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'MMM d, yyyy');
  }

  // Try parsing as date - but be careful with timezone
  try {
    const date = new Date(dateValue);
    if (!isNaN(date.getTime())) {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM d, yyyy');
    }
  } catch (e) {
    // Fall through to return string
  }

  return String(dateValue);
}

/**
 * Normalize value for comparison (handles dates, nulls, empty strings)
 */
function normalizeValueForComparison(val) {
  if (val === null || val === undefined) return '';
  if (val instanceof Date) {
    return formatDate(val);
  }
  return String(val).trim();
}

/**
 * Format field name for display (replace underscores with spaces)
 */
function formatFieldName(fieldName) {
  if (!fieldName) return '';
  return fieldName.replace(/_/g, ' ');
}

/**
 * Truncate value for display (max 30 chars)
 */
function truncateValue(val) {
  if (!val) return '';
  const str = String(val);
  if (str.length > 30) {
    return str.substring(0, 27) + '...';
  }
  return str;
}

// ============================================
// PHONE NORMALIZATION HELPER (N2)
// ============================================

function normalizePhone(phone) {
  if (!phone) return '';
  // Remove all non-numeric characters
  return String(phone).replace(/[^0-9]/g, '');
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
    // Normalize phone number - ensure it has country code
    var normalizedPhone = normalizePhoneForDialpad(phoneNumber);

    // Validate phone number
    if (!normalizedPhone || normalizedPhone.length < 10) {
      return { success: false, error: 'Invalid phone number' };
    }

    Logger.log('Dialpad Click-to-Call Request:');
    Logger.log('  User ID: ' + userId);
    Logger.log('  Original Phone: ' + phoneNumber);
    Logger.log('  Normalized Phone: ' + normalizedPhone);

    var url = DIALPAD_BASE_URL + '/users/' + userId + '/initiate_call';

    var options = {
      'method': 'post',
      'headers': {
        'Authorization': 'Bearer ' + DIALPAD_API_KEY,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify({
        'phone_number': normalizedPhone
      }),
      'muteHttpExceptions': true
    };

    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    Logger.log('Dialpad API Response:');
    Logger.log('  Status Code: ' + responseCode);
    Logger.log('  Response: ' + responseText);

    // Parse response
    var result;
    try {
      result = JSON.parse(responseText);
    } catch (parseErr) {
      result = { raw: responseText };
    }

    // Check for HTTP errors
    if (responseCode < 200 || responseCode >= 300) {
      var errorMsg = extractErrorMessage(result, 'HTTP ' + responseCode);
      Logger.log('Dialpad API Error: ' + errorMsg);
      return {
        success: false,
        error: errorMsg,
        statusCode: responseCode
      };
    }

    // Check for error in response body
    if (result.error || result.errors) {
      var errorMsg = extractErrorMessage(result, 'Unknown error');
      Logger.log('Dialpad returned error: ' + errorMsg);
      return {
        success: false,
        error: errorMsg
      };
    }

    Logger.log('Dialpad call initiated successfully');
    return { success: true, data: result };

  } catch (e) {
    Logger.log('Exception initiating Dialpad call: ' + e.toString());
    return { success: false, error: e.message || e.toString() };
  }
}

/**
 * Extract readable error message from API response
 */
function extractErrorMessage(response, fallback) {
  if (!response) return fallback;

  // If it's a string, return it directly
  if (typeof response === 'string') return response;

  // Try common error fields
  if (response.error) {
    if (typeof response.error === 'string') return response.error;
    if (response.error.message) return response.error.message;
    if (response.error.description) return response.error.description;
  }

  if (response.message && typeof response.message === 'string') return response.message;
  if (response.detail && typeof response.detail === 'string') return response.detail;
  if (response.errors && Array.isArray(response.errors) && response.errors.length > 0) {
    var firstError = response.errors[0];
    if (typeof firstError === 'string') return firstError;
    if (firstError.message) return firstError.message;
  }

  // Last resort - stringify but keep it short
  try {
    var str = JSON.stringify(response);
    return str.length > 100 ? str.substring(0, 100) + '...' : str;
  } catch (e) {
    return fallback;
  }
}

/**
 * Normalize phone number for Dialpad API
 * Ensures proper format with country code
 *
 * Examples:
 * "8135551234" → "+18135551234"
 * "1-813-555-1234" → "+18135551234"
 * "(813) 555-1234" → "+18135551234"
 * "18135551234" → "+18135551234"
 * "+971552244072" → "+971552244072" (unchanged)
 * "971552244072" → "+971552244072"
 */
function normalizePhoneForDialpad(phone) {
  // Convert to string and handle null/undefined/numbers
  if (!phone && phone !== 0) return '';
  phone = String(phone).trim();
  if (!phone) return '';

  // If already starts with +, just clean it (keep + and digits only)
  if (phone.charAt(0) === '+') {
    return '+' + phone.replace(/\D/g, '');
  }

  // Strip all non-digit characters (spaces, dashes, parentheses, dots)
  var digits = phone.replace(/\D/g, '');

  // Handle empty or too short
  if (!digits || digits.length < 7) {
    return digits; // Return as-is, let validation catch it
  }

  // 10 digits = US number without country code → add +1
  if (digits.length === 10) {
    return '+1' + digits;
  }

  // 11 digits starting with 1 = US number with country code → add +
  if (digits.length === 11 && digits.charAt(0) === '1') {
    return '+' + digits;
  }

  // Any other length > 10 = international number → add +
  if (digits.length > 10) {
    return '+' + digits;
  }

  // Less than 10 digits - return with + prefix anyway (might be short code or partial)
  return '+' + digits;
}

/**
 * Comprehensive Dialpad API debugging function
 * Run this from Apps Script editor to diagnose connection issues
 */
function testDialpadConnection() {
  Logger.log('========== DIALPAD API DIAGNOSTIC ==========');
  Logger.log('');

  // 1. Show current configuration
  Logger.log('1. CURRENT CONFIGURATION:');
  Logger.log('   API Key: ' + DIALPAD_API_KEY.substring(0, 10) + '...' + DIALPAD_API_KEY.substring(DIALPAD_API_KEY.length - 10));
  Logger.log('   Base URL: ' + DIALPAD_BASE_URL);
  Logger.log('');

  Logger.log('   User ID Mappings:');
  Logger.log('   - Dewayne: 5382638134673408');
  Logger.log('   - Christine: 5825830596165632');
  Logger.log('   - Jan: 5183185007984640');
  Logger.log('   - Dewayne 2: 5966784611270656');
  Logger.log('');

  // 2. Test API Key - Get company info
  Logger.log('2. TESTING API KEY (GET /company):');
  try {
    var companyUrl = DIALPAD_BASE_URL + '/company';
    var companyResponse = UrlFetchApp.fetch(companyUrl, {
      'method': 'get',
      'headers': { 'Authorization': 'Bearer ' + DIALPAD_API_KEY },
      'muteHttpExceptions': true
    });
    Logger.log('   Status: ' + companyResponse.getResponseCode());
    Logger.log('   Response: ' + companyResponse.getContentText().substring(0, 500));
  } catch (e) {
    Logger.log('   ERROR: ' + e.toString());
  }
  Logger.log('');

  // 3. List all users
  Logger.log('3. LISTING ALL USERS (GET /users):');
  try {
    var usersUrl = DIALPAD_BASE_URL + '/users';
    var usersResponse = UrlFetchApp.fetch(usersUrl, {
      'method': 'get',
      'headers': { 'Authorization': 'Bearer ' + DIALPAD_API_KEY },
      'muteHttpExceptions': true
    });
    Logger.log('   Status: ' + usersResponse.getResponseCode());
    var usersData = JSON.parse(usersResponse.getContentText());

    if (usersData.items && usersData.items.length > 0) {
      Logger.log('   Found ' + usersData.items.length + ' users:');
      usersData.items.forEach(function(user, i) {
        Logger.log('   [' + i + '] ID: ' + user.id + ' | Name: ' + (user.display_name || user.first_name + ' ' + user.last_name) + ' | Email: ' + user.email);
      });
    } else {
      Logger.log('   Response: ' + JSON.stringify(usersData).substring(0, 500));
    }
  } catch (e) {
    Logger.log('   ERROR: ' + e.toString());
  }
  Logger.log('');

  // 4. Test specific user IDs
  Logger.log('4. TESTING USER IDS (GET /users/{id}):');
  var testUserIds = {
    'Dewayne': '5382638134673408',
    'Christine': '5825830596165632',
    'Jan': '5183185007984640',
    'Dewayne 2': '5966784611270656'
  };

  for (var name in testUserIds) {
    var userId = testUserIds[name];
    try {
      var userUrl = DIALPAD_BASE_URL + '/users/' + userId;
      var userResponse = UrlFetchApp.fetch(userUrl, {
        'method': 'get',
        'headers': { 'Authorization': 'Bearer ' + DIALPAD_API_KEY },
        'muteHttpExceptions': true
      });
      var status = userResponse.getResponseCode();
      if (status === 200) {
        var userData = JSON.parse(userResponse.getContentText());
        Logger.log('   ' + name + ' (' + userId + '): OK - ' + (userData.display_name || userData.email));
      } else {
        Logger.log('   ' + name + ' (' + userId + '): FAILED - Status ' + status + ' - ' + userResponse.getContentText().substring(0, 100));
      }
    } catch (e) {
      Logger.log('   ' + name + ' (' + userId + '): ERROR - ' + e.toString());
    }
  }
  Logger.log('');

  // 5. Test initiate_call endpoint (without actually calling)
  Logger.log('5. TESTING CLICK-TO-CALL ENDPOINT:');
  Logger.log('   Note: This tests if the endpoint exists and accepts requests');
  try {
    var callUrl = DIALPAD_BASE_URL + '/users/5382638134673408/initiate_call'; // Dewayne's correct ID
    var callResponse = UrlFetchApp.fetch(callUrl, {
      'method': 'post',
      'headers': {
        'Authorization': 'Bearer ' + DIALPAD_API_KEY,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify({ 'phone_number': '+10000000000' }), // Fake number to test endpoint
      'muteHttpExceptions': true
    });
    Logger.log('   Status: ' + callResponse.getResponseCode());
    Logger.log('   Response: ' + callResponse.getContentText().substring(0, 300));
  } catch (e) {
    Logger.log('   ERROR: ' + e.toString());
  }
  Logger.log('');

  Logger.log('========== END DIAGNOSTIC ==========');

  return {
    message: 'Check the Logs (View > Logs or Executions) for full diagnostic output',
    apiKeyPrefix: DIALPAD_API_KEY.substring(0, 10)
  };
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

// ============================================
// PHASE 2: ACTIVITY LOG SYSTEM
// ============================================

/**
 * Log activity to Activities sheet
 * @param {Object} activity - Activity data
 * @returns {string} Activity_ID
 */
function logActivity(activity) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('Activities');

    // Create sheet if doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('Activities');
      sheet.appendRow(['Activity_ID', 'Lead_Record_ID', 'User_ID', 'User_Name', 'Type', 'Date', 'Title', 'Description', 'Old_Value', 'New_Value', 'Duration', 'Related_ID']);
      sheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
    }

    const activityId = generateActivityId();
    const timestamp = new Date();

    sheet.appendRow([
      activityId,
      activity.leadId || '',
      activity.userId || '',
      activity.userName || '',
      activity.type,
      timestamp,
      activity.title,
      activity.description || '',
      activity.oldValue || '',
      activity.newValue || '',
      activity.duration || '',
      activity.relatedId || ''
    ]);

    return activityId;
  } catch (e) {
    Logger.log('Error logging activity: ' + e.toString());
    return null;
  }
}

/**
 * Generate Activity ID
 */
function generateActivityId() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Activities');
  const lastRow = sheet ? sheet.getLastRow() : 0;
  const nextNum = lastRow; // Row 1 is header, so lastRow = count
  return 'ACT-' + String(nextNum + 1).padStart(5, '0');
}

/**
 * Get activities for a specific lead
 * @param {string} leadId - Lead Record ID
 * @returns {Array} Activities array
 */
function getLeadActivities(leadId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Activities');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const headers = data[0];
    const activities = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === leadId) { // Column B is Lead_Record_ID
        const activity = {};
        headers.forEach((header, index) => {
          let value = data[i][index];
          if (value instanceof Date) {
            value = value.toISOString();
          }
          activity[header] = value;
        });
        activities.push(activity);
      }
    }

    // Sort by date descending (newest first)
    activities.sort((a, b) => new Date(b.Date) - new Date(a.Date));

    return JSON.parse(JSON.stringify(activities)); // Serialize for client
  } catch (e) {
    Logger.log('Error getting lead activities: ' + e.toString());
    return [];
  }
}

/**
 * Get activities for web (with serialization)
 */
function getLeadActivitiesForWeb(leadId) {
  return getLeadActivities(leadId);
}

// ============================================
// PHASE 2: EMAIL TRACKING SYSTEM
// ============================================

/**
 * Create email with tracking pixel
 */
function createEmailWithTracking(leadId, recipientEmail, subject, body, userId, userName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('Email_Tracking');

    // Create sheet if doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('Email_Tracking');
      sheet.appendRow(['Tracking_ID', 'Lead_Record_ID', 'User_ID', 'User_Name', 'Subject', 'Sent_Date', 'Recipient_Email', 'Opened', 'Open_Count', 'First_Open', 'Last_Open']);
      sheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
    }

    // Generate tracking ID
    const trackingId = 'TRK-' + Date.now();

    // Log to Email_Tracking sheet
    sheet.appendRow([
      trackingId,
      leadId,
      userId,
      userName,
      subject,
      new Date(),
      recipientEmail,
      'No',
      0,
      '',
      ''
    ]);

    // Log activity
    logActivity({
      leadId: leadId,
      userId: userId,
      userName: userName,
      type: 'Email',
      title: 'Email Sent',
      description: subject,
      relatedId: trackingId
    });

    return {
      trackingId: trackingId,
      success: true
    };
  } catch (e) {
    Logger.log('Error creating email tracking: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Log email open from tracking pixel
 */
function logEmailOpen(trackingId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Email_Tracking');
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // Find column indices
    const colIdx = {};
    headers.forEach((h, i) => colIdx[h] = i + 1);

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === trackingId) {
        const row = i + 1;
        const currentCount = data[i][colIdx['Open_Count'] - 1] || 0;
        const now = new Date();

        // Update Opened = Yes
        sheet.getRange(row, colIdx['Opened']).setValue('Yes');
        // Increment Open_Count
        sheet.getRange(row, colIdx['Open_Count']).setValue(currentCount + 1);
        // Set First_Open if not set
        if (!data[i][colIdx['First_Open'] - 1]) {
          sheet.getRange(row, colIdx['First_Open']).setValue(now);
        }
        // Update Last_Open
        sheet.getRange(row, colIdx['Last_Open']).setValue(now);

        // Log to Activity
        logActivity({
          leadId: data[i][1],
          userId: data[i][2],
          userName: data[i][3],
          type: 'Email_Opened',
          title: 'Email Opened',
          description: data[i][4], // Subject
          relatedId: trackingId
        });

        break;
      }
    }
  } catch (e) {
    Logger.log('Error logging email open: ' + e.toString());
  }
}

/**
 * Get email tracking info
 */
function getEmailTracking(trackingId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Email_Tracking');
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === trackingId) {
        const tracking = {};
        headers.forEach((h, idx) => {
          let value = data[i][idx];
          if (value instanceof Date) {
            value = value.toISOString();
          }
          tracking[h] = value;
        });
        return JSON.parse(JSON.stringify(tracking));
      }
    }
    return null;
  } catch (e) {
    Logger.log('Error getting email tracking: ' + e.toString());
    return null;
  }
}

/**
 * Get email analytics for dashboard
 */
function getEmailAnalytics(period, userId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Email_Tracking');
    if (!sheet) return getEmptyEmailAnalytics();

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return getEmptyEmailAnalytics();

    const headers = data[0];
    const colIdx = {};
    headers.forEach((h, i) => colIdx[h] = i);

    // Filter by date range
    const now = new Date();
    let startDate;
    switch(period) {
      case 'today':
        startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        break;
      case 'week':
        startDate = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
        break;
      case 'month':
        startDate = new Date(now.getFullYear(), now.getMonth(), 1);
        break;
      default:
        startDate = new Date(0); // All time
    }

    let totalSent = 0;
    let totalOpened = 0;
    let totalOpenCount = 0;
    let totalResponseTime = 0;
    let responseTimeCount = 0;
    const userStats = {};
    const subjectStats = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const sentDate = new Date(row[colIdx['Sent_Date']]);
      const rowUserId = row[colIdx['User_ID']];

      // Apply filters
      if (sentDate < startDate) continue;
      if (userId && rowUserId !== userId) continue;

      totalSent++;

      const opened = row[colIdx['Opened']] === 'Yes';
      const openCount = row[colIdx['Open_Count']] || 0;
      const firstOpen = row[colIdx['First_Open']];
      const subject = row[colIdx['Subject']] || 'No Subject';
      const userName = row[colIdx['User_Name']] || 'Unknown';

      if (opened) {
        totalOpened++;
        totalOpenCount += openCount;

        // Calculate response time (time to first open)
        if (firstOpen) {
          const responseTime = new Date(firstOpen) - sentDate;
          totalResponseTime += responseTime;
          responseTimeCount++;
        }
      }

      // Track by user
      if (!userStats[userName]) {
        userStats[userName] = { sent: 0, opened: 0, userId: rowUserId };
      }
      userStats[userName].sent++;
      if (opened) userStats[userName].opened++;

      // Track by subject
      if (!subjectStats[subject]) {
        subjectStats[subject] = { sent: 0, opened: 0 };
      }
      subjectStats[subject].sent++;
      if (opened) subjectStats[subject].opened++;
    }

    // Calculate user stats with open rates
    const userBreakdown = Object.entries(userStats).map(([name, stats]) => ({
      name,
      userId: stats.userId,
      sent: stats.sent,
      opened: stats.opened,
      openRate: stats.sent > 0 ? Math.round((stats.opened / stats.sent) * 100) : 0
    })).sort((a, b) => b.openRate - a.openRate);

    // Find best performing subject
    let bestSubject = { subject: 'N/A', openRate: 0, sent: 0 };
    Object.entries(subjectStats).forEach(([subject, stats]) => {
      if (stats.sent >= 3) { // Minimum 3 sends for significance
        const rate = (stats.opened / stats.sent) * 100;
        if (rate > bestSubject.openRate) {
          bestSubject = { subject, openRate: Math.round(rate), sent: stats.sent };
        }
      }
    });

    // Calculate avg response time in minutes
    const avgResponseTime = responseTimeCount > 0
      ? Math.round(totalResponseTime / responseTimeCount / 60000)
      : 0;

    return {
      totalSent,
      totalOpened,
      openRate: totalSent > 0 ? Math.round((totalOpened / totalSent) * 100) : 0,
      totalOpenCount,
      avgOpensPerEmail: totalOpened > 0 ? (totalOpenCount / totalOpened).toFixed(1) : 0,
      avgResponseTime,
      bestSubject,
      userBreakdown,
      period
    };
  } catch (e) {
    Logger.log('Error getting email analytics: ' + e.toString());
    return getEmptyEmailAnalytics();
  }
}

function getEmptyEmailAnalytics() {
  return {
    totalSent: 0,
    totalOpened: 0,
    openRate: 0,
    totalOpenCount: 0,
    avgOpensPerEmail: 0,
    avgResponseTime: 0,
    bestSubject: { subject: 'N/A', openRate: 0, sent: 0 },
    userBreakdown: [],
    period: 'all'
  };
}

function getEmailAnalyticsForWeb(period, userId) {
  return JSON.parse(JSON.stringify(getEmailAnalytics(period, userId)));
}

// ============================================
// PHASE 2: AUTO-ASSIGN LEADS SYSTEM
// ============================================

/**
 * Get next agent using round-robin assignment
 */
function getNextAgentRoundRobin() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const usersSheet = ss.getSheetByName(SHEETS.USERS);
  const leadsSheet = ss.getSheetByName(SHEETS.LEADS);

  // Get active team members (not Admin)
  const users = usersSheet.getDataRange().getValues();
  const headers = users[0];
  const colIdx = {};
  headers.forEach((h, i) => colIdx[h] = i);

  const activeAgents = [];
  for (let i = 1; i < users.length; i++) {
    const active = users[i][colIdx['Active']];
    const role = users[i][colIdx['Role']];
    const name = users[i][colIdx['Name']];
    const userId = users[i][colIdx['User_ID']] || 'USR-' + String(i).padStart(3, '0');

    if ((active === 'Yes' || active === true) && role === 'Team Member') {
      activeAgents.push({
        userId: userId,
        name: name
      });
    }
  }

  if (activeAgents.length === 0) return null;

  // Count leads per agent
  const leads = leadsSheet.getDataRange().getValues();
  const leadsHeaders = leads[0];
  const leadsColIdx = {};
  leadsHeaders.forEach((h, i) => leadsColIdx[h] = i);

  const leadCounts = {};
  activeAgents.forEach(a => leadCounts[a.name] = 0);

  for (let i = 1; i < leads.length; i++) {
    const owner = leads[i][leadsColIdx['Lead_Owner']];
    const location = leads[i][leadsColIdx['Location']];
    if (location === 'Active' && leadCounts.hasOwnProperty(owner)) {
      leadCounts[owner]++;
    }
  }

  // Find agent with fewest leads
  let minAgent = activeAgents[0];
  let minCount = leadCounts[minAgent.name];

  activeAgents.forEach(agent => {
    if (leadCounts[agent.name] < minCount) {
      minCount = leadCounts[agent.name];
      minAgent = agent;
    }
  });

  return minAgent;
}

/**
 * Get next agent based on capacity settings
 */
function getNextAgentByCapacity() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const usersSheet = ss.getSheetByName(SHEETS.USERS);
  const leadsSheet = ss.getSheetByName(SHEETS.LEADS);

  const users = usersSheet.getDataRange().getValues();
  const headers = users[0];
  const colIdx = {};
  headers.forEach((h, i) => colIdx[h] = i);

  const leads = leadsSheet.getDataRange().getValues();
  const leadsHeaders = leads[0];
  const leadsColIdx = {};
  leadsHeaders.forEach((h, i) => leadsColIdx[h] = i);

  // Build agent capacity map
  const agents = [];
  for (let i = 1; i < users.length; i++) {
    const active = users[i][colIdx['Active']];
    const role = users[i][colIdx['Role']];
    if ((active === 'Yes' || active === true) && role === 'Team Member') {
      agents.push({
        userId: users[i][colIdx['User_ID']] || 'USR-' + String(i).padStart(3, '0'),
        name: users[i][colIdx['Name']],
        maxLeads: users[i][colIdx['Max_Leads']] || 100 // Default 100
      });
    }
  }

  // Count current active leads per agent
  const leadCounts = {};
  agents.forEach(a => leadCounts[a.name] = 0);

  for (let i = 1; i < leads.length; i++) {
    const owner = leads[i][leadsColIdx['Lead_Owner']];
    const location = leads[i][leadsColIdx['Location']];
    if (location === 'Active' && leadCounts.hasOwnProperty(owner)) {
      leadCounts[owner]++;
    }
  }

  // Find agent with most remaining capacity
  let bestAgent = null;
  let maxRemaining = -1;

  agents.forEach(agent => {
    const remaining = agent.maxLeads - leadCounts[agent.name];
    if (remaining > maxRemaining) {
      maxRemaining = remaining;
      bestAgent = agent;
    }
  });

  return bestAgent;
}

/**
 * Get agent based on lead source routing rules
 */
function getAgentBySource(leadSource) {
  // Get routing rules from settings or use defaults
  const routingRules = {
    'Bark.com': 'Christine',
    'Seamless': 'Jan',
    'Referral': 'Dewayne',
    'Website': 'Kevin',
    'LinkedIn': 'Tiki'
  };

  if (routingRules[leadSource]) {
    return { name: routingRules[leadSource] };
  }

  return getNextAgentRoundRobin();
}

/**
 * Get agent based on lead location/region
 */
function getAgentByRegion(state) {
  const regionAssignments = {
    // East Coast
    'NY': 'Christine', 'NJ': 'Christine', 'PA': 'Christine', 'MA': 'Christine',
    // Midwest
    'OH': 'Jan', 'MI': 'Jan', 'IL': 'Jan', 'IN': 'Jan',
    // South
    'FL': 'Dewayne', 'GA': 'Dewayne', 'TX': 'Dewayne', 'NC': 'Dewayne',
    // West
    'CA': 'Kevin', 'WA': 'Kevin', 'OR': 'Kevin', 'AZ': 'Kevin'
  };

  if (regionAssignments[state]) {
    return { name: regionAssignments[state] };
  }

  return getNextAgentRoundRobin();
}

/**
 * Auto-assign a lead to an agent
 */
function autoAssignLead(lead, method) {
  method = method || 'roundrobin';
  let agent;

  switch(method) {
    case 'capacity':
      agent = getNextAgentByCapacity();
      break;
    case 'source':
      agent = getAgentBySource(lead.Lead_Source);
      break;
    case 'region':
      agent = getAgentByRegion(lead.Company_State || lead.Venue_State);
      break;
    case 'roundrobin':
    default:
      agent = getNextAgentRoundRobin();
  }

  if (agent) {
    // Log the assignment
    logActivity({
      leadId: lead.Record_ID,
      userId: 'SYSTEM',
      userName: 'System',
      type: 'Owner_Change',
      title: 'Auto-assigned to ' + agent.name,
      description: 'Method: ' + method,
      oldValue: lead.Lead_Owner || 'Unassigned',
      newValue: agent.name
    });
  }

  return agent;
}

/**
 * Bulk auto-assign unassigned leads
 */
function bulkAutoAssign(method) {
  method = method || 'roundrobin';

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.LEADS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const colIdx = {};
  headers.forEach((h, i) => colIdx[h] = i + 1);

  let assignedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const owner = data[i][colIdx['Lead_Owner'] - 1];
    const location = data[i][colIdx['Location'] - 1];

    // Only assign unassigned active leads
    if (location === 'Active' && (!owner || owner === '' || owner === 'Unassigned')) {
      const lead = {};
      headers.forEach((h, idx) => lead[h] = data[i][idx]);

      const agent = autoAssignLead(lead, method);
      if (agent) {
        sheet.getRange(i + 1, colIdx['Lead_Owner']).setValue(agent.name);
        assignedCount++;
      }
    }
  }

  return { assigned: assignedCount, method: method };
}

/**
 * Get auto-assign settings
 */
function getAutoAssignSettings() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('System_Settings');
    if (!sheet) {
      return { enabled: true, method: 'roundrobin', onImport: true, onCreate: true };
    }

    const data = sheet.getDataRange().getValues();
    const settings = {};

    for (let i = 1; i < data.length; i++) {
      settings[data[i][0]] = data[i][1];
    }

    return {
      enabled: settings['auto_assign_enabled'] === 'Yes',
      method: settings['auto_assign_method'] || 'roundrobin',
      onImport: settings['auto_assign_on_import'] === 'Yes',
      onCreate: settings['auto_assign_on_create'] === 'Yes'
    };
  } catch (e) {
    return { enabled: true, method: 'roundrobin', onImport: true, onCreate: true };
  }
}

/**
 * Save auto-assign settings
 */
function saveAutoAssignSettings(settings) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('System_Settings');

    if (!sheet) {
      sheet = ss.insertSheet('System_Settings');
      sheet.appendRow(['Setting_Key', 'Setting_Value']);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
    }

    const settingsToSave = [
      ['auto_assign_enabled', settings.enabled ? 'Yes' : 'No'],
      ['auto_assign_method', settings.method || 'roundrobin'],
      ['auto_assign_on_import', settings.onImport ? 'Yes' : 'No'],
      ['auto_assign_on_create', settings.onCreate ? 'Yes' : 'No']
    ];

    // Clear existing and write new
    const data = sheet.getDataRange().getValues();
    const existingKeys = {};
    for (let i = 1; i < data.length; i++) {
      existingKeys[data[i][0]] = i + 1;
    }

    settingsToSave.forEach(([key, value]) => {
      if (existingKeys[key]) {
        sheet.getRange(existingKeys[key], 2).setValue(value);
      } else {
        sheet.appendRow([key, value]);
      }
    });

    return { success: true };
  } catch (e) {
    Logger.log('Error saving auto-assign settings: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Get team workload for dashboard widget
 */
function getTeamWorkload() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const usersSheet = ss.getSheetByName(SHEETS.USERS);
  const leadsSheet = ss.getSheetByName(SHEETS.LEADS);

  const users = usersSheet.getDataRange().getValues();
  const headers = users[0];
  const colIdx = {};
  headers.forEach((h, i) => colIdx[h] = i);

  const leads = leadsSheet.getDataRange().getValues();
  const leadsHeaders = leads[0];
  const leadsColIdx = {};
  leadsHeaders.forEach((h, i) => leadsColIdx[h] = i);

  // Count leads per agent
  const leadCounts = {};
  for (let i = 1; i < leads.length; i++) {
    const owner = leads[i][leadsColIdx['Lead_Owner']];
    const location = leads[i][leadsColIdx['Location']];
    if (location === 'Active' && owner) {
      if (!leadCounts[owner]) leadCounts[owner] = 0;
      leadCounts[owner]++;
    }
  }

  // Build workload data
  const workload = [];
  for (let i = 1; i < users.length; i++) {
    const active = users[i][colIdx['Active']];
    const role = users[i][colIdx['Role']];
    if ((active === 'Yes' || active === true) && role === 'Team Member') {
      const name = users[i][colIdx['Name']];
      const maxLeads = users[i][colIdx['Max_Leads']] || 100;
      const currentLeads = leadCounts[name] || 0;

      workload.push({
        name: name,
        currentLeads: currentLeads,
        maxLeads: maxLeads,
        percentage: Math.min(Math.round((currentLeads / maxLeads) * 100), 100)
      });
    }
  }

  return workload;
}

/**
 * Get unassigned lead count
 */
function getUnassignedLeadCount() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.LEADS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const colIdx = {};
  headers.forEach((h, i) => colIdx[h] = i);

  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const owner = data[i][colIdx['Lead_Owner']];
    const location = data[i][colIdx['Location']];
    if (location === 'Active' && (!owner || owner === '' || owner === 'Unassigned')) {
      count++;
    }
  }

  return count;
}

// ============================================
// PHASE 2: RESEARCH WORKFLOW IMPROVEMENTS
// ============================================

/**
 * Generate quick search links for Bark research
 */
function generateSearchLinks(barkLead) {
  const city = barkLead.City_Location || barkLead.city || '';
  const eventType = barkLead.Event_Type || barkLead.eventType || '';
  const message = barkLead.Message || barkLead.message || '';
  const date = barkLead.Date_Needed || barkLead.dateNeeded || '';

  // Extract venue names from message
  const venueMatch = message.match(/(at the |at |@ )([A-Z][a-zA-Z\s]+)/);
  const venue = venueMatch ? venueMatch[2].trim() : '';

  // Build search queries
  const linkedInQuery = encodeURIComponent(eventType + ' ' + city + ' ' + venue + ' event planner');
  const googleQuery = encodeURIComponent(eventType + ' ' + city + ' ' + date + ' ' + venue);
  const facebookQuery = encodeURIComponent(city + ' ' + eventType + ' events ' + venue);

  return {
    linkedin: 'https://www.linkedin.com/search/results/people/?keywords=' + linkedInQuery,
    google: 'https://www.google.com/search?q=' + googleQuery,
    facebook: 'https://www.facebook.com/search/top?q=' + facebookQuery,
    googleMaps: venue ? 'https://www.google.com/maps/search/' + encodeURIComponent(venue + ' ' + city) : null,
    searchTerms: eventType + ' ' + city + ' ' + date
  };
}

/**
 * Extract info from LinkedIn URL
 */
function extractFromLinkedInUrl(url) {
  const match = url.match(/linkedin\.com\/in\/([^\/\?]+)/);
  if (match) {
    const profileSlug = match[1];
    // Convert slug to name: "john-doe-123abc" -> "John Doe"
    const nameParts = profileSlug.split('-').filter(p => !p.match(/^[a-f0-9]+$/));
    const name = nameParts.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
    return { name: name, linkedinUrl: url };
  }
  return null;
}

// ============================================
// PHASE 2: PERFORMANCE OPTIMIZATION
// ============================================

/**
 * Get leads with pagination
 */
function getLeadsPaginated(view, page, pageSize, filters) {
  page = page || 1;
  pageSize = pageSize || 50;
  filters = filters || {};

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.LEADS);
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return { leads: [], total: 0, page: 1, pageSize: pageSize, totalPages: 0 };
  }

  const headers = data[0];
  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const currentUser = filters.userName;

  // Filter leads based on view
  const filtered = [];
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
    const contractId = row[colIndex['Contract_ID']];
    const followupDate = row[colIndex['Date_Next_Followup']];

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
        // Show leads that have a contract generated (Contract_ID exists)
        include = (location === 'Active' && contractId && String(contractId).trim() !== '');
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
      case 'clients':
        // Clients = past booked clients (event date is in the past) or Location = 'Clients'
        if (location === 'Clients') {
          include = true;
        } else if (location === 'Booked') {
          const eventDateRaw = row[colIndex['Date_of_Event']];
          // Handle both Date objects and string dates
          let eventDate = null;
          if (eventDateRaw instanceof Date) {
            eventDate = new Date(eventDateRaw);
          } else if (eventDateRaw && typeof eventDateRaw === 'string') {
            eventDate = new Date(eventDateRaw);
          }
          if (eventDate && !isNaN(eventDate.getTime())) {
            eventDate.setHours(0, 0, 0, 0);
            include = (eventDate.getTime() < today.getTime());
          }
        }
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
      case 'overdue':
        if (location === 'Active' && followupDate instanceof Date) {
          const followup = new Date(followupDate);
          followup.setHours(0, 0, 0, 0);
          include = (followup.getTime() < today.getTime());
        }
        break;
      case 'nofollowup':
        include = (location === 'Active' && !followupDate);
        break;
      default:
        include = (location === 'Active');
    }

    if (include) {
      const lead = { rowIndex: i + 1 };
      headers.forEach((header, index) => {
        let value = row[index];
        if (value instanceof Date) {
          // Use formatDateForDisplay for event dates to prevent timezone shift
          if (header === 'Date_of_Event' || header === 'Event_Date') {
            value = formatDateForDisplay(value);
          } else {
            value = formatDate(value);
          }
        }
        lead[header] = value;
      });
      filtered.push(lead);
    }
  }

  // Sort by Priority_Score descending
  filtered.sort((a, b) => {
    const scoreA = parseFloat(a.Priority_Score) || 0;
    const scoreB = parseFloat(b.Priority_Score) || 0;
    return scoreB - scoreA;
  });

  // Paginate
  const start = (page - 1) * pageSize;
  const end = start + pageSize;
  const pageData = filtered.slice(start, end);

  return {
    leads: JSON.parse(JSON.stringify(pageData)),
    total: filtered.length,
    page: page,
    pageSize: pageSize,
    totalPages: Math.ceil(filtered.length / pageSize)
  };
}

/**
 * Get optimized dashboard data in a single call
 */
function getDashboardDataOptimized(currentUser) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const leadsSheet = ss.getSheetByName(SHEETS.LEADS);
  const eodSheet = ss.getSheetByName(SHEETS.EOD);

  // Get all leads data in one read
  const leadsData = leadsSheet.getDataRange().getValues();
  const headers = leadsData[0];
  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const sevenDaysAgo = new Date(today);
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

  // Calculate all stats in one pass
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
  let overdueFollowups = 0;
  let noFollowupSet = 0;
  let staleLeads = 0;

  const recentHotLeads = [];
  const todaysQueue = [];

  for (let i = 1; i < leadsData.length; i++) {
    const row = leadsData[i];
    const location = row[colIndex['Location']];
    const status = row[colIndex['Lead_Status']];
    const quote = row[colIndex['Active_Quote']];
    const contractId = row[colIndex['Contract_ID']];
    const followupDate = row[colIndex['Date_Next_Followup']];
    const owner = row[colIndex['Lead_Owner']];
    const dhillReview = row[colIndex['Dhill_Call_List']];
    const lastTouch = row[colIndex['Date_Last_Touch']];

    if (location === 'Active') {
      totalActive++;

      if (status === 'Hot' || status === 'Super Hot' || status === 'Contract Lead') {
        hotLeads++;
        // Add to recent hot leads (max 5)
        if (recentHotLeads.length < 5) {
          const lead = { rowIndex: i + 1 };
          headers.forEach((h, idx) => {
            let val = row[idx];
            if (val instanceof Date) val = formatDate(val);
            lead[h] = val;
          });
          recentHotLeads.push(lead);
        }
      } else if (status === 'Warm') {
        warmLeads++;
      }

      if (quote) {
        const numericQuote = parseFloat(String(quote).replace(/[$,]/g, ''));
        if (!isNaN(numericQuote)) {
          pipelineValue += numericQuote;
        }
      }

      // Contracts to review (leads that have a contract generated)
      if (contractId && String(contractId).trim() !== '') contractsReview++;

      if (followupDate instanceof Date) {
        const followup = new Date(followupDate);
        followup.setHours(0, 0, 0, 0);
        if (followup.getTime() === today.getTime()) {
          todayFollowUps++;
          // Add to today's queue (max 10)
          if (todaysQueue.length < 10) {
            const lead = { rowIndex: i + 1 };
            headers.forEach((h, idx) => {
              let val = row[idx];
              if (val instanceof Date) val = formatDate(val);
              lead[h] = val;
            });
            todaysQueue.push(lead);
          }
        }
        if (followup.getTime() < today.getTime()) overdueFollowups++;
      } else {
        noFollowupSet++;
      }

      if (lastTouch instanceof Date) {
        const lastTouchDate = new Date(lastTouch);
        lastTouchDate.setHours(0, 0, 0, 0);
        if (lastTouchDate.getTime() < sevenDaysAgo.getTime()) staleLeads++;
      }

      if (currentUser && owner === currentUser) myLeadsCount++;
      if (dhillReview === 'Yes') dhillCount++;

    } else if (location === 'Booked') {
      bookedCount++;
    } else if (location === 'Parking Lot') {
      parkingLotCount++;
    } else if (location === 'Storage') {
      storageCount++;
    }
  }

  const hotLeadsPercentage = totalActive > 0 ? Math.round((hotLeads / totalActive) * 100) : 0;
  const eodAverages = getEODAverages();

  return {
    stats: {
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
      dhillCount,
      hotLeadsPercentage,
      avgCallsMade: eodAverages.avgCalls,
      avgConversations: eodAverages.avgConversations,
      overdueFollowups,
      noFollowupSet,
      staleLeads
    },
    recentHotLeads: JSON.parse(JSON.stringify(recentHotLeads)),
    todaysQueue: JSON.parse(JSON.stringify(todaysQueue))
  };
}

// ============================================
// PHASE 2: USER ID SYSTEM ENHANCEMENT
// ============================================

/**
 * Generate User ID
 */
function generateUserId(rowNum) {
  return 'USR-' + String(rowNum).padStart(3, '0');
}

/**
 * Get user with User_ID included
 */
function getUserWithId(name) {
  const users = getUsers();
  for (let i = 0; i < users.length; i++) {
    if (users[i].Name && users[i].Name.toLowerCase() === name.toLowerCase()) {
      // Ensure User_ID exists
      if (!users[i].User_ID) {
        users[i].User_ID = generateUserId(i + 1);
      }
      return users[i];
    }
  }
  return null;
}

/**
 * Get Dialpad User ID for BEM user
 */
function getDialpadUserIdForUser(userName) {
  const user = getUserWithId(userName);
  if (user && user.Dialpad_ID) {
    return user.Dialpad_ID;
  }

  // Fallback map - Updated with correct Dialpad User IDs
  const userMap = {
    'Dewayne': '5382638134673408',
    'Christine': '5825830596165632',
    'Jan': '5183185007984640',
    'Dewayne 2': '5966784611270656'
  };

  return userMap[userName] || null;
}

// ============================================
// PHASE 2: CLICK-TO-CALL ENHANCEMENT
// ============================================

/**
 * Initiate call via Dialpad with user lookup
 */
function initiateCallForUser(phoneNumber, userName) {
  const dialpadUserId = getDialpadUserIdForUser(userName);

  if (!dialpadUserId) {
    return { success: false, error: 'Dialpad user not configured for ' + userName };
  }

  return initiateDialpadCall(phoneNumber, dialpadUserId);
}

/**
 * Log a call activity to a lead's timeline
 * Called from frontend after Dialpad call is initiated
 */
function logCallActivity(leadRecordId, phoneNumber, userName, callResult) {
  try {
    if (!leadRecordId) {
      return { success: false, error: 'No lead Record ID provided' };
    }

    const user = getUserWithId(userName);
    const userId = user ? user.User_ID : '';

    let title = 'Call Made';
    let description = 'Called ' + (phoneNumber || 'Unknown');

    if (callResult) {
      title = 'Call: ' + callResult;
      description = 'Called ' + (phoneNumber || 'Unknown') + ' - ' + callResult;
    }

    logActivity({
      leadId: leadRecordId,
      userId: userId,
      userName: userName,
      type: 'Call',
      title: title,
      description: description
    });

    return { success: true };
  } catch (e) {
    Logger.log('Error logging call activity: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ============================================
// ANALYTICS - INTELLIGENCE HUB
// ============================================

/**
 * Get analytics data with custom date range
 */
function getAnalyticsDataWithRange(startDate, endDate, period) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var leadsSheet = ss.getSheetByName('Leads');
    var eodSheet = ss.getSheetByName('EOD_Metrics');

    var start = startDate ? new Date(startDate) : null;
    var end = endDate ? new Date(endDate) : new Date();

    if (period && !startDate) {
      var now = new Date();
      switch(period) {
        case 'today': start = new Date(now.getFullYear(), now.getMonth(), now.getDate()); break;
        case 'week': start = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000); break;
        case 'month': start = new Date(now.getFullYear(), now.getMonth(), 1); break;
        case 'quarter': start = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000); break;
        default: start = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
      }
    }

    if (!start) start = new Date(end.getTime() - 7 * 24 * 60 * 60 * 1000);
    start.setHours(0, 0, 0, 0);
    end.setHours(23, 59, 59, 999);

    var leadsData = leadsSheet.getDataRange().getValues();
    var headers = leadsData[0];
    var colIdx = {};
    headers.forEach(function(h, i) { colIdx[h] = i; });

    var leads = [];
    for (var i = 1; i < leadsData.length; i++) {
      var row = leadsData[i];
      leads.push({
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

    var activeLeads = leads.filter(function(l) { return l.Location === 'Active'; });
    var hotLeads = activeLeads.filter(function(l) {
      var s = (l.Lead_Status || '').toLowerCase();
      return s.indexOf('hot') >= 0 || s === 'contract lead';
    });
    var bookedLeads = leads.filter(function(l) { return l.Location === 'Booked'; });

    var bookedInRange = bookedLeads.filter(function(l) {
      if (!l.Date_Last_Touch) return false;
      var d = new Date(l.Date_Last_Touch);
      return d >= start && d <= end;
    });

    var pipelineValue = activeLeads.reduce(function(sum, l) { return sum + l.Active_Quote; }, 0);
    var hotValue = hotLeads.reduce(function(sum, l) { return sum + l.Active_Quote; }, 0);
    var bookedValue = bookedInRange.reduce(function(sum, l) { return sum + l.Active_Quote; }, 0);

    // Previous period for comparison
    var periodLength = end.getTime() - start.getTime();
    var prevStart = new Date(start.getTime() - periodLength);
    var prevEnd = new Date(start.getTime() - 1);
    var prevBooked = bookedLeads.filter(function(l) {
      if (!l.Date_Last_Touch) return false;
      var d = new Date(l.Date_Last_Touch);
      return d >= prevStart && d <= prevEnd;
    });
    var prevBookedValue = prevBooked.reduce(function(sum, l) { return sum + l.Active_Quote; }, 0);
    var bookedChange = prevBookedValue > 0 ? Math.round(((bookedValue - prevBookedValue) / prevBookedValue) * 100) : 0;

    var totalValue = pipelineValue + bookedValue;
    var conversionRate = totalValue > 0 ? Math.round((bookedValue / totalValue) * 100) : 0;

    // Sources
    var sourceMap = {};
    activeLeads.forEach(function(l) {
      var src = l.Lead_Source || 'Unknown';
      if (!sourceMap[src]) sourceMap[src] = { count: 0, value: 0 };
      sourceMap[src].count++;
      sourceMap[src].value += l.Active_Quote;
    });
    var sourceData = Object.keys(sourceMap).map(function(name) {
      return { name: name, count: sourceMap[name].count, value: sourceMap[name].value };
    }).sort(function(a, b) { return b.value - a.value; }).slice(0, 8);

    // Funnel
    var funnel = {
      total: leads.length,
      totalValue: leads.reduce(function(s, l) { return s + l.Active_Quote; }, 0),
      active: activeLeads.length,
      activeValue: pipelineValue,
      hot: hotLeads.length,
      hotValue: hotValue,
      booked: bookedInRange.length,
      bookedValue: bookedValue
    };

    // Team leaderboard
    var teamMap = {};
    bookedInRange.forEach(function(l) {
      var owner = l.Lead_Owner || 'Unknown';
      if (!teamMap[owner]) teamMap[owner] = { name: owner, bookedValue: 0, bookedCount: 0, pipelineValue: 0, calls: 0, conversations: 0, aiScore: 0, points: 0 };
      teamMap[owner].bookedValue += l.Active_Quote;
      teamMap[owner].bookedCount++;
    });
    activeLeads.forEach(function(l) {
      var owner = l.Lead_Owner || 'Unknown';
      if (!teamMap[owner]) teamMap[owner] = { name: owner, bookedValue: 0, bookedCount: 0, pipelineValue: 0, calls: 0, conversations: 0, aiScore: 0, points: 0 };
      teamMap[owner].pipelineValue += l.Active_Quote;
    });

    // EOD metrics
    if (eodSheet) {
      var eodData = eodSheet.getDataRange().getValues();
      var eodHeaders = eodData[0];
      var eodColIdx = {};
      eodHeaders.forEach(function(h, i) { eodColIdx[h] = i; });
      for (var i = 1; i < eodData.length; i++) {
        var row = eodData[i];
        var dateVal = row[eodColIdx['Date']];
        if (dateVal) {
          var d = new Date(dateVal);
          if (d >= start && d <= end) {
            var user = row[eodColIdx['User']] || '';
            if (user && teamMap[user]) {
              teamMap[user].calls += parseInt(row[eodColIdx['Calls_Made']] || 0);
              teamMap[user].conversations += parseInt(row[eodColIdx['Conversations']] || 0);
            }
          }
        }
      }
    }

    // Points
    Object.keys(teamMap).forEach(function(user) {
      var m = teamMap[user];
      m.points = Math.round((m.bookedValue / 100) + (m.bookedCount * 25) + (m.calls * 1) + (m.conversations * 3));
    });
    var teamLeaderboard = Object.values(teamMap).sort(function(a, b) { return b.points - a.points; });

    // Trends
    var trends = [];
    var days = Math.ceil((end - start) / (1000 * 60 * 60 * 24));
    for (var i = 0; i < Math.min(days, 30); i++) {
      var d = new Date(start.getTime() + i * 24 * 60 * 60 * 1000);
      trends.push({ date: d.toISOString().split('T')[0], calls: 0, conversations: 0 });
    }
    if (eodSheet) {
      var eodData2 = eodSheet.getDataRange().getValues();
      var eodHeaders2 = eodData2[0];
      var eodColIdx2 = {};
      eodHeaders2.forEach(function(h, i) { eodColIdx2[h] = i; });
      for (var i = 1; i < eodData2.length; i++) {
        var row = eodData2[i];
        var dateVal = row[eodColIdx2['Date']];
        if (dateVal) {
          var rowDate = new Date(dateVal);
          var dateStr = rowDate.toISOString().split('T')[0];
          for (var j = 0; j < trends.length; j++) {
            if (trends[j].date === dateStr) {
              trends[j].calls += parseInt(row[eodColIdx2['Calls_Made']] || 0);
              trends[j].conversations += parseInt(row[eodColIdx2['Conversations']] || 0);
              break;
            }
          }
        }
      }
    }

    // Email stats
    var emailStats = { sent: 0, opened: 0, openRate: 0, bestSubject: 'N/A' };
    var emailSheet = ss.getSheetByName('Email_Tracking');
    if (emailSheet) {
      var emailData = emailSheet.getDataRange().getValues();
      if (emailData.length > 1) {
        var emailHeaders = emailData[0];
        var emailColIdx = {};
        emailHeaders.forEach(function(h, i) { emailColIdx[h] = i; });
        var subjectStats = {};
        for (var i = 1; i < emailData.length; i++) {
          var row = emailData[i];
          var sentDate = new Date(row[emailColIdx['Sent_Date']]);
          if (sentDate >= start && sentDate <= end) {
            emailStats.sent++;
            if (row[emailColIdx['Opened']] === 'Yes') emailStats.opened++;
            var subject = String(row[emailColIdx['Subject']] || 'No Subject');
            if (!subjectStats[subject]) subjectStats[subject] = { sent: 0, opened: 0 };
            subjectStats[subject].sent++;
            if (row[emailColIdx['Opened']] === 'Yes') subjectStats[subject].opened++;
          }
        }
        emailStats.openRate = emailStats.sent > 0 ? Math.round((emailStats.opened / emailStats.sent) * 100) : 0;
        var bestRate = 0;
        Object.keys(subjectStats).forEach(function(subj) {
          var s = subjectStats[subj];
          if (s.sent >= 2) {
            var rate = s.opened / s.sent;
            if (rate > bestRate) {
              bestRate = rate;
              emailStats.bestSubject = subj.length > 35 ? subj.substring(0, 35) + '...' : subj;
            }
          }
        });
      }
    }

    var recentBookings = bookedInRange.slice(0, 5).map(function(l) {
      return { event: l.Event_Name || l.Contact_Person, contact: l.Contact_Person, owner: l.Lead_Owner, value: l.Active_Quote, date: l.Date_Last_Touch };
    });

    return {
      dateRange: { start: start.toISOString(), end: end.toISOString(), period: period || 'custom' },
      kpis: {
        pipeline: { value: pipelineValue, count: activeLeads.length },
        hotLeads: { value: hotValue, count: hotLeads.length },
        booked: { value: bookedValue, count: bookedInRange.length, change: bookedChange },
        conversionRate: { value: conversionRate }
      },
      funnel: funnel,
      sourceData: sourceData,
      teamLeaderboard: teamLeaderboard,
      trends: trends,
      emailStats: emailStats,
      recentBookings: recentBookings
    };
  } catch (e) {
    Logger.log('Error: ' + e.toString());
    return null;
  }
}

// ============================================
// ADD RESEARCH ITEM (Alias for addToResearchQueue)
// ============================================

function addResearchItem(data, currentUser) {
  return addToResearchQueue(data, currentUser);
}

// ============================================
// GET CONTRACTS LIST (For Contracts Folder View)
// ============================================

function getContractsList() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.CONTRACTS);

    if (!sheet) {
      return { contracts: [] };
    }

    const data = sheet.getDataRange().getValues();

    if (data.length < 2) {
      return { contracts: [] };
    }

    const headers = data[0];
    const colIndex = {};
    headers.forEach((h, i) => colIndex[h] = i);

    // Log actual column headers for debugging
    Logger.log('Contracts sheet headers: ' + JSON.stringify(headers));

    // Build a lookup map for lead emails from Leads sheet
    const leadsSheet = ss.getSheetByName(SHEETS.LEADS);
    const leadEmailMap = {};
    if (leadsSheet) {
      const leadsData = leadsSheet.getDataRange().getValues();
      if (leadsData.length > 1) {
        const leadsHeaders = leadsData[0];
        const leadColIndex = {};
        leadsHeaders.forEach((h, i) => leadColIndex[h] = i);
        const recordIdCol = leadColIndex['Record_ID'];
        const emailCol = leadColIndex['Email'];
        if (recordIdCol !== undefined && emailCol !== undefined) {
          for (let i = 1; i < leadsData.length; i++) {
            const recordId = leadsData[i][recordIdCol];
            const email = leadsData[i][emailCol];
            if (recordId) {
              leadEmailMap[recordId] = email || '';
            }
          }
        }
      }
    }

    const contracts = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows
      if (!row[0]) continue;

      // Date Generated - check multiple possible column names
      // Based on generateContract(): column 8 (index 7) is timestamp
      let rawDateGenerated = row[colIndex['Generated_Date']] || row[colIndex['Date_Generated']] || row[colIndex['Created_Date']] || row[colIndex['Timestamp']] || row[7] || '';

      // Get timestamp for sorting (preserve full datetime precision)
      let sortTimestamp = 0;
      if (rawDateGenerated instanceof Date) {
        sortTimestamp = rawDateGenerated.getTime();
      } else if (rawDateGenerated && typeof rawDateGenerated === 'string') {
        const parsed = new Date(rawDateGenerated);
        if (!isNaN(parsed.getTime())) {
          sortTimestamp = parsed.getTime();
        }
      }

      // Format date for display - use Utilities.formatDate with script timezone for accurate display
      let dateGenerated = '';
      if (rawDateGenerated instanceof Date) {
        dateGenerated = Utilities.formatDate(rawDateGenerated, Session.getScriptTimeZone(), 'MMM d, yyyy HH:mm');
      } else if (rawDateGenerated && typeof rawDateGenerated === 'string') {
        // For string dates, use formatDateForDisplay to handle timezone correctly
        dateGenerated = formatDateForDisplay(rawDateGenerated);
      }

      Logger.log('Contract ' + (row[0] || 'unknown') + ' raw: ' + rawDateGenerated + ' timestamp: ' + sortTimestamp + ' display: ' + dateGenerated);

      // Amount - check multiple possible column names
      let amount = row[colIndex['Amount']] || row[colIndex['Quote_Amount']] || row[colIndex['Active_Quote']] || row[6] || '';
      if (typeof amount === 'string') {
        amount = parseFloat(amount.replace(/[$,]/g, '')) || 0;
      } else if (typeof amount === 'number') {
        amount = amount;
      }

      // Event date for display - use formatDateForDisplay to handle timezone correctly
      let rawEventDate = row[colIndex['Date_of_Event']] || row[colIndex['Event_Date']] || row[5] || '';
      let eventDate = formatDateForDisplay(rawEventDate);

      // PDF URL - Based on generateContract(): column 12 (index 11) is file.getUrl()
      let pdfUrl = row[colIndex['Drive_URL']] || row[colIndex['PDF_URL']] || row[colIndex['File_URL']] || row[colIndex['Contract_URL']] || row[colIndex['URL']] || row[11] || '';

      // Created By - Based on generateContract(): column 9 (index 8) is currentUser
      let createdBy = row[colIndex['Generated_By']] || row[colIndex['Created_By']] || row[colIndex['User']] || row[colIndex['Lead_Owner']] || row[8] || '';

      // Contact Name
      let contactName = row[colIndex['Contact_Person']] || row[colIndex['Contact_Name']] || row[3] || '';

      // Event Name
      let eventName = row[colIndex['Event_Name']] || row[4] || '';

      // Lead Record ID for email lookup
      let leadRecordId = row[colIndex['Lead_Record_ID']] || row[2] || '';

      // Client Email - look up from leads using leadRecordId
      let clientEmail = leadEmailMap[leadRecordId] || '';

      contracts.push({
        contractId: row[colIndex['Contract_ID']] || row[0] || '',
        leadRecordId: leadRecordId,
        contractName: row[colIndex['Filename']] || row[colIndex['File_Name']] || row[colIndex['Contract_Name']] || row[9] || '',
        eventName: eventName,
        eventDate: eventDate,
        contactName: contactName,
        clientEmail: clientEmail,
        createdBy: createdBy,
        amount: amount,
        dateGenerated: dateGenerated,
        _sortTimestamp: sortTimestamp,  // Internal field for sorting
        pdfUrl: pdfUrl,
        fileId: row[colIndex['File_ID']] || row[10] || '',
        status: row[colIndex['Status']] || row[12] || 'Draft'
      });
    }

    // Sort by timestamp descending (newest first)
    Logger.log('Before sort - first 3 contracts: ' + contracts.slice(0, 3).map(c => c.contractId + ':' + c._sortTimestamp).join(', '));

    contracts.sort((a, b) => {
      // DESC order: newest first (higher timestamp first)
      return (b._sortTimestamp || 0) - (a._sortTimestamp || 0);
    });

    Logger.log('After sort - first 3 contracts (should be newest): ' + contracts.slice(0, 3).map(c => c.contractId + ':' + c._sortTimestamp + ':' + c.dateGenerated).join(', '));

    return { contracts: contracts };

  } catch (e) {
    Logger.log('Error in getContractsList: ' + e.toString());
    return { contracts: [], error: e.toString() };
  }
}

// ============================================
// SEND CONTRACT VIA EMAIL
// ============================================

function sendContractEmail(contractId, leadRecordId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const contractsSheet = ss.getSheetByName(SHEETS.CONTRACTS);
    const leadsSheet = ss.getSheetByName(SHEETS.LEADS);

    // Find contract by ID
    const contractData = contractsSheet.getDataRange().getValues();
    const contractHeaders = contractData[0];
    const contractColIndex = {};
    contractHeaders.forEach((h, i) => contractColIndex[h] = i);

    let contract = null;
    let contractRowIndex = -1;
    for (let i = 1; i < contractData.length; i++) {
      if (contractData[i][0] === contractId) {
        contract = contractData[i];
        contractRowIndex = i + 1;
        break;
      }
    }

    if (!contract) {
      return { success: false, error: 'Contract not found' };
    }

    // Get lead info by Record_ID
    const leadsData = leadsSheet.getDataRange().getValues();
    const leadHeaders = leadsData[0];
    const leadColIndex = {};
    leadHeaders.forEach((h, i) => leadColIndex[h] = i);

    let lead = null;
    for (let i = 1; i < leadsData.length; i++) {
      if (leadsData[i][leadColIndex['Record_ID']] === leadRecordId) {
        lead = {};
        leadHeaders.forEach((h, idx) => lead[h] = leadsData[i][idx]);
        break;
      }
    }

    if (!lead) {
      return { success: false, error: 'Lead not found' };
    }

    if (!lead.Email) {
      return { success: false, error: 'Lead has no email address' };
    }

    // Get PDF file
    const fileId = contract[contractColIndex['File_ID']] || contract[10] || '';
    if (!fileId) {
      return { success: false, error: 'Contract PDF not found' };
    }

    const file = DriveApp.getFileById(fileId);

    // Get contract details
    const eventName = contract[contractColIndex['Event_Name']] || contract[4] || '';
    const invoiceNumber = contract[contractColIndex['Invoice_Number']] || contract[1] || '';
    const dateOfEvent = contract[contractColIndex['Date_of_Event']] || contract[5] || '';
    const activeQuote = contract[contractColIndex['Amount']] || contract[6] || 0;

    // Format date
    let formattedDate = dateOfEvent;
    if (dateOfEvent instanceof Date) {
      formattedDate = Utilities.formatDate(dateOfEvent, 'America/New_York', 'MMMM d, yyyy');
    }

    // Format quote
    let formattedQuote = '$0';
    if (activeQuote) {
      const quoteNum = parseFloat(String(activeQuote).replace(/[$,]/g, '')) || 0;
      formattedQuote = '$' + quoteNum.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    // Compose professional HTML email
    const subject = 'Your Entertainment Contract - ' + (eventName || 'Your Event') + ' | Dewayne Hill';
    const depositAmount = formatDepositAmount(activeQuote);
    const contactName = lead.Contact_Person || 'Valued Client';

    // Professional HTML email body - using solid color for email client compatibility
    const htmlBody = '<!DOCTYPE html>\
<html>\
<head>\
  <meta charset="utf-8">\
  <meta name="viewport" content="width=device-width, initial-scale=1.0">\
</head>\
<body style="margin: 0; padding: 0; font-family: \'Helvetica Neue\', Arial, sans-serif; background-color: #f5f5f5;">\
  <table width="100%" cellpadding="0" cellspacing="0" style="max-width: 600px; margin: 0 auto; background-color: #ffffff;">\
    <tr>\
      <td style="background-color: #2c3e50; padding: 30px; text-align: center;">\
        <h1 style="color: #ffffff; font-size: 22px; margin: 0;">Dewayne Hill Entertainment</h1>\
        <p style="color: #bdc3c7; font-size: 13px; margin: 8px 0 0 0;">America\'s Funniest Comedy Magician</p>\
      </td>\
    </tr>\
    <tr>\
      <td style="padding: 40px 30px;">\
        <h2 style="color: #1a1a2e; font-size: 24px; margin: 0 0 20px 0;">\
          Hello ' + escapeHtmlEmail(contactName) + '!\
        </h2>\
        <p style="color: #444; font-size: 16px; line-height: 1.6; margin: 0 0 25px 0;">\
          Thank you for choosing Dewayne Hill Entertainment for your upcoming event! \
          I\'m thrilled to bring the magic and laughter to <strong>' + escapeHtmlEmail(eventName || 'your event') + '</strong>.\
        </p>\
        <table width="100%" cellpadding="0" cellspacing="0" style="background: #f8f9fa; border-radius: 12px; margin: 25px 0;">\
          <tr>\
            <td style="padding: 25px;">\
              <h3 style="color: #1a1a2e; font-size: 18px; margin: 0 0 20px 0; border-bottom: 2px solid #e0e0e0; padding-bottom: 10px;">\
                Contract Details\
              </h3>\
              <table width="100%" cellpadding="8" cellspacing="0">\
                <tr>\
                  <td style="color: #666; font-size: 14px;">Event</td>\
                  <td style="color: #1a1a2e; font-size: 14px; font-weight: 600; text-align: right;">' + escapeHtmlEmail(eventName || 'Your Event') + '</td>\
                </tr>\
                <tr>\
                  <td style="color: #666; font-size: 14px;">Date</td>\
                  <td style="color: #1a1a2e; font-size: 14px; font-weight: 600; text-align: right;">' + escapeHtmlEmail(formattedDate) + '</td>\
                </tr>\
                <tr>\
                  <td style="color: #666; font-size: 14px;">Total Amount</td>\
                  <td style="color: #1a1a2e; font-size: 14px; font-weight: 600; text-align: right;">' + formattedQuote + '</td>\
                </tr>\
                <tr style="background: #fff3cd; border-radius: 6px;">\
                  <td style="color: #856404; font-size: 14px; padding: 12px 8px;">Deposit Due (50%)</td>\
                  <td style="color: #856404; font-size: 16px; font-weight: 700; text-align: right; padding: 12px 8px;">' + depositAmount + '</td>\
                </tr>\
              </table>\
            </td>\
          </tr>\
        </table>\
        <p style="color: #444; font-size: 16px; line-height: 1.6; margin: 25px 0;">\
          Please review the attached contract, sign where indicated, and return it at your earliest convenience \
          to secure your date.\
        </p>\
        <table width="100%" cellpadding="0" cellspacing="0" style="background: #e8f4f8; border-radius: 12px; margin: 25px 0;">\
          <tr>\
            <td style="padding: 25px;">\
              <h3 style="color: #1a1a2e; font-size: 16px; margin: 0 0 15px 0;">\
                Payment Options\
              </h3>\
              <p style="color: #444; font-size: 14px; line-height: 1.6; margin: 0 0 10px 0;">\
                <strong>Credit Card:</strong> Simply reply to this email and I\'ll send you a secure payment link.\
              </p>\
              <p style="color: #444; font-size: 14px; line-height: 1.6; margin: 0;">\
                <strong>Check:</strong> Make payable to <em>Dewayne Hill</em> and mail to:<br>\
                <span style="display: inline-block; margin-top: 8px; padding-left: 15px;">\
                  Dewayne Hill<br>\
                  13913 Lazy Oak Drive<br>\
                  Tampa, FL 33613\
                </span>\
              </p>\
            </td>\
          </tr>\
        </table>\
        <p style="color: #444; font-size: 16px; line-height: 1.6; margin: 25px 0 10px 0;">\
          Questions? I\'m here to help!\
        </p>\
        <p style="color: #444; font-size: 14px; margin: 0;">\
          <a href="mailto:dhillmagic@gmail.com" style="color: #0066cc;">dhillmagic@gmail.com</a><br>\
          <a href="tel:855-306-2442" style="color: #0066cc;">855-306-2442</a>\
        </p>\
      </td>\
    </tr>\
    <tr>\
      <td style="padding: 20px 30px 40px; text-align: center; border-top: 1px solid #eee;">\
        <p style="color: #1a1a2e; font-size: 16px; font-weight: 600; margin: 5px 0;">Dewayne Hill</p>\
        <p style="color: #666; font-size: 13px; margin: 0;">America\'s Funniest Comedy Magician</p>\
        <p style="color: #999; font-size: 12px; margin: 10px 0 0 0;">Making Magic Happen Since 2005</p>\
      </td>\
    </tr>\
    <tr>\
      <td style="background: #1a1a2e; padding: 20px 30px; text-align: center;">\
        <p style="color: #888; font-size: 11px; margin: 0;">\
          Contract ' + escapeHtmlEmail(invoiceNumber) + ' | Dewayne Hill Entertainment<br>\
          Tampa, FL | <a href="mailto:dhillmagic@gmail.com" style="color: #888;">dhillmagic@gmail.com</a>\
        </p>\
      </td>\
    </tr>\
  </table>\
</body>\
</html>';

    // Plain text fallback
    const plainBody = 'Hello ' + contactName + '!\n\n' +
      'Thank you for choosing Dewayne Hill Entertainment for ' + (eventName || 'your event') + '!\n\n' +
      'CONTRACT DETAILS\n' +
      '----------------\n' +
      'Event: ' + (eventName || 'Your Event') + '\n' +
      'Date: ' + formattedDate + '\n' +
      'Total Amount: ' + formattedQuote + '\n' +
      'Deposit Due (50%): ' + depositAmount + '\n\n' +
      'Please review the attached contract, sign where indicated, and return it at your earliest convenience.\n\n' +
      'PAYMENT OPTIONS\n' +
      '---------------\n' +
      'Credit Card: Reply to this email and I\'ll send a secure payment link.\n' +
      'Check: Make payable to Dewayne Hill, mail to:\n' +
      '  Dewayne Hill\n' +
      '  13913 Lazy Oak Drive\n' +
      '  Tampa, FL 33613\n\n' +
      'Questions? Contact me:\n' +
      'Email: dhillmagic@gmail.com\n' +
      'Phone: 855-306-2442\n\n' +
      'Looking forward to making your event magical!\n\n' +
      'Dewayne Hill\n' +
      'America\'s Funniest Comedy Magician';

    // Send email with HTML
    GmailApp.sendEmail(lead.Email, subject, plainBody, {
      htmlBody: htmlBody,
      attachments: [file.getAs(MimeType.PDF)],
      name: 'Dewayne Hill Entertainment',
      replyTo: 'dhillmagic@gmail.com'
    });

    // Update contract status to 'Sent'
    const statusCol = contractColIndex['Status'] ? contractColIndex['Status'] + 1 : 13;
    contractsSheet.getRange(contractRowIndex, statusCol).setValue('Sent');

    // Log activity
    logActivity({
      leadId: leadRecordId,
      userId: '',
      userName: 'System',
      type: 'Email',
      title: 'Contract Sent',
      description: 'Contract ' + invoiceNumber + ' emailed to ' + lead.Email,
      relatedId: contractId
    });

    return { success: true, email: lead.Email, message: 'Contract sent successfully' };

  } catch (e) {
    Logger.log('Error sending contract email: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// Helper function to format deposit amount (50%)
function formatDepositAmount(quote) {
  if (!quote) return '$0.00';
  const quoteNum = parseFloat(String(quote).replace(/[$,]/g, '')) || 0;
  const deposit = quoteNum / 2;
  return '$' + deposit.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Helper function to escape HTML for email content
function escapeHtmlEmail(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
