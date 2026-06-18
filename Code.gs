// ============================================================
//  NegOr HS Hub — Community Management Portal
//  Google Apps Script Backend  (Code.gs)
//  Version 3.0  — sections-based recaps, photo archive
// ============================================================
//
//  SETUP INSTRUCTIONS:
//  1. Create two Google Sheets:
//     a) Membership Sheet — for family registrations
//     b) Events Sheet     — for events and event registrations (separate)
//  2. Paste this file into a new Apps Script project
//     (open the Membership Sheet → Extensions → Apps Script)
//  3. Fill in all CONFIG values below
//  4. Run setupSpreadsheet()       ONCE — sets up membership sheet tabs
//  5. Run setupEventsSpreadsheet() ONCE — sets up events sheet tabs
//  6. Deploy as Web App: Execute as Me, Who has access: Anyone
//  7. Copy the Web App URL into all HTML files (APPS_SCRIPT_URL)
// ============================================================


// ─────────────────────────────────────────────────────────────
//  CONFIGURATION  ← Edit ALL values before deploying
// ─────────────────────────────────────────────────────────────

const CONFIG = {

  // ── MEMBERSHIP SPREADSHEET ───────────────────────────────
  // URL: https://docs.google.com/spreadsheets/d/<<THIS_PART>>/edit
  SPREADSHEET_ID: '1ArP7C-s3lcxz9mnu9PErhDnOMCj9XDjnSGDoAI2Fpdw',

  // ── EVENTS SPREADSHEET (separate Google Sheet) ───────────
  // Create a NEW blank spreadsheet at sheets.new, copy its ID here
  EVENTS_SPREADSHEET_ID: '1xpCoJ5-RNDeFf0AaEFXEtoXZoqIUBRMG4tmViedZ2Cg',

  // ── GMAIL ────────────────────────────────────────────────
  // Receives new-registration and new-event-registration notifications
  ADMIN_EMAIL: 'negorhshub@gmail.com',

  // ── TOKENS ───────────────────────────────────────────────
  // ADMIN_TOKEN   — full access (admin dashboard, all data)
  // COORDINATOR_TOKEN — event setup, view registrations (share with coordinators & volunteers)
  // Generate tokens at: https://www.uuidgenerator.net/
  ADMIN_TOKEN:       'negor_admin',
  COORDINATOR_TOKEN: 'negor_coord',

  // ── COMMUNITY NAME ───────────────────────────────────────
  COMMUNITY_NAME: 'NegOr HS Hub',

  // ── MEMBERSHIP SHEET TAB NAMES ───────────────────────────
  SHEET_FAMILIES : 'Families',
  SHEET_PARENTS  : 'Parents',
  SHEET_CHILDREN : 'Children',

  // ── EVENTS SHEET TAB NAMES ───────────────────────────────
  SHEET_EVENTS    : 'Events',
  SHEET_EVENT_REGS: 'Registrations',
  SHEET_RECAPS    : 'Recaps',
  SHEET_BOOKS     : 'Books',
  SHEET_CHECKLIST : 'Checklist',
  SHEET_VENUES    : 'Venues',
  SHEET_EQUIPMENT : 'Equipment',
  SHEET_ACTIVITIES: 'ActivityLog',

  // ── GOOGLE DRIVE FOLDERS ─────────────────────────────────
  // Create folders in Drive, open each, copy the ID from the URL:
  // https://drive.google.com/drive/folders/<<THIS_PART>>

  // For event registration payment proofs and QR codes:
  PROOF_OF_PAYMENT_FOLDER_ID: '1Nb7-27Q4no3PDPXaoR6KyZZ8DoETgWOZ',

  // For recap section photos (displayed on Recap Corner page):
  // Leave empty to reuse payment proof folder as fallback.
  RECAP_PHOTOS_FOLDER_ID: '1PxN1HZ70ToaWoS2iZs_29Ukzygf87F7Y',
};


// ─────────────────────────────────────────────────────────────
//  POST HANDLER  — all form submissions land here
// ─────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    switch (data._type) {
      case 'event':             return doPostEventRegistration(data);
      case 'createEvent':       return doPostCreateEvent(data);
      case 'updateEvent':       return doPostUpdateEvent(data);
      case 'markPaid':          return doPostMarkPaid(data);
      case 'createRecap':       return doPostCreateRecap(data);
      case 'uploadPhoto':       return doPostUploadPhoto(data);
      case 'addBook':           return doPostAddBook(data);
      case 'editBook':          return doPostEditBook(data);
      case 'deleteBook':        return doPostDeleteBook(data);
      case 'borrowBook':        return doPostBorrowBook(data);
      case 'returnBook':        return doPostReturnBook(data);
      case 'markChecklist':     return doPostMarkChecklistItem(data);
      case 'addVenue':          return doPostAddVenue(data);
      case 'addEquipment':      return doPostAddEquipment(data);
      case 'addActivity':       return doPostAddActivity(data);
      default:                  return doPostMembership(data);
    }
  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return jsonResponse({ success: false, error: err.toString() });
  }
}


// ─────────────────────────────────────────────────────────────
//  GET HANDLER  — admin & coordinator data APIs
// ─────────────────────────────────────────────────────────────

function doGet(e) {
  const params = e.parameter || {};
  const token  = params.token || '';
  const action = params.action || '';

  const isAdmin       = token === CONFIG.ADMIN_TOKEN;
  const isCoordinator = token === CONFIG.COORDINATOR_TOKEN || isAdmin;

  // ── Public (no token) ────────────────────────────────────
  if (action === 'getEvent')        return getPublicEvent(params.eventId);
  if (action === 'getActiveEvents') return getActiveEventsData();
  if (action === 'getRecaps')       return getRecapsData();
  if (action === 'getBooks')        return getBooksData();
  if (action === 'getBirthdays')    return getBirthdaysData();
  if (action === 'getVenues')       return getVenuesData();
  if (action === 'getEquipment')    return getEquipmentData();
  if (action === 'getActivityLog')  return getActivityLogData();

  // ── Coordinator & Admin actions ──────────────────────────
  const coordActions = ['getEvents','getEventRegistrations','getChecklist'];
  if (coordActions.includes(action)) {
    if (!isCoordinator) return jsonResponse({ success: false, error: 'Unauthorized' });
    switch (action) {
      case 'getEvents':             return getEventsData();
      case 'getEventRegistrations': return getEventRegistrationsData(params.eventId);
      case 'getChecklist':          return getChecklistData(params.eventId);
    }
  }

  // ── Admin-only actions ───────────────────────────────────
  if (!isAdmin) return jsonResponse({ success: false, error: 'Unauthorized' });

  switch (action) {
    case 'getFamilies': return getFamiliesData(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID));
    case 'getStats':    return getStatsData(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID));
    case 'getFamily':   return getFamilyDetail(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID), params.familyId);
    default:            return jsonResponse({ success: false, error: 'Unknown action: ' + action });
  }
}


// ─────────────────────────────────────────────────────────────
//  MEMBERSHIP REGISTRATION
// ─────────────────────────────────────────────────────────────

function doPostMembership(data) {
  const ss          = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const familySheet = ss.getSheetByName(CONFIG.SHEET_FAMILIES);
  const nextRow     = familySheet.getLastRow();
  const familyId    = 'FAM-' + String(nextRow).padStart(4, '0');
  const timestamp   = new Date();

  var proofUrl = '';
  if (data.paymentProof && data.paymentProof.base64) {
    try { proofUrl = savePaymentProof(familyId, data.paymentProof); }
    catch(e) { Logger.log('Membership proof upload failed: ' + e.toString()); }
  }

  writeFamilyRow(familySheet, familyId, timestamp, data, proofUrl);
  writeParentRows(ss.getSheetByName(CONFIG.SHEET_PARENTS), familyId, data.familyName, data.parents);
  writeChildRows(ss.getSheetByName(CONFIG.SHEET_CHILDREN), familyId, data.familyName, data.children);

  try { sendConfirmationEmail(data.parents[0], data.familyName, familyId); }
  catch(e) { Logger.log('Confirmation email failed: ' + e.toString()); }

  try { sendAdminNotification(familyId, data, timestamp, proofUrl); }
  catch(e) { Logger.log('Admin notification failed: ' + e.toString()); }

  return jsonResponse({ success: true, familyId: familyId });
}


// ─────────────────────────────────────────────────────────────
//  EVENT MANAGEMENT (coordinator)
// ─────────────────────────────────────────────────────────────

function doPostCreateEvent(data) {
  if (data._token !== CONFIG.COORDINATOR_TOKEN && data._token !== CONFIG.ADMIN_TOKEN) {
    return jsonResponse({ success: false, error: 'Unauthorized' });
  }
  const ss        = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const evSheet   = ss.getSheetByName(CONFIG.SHEET_EVENTS);
  if (!evSheet)   return jsonResponse({ success: false, error: 'Events sheet not found. Run setupEventsSpreadsheet() first.' });

  const nextRow   = evSheet.getLastRow();
  const eventId   = 'EVT-' + String(nextRow).padStart(4, '0');
  const timestamp = new Date();

  var qrUrl = '';
  if (data.qrCode && data.qrCode.base64) {
    try { qrUrl = saveEventQrCode(eventId, data.qrCode); }
    catch(e) { Logger.log('QR save failed: ' + e.toString()); }
  }

  // Columns A–O (15 columns)
  evSheet.appendRow([
    eventId,
    timestamp,
    data.title            || '',
    data.date             || '',
    data.venue            || '',
    data.fee              || 0,
    data.payeeName        || 'Debbie Uy Matiao',
    data.description      || '',
    'Draft',
    0,
    data.startTime        || '',
    data.endTime          || '',
    data.registrationType || 'per_family',
    qrUrl,
    data.photoFolderUrl   || '',  // col O — Google Drive folder link for event photos
  ]);

  return jsonResponse({ success: true, eventId: eventId });
}

function doPostUpdateEvent(data) {
  if (data._token !== CONFIG.COORDINATOR_TOKEN && data._token !== CONFIG.ADMIN_TOKEN) {
    return jsonResponse({ success: false, error: 'Unauthorized' });
  }
  const ss      = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const evSheet = ss.getSheetByName(CONFIG.SHEET_EVENTS);
  if (!evSheet) return jsonResponse({ success: false, error: 'Events sheet not found.' });

  const rows = evSheet.getDataRange().getValues();
  var found  = false;
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.eventId) {
      if (data.status)                     evSheet.getRange(i + 1,  9).setValue(data.status);
      if (data.title)                      evSheet.getRange(i + 1,  3).setValue(data.title);
      if (data.date)                       evSheet.getRange(i + 1,  4).setValue(data.date);
      if (data.venue)                      evSheet.getRange(i + 1,  5).setValue(data.venue);
      if (data.fee !== undefined)          evSheet.getRange(i + 1,  6).setValue(data.fee);
      if (data.payeeName)                  evSheet.getRange(i + 1,  7).setValue(data.payeeName);
      if (data.description !== undefined)  evSheet.getRange(i + 1,  8).setValue(data.description);
      if (data.startTime   !== undefined)  evSheet.getRange(i + 1, 11).setValue(data.startTime);
      if (data.endTime     !== undefined)  evSheet.getRange(i + 1, 12).setValue(data.endTime);
      if (data.registrationType)                 evSheet.getRange(i + 1, 13).setValue(data.registrationType);
      if (data.photoFolderUrl !== undefined)     evSheet.getRange(i + 1, 15).setValue(data.photoFolderUrl);
      if (data.qrCode && data.qrCode.base64) {
        try {
          var qrUrl = saveEventQrCode(data.eventId, data.qrCode);
          evSheet.getRange(i + 1, 14).setValue(qrUrl);
        } catch(e) { Logger.log('QR update failed: ' + e.toString()); }
      }
      found = true;
      break;
    }
  }
  if (!found) return jsonResponse({ success: false, error: 'Event not found: ' + data.eventId });
  return jsonResponse({ success: true });
}

function doPostMarkPaid(data) {
  if (data._token !== CONFIG.COORDINATOR_TOKEN && data._token !== CONFIG.ADMIN_TOKEN) {
    return jsonResponse({ success: false, error: 'Unauthorized' });
  }
  if (!data.registrationId) return jsonResponse({ success: false, error: 'registrationId required' });
  const status = data.paymentStatus === 'Paid' ? 'Paid' : 'Unpaid';

  const ss      = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const rgSheet = ss.getSheetByName(CONFIG.SHEET_EVENT_REGS);
  if (!rgSheet) return jsonResponse({ success: false, error: 'Registrations sheet not found.' });

  const rows = rgSheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.registrationId) {
      rgSheet.getRange(i + 1, 15).setValue(status);
      return jsonResponse({ success: true, paymentStatus: status });
    }
  }
  return jsonResponse({ success: false, error: 'Registration not found: ' + data.registrationId });
}

function getActiveEventsData() {
  try {
    const ss      = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
    const evSheet = ss.getSheetByName(CONFIG.SHEET_EVENTS);
    if (!evSheet) return jsonResponse({ success: true, data: [] });
    const events  = sheetToObjects(evSheet);
    const active  = events.filter(function(ev) { return ev['status'] === 'Active'; });
    return jsonResponse({ success: true, data: active });
  } catch(err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function getPublicEvent(eventId) {
  if (!eventId) return jsonResponse({ success: false, error: 'eventId required' });
  const ss      = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const evSheet = ss.getSheetByName(CONFIG.SHEET_EVENTS);
  if (!evSheet) return jsonResponse({ success: false, error: 'Events sheet not found.' });

  const events = sheetToObjects(evSheet);
  const event  = events.find(function(ev) { return ev['event_id'] === eventId; });

  if (!event)                        return jsonResponse({ success: false, error: 'Event not found.' });
  if (event['status'] !== 'Active')  return jsonResponse({ success: false, error: 'Registration is not open for this event.', status: event['status'] });

  return jsonResponse({ success: true, data: event });
}

function getEventsData() {
  try {
    const ss      = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
    const evSheet = ss.getSheetByName(CONFIG.SHEET_EVENTS);
    const rgSheet = ss.getSheetByName(CONFIG.SHEET_EVENT_REGS);

    if (!evSheet) return jsonResponse({ success: true, data: [], _warning: 'Events sheet not found — run setupEventsSpreadsheet() in Apps Script.' });

    const events = sheetToObjects(evSheet);
    const regs   = rgSheet ? sheetToObjects(rgSheet) : [];

    const enriched = events.map(function(ev) {
      const count = regs.filter(function(r) { return r['event_id'] === ev['event_id']; }).length;
      return Object.assign({}, ev, { registration_count: count });
    });

    return jsonResponse({ success: true, data: enriched });
  } catch(err) {
    Logger.log('getEventsData error: ' + err.toString());
    return jsonResponse({ success: true, data: [], _warning: 'Could not open Events spreadsheet: ' + err.toString() });
  }
}

function getEventRegistrationsData(eventId) {
  if (!eventId) return jsonResponse({ success: false, error: 'eventId required' });
  try {
    const ss      = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
    const rgSheet = ss.getSheetByName(CONFIG.SHEET_EVENT_REGS);
    if (!rgSheet) return jsonResponse({ success: true, data: [] });
    const all  = sheetToObjects(rgSheet);
    const regs = all.filter(function(r) { return r['event_id'] === eventId; });
    return jsonResponse({ success: true, data: regs });
  } catch(err) {
    return jsonResponse({ success: true, data: [] });
  }
}


// ─────────────────────────────────────────────────────────────
//  EVENT REGISTRATION (family-facing form submission)
// ─────────────────────────────────────────────────────────────

function doPostEventRegistration(data) {
  const ss      = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const evSheet = ss.getSheetByName(CONFIG.SHEET_EVENTS);
  const rgSheet = ss.getSheetByName(CONFIG.SHEET_EVENT_REGS);

  if (!evSheet || !rgSheet) {
    return jsonResponse({ success: false, error: 'Events spreadsheet not set up. Run setupEventsSpreadsheet() first.' });
  }

  const events = sheetToObjects(evSheet);
  const event  = events.find(function(ev) { return ev['event_id'] === data.eventId; });
  if (!event)                       return jsonResponse({ success: false, error: 'Event not found.' });
  if (event['status'] !== 'Active') return jsonResponse({ success: false, error: 'Registration is not currently open.' });

  const existingRegs = sheetToObjects(rgSheet).filter(function(r) { return r['event_id'] === data.eventId; });
  const regNum       = existingRegs.length + 1;
  const regId        = data.eventId + '-R' + String(regNum).padStart(3, '0');
  const timestamp    = new Date();

  var proofUrl = '';
  if (!data.payLater && data.paymentProof && data.paymentProof.base64) {
    try {
      // Naming: yyyy_mm_eventname_parentname
      var now2   = new Date();
      var yyyy   = now2.getFullYear();
      var mm2    = String(now2.getMonth() + 1).padStart(2, '0');
      var evName = (data.eventTitle || event['event_title'] || 'event')
                     .toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/_+$/,'');
      var pName  = ((data.parents && data.parents[0] && data.parents[0].name) || 'family')
                     .toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/_+$/,'');
      data.paymentProof.name = yyyy + '_' + mm2 + '_' + evName + '_' + pName;
      proofUrl = savePaymentProof(regId, data.paymentProof);
    } catch(e) { Logger.log('Event proof upload failed: ' + e.toString()); }
  }

  const parents  = (data.parents  || []).map(function(p) { return p.name; }).join(', ');
  const children = (data.children || []).map(function(c) { return c.name + ' (age ' + c.age + ')'; }).join(', ');

  rgSheet.appendRow([
    regId,
    data.eventId              || '',
    data.eventTitle           || event['event_title'] || '',
    timestamp,
    data.familyName           || '',
    parents,
    children,
    data.excessHandling       || 'refund',
    data.voluntaryAmount      || '0',
    data.voluntaryDestination || '',
    data.organizingRole       || 'attendee',
    data.notes                || '',
    proofUrl,
    data.consentData ? 'Yes' : 'No',
    data.payLater ? 'Pay at Event' : (proofUrl ? 'Pending' : 'Unpaid'),
  ]);

  try { sendEventAdminNotification(regId, data, event, timestamp, proofUrl); }
  catch(e) { Logger.log('Event admin email failed: ' + e.toString()); }

  return jsonResponse({ success: true, registrationId: regId });
}


// ─────────────────────────────────────────────────────────────
//  FILE UPLOADS
// ─────────────────────────────────────────────────────────────

/**
 * Upload any photo to Drive and return a public embeddable URL.
 * Payload: { _type:'uploadPhoto', _token, base64, mimeType, fileName, category }
 * category: 'recap' | 'archive' | 'qr' | 'proof'
 * Coordinator token required for 'recap' and 'archive'.
 */
function doPostUploadPhoto(data) {
  // Determine folder
  var folderId;
  if (data.category === 'recap') {
    if (data._token !== CONFIG.COORDINATOR_TOKEN && data._token !== CONFIG.ADMIN_TOKEN) {
      return jsonResponse({ success: false, error: 'Unauthorized' });
    }
    folderId = CONFIG.RECAP_PHOTOS_FOLDER_ID || CONFIG.PROOF_OF_PAYMENT_FOLDER_ID;
  } else if (data.category === 'archive') {
    if (data._token !== CONFIG.COORDINATOR_TOKEN && data._token !== CONFIG.ADMIN_TOKEN) {
      return jsonResponse({ success: false, error: 'Unauthorized' });
    }
    folderId = CONFIG.PHOTO_ARCHIVE_FOLDER_ID || CONFIG.PROOF_OF_PAYMENT_FOLDER_ID;
  } else {
    folderId = CONFIG.PROOF_OF_PAYMENT_FOLDER_ID;
  }

  try {
    var raw      = data.base64 || '';
    // Strip data URL prefix if present (data:image/jpeg;base64,...)
    if (raw.indexOf(',') !== -1) raw = raw.split(',')[1];
    var decoded  = Utilities.base64Decode(raw);
    var mimeType = data.mimeType || 'image/jpeg';
    var fileName = (data.fileName || 'photo') + '_' + new Date().getTime();
    var blob     = Utilities.newBlob(decoded, mimeType, fileName);
    var folder   = DriveApp.getFolderById(folderId);
    var file     = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId   = file.getId();
    // Use thumbnail URL — works as <img src> in browsers
    var url      = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1200';
    return jsonResponse({ success: true, fileId: fileId, url: url });
  } catch(err) {
    Logger.log('doPostUploadPhoto error: ' + err.toString());
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// Note: Event photo archive is handled via Google Drive folder links stored per event.
// Coordinators create a Drive folder, paste the link into the event form (Photo Folder URL field),
// and community members upload directly to Drive via photos.html.

function saveEventQrCode(eventId, fileData) {
  var folder   = DriveApp.getFolderById(CONFIG.PROOF_OF_PAYMENT_FOLDER_ID);
  var fileName = eventId + '_qr_' + new Date().getTime() + '.' + ((fileData.name || 'qr.png').split('.').pop());
  var blob     = Utilities.newBlob(
    Utilities.base64Decode(fileData.base64),
    fileData.mimeType || 'image/png',
    fileName
  );
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?id=' + file.getId() + '&export=view';
}

function savePaymentProof(id, fileData) {
  const folder   = DriveApp.getFolderById(CONFIG.PROOF_OF_PAYMENT_FOLDER_ID);
  // Use caller-supplied name if provided (event reg sets yyyy_mm_eventname_parentname)
  const ext      = (fileData.name || 'proof').split('.').pop().toLowerCase();
  const baseName = fileData.name && fileData.name.includes('_')
                   ? fileData.name
                   : id + '_proof';
  const fileName = baseName.replace(/\.[^.]+$/, '') + '.' + ext;
  const blob     = Utilities.newBlob(
    Utilities.base64Decode(fileData.base64),
    fileData.mimeType || 'application/octet-stream',
    fileName
  );
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}




// ─────────────────────────────────────────────────────────────
//  MEMBERSHIP SHEET WRITERS
// ─────────────────────────────────────────────────────────────

function writeFamilyRow(sheet, familyId, timestamp, data, proofUrl) {
  const skills = data.skills || {};
  sheet.appendRow([
    familyId, timestamp,
    data.familyName                         || '',
    (data.address || {}).street             || '',
    (data.address || {}).barangay           || '',
    (data.address || {}).city               || '',
    (data.address || {}).province           || '',
    (data.address || {}).zipCode            || '',
    data.emergencyContactName               || '',
    data.emergencyContactPhone              || '',
    arrayToString(data.activities),
    arrayToString(data.clubs),
    arrayToString(data.volunteerRoles),
    data.volunteerOther                     || '',
    'Active',
    data.consent && data.consent.data             ? 'Yes' : 'No',
    data.consent && data.consent.photo            ? 'Yes' : 'No',
    data.consent && data.consent.fileContribution ? 'Yes' : 'No',
    timestamp,
    arrayToString(skills.creative),
    arrayToString(skills.practical),
    arrayToString(skills.people),
    arrayToString(skills.business),
    '', '', '', '',
    data.activitySuggestions  || '',
    data.growthSuggestions    || '',
    data.willingToLead        ? 'Yes' : 'No',
    proofUrl                  || '',
  ]);
}

function writeParentRows(sheet, familyId, familyName, parents) {
  if (!parents || parents.length === 0) return;
  parents.forEach(function(parent, index) {
    const parentId = familyId + '-P' + String(index + 1).padStart(2, '0');
    sheet.appendRow([
      parentId, familyId, familyName    || '',
      parent.name                       || '',
      parent.email                      || '',
      parent.phone                      || '',
      parent.relationship               || '',
      parent.preferredContact           || '',
      parent.bloodType                  || '',
      parent.occupation                 || '',
      parent.nickname                   || '',   // col K — Nickname
      parent.birthday                   || '',   // col L — Birthday (MM-DD, no year)
    ]);
  });
}

function writeChildRows(sheet, familyId, familyName, children) {
  if (!children || children.length === 0) return;
  children.forEach(function(child, index) {
    const childId = familyId + '-C' + String(index + 1).padStart(2, '0');
    sheet.appendRow([
      childId, familyId, familyName    || '',
      child.name                       || '',
      child.birthday                   || '',
      child.gradeLevel                 || '',
      child.bloodType                  || '',
      child.curriculum                 || '',
      child.medicalNotes               || '',
      child.nickname                   || '',   // col J — Nickname
    ]);
  });
}


// ─────────────────────────────────────────────────────────────
//  EMAIL FUNCTIONS
// ─────────────────────────────────────────────────────────────

function sendConfirmationEmail(primaryParent, familyName, familyId) {
  if (!primaryParent || !primaryParent.email) return;
  const subject = 'Welcome to ' + CONFIG.COMMUNITY_NAME + '! Your Registration is Confirmed';
  const body = [
    'Dear ' + (primaryParent.name || 'Parent') + ',',
    '',
    'Thank you for registering your family with ' + CONFIG.COMMUNITY_NAME + '!',
    'Your Family ID: ' + familyId,
    'Family Name: ' + familyName,
    '',
    'We are excited to have the ' + familyName + ' family as part of our community.',
    'Please allow 1–2 business days for your payment proof to be verified.',
    '',
    'With warmth,',
    'The ' + CONFIG.COMMUNITY_NAME + ' Team',
  ].join('\n');
  GmailApp.sendEmail(primaryParent.email, subject, body);
}

function sendAdminNotification(familyId, data, timestamp, proofUrl) {
  const parents  = (data.parents  || []).map(function(p) { return p.name + ' (' + p.relationship + ')'; }).join(', ');
  const children = (data.children || []).map(function(c) { return c.name + (c.gradeLevel ? ' — ' + c.gradeLevel : ''); }).join(', ');
  const subject  = '[NegOr HS Hub] New Member: ' + data.familyName + ' [' + familyId + ']';
  const body = [
    'New family registration received.',
    '',
    'Family ID  : ' + familyId,
    'Family Name: ' + data.familyName,
    'Registered : ' + timestamp.toLocaleString(),
    'City       : ' + ((data.address || {}).city || ''),
    '',
    'Parents    : ' + (parents  || '—'),
    'Children   : ' + (children || 'None'),
    '',
    'Proof of Payment: ' + (proofUrl || '(not uploaded)'),
  ].join('\n');
  GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, body);
}

function sendEventAdminNotification(regId, data, event, timestamp, proofUrl) {
  const parents    = (data.parents  || []).map(function(p) { return p.name; }).join(', ');
  const children   = (data.children || []).map(function(c) { return c.name + ' (age ' + c.age + ')'; }).join(', ');
  const eventTitle = (event && event['event_title']) || data.eventTitle || data.eventId;
  const subject    = '[NegOr HS Hub] Event Registration: ' + data.familyName + ' — ' + eventTitle + ' [' + regId + ']';
  const body = [
    'New event registration received.',
    '',
    'Registration ID : ' + regId,
    'Event           : ' + eventTitle,
    'Family          : ' + data.familyName,
    'Registered      : ' + timestamp.toLocaleString(),
    '',
    'Parents/Guardians : ' + (parents  || '—'),
    'Children          : ' + (children || 'None'),
    '',
    'Role at Event     : ' + (data.organizingRole || 'attendee'),
    'Excess Handling   : ' + (data.excessHandling === 'community-fund' ? 'Donate to Community Fund' : 'Refund'),
    'Voluntary Amount  : ₱' + (data.voluntaryAmount || '0'),
    '',
    'Notes             : ' + (data.notes || '—'),
    'Proof of Payment  : ' + (proofUrl || '(not uploaded)'),
  ].join('\n');
  GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, body);
}


// ─────────────────────────────────────────────────────────────
//  ADMIN DATA FETCHERS (membership)
// ─────────────────────────────────────────────────────────────

function getFamiliesData(ss) {
  const families = sheetToObjects(ss.getSheetByName(CONFIG.SHEET_FAMILIES));
  const parents  = sheetToObjects(ss.getSheetByName(CONFIG.SHEET_PARENTS));
  const children = sheetToObjects(ss.getSheetByName(CONFIG.SHEET_CHILDREN));

  const enriched = families.map(function(family) {
    var fid = family['family_id'];
    return Object.assign({}, family, {
      parent_count : parents.filter(function(p)  { return p['family_id'] === fid; }).length,
      child_count  : children.filter(function(c) { return c['family_id'] === fid; }).length,
      parent_names : parents.filter(function(p)  { return p['family_id'] === fid; })
                             .map(function(p)    { return p['full_name']; }).join(', '),
    });
  });
  return jsonResponse({ success: true, data: enriched });
}

function getFamilyDetail(ss, familyId) {
  if (!familyId) return jsonResponse({ success: false, error: 'familyId required' });
  const families = sheetToObjects(ss.getSheetByName(CONFIG.SHEET_FAMILIES));
  const parents  = sheetToObjects(ss.getSheetByName(CONFIG.SHEET_PARENTS));
  const children = sheetToObjects(ss.getSheetByName(CONFIG.SHEET_CHILDREN));

  const family = families.find(function(f) { return f['family_id'] === familyId; });
  if (!family) return jsonResponse({ success: false, error: 'Family not found: ' + familyId });

  return jsonResponse({ success: true, data: {
    family,
    parents:  parents.filter(function(p)  { return p['family_id'] === familyId; }),
    children: children.filter(function(c) { return c['family_id'] === familyId; }),
  }});
}

function getStatsData(ss) {
  const families = sheetToObjects(ss.getSheetByName(CONFIG.SHEET_FAMILIES));
  const now      = new Date();
  return jsonResponse({ success: true, data: {
    total        : families.length,
    active       : families.filter(function(f) { return f['status'] === 'Active'; }).length,
    photo_consent: families.filter(function(f) { return f['consent_-_photo'] === 'Yes'; }).length,
    this_month   : families.filter(function(f) {
      var d = new Date(f['registration_date']);
      return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
    }).length,
  }});
}


// ─────────────────────────────────────────────────────────────
//  UTILITIES
// ─────────────────────────────────────────────────────────────

function sheetToObjects(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(function(h) {
    return h.toString().toLowerCase().replace(/[\s()\/]+/g, '_').replace(/_+$/g, '');
  });
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      var v = row[i];
      if (v instanceof Date) {
        // Format as YYYY-MM-DD — avoids locale-dependent .toString() like "Sat Jul 18 2026 00:00:00 GMT+0800"
        if (v.getFullYear() > 1899) {
          var mm = String(v.getMonth() + 1).padStart(2, '0');
          var dd = String(v.getDate()).padStart(2, '0');
          obj[h] = v.getFullYear() + '-' + mm + '-' + dd;
        } else {
          obj[h] = '';
        }
      } else {
        obj[h] = v != null ? v.toString() : '';
      }
    });
    return obj;
  });
}

function arrayToString(arr) {
  if (!arr || !Array.isArray(arr)) return '';
  return arr.filter(Boolean).join(', ');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ─────────────────────────────────────────────────────────────
//  ONE-TIME SETUP — Membership Spreadsheet
// ─────────────────────────────────────────────────────────────

function setupSpreadsheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  var familySheet = ss.getSheetByName(CONFIG.SHEET_FAMILIES) || ss.insertSheet(CONFIG.SHEET_FAMILIES);
  familySheet.getRange(1, 1, 1, 31).setValues([[
    'Family ID','Registration Date','Family Name',
    'Street Address','Barangay','City / Municipality','Province','ZIP Code',
    'Emergency Contact Name','Emergency Contact Phone',
    'Activities Interested','Clubs / Co-ops Interested','Volunteer Roles','Volunteer - Other',
    'Status',
    'Consent - Data','Consent - Photo','Consent - File Contribution','Consent Date',
    'Skills - Creative Arts','Skills - Practical & Technical','Skills - People & Community','Skills - Business & Professional',
    'Reserved A','Reserved B','Reserved C','Reserved D',
    'Activity Suggestions','Growth Suggestions','Willing to Lead Suggestions',
    'Proof of Payment URL',
  ]]);
  familySheet.getRange(1,1,1,31).setBackground('#B5451B').setFontColor('#fff').setFontWeight('bold');
  familySheet.setFrozenRows(1);

  var parentSheet = ss.getSheetByName(CONFIG.SHEET_PARENTS) || ss.insertSheet(CONFIG.SHEET_PARENTS);
  parentSheet.getRange(1,1,1,12).setValues([[
    'Parent ID','Family ID','Family Name',
    'Full Name','Email','Phone',
    'Relationship','Preferred Contact','Blood Type','Occupation',
    'Nickname','Birthday',
  ]]);
  parentSheet.getRange(1,1,1,12).setBackground('#8B3214').setFontColor('#fff').setFontWeight('bold');
  parentSheet.setFrozenRows(1);

  var childSheet = ss.getSheetByName(CONFIG.SHEET_CHILDREN) || ss.insertSheet(CONFIG.SHEET_CHILDREN);
  childSheet.getRange(1,1,1,10).setValues([[
    'Child ID','Family ID','Family Name',
    'Full Name','Date of Birth','Grade Level',
    'Blood Type','Curriculum Approach','Medical / Allergy Notes','Nickname',
  ]]);
  childSheet.getRange(1,1,1,10).setBackground('#D4714E').setFontColor('#fff').setFontWeight('bold');
  childSheet.setFrozenRows(1);

  Logger.log('✅ Membership spreadsheet setup complete.');
  SpreadsheetApp.getUi().alert('Membership spreadsheet ready! Tabs: Families, Parents, Children.');
}


// ─────────────────────────────────────────────────────────────
//  RECAP CORNER
// ─────────────────────────────────────────────────────────────

/**
 * Public: list all recaps, newest first.
 * Each recap row has a Sections column (JSON array of {title, highlights, photoUrl}).
 */
function getRecapsData() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_RECAPS);
    if (!sheet) return jsonResponse({ success: true, data: [] });
    const rows  = sheetToObjects(sheet);
    // Parse sections JSON for each recap
    // Header "Sections (JSON)" → snake_case key "sections_json"
    const parsed = rows.map(function(r) {
      var sections = [];
      var raw = r['sections_json'] || r['sections'] || '';
      try { if (raw) sections = JSON.parse(raw); } catch(e) {}
      return Object.assign({}, r, { sections: sections });
    });
    parsed.sort(function(a,b){ return (b['recap_date']||'').localeCompare(a['recap_date']||''); });
    return jsonResponse({ success: true, data: parsed });
  } catch(err) { return jsonResponse({ success: false, error: err.toString() }); }
}

/**
 * Coordinator: create a recap with sections.
 * Payload: {
 *   _type:'createRecap', _token,
 *   eventTitle, eventDate, summary, writtenBy, attendanceCount,
 *   sections: [ { title, highlights, photoUrl } ]
 * }
 */
function doPostCreateRecap(data) {
  if (data._token !== CONFIG.COORDINATOR_TOKEN && data._token !== CONFIG.ADMIN_TOKEN) {
    return jsonResponse({ success: false, error: 'Unauthorized' });
  }
  const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_RECAPS);
  if (!sheet) return jsonResponse({ success: false, error: 'Recaps sheet not found. Run setupEventsSpreadsheet().' });

  const lastRow = sheet.getLastRow();
  const recapId = 'RCP-' + String(lastRow).padStart(3, '0');
  const now     = new Date().toISOString().slice(0, 10);

  var sectionsJson = '[]';
  try { sectionsJson = JSON.stringify(data.sections || []); } catch(e) {}

  // Columns A–H: Recap ID, Event Title, Event Date, Recap Date, Summary, Written By, Attendance Count, Sections
  sheet.appendRow([
    recapId,
    data.eventTitle      || '',
    data.eventDate       || '',
    now,
    data.summary         || '',
    data.writtenBy       || '',
    data.attendanceCount || '',
    sectionsJson,
  ]);
  return jsonResponse({ success: true, recapId });
}


// ─────────────────────────────────────────────────────────────
//  BOOK EXCHANGE CLUB
// ─────────────────────────────────────────────────────────────

function getBooksData() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_BOOKS);
    if (!sheet) return jsonResponse({ success: true, data: [] });
    return jsonResponse({ success: true, data: sheetToObjects(sheet) });
  } catch(err) { return jsonResponse({ success: false, error: err.toString() }); }
}

function doPostAddBook(data) {
  const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_BOOKS);
  if (!sheet) return jsonResponse({ success: false, error: 'Books sheet not found. Run setupEventsSpreadsheet().' });

  const lastRow = sheet.getLastRow();
  const bookId  = 'BK-' + String(lastRow).padStart(3, '0');
  const now     = new Date().toISOString().slice(0,10);
  sheet.appendRow([bookId, data.title||'', data.author||'', data.genre||'',
    data.condition||'Good', data.ownerName||'', data.ownerMessenger||'',
    '', '', '', 'Available', data.notes||'', now]);
  return jsonResponse({ success: true, bookId });
}

function doPostBorrowBook(data) {
  if (!data.bookId) return jsonResponse({ success: false, error: 'bookId required' });
  const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_BOOKS);
  if (!sheet) return jsonResponse({ success: false, error: 'Books sheet not found.' });
  const rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.bookId) {
      sheet.getRange(i+1,  8).setValue(data.borrowerName    || '');
      sheet.getRange(i+1,  9).setValue(data.borrowerContact || '');
      sheet.getRange(i+1, 10).setValue(new Date().toISOString().slice(0,10));
      sheet.getRange(i+1, 11).setValue('Borrowed');
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ success: false, error: 'Book not found.' });
}

function doPostReturnBook(data) {
  if (!data.bookId) return jsonResponse({ success: false, error: 'bookId required' });
  const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_BOOKS);
  if (!sheet) return jsonResponse({ success: false, error: 'Books sheet not found.' });
  const rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.bookId) {
      sheet.getRange(i+1, 8).setValue('');
      sheet.getRange(i+1, 9).setValue('');
      sheet.getRange(i+1,10).setValue('');
      sheet.getRange(i+1,11).setValue('Available');
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ success: false, error: 'Book not found.' });
}


function doPostEditBook(data) {
  if (!data.bookId) return jsonResponse({ success: false, error: 'bookId required' });
  const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_BOOKS);
  if (!sheet) return jsonResponse({ success: false, error: 'Books sheet not found.' });
  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.bookId) {
      if (data.title     !== undefined) sheet.getRange(i+1, 2).setValue(data.title);
      if (data.author    !== undefined) sheet.getRange(i+1, 3).setValue(data.author);
      if (data.genre     !== undefined) sheet.getRange(i+1, 4).setValue(data.genre);
      if (data.condition !== undefined) sheet.getRange(i+1, 5).setValue(data.condition);
      if (data.ownerName !== undefined) sheet.getRange(i+1, 6).setValue(data.ownerName);
      if (data.ownerMessenger !== undefined) sheet.getRange(i+1, 7).setValue(data.ownerMessenger);
      if (data.notes     !== undefined) sheet.getRange(i+1,12).setValue(data.notes);
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ success: false, error: 'Book not found.' });
}

function doPostDeleteBook(data) {
  if (!data.bookId) return jsonResponse({ success: false, error: 'bookId required' });
  const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_BOOKS);
  if (!sheet) return jsonResponse({ success: false, error: 'Books sheet not found.' });
  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.bookId) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ success: false, error: 'Book not found.' });
}

// ─────────────────────────────────────────────────────────────
//  ORGANIZER CHECKLIST
// ─────────────────────────────────────────────────────────────

function getChecklistData(eventId) {
  if (!eventId) return jsonResponse({ success: false, error: 'eventId required' });
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_CHECKLIST);
    if (!sheet) return jsonResponse({ success: true, data: [] });
    const all   = sheetToObjects(sheet);
    return jsonResponse({ success: true, data: all.filter(function(r){ return r['event_id'] === eventId; }) });
  } catch(err) { return jsonResponse({ success: false, error: err.toString() }); }
}

function doPostMarkChecklistItem(data) {
  if (data._token !== CONFIG.COORDINATOR_TOKEN && data._token !== CONFIG.ADMIN_TOKEN) {
    return jsonResponse({ success: false, error: 'Unauthorized' });
  }
  if (!data.eventId || !data.itemKey) return jsonResponse({ success: false, error: 'eventId and itemKey required' });
  const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_CHECKLIST);
  if (!sheet) return jsonResponse({ success: false, error: 'Checklist sheet not found. Run setupEventsSpreadsheet().' });

  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.eventId && rows[i][1] === data.itemKey) {
      sheet.getRange(i+1, 3).setValue(data.checked ? 'TRUE' : 'FALSE');
      sheet.getRange(i+1, 4).setValue(data.checked ? new Date().toISOString() : '');
      sheet.getRange(i+1, 5).setValue(data.checkedBy || '');
      return jsonResponse({ success: true });
    }
  }
  // New row
  sheet.appendRow([data.eventId, data.itemKey, data.checked ? 'TRUE' : 'FALSE',
    data.checked ? new Date().toISOString() : '', data.checkedBy || '']);
  return jsonResponse({ success: true });
}


// ─────────────────────────────────────────────────────────────
//  BIRTHDAYS — public, returns nicknames + day for current month
//  Includes both parents (MM-DD birthday, no age) and children
//  (full date of birth, age calculated).
// ─────────────────────────────────────────────────────────────

function getBirthdaysData() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const now   = new Date();
    const month = now.getMonth() + 1;
    const year  = now.getFullYear();
    const bdays = [];

    // ── Children: full date of birth → age available ──────────
    var childSheet = ss.getSheetByName(CONFIG.SHEET_CHILDREN);
    if (childSheet) {
      sheetToObjects(childSheet).forEach(function(child) {
        var dob = child['date_of_birth'];
        if (!dob) return;
        try {
          var d = new Date(dob);
          if (isNaN(d.getTime())) return;
          if (d.getMonth() + 1 !== month) return;
          var displayName = (child['nickname'] || '').trim()
                         || (child['full_name'] || '').split(' ')[0]
                         || 'A child';
          var age = year - d.getFullYear();
          bdays.push({ name: displayName, day: d.getDate(), age: age, type: 'child' });
        } catch(e) {}
      });
    }

    // ── Parents: MM-DD birthday only, no age ──────────────────
    var parentSheet = ss.getSheetByName(CONFIG.SHEET_PARENTS);
    if (parentSheet) {
      sheetToObjects(parentSheet).forEach(function(parent) {
        // Header is "Birthday" → key "birthday"
        var bday = (parent['birthday'] || '').trim();
        if (!bday) return;
        // Accept both MM-DD and MM/DD formats
        var parts = bday.replace('/', '-').split('-');
        if (parts.length < 2) return;
        var mm = parseInt(parts[0], 10);
        var dd = parseInt(parts[1], 10);
        if (isNaN(mm) || isNaN(dd)) return;
        if (mm !== month) return;
        var displayName = (parent['nickname'] || '').trim()
                       || (parent['full_name'] || '').split(' ')[0]
                       || 'A parent';
        bdays.push({ name: displayName, day: dd, type: 'parent' });
      });
    }

    bdays.sort(function(a, b) { return a.day - b.day; });
    return jsonResponse({ success: true, data: bdays });
  } catch(err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}


// ─────────────────────────────────────────────────────────────
//  MIGRATION — run ONCE on existing deployments to add new cols
// ─────────────────────────────────────────────────────────────

/**
 * addBirthdayNicknameColumns()
 * Run this ONCE in the Apps Script editor if you already have
 * data in your Parents / Children sheets.
 * It appends the new header columns without touching existing rows.
 */
function addBirthdayNicknameColumns() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Parents sheet: add Nickname (col K) and Birthday MM-DD (col L)
  var pSheet = ss.getSheetByName(CONFIG.SHEET_PARENTS);
  if (pSheet) {
    var pLastCol = pSheet.getLastColumn();
    if (pLastCol < 11) {
      pSheet.getRange(1, 11).setValue('Nickname');
      pSheet.getRange(1, 11).setBackground('#8B3214').setFontColor('#fff').setFontWeight('bold');
    }
    if (pLastCol < 12) {
      pSheet.getRange(1, 12).setValue('Birthday');
      pSheet.getRange(1, 12).setBackground('#8B3214').setFontColor('#fff').setFontWeight('bold');
    }
    Logger.log('Parents sheet updated.');
  }

  // Children sheet: add Nickname (col J)
  var cSheet = ss.getSheetByName(CONFIG.SHEET_CHILDREN);
  if (cSheet) {
    var cLastCol = cSheet.getLastColumn();
    if (cLastCol < 10) {
      cSheet.getRange(1, 10).setValue('Nickname');
      cSheet.getRange(1, 10).setBackground('#D4714E').setFontColor('#fff').setFontWeight('bold');
    }
    Logger.log('Children sheet updated.');
  }

  Logger.log('Migration complete. No existing data was changed.');
}


// ─────────────────────────────────────────────────────────────
//  RESOURCE REPOSITORY — Venues, Equipment, Activity Log
// ─────────────────────────────────────────────────────────────

// ── Venues ───────────────────────────────────────────────────

function getVenuesData() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_VENUES);
    if (!sheet) return jsonResponse({ success: true, data: [] });
    return jsonResponse({ success: true, data: sheetToObjects(sheet) });
  } catch(err) { return jsonResponse({ success: false, error: err.toString() }); }
}

function doPostAddVenue(data) {
  // Open to all community members — no token required
  const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_VENUES);
  if (!sheet) return jsonResponse({ success: false, error: 'Venues sheet not found. Run addResourcesSheets().' });
  const lastRow = sheet.getLastRow();
  const venueId = 'VEN-' + String(lastRow).padStart(3, '0');
  const now     = new Date().toISOString().slice(0, 10);
  // A–I: Venue ID, Name, Address, Type, Capacity, Contact Person, Contact Number, Notes, Status
  sheet.appendRow([
    venueId,
    data.name            || '',
    data.address         || '',
    data.type            || 'Indoor',
    data.capacity        || '',
    data.contactPerson   || '',
    data.contactNumber   || '',
    data.notes           || '',
    'Active',
    now,
  ]);
  return jsonResponse({ success: true, venueId: venueId });
}

// ── Equipment ─────────────────────────────────────────────────

function getEquipmentData() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_EQUIPMENT);
    if (!sheet) return jsonResponse({ success: true, data: [] });
    return jsonResponse({ success: true, data: sheetToObjects(sheet) });
  } catch(err) { return jsonResponse({ success: false, error: err.toString() }); }
}

function doPostAddEquipment(data) {
  // Open to all community members — no token required
  const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_EQUIPMENT);
  if (!sheet) return jsonResponse({ success: false, error: 'Equipment sheet not found. Run addResourcesSheets().' });
  const lastRow = sheet.getLastRow();
  const eqId    = 'EQ-' + String(lastRow).padStart(3, '0');
  const now     = new Date().toISOString().slice(0, 10);
  // A–I: Equipment ID, Item Name, Category, Quantity, Condition, Custodian, Contact, Notes, Date Added
  sheet.appendRow([
    eqId,
    data.itemName    || '',
    data.category    || 'Supplies',
    data.quantity    || 1,
    data.condition   || 'Good',
    data.custodian   || '',
    data.contact     || '',
    data.notes       || '',
    now,
  ]);
  return jsonResponse({ success: true, equipmentId: eqId });
}

// ── Activity Log ──────────────────────────────────────────────

function getActivityLogData() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_ACTIVITIES);
    if (!sheet) return jsonResponse({ success: true, data: [] });
    const rows  = sheetToObjects(sheet);
    rows.sort(function(a,b){ return (b['date_done']||'').localeCompare(a['date_done']||''); });
    return jsonResponse({ success: true, data: rows });
  } catch(err) { return jsonResponse({ success: false, error: err.toString() }); }
}

function doPostAddActivity(data) {
  // Open to all community members — no token required
  const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_ACTIVITIES);
  if (!sheet) return jsonResponse({ success: false, error: 'ActivityLog sheet not found. Run addResourcesSheets().' });
  const lastRow = sheet.getLastRow();
  const actId   = 'ACT-' + String(lastRow).padStart(3, '0');
  // A–H: Activity ID, Name, Category, Date Done, Venue Used, Attendance, Notes, Logged By
  sheet.appendRow([
    actId,
    data.activityName  || '',
    data.category      || 'Other',
    data.dateDone      || '',
    data.venueUsed     || '',
    data.attendance    || '',
    data.notes         || '',
    data.loggedBy      || '',
  ]);
  return jsonResponse({ success: true, activityId: actId });
}


// ─────────────────────────────────────────────────────────────
//  ONE-TIME SETUP — Events Spreadsheet (separate Google Sheet)
// ─────────────────────────────────────────────────────────────

function setupEventsSpreadsheet() {
  const ss = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);

  // Events  A–O (15 cols)
  var evSheet = ss.getSheetByName(CONFIG.SHEET_EVENTS) || ss.insertSheet(CONFIG.SHEET_EVENTS);
  evSheet.getRange(1,1,1,15).setValues([[
    'Event ID','Created Date','Event Title','Event Date','Venue',
    'Fee (₱)','Payee Name','Description','Status','Registration Count',
    'Start Time','End Time','Registration Type','QR Code URL','Photo Folder URL',
  ]]);
  evSheet.getRange(1,1,1,15).setBackground('#d97706').setFontColor('#fff').setFontWeight('bold');
  evSheet.setFrozenRows(1);

  // Registrations  A–O (15 cols)
  var rgSheet = ss.getSheetByName(CONFIG.SHEET_EVENT_REGS) || ss.insertSheet(CONFIG.SHEET_EVENT_REGS);
  rgSheet.getRange(1,1,1,15).setValues([[
    'Registration ID','Event ID','Event Title','Timestamp',
    'Family Name','Parents / Guardians','Children',
    'Excess Handling','Voluntary Amount (₱)','Voluntary Destination',
    'Role at Event','Notes','Proof of Payment URL','Data Consent','Payment Status',
  ]]);
  rgSheet.getRange(1,1,1,15).setBackground('#B5451B').setFontColor('#fff').setFontWeight('bold');
  rgSheet.setFrozenRows(1);

  // Recaps  A–H (8 cols) — redesigned v3: sections replace highlights+photo_link
  var rcSheet = ss.getSheetByName(CONFIG.SHEET_RECAPS) || ss.insertSheet(CONFIG.SHEET_RECAPS);
  rcSheet.getRange(1,1,1,8).setValues([[
    'Recap ID','Event Title','Event Date','Recap Date',
    'Summary','Written By','Attendance Count','Sections (JSON)',
  ]]);
  rcSheet.getRange(1,1,1,8).setBackground('#7c3aed').setFontColor('#fff').setFontWeight('bold');
  rcSheet.setFrozenRows(1);

  // Books  A–M (13 cols)
  var bkSheet = ss.getSheetByName(CONFIG.SHEET_BOOKS) || ss.insertSheet(CONFIG.SHEET_BOOKS);
  bkSheet.getRange(1,1,1,13).setValues([[
    'Book ID','Title','Author','Genre','Condition',
    'Owner Name','Owner Messenger','Current Holder','Holder Messenger','Date Borrowed',
    'Status','Notes','Date Added',
  ]]);
  bkSheet.getRange(1,1,1,13).setBackground('#059669').setFontColor('#fff').setFontWeight('bold');
  bkSheet.setFrozenRows(1);

  // Checklist  A–E (5 cols)
  var clSheet = ss.getSheetByName(CONFIG.SHEET_CHECKLIST) || ss.insertSheet(CONFIG.SHEET_CHECKLIST);
  clSheet.getRange(1,1,1,5).setValues([['Event ID','Item Key','Checked','Checked Date','Checked By']]);
  clSheet.getRange(1,1,1,5).setBackground('#0369a1').setFontColor('#fff').setFontWeight('bold');
  clSheet.setFrozenRows(1);

  // Venues  A–J (10 cols)
  var vnSheet = ss.getSheetByName(CONFIG.SHEET_VENUES) || ss.insertSheet(CONFIG.SHEET_VENUES);
  vnSheet.getRange(1,1,1,10).setValues([[
    'Venue ID','Name','Address','Type','Capacity',
    'Contact Person','Contact Number','Notes','Status','Date Added',
  ]]);
  vnSheet.getRange(1,1,1,10).setBackground('#ea580c').setFontColor('#fff').setFontWeight('bold');
  vnSheet.setFrozenRows(1);

  // Equipment  A–I (9 cols)
  var eqSheet = ss.getSheetByName(CONFIG.SHEET_EQUIPMENT) || ss.insertSheet(CONFIG.SHEET_EQUIPMENT);
  eqSheet.getRange(1,1,1,9).setValues([[
    'Equipment ID','Item Name','Category','Quantity','Condition',
    'Custodian','Contact','Notes','Date Added',
  ]]);
  eqSheet.getRange(1,1,1,9).setBackground('#c2410c').setFontColor('#fff').setFontWeight('bold');
  eqSheet.setFrozenRows(1);

  // Activity Log  A–H (8 cols)
  var alSheet = ss.getSheetByName(CONFIG.SHEET_ACTIVITIES) || ss.insertSheet(CONFIG.SHEET_ACTIVITIES);
  alSheet.getRange(1,1,1,8).setValues([[
    'Activity ID','Activity Name','Category','Date Done',
    'Venue Used','Attendance','Notes','Logged By',
  ]]);
  alSheet.getRange(1,1,1,8).setBackground('#9a3412').setFontColor('#fff').setFontWeight('bold');
  alSheet.setFrozenRows(1);

  Logger.log('✅ Events spreadsheet setup complete.');
  SpreadsheetApp.getUi().alert('Events spreadsheet ready! Tabs: Events, Registrations, Recaps, Books, Checklist, Venues, Equipment, ActivityLog.');
}


// ─────────────────────────────────────────────────────────────
//  MIGRATIONS — run once if upgrading from a previous version
// ─────────────────────────────────────────────────────────────

/**
 * Add "Payment Status" column O to existing Registrations sheet.
 * Run once if your sheet was set up before v2.0.
 */
function updateRegsSheet() {
  const ss      = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const rgSheet = ss.getSheetByName(CONFIG.SHEET_EVENT_REGS);
  if (!rgSheet) { SpreadsheetApp.getUi().alert('Registrations sheet not found.'); return; }
  const headers = rgSheet.getRange(1, 1, 1, rgSheet.getLastColumn()).getValues()[0];
  if (headers.includes('Payment Status')) {
    SpreadsheetApp.getUi().alert('Payment Status column already exists — nothing to do.');
    return;
  }
  const colO = rgSheet.getLastColumn() + 1;
  rgSheet.getRange(1, colO).setValue('Payment Status');
  rgSheet.getRange(1, colO).setBackground('#B5451B').setFontColor('#fff').setFontWeight('bold');
  SpreadsheetApp.getUi().alert('✅ Payment Status column added to Registrations sheet!');
}

/**
 * Add 4 new columns to existing Events sheet (Start Time, End Time, Registration Type, QR Code URL).
 * Run once if your sheet was set up before v2.0.
 */
function updateEventsSheet() {
  const ss      = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const evSheet = ss.getSheetByName(CONFIG.SHEET_EVENTS);
  if (!evSheet) { SpreadsheetApp.getUi().alert('Events sheet not found. Run setupEventsSpreadsheet() first.'); return; }
  const newCols   = ['Start Time', 'End Time', 'Registration Type', 'QR Code URL'];
  const headerRow = evSheet.getRange(1, 1, 1, Math.max(evSheet.getLastColumn(), 14)).getValues()[0];
  newCols.forEach(function(name, i) {
    const col = 11 + i;
    if (!headerRow[col - 1] || headerRow[col - 1].toString().trim() === '') {
      const cell = evSheet.getRange(1, col);
      cell.setValue(name);
      cell.setBackground('#d97706').setFontColor('#fff').setFontWeight('bold');
    }
  });
  SpreadsheetApp.getUi().alert('Done! Columns added: Start Time, End Time, Registration Type, QR Code URL.');
}

/**
 * Add "Sections (JSON)" column H to existing Recaps sheet + rename col E to Summary.
 * Run once if you have an existing Recaps sheet from v2.0.
 */
function updateRecapsSheet() {
  const ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_RECAPS);
  if (!sheet) { SpreadsheetApp.getUi().alert('Recaps sheet not found.'); return; }

  // Rename col E header from "Highlights" to "Summary"
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  if (headers[4] && headers[4].toString() === 'Highlights') {
    sheet.getRange(1,5).setValue('Summary');
  }
  // Drop col F (Photo Link) and add Sections at col H — just add Sections at end if not present
  if (!headers.includes('Sections (JSON)')) {
    const col = sheet.getLastColumn() + 1;
    var cell  = sheet.getRange(1, col);
    cell.setValue('Sections (JSON)');
    cell.setBackground('#7c3aed').setFontColor('#fff').setFontWeight('bold');
  }
  SpreadsheetApp.getUi().alert('✅ Recaps sheet updated with Sections column.');
}

/**
 * Add "Photo Folder URL" column O to existing Events sheet.
 * Run once if upgrading from v2.0 (which had only 14 cols).
 */
function updateEventsSheetV3() {
  const ss      = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  const evSheet = ss.getSheetByName(CONFIG.SHEET_EVENTS);
  if (!evSheet) { SpreadsheetApp.getUi().alert('Events sheet not found. Run setupEventsSpreadsheet() first.'); return; }
  const headers = evSheet.getRange(1,1,1,Math.max(evSheet.getLastColumn(),15)).getValues()[0];
  if (headers[14] && headers[14].toString().trim() !== '') {
    SpreadsheetApp.getUi().alert('Photo Folder URL column already exists — nothing to do.');
    return;
  }
  var cell = evSheet.getRange(1, 15);
  cell.setValue('Photo Folder URL');
  cell.setBackground('#d97706').setFontColor('#fff').setFontWeight('bold');
  SpreadsheetApp.getUi().alert('✅ Photo Folder URL column added to Events sheet (col O).');
}

/**
 * Rename "Owner Contact" → "Owner Messenger" and "Borrower Contact" → "Borrower Messenger"
 * in the Books sheet. Run once if your sheet was set up before v3.1.
 */
function updateBooksSheetV31() {
  var ss    = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_BOOKS);
  if (!sheet) { SpreadsheetApp.getUi().alert('Books sheet not found.'); return; }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var changed = false;
  headers.forEach(function(h, i) {
    if (h === 'Owner Contact')  { sheet.getRange(1, i+1).setValue('Owner Messenger');  changed = true; }
    if (h === 'Borrower Contact') { sheet.getRange(1, i+1).setValue('Borrower Messenger'); changed = true; }
  });
  SpreadsheetApp.getUi().alert(changed ? '✅ Books sheet headers updated!' : 'Headers already up to date — nothing to do.');
}

/**
 * Add Venues, Equipment, and ActivityLog sheets to the Events spreadsheet.
 * Run once if upgrading from v3.0 (which had only 5 module tabs).
 */
function addResourcesSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.EVENTS_SPREADSHEET_ID);

  function ensureSheet(name, headers, color) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1,1,1,headers.length).setValues([headers]);
      sheet.getRange(1,1,1,headers.length).setBackground(color).setFontColor('#fff').setFontWeight('bold');
      sheet.setFrozenRows(1);
      return 'Created';
    }
    return 'Already exists';
  }

  var v = ensureSheet(CONFIG.SHEET_VENUES,
    ['Venue ID','Name','Address','Type','Capacity','Contact Person','Contact Number','Notes','Status','Date Added'],
    '#ea580c');
  var e = ensureSheet(CONFIG.SHEET_EQUIPMENT,
    ['Equipment ID','Item Name','Category','Quantity','Condition','Custodian','Contact','Notes','Date Added'],
    '#c2410c');
  var a = ensureSheet(CONFIG.SHEET_ACTIVITIES,
    ['Activity ID','Activity Name','Category','Date Done','Venue Used','Attendance','Notes','Logged By'],
    '#9a3412');

  SpreadsheetApp.getUi().alert(
    'Resource Repository sheets:\n' +
    '• Venues: ' + v + '\n' +
    '• Equipment: ' + e + '\n' +
    '• ActivityLog: ' + a
  );
}


// ─────────────────────────────────────────────────────────────
//  ADMIN TOKEN UTILITY
// ─────────────────────────────────────────────────────────────

function setAdminToken() {
  var token = 'REPLACE_WITH_YOUR_SECURE_TOKEN';
  PropertiesService.getScriptProperties().setProperty('ADMIN_TOKEN', token);
  Logger.log('Admin token saved.');
}
