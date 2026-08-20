// ============================================================
// EHS Dog Log — Code.gs
//
// SETUP INSTRUCTIONS:
//   1. Paste this file into a new standalone Apps Script project.
//   2. Paste index.html as a second file named "index".
//   3. Run runSetup() from the editor (Run → Run function → runSetup).
//   4. Open Execution Log and copy the Spreadsheet ID printed there.
//   5. Paste it into SHEET_ID below. Save.
//   6. Deploy as web app:
//        Execute as:  "Me" (the account that owns this script)
//        Who has access: "Anyone, even anonymous"
//   7. Update ADMINS with the correct email addresses.
// ============================================================

const SHEET_ID = "12FHdWA-5u6nNlvvsNu9K0xYqsWVYRHRHHezMkoVvYk8";

// Google Group whose members receive view-only access.
// Must match the group email exactly.
const VOLUNTEER_GROUP_EMAIL = "dog.volunteers.list@etobicokehumanesociety.com";

// Any @EHS-domain email also receives view-only access.
const EHS_DOMAIN = "etobicokehumanesociety.com";

const ADMINS = [
  "dog.volunteer.scheduler@etobicokehumanesociety.com",
  "dog.senior.manager@etobicokehumanesociety.com"
];

const SHIFTS = [
  { label: "Morning",   key: "AM"  },
  { label: "Afternoon", key: "AFT" },
  { label: "Evening",   key: "EVE" }
];

const PROFILES_TAB         = "Dog Profiles";
const PROFILE_HISTORY_TAB  = "Profile History";
const LOG_TAB              = "Dog Log";
const ALERTS_TAB           = "Shift Alerts";
const USERS_TAB            = "Users";
const RESPITE_TAB          = "Respite Schedule";

// 0-based column indices — Users
const UC = { EMAIL: 0, ROLE: 1, NAME: 2 };
const USER_HEADERS = ["Email", "Role (Admin / Regular)", "Display Name"];

// 0-based column indices — Dog Profiles
const PC = {
  DOG_NAME:            0,
  STATUS:              1,
  BREED:               2,
  AGE:                 3,
  WEIGHT:              4,
  SEX:                 5,
  ARRIVAL_DATE:        6,
  ADOPTION_DATE:       7,
  COLOUR:              8,
  COLOUR_REASON:       9,
  BREAKFAST:           10,
  LUNCH:               11,
  DINNER:              12,
  MEDICATION_REQUIRED: 13,
  MED_B:               14,
  MED_L:               15,
  MED_D:               16,
  HISTORY:             17,
  FEEDING_NOTES:       18,
  MEDICAL:             19,
  BEHAVIOUR:           20,
  KENNEL:              21,
  WALKS:               22,
  WALKS_ALLOWED:       23,
  BACKYARD:            24,
  LIKES:               25,
  DISLIKES:            26,
  STATUS_CHANGED_AT:   27,
  PHOTO_URL:           28
};

// 0-based column indices — Shift Alerts
const ALC = {
  MESSAGE:      0,
  CREATED_BY:   1,
  CREATED_AT:   2,
  EXPIRES_ON:   3,
  ACTIVE:       4,
  CATEGORY:     5,
  SHIFT_FILTER: 6,
  VISIBLE_FROM: 7,
  HANDOFF:      8
};

// 0-based column indices — Dog Log
const LC = {
  DATE:           0,
  SHIFT:          1,
  DOG_NAME:       2,
  BM:             3,
  MED:            4,
  FOOD:           5,
  WALK:           6,
  COMMENTS:       7,
  LAST_EDITED_BY: 8,
  LAST_EDITED_AT: 9
};

const PROFILE_HEADERS = [
  // Basic info (0–6)
  "Dog Name", "Status", "Breed", "Age", "Weight", "Sex", "Arrival Date",
  // Scheduling (7)
  "Adoption Date",
  // Colour / safety (8–9)
  "Colour", "Colour Reason",
  // Feeding & meds (10–16)
  "Breakfast", "Lunch", "Dinner",
  "Medication Required", "Med (Breakfast)", "Med (Lunch)", "Med (Dinner)",
  // Long-form notes (17–26)
  "History", "Feeding Notes", "Medical", "Behaviour & Handling",
  "Kennel", "Walks", "Walks Allowed", "Backyard", "Likes", "Dislikes",
  // System (27–28)
  "Status Changed At", "Photo URL"
];

// 0-based column indices — Respite Schedule
const RSC = { DOG_NAME: 0, START: 1, END: 2, NOTES: 3, RETURNED_EARLY: 4, RETURNED_AT: 5 };
const RESPITE_HEADERS = ["Dog Name", "Start", "End", "Notes", "Returned Early", "Returned At"];

const ALERT_HEADERS = ["Message", "Created By", "Created At", "Expires On", "Active", "Category", "Shift Filter", "Visible From", "Handoff"];

const LOG_HEADERS = [
  "Date", "Shift", "Dog Name", "BM", "Med", "Food", "Walk",
  "Comments", "Last Edited By", "Last Edited At"
];


// ------------------------------------------------------------
// runSetup — one-time scaffold. Run from the Apps Script editor.
// ------------------------------------------------------------
function runSetup() {
  var ss;

  if (SHEET_ID) {
    ss = SpreadsheetApp.openById(SHEET_ID);
    Logger.log("Using existing spreadsheet: " + ss.getUrl());
  } else {
    ss = SpreadsheetApp.create("EHS Dog Log");
    Logger.log("Created new spreadsheet: " + ss.getUrl());
    Logger.log("Set SHEET_ID in Code.gs to: " + ss.getId());
  }

  // Dog Profiles tab
  var profilesSheet = ss.getSheetByName(PROFILES_TAB);
  if (!profilesSheet) {
    profilesSheet = ss.insertSheet(PROFILES_TAB);
    Logger.log("Created tab: " + PROFILES_TAB);
  } else {
    Logger.log("Tab already exists, skipping: " + PROFILES_TAB);
  }
  if (profilesSheet.getLastRow() === 0) {
    profilesSheet
      .getRange(1, 1, 1, PROFILE_HEADERS.length)
      .setValues([PROFILE_HEADERS])
      .setFontWeight("bold")
      .setBackground("#e8f0eb");
    profilesSheet.setFrozenRows(1);
    profilesSheet.setColumnWidth(PC.DOG_NAME + 1, 140);
    profilesSheet.setColumnWidth(PC.STATUS + 1, 90);
    [PC.HISTORY, PC.WALKS, PC.BACKYARD, PC.MEDICAL, PC.BEHAVIOUR].forEach(function(col) {
      profilesSheet.setColumnWidth(col + 1, 280);
    });
  }

  // Dog Log tab
  var logSheet = ss.getSheetByName(LOG_TAB);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_TAB);
    Logger.log("Created tab: " + LOG_TAB);
  } else {
    Logger.log("Tab already exists, skipping: " + LOG_TAB);
  }
  if (logSheet.getLastRow() === 0) {
    logSheet
      .getRange(1, 1, 1, LOG_HEADERS.length)
      .setValues([LOG_HEADERS])
      .setFontWeight("bold")
      .setBackground("#e8f0eb");
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidth(LC.COMMENTS + 1, 320);
  }

  // Shift Alerts tab
  var alertsSheet = ss.getSheetByName(ALERTS_TAB);
  if (!alertsSheet) {
    alertsSheet = ss.insertSheet(ALERTS_TAB);
    Logger.log("Created tab: " + ALERTS_TAB);
  } else {
    Logger.log("Tab already exists, skipping: " + ALERTS_TAB);
  }
  if (alertsSheet.getLastRow() === 0) {
    alertsSheet
      .getRange(1, 1, 1, ALERT_HEADERS.length)
      .setValues([ALERT_HEADERS])
      .setFontWeight("bold")
      .setBackground("#fef9c3");
    alertsSheet.setFrozenRows(1);
    alertsSheet.setColumnWidth(ALC.MESSAGE + 1, 400);
    alertsSheet.setColumnWidth(ALC.EXPIRES_ON + 1, 120);
  }

  // Profile History tab
  var histSheet = ss.getSheetByName(PROFILE_HISTORY_TAB);
  if (!histSheet) {
    histSheet = ss.insertSheet(PROFILE_HISTORY_TAB);
    Logger.log("Created tab: " + PROFILE_HISTORY_TAB);
  } else {
    Logger.log("Tab already exists, skipping: " + PROFILE_HISTORY_TAB);
  }
  if (histSheet.getLastRow() === 0) {
    var histHeaders = ["Timestamp", "Changed By"].concat(PROFILE_HEADERS);
    histSheet
      .getRange(1, 1, 1, histHeaders.length)
      .setValues([histHeaders])
      .setFontWeight("bold")
      .setBackground("#ede9fe");
    histSheet.setFrozenRows(1);
    histSheet.setColumnWidth(1, 160);
    histSheet.setColumnWidth(2, 220);
    histSheet.setColumnWidth(3, 140);
  }

  // Users tab
  var usersSheet = ss.getSheetByName(USERS_TAB);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(USERS_TAB);
    Logger.log("Created tab: " + USERS_TAB);
  } else {
    Logger.log("Tab already exists, skipping: " + USERS_TAB);
  }
  if (usersSheet.getLastRow() === 0) {
    usersSheet
      .getRange(1, 1, 1, USER_HEADERS.length)
      .setValues([USER_HEADERS])
      .setFontWeight("bold")
      .setBackground("#dbeafe");
    usersSheet.setFrozenRows(1);
    usersSheet.setColumnWidth(UC.EMAIL + 1, 280);
    usersSheet.setColumnWidth(UC.ROLE + 1, 160);
    usersSheet.setColumnWidth(UC.NAME + 1, 200);
    ADMINS.forEach(function(email) {
      usersSheet.appendRow([email, "Admin", ""]);
    });
  }

  // Respite Schedule tab
  var respiteSheet = ss.getSheetByName(RESPITE_TAB);
  if (!respiteSheet) {
    respiteSheet = ss.insertSheet(RESPITE_TAB);
    Logger.log("Created tab: " + RESPITE_TAB);
  } else {
    Logger.log("Tab already exists, skipping: " + RESPITE_TAB);
  }
  if (respiteSheet.getLastRow() === 0) {
    respiteSheet
      .getRange(1, 1, 1, RESPITE_HEADERS.length)
      .setValues([RESPITE_HEADERS])
      .setFontWeight("bold")
      .setBackground("#fce7f3");
    respiteSheet.setFrozenRows(1);
    respiteSheet.setColumnWidth(RSC.DOG_NAME + 1, 160);
    respiteSheet.setColumnWidth(RSC.START + 1, 160);
    respiteSheet.setColumnWidth(RSC.END + 1, 160);
    respiteSheet.setColumnWidth(RSC.NOTES + 1, 280);
  }

  Logger.log("=".repeat(60));
  Logger.log("runSetup complete! Check the spreadsheet:");
  Logger.log(ss.getUrl());
  Logger.log("=".repeat(60));
}


// ------------------------------------------------------------
// Private helpers
// ------------------------------------------------------------

function getSpreadsheet_() {
  if (!SHEET_ID) {
    throw new Error("SHEET_ID is not configured. Run runSetup() first.");
  }
  return SpreadsheetApp.openById(SHEET_ID);
}

// Role resolution by email — replaces the old Session-based getUserRole_().
// Caches results in ScriptCache (5-min TTL) keyed by email.
function resolveRoleByEmail_(email) {
  if (!email) return "none";
  email = String(email).trim().toLowerCase();

  var cache    = CacheService.getScriptCache();
  var cacheKey = "role_" + email;
  var cached   = cache.get(cacheKey);
  if (cached) return cached; // empty string = stale/blank role, re-evaluate

  try {
    var sheet = getSpreadsheet_().getSheetByName(USERS_TAB);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][UC.EMAIL]).trim().toLowerCase() === email) {
          var role = String(data[i][UC.ROLE]).trim();
          if (role) { // skip rows with blank role — fall through to ADMINS/domain checks
            cache.put(cacheKey, role, 300);
            return role;
          }
        }
      }
    }
  } catch (e) { /* fall through */ }

  if (ADMINS.map(function(a) { return a.toLowerCase(); }).indexOf(email) !== -1) {
    cache.put(cacheKey, "Admin", 300);
    return "Admin";
  }

  if (VOLUNTEER_GROUP_EMAIL) {
    try {
      if (GroupsApp.getGroup(VOLUNTEER_GROUP_EMAIL).hasMember(email)) {
        cache.put(cacheKey, "View", 300);
        return "View";
      }
    } catch (e) { /* group not accessible */ }
  }

  if (EHS_DOMAIN && email.endsWith("@" + EHS_DOMAIN.toLowerCase())) {
    cache.put(cacheKey, "View", 300);
    return "View";
  }

  cache.put(cacheKey, "none", 300);
  return "none";
}

function isAdminEmail_(email)         { return resolveRoleByEmail_(email) === "Admin"; }
function isAlertAccessEmail_(email)   { return resolveRoleByEmail_(email) === "Alert Access"; }
function canManageAlertsEmail_(email) { return isAdminEmail_(email) || isAlertAccessEmail_(email); }

function requireAdminEmail_(email) {
  if (!isAdminEmail_(email)) throw new Error("This action requires admin access.");
}

function requireAlertManagerEmail_(email) {
  if (!canManageAlertsEmail_(email)) throw new Error("This action requires Alert Access or Admin.");
}

function rowDate_(row) {
  var val = row[LC.DATE];
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(val || "");
}

function profileDate_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  }
  return String(val || "");
}

function rowToProfile_(row) {
  return {
    dogName:     String(row[PC.DOG_NAME]     || ""),
    status:      String(row[PC.STATUS]       || ""),
    breed:       String(row[PC.BREED]        || ""),
    age:         String(row[PC.AGE]          || ""),
    weight:      String(row[PC.WEIGHT]       || ""),
    sex:         String(row[PC.SEX]          || ""),
    arrivalDate: profileDate_(row[PC.ARRIVAL_DATE]),
    history:     String(row[PC.HISTORY]      || ""),
    feedingNotes:        String(row[PC.FEEDING_NOTES] || ""),
    medical:     String(row[PC.MEDICAL]      || ""),
    walks:       String(row[PC.WALKS]        || ""),
    backyard:    String(row[PC.BACKYARD]     || ""),
    behaviour:   String(row[PC.BEHAVIOUR]    || ""),
    likes:       String(row[PC.LIKES]        || ""),
    dislikes:    String(row[PC.DISLIKES]     || ""),
    breakfast:   String(row[PC.BREAKFAST]    || ""),
    lunch:       String(row[PC.LUNCH]        || ""),
    dinner:              String(row[PC.DINNER]              || ""),
    colour:              String(row[PC.COLOUR]              || ""),
    colourReason:        String(row[PC.COLOUR_REASON]       || ""),
    walksAllowed:        row[PC.WALKS_ALLOWED]       === "Y",
    medicationRequired:  row[PC.MEDICATION_REQUIRED] === "Y",
    medBreakfast:        String(row[PC.MED_B] || ""),
    medLunch:            String(row[PC.MED_L] || ""),
    medDinner:           String(row[PC.MED_D] || ""),
    statusChangedAt:     profileDate_(row[PC.STATUS_CHANGED_AT]),
    adoptionDate:        profileDate_(row[PC.ADOPTION_DATE]),
    kennel:              String(row[PC.KENNEL]       || ""),
    photoUrl:            String(row[PC.PHOTO_URL]   || "")
  };
}

function rowToLogEntry_(row) {
  var lastAt = row[LC.LAST_EDITED_AT];
  return {
    date:         rowDate_(row),
    shift:        String(row[LC.SHIFT]          || ""),
    dogName:      String(row[LC.DOG_NAME]        || ""),
    bm:           (row[LC.BM] !== "" && row[LC.BM] !== null) ? String(row[LC.BM]) : "",
    med:          String(row[LC.MED]             || ""),
    food:         String(row[LC.FOOD]            || ""),
    walk:         String(row[LC.WALK]            || ""),
    comments:     String(row[LC.COMMENTS]        || ""),
    lastEditedBy: String(row[LC.LAST_EDITED_BY]  || ""),
    lastEditedAt: lastAt instanceof Date ? lastAt.toISOString() : String(lastAt || "")
  };
}

function normBm_(bm) {
  var n = parseFloat(bm);
  return (!isNaN(n) && n > 0) ? n : "";
}


// ------------------------------------------------------------
// doGet — renders the page; handles magic-link token auth
// ------------------------------------------------------------
function doGet(e) {
  var tz  = Session.getScriptTimeZone();
  var now = new Date();
  var todayStr       = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  var todayFormatted = Utilities.formatDate(now, tz, "EEEE, MMMM d, yyyy");

  var hour = parseInt(Utilities.formatDate(now, tz, "H"));
  var min  = parseInt(Utilities.formatDate(now, tz, "m"));
  var t    = hour * 60 + min;
  var suggestedShift = t < 11 * 60 ? "AM" : (t < 18 * 60 ? "AFT" : "EVE");

  var userEmail  = "";
  var userRole   = "";
  var tokenError = "";

  var token = e && e.parameter && e.parameter.token ? e.parameter.token : "";
  if (token) {
    var cache      = CacheService.getScriptCache();
    var cachedEmail = cache.get("ml_" + token);
    if (cachedEmail) {
      cache.remove("ml_" + token); // one-time use
      var role = resolveRoleByEmail_(cachedEmail);
      if (role && role !== "none") {
        userEmail = cachedEmail;
        userRole  = role;
      } else {
        tokenError = "This email doesn't have access to the EHS Dog Log.";
      }
    } else {
      tokenError = "This sign-in link has expired or was already used. Please request a new one.";
    }
  }

  var template = HtmlService.createTemplateFromFile("index");

  template.today          = todayStr;
  template.todayFormatted = todayFormatted;
  template.shiftsJson     = JSON.stringify(SHIFTS);
  template.userEmail      = userEmail;
  template.userRole       = userRole;
  template.tokenError     = tokenError;

  if (userEmail && userRole) {
    // Valid magic link — pre-load all data to eliminate first round trips
    template.isAdmin         = userRole === "Admin";
    template.isAlertAccess   = userRole === "Alert Access";
    template.isView          = userRole === "View";
    template.dogsJson        = JSON.stringify(getActiveDogs());
    template.initialLogJson      = JSON.stringify(getLogForDate(todayStr, suggestedShift));
    template.initialLogShift     = suggestedShift;
    template.initialAlertsJson   = JSON.stringify(getActiveAlerts());
    template.recentDepartedJson  = JSON.stringify(getRecentFosterAndAdopted());
    template.respiteJson         = JSON.stringify(getActiveRespites());
  } else {
    // No valid session — serve shell; client will show login or load data via AJAX
    template.isAdmin         = false;
    template.isAlertAccess   = false;
    template.isView          = false;
    template.dogsJson        = "[]";
    template.initialLogJson      = "[]";
    template.initialLogShift     = suggestedShift;
    template.initialAlertsJson   = "[]";
    template.recentDepartedJson  = "{\"foster\":[],\"adopted\":[]}";
    template.respiteJson         = "[]";
  }

  return template.evaluate()
    .setTitle("EHS Dog Log")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

// Kept for legacy error pages if needed
function authPageHtml_(title, subtitle, bodyHtml) {
  return "<!DOCTYPE html><html><head>" +
    "<meta charset='utf-8'>" +
    "<link rel='preconnect' href='https://fonts.googleapis.com'>" +
    "<link href='https://fonts.googleapis.com/css2?family=Nunito:wght@400;700;800;900&display=swap' rel='stylesheet'>" +
    "<style>" +
      "* { box-sizing: border-box; margin: 0; padding: 0; }" +
      "body { font-family: 'Nunito', sans-serif; background: #f5f1ee; display: flex; align-items: center; justify-content: center; min-height: 100vh; padding: 24px; }" +
      ".auth-card { background: white; border-radius: 16px; box-shadow: 0 4px 24px rgba(0,0,0,0.1); padding: 48px 40px; max-width: 480px; width: 100%; text-align: center; }" +
      ".auth-logo { font-size: 36px; font-weight: 900; color: #22c55e; margin-bottom: 8px; }" +
      ".auth-title { font-size: 24px; font-weight: 800; margin-bottom: 8px; }" +
      ".auth-subtitle { font-size: 17px; color: #6e6460; margin-bottom: 28px; }" +
      ".auth-body { font-size: 18px; line-height: 1.6; color: #2c2723; }" +
      ".auth-btn { display: inline-block; margin-top: 20px; padding: 14px 32px; background: #22c55e; color: white; border-radius: 10px; text-decoration: none; font-size: 18px; font-weight: 800; }" +
      ".auth-btn:hover { background: #16a34a; }" +
    "</style></head><body>" +
    "<div class='auth-card'>" +
      "<div class='auth-logo'>🐾 EHS</div>" +
      "<div class='auth-title'>" + title + "</div>" +
      "<div class='auth-subtitle'>" + subtitle + "</div>" +
      "<div class='auth-body'>" + bodyHtml + "</div>" +
    "</div>" +
  "</body></html>";
}


// ------------------------------------------------------------
// sendMagicLink — called from client login screen
// ------------------------------------------------------------
function sendMagicLink(email) {
  email = String(email || "").trim().toLowerCase();
  if (!email || email.indexOf("@") === -1) {
    return { ok: false, message: "Please enter a valid email address." };
  }

  var role = resolveRoleByEmail_(email);
  if (role === "none") {
    return { ok: false, message: "This email doesn't have access to the EHS Dog Log. Contact your shelter administrator." };
  }

  var token  = Utilities.getUuid();
  CacheService.getScriptCache().put("ml_" + token, email, 900); // 15-min TTL

  var appUrl = ScriptApp.getService().getUrl();
  MailApp.sendEmail({
    to:      email,
    subject: "EHS Dog Log — Sign In Link",
    body:    "Click the link below to sign in to the EHS Dog Log.\n\n" +
             "This link expires in 15 minutes and can only be used once.\n\n" +
             appUrl + "?token=" + token + "\n\n" +
             "If you didn't request this, you can ignore this email."
  });

  return { ok: true };
}


// ------------------------------------------------------------
// getInitialData — called by returning users (localStorage session)
// Returns all data needed to boot the app in one round trip
// ------------------------------------------------------------
function getInitialData(email) {
  var role = resolveRoleByEmail_(email);
  if (role === "none") return { error: "no_access" };

  var tz  = Session.getScriptTimeZone();
  var now = new Date();
  return {
    role:           role,
    isAdmin:        role === "Admin",
    isAlertAccess:  role === "Alert Access",
    isView:         role === "View",
    dogs:           getActiveDogs(),
    recentDeparted: getRecentFosterAndAdopted(),
    respites:       getActiveRespites(),
    alerts:         getActiveAlerts(),
    today:          Utilities.formatDate(now, tz, "yyyy-MM-dd"),
    todayFormatted: Utilities.formatDate(now, tz, "EEEE, MMMM d, yyyy")
  };
}


// ------------------------------------------------------------
// getActiveDogs — all Active dogs, sorted alphabetically
// ------------------------------------------------------------
function getActiveDogs() {
  var sheet = getSpreadsheet_().getSheetByName(PROFILES_TAB);
  var data  = sheet.getDataRange().getValues();
  var dogs  = [];

  var tz     = Session.getScriptTimeZone();
  var nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");
  for (var i = 1; i < data.length; i++) {
    if (data[i][PC.STATUS] === "Active" && data[i][PC.DOG_NAME]) {
      var arrivalStr = profileDate_(data[i][PC.ARRIVAL_DATE]);
      if (arrivalStr && arrivalStr > nowStr) continue; // Incoming
      var adoptionStr = profileDate_(data[i][PC.ADOPTION_DATE]);
      if (adoptionStr && adoptionStr <= nowStr) continue; // Auto-adopted
      dogs.push(rowToProfile_(data[i]));
    }
  }
  dogs.sort(function(a, b) { return a.dogName.localeCompare(b.dogName); });
  return dogs;
}


// ------------------------------------------------------------
// getAllDogs — admin: all dogs (active + inactive)
// ------------------------------------------------------------
function getAllDogs(email) {
  requireAdminEmail_(email);
  var sheet = getSpreadsheet_().getSheetByName(PROFILES_TAB);
  var data  = sheet.getDataRange().getValues();
  var dogs  = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][PC.DOG_NAME]) {
      dogs.push(rowToProfile_(data[i]));
    }
  }
  dogs.sort(function(a, b) { return a.dogName.localeCompare(b.dogName); });
  return dogs;
}


// ------------------------------------------------------------
// getLogForDate — returns log entries for a given date + shift
// ------------------------------------------------------------
function getLogForDate(date, shift) {
  var sheet   = getSpreadsheet_().getSheetByName(LOG_TAB);
  var data    = sheet.getDataRange().getValues();
  var entries = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (rowDate_(row) === date && row[LC.SHIFT] === shift) {
      entries.push(rowToLogEntry_(row));
    }
  }
  return entries;
}


// ------------------------------------------------------------
// saveLogEntry — upsert: update existing row or append new one
// ------------------------------------------------------------
function saveLogEntry(date, shift, dog, bm, med, food, walk, comments, email) {
  var sheet = getSpreadsheet_().getSheetByName(LOG_TAB);
  var data  = sheet.getDataRange().getValues();
  var now   = new Date();
  var poster = email || "unknown";
  var bmVal  = normBm_(bm);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (rowDate_(row) === date && row[LC.SHIFT] === shift && row[LC.DOG_NAME] === dog) {
      var r = i + 1;
      sheet.getRange(r, LC.BM + 1, 1, 7).setValues([[
        bmVal, med || "", food || "", walk || "", comments || "", poster, now
      ]]);
      return { success: true, action: "updated" };
    }
  }

  sheet.appendRow([
    date, shift, dog, bmVal,
    med || "", food || "", walk || "", comments || "",
    poster, now
  ]);
  return { success: true, action: "created" };
}


// ------------------------------------------------------------
// getDogDetails — full profile for one dog (used by details panel)
// ------------------------------------------------------------
function getDogDetails(dogName) {
  var sheet = getSpreadsheet_().getSheetByName(PROFILES_TAB);
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][PC.DOG_NAME] === dogName) {
      return rowToProfile_(data[i]);
    }
  }
  return null;
}


// ------------------------------------------------------------
// getLogHistory — historical log queries
// ------------------------------------------------------------
function getLogHistory(mode, param, days) {
  var sheet = getSpreadsheet_().getSheetByName(LOG_TAB);
  var data  = sheet.getDataRange().getValues();

  if (mode === "byDate") {
    var result = { AM: [], AFT: [], EVE: [] };
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (rowDate_(row) === param) {
        var s = row[LC.SHIFT];
        if (result[s]) result[s].push(rowToLogEntry_(row));
      }
    }
    return result;
  }

  if (mode === "byDog") {
    var numDays = days || 14;
    var cutoff  = new Date();
    cutoff.setDate(cutoff.getDate() - numDays);
    var cutoffStr = Utilities.formatDate(cutoff, Session.getScriptTimeZone(), "yyyy-MM-dd");

    var entries = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[LC.DOG_NAME] !== param) continue;
      var d = rowDate_(row);
      if (d >= cutoffStr) entries.push(rowToLogEntry_(row));
    }

    var shiftOrder = { EVE: 0, AFT: 1, AM: 2 };
    entries.sort(function(a, b) {
      if (b.date !== a.date) return b.date.localeCompare(a.date);
      return (shiftOrder[a.shift] || 9) - (shiftOrder[b.shift] || 9);
    });
    return entries;
  }

  return [];
}


// ------------------------------------------------------------
// saveDogProfile — admin: add or update a dog profile
// ------------------------------------------------------------
function saveDogProfile(dogData, email) {
  requireAdminEmail_(email);
  var ss     = getSpreadsheet_();
  var sheet  = ss.getSheetByName(PROFILES_TAB);
  var data   = sheet.getDataRange().getValues();
  var newRow = [
    dogData.dogName             || "",        // 0  DOG_NAME
    dogData.status              || "Active",  // 1  STATUS
    dogData.breed               || "",        // 2  BREED
    dogData.age                 || "",        // 3  AGE
    dogData.weight              || "",        // 4  WEIGHT
    dogData.sex                 || "",        // 5  SEX
    dogData.arrivalDate         || "",        // 6  ARRIVAL_DATE
    dogData.adoptionDate        || "",        // 7  ADOPTION_DATE
    dogData.colour              || "",        // 8  COLOUR
    dogData.colourReason        || "",        // 9  COLOUR_REASON
    dogData.breakfast           || "",        // 10 BREAKFAST
    dogData.lunch               || "",        // 11 LUNCH
    dogData.dinner              || "",        // 12 DINNER
    dogData.medicationRequired  ? "Y" : "N", // 13 MEDICATION_REQUIRED
    dogData.medBreakfast        || "",        // 14 MED_B
    dogData.medLunch            || "",        // 15 MED_L
    dogData.medDinner           || "",        // 16 MED_D
    dogData.history             || "",        // 17 HISTORY
    dogData.feedingNotes        || "",        // 18 FEEDING_NOTES
    dogData.medical             || "",        // 19 MEDICAL
    dogData.behaviour           || "",        // 20 BEHAVIOUR
    dogData.kennel              || "",        // 21 KENNEL
    dogData.walks               || "",        // 22 WALKS
    dogData.walksAllowed        ? "Y" : "N", // 23 WALKS_ALLOWED
    dogData.backyard            || "",        // 24 BACKYARD
    dogData.likes               || "",        // 25 LIKES
    dogData.dislikes            || "",        // 26 DISLIKES
    "",                                       // 27 STATUS_CHANGED_AT — preserved on update below
    ""                                        // 28 PHOTO_URL — preserved on update below
  ];

  var searchName = dogData.originalDogName || dogData.dogName;
  for (var i = 1; i < data.length; i++) {
    if (data[i][PC.DOG_NAME] === searchName) {
      var histSheet = ss.getSheetByName(PROFILE_HISTORY_TAB);
      if (histSheet) {
        var oldRow = data[i].slice(0, PROFILE_HEADERS.length);
        histSheet.appendRow([new Date(), email].concat(oldRow));
      }
      newRow[PC.STATUS_CHANGED_AT] = data[i][PC.STATUS_CHANGED_AT] || "";
      sheet.getRange(i + 1, 1, 1, 28).setValues([newRow.slice(0, 28)]);
      // PHOTO_URL (index 28, col 29) is not included in slice — preserved as-is
      return { success: true, action: "updated" };
    }
  }
  newRow[PC.STATUS_CHANGED_AT] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  sheet.appendRow(newRow);
  return { success: true, action: "created" };
}


// ------------------------------------------------------------
// Photo management — admin only
// ------------------------------------------------------------
function getOrCreatePhotoFolder_() {
  var folders = DriveApp.getFoldersByName("EHS Dog Photos");
  return folders.hasNext() ? folders.next() : DriveApp.createFolder("EHS Dog Photos");
}

function saveProfilePhoto(dogName, base64Data, mimeType, email) {
  requireAdminEmail_(email);
  var sheet = getSpreadsheet_().getSheetByName(PROFILES_TAB);
  var data  = sheet.getDataRange().getValues();
  var rowIndex = -1;
  var oldUrl   = "";
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][PC.DOG_NAME]) === dogName) {
      rowIndex = i;
      oldUrl   = data[i][PC.PHOTO_URL] ? String(data[i][PC.PHOTO_URL]) : "";
      break;
    }
  }
  if (rowIndex === -1) throw new Error("Dog not found: " + dogName);

  if (oldUrl) {
    var m = oldUrl.match(/[?&]id=([^&]+)/);
    if (m) {
      try { DriveApp.getFileById(m[1]).setTrashed(true); } catch (e) { /* ignore */ }
    }
  }

  var ext  = mimeType.indexOf("png") !== -1 ? "png" : mimeType.indexOf("gif") !== -1 ? "gif" : "jpg";
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, dogName + "." + ext);
  var file = getOrCreatePhotoFolder_().createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) { /* domain policy may restrict public sharing */ }
  var url = "https://lh3.googleusercontent.com/d/" + file.getId();

  sheet.getRange(rowIndex + 1, PC.PHOTO_URL + 1).setValue(url);
  return { success: true, url: url };
}

function removeProfilePhoto(dogName, email) {
  requireAdminEmail_(email);
  var sheet = getSpreadsheet_().getSheetByName(PROFILES_TAB);
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][PC.DOG_NAME]) === dogName) {
      sheet.getRange(i + 1, PC.PHOTO_URL + 1).setValue("");
      return { success: true };
    }
  }
  return { success: false, error: "Dog not found" };
}

function migrateProfileSchemaV4() {
  var sheet   = getSpreadsheet_().getSheetByName(PROFILES_TAB);
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf("Photo URL") === -1) {
    sheet.getRange(1, lastCol + 1).setValue("Photo URL");
    Logger.log("migrateProfileSchemaV4: added Photo URL header at col " + (lastCol + 1));
  } else {
    Logger.log("migrateProfileSchemaV4: Photo URL header already present");
  }
}


// ------------------------------------------------------------
// setDogStatus — admin: set a dog's status
// ------------------------------------------------------------
function setDogStatus(dogName, status, email) {
  requireAdminEmail_(email);
  var sheet = getSpreadsheet_().getSheetByName(PROFILES_TAB);
  var data  = sheet.getDataRange().getValues();

  var now = new Date();
  for (var i = 1; i < data.length; i++) {
    if (data[i][PC.DOG_NAME] === dogName) {
      sheet.getRange(i + 1, PC.STATUS + 1).setValue(status);
      sheet.getRange(i + 1, PC.STATUS_CHANGED_AT + 1).setValue(
        Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
      );
      if (status === "Adopted") {
        sheet.getRange(i + 1, PC.ADOPTION_DATE + 1).setValue("");
      }
      return { success: true };
    }
  }
  throw new Error("Dog not found: " + dogName);
}


// ------------------------------------------------------------
// getProfileHistory — admin: last 20 profile snapshots for a dog
// ------------------------------------------------------------
function getProfileHistory(dogName, email) {
  requireAdminEmail_(email);
  var sheet = getSpreadsheet_().getSheetByName(PROFILE_HISTORY_TAB);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data    = sheet.getDataRange().getValues();
  var tz      = Session.getScriptTimeZone();
  var entries = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[2] || "") !== dogName) continue;
    var ts = row[0];
    entries.push({
      timestamp:         ts instanceof Date ? Utilities.formatDate(ts, tz, "MMM d, yyyy h:mm a") : String(ts || ""),
      changedBy:         String(row[1] || ""),
      colour:            String(row[PC.COLOUR            + 2] || ""),
      colourReason:      String(row[PC.COLOUR_REASON     + 2] || ""),
      walksAllowed:      row[PC.WALKS_ALLOWED     + 2] === "Y",
      medicationRequired:row[PC.MEDICATION_REQUIRED + 2] === "Y",
      medBreakfast:      String(row[PC.MED_B        + 2] || ""),
      medLunch:          String(row[PC.MED_L        + 2] || ""),
      medDinner:         String(row[PC.MED_D        + 2] || ""),
      medical:           String(row[PC.MEDICAL      + 2] || "")
    });
  }

  entries.reverse();
  return entries.slice(0, 20);
}


// ------------------------------------------------------------
// getActiveAlerts — returns alerts that haven't expired and are active
// ------------------------------------------------------------
function getActiveAlerts() {
  var ss        = getSpreadsheet_();
  var sheet     = ss.getSheetByName(ALERTS_TAB);
  if (!sheet) return [];
  var data    = sheet.getDataRange().getValues();
  var tz      = Session.getScriptTimeZone();
  var today   = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  var nameMap = {};
  var usersSheet = ss.getSheetByName(USERS_TAB);
  if (usersSheet) {
    var users = usersSheet.getDataRange().getValues();
    for (var u = 1; u < users.length; u++) {
      var uEmail = String(users[u][UC.EMAIL] || "").trim().toLowerCase();
      var name   = String(users[u][UC.NAME]  || "").trim();
      if (uEmail && name) nameMap[uEmail] = name;
    }
  }

  var alerts = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[ALC.MESSAGE]) continue;
    if (row[ALC.ACTIVE] !== "Y") continue;
    var expires = String(row[ALC.EXPIRES_ON] instanceof Date
      ? Utilities.formatDate(row[ALC.EXPIRES_ON], tz, "yyyy-MM-dd")
      : row[ALC.EXPIRES_ON] || "");
    if (expires < today) continue;

    var visibleFrom = String(row[ALC.VISIBLE_FROM] instanceof Date
      ? Utilities.formatDate(row[ALC.VISIBLE_FROM], tz, "yyyy-MM-dd")
      : row[ALC.VISIBLE_FROM] || "");
    if (visibleFrom && visibleFrom > today) continue;

    var rawEmail = String(row[ALC.CREATED_BY] || "");
    alerts.push({
      rowIndex:    i + 1,
      message:     String(row[ALC.MESSAGE] || ""),
      createdBy:   nameMap[rawEmail.toLowerCase()] || rawEmail,
      expiresOn:   expires,
      category:    String(row[ALC.CATEGORY]     || "GENERAL"),
      shiftFilter: String(row[ALC.SHIFT_FILTER] || "ALL"),
      visibleFrom: visibleFrom,
      handoff:     String(row[ALC.HANDOFF] || "")
    });
  }
  return alerts;
}


// ------------------------------------------------------------
// saveAlert — any user; non-admins limited to today/tomorrow expiry
// ------------------------------------------------------------
function saveAlert(message, expiresOn, category, shiftFilter, visibleFrom, handoff, email) {
  if (!message || !message.trim()) {
    throw new Error("Message is required.");
  }
  var tz      = Session.getScriptTimeZone();
  var today   = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var poster  = email || "unknown";

  var tomorrow = Utilities.formatDate(
    new Date(new Date().getTime() + 86400000), tz, "yyyy-MM-dd"
  );
  var expiry = canManageAlertsEmail_(poster)
    ? (expiresOn || today)
    : (expiresOn === tomorrow ? tomorrow : today);

  var cat   = category    || "GENERAL";
  var shift = shiftFilter || "ALL";
  var from  = visibleFrom || today;

  var sheet = getSpreadsheet_().getSheetByName(ALERTS_TAB);
  sheet.appendRow([message.trim(), poster, new Date(), expiry, "Y", cat, shift, from, handoff || ""]);
  return { success: true };
}


// ------------------------------------------------------------
// Migration functions (unchanged)
// ------------------------------------------------------------
function migrateDogProfileSchema() {
  var sheet = getSpreadsheet_().getSheetByName(PROFILES_TAB);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers[21] === "Med (Breakfast)") { Logger.log("Already migrated — skipping."); return; }
  sheet.insertColumnsAfter(21, 3);
  sheet.getRange(1, 22).setValue("Med (Breakfast)");
  sheet.getRange(1, 23).setValue("Med (Lunch)");
  sheet.getRange(1, 24).setValue("Med (Dinner)");
  sheet.getRange(1, 22, 1, 3).setFontWeight("bold").setBackground("#e8f0eb");
  Logger.log("Migration complete. Columns 22-24 added to Dog Profiles.");
}

function migrateAlertSchema() {
  var sheet = getSpreadsheet_().getSheetByName(ALERTS_TAB);
  if (!sheet) { Logger.log("Shift Alerts tab not found."); return; }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers[5] === "Category") { Logger.log("Already migrated — skipping."); return; }
  sheet.getRange(1, 6).setValue("Category");
  sheet.getRange(1, 7).setValue("Shift Filter");
  sheet.getRange(1, 8).setValue("Visible From");
  sheet.getRange(1, 6, 1, 3).setFontWeight("bold").setBackground("#e8f0eb");
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var tz    = Session.getScriptTimeZone();
    var today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    for (var r = 2; r <= lastRow; r++) {
      sheet.getRange(r, 6).setValue("GENERAL");
      sheet.getRange(r, 7).setValue("ALL");
      sheet.getRange(r, 8).setValue(today);
    }
  }
  Logger.log("Alert schema migration complete. Columns 6-8 added.");
}

function migrateProfileSchemaV2() {
  var sheet = getSpreadsheet_().getSheetByName(PROFILES_TAB);
  if (!sheet) { Logger.log("Dog Profiles tab not found."); return; }
  var lastCol = sheet.getLastColumn();
  if (lastCol < 26) { sheet.insertColumnsAfter(lastCol, 26 - lastCol); }
  var headers = sheet.getRange(1, 1, 1, 26).getValues()[0];
  var changed = false;
  if (!headers[24] || headers[24] === "") { sheet.getRange(1, 25).setValue("Status Changed At"); changed = true; }
  if (!headers[25] || headers[25] === "") { sheet.getRange(1, 26).setValue("Adoption Date"); changed = true; }
  if (changed) {
    sheet.getRange(1, 25, 1, 2).setFontWeight("bold").setBackground("#e8f0eb");
    Logger.log("migrateProfileSchemaV2: headers added to columns 25–26.");
  } else {
    Logger.log("migrateProfileSchemaV2: already up to date — no changes made.");
  }
}

function migrateAlertSchemaV2() {
  var sheet = getSpreadsheet_().getSheetByName(ALERTS_TAB);
  if (!sheet) { Logger.log("Shift Alerts tab not found."); return; }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers[8] === "Handoff") { Logger.log("migrateAlertSchemaV2: already up to date — skipping."); return; }
  sheet.getRange(1, 9).setValue("Handoff");
  sheet.getRange(1, 9).setFontWeight("bold").setBackground("#e8f0eb");
  Logger.log("migrateAlertSchemaV2: 'Handoff' header added to column 9.");
}

function migrateProfileSchemaV3() {
  var sheet = getSpreadsheet_().getSheetByName(PROFILES_TAB);
  if (!sheet) { Logger.log("Dog Profiles tab not found."); return; }
  var expected = PROFILE_HEADERS;
  var lastCol  = sheet.getLastColumn();
  if (lastCol < expected.length) {
    sheet.insertColumnsAfter(lastCol, expected.length - lastCol);
    Logger.log("Extended sheet to " + expected.length + " columns.");
  }
  var actual  = sheet.getRange(1, 1, 1, expected.length).getValues()[0];
  var changed = false;
  for (var i = 0; i < expected.length; i++) {
    var actualVal   = String(actual[i] || "").trim();
    var expectedVal = expected[i];
    Logger.log("Col " + (i + 1) + ": actual='" + actualVal + "'  expected='" + expectedVal + "'" +
      (actualVal === expectedVal ? "  ✓" : "  ← MISMATCH"));
    if (actualVal !== expectedVal) { sheet.getRange(1, i + 1).setValue(expectedVal); changed = true; }
  }
  if (changed) {
    sheet.getRange(1, 1, 1, expected.length).setFontWeight("bold").setBackground("#e8f0eb");
    Logger.log("migrateProfileSchemaV3: headers updated.");
  } else {
    Logger.log("migrateProfileSchemaV3: all headers match — no changes needed.");
  }
}


// ------------------------------------------------------------
// checkProfileHeaders — returns { ok, detail }
// ------------------------------------------------------------
function checkProfileHeaders() {
  var sheet = getSpreadsheet_().getSheetByName(PROFILES_TAB);
  if (!sheet) return { ok: false, detail: "Dog Profiles tab not found" };
  var lastCol = sheet.getLastColumn();
  if (lastCol < PROFILE_HEADERS.length) {
    return { ok: false, detail: "Sheet has " + lastCol + " columns, expected " + PROFILE_HEADERS.length };
  }
  var actual     = sheet.getRange(1, 1, 1, PROFILE_HEADERS.length).getValues()[0];
  var mismatches = [];
  for (var i = 0; i < PROFILE_HEADERS.length; i++) {
    var a = String(actual[i] || "").trim();
    if (a !== PROFILE_HEADERS[i]) {
      mismatches.push("col " + (i + 1) + ": got '" + a + "', expected '" + PROFILE_HEADERS[i] + "'");
    }
  }
  return mismatches.length === 0
    ? { ok: true,  detail: "" }
    : { ok: false, detail: mismatches.join("; ") };
}


// ------------------------------------------------------------
// runTests — run from the Apps Script editor
// ------------------------------------------------------------
function runTests() {
  var passed = 0, failed = 0;
  function assert(name, ok, detail) {
    if (ok) { Logger.log("✓ " + name); passed++; }
    else     { Logger.log("✗ FAIL: " + name + (detail ? " — " + detail : "")); failed++; }
  }

  Logger.log("═══ EHS Dog Log — Backend Tests ═══");
  var tz        = Session.getScriptTimeZone();
  var today     = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var testEmail = Session.getEffectiveUser().getEmail(); // script owner — must be an admin or alert manager

  // ── Fix 7: Profile schema ──────────────────────────────────
  Logger.log("\n── Fix 7: Dog Profiles schema ──");
  var headerResult = checkProfileHeaders();
  assert("All " + PROFILE_HEADERS.length + " Dog Profiles headers correct",
    headerResult.ok, headerResult.detail);

  // ── Fix 5: Alert handoff round-trip ───────────────────────
  Logger.log("\n── Fix 5: Alert handoff field ──");

  var msg1 = "__TEST_" + new Date().getTime() + "__";
  try {
    saveAlert(msg1, today, "GENERAL", "ALL", today, "Y", testEmail);
    var alerts1 = getActiveAlerts();
    var found1  = null;
    for (var i = 0; i < alerts1.length; i++) {
      if (alerts1[i].message === msg1) { found1 = alerts1[i]; break; }
    }
    assert("Test alert (handoff='Y') created and retrieved", found1 !== null);
    if (found1) {
      assert("handoff field is 'Y'", found1.handoff === "Y", "got '" + found1.handoff + "'");
      dismissAlert(found1.rowIndex, testEmail);
    }
  } catch(e) { assert("saveAlert/getActiveAlerts handoff='Y' round-trip", false, e.message); }

  var msg2 = "__TEST2_" + new Date().getTime() + "__";
  try {
    saveAlert(msg2, today, "GENERAL", "ALL", today, "", testEmail);
    var alerts2 = getActiveAlerts();
    var found2  = null;
    for (var i = 0; i < alerts2.length; i++) {
      if (alerts2[i].message === msg2) { found2 = alerts2[i]; break; }
    }
    assert("Test alert (no handoff) retrieved", found2 !== null);
    if (found2) {
      assert("handoff field is '' when not a handoff", found2.handoff === "", "got '" + found2.handoff + "'");
      dismissAlert(found2.rowIndex, testEmail);
    }
  } catch(e) { assert("saveAlert without handoff round-trip", false, e.message); }

  Logger.log("\n═══ " + passed + " passed · " + failed + " failed ═══");
}


// ------------------------------------------------------------
// clearRoleCache — clears the cached role for a specific email
// Run manually from the editor after changing a user's role
// ------------------------------------------------------------
function clearRoleCache(email) {
  if (email) {
    CacheService.getScriptCache().remove("role_" + String(email).trim().toLowerCase());
    Logger.log("Role cache cleared for: " + email);
  } else {
    Logger.log("Pass an email address to clearRoleCache(email).");
  }
}


// ------------------------------------------------------------
// dismissAlert — admin or alert manager only
// ------------------------------------------------------------
function dismissAlert(rowIndex, email) {
  requireAlertManagerEmail_(email);
  var sheet = getSpreadsheet_().getSheetByName(ALERTS_TAB);
  sheet.getRange(rowIndex, ALC.ACTIVE + 1).setValue("N");
  return { success: true };
}


// ------------------------------------------------------------
// getIncomingDogs — admin: Active dogs with future arrival date
// ------------------------------------------------------------
function getIncomingDogs(email) {
  requireAdminEmail_(email);
  var sheet = getSpreadsheet_().getSheetByName(PROFILES_TAB);
  var data  = sheet.getDataRange().getValues();
  var tz     = Session.getScriptTimeZone();
  var nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");
  var dogs   = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][PC.STATUS] === "Active" && data[i][PC.DOG_NAME]) {
      var arrivalStr = profileDate_(data[i][PC.ARRIVAL_DATE]);
      if (arrivalStr && arrivalStr > nowStr) {
        dogs.push(rowToProfile_(data[i]));
      }
    }
  }
  dogs.sort(function(a, b) { return a.dogName.localeCompare(b.dogName); });
  return dogs;
}


// ------------------------------------------------------------
// getRecentFosterAndAdopted
// ------------------------------------------------------------
function getRecentFosterAndAdopted() {
  var sheet = getSpreadsheet_().getSheetByName(PROFILES_TAB);
  var data  = sheet.getDataRange().getValues();
  var tz    = Session.getScriptTimeZone();
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 7);
  var cutoffStr = Utilities.formatDate(cutoff, tz, "yyyy-MM-dd");

  var foster  = [];
  var adopted = [];

  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][PC.STATUS] || "");
    if (!data[i][PC.DOG_NAME]) continue;

    if (status === "Foster" || status === "Adopted") {
      var changedStr  = String(data[i][PC.STATUS_CHANGED_AT] || "");
      var changedDate = changedStr ? changedStr.substring(0, 10) : "";
      if (changedDate && changedDate < cutoffStr) continue;
      var profile = rowToProfile_(data[i]);
      if (status === "Foster") foster.push(profile);
      else                     adopted.push(profile);
    } else if (status === "Active") {
      var adoptStr = profileDate_(data[i][PC.ADOPTION_DATE]);
      if (!adoptStr) continue;
      var nowStr2 = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");
      if (adoptStr > nowStr2) continue;
      if (adoptStr.substring(0, 10) < cutoffStr) continue;
      adopted.push(rowToProfile_(data[i]));
    }
  }

  foster.sort(function(a, b)  { return a.dogName.localeCompare(b.dogName); });
  adopted.sort(function(a, b) { return a.dogName.localeCompare(b.dogName); });
  return { foster: foster, adopted: adopted };
}


// ------------------------------------------------------------
// getActiveRespites
// ------------------------------------------------------------
function getActiveRespites() {
  var sheet = getSpreadsheet_().getSheetByName(RESPITE_TAB);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data   = sheet.getDataRange().getValues();
  var tz     = Session.getScriptTimeZone();
  var nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");
  var respites = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[RSC.DOG_NAME]) continue;
    if (String(row[RSC.RETURNED_EARLY] || "") === "Y") continue;
    var endStr = profileDate_(row[RSC.END]);
    if (!endStr || endStr <= nowStr) continue;
    var startStr = profileDate_(row[RSC.START]);
    respites.push({
      dogName:  String(row[RSC.DOG_NAME]),
      start:    startStr,
      end:      endStr,
      notes:    String(row[RSC.NOTES] || ""),
      rowIndex: i + 1
    });
  }
  return respites;
}


// ------------------------------------------------------------
// scheduleRespite — admin only
// ------------------------------------------------------------
function scheduleRespite(dogName, start, end, notes, email) {
  requireAdminEmail_(email);
  if (!dogName || !start || !end) throw new Error("Dog name, start, and end are required.");
  if (end <= start) throw new Error("End must be after start.");
  var sheet = getSpreadsheet_().getSheetByName(RESPITE_TAB);
  sheet.appendRow([dogName, start, end, notes || "", "", ""]);
  return { success: true };
}


// ------------------------------------------------------------
// markReturnedEarly — any user
// ------------------------------------------------------------
function markReturnedEarly(dogName) {
  var sheet = getSpreadsheet_().getSheetByName(RESPITE_TAB);
  if (!sheet || sheet.getLastRow() < 2) throw new Error("No respite records found.");

  var data   = sheet.getDataRange().getValues();
  var tz     = Session.getScriptTimeZone();
  var nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[RSC.DOG_NAME]) !== dogName) continue;
    if (String(row[RSC.RETURNED_EARLY] || "") === "Y") continue;
    var endStr = profileDate_(row[RSC.END]);
    if (!endStr || endStr <= nowStr) continue;
    sheet.getRange(i + 1, RSC.RETURNED_EARLY + 1).setValue("Y");
    sheet.getRange(i + 1, RSC.RETURNED_AT + 1).setValue(nowStr);
    return { success: true };
  }
  throw new Error("No active respite found for: " + dogName);
}


// ------------------------------------------------------------
// cancelRespite — admin only
// ------------------------------------------------------------
function cancelRespite(rowIndex, email) {
  requireAdminEmail_(email);
  var sheet = getSpreadsheet_().getSheetByName(RESPITE_TAB);
  var tz    = Session.getScriptTimeZone();
  sheet.getRange(rowIndex, RSC.RETURNED_EARLY + 1).setValue("Y");
  sheet.getRange(rowIndex, RSC.RETURNED_AT    + 1).setValue(
    Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm")
  );
  return { success: true };
}


// ============================================================
// Concurrency Locking — CacheService.getScriptCache()
// ============================================================
var LOCK_TTL_SECONDS = 300;

function getUserDisplayName_(email) {
  try {
    var sheet = getSpreadsheet_().getSheetByName(USERS_TAB);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][UC.EMAIL]).trim().toLowerCase() === email.toLowerCase()) {
          var name = String(data[i][UC.NAME] || "").trim();
          if (name) return name;
        }
      }
    }
  } catch (e) { /* fall through */ }
  return email.split("@")[0];
}

function acquireProfileLock(dogName, email) {
  requireAdminEmail_(email);
  var cache = CacheService.getScriptCache();
  var key   = "doglock_" + dogName;
  var existing = cache.get(key);
  if (existing) {
    var info = JSON.parse(existing);
    if (info.email !== email) return { success: false, lockedBy: info.displayName };
  }
  var payload = JSON.stringify({ email: email, displayName: getUserDisplayName_(email), lockedAt: new Date().toISOString() });
  cache.put(key, payload, LOCK_TTL_SECONDS);
  return { success: true };
}

function renewProfileLock(dogName, email) {
  var cache = CacheService.getScriptCache();
  var key   = "doglock_" + dogName;
  var existing = cache.get(key);
  if (!existing) return { success: false };
  var info = JSON.parse(existing);
  if (info.email !== email) return { success: false, lockedBy: info.displayName };
  cache.put(key, existing, LOCK_TTL_SECONDS);
  return { success: true };
}

function releaseProfileLock(dogName, email) {
  var cache = CacheService.getScriptCache();
  var key   = "doglock_" + dogName;
  var existing = cache.get(key);
  if (existing && JSON.parse(existing).email === email) cache.remove(key);
  return { success: true };
}

function getAllProfileLocks(dogNames, email) {
  var cache = CacheService.getScriptCache();
  var keys  = dogNames.map(function(n) { return "doglock_" + n; });
  var raw   = cache.getAll(keys);
  var result = {};
  dogNames.forEach(function(name) {
    var val = raw["doglock_" + name];
    if (val) {
      var info = JSON.parse(val);
      if (info.email !== email) result[name] = info.displayName;
    }
  });
  return result;
}

function acquireLogLock(date, shift, dogName, email) {
  var cache = CacheService.getScriptCache();
  var key   = "loglock_" + date + "_" + shift + "_" + dogName;
  var existing = cache.get(key);
  if (existing) {
    var info = JSON.parse(existing);
    if (info.email !== email) return { success: false, lockedBy: info.displayName };
  }
  var payload = JSON.stringify({ email: email, displayName: getUserDisplayName_(email), lockedAt: new Date().toISOString() });
  cache.put(key, payload, LOCK_TTL_SECONDS);
  return { success: true };
}

function renewLogLock(date, shift, dogName, email) {
  var cache = CacheService.getScriptCache();
  var key   = "loglock_" + date + "_" + shift + "_" + dogName;
  var existing = cache.get(key);
  if (!existing) return { success: false };
  var info = JSON.parse(existing);
  if (info.email !== email) return { success: false, lockedBy: info.displayName };
  cache.put(key, existing, LOCK_TTL_SECONDS);
  return { success: true };
}

function releaseLogLock(date, shift, dogName, email) {
  var cache = CacheService.getScriptCache();
  var key   = "loglock_" + date + "_" + shift + "_" + dogName;
  var existing = cache.get(key);
  if (existing && JSON.parse(existing).email === email) cache.remove(key);
  return { success: true };
}

function getAllLogLocks(date, shift, dogNames, email) {
  var cache = CacheService.getScriptCache();
  var keys  = dogNames.map(function(n) { return "loglock_" + date + "_" + shift + "_" + n; });
  var raw   = cache.getAll(keys);
  var result = {};
  dogNames.forEach(function(name) {
    var val = raw["loglock_" + date + "_" + shift + "_" + name];
    if (val) {
      var info = JSON.parse(val);
      if (info.email !== email) result[name] = info.displayName;
    }
  });
  return result;
}
