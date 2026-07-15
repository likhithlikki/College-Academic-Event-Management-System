/* ==========================================================================
  
   --------------------------------------------------------------------------
   SETUP INSTRUCTIONS
   --------------------------------------------------------------------------
   1. Create a Google Sheet with two tabs (exact names matter):

      a) "Events" — header row (row 1) with these exact column names:
         id | title | category | department | description | date | time |
         venue | coordinator | eligibility | registrationDeadline |
         registrationLink | imageUrl | pdfUrl | contact | status |
         featured | pinned | createdAt | updatedAt

      b) "Admins" — header row with:
         username | password

         Add at least one row, e.g. username="admin", password="changeme".
         (For a real deployment, replace plaintext passwords with hashes —
         see the note above adminLogin_ below.)

      Categories are NOT a separate sheet — they are derived automatically
      from the distinct "category" values already used in the Events sheet,
      so admins never have to maintain a category list by hand. If you'd
      rather curate a fixed category list, add a "Categories" tab with a
      single "name" column and the code below will use it automatically.

   2. Extensions -> Apps Script, paste this file in as Code.gs.

   3. Deploy -> New deployment -> Web app.
        - Execute as: Me
        - Who has access: Anyone
      Copy the resulting URL into API_URL in app.js.

   4. Every request from app.js is a GET (for reads) or a POST with a
      text/plain JSON body (for writes). Both are handled below.
   ========================================================================== */

/* ==========================================================================
   1. CONFIG
   ========================================================================== */

// Paste the ID of your Google Sheet here (the long string in its URL,
// between /d/ and /edit). Required if this is a standalone Apps Script
// project (Extensions > Apps Script from a blank script, not opened
// from inside the Sheet itself).
const SHEET_ID = "1VmeTgOGTe94f4Y-iuAGdkXdU3EFV6Oc_MwUqvZoYYtw";

const SHEET_NAME_EVENTS = "Events";
const SHEET_NAME_ADMINS = "Admins";
const SHEET_NAME_CATEGORIES = "Categories"; // optional

/**
 * Returns the spreadsheet to use. Works whether the script is bound to
 * the sheet (Extensions > Apps Script from inside the Sheet) or is a
 * standalone project (in which case getActiveSpreadsheet() is null and
 * we fall back to opening SHEET_ID directly).
 */
function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  if (!SHEET_ID || SHEET_ID === "PASTE_YOUR_SHEET_ID_HERE") {
    throw new Error(
      "No active spreadsheet found and SHEET_ID is not set. " +
      "Open your Google Sheet, copy the ID from its URL, and paste it into SHEET_ID at the top of Code.gs."
    );
  }
  return SpreadsheetApp.openById(SHEET_ID);
}

const EVENT_COLUMNS = [
  "id", "title", "category", "department", "description", "date", "time",
  "venue", "coordinator", "eligibility", "registrationDeadline",
  "registrationLink", "imageUrl", "pdfUrl", "contact", "status",
  "featured", "pinned", "createdAt", "updatedAt"
];


/* ==========================================================================
   2. ENTRY POINTS (doGet / doPost)
   ========================================================================== */

function doGet(e) {
  return handleRequest_(e, "GET");
}

function doPost(e) {
  return handleRequest_(e, "POST");
}

/**
 * Routes an incoming request to the correct handler based on "action".
 */
function handleRequest_(e, method) {
  let action = "";
  let params = {};

  try {
    if (method === "GET") {
      params = (e && e.parameter) || {};
      action = params.action || "";
    } else {
      // app.js POSTs a JSON string as the raw body (text/plain).
      const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
      params = body;
      action = body.action || "";
    }

    switch (action) {
      // ---- GET actions ----
      case "getEvents":
        return jsonResponse_({ success: true, events: getAllEvents_() });

      case "getCategories":
        return jsonResponse_({ success: true, categories: getAllCategories_() });

      case "getPinnedNotices":
        return jsonResponse_({ success: true, notices: getPinnedNotices_() });

      case "getUpcomingEvents":
        return jsonResponse_({ success: true, events: getUpcomingEvents_() });

      case "searchEvents":
        return jsonResponse_({ success: true, events: searchEvents_(params.q || "") });

      case "getEventById":
        return jsonResponse_({ success: true, event: getEventById_(params.id) });

      // ---- POST (admin) actions ----
      case "adminLogin":
        return jsonResponse_(adminLogin_(params.username, params.password));

      case "addEvent":
        return jsonResponse_(addEvent_(params));

      case "updateEvent":
        return jsonResponse_(updateEvent_(params));

      case "deleteEvent":
        return jsonResponse_(deleteEvent_(params.id));

      default:
        return jsonResponse_({ success: false, message: "Unknown action: " + action });
    }
  } catch (err) {
    return jsonResponse_({ success: false, message: err && err.message ? err.message : String(err) });
  }
}

/**
 * Wraps a JS object as a JSON text output (Apps Script web apps can only
 * return text; app.js reads it via response.json()).
 */
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ==========================================================================
   3. SHEET HELPERS
   ========================================================================== */

function getSheet_(name) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" not found.');
  return sheet;
}

/**
 * Reads a sheet into an array of plain objects keyed by the header row.
 */
function sheetToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map((h) => String(h).trim());
  const rows = values.slice(1);

  return rows
    .filter((row) => row.some((cell) => cell !== "" && cell !== null))
    .map((row) => {
      const obj = {};
      headers.forEach((header, i) => {
        obj[header] = row[i];
      });
      return obj;
    });
}

function normalizeEvent_(raw) {
  const ev = Object.assign({}, raw);
  ev.id = String(ev.id);
  ev.date = formatDateForClient_(ev.date);
  ev.registrationDeadline = formatDateForClient_(ev.registrationDeadline);
  ev.featured = ev.featured === true || ev.featured === "TRUE" || ev.featured === "true" || ev.featured === 1;
  ev.pinned = ev.pinned === true || ev.pinned === "TRUE" || ev.pinned === "true" || ev.pinned === 1;
  ev.status = ev.status || "active";
  return ev;
}

function formatDateForClient_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value);
}


/* ==========================================================================
   4. READ OPERATIONS
   ========================================================================== */

function getAllEvents_() {
  const sheet = getSheet_(SHEET_NAME_EVENTS);
  return sheetToObjects_(sheet).map(normalizeEvent_);
}

function getAllCategories_() {
  const ss = getSpreadsheet_();
  const catSheet = ss.getSheetByName(SHEET_NAME_CATEGORIES);

  if (catSheet) {
    return sheetToObjects_(catSheet)
      .filter((row) => row.name)
      .map((row) => ({ name: String(row.name) }));
  }

  // Fallback: derive distinct categories from the Events sheet.
  const events = getAllEvents_();
  const seen = {};
  const categories = [];
  events.forEach((ev) => {
    const name = (ev.category || "").trim();
    if (name && !seen[name]) {
      seen[name] = true;
      categories.push({ name: name });
    }
  });
  return categories;
}

function getPinnedNotices_() {
  return getAllEvents_()
    .filter((ev) => ev.pinned)
    .sort((a, b) => new Date(a.date || 0) - new Date(b.date || 0))
    .map((ev) => ({ id: ev.id, title: ev.title, date: ev.date }));
}

function getUpcomingEvents_() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return getAllEvents_()
    .filter((ev) => {
      if (!ev.date) return false;
      const d = new Date(ev.date);
      return !isNaN(d.getTime()) && d >= today && ev.status !== "cancelled";
    })
    .sort((a, b) => new Date(a.date) - new Date(b.date));
}

function searchEvents_(query) {
  const term = String(query || "").trim().toLowerCase();
  if (!term) return getAllEvents_();

  return getAllEvents_().filter((ev) => {
    const haystack = [ev.title, ev.category, ev.department, ev.venue, ev.description]
      .map((v) => String(v || "").toLowerCase())
      .join(" ");
    return haystack.indexOf(term) !== -1;
  });
}

function getEventById_(id) {
  if (!id) throw new Error("Missing event id.");
  const event = getAllEvents_().find((ev) => String(ev.id) === String(id));
  if (!event) throw new Error("Event not found.");
  return event;
}


/* ==========================================================================
   5. WRITE OPERATIONS (ADMIN)
   ========================================================================== */

/**
 * NOTE ON SECURITY: this compares plaintext passwords stored in the Admins
 * sheet, matching the simple prompt()-based login already in app.js. For a
 * real deployment, store a hash (e.g. Utilities.computeDigest) instead of
 * plaintext and compare hashes here.
 */
function adminLogin_(username, password) {
  if (!username || !password) {
    return { success: false, message: "Username and password are required." };
  }

  const admins = sheetToObjects_(getSheet_(SHEET_NAME_ADMINS));
  const match = admins.find(
    (a) => String(a.username) === String(username) && String(a.password) === String(password)
  );

  if (!match) {
    return { success: false, message: "Invalid credentials." };
  }

  return { success: true };
}

function addEvent_(payload) {
  if (!payload.title || !payload.category) {
    return { success: false, message: "Title and Category are required." };
  }

  const sheet = getSheet_(SHEET_NAME_EVENTS);
  const now = new Date();
  const id = Utilities.getUuid();

  const row = EVENT_COLUMNS.map((col) => {
    if (col === "id") return id;
    if (col === "createdAt" || col === "updatedAt") return now;
    if (col === "featured" || col === "pinned") return !!payload[col];
    return payload[col] !== undefined && payload[col] !== null ? payload[col] : "";
  });

  sheet.appendRow(row);

  return { success: true, id: id };
}

function updateEvent_(payload) {
  if (!payload.id) {
    return { success: false, message: "Missing event id." };
  }

  const sheet = getSheet_(SHEET_NAME_EVENTS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map((h) => String(h).trim());
  const idCol = headers.indexOf("id");

  if (idCol === -1) {
    return { success: false, message: 'Events sheet is missing an "id" column.' };
  }

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) === String(payload.id)) {
      const now = new Date();
      const newRow = headers.map((col) => {
        if (col === "id") return payload.id;
        if (col === "createdAt") return values[r][headers.indexOf("createdAt")];
        if (col === "updatedAt") return now;
        if (col === "featured" || col === "pinned") return !!payload[col];
        return payload[col] !== undefined && payload[col] !== null ? payload[col] : values[r][headers.indexOf(col)];
      });

      sheet.getRange(r + 1, 1, 1, newRow.length).setValues([newRow]);
      return { success: true };
    }
  }

  return { success: false, message: "Event not found." };
}

function deleteEvent_(id) {
  if (!id) {
    return { success: false, message: "Missing event id." };
  }

  const sheet = getSheet_(SHEET_NAME_EVENTS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map((h) => String(h).trim());
  const idCol = headers.indexOf("id");

  if (idCol === -1) {
    return { success: false, message: 'Events sheet is missing an "id" column.' };
  }

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) === String(id)) {
      sheet.deleteRow(r + 1);
      return { success: true };
    }
  }

  return { success: false, message: "Event not found." };
}


/* ==========================================================================
   6. ONE-TIME SETUP HELPER (optional)
   --------------------------------------------------------------------------
   Run this once from the Apps Script editor (select setupSheets from the
   function dropdown and click Run) to auto-create the Events and Admins
   sheets with correct headers if they don't already exist. Safe to run
   more than once — it will not overwrite existing sheets/data.
   ========================================================================== */

function setupSheets() {
  const ss = getSpreadsheet_();

  if (!ss.getSheetByName(SHEET_NAME_EVENTS)) {
    const sheet = ss.insertSheet(SHEET_NAME_EVENTS);
    sheet.getRange(1, 1, 1, EVENT_COLUMNS.length).setValues([EVENT_COLUMNS]);
    sheet.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_NAME_ADMINS)) {
    const sheet = ss.insertSheet(SHEET_NAME_ADMINS);
    sheet.getRange(1, 1, 1, 2).setValues([["username", "password"]]);
    sheet.appendRow(["admin", "changeme"]);
    sheet.setFrozenRows(1);
  }

  SpreadsheetApp.flush();
  Logger.log("Setup complete.");
}
