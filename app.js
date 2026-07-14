/* ==========================================================================
   CampusHub - College Academic & Event Management System
   app.js
   --------------------------------------------------------------------------
   Vanilla JS, no frameworks. Talks to a Google Apps Script Web App REST
   API that reads/writes a Google Sheets backend.

   NOTE: This file is written to be style.css / Code.gs agnostic.
   No inline styles are set from JS (classes are toggled only), so a
   future style.css can restyle everything without touching this file.
   ========================================================================== */

/* ==========================================================================
   1. CONFIG
   ========================================================================== */

// Replace with your deployed Google Apps Script Web App URL.
const API_URL = "https://script.google.com/macros/s/AKfycbxe2mbwDle6sOvhTVxk3CnDatki6GfyEW9AdtyvMwd2I3pq_2k4s1HQfAtMzM9vlkYp/exec";

// Local state, kept as the single source of truth for the current session.
const state = {
  events: [],
  categories: [],
  notices: [],
  upcomingEvents: [],
  activeCategory: null,
  isAdminLoggedIn: false,
  editingEventId: null
};


/* ==========================================================================
   2. GENERIC API HELPER
   ========================================================================== */

/**
 * Generic wrapper around fetch() for talking to the Apps Script backend.
 * @param {string} action - backend action name (e.g. "getEvents")
 * @param {string} method - "GET" or "POST"
 * @param {object|null} payload - body data for POST requests
 * @returns {Promise<any>} parsed JSON response, or throws on failure
 */
async function apiRequest(action, method = "GET", payload = null) {
  try {
    let url = `${API_URL}?action=${encodeURIComponent(action)}`;
    const options = { method };

    if (method === "GET") {
      // Apps Script web apps work best with query params for GET.
      if (payload) {
        Object.keys(payload).forEach((key) => {
          url += `&${encodeURIComponent(key)}=${encodeURIComponent(payload[key])}`;
        });
      }
    } else {
      options.headers = { "Content-Type": "text/plain;charset=utf-8" };
      options.body = JSON.stringify({ action, ...payload });
    }

    const response = await fetch(url, options);

    if (!response.ok) {
      throw new Error(`Request failed with status ${response.status}`);
    }

    const data = await response.json();

    if (data && data.success === false) {
      throw new Error(data.message || "Unknown API error");
    }

    return data;
  } catch (err) {
    console.error(`[apiRequest] "${action}" failed:`, err);
    throw err;
  }
}


/* ==========================================================================
   3. DOM HELPERS
   ========================================================================== */

function qs(selector) {
  return document.querySelector(selector);
}

function setLoading(containerEl, message) {
  if (!containerEl) return;
  containerEl.innerHTML = `<div class="state-msg">${escapeHtml(message)}</div>`;
}

function setEmpty(containerEl, message) {
  if (!containerEl) return;
  containerEl.innerHTML = `<div class="state-msg">${escapeHtml(message)}</div>`;
}

function setError(containerEl, message) {
  if (!containerEl) return;
  containerEl.innerHTML = `<div class="state-msg error">${escapeHtml(message)}</div>`;
}

function escapeHtml(str) {
  if (str === null || str === undefined) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatDate(dateStr) {
  if (!dateStr) return "TBA";
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  return d.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
}


/* ==========================================================================
   4. DATA LOADING FUNCTIONS
   ========================================================================== */

/**
 * Loads all events from the backend and stores them in state.
 */
async function loadEvents() {
  const grid = qs("#latest-events-grid");
  setLoading(grid, "Loading events...");

  try {
    const data = await apiRequest("getEvents", "GET");
    state.events = Array.isArray(data.events) ? data.events : [];
    renderCards();
    renderDashboard();
  } catch (err) {
    setError(grid, "Unable to load events. Please try again later.");
  }
}

/**
 * Loads categories from the backend and renders the category grid + quick buttons.
 */
async function loadCategories() {
  const grid = qs("#category-grid");
  setLoading(grid, "Loading categories...");

  try {
    const data = await apiRequest("getCategories", "GET");
    state.categories = Array.isArray(data.categories) ? data.categories : [];
    renderCategoryGrid();
    renderQuickCategoryButtons();
    populateCategorySelect();
  } catch (err) {
    setError(grid, "Unable to load categories.");
  }
}

/**
 * Loads pinned notices from the backend.
 */
async function loadPinnedNotices() {
  const list = qs("#notices-list");
  setLoading(list, "Loading notices...");

  try {
    const data = await apiRequest("getPinnedNotices", "GET");
    state.notices = Array.isArray(data.notices) ? data.notices : [];
    renderNotices();
  } catch (err) {
    setError(list, "Unable to load notices.");
  }
}

/**
 * Loads upcoming events (date >= today) from the backend.
 */
async function loadUpcomingEvents() {
  const grid = qs("#upcoming-events-grid");
  setLoading(grid, "Loading upcoming events...");

  try {
    const data = await apiRequest("getUpcomingEvents", "GET");
    state.upcomingEvents = Array.isArray(data.events) ? data.events : [];
    renderCalendar();
  } catch (err) {
    setError(grid, "Unable to load upcoming events.");
  }
}


/* ==========================================================================
   5. SEARCH & FILTER
   ========================================================================== */

/**
 * Searches events by keyword (title, category, department, venue) and re-renders.
 * @param {string} keyword
 */
async function searchEvents(keyword) {
  const grid = qs("#latest-events-grid");
  const term = (keyword || "").trim().toLowerCase();

  if (!term) {
    renderCards();
    return;
  }

  setLoading(grid, "Searching...");

  try {
    // Prefer server-side search if backend supports it, otherwise filter locally.
    const data = await apiRequest("searchEvents", "GET", { q: term });
    const results = Array.isArray(data.events) ? data.events : [];
    renderCards(results);
  } catch (err) {
    // Fallback: local filter over already-loaded events.
    const results = state.events.filter((ev) => {
      const haystack = `${ev.title || ""} ${ev.category || ""} ${ev.department || ""} ${ev.venue || ""}`.toLowerCase();
      return haystack.includes(term);
    });
    renderCards(results);
  }
}

/**
 * Filters the currently displayed events by category.
 * @param {string} categoryName
 */
function filterByCategory(categoryName) {
  state.activeCategory = categoryName || null;

  // Highlight active category in the grid.
  document.querySelectorAll(".category-item").forEach((el) => {
    el.classList.toggle("active", el.dataset.category === categoryName);
  });

  if (!categoryName) {
    renderCards();
    return;
  }

  const filtered = state.events.filter(
    (ev) => (ev.category || "").toLowerCase() === categoryName.toLowerCase()
  );
  renderCards(filtered);

  // Scroll to the events section so the filter result is visible.
  const target = qs("#latest-events-section");
  if (target) target.scrollIntoView({ behavior: "smooth" });
}


/* ==========================================================================
   6. EVENT DETAIL MODAL
   ========================================================================== */

/**
 * Opens the event details modal for a given event id.
 * @param {string|number} id
 */
async function openDetails(id) {
  let event = state.events.find((ev) => String(ev.id) === String(id));

  // If not found locally (e.g. came from a search result set), fetch directly.
  if (!event) {
    try {
      const data = await apiRequest("getEventById", "GET", { id });
      event = data.event;
    } catch (err) {
      alert("Unable to load event details.");
      return;
    }
  }

  if (!event) {
    alert("Event not found.");
    return;
  }

  qs("#detail-banner-img").src = event.imageUrl || "";
  qs("#detail-title").textContent = event.title || "Untitled Event";
  qs("#detail-category").textContent = event.category || "—";
  qs("#detail-department").textContent = event.department || "—";
  qs("#detail-date").textContent = formatDate(event.date);
  qs("#detail-time").textContent = event.time || "—";
  qs("#detail-venue").textContent = event.venue || "—";
  qs("#detail-coordinator").textContent = event.coordinator || "—";
  qs("#detail-description").textContent = event.description || "No description provided.";
  qs("#detail-eligibility").textContent = event.eligibility || "Open to all";
  qs("#detail-reg-deadline").textContent = formatDate(event.registrationDeadline);
  qs("#detail-contact").textContent = event.contact || "—";

  const regLink = qs("#detail-reg-link");
  const pdfLink = qs("#detail-pdf-link");

  if (event.registrationLink) {
    regLink.href = event.registrationLink;
    regLink.style.display = "inline-block";
  } else {
    regLink.style.display = "none";
  }

  if (event.pdfUrl) {
    pdfLink.href = event.pdfUrl;
    pdfLink.style.display = "inline-block";
  } else {
    pdfLink.style.display = "none";
  }

  qs("#details-modal-overlay").classList.add("open");
}

/**
 * Closes the event details modal.
 */
function closeDetails() {
  qs("#details-modal-overlay").classList.remove("open");
}


/* ==========================================================================
   7. ADMIN PANEL
   ========================================================================== */

/**
 * Opens the admin panel (assumes login already happened).
 */
function openAdmin() {
  state.isAdminLoggedIn = true;
  qs("#admin-panel").classList.add("open");
  loadAdminData();
}

/**
 * Closes the admin panel and returns to the public site.
 */
function closeAdmin() {
  qs("#admin-panel").classList.remove("open");
}

/**
 * Loads all data needed by the admin panel (events table + stats).
 */
async function loadAdminData() {
  await loadEvents();
  renderEventsTable();
  renderStatistics();
}

/**
 * Switches between admin tabs.
 * @param {string} tabId
 */
function switchAdminTab(tabId) {
  document.querySelectorAll(".admin-section").forEach((sec) => {
    sec.classList.toggle("active", sec.id === tabId);
  });
  document.querySelectorAll(".admin-tab-btn").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.tab === tabId);
  });
}

/**
 * Handles simple admin login. Real auth should happen on the backend.
 */
async function handleLogin(username, password) {
  try {
    const data = await apiRequest("adminLogin", "POST", { username, password });
    if (data.success) {
      openAdmin();
    } else {
      alert(data.message || "Invalid credentials.");
    }
  } catch (err) {
    alert("Login failed. Please try again.");
  }
}

/**
 * Logs the admin out.
 */
function handleLogout() {
  state.isAdminLoggedIn = false;
  closeAdmin();
}


/* ==========================================================================
   8. CRUD OPERATIONS (ADMIN)
   ========================================================================== */

/**
 * Reads the Add/Edit Event form into a plain object.
 */
function readEventForm() {
  return {
    id: qs("#event-id").value || null,
    title: qs("#field-title").value.trim(),
    category: qs("#field-category").value,
    department: qs("#field-department").value.trim(),
    description: qs("#field-description").value.trim(),
    date: qs("#field-date").value,
    time: qs("#field-time").value,
    venue: qs("#field-venue").value.trim(),
    coordinator: qs("#field-coordinator").value.trim(),
    eligibility: qs("#field-eligibility").value.trim(),
    registrationDeadline: qs("#field-reg-deadline").value,
    registrationLink: qs("#field-reg-link").value.trim(),
    imageUrl: qs("#field-image-url").value.trim(),
    pdfUrl: qs("#field-pdf-url").value.trim(),
    contact: qs("#field-contact").value.trim(),
    status: qs("#field-status").value,
    featured: qs("#field-featured").checked,
    pinned: qs("#field-pinned").checked
  };
}

/**
 * Clears the Add/Edit Event form back to defaults.
 */
function resetEventForm() {
  qs("#event-form").reset();
  qs("#event-id").value = "";
  state.editingEventId = null;
  qs("#event-form-heading").textContent = "Add Event";
  qs("#event-form-submit-btn").textContent = "Submit";
}

/**
 * Adds a new event via the backend.
 */
async function addEvent() {
  const payload = readEventForm();

  if (!payload.title || !payload.category) {
    alert("Title and Category are required.");
    return;
  }

  try {
    await apiRequest("addEvent", "POST", payload);
    resetEventForm();
    await loadAdminData();
    alert("Event added successfully.");
  } catch (err) {
    alert("Failed to add event. Please try again.");
  }
}

/**
 * Updates an existing event via the backend.
 */
async function updateEvent() {
  const payload = readEventForm();

  if (!payload.id) {
    alert("No event selected for update.");
    return;
  }

  try {
    await apiRequest("updateEvent", "POST", payload);
    resetEventForm();
    await loadAdminData();
    alert("Event updated successfully.");
  } catch (err) {
    alert("Failed to update event. Please try again.");
  }
}

/**
 * Deletes an event by id via the backend.
 * @param {string|number} id
 */
async function deleteEvent(id) {
  if (!confirm("Are you sure you want to delete this event?")) return;

  try {
    await apiRequest("deleteEvent", "POST", { id });
    await loadAdminData();
  } catch (err) {
    alert("Failed to delete event. Please try again.");
  }
}

/**
 * Loads an event into the form for editing.
 * @param {string|number} id
 */
function editEventById(id) {
  const event = state.events.find((ev) => String(ev.id) === String(id));
  if (!event) {
    alert("Event not found.");
    return;
  }

  state.editingEventId = id;
  qs("#event-id").value = event.id;
  qs("#field-title").value = event.title || "";
  qs("#field-category").value = event.category || "";
  qs("#field-department").value = event.department || "";
  qs("#field-description").value = event.description || "";
  qs("#field-date").value = event.date || "";
  qs("#field-time").value = event.time || "";
  qs("#field-venue").value = event.venue || "";
  qs("#field-coordinator").value = event.coordinator || "";
  qs("#field-eligibility").value = event.eligibility || "";
  qs("#field-reg-deadline").value = event.registrationDeadline || "";
  qs("#field-reg-link").value = event.registrationLink || "";
  qs("#field-image-url").value = event.imageUrl || "";
  qs("#field-pdf-url").value = event.pdfUrl || "";
  qs("#field-contact").value = event.contact || "";
  qs("#field-status").value = event.status || "active";
  qs("#field-featured").checked = !!event.featured;
  qs("#field-pinned").checked = !!event.pinned;

  qs("#event-form-heading").textContent = "Edit Event";
  qs("#event-form-submit-btn").textContent = "Update Event";

  switchAdminTab("admin-add-event-tab");
}


/* ==========================================================================
   9. RENDER FUNCTIONS
   ========================================================================== */

/**
 * Renders the top dashboard summary cards based on loaded events.
 */
function renderDashboard() {
  const counts = {
    events: state.events.length,
    notices: state.events.filter((e) => e.pinned).length,
    exams: state.events.filter((e) => (e.category || "").toLowerCase().includes("exam")).length,
    placements: state.events.filter((e) => (e.category || "").toLowerCase().includes("placement")).length,
    workshops: state.events.filter((e) => (e.category || "").toLowerCase().includes("workshop")).length,
    holidays: state.events.filter((e) => (e.category || "").toLowerCase().includes("holiday")).length
  };

  const map = {
    "#count-events": counts.events,
    "#count-notices": counts.notices,
    "#count-exams": counts.exams,
    "#count-placements": counts.placements,
    "#count-workshops": counts.workshops,
    "#count-holidays": counts.holidays
  };

  Object.keys(map).forEach((selector) => {
    const el = qs(selector);
    if (el) el.textContent = map[selector];
  });
}

/**
 * Renders the category grid buttons.
 */
function renderCategoryGrid() {
  const grid = qs("#category-grid");

  if (!state.categories.length) {
    setEmpty(grid, "No categories available.");
    return;
  }

  grid.innerHTML = state.categories
    .map(
      (cat) => `
      <button class="category-item" type="button" data-category="${escapeHtml(cat.name)}" onclick="filterByCategory('${escapeHtml(cat.name)}')">
        ${escapeHtml(cat.name)}
      </button>`
    )
    .join("");
}

/**
 * Renders quick category buttons inside the hero section.
 */
function renderQuickCategoryButtons() {
  const container = qs("#quick-categories");
  if (!container) return;

  if (!state.categories.length) {
    container.innerHTML = "";
    return;
  }

  // Show only the first few for a clean hero layout.
  const quickList = state.categories.slice(0, 6);

  container.innerHTML = quickList
    .map(
      (cat) => `
      <button class="quick-cat-btn" type="button" onclick="filterByCategory('${escapeHtml(cat.name)}')">
        ${escapeHtml(cat.name)}
      </button>`
    )
    .join("");
}

/**
 * Fills the Category <select> in the Add/Edit Event form.
 */
function populateCategorySelect() {
  const select = qs("#field-category");
  if (!select) return;

  if (!state.categories.length) {
    select.innerHTML = `<option value="">No categories found</option>`;
    return;
  }

  select.innerHTML =
    `<option value="">Select category</option>` +
    state.categories.map((cat) => `<option value="${escapeHtml(cat.name)}">${escapeHtml(cat.name)}</option>`).join("");
}

/**
 * Renders event cards into the Latest Events grid.
 * @param {Array} [eventsList] optional override list (e.g. search/filter results)
 */
function renderCards(eventsList) {
  const grid = qs("#latest-events-grid");
  const list = eventsList || state.events;

  if (!list.length) {
    setEmpty(grid, "No events found.");
    return;
  }

  grid.innerHTML = list.map((ev) => buildEventCardHtml(ev)).join("");
}

/**
 * Builds the HTML markup for a single event card.
 */
function buildEventCardHtml(ev) {
  return `
    <article class="event-card" data-id="${escapeHtml(ev.id)}">
      <img class="event-card-img" src="${escapeHtml(ev.imageUrl || "")}" alt="${escapeHtml(ev.title || "Event")}" onerror="this.style.opacity='0.3'">
      <div class="event-card-body">
        <span class="event-card-badge">${escapeHtml(ev.category || "General")}</span>
        <div class="event-card-title">${escapeHtml(ev.title || "Untitled Event")}</div>
        <div class="event-card-meta">📅 ${escapeHtml(formatDate(ev.date))}</div>
        <div class="event-card-meta">📍 ${escapeHtml(ev.venue || "TBA")}</div>
        <div class="event-card-meta">🏢 ${escapeHtml(ev.department || "—")}</div>
        <button class="view-details-btn" type="button" onclick="openDetails('${escapeHtml(ev.id)}')">View Details</button>
      </div>
    </article>`;
}

/**
 * Renders the Upcoming Events section.
 */
function renderCalendar() {
  const grid = qs("#upcoming-events-grid");

  if (!state.upcomingEvents.length) {
    setEmpty(grid, "No upcoming events right now.");
    return;
  }

  grid.innerHTML = state.upcomingEvents.map((ev) => buildEventCardHtml(ev)).join("");
}

/**
 * Renders the Pinned Notices list.
 */
function renderNotices() {
  const list = qs("#notices-list");

  if (!state.notices.length) {
    setEmpty(list, "No pinned notices at the moment.");
    return;
  }

  list.innerHTML = state.notices
    .map(
      (n) => `
      <div class="notice-item">
        <span class="notice-title">${escapeHtml(n.title || "Notice")}</span>
        <span class="notice-date">${escapeHtml(formatDate(n.date))}</span>
      </div>`
    )
    .join("");
}

/**
 * Renders the admin events management table.
 */
function renderEventsTable() {
  const tbody = qs("#events-table-body");
  if (!tbody) return;

  if (!state.events.length) {
    tbody.innerHTML = `<tr><td colspan="5" class="state-msg">No events found.</td></tr>`;
    return;
  }

  tbody.innerHTML = state.events
    .map(
      (ev) => `
      <tr>
        <td>${escapeHtml(ev.title || "")}</td>
        <td>${escapeHtml(ev.category || "")}</td>
        <td>${escapeHtml(formatDate(ev.date))}</td>
        <td>${escapeHtml(ev.status || "active")}</td>
        <td>
          <button class="table-action-btn edit-btn" type="button" onclick="editEventById('${escapeHtml(ev.id)}')">Edit</button>
          <button class="table-action-btn delete-btn" type="button" onclick="deleteEvent('${escapeHtml(ev.id)}')">Delete</button>
        </td>
      </tr>`
    )
    .join("");

  // Also mirror stats/dashboard counts in admin dashboard.
  qs("#admin-count-total").textContent = state.events.length;
  qs("#admin-count-featured").textContent = state.events.filter((e) => e.featured).length;
  qs("#admin-count-pinned").textContent = state.events.filter((e) => e.pinned).length;
  qs("#admin-count-active").textContent = state.events.filter((e) => (e.status || "active") === "active").length;
}

/**
 * Renders the Statistics tab (basic counts by category).
 */
function renderStatistics() {
  const container = qs("#statistics-container");
  if (!container) return;

  if (!state.events.length) {
    setEmpty(container, "No data available for statistics.");
    return;
  }

  const byCategory = {};
  state.events.forEach((ev) => {
    const cat = ev.category || "Uncategorized";
    byCategory[cat] = (byCategory[cat] || 0) + 1;
  });

  container.innerHTML = `
    <div id="admin-stats-grid">
      ${Object.keys(byCategory)
        .map(
          (cat) => `
        <div class="dash-card">
          <div class="dash-count">${byCategory[cat]}</div>
          <div class="dash-label">${escapeHtml(cat)}</div>
        </div>`
        )
        .join("")}
    </div>`;
}


/* ==========================================================================
   10. EVENT LISTENERS / INIT
   ========================================================================== */

function attachEventListeners() {
  // Header search.
  const globalSearch = qs("#global-search");
  if (globalSearch) {
    globalSearch.addEventListener("keydown", (e) => {
      if (e.key === "Enter") searchEvents(globalSearch.value);
    });
  }

  // Hero search form.
  const heroForm = qs("#hero-search-form");
  if (heroForm) {
    heroForm.addEventListener("submit", (e) => {
      e.preventDefault();
      searchEvents(qs("#hero-search-input").value);
    });
  }

  // Login button opens a very basic prompt-based login (replace with real modal later).
  const loginBtn = qs("#login-btn");
  if (loginBtn) {
    loginBtn.addEventListener("click", () => {
      const username = prompt("Admin username:");
      if (username === null) return;
      const password = prompt("Admin password:");
      if (password === null) return;
      handleLogin(username, password);
    });
  }

  // Details modal close handlers.
  qs("#details-close-btn").addEventListener("click", closeDetails);
  qs("#detail-close-action-btn").addEventListener("click", closeDetails);
  qs("#details-modal-overlay").addEventListener("click", (e) => {
    if (e.target.id === "details-modal-overlay") closeDetails();
  });

  // Admin logout.
  qs("#admin-logout-btn").addEventListener("click", handleLogout);

  // Admin tab switching.
  document.querySelectorAll(".admin-tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => switchAdminTab(btn.dataset.tab));
  });

  // Event form submit (handles both add + update).
  qs("#event-form").addEventListener("submit", (e) => {
    e.preventDefault();
    if (state.editingEventId) {
      updateEvent();
    } else {
      addEvent();
    }
  });
}

/**
 * Initializes the app: wires up listeners and loads all initial data.
 */
async function initApp() {
  qs("#footer-year").textContent = new Date().getFullYear();
  attachEventListeners();

  // Load everything in parallel for faster first paint.
  await Promise.all([
    loadCategories(),
    loadEvents(),
    loadPinnedNotices(),
    loadUpcomingEvents()
  ]);
}

document.addEventListener("DOMContentLoaded", initApp);
