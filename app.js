"use strict";

const SESSION_KEY = "pinakaPortalActiveApp";
const PORTAL_CONFIG_URL = "/api/portal/config";
const PORTAL_UNLOCK_URL = "/api/portal/unlock";
const PORTAL_LOGOUT_URL = "/api/portal/logout";

let publicApps = ["maintenance", "joining", "complaints"];

const moduleAccess = [
  {
    key: "fuel",
    marker: "FM",
    title: "Fuel Manager",
    file: "apps/fuel-manager-final.html?v=fuel-log-date-fix-1"
  },
  {
    key: "weighbridge",
    marker: "WB",
    title: "Weighbridge",
    file: "apps/weighbridge-app.html"
  },
  {
    key: "userCharge",
    marker: "UC",
    title: "User Charge Collection",
    file: "apps/user-charge-final.html"
  },
  {
    key: "mrf",
    marker: "MR",
    title: "MRF Supervisor",
    file: "apps/mrf-supervisor-app-v2.html"
  },
  {
    key: "joining",
    marker: "SJ",
    title: "Staff Joining",
    file: "apps/joining-app.html"
  },
  {
    key: "maintenance",
    marker: "MT",
    title: "Maintenance",
    file: "apps/maintenance-app.html"
  },
  {
    key: "complaints",
    marker: "CP",
    title: "Complaints",
    file: "/complaints"
  }
];

const app = document.getElementById("app");

const state = {
  activeApp: null,
  role: null,
  pendingApp: null,
  loginError: "",
  authLoading: true,
  authSource: "Loading portal settings...",
  unlocking: false,
};

function currentRole() {
  return state.role || null;
}

function allowedModules() {
  const role = currentRole();
  if (!role) return [];
  return moduleAccess.filter((item) => role.apps.includes(item.key));
}

function applyPortalApps(apps) {
  if (!Array.isArray(apps) || apps.length === 0) return;
  publicApps = apps
    .filter((item) => item && item.key && item.pinRequired === false)
    .map((item) => item.key);
  if (!publicApps.includes("complaints")) publicApps.push("complaints");
}

async function loadPortalRoles() {
  if (!PORTAL_CONFIG_URL) {
    state.authLoading = false;
      state.authSource = "Portal backend missing, protected apps locked";
    render();
    return;
  }
  try {
    const response = await fetch(PORTAL_CONFIG_URL, {
      method: "GET",
      redirect: "follow",
      cache: "no-store"
    });
    if (!response.ok) throw new Error(`Portal backend returned ${response.status}`);
    const data = await response.json();
    if (data.ok) applyPortalApps(data.apps || []);
    if (!data.ok) throw new Error(data.error || "Portal backend did not return settings");
    state.role = data.auth || null;
    state.authSource = "Portal settings loaded from backend";
    restoreActiveApp();
  } catch (error) {
    state.role = null;
    state.authSource = "Portal backend offline, protected apps locked";
    console.warn("Portal config load failed:", error);
  } finally {
    state.authLoading = false;
    render();
  }
}

function restoreActiveApp() {
  const savedKey = sessionStorage.getItem(SESSION_KEY);
  if (!savedKey) return;
  const saved = moduleAccess.find((item) => item.key === savedKey);
  if (!saved) {
    sessionStorage.removeItem(SESSION_KEY);
    return;
  }
  if (publicApps.includes(saved.key) || allowedModules().some((item) => item.key === saved.key)) {
    state.activeApp = saved;
    return;
  }
  sessionStorage.removeItem(SESSION_KEY);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function render() {
  app.innerHTML = state.activeApp ? AppFrame() : AppHub();
  bindEvents();
}

function AppHub() {
  const role = currentRole();
  return `
    <main class="portal-hub-page">
      <section class="portal-hub">
        <div class="brand-lockup portal-brand">
          <div class="brand-mark">PI</div>
          <div>
            <h1 class="brand-title">Pinaka Infra Unified Operations Portal</h1>
            <p class="brand-subtitle">${role ? escapeHtml(role.label) + " · " : ""}Select an app to open</p>
            <p class="portal-sync-note">${escapeHtml(state.authSource)}</p>
          </div>
          ${role ? '<button class="btn ghost portal-logout" type="button" data-action="logout">Logout</button>' : ""}
        </div><div class="tile-grid portal-tile-grid">
          ${moduleAccess.map(AppTile).join("")}
        </div>
        ${state.pendingApp ? AppUnlock() : ""}
      </section>
    </main>
  `;
}

function AppUnlock() {
  const item = moduleAccess.find((appItem) => appItem.key === state.pendingApp);
  if (!item) return "";
  return `
    <div class="portal-unlock-backdrop" data-action="cancel-unlock">
      <section class="portal-login-card portal-unlock-card" role="dialog" aria-modal="true" aria-labelledby="unlock-title">
        <div class="brand-lockup">
          <div class="brand-mark">${escapeHtml(item.marker)}</div>
          <div>
            <h2 class="brand-title" id="unlock-title">${escapeHtml(item.title)}</h2>
            <p class="brand-subtitle">${state.authLoading ? "Loading portal settings..." : "Enter passcode to open"}</p>
          </div>
        </div>
        <form class="portal-login-form" data-action="unlock">
          <label class="field">
            <span>Passcode</span>
            <input id="portal-pin" name="pin" type="password" inputmode="numeric" autocomplete="off" pattern="[0-9]*" autofocus />
          </label>
          <div class="error-text">${escapeHtml(state.loginError)}</div>
          <button class="btn primary portal-login-btn" type="submit" ${state.authLoading || state.unlocking ? "disabled" : ""}>${state.unlocking ? "Opening..." : "Open App"}</button>
          <button class="btn ghost portal-login-btn" type="button" data-action="cancel-unlock">Cancel</button>
        </form>
      </section>
    </div>
  `;
}

function AppTile(item) {
  return `
    <button class="module-tile portal-app-tile" type="button" data-app-key="${escapeHtml(item.key)}">
      <span class="tile-marker">${escapeHtml(item.marker)}</span>
      <span class="tile-copy">${escapeHtml(item.title)}</span>
    </button>
  `;
}

function AppFrame() {
  return `
    <div class="portal-shell">
      <header class="portal-topbar">
        <div class="brand-lockup">
          <div class="brand-mark">PI</div>
          <div>
            <h1 class="brand-title">Pinaka Infra Unified Operations Portal</h1>
            <p class="brand-subtitle">${escapeHtml(state.activeApp.title)}</p>
          </div>
        </div>
        <div class="portal-actions">
          <button class="btn ghost" data-action="hub">Back to Apps</button>
          <button class="btn ghost" data-action="logout">Logout</button>
        </div>
      </header>
      <iframe
        class="app-frame"
        title="${escapeHtml(state.activeApp.title)}"
        src="${escapeHtml(state.activeApp.file)}"
      ></iframe>
    </div>
  `;
}

function bindEvents() {
  document.querySelectorAll("[data-action='unlock']").forEach((form) => {
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      const item = moduleAccess.find((appItem) => appItem.key === state.pendingApp);
      if (!item) {
        state.unlocking = false;
        state.loginError = "Invalid app selection.";
        render();
        return;
      }
      try {
        const pin = form.pin.value;
        state.unlocking = true;
        state.loginError = "";
        render();
        const response = await fetch(PORTAL_UNLOCK_URL, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ pin, appKey: item.key })
        });
        const data = await response.json().catch(() => ({}));
        if (!response.ok || !data.ok) {
          throw new Error(data.error || "Invalid passcode for this app.");
        } else {
          state.role = data.auth || null;
        }
        state.unlocking = false;
        state.loginError = "";
        state.pendingApp = null;
        state.activeApp = item;
        sessionStorage.setItem(SESSION_KEY, item.key);
        render();
      } catch (error) {
        state.unlocking = false;
        state.loginError = error.message || "Invalid passcode for this app.";
        render();
      }
    });
  });

  document.querySelectorAll("[data-app-key]").forEach((button) => {
    button.addEventListener("click", () => {
      const found = moduleAccess.find((item) => item.key === button.dataset.appKey);
      if (!found) return;
      if (!publicApps.includes(found.key) && (!currentRole() || !allowedModules().some((item) => item.key === found.key))) {
        state.pendingApp = found.key;
        state.loginError = "";
        state.unlocking = false;
        render();
        return;
      }
      state.activeApp = found;
      sessionStorage.setItem(SESSION_KEY, found.key);
      render();
    });
  });

  document.querySelectorAll("[data-action='hub']").forEach((button) => {
    button.addEventListener("click", () => {
      state.activeApp = null;
      sessionStorage.removeItem(SESSION_KEY);
      render();
    });
  });

  document.querySelectorAll("[data-action='cancel-unlock']").forEach((button) => {
    button.addEventListener("click", (event) => {
      if (event.target !== button && button.classList.contains("portal-unlock-backdrop")) return;
      state.pendingApp = null;
      state.loginError = "";
      state.unlocking = false;
      render();
    });
  });

  document.querySelectorAll("[data-action='logout']").forEach((button) => {
    button.addEventListener("click", async () => {
      try {
        await fetch(PORTAL_LOGOUT_URL, { method: "POST" });
      } catch (error) {
        console.warn("Portal logout failed:", error);
      }
      state.role = null;
      state.activeApp = null;
      state.pendingApp = null;
      state.loginError = "";
      state.unlocking = false;
      sessionStorage.removeItem(SESSION_KEY);
      render();
    });
  });
}

render();
loadPortalRoles();
