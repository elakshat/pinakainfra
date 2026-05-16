import http from "node:http";
import fs from "node:fs/promises";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import crypto from "node:crypto";

const root = path.dirname(fileURLToPath(import.meta.url));
loadEnvFile();

function loadEnvFile() {
  try {
    const envText = readFileSync(path.join(root, ".env"), "utf8");
    envText.split(/\r?\n/).forEach((line) => {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("#")) return;
      const eq = trimmed.indexOf("=");
      if (eq < 1) return;
      const key = trimmed.slice(0, eq).trim();
      const value = trimmed.slice(eq + 1).trim().replace(/^["']|["']$/g, "");
      if (key && process.env[key] === undefined) process.env[key] = value;
    });
  } catch {
    // Hosts can still provide environment variables directly.
  }
}

const port = Number(process.env.PORT || 4173);
const disableSheetsBackend = process.env.DISABLE_SHEETS_BACKEND === "true" || process.env.BACKEND_TYPE === "supabase";
const scriptTargets = disableSheetsBackend ? {} : {
  master: process.env.MASTER_SCRIPT_URL,
  fuel: process.env.FUEL_SCRIPT_URL,
  maintenance: process.env.MAINTENANCE_SCRIPT_URL,
  joining: process.env.JOINING_SCRIPT_URL,
  mrf: process.env.MRF_SCRIPT_URL,
  "user-charge": process.env.USER_CHARGE_SCRIPT_URL,
  weighbridge: process.env.WEIGHBRIDGE_SCRIPT_URL,
  complaints: process.env.COMPLAINTS_SCRIPT_URL
};
if (disableSheetsBackend) {
  console.log("Google Sheets backend disabled for Supabase testing.");
}
const appPinOverrides = {
  fuel: process.env.FUEL_PORTAL_PIN || "1111",
  weighbridge: process.env.WEIGHBRIDGE_PORTAL_PIN || "2222",
  userCharge: process.env.USER_CHARGE_PORTAL_PIN || "3333",
  mrf: process.env.MRF_PORTAL_PIN || "4444"
};
const types = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8"
};
const blockedExtensions = new Set([".mjs", ".md", ".gs", ".txt", ".csv", ".xlsx", ".pdf"]);
const blockedPrefixes = ["/apps/scripts/", "/data-import/"];
const appHtmlRoutes = new Map([
  ["/apps/fuel-manager-final.html", "fuel"],
  ["/apps/weighbridge-app.html", "weighbridge"],
  ["/apps/user-charge-final.html", "userCharge"],
  ["/apps/mrf-supervisor-app-v2.html", "mrf"],
  ["/apps/joining-app.html", "joining"],
  ["/apps/maintenance-app.html", "maintenance"]
]);
const sessions = new Map();
const PORTAL_CONFIG_TTL_MS = Number(process.env.PORTAL_CONFIG_TTL_MS || 300000);
let portalConfigCache = null;
let portalConfigPromise = null;

function parseCookies(req) {
  return Object.fromEntries(String(req.headers.cookie || "")
    .split(";")
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => {
      const eq = part.indexOf("=");
      return eq < 0 ? [part, ""] : [part.slice(0, eq), decodeURIComponent(part.slice(eq + 1))];
    }));
}

function sendJson(res, status, data, headers = {}) {
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    ...headers
  });
  res.end(JSON.stringify(data));
}

function isSecureRequest(req) {
  const proto = String(req.headers["x-forwarded-proto"] || "").split(",")[0].trim().toLowerCase();
  const host = String(req.headers.host || "");
  return proto === "https" || (!host.startsWith("localhost") && !host.startsWith("127.0.0.1"));
}

function sessionCookie(req, value, maxAge) {
  const secure = isSecureRequest(req) ? "; Secure" : "";
  return `pinaka_session=${encodeURIComponent(value)}; HttpOnly; SameSite=Lax; Path=/; Max-Age=${maxAge}${secure}`;
}

function sessionFor(req) {
  const token = parseCookies(req).pinaka_session;
  if (!token) return null;
  const session = sessions.get(token);
  if (!session || session.expiresAt < Date.now()) {
    sessions.delete(token);
    return null;
  }
  return session;
}

async function readBody(req) {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  return Buffer.concat(chunks);
}

async function fetchScriptJson(key, search = "", options = {}) {
  const target = scriptTargets[key];
  if (!target) throw new Error(`Backend URL missing for ${key}`);
  const targetUrl = new URL(target);
  targetUrl.search = search;
  const upstream = await fetch(targetUrl, {
    method: options.method || "GET",
    redirect: "follow",
    headers: options.headers || { "Content-Type": "text/plain;charset=utf-8" },
    body: options.body
  });
  const text = await upstream.text();
  try {
    return JSON.parse(text);
  } catch {
    throw new Error("Backend did not return JSON");
  }
}

async function portalConfig(options = {}) {
  const now = Date.now();
  if (!options.force && portalConfigCache && portalConfigCache.expiresAt > now) {
    return portalConfigCache.data;
  }
  if (!options.force && portalConfigCache) {
    if (!portalConfigPromise) {
      portalConfigPromise = fetchScriptJson("master", "?action=portalConfig")
        .then((data) => {
          if (!data.ok) throw new Error(data.error || "Master portal config failed");
          portalConfigCache = { data, expiresAt: Date.now() + PORTAL_CONFIG_TTL_MS };
          return data;
        })
        .finally(() => {
          portalConfigPromise = null;
        });
    }
    portalConfigPromise.catch((error) => console.warn("Portal config refresh failed:", error.message));
    return portalConfigCache.data;
  }
  if (!portalConfigPromise) {
    portalConfigPromise = fetchScriptJson("master", "?action=portalConfig")
      .then((data) => {
        if (!data.ok) throw new Error(data.error || "Master portal config failed");
        portalConfigCache = { data, expiresAt: Date.now() + PORTAL_CONFIG_TTL_MS };
        return data;
      })
      .finally(() => {
        portalConfigPromise = null;
      });
  }
  return portalConfigPromise;
}

function appAccessAllowed(access, key) {
  if (!access) return false;
  if (key === "userCharge") return access.userCharge === true || access.usercharge === true || access.uc === true;
  return access[key] === true;
}

function appTitle(appKey) {
  const titles = {
    fuel: "Fuel Manager",
    weighbridge: "Weighbridge",
    userCharge: "User Charge Collection",
    mrf: "MRF Supervisor",
    joining: "Staff Joining",
    maintenance: "Maintenance",
    complaints: "Complaints"
  };
  return titles[appKey] || "Portal User";
}

function overrideUserForPin(appKey, pin) {
  const expected = String(appPinOverrides[appKey] || "").trim();
  if (!expected || expected !== String(pin || "").trim()) return null;
  return {
    name: appTitle(appKey),
    access: { [appKey]: true }
  };
}

function hasSheetUserForApp(users, appKey) {
  return (users || []).some((user) => appAccessAllowed(user.access, appKey));
}

function appKeyForRoute(key, localUrl) {
  if (key === "user-charge") return "userCharge";
  if (key === "master") return localUrl.searchParams.get("app") || "master";
  return key;
}

async function isAppPublic(appKey) {
  if (appKey === "maintenance" || appKey === "joining" || appKey === "complaints") return true;
  const config = await portalConfig().catch(() => null);
  const app = config && (config.apps || []).find((item) => item.key === appKey);
  return app ? app.pinRequired === false : false;
}

async function canAccessApi(req, key, localUrl) {
  const appKey = appKeyForRoute(key, localUrl);
  if (appKey !== "master" && await isAppPublic(appKey)) return true;
  const session = sessionFor(req);
  if (!session) return false;
  if (appKey === "master") {
    const action = localUrl.searchParams.get("action") || "";
    if (action === "config") {
      const app = localUrl.searchParams.get("app") || "";
      return session.apps.includes(app);
    }
    return session.apps.length > 0;
  }
  return session.apps.includes(appKey);
}

async function canAccessAppHtml(req, appKey) {
  if (!appKey) return true;
  if (await isAppPublic(appKey)) return true;
  const session = sessionFor(req);
  return Boolean(session && session.apps.includes(appKey));
}

async function handlePortalApi(req, res, localUrl) {
  if (localUrl.pathname === "/api/portal/config" && req.method === "GET") {
    try {
      const data = await portalConfig({ force: localUrl.searchParams.get("refresh") === "1" });
      const session = sessionFor(req);
      sendJson(res, 200, {
        ok: true,
        apps: data.apps || [],
        auth: session ? { label: session.label, apps: session.apps } : null
      });
    } catch (error) {
      sendJson(res, 503, { ok: false, error: error.message || "Portal backend unavailable" });
    }
    return;
  }
  if (localUrl.pathname === "/api/portal/unlock" && req.method === "POST") {
    try {
      const body = JSON.parse((await readBody(req)).toString("utf8") || "{}");
      const pin = String(body.pin || "").trim();
      const appKey = String(body.appKey || "").trim();
      const data = await portalConfig().catch(() => null);
      const users = data ? (data.users || []) : [];
      const user = users.find((item) => String(item.pin || "").trim() === pin && appAccessAllowed(item.access, appKey))
        || (!data || !hasSheetUserForApp(users, appKey) ? overrideUserForPin(appKey, pin) : null);
      if (!user) {
        sendJson(res, 401, { ok: false, error: "Invalid passcode for this app." });
        return;
      }
      const apps = data
        ? (data.apps || [])
          .filter((item) => appAccessAllowed(user.access, item.key))
          .map((item) => item.key)
        : [appKey];
      const token = crypto.randomUUID();
      const session = {
        label: String(user.name || user.id || "Portal User"),
        apps,
        expiresAt: Date.now() + 8 * 60 * 60 * 1000
      };
      sessions.set(token, session);
      sendJson(res, 200, {
        ok: true,
        auth: { label: session.label, apps: session.apps }
      }, {
        "Set-Cookie": sessionCookie(req, token, 28800)
      });
    } catch (error) {
      sendJson(res, 503, { ok: false, error: error.message || "Portal unlock unavailable" });
    }
    return;
  }
  if (localUrl.pathname === "/api/portal/logout" && req.method === "POST") {
    const token = parseCookies(req).pinaka_session;
    if (token) sessions.delete(token);
    sendJson(res, 200, { ok: true }, {
      "Set-Cookie": sessionCookie(req, "", 0)
    });
    return;
  }
  sendJson(res, 404, { ok: false, error: "Unknown portal API route" });
}

async function proxyToScript(req, res, localUrl, key) {
  if (!await canAccessApi(req, key, localUrl)) {
    sendJson(res, 401, { ok: false, error: "Not authorized for this app." });
    return;
  }
  const target = scriptTargets[key];
  if (!target) {
    res.writeHead(503, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
    res.end(JSON.stringify({ ok: false, error: `Backend URL missing for ${key}` }));
    return;
  }
  try {
    const targetUrl = new URL(target);
    targetUrl.search = localUrl.search;
    const body = req.method === "GET" || req.method === "HEAD" ? undefined : await readBody(req);
    const upstream = await fetch(targetUrl, {
      method: req.method,
      redirect: "follow",
      headers: {
        "Content-Type": req.headers["content-type"] || "text/plain;charset=utf-8"
      },
      body
    });
    const text = await upstream.text();
    res.writeHead(upstream.status, {
      "Content-Type": upstream.headers.get("content-type") || "application/json; charset=utf-8",
      "Cache-Control": "no-store"
    });
    res.end(text);
  } catch (error) {
    res.writeHead(502, { "Content-Type": "application/json; charset=utf-8" });
    res.end(JSON.stringify({ ok: false, error: "Backend proxy failed", detail: error.message }));
  }
}

const server = http.createServer(async (req, res) => {
  try {
    const localUrl = new URL(req.url, `http://localhost:${port}`);
    const urlPath = decodeURIComponent(localUrl.pathname);
    if (urlPath.startsWith("/api/portal/")) {
      await handlePortalApi(req, res, localUrl);
      return;
    }
    if (urlPath.startsWith("/api/")) {
      await proxyToScript(req, res, localUrl, urlPath.slice(5));
      return;
    }
    const basename = path.basename(urlPath);
    if (basename.startsWith(".") || blockedPrefixes.some((prefix) => urlPath.startsWith(prefix)) || blockedExtensions.has(path.extname(urlPath))) {
      res.writeHead(403, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Forbidden");
      return;
    }
    let requested = urlPath === "/" ? "/index.html" : urlPath;
    if (requested === "/complaints" || requested === "/complaints/") {
      requested = "/complaints/index.html";
    }
    if (requested === "/complaints/admin") {
      requested = "/complaints/admin.html";
    }
    const appKey = appHtmlRoutes.get(requested);
    if (appKey && !await canAccessAppHtml(req, appKey)) {
      res.writeHead(401, {
        "Content-Type": "text/html; charset=utf-8",
        "Cache-Control": "no-store"
      });
      res.end("<!doctype html><meta charset=\"utf-8\"><title>Locked</title><p>This app is locked. Open it from the portal and enter your passcode.</p>");
      return;
    }
    const filePath = path.resolve(root, `.${requested}`);
    if (!filePath.startsWith(root)) {
      res.writeHead(403);
      res.end("Forbidden");
      return;
    }
    const data = await fs.readFile(filePath);
    res.writeHead(200, { "Content-Type": types[path.extname(filePath)] || "application/octet-stream" });
    res.end(data);
  } catch {
    res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Not found");
  }
});

server.listen(port, () => {
  console.log(`Pinaka portal running at http://localhost:${port}`);
});
import { createClient } from "@supabase/supabase-js";

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);
