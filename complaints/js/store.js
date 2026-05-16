(function () {
  var STORAGE_KEY = "msw_crm_complaints";
  var COLLECTION = "complaints";
  var SHEETS_API = "/api/complaints";
  var appStarted = false;
  var firestore = null;
  var sheetsAvailable = null;

  function getConfig() {
    return window.MSW_CONFIG || {};
  }

  function hasFirebaseConfig() {
    var cfg = getConfig().firebaseConfig || {};
    return Boolean(cfg.apiKey && cfg.projectId && cfg.appId);
  }

  function initFirebaseApp() {
    if (!hasFirebaseConfig() || !window.firebase) return false;
    if (!window.firebase.apps.length) {
      window.firebase.initializeApp(getConfig().firebaseConfig);
    }
    return true;
  }

  function init() {
    if (appStarted) return getMode();
    appStarted = true;

    if (initFirebaseApp()) {
      try {
        firestore = window.firebase.firestore ? window.firebase.firestore() : null;
      } catch (err) {
        console.warn("Firebase could not start. Falling back to demo mode.", err);
        firestore = null;
      }
    }

    return getMode();
  }

  function getMode() {
    if (sheetsAvailable === true) return "sheets";
    if (firestore) return "firebase";
    return "demo";
  }

  function updateModeBadge() {
    var el = document.getElementById("backend-mode");
    if (!el) return;
    init();
    var mode = getMode();
    if (mode === "sheets") el.textContent = "Google Sheets backend";
    else el.textContent = mode === "firebase" ? "Firebase backend" : "Demo local backend";
    el.className = "mode-pill " + mode;
  }

  function generateComplaintId() {
    var rnd = Math.floor(100 + Math.random() * 900);
    return "CMP-" + Date.now() + rnd;
  }

  function getLocalComplaints() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      return raw ? JSON.parse(raw) : [];
    } catch (err) {
      return [];
    }
  }

  function saveLocalComplaints(list) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
  }

  function normalizeDate(value) {
    if (!value) return new Date().toISOString();
    if (value && typeof value.toDate === "function") return value.toDate().toISOString();
    return value;
  }

  function normalizeRecord(record) {
    record = record || {};
    return {
      complaintId: record.complaintId || generateComplaintId(),
      name: record.name || "",
      mobile: record.mobile || "",
      ward: record.ward || "",
      type: record.type || "",
      description: record.description || "",
      status: record.status || "Open",
      assignedOfficer: record.assignedOfficer || "",
      adminRemark: record.adminRemark || "",
      priority: record.priority || "Medium",
      createdAt: normalizeDate(record.createdAt),
      updatedAt: normalizeDate(record.updatedAt),
      createdByUid: record.createdByUid || "",
      createdByPhone: record.createdByPhone || "",
    };
  }

  function currentUser() {
    if (!window.firebase || !window.firebase.auth) return null;
    return window.firebase.auth().currentUser;
  }

  function fetchSheets(path, options) {
    return fetch(SHEETS_API + path, options || {})
      .then(function (response) {
        return response.text().then(function (text) {
          var data = {};
          try {
            data = text ? JSON.parse(text) : {};
          } catch (err) {
            throw new Error("Google Sheets backend did not return JSON.");
          }
          if (!response.ok || data.ok === false) {
            throw new Error(data.error || data.msg || "Google Sheets backend failed.");
          }
          sheetsAvailable = true;
          updateModeBadge();
          return data;
        });
      })
      .catch(function (err) {
        if (sheetsAvailable === null) sheetsAvailable = false;
        updateModeBadge();
        throw err;
      });
  }

  function sheetsCreateComplaint(record) {
    return fetchSheets("", {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action: "createComplaint", complaint: record }),
    }).then(function (data) {
      return normalizeRecord(data.complaint || record);
    });
  }

  function sheetsListComplaints() {
    return fetchSheets("?action=list", { method: "GET", cache: "no-store" }).then(function (data) {
      return (data.complaints || []).map(normalizeRecord);
    });
  }

  function sheetsUpdateComplaint(complaintId, patch) {
    return fetchSheets("", {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action: "updateComplaint", complaintId: complaintId, patch: patch }),
    }).then(function (data) {
      return normalizeRecord(data.complaint || {});
    });
  }

  function createComplaint(payload) {
    init();
    var now = new Date().toISOString();
    var user = currentUser();
    var record = normalizeRecord(
      Object.assign({}, payload, {
        complaintId: generateComplaintId(),
        status: "Open",
        assignedOfficer: "",
        adminRemark: "",
        priority: "Medium",
        createdAt: now,
        updatedAt: now,
        createdByUid: user ? user.uid : "",
        createdByPhone: user ? user.phoneNumber || "" : "",
      })
    );

    return sheetsCreateComplaint(record).catch(function (sheetErr) {
      console.warn("Google Sheets backend unavailable. Falling back to secondary storage.", sheetErr);
      if (firestore) {
        return firestore
          .collection(COLLECTION)
          .doc(record.complaintId)
          .set(record)
          .then(function () {
            return record;
          });
      }
      var list = getLocalComplaints();
      list.unshift(record);
      saveLocalComplaints(list);
      return record;
    });
  }

  function listComplaints() {
    init();
    return sheetsListComplaints().catch(function (sheetErr) {
      console.warn("Google Sheets backend unavailable. Falling back to secondary storage.", sheetErr);
      if (!firestore) return getLocalComplaints();
      return firestore
        .collection(COLLECTION)
        .orderBy("createdAt", "desc")
        .get()
        .then(function (snapshot) {
          var records = [];
          snapshot.forEach(function (doc) {
            records.push(normalizeRecord(doc.data()));
          });
          return records;
        });
    });
  }

  function updateComplaint(complaintId, patch) {
    init();
    var cleanPatch = Object.assign({}, patch, {
      updatedAt: new Date().toISOString(),
    });

    return sheetsUpdateComplaint(complaintId, cleanPatch).catch(function (sheetErr) {
      console.warn("Google Sheets backend unavailable. Falling back to secondary storage.", sheetErr);
      if (firestore) {
        return firestore
          .collection(COLLECTION)
          .doc(complaintId)
          .update(cleanPatch)
          .then(function () {
            return firestore.collection(COLLECTION).doc(complaintId).get();
          })
          .then(function (doc) {
            return normalizeRecord(doc.data() || {});
          });
      }
      var list = getLocalComplaints();
      var idx = list.findIndex(function (item) {
        return item.complaintId === complaintId;
      });
      if (idx === -1) return null;
      list[idx] = normalizeRecord(Object.assign({}, list[idx], cleanPatch));
      saveLocalComplaints(list);
      return list[idx];
    });
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  window.MSWStore = {
    init: init,
    getMode: getMode,
    updateModeBadge: updateModeBadge,
    createComplaint: createComplaint,
    listComplaints: listComplaints,
    updateComplaint: updateComplaint,
    escapeHtml: escapeHtml,
  };
})();
