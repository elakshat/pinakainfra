(function () {
  function byId(id) {
    return document.getElementById(id);
  }

  function config() {
    return window.MSW_CONFIG || {};
  }

  function isLoggedIn() {
    return sessionStorage.getItem("msw_admin_ok") === "1";
  }

  function setLoggedIn(yes) {
    if (yes) sessionStorage.setItem("msw_admin_ok", "1");
    else sessionStorage.removeItem("msw_admin_ok");
  }

  function badgeClass(status) {
    if (status === "Open") return "badge badge-open";
    if (status === "In Progress") return "badge badge-progress";
    if (status === "Resolved") return "badge badge-resolved";
    return "badge";
  }

  function formatDate(iso) {
    if (!iso) return "";
    var date = new Date(iso);
    if (Number.isNaN(date.getTime())) return "";
    return date.toLocaleString();
  }

  function filterByDate(list, fromStr, toStr) {
    if (!fromStr && !toStr) return list;
    return list.filter(function (item) {
      var time = new Date(item.createdAt).getTime();
      if (Number.isNaN(time)) return false;
      if (fromStr) {
        var from = new Date(fromStr);
        from.setHours(0, 0, 0, 0);
        if (time < from.getTime()) return false;
      }
      if (toStr) {
        var to = new Date(toStr);
        to.setHours(23, 59, 59, 999);
        if (time > to.getTime()) return false;
      }
      return true;
    });
  }

  function renderStats(list) {
    var total = list.length;
    var open = list.filter(function (c) {
      return c.status === "Open";
    }).length;
    var progress = list.filter(function (c) {
      return c.status === "In Progress";
    }).length;
    var resolved = list.filter(function (c) {
      return c.status === "Resolved";
    }).length;

    byId("stats-bar").innerHTML =
      '<div class="stat-card"><span>Total</span><strong>' +
      total +
      '</strong></div><div class="stat-card"><span>Open</span><strong>' +
      open +
      '</strong></div><div class="stat-card"><span>In Progress</span><strong>' +
      progress +
      '</strong></div><div class="stat-card"><span>Resolved</span><strong>' +
      resolved +
      "</strong></div>";
  }

  function officerOptions(selected) {
    var officers = config().officers || [];
    var html = '<option value="">Unassigned</option>';
    officers.forEach(function (officer) {
      var name = window.MSWStore.escapeHtml(officer.name);
      var ward = window.MSWStore.escapeHtml(officer.ward || "");
      html +=
        '<option value="' +
        name +
        '"' +
        (selected === officer.name ? " selected" : "") +
        ">" +
        name +
        (ward ? " (" + ward + ")" : "") +
        "</option>";
    });
    return html;
  }

  function statusOptions(selected) {
    return ["Open", "In Progress", "Resolved"]
      .map(function (status) {
        return '<option value="' + status + '"' + (selected === status ? " selected" : "") + ">" + status + "</option>";
      })
      .join("");
  }

  function renderComplaint(item) {
    var esc = window.MSWStore.escapeHtml;
    var id = esc(item.complaintId);

    return (
      '<article class="complaint-item" data-id="' +
      id +
      '"><button type="button" class="complaint-head" data-toggle="' +
      id +
      '"><span><span class="complaint-id">' +
      id +
      '</span><span class="complaint-meta">' +
      esc(item.ward) +
      " | " +
      esc(item.type) +
      '</span><span class="complaint-time">' +
      esc(formatDate(item.createdAt)) +
      '</span></span><span class="' +
      badgeClass(item.status) +
      '">' +
      esc(item.status) +
      '</span></button><div class="complaint-body hidden" id="body-' +
      id +
      '"><dl class="detail-grid"><div><dt>Name</dt><dd>' +
      esc(item.name) +
      '</dd></div><div><dt>Mobile</dt><dd>' +
      esc(item.mobile) +
      '</dd></div><div class="full"><dt>Description</dt><dd>' +
      esc(item.description).replace(/\n/g, "<br />") +
      "</dd></div>" +
      '</dl><div class="admin-edit-grid"><div class="field"><label>Assign officer</label><select class="officer-sel" data-id="' +
      id +
      '">' +
      officerOptions(item.assignedOfficer) +
      '</select></div><div class="field"><label>Status</label><select class="status-sel" data-id="' +
      id +
      '">' +
      statusOptions(item.status) +
      '</select></div><div class="field"><label>Admin remark</label><textarea class="remark-ta" data-id="' +
      id +
      '">' +
      esc(item.adminRemark || "") +
      '</textarea></div><button type="button" class="btn save-btn" data-id="' +
      id +
      '">Save changes</button></div></div></article>'
    );
  }

  function renderList() {
    var fromStr = byId("filter-from").value;
    var toStr = byId("filter-to").value;
    var container = byId("complaint-list");

    container.innerHTML = '<p class="hint">Loading complaints...</p>';

    window.MSWStore.listComplaints()
      .then(function (all) {
        var list = filterByDate(all, fromStr, toStr);
        renderStats(list);

        if (!list.length) {
          container.innerHTML = '<section class="panel"><p class="hint">No complaints found.</p></section>';
          return;
        }

        container.innerHTML = list.map(renderComplaint).join("");

        container.querySelectorAll("[data-toggle]").forEach(function (button) {
          button.addEventListener("click", function () {
            var id = button.getAttribute("data-toggle");
            var body = byId("body-" + id);
            if (body) body.classList.toggle("hidden");
          });
        });

        container.querySelectorAll(".save-btn").forEach(function (button) {
          button.addEventListener("click", function () {
            var id = button.getAttribute("data-id");
            var officer = container.querySelector('.officer-sel[data-id="' + id + '"]').value;
            var status = container.querySelector('.status-sel[data-id="' + id + '"]').value;
            var remark = container.querySelector('.remark-ta[data-id="' + id + '"]').value;
            button.disabled = true;

            window.MSWStore.updateComplaint(id, {
              assignedOfficer: officer,
              status: status,
              adminRemark: remark,
            })
              .then(renderList)
              .catch(function (err) {
                console.error(err);
                alert("Could not save changes. Check backend setup.");
              })
              .finally(function () {
                button.disabled = false;
              });
          });
        });
      })
      .catch(function (err) {
        console.error(err);
        container.innerHTML = '<section class="panel"><p class="error">Could not load complaints.</p></section>';
      });
  }

  function showDashboard() {
    byId("login-section").classList.add("hidden");
    byId("dashboard").classList.remove("hidden");
    renderList();
  }

  function showLogin() {
    byId("login-section").classList.remove("hidden");
    byId("dashboard").classList.add("hidden");
  }

  byId("btn-send-otp").addEventListener("click", function () {
    var mobile = byId("admin-mobile").value.replace(/\D/g, "");
    var err = byId("err-login");
    err.classList.add("hidden");
    if (mobile.length !== 10) {
      err.textContent = "Enter a valid 10-digit mobile number.";
      err.classList.remove("hidden");
      return;
    }
    if (mobile !== config().adminMobile) {
      err.textContent = "Unauthorized admin mobile.";
      err.classList.remove("hidden");
      return;
    }
    setLoggedIn(true);
    showDashboard();
  });

  byId("admin-otp-cancel").addEventListener("click", function () {
    byId("admin-otp-modal").classList.add("hidden");
  });

  byId("admin-otp-form").addEventListener("submit", function (event) {
    event.preventDefault();
    var code = byId("admin-otp-input").value.replace(/\D/g, "");
    var mobile = byId("admin-mobile").value.replace(/\D/g, "");
    var err = byId("err-admin-otp");
    err.classList.add("hidden");

    if (code.length !== 6) {
      err.textContent = "Enter 6 digits.";
      err.classList.remove("hidden");
      return;
    }

    window.MSWOtp.verify(code)
      .then(function () {
        if (mobile !== config().adminMobile) {
          err.textContent = "Unauthorized admin mobile.";
          err.classList.remove("hidden");
          return;
        }

        byId("admin-otp-modal").classList.add("hidden");
        setLoggedIn(true);
        showDashboard();
      })
      .catch(function (verifyErr) {
        console.error(verifyErr);
        err.textContent = "Invalid OTP.";
        err.classList.remove("hidden");
      });
  });

  byId("btn-logout").addEventListener("click", function () {
    setLoggedIn(false);
    showLogin();
  });

  byId("filter-from").addEventListener("change", renderList);
  byId("filter-to").addEventListener("change", renderList);
  byId("filter-clear").addEventListener("click", function () {
    byId("filter-from").value = "";
    byId("filter-to").value = "";
    renderList();
  });

  window.MSWStore.updateModeBadge();
  if (isLoggedIn()) showDashboard();
})();
