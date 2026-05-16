(function () {
  var complaintTypes = [
    { value: "Garbage not collected", label: "Garbage not collected / कचरा नहीं उठाया" },
    { value: "Overflowing dustbin", label: "Overflowing dustbin / कूड़ेदान भरा हुआ" },
    { value: "Illegal dumping", label: "Illegal dumping / अवैध कचरा फेंकना" },
    { value: "Bad odor / unhygienic area", label: "Bad odor / गंदा क्षेत्र" },
    { value: "Missed pickup schedule", label: "Missed pickup schedule / समय पर सफाई नहीं हुई" },
  ];

  function byId(id) {
    return document.getElementById(id);
  }

  function fillWards() {
    var sel = byId("ward");
    var count = (window.MSW_CONFIG && window.MSW_CONFIG.wards) || 20;
    for (var i = 1; i <= count; i += 1) {
      var opt = document.createElement("option");
      opt.value = "Ward " + i;
      opt.textContent = "Ward " + i + " / वार्ड " + i;
      sel.appendChild(opt);
    }
  }

  function fillTypes() {
    var sel = byId("type");
    complaintTypes.forEach(function (item) {
      var opt = document.createElement("option");
      opt.value = item.value;
      opt.textContent = item.label;
      sel.appendChild(opt);
    });
  }

  function hideErrors() {
    document.querySelectorAll(".error").forEach(function (el) {
      el.classList.add("hidden");
      el.textContent = "";
    });
  }

  function showError(id, message) {
    var el = byId(id);
    el.textContent = message;
    el.classList.remove("hidden");
  }

  function validate() {
    hideErrors();
    var ok = true;
    var name = byId("name").value.trim();
    var mobile = byId("mobile").value.replace(/\D/g, "");
    var ward = byId("ward").value;
    var type = byId("type").value;
    var desc = byId("description").value.trim();

    if (!name) {
      showError("err-name", "Full name is required.");
      ok = false;
    }
    if (mobile.length !== 10) {
      showError("err-mobile", "Enter a valid 10-digit mobile number.");
      ok = false;
    }
    if (!ward) {
      showError("err-ward", "Please select a ward.");
      ok = false;
    }
    if (!type) {
      showError("err-type", "Please select complaint type.");
      ok = false;
    }
    if (!desc) {
      showError("err-desc", "Description is required.");
      ok = false;
    }
    return ok;
  }

  function setBusy(isBusy) {
    byId("btn-submit").disabled = isBusy;
    byId("otp-verify").disabled = isBusy;
  }

  byId("complaint-form").addEventListener("submit", function (event) {
    event.preventDefault();
    if (!validate()) return;

    var payload = {
      name: byId("name").value.trim(),
      mobile: byId("mobile").value.replace(/\D/g, ""),
      ward: byId("ward").value,
      type: byId("type").value,
      description: byId("description").value.trim(),
    };

    setBusy(true);
    window.MSWStore.createComplaint(payload)
      .then(function (record) {
        byId("complaint-form").classList.add("hidden");
        byId("success-screen").classList.remove("hidden");
        byId("success-id").textContent = record.complaintId;
        byId("complaint-form").reset();
      })
      .catch(function (err) {
        console.error(err);
        showError(
          "err-form",
          "Could not send OTP: " +
            ((err && err.code) || "") +
            " " +
            ((err && err.message) || "Check Firebase Phone Auth setup.")
        );
      })
      .finally(function () {
        setBusy(false);
      });
  });

  byId("otp-cancel").addEventListener("click", function () {
    byId("otp-modal").classList.add("hidden");
  });

  byId("otp-form").addEventListener("submit", function (event) {
    event.preventDefault();
    var code = byId("otp-input").value.replace(/\D/g, "");
    if (code.length !== 6) {
      showError("err-otp", "Enter 6 digits.");
      return;
    }
    setBusy(true);

    window.MSWOtp.verify(code)
      .then(function () {
        return null;
      })
      .then(function (record) {
        byId("otp-modal").classList.add("hidden");
        byId("complaint-form").classList.add("hidden");
        byId("success-screen").classList.remove("hidden");
        byId("success-id").textContent = record ? record.complaintId : "";
        byId("complaint-form").reset();
      })
      .catch(function (err) {
        console.error(err);
        showError("err-otp", "Could not save complaint. Check Firebase setup or try demo mode.");
      })
      .finally(function () {
        setBusy(false);
      });
  });

  byId("btn-another").addEventListener("click", function () {
    byId("success-screen").classList.add("hidden");
    byId("complaint-form").classList.remove("hidden");
  });

  fillWards();
  fillTypes();
  window.MSWStore.updateModeBadge();
})();
