(function () {
  var confirmation = null;
  var recaptcha = null;

  function getConfig() {
    return window.MSW_CONFIG || {};
  }

  function hasFirebaseConfig() {
    var cfg = getConfig().firebaseConfig || {};
    return Boolean(cfg.apiKey && cfg.projectId && cfg.appId);
  }

  function shouldUseFirebaseOtp() {
    return Boolean(getConfig().useFirebasePhoneAuth && hasFirebaseConfig() && window.firebase && window.firebase.auth);
  }

  function initFirebaseApp() {
    if (!hasFirebaseConfig() || !window.firebase) return false;
    if (!window.firebase.apps.length) {
      window.firebase.initializeApp(getConfig().firebaseConfig);
    }
    return true;
  }

  function formatPhone(mobile) {
    var country = getConfig().defaultCountryCode || "+91";
    return country + String(mobile || "").replace(/\D/g, "");
  }

  function getVerifier() {
    if (recaptcha) return recaptcha;
    var containerId = "recaptcha-container";
    recaptcha = new window.firebase.auth.RecaptchaVerifier(containerId, {
      size: "invisible",
    });
    return recaptcha;
  }

  function withTimeout(promise, message) {
    var timer = null;
    var timeout = new Promise(function (_, reject) {
      timer = window.setTimeout(function () {
        reject(new Error(message));
      }, 15000);
    });

    return Promise.race([promise, timeout]).finally(function () {
      window.clearTimeout(timer);
    });
  }

  function send(mobile) {
    if (!shouldUseFirebaseOtp()) {
      confirmation = null;
      return Promise.resolve({ mode: "demo" });
    }

    initFirebaseApp();
    return withTimeout(
      window.firebase.auth().signInWithPhoneNumber(formatPhone(mobile), getVerifier()),
      "Firebase reCAPTCHA did not respond. Reload the page and check internet access."
    )
      .then(function (result) {
        confirmation = result;
        return { mode: "firebase" };
      });
  }

  function verify(code) {
    if (!confirmation) {
      return Promise.resolve({ mode: "demo" });
    }
    return confirmation.confirm(code).then(function (result) {
      return { mode: "firebase", user: result.user };
    });
  }

  window.MSWOtp = {
    send: send,
    verify: verify,
    shouldUseFirebaseOtp: shouldUseFirebaseOtp,
  };
})();
