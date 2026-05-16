window.MSW_CONFIG = {
  adminMobile: "7052000032",

  // Leave these values blank to use browser demo mode.
  // After you create Firebase, paste your web app config here.
  firebaseConfig: {
    apiKey: "AIzaSyDawXR0Z3sbsMayfxPXkdgyVQxKM5ECnRk",
    authDomain: "pinaka-crm.firebaseapp.com",
    projectId: "pinaka-crm",
    messagingSenderId: "1012814380396",
    appId: "1:1012814380396:web:61480633f289d1efaf8749",
    measurementId: "G-NBW26PL4JN",
  },

  // Keep false until your Firebase Phone Auth domain and SMS setup are ready.
  useFirebasePhoneAuth: false,
  defaultCountryCode: "+91",

  wards: 20,

  officers: [
    { name: "Ramesh Sharma", ward: "Ward 7" },
    { name: "Suresh Verma", ward: "Ward 3" },
    { name: "Priya Singh", ward: "Ward 1" },
  ],
};
