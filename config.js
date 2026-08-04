(function () {
  "use strict";

  const isLocal = ["localhost", "127.0.0.1"].includes(window.location.hostname);

  window.auditDashboardConfig = {
    apiBaseUrl: isLocal
      ? "http://localhost:3000"
      : "https://dpafprd.emrsn.com/api"
  };
})();
