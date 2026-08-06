// Selects the API origin for local development and deployed environments.
(function () {
  "use strict";

  const isLocal = ["localhost", "127.0.0.1"].includes(window.location.hostname);

  window.auditDashboardConfig = {
    apiBaseUrl: isLocal
      ? "http://localhost:3000"
      : `${window.location.origin}/api`
  };
})();
