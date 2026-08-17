// Authentication and snapshot-data client used by the dashboard pages.
(function () {
  "use strict";

  const TOKEN_KEY = "auditDashboardAccessToken";
  const USER_KEY = "auditDashboardUser";
  const SESSION_KEY = "auditDashboardSessionId";
  const SOURCE_COLORS = {
    "100% Hybrid Forecast": "#e49a3a",
    "100% Final Forecast": "#99a9b8",
    "Mixed Forecast Source": "#2f6fed",
    "Volume-adjusted / scaled blend": "#0b7f82",
    "No active blend weight": "#b55f62"
  };
  const SOURCE_ORDER = Object.keys(SOURCE_COLORS);

  function apiUrl(path) {
    const configuredBase = window.auditDashboardConfig?.apiBaseUrl;
    if (!configuredBase) throw new Error("Audit Dashboard API URL is not configured.");
    return `${configuredBase.replace(/\/$/, "")}${path}`;
  }

  function readStoredJson(key) {
    try {
      return JSON.parse(sessionStorage.getItem(key));
    } catch (error) {
      return null;
    }
  }

  function storeSession(response) {
    sessionStorage.setItem(TOKEN_KEY, response.accessToken);
    sessionStorage.setItem(USER_KEY, JSON.stringify(response.user));
    sessionStorage.setItem(SESSION_KEY, response.sessionId);
  }

  function clearSession() {
    sessionStorage.removeItem(TOKEN_KEY);
    sessionStorage.removeItem(USER_KEY);
    sessionStorage.removeItem(SESSION_KEY);
  }

  function hasSession() {
    return Boolean(sessionStorage.getItem(TOKEN_KEY) && readStoredJson(USER_KEY));
  }

  async function parseError(response) {
    try {
      const body = await response.json();
      return body.message || body.error || `Request failed with status ${response.status}`;
    } catch (error) {
      return `Request failed with status ${response.status}`;
    }
  }

  async function login(username, password) {
    const response = await fetch(apiUrl("/api/planning/audit-dashboard/login"), {
      method: "POST",
      credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username, password })
    });
    if (!response.ok) {
      if (response.status === 401) {
        throw new Error("Invalid username or password.");
      }
      throw new Error(await parseError(response));
    }
    const body = await response.json();
    storeSession(body);
    return body;
  }

  async function refreshSession() {
    const user = readStoredJson(USER_KEY);
    if (!user?.id) return false;

    const response = await fetch(apiUrl("/refresh_token"), {
      method: "POST",
      credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ userID: user.id, isActive: "true" })
    });
    if (!response.ok) return false;
    storeSession(await response.json());
    return true;
  }

  async function request(path, options = {}, retry = true) {
    const headers = new Headers(options.headers || {});
    const token = sessionStorage.getItem(TOKEN_KEY);
    if (token) headers.set("Authorization", `Bearer ${token}`);

    const response = await fetch(apiUrl(path), {
      ...options,
      credentials: "include",
      headers
    });
    if (response.ok) return response.json();

    const message = await parseError(response);
    const tokenRejected = response.status === 401 ||
      (response.status === 400 && /authorization token/i.test(message));
    if (retry && tokenRejected && await refreshSession()) {
      return request(path, options, false);
    }
    if (tokenRejected) clearSession();
    const requestError = new Error(message);
    requestError.status = response.status;
    throw requestError;
  }

  async function logout() {
    const user = readStoredJson(USER_KEY);
    const sessionId = sessionStorage.getItem(SESSION_KEY);
    try {
      if (user?.id && sessionId) {
        await request("/api/preferences/users/logout", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ id: user.id, sessionLogout: sessionId })
        });
      }
    } catch (error) {
      // Local logout must still complete when the API is unavailable.
    } finally {
      clearSession();
    }
  }

  function emptyPlanner() {
    return {
      reconciled: { withOverrides: 0, withoutOverrides: 0 },
      exceptions: { withOverrides: 0, withoutOverrides: 0 }
    };
  }

  function emptyReconciliationOverrides() {
    return {
      automatic: { withOverrides: 0, withoutOverrides: 0 },
      manual: { withOverrides: 0, withoutOverrides: 0 }
    };
  }

  function sourceRows(items) {
    const counts = new Map(SOURCE_ORDER.map(label => [label, {
      label,
      exceptions: 0,
      reconciled: 0,
      count: 0,
      color: SOURCE_COLORS[label]
    }]));
    items.forEach(item => {
      const source = counts.get(item.forecastSourceClassification) || {
        label: item.forecastSourceClassification,
        exceptions: 0,
        reconciled: 0,
        count: 0,
        color: "#6b7f8e"
      };
      source.count += 1;
      if (item.reconciled === 0) source.exceptions += 1;
      else source.reconciled += 1;
      counts.set(source.label, source);
    });
    return [...counts.values()];
  }

  function summarizeItems(items) {
    const total = items.length;
    const reconciled = items.filter(item => item.reconciled >= 1).length;
    const exceptions = total - reconciled;
    const automatic = items.filter(item => item.reconciled === 2).length;
    const manual = items.filter(item => item.reconciled === 1 || item.reconciled === 3).length;
    const planner = emptyPlanner();
    const reconciliationOverrides = emptyReconciliationOverrides();

    items.forEach(item => {
      const plannerGroup = item.reconciled === 0 ? planner.exceptions : planner.reconciled;
      plannerGroup[item.hasUserOverride ? "withOverrides" : "withoutOverrides"] += 1;
      if (item.reconciled === 2) {
        reconciliationOverrides.automatic[
          item.hasUserOverride ? "withOverrides" : "withoutOverrides"
        ] += 1;
      } else if (item.reconciled === 1 || item.reconciled === 3) {
        reconciliationOverrides.manual[
          item.hasUserOverride ? "withOverrides" : "withoutOverrides"
        ] += 1;
      }
    });

    return {
      total,
      reconciled,
      exceptions,
      automatic,
      manual,
      originalExceptions: exceptions + manual,
      reconciliationOverrides,
      planner,
      overrides: planner.reconciled.withOverrides + planner.exceptions.withOverrides,
      sources: sourceRows(items)
    };
  }

  function summaryFromStatusRows(statusRows) {
    const total = statusRows.reduce((sum, row) => sum + row.auditedItems, 0);
    const reconciled = statusRows.reduce((sum, row) => sum + row.reconciledItems, 0);
    const exceptions = statusRows.reduce((sum, row) => sum + row.exceptionItems, 0);
    return {
      total,
      reconciled,
      exceptions,
      automatic: 0,
      manual: 0,
      originalExceptions: exceptions,
      reconciliationOverrides: emptyReconciliationOverrides(),
      planner: emptyPlanner(),
      overrides: 0,
      sources: []
    };
  }

  function createDashboardData(monthResponse) {
    const months = monthResponse.months.map(apiMonth => {
      const abc = apiMonth.statusByAbc.map(row => ({
        label: row.abcClass,
        total: row.auditedItems,
        reconciled: row.reconciledItems,
        exceptions: row.exceptionItems
      }));
      const kpiByAbc = Object.fromEntries(apiMonth.statusByAbc.map(row => [
        row.abcClass,
        {
          ...summaryFromStatusRows([row]),
          abc: [{
            label: row.abcClass,
            total: row.auditedItems,
            reconciled: row.reconciledItems,
            exceptions: row.exceptionItems
          }]
        }
      ]));
      return {
        key: apiMonth.targetMonthStart.slice(0, 7),
        targetMonthStart: apiMonth.targetMonthStart,
        cycleDate: apiMonth.cycleDate,
        generatedAt: apiMonth.generatedAtEpochMs,
        generatedBy: apiMonth.generatedBy,
        label: apiMonth.label,
        year: String(apiMonth.year),
        ...summaryFromStatusRows(apiMonth.statusByAbc),
        abc,
        kpiByAbc,
        loaded: false
      };
    });
    return {
      currentMonth: monthResponse.currentMonth || null,
      months,
      partsByMonth: {}
    };
  }

  function applyMonthItems(data, targetMonthStart, items) {
    const month = data.months.find(candidate => candidate.targetMonthStart === targetMonthStart);
    if (!month) return;

    const summary = summarizeItems(items);
    const groups = new Map();
    items.forEach(item => {
      if (!groups.has(item.abcClass)) groups.set(item.abcClass, []);
      groups.get(item.abcClass).push(item);
    });
    const kpiByAbc = Object.fromEntries([...groups.entries()].map(([abcClass, rows]) => [
      abcClass,
      {
        ...summarizeItems(rows),
        abc: [{
          label: abcClass,
          total: rows.length,
          reconciled: rows.filter(row => row.reconciled >= 1).length,
          exceptions: rows.filter(row => row.reconciled === 0).length
        }]
      }
    ]));
    const abc = [...groups.entries()]
      .map(([label, rows]) => ({
        label,
        total: rows.length,
        reconciled: rows.filter(row => row.reconciled >= 1).length,
        exceptions: rows.filter(row => row.reconciled === 0).length
      }))
      .sort((left, right) => left.label.localeCompare(right.label));

    Object.assign(month, summary, { abc, kpiByAbc, loaded: true });
    data.partsByMonth[month.key] = items.map(item => [
      item.item,
      item.itemDescription || "",
      item.abcClass,
      item.monthStatus,
      item.forecastSourceClassification,
      item.hasUserOverride,
      item.itemId,
      item.locationId,
      item.organization || ""
    ]);
  }

  async function ensureMonths(data, monthKeys) {
    const requested = [...new Set(monthKeys)]
      .map(key => data.months.find(month => month.key === key))
      .filter(month => month && !month.loaded);
    if (requested.length === 0) return;

    const query = requested.map(month => month.targetMonthStart).join(",");
    const response = await request(
      `/api/planning/audit-dashboard/data?months=${encodeURIComponent(query)}`
    );
    const itemsByMonth = new Map(requested.map(month => [month.targetMonthStart, []]));
    response.items.forEach(item => itemsByMonth.get(item.targetMonthStart)?.push(item));
    requested.forEach(month => {
      applyMonthItems(data, month.targetMonthStart, itemsByMonth.get(month.targetMonthStart));
    });
  }

  async function loadAvailableMonths() {
    return request("/api/planning/audit-dashboard/months");
  }

  async function loadDashboardData() {
    const response = await loadAvailableMonths();
    const data = createDashboardData(response);
    if (data.months.length === 0) return data;
    const latest = data.months.at(-1);
    const previous = data.months.at(-2);
    await ensureMonths(data, [latest?.key, previous?.key].filter(Boolean));
    return data;
  }

  async function generateSnapshot(replaceExisting = false) {
    return request("/api/planning/audit-dashboard/snapshots", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ replaceExisting })
    });
  }

  window.AuditDashboardApi = {
    clearSession,
    ensureMonths,
    generateSnapshot,
    hasSession,
    loadAvailableMonths,
    loadDashboardData,
    login,
    logout
  };
})();
