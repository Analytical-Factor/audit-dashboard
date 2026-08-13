// Admin-only user creation modal for the Audit Dashboard.
(function () {
  "use strict";

  function validatePasswordPolicy(password) {
    if (!password || typeof password !== "string") return { isValid: false, message: "Password is required" };
    if (password.length < 8) return { isValid: false, message: "Password must be at least 8 characters long" };
    if (password.trim().length === 0) return { isValid: false, message: "Password cannot be only whitespace" };
    if (!/[a-z]/.test(password)) return { isValid: false, message: "Password must contain at least one lowercase letter" };
    if (!/[A-Z]/.test(password)) return { isValid: false, message: "Password must contain at least one uppercase letter" };
    if (!/\d/.test(password)) return { isValid: false, message: "Password must contain at least one number" };
    if (!/[!@#$%^&*(),.?":{}|<>]/.test(password)) return { isValid: false, message: "Password must contain at least one special character" };
    return { isValid: true, message: "Valid" };
  }

  function init({ isAdmin } = {}) {
    if (!isAdmin || document.getElementById("createUserDialog")) return;
    const userMenu = document.querySelector(".user-menu");
    const userName = document.getElementById("dashboardUser");
    if (!userMenu || !userName) return;

    const style = document.createElement("style");
    style.textContent = `
      .create-user-button { white-space: nowrap; }
      .password-policy { margin: -2px 0 0; color: var(--muted); font-size: 10px; line-height: 1.5; }
      .password-policy strong { color: var(--ink); }
      .create-user-status { min-height: 34px; border-radius: 8px; padding: 9px 11px; font-size: 11px; }
      .create-user-status:empty { display: none; }
      .create-user-status.pending { color: var(--navy); background: #edf3f6; }
      .create-user-status.success { color: #176a68; background: var(--teal-soft); }
      .create-user-status.error { color: var(--red); background: var(--red-soft); border: 1px solid var(--red); }
      .snapshot-field input.input-error { border-color: var(--red); outline: 2px solid rgba(184, 79, 74, .16); }
      .field-error { margin: 0; color: var(--red); font-size: 10px; line-height: 1.35; }
      .required { color: var(--red); margin-left: 4px; }
    `;
    document.head.append(style);

    const button = document.createElement("button");
    button.className = "button create-user-button";
    button.type = "button";
    button.textContent = "Create user";
    userMenu.insertBefore(button, userName);

    const modal = document.createElement("div");
    modal.className = "snapshot-modal";
    modal.id = "createUserDialog";
    modal.hidden = true;
    modal.innerHTML = `
      <section class="snapshot-dialog card" role="dialog" aria-modal="true" aria-labelledby="createUserDialogTitle">
        <div class="snapshot-dialog-head">
          <div><h2 id="createUserDialogTitle">Create user</h2><p>Add a dashboard user with a secure password.</p></div>
          <button class="snapshot-close" type="button" data-create-user-close aria-label="Close">&times;</button>
        </div>
        <form class="snapshot-form" id="createUserForm" novalidate>
          <div class="snapshot-field"><label for="createUserFullName">Full name<span class="required" aria-hidden="true">*</span></label><input id="createUserFullName" name="fullName" type="text" autocomplete="name" required><p class="field-error" data-error-for="createUserFullName" hidden></p></div>
          <div class="snapshot-field"><label for="createUsername">Username<span class="required" aria-hidden="true">*</span></label><input id="createUsername" name="username" type="text" autocomplete="username" required><p class="field-error" data-error-for="createUsername" hidden></p></div>
          <div class="snapshot-field"><label for="createUserPassword">Password<span class="required" aria-hidden="true">*</span></label><input id="createUserPassword" name="password" type="password" autocomplete="new-password" aria-describedby="passwordPolicy" required><p class="password-policy" id="passwordPolicy"><strong>Password policy:</strong> at least 8 characters with lowercase, uppercase, number, and special character. Common and whitespace-only passwords are not allowed.</p><p class="field-error" data-error-for="createUserPassword" hidden></p></div>
          <div class="snapshot-field"><label for="confirmUserPassword">Confirm password<span class="required" aria-hidden="true">*</span></label><input id="confirmUserPassword" name="confirmPassword" type="password" autocomplete="new-password" required><p class="field-error" data-error-for="confirmUserPassword" hidden></p></div>
          <div class="create-user-status" id="createUserStatus" role="status" aria-live="polite"></div>
          <div class="snapshot-actions"><button class="button" type="button" data-create-user-close>Cancel</button><button class="button primary" id="submitCreateUser" type="submit">Create user</button></div>
        </form>
      </section>`;
    document.body.append(modal);

    const form = modal.querySelector("form");
    const status = modal.querySelector("#createUserStatus");
    const submitButton = modal.querySelector("#submitCreateUser");
    let pending = false;

    function setStatus(message = "", type = "") {
      status.textContent = message;
      status.className = `create-user-status ${type}`.trim();
    }

    function setFieldError(id, message) {
      const input = modal.querySelector(`#${id}`);
      const error = modal.querySelector(`[data-error-for="${id}"]`);
      input.classList.toggle("input-error", Boolean(message));
      input.setAttribute("aria-invalid", String(Boolean(message)));
      error.textContent = message || "";
      error.hidden = !message;
    }

    function clearErrors() {
      modal.querySelectorAll("input").forEach(input => setFieldError(input.id, ""));
      setStatus();
    }

    function close() {
      if (pending) return;
      modal.hidden = true;
      form.reset();
      clearErrors();
    }

    function open() {
      modal.hidden = false;
      form.reset();
      clearErrors();
      modal.querySelector("#createUserFullName").focus();
    }

    button.addEventListener("click", open);
    modal.querySelectorAll("[data-create-user-close]").forEach(control => control.addEventListener("click", close));
    modal.addEventListener("click", event => { if (event.target === modal) close(); });
    document.addEventListener("keydown", event => { if (event.key === "Escape" && !modal.hidden) close(); });
    modal.querySelectorAll("input").forEach(input => input.addEventListener("input", () => setFieldError(input.id, "")));

    form.addEventListener("submit", async event => {
      event.preventDefault();
      if (pending) return;
      clearErrors();
      const fullUserName = modal.querySelector("#createUserFullName").value.trim();
      const username = modal.querySelector("#createUsername").value.trim();
      const password = modal.querySelector("#createUserPassword").value;
      const confirmation = modal.querySelector("#confirmUserPassword").value;
      const errors = [];
      const usernameValidation = /^[a-zA-Z0-9._-]{3,30}$/.test(username);
      const fullNameValidation = /^[a-zA-Z0-9 ._-]{1,50}$/.test(fullUserName || "");

      if (!fullUserName) errors.push(["createUserFullName", "Full name is required."]);
      else if (!fullNameValidation) errors.push(["createUserFullName", "Full name must be 1-50 characters and use only letters, numbers, spaces, periods, underscores, or hyphens."]);
      if (!username) errors.push(["createUsername", "Username is required."]);
      else if (!usernameValidation) errors.push(["createUsername", "Username must be 3-30 characters and use only letters, numbers, periods, underscores, or hyphens."]);
      const passwordValidation = validatePasswordPolicy(password);
      if (!passwordValidation.isValid) errors.push(["createUserPassword", passwordValidation.message]);
      if (!confirmation) errors.push(["confirmUserPassword", "Confirm password is required."]);
      else if (password !== confirmation) errors.push(["confirmUserPassword", "Passwords do not match."]);
      if (errors.length) {
        errors.forEach(([id, message]) => setFieldError(id, message));
        modal.querySelector(`#${errors[0][0]}`).focus();
        return;
      }
      if (typeof window.AuditDashboardApi?.createUser !== "function") {
        setStatus("User creation is unavailable. Reload the dashboard and try again.", "error");
        return;
      }

      pending = true;
      submitButton.disabled = true;
      setStatus("Creating user...", "pending");
      try {
        await window.AuditDashboardApi.createUser({ username, password, fullUserName });
        setStatus("User created successfully.", "success");
        form.reset();
      } catch (error) {
        setStatus(error.message || "The user could not be created.", "error");
      } finally {
        pending = false;
        submitButton.disabled = false;
      }
    });
  }

  window.CreateUserModal = { init, validatePasswordPolicy };
})();
