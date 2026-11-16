// Page identity, API endpoint selection, session state, and shared currency formatting.
const page = document.body.dataset.page;
const apiBase = location.port === "5500" ? `${location.protocol}//${location.hostname}:3000` : "";
let sessionUser = null;
const pesoFormatter = new Intl.NumberFormat("en-PH", {
  style: "currency",
  currency: "PHP",
  minimumFractionDigits: 2,
});

function localDate() {
  const date = new Date();
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function localMonth() {
  return localDate().slice(0, 7);
}

// Shared API, role visibility, form, message, and DOM-rendering helpers.
function setMessage(message, isError = false, panel = document.querySelector("main")) {
  const element = panel?.querySelector("[data-message]") || document.querySelector("[data-message]");
  if (!element) return;
  element.textContent = message;
  element.classList.toggle("message-error", isError);
}

async function api(url, options = {}) {
  const response = await fetch(`${apiBase}${url}`, {
    headers: {
      ...(options.body ? { "Content-Type": "application/json" } : {}),
      ...options.headers,
    },
    ...options,
    credentials: "include",
  });
  if (response.status === 204) return null;

  const result = await response.json();
  if (!response.ok) {
    const error = new Error(result.error || "The request could not be completed.");
    error.status = response.status;
    throw error;
  }
  return result;
}

async function downloadExcel(url) {
  const response = await fetch(`${apiBase}${url}`, { credentials: "include" });
  if (!response.ok) {
    const result = await response.json();
    throw new Error(result.error || "The Excel file could not be created.");
  }
  const blob = await response.blob();
  const contentDisposition = response.headers.get("Content-Disposition");
  const filename = contentDisposition?.match(/filename="?([^";]+)"?/i)?.[1];
  if (!filename) throw new Error("The server did not provide an Excel filename. Please refresh and try again.");
  const link = document.createElement("a");
  const objectUrl = URL.createObjectURL(blob);
  link.href = objectUrl;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
}

function isManager() {
  return sessionUser?.role === "admin" || sessionUser?.role === "hr";
}

function applyRoleVisibility() {
  document.querySelectorAll("[data-roles]").forEach((element) => {
    const allowedRoles = element.dataset.roles.split(/\s+/);
    element.hidden = !allowedRoles.includes(sessionUser.role);
  });
  document.querySelectorAll("[data-current-user]").forEach((element) => {
    element.textContent = `${sessionUser.username} (${sessionUser.accountRole})`;
  });
  if (!isManager()) {
    document.querySelectorAll("[data-management]").forEach((element) => {
      element.hidden = true;
    });
    document.querySelectorAll("[data-manager-only]").forEach((element) => {
      element.hidden = true;
    });
  }
}

function formData(form) {
  return Object.fromEntries(new FormData(form).entries());
}

function setText(id, value) {
  const element = document.getElementById(id);
  if (element) element.textContent = String(value);
}

function addCell(row, value) {
  const cell = document.createElement("td");
  cell.textContent = value === null || value === undefined || value === "" ? "—" : String(value);
  row.append(cell);
  return cell;
}

function setEmptyRow(body, columns, message) {
  const row = document.createElement("tr");
  const cell = document.createElement("td");
  cell.colSpan = columns;
  cell.textContent = message;
  row.append(cell);
  body.replaceChildren(row);
}

function formatTime(value) {
  return value ? String(value).slice(0, 5) : "—";
}

function formatMoney(value) {
  return pesoFormatter.format(Number(value) || 0);
}

function addActionButton(row, label, action, id, danger = false) {
  const cell = row.lastElementChild?.dataset.actions === "true"
    ? row.lastElementChild
    : (() => {
        const newCell = document.createElement("td");
        newCell.dataset.actions = "true";
        row.append(newCell);
        return newCell;
      })();
  const button = document.createElement("button");
  button.type = "button";
  button.textContent = label;
  button.dataset.action = action;
  button.dataset.recordId = id;
  if (danger) button.className = "button-danger";
  cell.append(button);
}

// Handle login, existing sessions, and invalid-credential messages.
async function initializeLogin() {
  const form = document.getElementById("login-form");
  const errorElement = document.getElementById("login-error");
  try {
    await api("/api/session");
    location.replace("dashboard.html");
    return;
  } catch (error) {
    if (error.status !== 401) {
      errorElement.textContent = error.message;
      errorElement.hidden = false;
    }
  }

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const submitButton = form.querySelector('button[type="submit"]');
    submitButton.disabled = true;
    errorElement.hidden = true;
    try {
      await api("/api/login", {
        method: "POST",
        body: JSON.stringify({
          username: form.elements.username.value,
          password: form.elements.password.value,
        }),
      });
      location.replace("dashboard.html");
    } catch (error) {
      errorElement.textContent = error.message;
      errorElement.hidden = false;
    } finally {
      submitButton.disabled = false;
    }
  });
}

// Load the signed-in user's dashboard summary and attendance rows.
async function initializeDashboard() {
  const date = localDate();
  const data = await api(`/api/dashboard?date=${encodeURIComponent(date)}`);
  setText("dashboard-employee-count", data.totalEmployees);
  setText("dashboard-present-count", data.presentToday);
  setText("dashboard-absent-count", data.absentToday);
  const body = document.getElementById("dashboard-attendance-rows");
  if (!data.attendance.length) {
    setEmptyRow(body, 8, "No attendance records for today.");
    return;
  }
  body.replaceChildren(...data.attendance.map((record) => {
    const row = document.createElement("tr");
    addCell(row, record.attendanceId);
    addCell(row, record.employeeId);
    addCell(row, record.date);
    addCell(row, record.isAbsent ? "Absent" : "Present");
    addCell(row, formatTime(record.timeIn));
    addCell(row, formatTime(record.timeOut));
    addCell(row, Number(record.overtime).toFixed(2));
    addCell(row, Number(record.totalWorkingHours).toFixed(2));
    return row;
  }));
}

// Search, display, create, update, and remove employee records.
async function loadEmployees(search = "") {
  const query = new URLSearchParams();
  if (search) query.set("search", search);
  return api(`/api/employees?${query}`);
}

async function initializeEmployees() {
  const form = document.getElementById("employee-form");
  const body = document.getElementById("employee-rows");
  const searchForm = document.querySelector(".search-form[role='search']");
  const searchInput = document.getElementById("employee-search");
  const positionSelect = form.elements.Position;
  let editingId = null;
  let originalEmployeeData = null;

  function resetEmployeeForm() {
    form.reset();
    positionSelect.querySelector("option[data-existing-position]")?.remove();
    editingId = null;
    originalEmployeeData = null;
    document.getElementById("employee-form-heading").textContent = "Add employee";
    form.querySelector('button[type="submit"]').textContent = "Save employee";
    document.getElementById("employee-reset").textContent = "Clear form";
    setMessage("", false, form.closest(".panel"));
  }

  if (!isManager()) {
    document.querySelector('[aria-labelledby="employee-form-heading"]').hidden = true;
    document.querySelector("#employee-rows").closest("table").querySelector("th:last-child").hidden = true;
  }

  async function render() {
    const employees = await loadEmployees(searchInput.value.trim());
    if (!employees.length) {
      setEmptyRow(body, isManager() ? 9 : 8, "No employee records found.");
    } else {
      const rows = employees.map((employee) => {
        const row = document.createElement("tr");
        addCell(row, employee.employeeId);
        addCell(row, employee.name);
        addCell(row, employee.gender);
        addCell(row, employee.birthDate);
        addCell(row, employee.contactNumber);
        addCell(row, employee.email);
        addCell(row, employee.position);
        addCell(row, employee.address);
        if (isManager()) {
          const actions = document.createElement("td");
          actions.dataset.actions = "true";
          row.append(actions);
          addActionButton(row, "Edit", "edit-employee", employee.employeeId);
          addActionButton(row, "Delete", "delete-employee", employee.employeeId, true);
        }
        return row;
      });
      body.replaceChildren(...rows);
    }

    const allEmployees = await loadEmployees();
    setText("employee-total", allEmployees.length);
    setText("employee-recent", allEmployees[0]?.name || "None");
  }

  searchForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    try {
      await render();
      setMessage("");
    } catch (error) {
      setMessage(error.message, true, searchForm.closest(".panel"));
    }
  });

  document.querySelector("[data-export='employees']")?.addEventListener("click", async () => {
    const button = document.querySelector("[data-export='employees']");
    button.disabled = true;
    try {
      await downloadExcel("/api/exports/employees.xlsx");
      setMessage("Employee directory Excel file saved.");
    } catch (error) {
      setMessage(error.message, true, searchForm.closest(".panel"));
    } finally {
      button.disabled = false;
    }
  });

  // Save employee form changes or create a new employee.
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const employeeData = formData(form);
    if (editingId !== null) {
      const hasChanges = Object.entries(originalEmployeeData || {}).some(
        ([field, value]) => employeeData[field] !== value,
      );
      if (!hasChanges) {
        setMessage("No changes detected.", false, form.closest(".panel"));
        return;
      }
      if (!window.confirm(`Update employee #${editingId} with these details?`)) return;
    }
    const panel = form.closest(".panel");
    const submitButton = form.querySelector('button[type="submit"]');
    submitButton.disabled = true;
    try {
      await api(editingId ? `/api/employees/${editingId}` : "/api/employees", {
        method: editingId ? "PUT" : "POST",
        body: JSON.stringify(employeeData),
      });
      const message = editingId === null ? "Employee added." : `Employee #${editingId} updated.`;
      resetEmployeeForm();
      await render();
      setMessage(message, false, panel);
    } catch (error) {
      setMessage(error.message, true, panel);
    } finally {
      submitButton.disabled = false;
    }
  });

  // Populate the employee form and preserve original values for change checks.
  async function loadEmployeeForEdit(employeeId) {
    const employee = (await loadEmployees(String(employeeId)))
      .find((item) => String(item.employeeId) === String(employeeId));
    if (!employee) throw new Error(`Employee #${employeeId} was not found.`);

    editingId = employee.employeeId;
    originalEmployeeData = {
      Name: employee.name,
      Gender: employee.gender,
      BirthDate: employee.birthDate,
      ContactNumber: employee.contactNumber,
      Email: employee.email,
      Position: employee.position,
      Address: employee.address,
    };
    form.elements.Name.value = originalEmployeeData.Name;
    form.elements.Gender.value = originalEmployeeData.Gender;
    form.elements.BirthDate.value = originalEmployeeData.BirthDate;
    form.elements.ContactNumber.value = originalEmployeeData.ContactNumber;
    form.elements.Email.value = originalEmployeeData.Email;
    positionSelect.querySelector("option[data-existing-position]")?.remove();
    if (![...positionSelect.options].some((option) => option.value === employee.position)) {
      const existingPosition = new Option(`Current: ${employee.position}`, employee.position);
      existingPosition.dataset.existingPosition = "true";
      positionSelect.add(existingPosition);
    }
    positionSelect.value = originalEmployeeData.Position;
    form.elements.Address.value = originalEmployeeData.Address;
    document.getElementById("employee-form-heading").textContent = `Edit employee #${employeeId}`;
    form.querySelector('button[type="submit"]').textContent = "Update employee";
    document.getElementById("employee-reset").textContent = "Cancel edit";
    form.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  document.getElementById("employee-reset").addEventListener("click", () => {
    window.setTimeout(resetEmployeeForm, 0);
  });

  body.addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-action]");
    if (!button) return;
    const panel = form.closest(".panel");
    const employeeId = button.dataset.recordId;
    try {
      if (button.dataset.action === "edit-employee") {
        await loadEmployeeForEdit(employeeId);
      } else if (button.dataset.action === "delete-employee") {
        if (!window.confirm(`Delete employee #${employeeId}, their linked employee account, and their attendance and salary records?`)) return;
        await api(`/api/employees/${employeeId}`, { method: "DELETE" });
        if (String(editingId) === employeeId) resetEmployeeForm();
        await render();
        setMessage("Employee and linked records deleted.", false, panel);
      }
    } catch (error) {
      setMessage(error.message, true, panel);
    }
  });

  await render();
}

// Calculate the weekday lateness used by the attendance summary.
function calculateLateHours(date, timeIn) {
  const day = new Date(`${date}T00:00:00Z`).getUTCDay();
  if (day === 0 || day === 6) return 0;
  const [hours, minutes] = timeIn.split(":").map(Number);
  const arrival = hours * 60 + minutes;
  if (arrival <= 480) return 0;
  if (arrival <= 720) return (arrival - 480) / 60;
  if (arrival <= 780) return 4;
  return 4 + (arrival - 780) / 60;
}

async function initializeAttendance() {
  const filterForm = document.getElementById("attendance-filter-form");
  const exportForm = document.getElementById("attendance-export-form");
  const automaticExportToggle = document.getElementById("automatic-monthly-export");
  const automaticExportState = document.getElementById("automatic-monthly-export-state");
  const form = document.getElementById("attendance-form");
  const body = document.getElementById("attendance-rows");
  const dateFilter = document.getElementById("attendance-filter-date");
  const exportMonth = document.getElementById("attendance-export-month");
  const attendanceDate = document.getElementById("attendance-date");
  let editingId = null;
  let originalAttendanceData = null;

  if (!isManager()) {
    document.querySelector('[aria-labelledby="attendance-form-heading"]').hidden = true;
    document.querySelector("#attendance-rows").closest("table").querySelector("th:last-child").hidden = true;
  }

  async function loadEmployeeOptions() {
    if (!isManager()) return;
    const select = document.getElementById("attendance-employee");
    const employees = await loadEmployees();
    for (const employee of employees) {
      const option = document.createElement("option");
      option.value = employee.employeeId;
      option.textContent = `${employee.name} (#${employee.employeeId})`;
      select.append(option);
    }
  }

  async function render(date = dateFilter.value) {
    const query = date ? `?date=${encodeURIComponent(date)}` : "";
    const panel = body.closest(".panel");
    setMessage("", false, panel);
    let savedRecords;
    let employees;
    let absenceSaveError = null;
    let savedAbsences = 0;
    const weekday = date ? new Date(`${date}T00:00:00Z`).getUTCDay() : null;
    const completedWeekday = Boolean(date && date < localDate() && weekday !== 0 && weekday !== 6);
    [savedRecords, employees] = await Promise.all([
      api(`/api/attendance${query}`),
      isManager() ? loadEmployees() : Promise.resolve([{ employeeId: sessionUser.employeeId }]),
    ]);

    if (isManager() && completedWeekday) {
      try {
        const result = await api("/api/attendance/absences", {
          method: "POST",
          body: JSON.stringify({ date }),
        });
        savedAbsences = result.created;
        if (savedAbsences > 0) {
          savedRecords = await api(`/api/attendance${query}`);
        }
      } catch (error) {
        absenceSaveError = error;
      }
    }

    const recordedEmployeeIds = new Set(savedRecords.map((record) => Number(record.employeeId)));
    const records = [
      ...savedRecords,
      ...(completedWeekday ? employees.filter((employee) => !recordedEmployeeIds.has(Number(employee.employeeId))) : [])
        .map((employee) => ({
          attendanceId: null,
          employeeId: employee.employeeId,
          date,
          timeIn: "00:00:00",
          timeOut: "00:00:00",
          overtime: 0,
          totalWorkingHours: 0,
          isAbsent: true,
        })),
    ];
    if (!records.length) {
      setEmptyRow(body, isManager() ? 9 : 8, "No employees found.");
    } else {
      body.replaceChildren(...records.map((record) => {
        const row = document.createElement("tr");
        const attendanceId = record.attendanceId;
        const idCell = addCell(row, attendanceId === null ? "— (no saved record)" : String(attendanceId));
        idCell.className = "attendance-id";
        addCell(row, record.employeeId);
        addCell(row, record.date);
        addCell(row, record.isAbsent ? "Absent" : "Present");
        addCell(row, formatTime(record.timeIn));
        addCell(row, formatTime(record.timeOut));
        addCell(row, Number(record.overtime).toFixed(2));
        addCell(row, Number(record.totalWorkingHours).toFixed(2));
        if (isManager()) {
          const actions = document.createElement("td");
          actions.dataset.actions = "true";
          row.append(actions);
          if (attendanceId !== null) {
            addActionButton(row, "Edit", "edit-attendance", String(attendanceId));
            addActionButton(row, "Delete", "delete-attendance", String(attendanceId), true);
          }
        }
        return row;
      }));
    }

    setText("attendance-scheduled", employees.length);
    const presentEmployees = new Set(records.filter((record) => !record.isAbsent).map((record) => record.employeeId));
    setText("attendance-present", presentEmployees.size);
    setText("attendance-absent", completedWeekday ? Math.max(0, employees.length - presentEmployees.size) : 0);
    setText("attendance-late", records.filter((record) => !record.isAbsent && calculateLateHours(record.date, formatTime(record.timeIn)) > 0).length);
    if (savedAbsences > 0) {
      setMessage(`Saved ${savedAbsences} missing weekday attendance record${savedAbsences === 1 ? "" : "s"} as absent.`, false, panel);
    } else if (absenceSaveError) {
      setMessage(`Absent rows are shown, but they could not be saved: ${absenceSaveError.message}`, true, panel);
    }
  }

  dateFilter.value = localDate();
  exportMonth.value = localMonth();
  attendanceDate.value = localDate();
  if (isManager()) {
    try {
      const setting = await api("/api/settings/automatic-monthly-export");
      automaticExportToggle.checked = setting.enabled;
      automaticExportState.textContent = `Auto-save monthly: ${setting.enabled ? "Yes" : "No"}`;
    } catch (error) {
      setMessage(error.message, true, body.closest(".panel"));
    }
  }
  await loadEmployeeOptions();
  await render();

  filterForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    try {
      await render();
    } catch (error) {
      setMessage(error.message, true, filterForm.closest(".panel"));
    }
  });

  exportForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const exportButton = exportForm.querySelector("button[type='submit']");
    exportButton.disabled = true;
    try {
      const month = encodeURIComponent(exportMonth.value);
      await downloadExcel(`/api/exports/attendance.xlsx?month=${month}`);
      setMessage(`Attendance Excel file for ${exportMonth.value} saved.`, false, body.closest(".panel"));
    } catch (error) {
      setMessage(error.message, true, body.closest(".panel"));
    } finally {
      exportButton.disabled = false;
    }
  });

  automaticExportToggle.addEventListener("change", async () => {
    const previousValue = !automaticExportToggle.checked;
    automaticExportToggle.disabled = true;
    try {
      const setting = await api("/api/settings/automatic-monthly-export", {
        method: "PUT",
        body: JSON.stringify({ enabled: automaticExportToggle.checked }),
      });
      automaticExportToggle.checked = setting.enabled;
      automaticExportState.textContent = `Auto-save monthly: ${setting.enabled ? "Yes" : "No"}`;
      setMessage(`Automatic monthly export ${setting.enabled ? "enabled" : "disabled"}.`, false, body.closest(".panel"));
    } catch (error) {
      automaticExportToggle.checked = previousValue;
      setMessage(error.message, true, body.closest(".panel"));
    } finally {
      automaticExportToggle.disabled = false;
    }
  });

  // Save or edit attendance and refresh the affected summaries.
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const panel = form.closest(".panel");
    const submitButton = form.querySelector('button[type="submit"]');
    const attendanceData = formData(form);
    if (editingId !== null) {
      const hasChanges = Object.entries(originalAttendanceData || {}).some(
        ([field, value]) => attendanceData[field] !== value,
      );
      if (!hasChanges) {
        setMessage("No changes detected.", false, panel);
        return;
      }
    }
    submitButton.disabled = true;
    try {
      const result = await api(editingId ? `/api/attendance/${editingId}` : "/api/attendance", {
        method: editingId ? "PUT" : "POST",
        body: JSON.stringify(attendanceData),
      });
      const savedDate = attendanceDate.value;
      form.reset();
      attendanceDate.value = savedDate || localDate();
      dateFilter.value = savedDate || localDate();
      editingId = null;
      originalAttendanceData = null;
      document.getElementById("attendance-form-heading").textContent = "Add attendance record";
      submitButton.textContent = "Save attendance";
      await render(dateFilter.value);
      const actionMessage = result.action === "updated"
        ? "An attendance record already existed for this employee and date, so it was updated with the latest times."
        : "A new attendance record was added.";
      setMessage(`${actionMessage} Monthly salary totals have been recalculated.`, false, panel);
      await refreshPayrollIfOpen();
    } catch (error) {
      setMessage(error.message, true, panel);
    } finally {
      submitButton.disabled = false;
    }
  });

  document.getElementById("attendance-reset").addEventListener("click", () => {
    editingId = null;
    originalAttendanceData = null;
    document.getElementById("attendance-form-heading").textContent = "Add attendance record";
    form.querySelector('button[type="submit"]').textContent = "Save attendance";
    setMessage("", false, form.closest(".panel"));
    window.setTimeout(() => { attendanceDate.value = localDate(); }, 0);
  });

  body.addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-action]");
    if (!button) return;
    const panel = form.closest(".panel");
    const attendanceId = button.dataset.recordId;
    try {
      if (button.dataset.action === "edit-attendance") {
        const record = (await api("/api/attendance")).find((item) => String(item.attendanceId) === attendanceId);
        if (!record) throw new Error("Attendance record not found.");
        editingId = record.attendanceId;
        originalAttendanceData = {
          EmployeeID: String(record.employeeId),
          Date: String(record.date).slice(0, 10),
          TimeIn: formatTime(record.timeIn),
          TimeOut: formatTime(record.timeOut),
        };
        form.elements.EmployeeID.value = originalAttendanceData.EmployeeID;
        form.elements.Date.value = originalAttendanceData.Date;
        form.elements.TimeIn.value = originalAttendanceData.TimeIn;
        form.elements.TimeOut.value = originalAttendanceData.TimeOut;
        document.getElementById("attendance-form-heading").textContent = `Edit attendance #${attendanceId}`;
        form.querySelector('button[type="submit"]').textContent = "Update attendance";
        form.scrollIntoView({ behavior: "smooth", block: "start" });
      } else if (button.dataset.action === "delete-attendance") {
        if (!window.confirm(`Delete attendance record #${attendanceId}? Monthly salary will be recalculated.`)) return;
        await api(`/api/attendance/${attendanceId}`, { method: "DELETE" });
        await render();
        setMessage("Attendance deleted. Monthly salary totals have been recalculated.", false, panel);
        await refreshPayrollIfOpen();
      }
    } catch (error) {
      setMessage(error.message, true, panel);
    }
  });
}

// Load payroll rows allowed by the signed-in role and calculate display totals.
async function refreshPayrollIfOpen() {
  if (document.body.dataset.page === "payroll") await loadPayroll(document.getElementById("payroll-month").value);
}

async function loadPayroll(month) {
  const result = await api(`/api/salaries?month=${encodeURIComponent(month)}`);
  const body = document.getElementById("payroll-rows");
  if (!result.salaries.length) {
    setEmptyRow(body, 11, "No salary records found.");
  } else {
    body.replaceChildren(...result.salaries.map((salary) => {
      const row = document.createElement("tr");
      addCell(row, salary.salaryId);
      addCell(row, salary.employeeId);
      addCell(row, formatMoney(salary.basicSalary));
      addCell(row, Number(salary.totalWorkingHours).toFixed(2));
      addCell(row, Number(salary.overtimeHours).toFixed(2));
      addCell(row, formatMoney(salary.overtimePay));
      addCell(row, salary.absentDays);
      addCell(row, formatMoney(salary.absenceDeduction));
      addCell(row, formatMoney(salary.lateDeduction));
      addCell(row, formatMoney(salary.deductions));
      addCell(row, formatMoney(salary.netSalary));
      return row;
    }));
  }

  const totals = result.salaries.reduce((sum, salary) => ({
    overtimePay: sum.overtimePay + Number(salary.overtimePay),
    absenceDeduction: sum.absenceDeduction + Number(salary.absenceDeduction),
    deductions: sum.deductions + Number(salary.deductions),
    netSalary: sum.netSalary + Number(salary.netSalary),
  }), { overtimePay: 0, absenceDeduction: 0, deductions: 0, netSalary: 0 });
  setText("salary-count", result.salaries.length);
  setText("overtime-pay-total", formatMoney(totals.overtimePay));
  setText("absence-deduction-total", formatMoney(totals.absenceDeduction));
  setText("deductions-total", formatMoney(totals.deductions));
  setText("net-salary-total", formatMoney(totals.netSalary));
  setText("payroll-month-label", `Monthly salary records for ${result.month}.`);
}

async function initializePayroll() {
  const filterForm = document.getElementById("payroll-filter-form");
  const monthInput = document.getElementById("payroll-month");
  monthInput.value = localMonth();
  await loadPayroll(monthInput.value);
  filterForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    try {
      await loadPayroll(monthInput.value);
      setMessage("");
    } catch (error) {
      setMessage(error.message, true, filterForm.closest(".panel"));
    }
  });
}

// Load employee account settings and handle username/password changes.
async function initializeSettings() {
  const form = document.getElementById("account-settings-form");
  const submitButton = form.querySelector('button[type="submit"]');
  const passwordToggle = form.querySelector("[data-toggle-password]");
  const currentPasswordInput = document.getElementById(passwordToggle.dataset.togglePassword);
  const account = await api("/api/account-settings");
  form.elements.username.value = account.username;
  form.elements.accountRole.value = account.accountRole;

  passwordToggle.addEventListener("click", () => {
    const showPassword = currentPasswordInput.type === "password";
    currentPasswordInput.type = showPassword ? "text" : "password";
    passwordToggle.setAttribute("aria-label", `${showPassword ? "Hide" : "Show"} current password`);
    passwordToggle.setAttribute("aria-pressed", String(showPassword));
    passwordToggle.querySelector("[data-eye-icon]").hidden = showPassword;
    passwordToggle.querySelector("[data-eye-off-icon]").hidden = !showPassword;
  });

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    submitButton.disabled = true;
    try {
      const updatedAccount = await api("/api/account-settings", {
        method: "PUT",
        body: JSON.stringify(formData(form)),
      });
      sessionUser.username = updatedAccount.username;
      document.querySelectorAll("[data-current-user]").forEach((element) => {
        element.textContent = `${sessionUser.username} (${sessionUser.accountRole})`;
      });
      form.elements.currentPassword.value = "";
      form.elements.newPassword.value = "";
      form.elements.confirmPassword.value = "";
      setMessage("Account settings saved. Your account role has not changed.", false, form.closest(".panel"));
    } catch (error) {
      setMessage(error.message, true, form.closest(".panel"));
    } finally {
      submitButton.disabled = false;
    }
  });
}

// Manage administrator accounts and monthly attendance export downloads.
async function initializeAdmin() {
  const accountForm = document.getElementById("account-form");
  const body = document.getElementById("user-account-rows");
  const missingAccountsBody = document.getElementById("employees-without-accounts-rows");
  const monthlyExportsBody = document.getElementById("monthly-attendance-export-rows");
  const usernameInput = accountForm.elements.username;
  const employeeIdInput = accountForm.elements.employeeId;
  const roleSelect = accountForm.elements.accountRole;
  const submitButton = accountForm.querySelector("button[type='submit']");
  const cancelButton = document.getElementById("account-cancel");

  function resetAccountForm() {
    accountForm.reset();
    employeeIdInput.value = "";
    usernameInput.readOnly = false;
    roleSelect.querySelector("option[value='Employee']")?.remove();
    document.getElementById("add-account-heading").textContent = "Add account";
    submitButton.textContent = "Add account";
  }

  async function loadUsers() {
    const [users, employeesWithoutAccounts] = await Promise.all([
      api("/api/users"),
      api("/api/employees/without-account"),
    ]);
    const roleName = (role) => String(role).trim().toLowerCase();
    const superAdminCount = users.filter((user) =>
      ["super admin", "super administrator"].includes(roleName(user.accountRole))).length;
    const hrCount = users.filter((user) =>
      ["hr", "hr personnel", "administrator / hr personnel"].includes(roleName(user.accountRole))).length;

    setText("account-total", users.length);
    setText("super-admin-total", superAdminCount);
    setText("hr-total", hrCount);

    if (!employeesWithoutAccounts.length) {
      setEmptyRow(missingAccountsBody, 5, "All employees have a matching account username.");
    } else {
      missingAccountsBody.replaceChildren(...employeesWithoutAccounts.map((employee) => {
        const row = document.createElement("tr");
        addCell(row, employee.employeeId);
        addCell(row, employee.name);
        addCell(row, employee.email);
        addCell(row, employee.suggestedUsername);
        const actionCell = document.createElement("td");
        const createButton = document.createElement("button");
        createButton.type = "button";
        createButton.dataset.action = "create-employee-account";
        createButton.dataset.employeeId = employee.employeeId;
        createButton.dataset.username = employee.suggestedUsername;
        createButton.textContent = "Create account";
        actionCell.append(createButton);
        row.append(actionCell);
        return row;
      }));
    }

    if (!users.length) {
      setEmptyRow(body, 4, "No user accounts found.");
      return;
    }

    body.replaceChildren(...users.map((user) => {
      const row = document.createElement("tr");
      addCell(row, user.userId);
      addCell(row, user.username);
      addCell(row, user.accountRole);
      const actionCell = document.createElement("td");
      const isCurrentUser = Number(user.userId) === Number(sessionUser.userId);
      if (isCurrentUser) {
        actionCell.textContent = "Signed in";
      } else {
        const deleteButton = document.createElement("button");
        deleteButton.type = "button";
        deleteButton.className = "button-danger";
        deleteButton.dataset.action = "delete-user";
        deleteButton.dataset.recordId = user.userId;
        deleteButton.textContent = "Delete";
        deleteButton.setAttribute("aria-label", `Delete account ${user.username}`);
        actionCell.append(deleteButton);
      }
      row.append(actionCell);
      return row;
    }));
  }

  async function loadMonthlyAttendanceExports() {
    const exports = await api("/api/exports/attendance/monthly");
    if (!exports.length) {
      setEmptyRow(monthlyExportsBody, 3, "No automatic monthly exports have been generated yet.");
      return;
    }

    monthlyExportsBody.replaceChildren(...exports.map((file) => {
      const row = document.createElement("tr");
      addCell(row, file.month);
      addCell(row, file.filename);
      const actionCell = document.createElement("td");
      const link = document.createElement("a");
      link.href = `${apiBase}/api/exports/attendance/monthly/${encodeURIComponent(file.month)}.xlsx`;
      link.textContent = "Download";
      actionCell.append(link);
      row.append(actionCell);
      return row;
    }));
  }

  await loadUsers();
  await loadMonthlyAttendanceExports();
  cancelButton.addEventListener("click", () => {
    resetAccountForm();
    setMessage("Account creation cancelled.", false, accountForm.closest(".panel"));
  });

  body.addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-action='delete-user']");
    if (!button) return;
    const username = button.closest("tr").children[1].textContent;
    if (!window.confirm(`Delete the account "${username}"? A linked employee record will remain but be unlinked.`)) return;

    button.disabled = true;
    try {
      await api(`/api/users/${encodeURIComponent(button.dataset.recordId)}`, { method: "DELETE" });
      await loadUsers();
      setMessage(`Account "${username}" deleted.`, false, body.closest(".panel"));
    } catch (error) {
      setMessage(error.message, true, body.closest(".panel"));
    } finally {
      button.disabled = false;
    }
  });

  missingAccountsBody.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action='create-employee-account']");
    if (!button) return;

    employeeIdInput.value = button.dataset.employeeId;
    usernameInput.value = button.dataset.username;
    usernameInput.readOnly = true;
    if (!roleSelect.querySelector("option[value='Employee']")) {
      roleSelect.add(new Option("Employee", "Employee"));
    }
    roleSelect.value = "Employee";
    document.getElementById("add-account-heading").textContent = `Create account for ${button.closest("tr").children[1].textContent}`;
    submitButton.textContent = "Create employee account";
    accountForm.scrollIntoView({ behavior: "smooth", block: "center" });
    accountForm.elements.password.focus();
  });

  accountForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    submitButton.disabled = true;
    cancelButton.disabled = true;
    try {
      await api("/api/users", {
        method: "POST",
        body: JSON.stringify(formData(accountForm)),
      });
      const creatingEmployeeAccount = Boolean(employeeIdInput.value);
      resetAccountForm();
      await loadUsers();
      setMessage(creatingEmployeeAccount ? "Employee account created and linked." : "Account added successfully.");
    } catch (error) {
      setMessage(error.message, true, accountForm.closest(".panel"));
    } finally {
      submitButton.disabled = false;
      cancelButton.disabled = false;
    }
  });
}

async function initialize() {
  if (page === "login") {
    await initializeLogin();
    return;
  }

  try {
    sessionUser = await api("/api/session");
  } catch {
    location.replace("login.html");
    return;
  }

  const restrictedPage = (sessionUser.role === "employee" && page === "admin")
    || (sessionUser.role !== "employee" && page === "settings");
  if (restrictedPage) {
    location.replace("dashboard.html");
    return;
  }

  applyRoleVisibility();
  document.body.dataset.authorized = "true";
  try {
    if (page === "dashboard") await initializeDashboard();
    if (page === "employees") await initializeEmployees();
    if (page === "attendance") await initializeAttendance();
    if (page === "payroll") await initializePayroll();
    if (page === "settings") await initializeSettings();
    if (page === "admin") await initializeAdmin();
  } catch (error) {
    if (error.status === 401) {
      location.replace("login.html");
      return;
    }
    if (page === "attendance") {
      setEmptyRow(
        document.getElementById("attendance-rows"),
        isManager() ? 9 : 8,
        `Attendance could not be loaded: ${error.message}`,
      );
    }
    setMessage(error.message, true);
  }
}

// End the server session before returning to the login page.
document.querySelectorAll("[data-logout]").forEach((link) => {
  link.addEventListener("click", async (event) => {
    event.preventDefault();
    if (!window.confirm("Are you sure you want to log out?")) return;
    try {
      await api("/api/logout", { method: "POST" });
      location.replace("login.html");
    } catch (error) {
      setMessage(error.message, true);
    }
  });
});

initialize();