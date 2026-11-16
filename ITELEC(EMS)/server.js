require("dotenv").config();

const path = require("node:path");
const crypto = require("node:crypto");
const fs = require("node:fs/promises");
const { promisify } = require("node:util");
const express = require("express");
const mysql = require("mysql2/promise");
const ExcelJS = require("exceljs");

// Runtime configuration, salary rates, session storage, and the MySQL connection pool.
const scrypt = promisify(crypto.scrypt);
const app = express();
const port = Number(process.env.PORT || 3000);
const publicDirectory = path.join(__dirname, "public");
const exportsDirectory = path.join(__dirname, "exports");
const automaticMonthlyExportSetting = "automatic_monthly_attendance_export";
const automaticMonthlyExportSettingsTable = "tbl_monthly_attendance_export_settings";
const legacySettingsTable = "tbl_ems_settings";
const basicSalary = 25000;
const overtimeRate = 200;
const lateDeductionRate = 300;
const sessionLifetimeMs = 8 * 60 * 60 * 1000;
const sessions = new Map();

const pool = mysql.createPool({
  host: process.env.DB_HOST || "127.0.0.1",
  port: Number(process.env.DB_PORT || 3306),
  user: process.env.DB_USER || "root",
  password: process.env.DB_PASSWORD || "",
  database: process.env.DB_NAME || "employeemanagementsystem_db",
  waitForConnections: true,
  connectionLimit: 10,
  dateStrings: true,
});

// Shared authentication, authorization, cookie, and request-validation helpers.
function httpError(status, message) {
  const error = new Error(message);
  error.status = status;
  return error;
}

function getCookie(request, name) {
  const cookies = request.headers.cookie || "";
  const prefix = `${name}=`;
  const entry = cookies.split(";").map((cookie) => cookie.trim()).find((cookie) => cookie.startsWith(prefix));
  return entry ? decodeURIComponent(entry.slice(prefix.length)) : null;
}

function requireRoles(...roles) {
  return (request, response, next) => {
    if (!roles.includes(request.user.role)) {
      return response.status(403).json({ error: "You do not have permission to perform this action." });
    }
    next();
  };
}

function authenticate(request, response, next) {
  const sessionId = getCookie(request, "ems_session");
  const session = sessionId ? sessions.get(sessionId) : null;
  if (!session || session.expiresAt <= Date.now()) {
    if (sessionId) sessions.delete(sessionId);
    return response.status(401).json({ error: "Please sign in to continue." });
  }
  request.user = session.user;
  session.expiresAt = Date.now() + sessionLifetimeMs;
  next();
}

function normalizeRole(accountRole) {
  const role = String(accountRole).trim().toLowerCase();
  if (["super admin", "super administrator", "admin"].includes(role)) return "admin";
  if (["hr", "hr personnel", "administrator / hr personnel"].includes(role)) return "hr";
  if (role === "employee") return "employee";
  return null;
}

function parsePasswordHash(passwordHash) {
  if (typeof passwordHash !== "string") return null;
  const [algorithm, saltText, hashText, extra] = passwordHash.split("$");
  if (algorithm !== "scrypt" || !saltText || !hashText || extra !== undefined) return null;
  const salt = Buffer.from(saltText, "base64url");
  const expectedHash = Buffer.from(hashText, "base64url");
  if (
    salt.length !== 16
    || expectedHash.length !== 64
    || salt.toString("base64url") !== saltText
    || expectedHash.toString("base64url") !== hashText
  ) {
    return null;
  }
  return { salt, expectedHash };
}

function normalizeAccountUsername(value) {
  return String(value).toLocaleLowerCase().replace(/[^a-z0-9]/g, "");
}

function employeeUsername(name, employeeId) {
  const suffix = String(employeeId).padStart(2, "0");
  const prefix = normalizeAccountUsername(name).slice(0, 100 - suffix.length);
  if (!prefix) throw httpError(400, "The employee name cannot produce a username.");
  return `${prefix}${suffix}`;
}

async function getEmployeeIdForUser(user) {
  const [employees] = await pool.execute(
    "SELECT EmployeeID AS employeeId FROM tbl_employees WHERE Username = ? LIMIT 1",
    [user.username],
  );
  if (!employees.length) {
    throw httpError(403, "Your employee account is not linked to an employee record.");
  }
  return Number(employees[0].employeeId);
}

async function sendExcelWorkbook(response, filename, worksheetName, columns, rows) {
  const buffer = await createExcelWorkbook(worksheetName, columns, rows);
  response
    .status(200)
    .set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    .set("Content-Disposition", `attachment; filename="${filename}"`)
    .set("Content-Length", String(buffer.byteLength))
    .send(Buffer.from(buffer));
}

async function createExcelWorkbook(worksheetName, columns, rows) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "EMS";
  workbook.created = new Date();
  const worksheet = workbook.addWorksheet(worksheetName);
  worksheet.columns = columns;
  worksheet.addRows(rows);
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  worksheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: columns.length },
  };
  worksheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  worksheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF173B63" },
  };
  return workbook.xlsx.writeBuffer();
}

function setSessionCookie(response, sessionId) {
  response.cookie("ems_session", sessionId, {
    httpOnly: true,
    sameSite: "lax",
    maxAge: sessionLifetimeMs,
    path: "/",
  });
}

function clearSessionCookie(response) {
  response.clearCookie("ems_session", { httpOnly: true, sameSite: "lax", path: "/" });
}

function readText(value, label, maximum) {
  const text = typeof value === "string" ? value.trim() : "";
  if (!text) throw httpError(400, `${label} is required.`);
  if (text.length > maximum) throw httpError(400, `${label} must be ${maximum} characters or fewer.`);
  return text;
}

async function hashPassword(password) {
  const salt = crypto.randomBytes(16);
  const hash = await scrypt(password, salt, 64);
  return `scrypt$${salt.toString("base64url")}$${hash.toString("base64url")}`;
}

async function verifyPassword(password, passwordHash) {
  const passwordData = parsePasswordHash(passwordHash);
  if (!passwordData) return false;
  const actualHash = await scrypt(password, passwordData.salt, passwordData.expectedHash.length);
  return crypto.timingSafeEqual(passwordData.expectedHash, actualHash);
}

function readId(value, label) {
  const id = Number(value);
  if (!Number.isSafeInteger(id) || id < 1) throw httpError(400, `Invalid ${label}.`);
  return id;
}

function readDate(value) {
  const date = readText(value, "Date", 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw httpError(400, "Date must be a valid calendar date.");
  const parsed = new Date(`${date}T00:00:00Z`);
  if (Number.isNaN(parsed.getTime()) || parsed.toISOString().slice(0, 10) !== date) {
    throw httpError(400, "Date must be a valid calendar date.");
  }
  return date;
}

function readMonth(value) {
  const month = readText(value, "Month", 7);
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw httpError(400, "Month must use YYYY-MM format.");
  return month;
}

function readTime(value, label) {
  const time = readText(value, label, 8);
  if (!/^(?:[01]\d|2[0-3]):[0-5]\d(?::[0-5]\d)?$/.test(time)) {
    throw httpError(400, `${label} must be a valid time.`);
  }
  return time.length === 5 ? `${time}:00` : time;
}

function timeToMinutes(time) {
  const [hours, minutes] = String(time).split(":").map(Number);
  return hours * 60 + minutes;
}

// Attendance time calculations and absence checks used by attendance and payroll.
function isAbsentAttendance(record) {
  return timeToMinutes(record.timeIn) === 0 && timeToMinutes(record.timeOut) === 0;
}

function emptyWorkday() {
  return {
    totalWorkingHours: 0,
    totalWorkingMinutes: 0,
    overtimeMinutes: 0,
    overtimeHours: 0,
    lateMinutes: 0,
    lateHours: 0,
  };
}

function calculateWorkday(date, timeIn, timeOut) {
  const inMinutes = timeToMinutes(timeIn);
  const outMinutes = timeToMinutes(timeOut);
  if (outMinutes <= inMinutes) throw httpError(400, "Time out must be later than time in.");

  const regularMinutes =
    Math.max(0, Math.min(outMinutes, 12 * 60) - Math.max(inMinutes, 8 * 60)) +
    Math.max(0, Math.min(outMinutes, 17 * 60) - Math.max(inMinutes, 13 * 60));
  const overtimeMinutes = Math.max(0, outMinutes - Math.max(inMinutes, 17 * 60));
  const weekday = new Date(`${date}T00:00:00Z`).getUTCDay();
  let lateMinutes = 0;
  if (weekday !== 0 && weekday !== 6) {
    if (inMinutes > 8 * 60) {
      lateMinutes = inMinutes <= 12 * 60
        ? inMinutes - 8 * 60
        : inMinutes <= 13 * 60
          ? 4 * 60
          : 4 * 60 + inMinutes - 13 * 60;
    }
  }

  return {
    totalWorkingHours: Number(((regularMinutes + overtimeMinutes) / 60).toFixed(2)),
    totalWorkingMinutes: regularMinutes + overtimeMinutes,
    overtimeMinutes,
    overtimeHours: overtimeMinutes / 60,
    lateMinutes,
    lateHours: lateMinutes / 60,
  };
}

function attendanceWithCalculatedHours(record) {
  const workday = isAbsentAttendance(record)
    ? emptyWorkday()
    : calculateWorkday(record.date, record.timeIn, record.timeOut);
  return {
    ...record,
    overtime: workday.overtimeHours,
    totalWorkingHours: workday.totalWorkingHours,
    isAbsent: isAbsentAttendance(record),
  };
}

function includeMissingEmployeesAsAbsent(records, employees, date) {
  const attendance = records.map(attendanceWithCalculatedHours);
  if (!canMarkAbsentForDate(date)) return attendance;
  const recordedEmployeeIds = new Set(attendance.map((record) => Number(record.employeeId)));
  for (const employee of employees) {
    if (recordedEmployeeIds.has(Number(employee.employeeId))) continue;
    attendance.push({
      attendanceId: null,
      employeeId: employee.employeeId,
      date,
      timeIn: "00:00:00",
      timeOut: "00:00:00",
      ...emptyWorkday(),
      overtime: 0,
      isAbsent: true,
    });
  }
  return attendance;
}

function readEmployee(body) {
  const employee = {
    name: readText(body.name ?? body.Name, "Name", 100),
    gender: readText(body.gender ?? body.Gender, "Gender", 30),
    birthDate: readDate(body.birthDate ?? body.BirthDate),
    contactNumber: readText(body.contactNumber ?? body.ContactNumber, "Contact number", 25),
    email: readText(body.email ?? body.Email, "Email", 255),
    position: readText(body.position ?? body.Position, "Position", 100),
    address: readText(body.address ?? body.Address, "Address", 255),
  };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(employee.email)) {
    throw httpError(400, "Enter a valid email address.");
  }
  return employee;
}

// Manila-based date helpers and monthly payroll calculations persisted to the database.
function currentMonth() {
  return currentDate().slice(0, 7);
}

function currentDate() {
  const now = new Date();
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: "Asia/Manila",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(now);
  const values = Object.fromEntries(parts.map(({ type, value }) => [type, value]));
  return `${values.year}-${values.month}-${values.day}`;
}

function canMarkAbsentForDate(date) {
  const weekday = new Date(`${date}T00:00:00Z`).getUTCDay();
  return date < currentDate() && weekday !== 0 && weekday !== 6;
}

function addDays(date, amount) {
  const result = new Date(`${date}T00:00:00Z`);
  result.setUTCDate(result.getUTCDate() + amount);
  return result.toISOString().slice(0, 10);
}

async function saveAbsentAttendanceForDate(date) {
  if (!canMarkAbsentForDate(date)) {
    return 0;
  }

  const connection = await pool.getConnection();
  try {
    await connection.beginTransaction();
    const [employees] = await connection.query(
      "SELECT EmployeeID FROM tbl_employees ORDER BY EmployeeID FOR UPDATE",
    );
    const [existingRecords] = await connection.execute(
      "SELECT EmployeeID FROM tbl_attendance WHERE Date = ? FOR UPDATE",
      [date],
    );
    const recordedEmployeeIds = new Set(existingRecords.map((record) => Number(record.EmployeeID)));
    let created = 0;
    for (const employee of employees) {
      if (recordedEmployeeIds.has(Number(employee.EmployeeID))) continue;
      await connection.execute(
        `INSERT INTO tbl_attendance
          (EmployeeID, Date, \`Time in\`, \`Time out\`, Overtime, \`Total Working Hours\`)
         VALUES (?, ?, '00:00:00', '00:00:00', 0, 0)`,
        [employee.EmployeeID, date],
      );
      created += 1;
    }
    if (created > 0) {
      for (const employee of employees) {
        await updateEmployeeSalary(connection, employee.EmployeeID, date.slice(0, 7));
      }
    }
    await connection.commit();
    return created;
  } catch (error) {
    await connection.rollback();
    throw error;
  } finally {
    connection.release();
  }
}

let lastAbsenceProcessedDate = addDays(currentDate(), -2);
let absenceSchedulerRunning = false;

async function processCompletedWeekdays() {
  if (absenceSchedulerRunning) return;
  absenceSchedulerRunning = true;
  try {
    const lastCompletedDate = addDays(currentDate(), -1);
    while (lastAbsenceProcessedDate < lastCompletedDate) {
      const date = addDays(lastAbsenceProcessedDate, 1);
      const created = await saveAbsentAttendanceForDate(date);
      lastAbsenceProcessedDate = date;
      if (created > 0) {
        console.log(`Saved ${created} absent attendance record(s) for ${date}.`);
      }
    }
  } catch (error) {
    console.error("Automatic absence recording failed:", error.code || error.message);
  } finally {
    absenceSchedulerRunning = false;
  }
}

function absenceCutoff(month) {
  const today = currentDate();
  if (month > today.slice(0, 7)) return null;
  if (month === today.slice(0, 7)) return addDays(today, -1);
  const [year, monthNumber] = month.split("-").map(Number);
  return new Date(Date.UTC(year, monthNumber, 0)).toISOString().slice(0, 10);
}

function previousMonth(month) {
  const [year, monthNumber] = month.split("-").map(Number);
  return new Date(Date.UTC(year, monthNumber - 2, 1)).toISOString().slice(0, 7);
}

const attendanceExportColumns = [
  { header: "AttendanceID", key: "attendanceId", width: 16 },
  { header: "EmployeeID", key: "employeeId", width: 14 },
  { header: "Employee Name", key: "employeeName", width: 28 },
  { header: "Date", key: "date", width: 14 },
  { header: "Status", key: "status", width: 14 },
  { header: "Time In", key: "timeIn", width: 14 },
  { header: "Time Out", key: "timeOut", width: 14 },
  { header: "Overtime", key: "overtime", width: 14 },
  { header: "Total Working Hours", key: "totalWorkingHours", width: 22 },
];

async function monthlyAttendanceExportRows(month) {
  const startDate = `${month}-01`;
  const endDate = `${addDays(startDate, 32).slice(0, 7)}-01`;
  const [rawRecords] = await pool.execute(
    `SELECT a.AttendanceID AS attendanceId, a.EmployeeID AS employeeId, e.Name AS employeeName,
            a.Date AS date, a.\`Time in\` AS timeIn, a.\`Time out\` AS timeOut,
            a.Overtime AS overtime, a.\`Total Working Hours\` AS totalWorkingHours
     FROM tbl_attendance a
     JOIN tbl_employees e ON e.EmployeeID = a.EmployeeID
     WHERE a.Date >= ? AND a.Date < ?
     ORDER BY a.Date, a.EmployeeID, a.\`Time in\``,
    [startDate, endDate],
  );
  const records = rawRecords.map(({ employeeName, ...record }) => ({
    ...attendanceWithCalculatedHours(record),
    employeeName,
  }));
  const [employees] = await pool.query(
    "SELECT EmployeeID AS employeeId, Name AS employeeName FROM tbl_employees",
  );
  const employeeNames = new Map(
    employees.map((employee) => [Number(employee.employeeId), employee.employeeName]),
  );
  const recordsByDate = new Map();
  for (const record of records) {
    const key = String(record.date).slice(0, 10);
    if (!recordsByDate.has(key)) recordsByDate.set(key, []);
    recordsByDate.get(key).push(record);
  }
  const exportRecords = records.slice();
  const cutoff = absenceCutoff(month);
  if (cutoff) {
    const cursor = new Date(`${month}-01T00:00:00Z`);
    const end = new Date(`${cutoff}T00:00:00Z`);
    while (cursor <= end) {
      const weekday = cursor.getUTCDay();
      const date = cursor.toISOString().slice(0, 10);
      if (weekday !== 0 && weekday !== 6) {
        const dayRecords = includeMissingEmployeesAsAbsent(
          recordsByDate.get(date) || [],
          employees,
          date,
        );
        for (const record of dayRecords) {
          if (record.attendanceId !== null) continue;
          exportRecords.push({
            ...record,
            employeeName: employeeNames.get(Number(record.employeeId)),
          });
        }
      }
      cursor.setUTCDate(cursor.getUTCDate() + 1);
    }
  }
  return exportRecords
    .sort((left, right) =>
      String(left.date).localeCompare(String(right.date))
      || Number(left.employeeId) - Number(right.employeeId),
    )
    .map((record) => ({
      attendanceId: record.attendanceId,
      employeeId: record.employeeId,
      employeeName: record.employeeName,
      date: record.date,
      status: record.isAbsent ? "Absent" : "Present",
      timeIn: record.timeIn,
      timeOut: record.timeOut,
      overtime: Number(record.overtime),
      totalWorkingHours: Number(record.totalWorkingHours),
    }));
}

function countAbsentWeekdays(month, cutoff, records) {
  if (!cutoff) return 0;
  const attendedDates = new Set(
    records
      .filter((record) => !isAbsentAttendance(record))
      .map((record) => String(record.date).slice(0, 10))
      .filter((date) => date <= cutoff),
  );
  const cursor = new Date(`${month}-01T00:00:00Z`);
  const end = new Date(`${cutoff}T00:00:00Z`);
  let absentDays = 0;

  while (cursor <= end) {
    const weekday = cursor.getUTCDay();
    const date = cursor.toISOString().slice(0, 10);
    if (weekday !== 0 && weekday !== 6 && !attendedDates.has(date)) absentDays += 1;
    cursor.setUTCDate(cursor.getUTCDate() + 1);
  }
  return absentDays;
}

async function updateEmployeeSalary(connection, employeeId, month) {
  const startDate = `${month}-01`;
  const [records] = await connection.execute(
    `SELECT Date AS date, \`Time in\` AS timeIn, \`Time out\` AS timeOut
     FROM tbl_attendance
     WHERE EmployeeID = ? AND Date >= ? AND Date < DATE_ADD(?, INTERVAL 1 MONTH)`,
    [employeeId, startDate, startDate],
  );

  let overtimeMinutes = 0;
  let lateMinutes = 0;
  let totalWorkingMinutes = 0;
  for (const record of records) {
    if (isAbsentAttendance(record)) continue;
    const workday = calculateWorkday(record.date, record.timeIn, record.timeOut);
    overtimeMinutes += workday.overtimeMinutes;
    lateMinutes += workday.lateMinutes;
    totalWorkingMinutes += workday.totalWorkingMinutes;
  }

  const totalWorkingHours = Number((totalWorkingMinutes / 60).toFixed(2));
  const overtimeHours = overtimeMinutes / 60;
  const lateHours = lateMinutes / 60;
  const absentDays = countAbsentWeekdays(month, absenceCutoff(month), records);
  const absentHours = absentDays * 8;
  const overtimePay = Number((overtimeMinutes * overtimeRate / 60).toFixed(2));
  const lateDeduction = Number((lateMinutes * lateDeductionRate / 60).toFixed(2));
  const absenceDeduction = Number((absentHours * lateDeductionRate).toFixed(2));
  const deductions = Number((lateDeduction + absenceDeduction).toFixed(2));
  const netSalary = Number((basicSalary + overtimePay - deductions).toFixed(2));
  const [existing] = await connection.execute(
    "SELECT SalaryID FROM tbl_salary WHERE EmployeeID = ? ORDER BY SalaryID FOR UPDATE",
    [employeeId],
  );

  if (existing.length) {
    await connection.execute(
      `UPDATE tbl_salary
       SET \`Basic Salary\` = ?, \`Working Hours\` = ?, \`Overtime Hours\` = ?,
           \`Overtime Pay\` = ?, \`Absent Days\` = ?, \`Absence Deduction\` = ?,
           \`Late Deduction\` = ?, Deductions = ?, \`Net Salary\` = ?
       WHERE EmployeeID = ?`,
      [basicSalary, totalWorkingHours, overtimeHours, overtimePay, absentDays, absenceDeduction, lateDeduction, deductions, netSalary, employeeId],
    );
  } else {
    await connection.execute(
      `INSERT INTO tbl_salary
        (EmployeeID, \`Basic Salary\`, \`Working Hours\`, \`Overtime Hours\`,
         \`Overtime Pay\`, \`Absent Days\`, \`Absence Deduction\`, \`Late Deduction\`,
         Deductions, \`Net Salary\`)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [employeeId, basicSalary, totalWorkingHours, overtimeHours, overtimePay, absentDays, absenceDeduction, lateDeduction, deductions, netSalary],
    );
  }

  return {
    totalWorkingHours,
    overtimeHours,
    lateHours,
    absentDays,
    absentHours,
    overtimePay,
    lateDeduction,
    absenceDeduction,
    deductions,
    netSalary,
  };
}

async function refreshMonthlySalaries(connection, month) {
  const [employees] = await connection.query("SELECT EmployeeID FROM tbl_employees");
  const calculations = new Map();
  for (const employee of employees) {
    calculations.set(
      Number(employee.EmployeeID),
      await updateEmployeeSalary(connection, employee.EmployeeID, month),
    );
  }
  return calculations;
}

function asyncRoute(handler) {
  return (request, response, next) => Promise.resolve(handler(request, response, next)).catch(next);
}

// Express request parsing, local frontend CORS access, and the default dashboard route.
app.disable("x-powered-by");
app.use(express.json({ limit: "32kb" }));

const liveServerOrigins = new Set([
  "http://localhost:5500",
  "http://127.0.0.1:5500",
]);

app.use((request, response, next) => {
  const origin = request.get("Origin");
  if (liveServerOrigins.has(origin)) {
    response.set("Access-Control-Allow-Origin", origin);
    response.set("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
    response.set("Access-Control-Allow-Headers", "Content-Type");
    response.set("Access-Control-Expose-Headers", "Content-Disposition");
    response.set("Access-Control-Allow-Credentials", "true");
    response.vary("Origin");
  }
  if (request.method === "OPTIONS") {
    return response.sendStatus(liveServerOrigins.has(origin) ? 204 : 403);
  }
  next();
});

app.get("/", (request, response) => response.redirect("/dashboard.html"));

// Authenticate users and create the HTTP-only session cookie.
app.post("/api/login", asyncRoute(async (request, response) => {
  const body = request.body || {};
  const username = typeof body.username === "string" ? body.username.trim() : "";
  const password = typeof body.password === "string" ? body.password : "";
  if (!username || !password || username.length > 100 || password.length > 255) {
    throw httpError(401, "Invalid username or password.");
  }

  const [users] = await pool.execute(
    `SELECT UserID AS userId, Username AS username, \`Password Hash\` AS passwordHash,
            \`Account Role\` AS accountRole
     FROM tbl_users
     WHERE Username = ?
     LIMIT 1`,
    [username],
  );
  const user = users[0];
  if (!await verifyPassword(password, user?.passwordHash)) {
    throw httpError(401, "Invalid username or password.");
  }

  const role = normalizeRole(user.accountRole);
  if (!role) throw httpError(403, "This account does not have an application role.");

  const sessionId = crypto.randomBytes(32).toString("hex");
  sessions.set(sessionId, {
    user: { userId: user.userId, username: user.username, accountRole: user.accountRole, role },
    expiresAt: Date.now() + sessionLifetimeMs,
  });
  setSessionCookie(response, sessionId);
  response.json({ username: user.username, accountRole: user.accountRole, role });
}));

app.use("/api", authenticate);

// Return the signed-in identity and its linked employee ID, if applicable.
app.get("/api/session", asyncRoute(async (request, response) => {
  const employeeId = request.user.role === "employee"
    ? await getEmployeeIdForUser(request.user)
    : null;
  response.json({ ...request.user, employeeId });
}));

// Read or update the signed-in employee's own username and password.
app.get("/api/account-settings", requireRoles("employee"), asyncRoute(async (request, response) => {
  await getEmployeeIdForUser(request.user);
  const [users] = await pool.execute(
    "SELECT Username AS username, `Account Role` AS accountRole FROM tbl_users WHERE UserID = ?",
    [request.user.userId],
  );
  if (!users.length) throw httpError(404, "Account not found.");
  response.json(users[0]);
}));

app.put("/api/account-settings", requireRoles("employee"), asyncRoute(async (request, response) => {
  const body = request.body || {};
  const username = readText(body.username, "Username", 100);
  const currentPassword = typeof body.currentPassword === "string" ? body.currentPassword : "";
  const newPassword = typeof body.newPassword === "string" ? body.newPassword : "";
  const confirmPassword = typeof body.confirmPassword === "string" ? body.confirmPassword : "";
  if (!currentPassword || currentPassword.length > 255) {
    throw httpError(400, "Enter your current password.");
  }
  if (newPassword && (newPassword.length < 8 || newPassword.length > 255)) {
    throw httpError(400, "New password must be between 8 and 255 characters.");
  }
  if (newPassword !== confirmPassword) {
    throw httpError(400, "New password and confirmation do not match.");
  }

  const connection = await pool.getConnection();
  try {
    await connection.beginTransaction();
    const [users] = await connection.execute(
      `SELECT u.UserID AS userId, u.Username AS username,
              u.\`Password Hash\` AS passwordHash, u.\`Account Role\` AS accountRole
       FROM tbl_users u
       JOIN tbl_employees e ON e.Username = u.Username
       WHERE u.UserID = ?
       FOR UPDATE`,
      [request.user.userId],
    );
    const user = users[0];
    if (!user || normalizeRole(user.accountRole) !== "employee") {
      throw httpError(403, "Your employee account is not linked to an employee record.");
    }
    if (!await verifyPassword(currentPassword, user.passwordHash)) {
      throw httpError(401, "Current password is incorrect.");
    }
    if (username === user.username && !newPassword) {
      throw httpError(400, "There are no account changes to save.");
    }

    const passwordHash = newPassword ? await hashPassword(newPassword) : user.passwordHash;
    await connection.execute(
      "UPDATE tbl_users SET Username = ?, `Password Hash` = ? WHERE UserID = ?",
      [username, passwordHash, user.userId],
    );
    await connection.commit();
    const currentSessionId = getCookie(request, "ems_session");
    for (const [sessionId, session] of sessions) {
      if (Number(session.user.userId) !== Number(user.userId)) continue;
      if (newPassword && sessionId !== currentSessionId) {
        sessions.delete(sessionId);
      } else {
        session.user.username = username;
      }
    }
    request.user.username = username;
    response.json({ username, accountRole: user.accountRole });
  } catch (error) {
    await connection.rollback();
    if (error.code === "ER_DUP_ENTRY") {
      throw httpError(409, "That username is already in use.");
    }
    throw error;
  } finally {
    connection.release();
  }
}));

// Serve dashboard counts and attendance scoped to Employee accounts.
app.get("/api/dashboard", asyncRoute(async (request, response) => {
  const date = request.query.date ? readDate(request.query.date) : currentDate();
  const employeeId = request.user.role === "employee"
    ? await getEmployeeIdForUser(request.user)
    : null;
  const [[employeeCount]] = await pool.execute(
    "SELECT COUNT(*) AS total FROM tbl_employees WHERE (? IS NULL OR EmployeeID = ?)",
    [employeeId, employeeId],
  );
  const [[presentCount]] = await pool.execute(
    `SELECT COUNT(DISTINCT EmployeeID) AS total
     FROM tbl_attendance
     WHERE Date = ? AND (? IS NULL OR EmployeeID = ?)
       AND NOT (
         COALESCE(TIME_TO_SEC(\`Time in\`), 0) = 0
         AND COALESCE(TIME_TO_SEC(\`Time out\`), 0) = 0
       )`,
    [date, employeeId, employeeId],
  );
  const [attendance] = await pool.execute(
    `SELECT AttendanceID AS attendanceId, EmployeeID AS employeeId, Date AS date,
            \`Time in\` AS timeIn, \`Time out\` AS timeOut, Overtime AS overtime,
            \`Total Working Hours\` AS totalWorkingHours
     FROM tbl_attendance
     WHERE Date = ? AND (? IS NULL OR EmployeeID = ?)
     ORDER BY \`Time in\``,
    [date, employeeId, employeeId],
  );
  const [employees] = await pool.execute(
    "SELECT EmployeeID AS employeeId FROM tbl_employees WHERE (? IS NULL OR EmployeeID = ?)",
    [employeeId, employeeId],
  );
  const weekday = new Date(`${date}T00:00:00Z`).getUTCDay();
  const completedWeekday = date < currentDate() && weekday !== 0 && weekday !== 6;
  response.json({
    totalEmployees: employeeCount.total,
    presentToday: presentCount.total,
    absentToday: completedWeekday
      ? Math.max(0, Number(employeeCount.total) - Number(presentCount.total))
      : 0,
    attendance: includeMissingEmployeesAsAbsent(attendance, employees, date),
  });
}));

// List searchable employee details; manager-only routes export or change records.
app.get("/api/employees", asyncRoute(async (request, response) => {
  const search = typeof request.query.search === "string" ? request.query.search.trim() : "";
  const pattern = `%${search.replace(/[!%_]/g, "!$&")}%`;
  const [employees] = await pool.execute(
    `SELECT EmployeeID AS employeeId, Name AS name, Gender AS gender,
            BirthDate AS birthDate, ContactNumber AS contactNumber,
            Email AS email, Position AS position, Address AS address
     FROM tbl_employees
     WHERE ? = '' OR Name LIKE ? ESCAPE '!' OR CAST(EmployeeID AS CHAR) LIKE ? OR Email LIKE ? ESCAPE '!'
     ORDER BY EmployeeID DESC`,
    [search, pattern, pattern, pattern],
  );
  response.json(employees);
}));

app.get("/api/exports/employees.xlsx", requireRoles("admin", "hr"), asyncRoute(async (request, response) => {
  const [employees] = await pool.query(
    `SELECT EmployeeID AS employeeId, Name AS name, Gender AS gender,
            BirthDate AS birthDate, ContactNumber AS contactNumber,
            Email AS email, Position AS position, Address AS address
     FROM tbl_employees
     ORDER BY EmployeeID`,
  );
  const rows = employees.map((employee) => ({
    employeeId: employee.employeeId,
    name: employee.name,
    gender: employee.gender,
    birthDate: employee.birthDate,
    contactNumber: employee.contactNumber,
    email: employee.email,
    position: employee.position,
    address: employee.address,
  }));
  await sendExcelWorkbook(
    response,
    `Employee_List_${currentDate()}.xlsx`,
    "Employee Directory",
    [
      { header: "EmployeeID", key: "employeeId", width: 14 },
      { header: "Name", key: "name", width: 28 },
      { header: "Gender", key: "gender", width: 18 },
      { header: "Birth Date", key: "birthDate", width: 14 },
      { header: "Contact Number", key: "contactNumber", width: 20 },
      { header: "Email", key: "email", width: 32 },
      { header: "Position", key: "position", width: 24 },
      { header: "Address", key: "address", width: 40 },
    ],
    rows,
  );
}));

// Create, update, and delete employee records and their linked data.
app.post("/api/employees", requireRoles("admin", "hr"), asyncRoute(async (request, response) => {
  const employee = readEmployee(request.body);
  const connection = await pool.getConnection();
  try {
    await connection.beginTransaction();
    const [existingEmails] = await connection.execute(
      "SELECT EmployeeID FROM tbl_employees WHERE LOWER(Email) = LOWER(?) LIMIT 1 FOR UPDATE",
      [employee.email],
    );
    if (existingEmails.length) throw httpError(409, "An employee with this email address already exists.");
    const [result] = await connection.execute(
      `INSERT INTO tbl_employees
        (Name, Gender, BirthDate, ContactNumber, Email, Position, Address)
       VALUES (?, ?, ?, ?, ?, ?, ?)`,
      [
        employee.name,
        employee.gender,
        employee.birthDate,
        employee.contactNumber,
        employee.email,
        employee.position,
        employee.address,
      ],
    );
    await connection.execute(
      "INSERT INTO tbl_salary (EmployeeID, `Basic Salary`, `Overtime Pay`, Deductions, `Net Salary`) VALUES (?, ?, 0, 0, ?)",
      [result.insertId, basicSalary, basicSalary],
    );
    await connection.commit();
    response.status(201).json({ employeeId: result.insertId });
  } catch (error) {
    await connection.rollback();
    throw error;
  } finally {
    connection.release();
  }
}));

app.put("/api/employees/:employeeId", requireRoles("admin", "hr"), asyncRoute(async (request, response) => {
  const employeeId = readId(request.params.employeeId, "employee ID");
  const employee = readEmployee(request.body);
  const connection = await pool.getConnection();
  try {
    await connection.beginTransaction();
    const [employees] = await connection.execute(
      "SELECT EmployeeID FROM tbl_employees WHERE EmployeeID = ? FOR UPDATE",
      [employeeId],
    );
    if (!employees.length) throw httpError(404, "Employee not found.");
    const [existingEmails] = await connection.execute(
      "SELECT EmployeeID FROM tbl_employees WHERE LOWER(Email) = LOWER(?) AND EmployeeID <> ? LIMIT 1 FOR UPDATE",
      [employee.email, employeeId],
    );
    if (existingEmails.length) throw httpError(409, "Another employee is already using this email address.");
    await connection.execute(
      `UPDATE tbl_employees
       SET Name = ?, Gender = ?, BirthDate = ?, ContactNumber = ?, Email = ?, Position = ?, Address = ?
       WHERE EmployeeID = ?`,
      [
        employee.name,
        employee.gender,
        employee.birthDate,
        employee.contactNumber,
        employee.email,
        employee.position,
        employee.address,
        employeeId,
      ],
    );
    await connection.commit();
    response.json({ employeeId });
  } catch (error) {
    await connection.rollback();
    throw error;
  } finally {
    connection.release();
  }
}));

app.delete("/api/employees/:employeeId", requireRoles("admin", "hr"), asyncRoute(async (request, response) => {
  const employeeId = readId(request.params.employeeId, "employee ID");
  const connection = await pool.getConnection();
  try {
    await connection.beginTransaction();
    const [employees] = await connection.execute(
      "SELECT EmployeeID, Username AS username FROM tbl_employees WHERE EmployeeID = ? FOR UPDATE",
      [employeeId],
    );
    if (!employees.length) throw httpError(404, "Employee not found.");
    await connection.execute("DELETE FROM tbl_attendance WHERE EmployeeID = ?", [employeeId]);
    await connection.execute("DELETE FROM tbl_salary WHERE EmployeeID = ?", [employeeId]);
    await connection.execute("DELETE FROM tbl_employees WHERE EmployeeID = ?", [employeeId]);
    if (employees[0].username) {
      await connection.execute(
        "DELETE FROM tbl_users WHERE Username = ? AND LOWER(TRIM(`Account Role`)) = 'employee'",
        [employees[0].username],
      );
    }
    await connection.commit();
    response.status(204).end();
  } catch (error) {
    await connection.rollback();
    throw error;
  } finally {
    connection.release();
  }
}));

// List scoped attendance, mark completed weekdays absent, and manage exports.
app.post("/api/attendance/absences", requireRoles("admin", "hr"), asyncRoute(async (request, response) => {
  const date = readDate(request.body.date);
  if (!canMarkAbsentForDate(date)) {
    throw httpError(400, "Absent attendance can only be saved for weekdays up to today.");
  }
  const created = await saveAbsentAttendanceForDate(date);
  response.status(created > 0 ? 201 : 200).json({ date, created });
}));

app.get("/api/attendance", asyncRoute(async (request, response) => {
  const date = request.query.date ? readDate(request.query.date) : null;
  const employeeId = request.user.role === "employee"
    ? await getEmployeeIdForUser(request.user)
    : null;
  const [records] = await pool.execute(
    `SELECT AttendanceID AS attendanceId, EmployeeID AS employeeId, Date AS date,
            \`Time in\` AS timeIn, \`Time out\` AS timeOut, Overtime AS overtime,
            \`Total Working Hours\` AS totalWorkingHours
     FROM tbl_attendance
     WHERE (? IS NULL OR Date = ?) AND (? IS NULL OR EmployeeID = ?)
     ORDER BY Date DESC, \`Time in\` DESC`,
    [date, date, employeeId, employeeId],
  );
  if (!date) {
    response.json(records.map(attendanceWithCalculatedHours));
    return;
  }
  const [employees] = await pool.execute(
    "SELECT EmployeeID AS employeeId FROM tbl_employees WHERE (? IS NULL OR EmployeeID = ?)",
    [employeeId, employeeId],
  );
  response.json(includeMissingEmployeesAsAbsent(records, employees, date));
}));

app.get("/api/exports/attendance/monthly", requireRoles("admin", "hr"), asyncRoute(async (request, response) => {
  const filenames = (await fs.readdir(exportsDirectory))
    .filter((filename) => /^Monthly_Attendance_\d{4}-(0[1-9]|1[0-2])\.xlsx$/.test(filename))
    .sort()
    .reverse();
  response.json(filenames.map((filename) => ({
    month: filename.slice("Monthly_Attendance_".length, -".xlsx".length),
    filename,
  })));
}));

app.get("/api/settings/automatic-monthly-export", requireRoles("admin", "hr"), asyncRoute(async (request, response) => {
  const [settings] = await pool.execute(
    `SELECT SettingValue AS settingValue FROM ${automaticMonthlyExportSettingsTable} WHERE SettingKey = ?`,
    [automaticMonthlyExportSetting],
  );
  if (!settings.length) throw httpError(500, "Automatic export setting is missing.");
  response.json({ enabled: settings[0].settingValue === "true" });
}));

app.put("/api/settings/automatic-monthly-export", requireRoles("admin", "hr"), asyncRoute(async (request, response) => {
  const enabled = request.body?.enabled;
  if (typeof enabled !== "boolean") {
    throw httpError(400, "Automatic monthly export must be set to Yes or No.");
  }
  await pool.execute(
    `INSERT INTO ${automaticMonthlyExportSettingsTable} (SettingKey, SettingValue)
     VALUES (?, ?)
     ON DUPLICATE KEY UPDATE SettingValue = VALUES(SettingValue)`,
    [automaticMonthlyExportSetting, enabled ? "true" : "false"],
  );
  response.json({ enabled });
}));

app.get("/api/exports/attendance/monthly/:month.xlsx", requireRoles("admin", "hr"), asyncRoute(async (request, response) => {
  const month = readMonth(request.params.month);
  const filename = `Monthly_Attendance_${month}.xlsx`;
  const filePath = path.join(exportsDirectory, filename);
  try {
    await fs.access(filePath);
  } catch (error) {
    if (error.code === "ENOENT") throw httpError(404, "That monthly attendance export was not found.");
    throw error;
  }
  response
    .status(200)
    .set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    .set("Content-Disposition", `attachment; filename="${filename}"`)
    .send(await fs.readFile(filePath));
}));

app.get("/api/exports/attendance.xlsx", requireRoles("admin", "hr"), asyncRoute(async (request, response) => {
  const month = readMonth(request.query.month);
  const exportRows = await monthlyAttendanceExportRows(month);
  await sendExcelWorkbook(
    response,
    `Monthly_Attendance_${month}.xlsx`,
    "Attendance Records",
    attendanceExportColumns,
    exportRows,
  );
}));

// Validate attendance changes and recalculate payroll within database transactions.
async function saveAttendance(request, response, update) {
  const employeeId = readId(request.body.employeeId ?? request.body.EmployeeID, "employee ID");
  const date = readDate(request.body.date ?? request.body.Date);
  const timeIn = readTime(request.body.timeIn ?? request.body.TimeIn, "Time in");
  const timeOut = readTime(request.body.timeOut ?? request.body.TimeOut, "Time out");
  const requestedAttendanceId = update
    ? readId(request.params.attendanceId, "attendance ID")
    : null;
  const workday = isAbsentAttendance({ timeIn, timeOut })
    ? emptyWorkday()
    : calculateWorkday(date, timeIn, timeOut);
  const connection = await pool.getConnection();
  try {
    await connection.beginTransaction();
    const [employees] = await connection.execute(
      "SELECT EmployeeID FROM tbl_employees WHERE EmployeeID = ? FOR UPDATE",
      [employeeId],
    );
    if (!employees.length) throw httpError(400, "Select an existing employee.");

    let attendanceId;
    let previousEmployeeId = null;
    let action = "added";
    if (update) {
      attendanceId = requestedAttendanceId;
      const [existing] = await connection.execute(
        "SELECT EmployeeID FROM tbl_attendance WHERE AttendanceID = ? FOR UPDATE",
        [attendanceId],
      );
      if (!existing.length) throw httpError(404, "Attendance record not found.");
      const [duplicateRecords] = await connection.execute(
        "SELECT AttendanceID FROM tbl_attendance WHERE EmployeeID = ? AND Date = ? AND AttendanceID <> ?",
        [employeeId, date, requestedAttendanceId],
      );
      if (duplicateRecords.length) {
        throw httpError(409, "This employee already has another attendance record for that date.");
      }
      previousEmployeeId = existing[0].EmployeeID;
      await connection.execute(
        `UPDATE tbl_attendance
         SET EmployeeID = ?, Date = ?, \`Time in\` = ?, \`Time out\` = ?,
             Overtime = ?, \`Total Working Hours\` = ?
         WHERE AttendanceID = ?`,
        [employeeId, date, timeIn, timeOut, workday.overtimeHours, workday.totalWorkingHours, attendanceId],
      );
    } else {
        const [existing] = await connection.execute(
          "SELECT AttendanceID FROM tbl_attendance WHERE EmployeeID = ? AND Date = ? FOR UPDATE",
          [employeeId, date],
      );
      if (existing.length) {
          attendanceId = existing[0].AttendanceID;
          action = "updated";
          await connection.execute(
            `UPDATE tbl_attendance
             SET \`Time in\` = ?, \`Time out\` = ?, Overtime = ?, \`Total Working Hours\` = ?
             WHERE AttendanceID = ?`,
            [timeIn, timeOut, workday.overtimeHours, workday.totalWorkingHours, attendanceId],
          );
      } else {
          const [result] = await connection.execute(
            `INSERT INTO tbl_attendance
              (EmployeeID, Date, \`Time in\`, \`Time out\`, Overtime, \`Total Working Hours\`)
             VALUES (?, ?, ?, ?, ?, ?)`,
            [employeeId, date, timeIn, timeOut, workday.overtimeHours, workday.totalWorkingHours],
          );
          attendanceId = result.insertId;
      }
    }

    const payrollMonth = currentMonth();
    await updateEmployeeSalary(connection, employeeId, payrollMonth);
    if (previousEmployeeId && Number(previousEmployeeId) !== employeeId) {
      await updateEmployeeSalary(connection, previousEmployeeId, payrollMonth);
    }
    await connection.commit();
    response.status(update || action === "updated" ? 200 : 201).json({
      attendanceId,
      action: update ? "updated" : action,
      ...workday,
      lateDeduction: Number((workday.lateMinutes * lateDeductionRate / 60).toFixed(2)),
    });
  } catch (error) {
    await connection.rollback();
    throw error;
  } finally {
    connection.release();
  }
}

app.post("/api/attendance", requireRoles("admin", "hr"), asyncRoute((request, response) => saveAttendance(request, response, false)));
app.put("/api/attendance/:attendanceId", requireRoles("admin", "hr"), asyncRoute((request, response) => saveAttendance(request, response, true)));

app.delete("/api/attendance/:attendanceId", requireRoles("admin", "hr"), asyncRoute(async (request, response) => {
  const attendanceId = readId(request.params.attendanceId, "attendance ID");
  const connection = await pool.getConnection();
  try {
    await connection.beginTransaction();
    const [records] = await connection.execute(
      "SELECT EmployeeID FROM tbl_attendance WHERE AttendanceID = ? FOR UPDATE",
      [attendanceId],
    );
    if (!records.length) throw httpError(404, "Attendance record not found.");
    await connection.execute("DELETE FROM tbl_attendance WHERE AttendanceID = ?", [attendanceId]);
    await updateEmployeeSalary(connection, records[0].EmployeeID, currentMonth());
    await connection.commit();
    response.status(204).end();
  } catch (error) {
    await connection.rollback();
    throw error;
  } finally {
    connection.release();
  }
}));

// Refresh payroll and return all manager rows or the signed-in employee's row.
app.get("/api/salaries", requireRoles("admin", "hr", "employee"), asyncRoute(async (request, response) => {
  const month = request.query.month ? readMonth(request.query.month) : currentMonth();
  const employeeId = request.user.role === "employee"
    ? await getEmployeeIdForUser(request.user)
    : null;
  const connection = await pool.getConnection();
  try {
    await connection.beginTransaction();
    const calculations = employeeId === null
      ? await refreshMonthlySalaries(connection, month)
      : new Map([[employeeId, await updateEmployeeSalary(connection, employeeId, month)]]);
    const [salaries] = await connection.execute(
      `SELECT s.SalaryID AS salaryId, s.EmployeeID AS employeeId,
              s.\`Basic Salary\` AS basicSalary, s.\`Working Hours\` AS totalWorkingHours,
              s.\`Overtime Hours\` AS overtimeHours, s.\`Overtime Pay\` AS overtimePay,
              s.\`Absent Days\` AS absentDays, s.\`Absence Deduction\` AS absenceDeduction,
              s.\`Late Deduction\` AS lateDeduction,
              s.Deductions AS deductions, s.\`Net Salary\` AS netSalary
       FROM tbl_salary s
       JOIN tbl_employees e ON e.EmployeeID = s.EmployeeID
       WHERE (? IS NULL OR s.EmployeeID = ?)
       ORDER BY s.EmployeeID`,
      [employeeId, employeeId],
    );
    await connection.commit();
    response.json({
      month,
      salaries: salaries.map((salary) => ({
        ...salary,
        ...calculations.get(Number(salary.employeeId)),
      })),
    });
  } catch (error) {
    await connection.rollback();
    throw error;
  } finally {
    connection.release();
  }
}));

// Manage accounts, detect unlinked employees, and list export files.
app.get("/api/users", requireRoles("admin", "hr"), asyncRoute(async (request, response) => {
  const [users] = await pool.query(
    `SELECT UserID AS userId, Username AS username, \`Account Role\` AS accountRole
     FROM tbl_users
     ORDER BY UserID`,
  );
  response.json(users);
}));

app.delete("/api/users/:userId", requireRoles("admin", "hr"), asyncRoute(async (request, response) => {
  const userId = readId(request.params.userId, "user ID");
  if (userId === Number(request.user.userId)) {
    throw httpError(409, "You cannot delete the account currently signed in.");
  }

  const connection = await pool.getConnection();
  try {
    await connection.beginTransaction();
    const [users] = await connection.execute(
      "SELECT UserID AS userId, `Account Role` AS accountRole FROM tbl_users WHERE UserID = ? FOR UPDATE",
      [userId],
    );
    if (!users.length) throw httpError(404, "User account not found.");

    if (["admin", "hr"].includes(normalizeRole(users[0].accountRole))) {
      const [otherUsers] = await connection.execute(
        "SELECT `Account Role` AS accountRole FROM tbl_users WHERE UserID <> ? FOR UPDATE",
        [userId],
      );
      const hasOtherManager = otherUsers.some((user) => ["admin", "hr"].includes(normalizeRole(user.accountRole)));
      if (!hasOtherManager) {
        throw httpError(409, "At least one HR Personnel or Super Administrator account must remain.");
      }
    }

    await connection.execute("DELETE FROM tbl_users WHERE UserID = ?", [userId]);
    await connection.commit();
    for (const [sessionId, session] of sessions) {
      if (Number(session.user.userId) === userId) sessions.delete(sessionId);
    }
    response.status(204).end();
  } catch (error) {
    await connection.rollback();
    throw error;
  } finally {
    connection.release();
  }
}));

app.get("/api/employees/without-account", requireRoles("admin", "hr"), asyncRoute(async (request, response) => {
  const [[employees], [users]] = await Promise.all([
    pool.query(
      `SELECT EmployeeID AS employeeId, Name AS name, Email AS email, Username AS username
       FROM tbl_employees
       ORDER BY EmployeeID`,
    ),
    pool.query(
      `SELECT Username AS username, \`Password Hash\` AS passwordHash,
              \`Account Role\` AS accountRole
       FROM tbl_users`,
    ),
  ]);
  const existingAccounts = new Set(
    users
      .filter((user) => normalizeRole(user.accountRole) && parsePasswordHash(user.passwordHash))
      .map((user) => normalizeAccountUsername(user.username)),
  );
  response.json(
    employees
      .filter((employee) => !employee.username || !existingAccounts.has(normalizeAccountUsername(employee.username)))
      .map(({ username, ...employee }) => ({
        ...employee,
        suggestedUsername: username || employeeUsername(employee.name, employee.employeeId),
      })),
  );
}));

app.post("/api/users", requireRoles("admin", "hr"), asyncRoute(async (request, response) => {
  const body = request.body || {};
  const password = typeof body.password === "string" ? body.password : "";
  const accountRole = body.accountRole;
  const employeeId = accountRole === "Employee" ? readId(body.employeeId, "employee ID") : null;
  const username = readText(body.username, "Username", 100);
  const allowedRoles = ["Super Administrator", "HR Personnel"];

  if (password.length < 8 || password.length > 255) {
    throw httpError(400, "Password must be between 8 and 255 characters.");
  }
  if (!allowedRoles.includes(accountRole) && accountRole !== "Employee") {
    throw httpError(400, "Choose a valid account role.");
  }

  if (employeeId) {
    const connection = await pool.getConnection();
    try {
      await connection.beginTransaction();
      const [employees] = await connection.execute(
        "SELECT EmployeeID AS employeeId, Name AS name, Username AS username FROM tbl_employees WHERE EmployeeID = ? FOR UPDATE",
        [employeeId],
      );
      if (!employees.length) throw httpError(404, "Employee not found.");
      const employee = employees[0];
      const employeeAccountUsername = employee.username || employeeUsername(employee.name, employee.employeeId);
      if (username !== employeeAccountUsername) {
        throw httpError(400, "The employee username does not match the employee account.");
      }

      const [existingUsers] = await connection.execute(
        "SELECT UserID AS userId, `Account Role` AS accountRole FROM tbl_users WHERE Username = ? FOR UPDATE",
        [employeeAccountUsername],
      );
      if (existingUsers.length) {
        if (normalizeRole(existingUsers[0].accountRole) !== "employee") {
          throw httpError(409, "That username already belongs to a non-employee account.");
        }
        await connection.execute(
          "UPDATE tbl_users SET `Password Hash` = ?, `Account Role` = 'Employee' WHERE UserID = ?",
          [await hashPassword(password), existingUsers[0].userId],
        );
      } else {
        await connection.execute(
          "INSERT INTO tbl_users (Username, `Password Hash`, `Account Role`) VALUES (?, ?, 'Employee')",
          [employeeAccountUsername, await hashPassword(password)],
        );
      }

      await connection.execute(
        "UPDATE tbl_employees SET Username = ? WHERE EmployeeID = ?",
        [employeeAccountUsername, employeeId],
      );
      await connection.commit();
      response.status(201).json({
        username: employeeAccountUsername,
        accountRole: "Employee",
        employeeId,
      });
    } catch (error) {
      await connection.rollback();
      throw error;
    } finally {
      connection.release();
    }
    return;
  }

  const [existingUsers] = await pool.execute(
    "SELECT UserID FROM tbl_users WHERE Username = ? LIMIT 1",
    [username],
  );
  if (existingUsers.length) throw httpError(409, "That username already exists.");

  const passwordHash = await hashPassword(password);
  try {
    const [result] = await pool.execute(
      "INSERT INTO tbl_users (Username, `Password Hash`, `Account Role`) VALUES (?, ?, ?)",
      [username, passwordHash, accountRole],
    );
    response.status(201).json({
      userId: result.insertId,
      username,
      accountRole,
    });
  } catch (error) {
    if (error.code === "ER_DUP_ENTRY") {
      throw httpError(409, "That username already exists.");
    }
    throw error;
  }
}));

// Invalidate the current session and clear its browser cookie.
app.post("/api/logout", (request, response) => {
  const sessionId = getCookie(request, "ems_session");
  if (sessionId) sessions.delete(sessionId);
  clearSessionCookie(response);
  response.status(204).end();
});

// Serve frontend files and convert uncaught route errors into consistent API responses.
app.use(express.static(publicDirectory, { index: false }));

app.use((error, request, response, next) => {
  if (response.headersSent) return next(error);
  const status = Number(error.status) || 500;
  if (status >= 500) console.error("EMS request failed:", error.code || error.message);
  response.status(status).json({
    error: status >= 500 ? "The request could not be completed." : error.message,
  });
});

// Generate the previous month's attendance workbook when automatic export is enabled.
let monthlyAttendanceExportRunning = false;

async function processMonthlyAttendanceExport() {
  if (monthlyAttendanceExportRunning) return;
  monthlyAttendanceExportRunning = true;
  let temporaryPath = null;
  try {
    const [settings] = await pool.execute(
      `SELECT SettingValue AS settingValue FROM ${automaticMonthlyExportSettingsTable} WHERE SettingKey = ?`,
      [automaticMonthlyExportSetting],
    );
    if (!settings.length) throw new Error("Automatic monthly export setting is missing.");
    if (settings[0].settingValue !== "true") return;

    const month = previousMonth(currentMonth());
    await fs.mkdir(exportsDirectory, { recursive: true });
    const filename = `Monthly_Attendance_${month}.xlsx`;
    const filePath = path.join(exportsDirectory, filename);
    try {
      await fs.access(filePath);
      return;
    } catch (error) {
      if (error.code !== "ENOENT") throw error;
    }

    const rows = await monthlyAttendanceExportRows(month);
    const workbook = await createExcelWorkbook("Attendance Records", attendanceExportColumns, rows);
    temporaryPath = path.join(
      exportsDirectory,
      `.${filename}.${process.pid}.${crypto.randomBytes(6).toString("hex")}.tmp`,
    );
    await fs.writeFile(temporaryPath, workbook);
    await fs.rename(temporaryPath, filePath);
    temporaryPath = null;
    console.log(`Saved monthly attendance export: ${filePath}`);
  } catch (error) {
    console.error("Automatic monthly attendance export failed:", error.code || error.message);
  } finally {
    if (temporaryPath) {
      try {
        await fs.rm(temporaryPath, { force: true });
      } catch (error) {
        console.error("Could not remove temporary attendance export:", error.code || error.message);
      }
    }
    monthlyAttendanceExportRunning = false;
  }
}

// Rename workbooks created by the earlier attendance-export naming convention.
async function renameLegacyAttendanceExports() {
  let filenames;
  try {
    filenames = await fs.readdir(exportsDirectory);
  } catch (error) {
    if (error.code === "ENOENT") return;
    throw error;
  }
  for (const filename of filenames) {
    const match = filename.match(/^attendance_records_(\d{4}-(?:0[1-9]|1[0-2]))\.xlsx$/);
    if (!match) continue;
    const newPath = path.join(exportsDirectory, `Monthly_Attendance_${match[1]}.xlsx`);
    try {
      await fs.access(newPath);
    } catch (error) {
      if (error.code !== "ENOENT") throw error;
      await fs.rename(path.join(exportsDirectory, filename), newPath);
    }
  }
}

// Migrate the persisted auto-export preference to its feature-specific table.
async function ensureAutomaticMonthlyExportSetting(connection) {
  const [tables] = await connection.execute(
    `SELECT TABLE_NAME AS tableName FROM information_schema.TABLES
     WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME IN (?, ?)`,
    [automaticMonthlyExportSettingsTable, legacySettingsTable],
  );
  const existingTables = new Set(tables.map((table) => table.tableName));
  if (existingTables.has(legacySettingsTable) && !existingTables.has(automaticMonthlyExportSettingsTable)) {
    await connection.query(
      `RENAME TABLE ${legacySettingsTable} TO ${automaticMonthlyExportSettingsTable}`,
    );
  } else if (existingTables.has(legacySettingsTable)) {
    await connection.execute(
      `INSERT IGNORE INTO ${automaticMonthlyExportSettingsTable} (SettingKey, SettingValue)
       SELECT SettingKey, SettingValue FROM ${legacySettingsTable} WHERE SettingKey = ?`,
      [automaticMonthlyExportSetting],
    );
    await connection.execute(
      `DELETE FROM ${legacySettingsTable} WHERE SettingKey = ?`,
      [automaticMonthlyExportSetting],
    );
  }

  await connection.query(
    `CREATE TABLE IF NOT EXISTS ${automaticMonthlyExportSettingsTable} (
      SettingKey VARCHAR(100) NOT NULL PRIMARY KEY,
      SettingValue VARCHAR(255) NOT NULL
    )`,
  );
  await connection.execute(
    `INSERT IGNORE INTO ${automaticMonthlyExportSettingsTable} (SettingKey, SettingValue)
     VALUES (?, 'true')`,
    [automaticMonthlyExportSetting],
  );
}

// Migrate the username link, repair safe matches, and enforce referential integrity.
async function ensureEmployeeUsernameRelationship(connection) {
  const [columns] = await connection.execute(
    `SELECT COLUMN_NAME FROM information_schema.COLUMNS
     WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'tbl_employees'`,
  );
  const columnNames = new Set(columns.map((column) => column.COLUMN_NAME));
  if (!columnNames.has("Username")) {
    await connection.query(
      "ALTER TABLE tbl_employees ADD COLUMN Username VARCHAR(100) NULL DEFAULT NULL AFTER Name",
    );
  }

  const [[employees], [users]] = await Promise.all([
    connection.query(
      "SELECT EmployeeID AS employeeId, Name AS name, Username AS username FROM tbl_employees",
    ),
    connection.query(
      "SELECT Username AS username, `Password Hash` AS passwordHash, `Account Role` AS accountRole FROM tbl_users",
    ),
  ]);
  const employeesByName = new Map();
  for (const employee of employees) {
    if (employee.username) continue;
    const key = normalizeAccountUsername(employee.name);
    if (!employeesByName.has(key)) employeesByName.set(key, []);
    employeesByName.get(key).push(employee);
  }
  const usersByName = new Map();
  for (const user of users) {
    if (normalizeRole(user.accountRole) !== "employee" || !parsePasswordHash(user.passwordHash)) continue;
    const key = normalizeAccountUsername(user.username);
    if (!usersByName.has(key)) usersByName.set(key, []);
    usersByName.get(key).push(user);
  }
  for (const [key, matchingEmployees] of employeesByName) {
    const matchingUsers = usersByName.get(key) || [];
    if (matchingEmployees.length !== 1 || matchingUsers.length !== 1) continue;
    await connection.execute(
      "UPDATE tbl_employees SET Username = ? WHERE EmployeeID = ? AND Username IS NULL",
      [matchingUsers[0].username, matchingEmployees[0].employeeId],
    );
  }

  const [indexes] = await connection.execute(
    `SELECT INDEX_NAME FROM information_schema.STATISTICS
     WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'tbl_employees'
       AND INDEX_NAME = 'uq_tbl_employees_username'`,
  );
  if (!indexes.length) {
    await connection.query(
      "ALTER TABLE tbl_employees ADD UNIQUE INDEX uq_tbl_employees_username (Username)",
    );
  }

  const [foreignKeys] = await connection.execute(
    `SELECT REFERENCED_TABLE_NAME AS referencedTable, REFERENCED_COLUMN_NAME AS referencedColumn
     FROM information_schema.KEY_COLUMN_USAGE
     WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'tbl_employees'
       AND COLUMN_NAME = 'Username' AND REFERENCED_TABLE_NAME IS NOT NULL`,
  );
  if (!foreignKeys.length) {
    await connection.query(
      `ALTER TABLE tbl_employees
       ADD CONSTRAINT fk_tbl_employees_username_user_username
       FOREIGN KEY (Username) REFERENCES tbl_users (Username)
       ON UPDATE CASCADE ON DELETE SET NULL`,
    );
  } else if (foreignKeys.length !== 1
      || foreignKeys[0].referencedTable !== "tbl_users"
      || foreignKeys[0].referencedColumn !== "Username") {
    throw new Error("tbl_employees.Username has an unexpected foreign-key constraint.");
  }
}

// Apply schema migrations, recalculate totals, catch up schedulers, and start Express.
async function start() {
  const connection = await pool.getConnection();
  try {
    await ensureAutomaticMonthlyExportSetting(connection);
    await ensureEmployeeUsernameRelationship(connection);
    const [existingColumns] = await connection.execute(
      `SELECT COLUMN_NAME FROM information_schema.COLUMNS
       WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'tbl_salary'`,
    );
    const columnNames = new Set(existingColumns.map((column) => column.COLUMN_NAME));
    const payrollColumns = [
      ["Working Hours", "DECIMAL(7,2) NOT NULL DEFAULT 0"],
      ["Overtime Hours", "DECIMAL(7,2) NOT NULL DEFAULT 0"],
      ["Absent Days", "INT NOT NULL DEFAULT 0"],
      ["Absence Deduction", "DECIMAL(12,2) NOT NULL DEFAULT 0"],
      ["Late Deduction", "DECIMAL(12,2) NOT NULL DEFAULT 0"],
    ];
    for (const [column, definition] of payrollColumns) {
      if (!columnNames.has(column)) {
        await connection.query(`ALTER TABLE tbl_salary ADD COLUMN \`${column}\` ${definition}`);
      }
    }

    await connection.beginTransaction();
    const [records] = await connection.query(
      "SELECT AttendanceID, Date AS date, `Time in` AS timeIn, `Time out` AS timeOut FROM tbl_attendance",
    );
    for (const record of records) {
      if (isAbsentAttendance(record)) {
        await connection.execute(
          "UPDATE tbl_attendance SET Overtime = 0, `Total Working Hours` = 0 WHERE AttendanceID = ?",
          [record.AttendanceID],
        );
        continue;
      }
      const workday = calculateWorkday(record.date, record.timeIn, record.timeOut);
      await connection.execute(
        "UPDATE tbl_attendance SET Overtime = ?, `Total Working Hours` = ? WHERE AttendanceID = ?",
        [workday.overtimeHours, workday.totalWorkingHours, record.AttendanceID],
      );
    }
    await refreshMonthlySalaries(connection, currentMonth());
    await connection.commit();
  } catch (error) {
    await connection.rollback();
    throw error;
  } finally {
    connection.release();
  }
  await processCompletedWeekdays();
  await renameLegacyAttendanceExports();
  await processMonthlyAttendanceExport();
  setInterval(() => {
    processCompletedWeekdays();
    processMonthlyAttendanceExport();
  }, 30 * 1000);
  app.listen(port, process.env.HOST || "127.0.0.1", () => {
    console.log(`EMS is listening at http://localhost:${port}`);
  });
}

start().catch((error) => {
  console.error("Could not start EMS:", error.code || error.message);
  process.exitCode = 1;
});