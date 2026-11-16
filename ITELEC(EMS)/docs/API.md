# EMS API Reference

Base URL: `http://127.0.0.1:3000`

All API routes use the `/api` prefix. Send JSON request bodies with
`Content-Type: application/json`. Except for `POST /api/login`, routes require the
HTTP-only `ems_session` cookie created at login. Postman retains and sends this cookie
automatically when requests use the same host.

## Authentication and response format

| Method | Path | Access | Description |
|---|---|---|---|
| `POST` | `/api/login` | Public | Sign in and create a session cookie. Body: `{"username":"...","password":"..."}`. Returns the username, account role, and normalized role. |
| `GET` | `/api/session` | Signed-in user | Return the current identity; includes `employeeId` for a linked Employee account. |
| `POST` | `/api/logout` | Signed-in user | End the current session and clear its cookie. Returns `204 No Content`. |

JSON errors have the shape `{"error":"..."}`. Common status codes are `400` for
invalid input, `401` for missing/invalid authentication, `403` for insufficient role,
`404` for missing records/files, and `409` for duplicate/conflicting data. Unexpected
server errors return a generic message; the server logs the underlying error.

## Dashboard and employees

| Method | Path | Access | Description |
|---|---|---|---|
| `GET` | `/api/dashboard?date=YYYY-MM-DD` | Signed-in user | Counts and attendance for a date (defaults to today). Employee data is scoped to the signed-in employee. |
| `GET` | `/api/employees?search=text` | Signed-in user | Search the employee directory by name, ID, or email. The search parameter is optional. Currently returns the full matching directory to any signed-in role. |
| `POST` | `/api/employees` | HR/Admin | Create an employee and initial salary row. Requires the employee fields below; returns `201` and `employeeId`. |
| `PUT` | `/api/employees/:employeeId` | HR/Admin | Replace the employee fields below. Returns the employee ID. |
| `DELETE` | `/api/employees/:employeeId` | HR/Admin | Delete the employee and associated attendance, salary, and linked Employee account. Returns `204`. |
| `GET` | `/api/exports/employees.xlsx` | HR/Admin | Download the employee directory as an Excel workbook. |

Employee create/update JSON:

```json
{
  "name": "Alex Example",
  "gender": "Other",
  "birthDate": "1995-01-20",
  "contactNumber": "555-0100",
  "email": "alex@example.com",
  "position": "Analyst",
  "address": "Example address"
}
```

## Attendance

| Method | Path | Access | Description |
|---|---|---|---|
| `GET` | `/api/attendance?date=YYYY-MM-DD` | Signed-in user | List attendance. Without a date, returns recorded rows; with a date, includes eligible missing weekday absences. Employee records are scoped to the signed-in employee. |
| `POST` | `/api/attendance` | HR/Admin | Create attendance, or update the existing row for that employee/date. |
| `PUT` | `/api/attendance/:attendanceId` | HR/Admin | Update an attendance row. |
| `DELETE` | `/api/attendance/:attendanceId` | HR/Admin | Delete an attendance row. Returns `204`. |
| `POST` | `/api/attendance/absences` | HR/Admin | Save missing weekday absence rows. Body: `{"date":"YYYY-MM-DD"}`. |
| `GET` | `/api/exports/attendance.xlsx?month=YYYY-MM` | HR/Admin | Generate and download the selected month's attendance workbook. |
| `GET` | `/api/exports/attendance/monthly` | HR/Admin | List generated monthly workbooks. |
| `GET` | `/api/exports/attendance/monthly/:month.xlsx` | HR/Admin | Download an existing generated workbook, where `:month` is `YYYY-MM`. |

Attendance create/update JSON:

```json
{
  "employeeId": 1,
  "date": "2026-11-13",
  "timeIn": "08:00",
  "timeOut": "17:00"
}
```

Use `00:00` for both times to record an absence. Dates must be real calendar dates;
months must use `YYYY-MM`; times use `HH:mm` or `HH:mm:ss`.

## Payroll, accounts, and settings

| Method | Path | Access | Description |
|---|---|---|---|
| `GET` | `/api/salaries?month=YYYY-MM` | Signed-in user | Recalculate and return payroll for the month. Employee users receive only their own row. |
| `GET` | `/api/users` | HR/Admin | List user accounts without password hashes. |
| `POST` | `/api/users` | HR/Admin | Create an HR/Admin or Employee account. Requires `username`, `password` (8–255 characters), and `accountRole` (`Super Administrator`, `HR Personnel`, or `Employee`). Employee accounts also require `employeeId` and the employee's matching username. Returns `201`. |
| `DELETE` | `/api/users/:userId` | HR/Admin | Delete an account and invalidate its active sessions. Cannot delete the signed-in account or the last HR/Admin account. A linked employee record is retained and unlinked. Returns `204`. |
| `GET` | `/api/employees/without-account` | HR/Admin | List employees without a valid linked account and suggested usernames. |
| `GET` | `/api/account-settings` | Employee | Return the signed-in Employee's username and account role. |
| `PUT` | `/api/account-settings` | Employee | Change account credentials. Requires `username` and `currentPassword`; `newPassword` and matching `confirmPassword` are optional. A new password must be 8–255 characters. |
| `GET` | `/api/settings/automatic-monthly-export` | HR/Admin | Return `{"enabled":true}` or `{"enabled":false}`. |
| `PUT` | `/api/settings/automatic-monthly-export` | HR/Admin | Set the preference with `{"enabled":true}` or `{"enabled":false}`. |

## Role names

The server normalizes `Super Administrator` and `Admin` to `admin`; `HR Personnel`
and `HR` to `hr`; and `Employee` to `employee`. In this reference, **HR/Admin** means
either normalized manager role.

## CORS and local frontend

The server allows credentialed requests from `http://localhost:5500` and
`http://127.0.0.1:5500` for Live Server development. Requests from other browser
origins are not enabled by default.
