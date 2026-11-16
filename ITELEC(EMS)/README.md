# Employee Management System

A web-based employee management system built with HTML, CSS, browser JavaScript,
Node.js, Express, and MySQL (`mysql2`). The frontend communicates with the server
through JSON API endpoints and Fetch API requests.

## Requirements

- Node.js 18 or newer
- npm
- MySQL with the existing EMS tables: `tbl_employees`, `tbl_users`,
  `tbl_attendance`, and `tbl_salary`

The server creates the EMS settings table and applies the additional schema changes it
needs, but it does not create the four base tables. See
[md files/explanation.md](<./md files/explanation.md>) for the expected schema and
application behavior.

## Setup

1. Install dependencies:

   ```sh
   npm install
   ```

2. Copy `.env.example` to `.env` and set the database values for your local MySQL
   installation. In particular, set `DB_PASSWORD` to the password for `DB_USER`.
   Keep `.env` private; it is excluded from Git.
3. Start the server:

   ```sh
   npm start
   ```

4. Open [http://127.0.0.1:3000](http://127.0.0.1:3000) in a browser and sign in with
   an account already present in `tbl_users`.

The server also accepts these environment variables: `HOST`, `PORT`, `DB_HOST`,
`DB_PORT`, `DB_USER`, `DB_PASSWORD`, and `DB_NAME`. A `.env` file is optional when
the variables are supplied by the operating system. Existing operating-system values
take precedence over `.env`.

For local development with VS Code Live Server, the frontend may be opened on port
5500; the application is configured to send API requests to the server on port 3000.

## Checks and API testing

Run the local JavaScript syntax checks with:

```sh
npm test
```

For API smoke testing, import
[`postman/EMS.postman_collection.json`](./postman/EMS.postman_collection.json) into
Postman. Set its `baseUrl`, `username`, `password`, `date`, and `month` collection
variables. Use a valid HR Personnel or Super Administrator account, start the server,
and run the collection in order. The collection does not create, edit, or delete
employee or attendance records. Its payroll request does recalculate and persist
derived salary totals, as the payroll API normally does. Postman keeps the session
cookie from the login response for subsequent requests. Store real credentials in a
local Postman environment rather than saving them in or exporting the collection.

See [docs/API.md](./docs/API.md) for endpoint methods, roles, request formats, and
response behavior.

## Project documentation

- [API reference](./docs/API.md)
- [Code and project explanation](<./md files/explanation.md>)
- [Features and limitations](<./md files/features.md>)
- [Presentation script](<./md files/presentation.md>)
- [Final demonstration report](<./md files/reporting.md>)
