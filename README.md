# Church Dashboard (Node.js)

A focused, self-hosted church operations interface for:

- People management (names, numbers, birthdays, notes)
- Gracely-style person profile layout (tabs + cards + quick actions)
- Follow-ups and visit logging per person
- Global follow-up queue
- CSV and Excel people import (XLS/XLSX export migration)
- Membership type classification (Prospect, Member, Voting Member)
- Metrics and attendance integration
- Yearly calendar planning
- Calendar birthday overlay toggle
- Registration forms (starter level)
- Visitation map sections (draw and track)

## Stack

- Node.js + Express
- EJS server-rendered UI
- Tabler UI kit (open source admin framework via CDN)
- Persistence:
  - MySQL-backed JSON state when `DB_HOST/DB_NAME/DB_USER/DB_PASSWORD` env vars are set
  - File fallback (`data/data.json`) when DB vars are not set
- csv-parse + xlsx + multer (CSV/Excel import parser and upload)
- FullCalendar (calendar UI)
- Leaflet + Leaflet.draw (visitation map sections)

## Local run

```bash
npm install
npm run dev
```

Open `http://localhost:3000`.

## Environment variables

- `PORT` (default `3000`)
- `ATTENDANCE_APP_URL` (optional): URL to your existing attendance app to embed in Metrics tab
- `APP_SECRET` (recommended): session secret for login cookies
- `CRM_ADMIN_PASSWORD` (recommended): enables login protection for all CRM routes except `/register/*`
  - Compatibility aliases also accepted: `ADMIN_PASSWORD`, `APP_PASSWORD`, `PASSWORD`
- `DB_HOST` (optional): MySQL host
- `DB_PORT` (optional, default `3306`): MySQL port
- `DB_NAME` (optional): MySQL database
- `DB_USER` (optional): MySQL user
- `DB_PASSWORD` (optional): MySQL password

## Hostinger deployment (Node.js Apps)

1. Push this project to GitHub.
2. In hPanel choose `Websites -> Add Website -> Node.js Apps`.
3. Import Git repository.
4. For framework type, `Express.js` or `Other` both work.
5. Build command: `npm install`
6. Start/entry file: `server.js`
7. Deploy.

If using `ATTENDANCE_APP_URL`, DB vars, or login vars, set them in Node.js app environment variables in hPanel.

## Data backup

Primary data is in:

- `data/data.json`

Back this file up regularly.
