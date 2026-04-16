# Church Dashboard (Node.js)

A focused, self-hosted church operations interface for:

- People management (names, numbers, birthdays, notes)
- Gracely-style person profile layout (tabs + cards + quick actions)
- Follow-ups and visit logging per person
- Global follow-up queue
- CSV people import (for exported data migration)
- Metrics and attendance integration
- Yearly calendar planning
- Registration forms (starter level)
- Visitation map sections (draw and track)

## Stack

- Node.js + Express
- EJS server-rendered UI
- Tabler UI kit (open source admin framework via CDN)
- File-based JSON persistence (`data/data.json`)
- csv-parse (CSV import parser)
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

## Hostinger deployment (Node.js Apps)

1. Push this project to GitHub.
2. In hPanel choose `Websites -> Add Website -> Node.js Apps`.
3. Import Git repository.
4. For framework type, `Express.js` or `Other` both work.
5. Build command: `npm install`
6. Start/entry file: `server.js`
7. Deploy.

If using `ATTENDANCE_APP_URL`, set it in Node.js app environment variables in hPanel.

## Data backup

Primary data is in:

- `data/data.json`

Back this file up regularly.
