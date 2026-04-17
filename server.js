const path = require('path');
const crypto = require('crypto');
const express = require('express');
const session = require('express-session');
const multer = require('multer');
const XLSX = require('xlsx');
const cheerio = require('cheerio');
const { parse } = require('csv-parse/sync');
const { readData, updateData } = require('./lib/dataStore');

const app = express();
const PORT = process.env.PORT || 3000;
app.set('trust proxy', 1);

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use('/static', express.static(path.join(__dirname, 'public')));
app.use(
  session({
    name: 'crm.sid',
    secret: process.env.APP_SECRET || 'dev-only-secret-change-me',
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: 'lax',
      secure: 'auto',
      maxAge: 1000 * 60 * 60 * 12
    }
  })
);
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 8 * 1024 * 1024 }
});

function id() {
  return crypto.randomUUID();
}

function sortByName(list) {
  return [...list].sort((a, b) => a.name.localeCompare(b.name));
}

function pad2(value) {
  return String(value).padStart(2, '0');
}

function calculateAge(birthday) {
  if (!birthday) return null;

  const today = new Date();
  const normalizedBirthday = toIsoDate(birthday);
  const dob = new Date(normalizedBirthday);

  if (Number.isNaN(dob.getTime())) return null;

  let age = today.getFullYear() - dob.getFullYear();
  const monthDiff = today.getMonth() - dob.getMonth();

  if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < dob.getDate())) {
    age -= 1;
  }

  return age;
}

function profileCompletion(person) {
  const fields = [
    'name',
    'phone',
    'email',
    'birthday',
    'sectionId',
    'notes',
    'gender',
    'ageGroup',
    'occupation',
    'language',
    'maritalStatus',
    'allergies',
    'emergencyContact',
    'medicalNotes',
    'address',
    'city',
    'state',
    'zipCode'
  ];

  const completed = fields.filter((field) => (person[field] || '').toString().trim().length > 0).length;

  return Math.round((completed / fields.length) * 100);
}

function enrichPerson(person) {
  return {
    ...person,
    age: calculateAge(person.birthday),
    profilePercent: profileCompletion(person)
  };
}

function normalize(value) {
  return (value || '').toString().trim();
}

function normalizePhone(value) {
  return normalize(value).replace(/[^\d]/g, '');
}

function normalizeHeaderKey(value) {
  return normalize(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function normalizeImportRow(row) {
  return Object.entries(row || {}).reduce((acc, [key, value]) => {
    const normalizedKey = normalizeHeaderKey(key);
    if (normalizedKey) {
      acc[normalizedKey] = value;
    }
    return acc;
  }, {});
}

function csvField(row, candidates) {
  const normalizedRow = normalizeImportRow(row);

  for (const key of candidates) {
    const normalizedKey = normalizeHeaderKey(key);
    if (Object.prototype.hasOwnProperty.call(normalizedRow, normalizedKey) && normalize(normalizedRow[normalizedKey])) {
      return normalize(normalizedRow[normalizedKey]);
    }
  }
  return '';
}

function detectDelimiter(csvText) {
  const firstDataLine = csvText
    .split(/\r?\n/)
    .find((line) => normalize(line).length > 0) || '';

  const delimiters = [',', ';', '\t', '|'];
  const counts = delimiters.map((delimiter) => ({
    delimiter,
    count: firstDataLine.split(delimiter).length - 1
  }));

  counts.sort((a, b) => b.count - a.count);
  return counts[0].count > 0 ? counts[0].delimiter : ',';
}

function toIsoDate(value) {
  const raw = normalize(value);
  if (!raw) return '';

  const dmyMatch = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmyMatch) {
    const day = Number(dmyMatch[1]);
    const month = Number(dmyMatch[2]);
    const year = Number(dmyMatch[3]);

    const parsed = new Date(Date.UTC(year, month - 1, day));
    if (
      parsed.getUTCFullYear() === year &&
      parsed.getUTCMonth() === month - 1 &&
      parsed.getUTCDate() === day
    ) {
      return `${year}-${pad2(month)}-${pad2(day)}`;
    }
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    return raw;
  }

  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) {
    return '';
  }

  return parsed.toISOString().slice(0, 10);
}

function formatDateDMY(value) {
  const iso = toIsoDate(value);
  if (!iso) return '-';
  const [year, month, day] = iso.split('-');
  return `${day}/${month}/${year}`;
}

function nextBirthdayDate(birthday, now = new Date()) {
  const iso = toIsoDate(birthday);
  if (!iso) return null;

  const [, monthString, dayString] = iso.split('-');
  const month = Number(monthString);
  const day = Number(dayString);
  const year = now.getFullYear();

  const currentYearBirthday = new Date(year, month - 1, day);
  if (Number.isNaN(currentYearBirthday.getTime()) || currentYearBirthday.getMonth() !== month - 1) {
    // Feb 29 fallback for non-leap years.
    if (month === 2 && day === 29) {
      const fallback = new Date(year, 1, 28);
      if (fallback >= new Date(now.getFullYear(), now.getMonth(), now.getDate())) {
        return fallback;
      }
      return new Date(year + 1, 1, 28);
    }
    return null;
  }

  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  if (currentYearBirthday >= todayStart) {
    return currentYearBirthday;
  }
  return new Date(year + 1, month - 1, day);
}

function diffDays(dateA, dateB) {
  const startA = new Date(dateA.getFullYear(), dateA.getMonth(), dateA.getDate());
  const startB = new Date(dateB.getFullYear(), dateB.getMonth(), dateB.getDate());
  return Math.round((startA - startB) / (24 * 60 * 60 * 1000));
}

function normalizeServiceType(value) {
  const serviceType = normalize(value);
  return ['wed', 'sun_am', 'sun_pm'].includes(serviceType) ? serviceType : 'sun_am';
}

function getLegacyAppUrl() {
  return normalize(process.env.METRICS_APP_URL) || normalize(process.env.ATTENDANCE_APP_URL);
}

function getLegacyAppPassword() {
  const candidates = [
    process.env.LEGACY_APP_PASSWORD,
    process.env.LEGACY_ADMIN_PASSWORD,
    process.env.ATTENDANCE_APP_PASSWORD,
    process.env.ADMIN_PASSWORD,
    process.env.CRM_ADMIN_PASSWORD,
    process.env.APP_PASSWORD,
    process.env.PASSWORD
  ];
  const first = candidates.find((value) => normalize(value).length > 0);
  return normalize(first);
}

function getSetCookieHeaders(headers) {
  if (!headers) return [];

  if (typeof headers.getSetCookie === 'function') {
    return headers.getSetCookie();
  }

  const headerValue = headers.get('set-cookie');
  if (!headerValue) return [];
  return [headerValue];
}

function parseCookiePairs(rawSetCookie) {
  const pairs = [];
  const value = normalize(rawSetCookie);
  if (!value) return pairs;

  const directPair = value.split(';')[0];
  if (directPair.includes('=')) {
    pairs.push(directPair.trim());
  }

  const pattern = /(?:^|,\s*)([^=,\s;]+)=([^;,\r\n]*)/g;
  let match;
  while ((match = pattern.exec(value)) !== null) {
    const cookieName = normalize(match[1]);
    const cookieValue = normalize(match[2]);
    if (!cookieName) continue;
    const pair = `${cookieName}=${cookieValue}`;
    if (!pairs.includes(pair)) {
      pairs.push(pair);
    }
  }

  return pairs;
}

function mergeCookieHeader(existingHeader, setCookieHeaders) {
  const cookieMap = new Map();

  normalize(existingHeader)
    .split(';')
    .map((part) => part.trim())
    .filter(Boolean)
    .forEach((pair) => {
      const separatorIndex = pair.indexOf('=');
      if (separatorIndex <= 0) return;
      const key = pair.slice(0, separatorIndex).trim();
      const value = pair.slice(separatorIndex + 1).trim();
      if (key) cookieMap.set(key, value);
    });

  (setCookieHeaders || []).forEach((setCookie) => {
    parseCookiePairs(setCookie).forEach((pair) => {
      const separatorIndex = pair.indexOf('=');
      if (separatorIndex <= 0) return;
      const key = pair.slice(0, separatorIndex).trim();
      const value = pair.slice(separatorIndex + 1).trim();
      if (key) cookieMap.set(key, value);
    });
  });

  return Array.from(cookieMap.entries())
    .map(([key, value]) => `${key}=${value}`)
    .join('; ');
}

async function legacyFetch(url, options = {}) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15000);

  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(timeout);
  }
}

async function fetchLegacyHtml(url, cookieHeader = '', redirectsLeft = 5) {
  const headers = {
    Accept: 'text/html,application/xhtml+xml',
    ...(cookieHeader ? { Cookie: cookieHeader } : {})
  };

  const response = await legacyFetch(url, {
    method: 'GET',
    headers,
    redirect: 'manual'
  });

  let updatedCookieHeader = mergeCookieHeader(cookieHeader, getSetCookieHeaders(response.headers));
  const location = response.headers.get('location');
  const isRedirect = response.status >= 300 && response.status < 400 && location;

  if (isRedirect && redirectsLeft > 0) {
    const nextUrl = new URL(location, url).toString();
    return fetchLegacyHtml(nextUrl, updatedCookieHeader, redirectsLeft - 1);
  }

  const html = await response.text();
  return {
    html,
    status: response.status,
    cookieHeader: updatedCookieHeader
  };
}

function parseLegacyServiceType(value) {
  const raw = normalize(value).toLowerCase();
  if (!raw) return '';

  if (['wed', 'wednesday'].includes(raw)) return 'wed';
  if (['sun_am', 'sun am', 'sunday am', 'sunday_am'].includes(raw)) return 'sun_am';
  if (['sun_pm', 'sun pm', 'sunday pm', 'sunday_pm'].includes(raw)) return 'sun_pm';

  if (raw.includes('wed')) return 'wed';
  if (raw.includes('sun') && raw.includes('pm')) return 'sun_pm';
  if (raw.includes('sun') && raw.includes('am')) return 'sun_am';

  return '';
}

function addLegacyRow(rowsMap, row) {
  const serviceDate = toIsoDate(row.serviceDate);
  const serviceType = parseLegacyServiceType(row.serviceType);
  const headcount = Number.parseInt(row.headcount, 10);
  const note = normalize(row.note);

  if (!serviceDate || !serviceType || Number.isNaN(headcount) || headcount < 0) {
    return;
  }

  const key = `${serviceDate}|${serviceType}`;
  const existing = rowsMap.get(key);
  if (existing) {
    existing.headcount = headcount;
    if (note) {
      existing.note = note;
    }
    return;
  }

  rowsMap.set(key, {
    serviceDate,
    serviceType,
    headcount,
    note
  });
}

function parseLegacyAttendancePage(html) {
  const $ = cheerio.load(html || '');
  const rowsMap = new Map();
  const years = $('#year option')
    .map((index, option) => {
      const raw = normalize($(option).attr('value') || $(option).text());
      return Number.parseInt(raw, 10);
    })
    .get()
    .filter((value) => Number.isInteger(value));

  const chartScript = $('script')
    .map((index, script) => $(script).html() || '')
    .get()
    .find((content) => content.includes('const chartData ='));

  if (chartScript) {
    const chartMatch = chartScript.match(/const\s+chartData\s*=\s*(\{[\s\S]*?\})\s*;/);
    if (chartMatch) {
      try {
        const parsed = JSON.parse(chartMatch[1]);
        ['wed', 'sun_am', 'sun_pm'].forEach((serviceType) => {
          const points = Array.isArray(parsed[serviceType]) ? parsed[serviceType] : [];
          points.forEach((point) => {
            addLegacyRow(rowsMap, {
              serviceDate: point.x,
              serviceType,
              headcount: point.y,
              note: ''
            });
          });
        });
      } catch (error) {
        // Ignore chart parse errors and continue with table parsing.
      }
    }
  }

  $('table tbody tr').each((index, row) => {
    const cells = $(row).find('td');
    if (cells.length < 3) return;

    const badgeClass = normalize($(cells[1]).find('.badge').attr('class') || '');
    const badgeTypeMatch = badgeClass.match(/badge-(wed|sun_am|sun_pm)/);
    const fallbackType = normalize($(cells[1]).text());
    const serviceType = badgeTypeMatch ? badgeTypeMatch[1] : parseLegacyServiceType(fallbackType);

    addLegacyRow(rowsMap, {
      serviceDate: normalize($(cells[0]).text()),
      serviceType,
      headcount: normalize($(cells[2]).text()).replace(/[^\d-]/g, ''),
      note: cells.length > 3 ? normalize($(cells[3]).text()) : ''
    });
  });

  const hasLoginForm = $('form[action="/login"], form[action*="login"]').length > 0;
  const hasPasswordError = /incorrect password/i.test($.text());
  const requiresLogin = hasLoginForm || hasPasswordError;

  return {
    rows: Array.from(rowsMap.values()),
    years: Array.from(new Set(years)),
    requiresLogin,
    hasPasswordError
  };
}

function mergeLegacyRows(targetMap, rows) {
  (rows || []).forEach((row) => addLegacyRow(targetMap, row));
}

async function readLegacyAttendanceRowsByScrape() {
  const legacyUrl = getLegacyAppUrl();
  if (!legacyUrl) {
    return {
      ok: false,
      reason: 'Set METRICS_APP_URL (or ATTENDANCE_APP_URL) before importing legacy data.',
      rows: []
    };
  }

  let rootUrl;
  let loginUrl;
  try {
    rootUrl = new URL('/', legacyUrl).toString();
    loginUrl = new URL('/login', legacyUrl).toString();
  } catch (error) {
    return { ok: false, reason: 'Legacy metrics URL is invalid.', rows: [] };
  }

  try {
    let cookieHeader = '';
    let dashboard = await fetchLegacyHtml(rootUrl, cookieHeader);
    cookieHeader = dashboard.cookieHeader;

    let parsed = parseLegacyAttendancePage(dashboard.html);
    if (!parsed.rows.length || parsed.requiresLogin) {
      const legacyPassword = getLegacyAppPassword();
      if (!legacyPassword) {
        return {
          ok: false,
          reason: 'Legacy app needs login. Set LEGACY_APP_PASSWORD (or reuse ADMIN_PASSWORD) and try again.',
          rows: []
        };
      }

      const loginResponse = await legacyFetch(loginUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          Accept: 'text/html,application/xhtml+xml',
          ...(cookieHeader ? { Cookie: cookieHeader } : {})
        },
        body: new URLSearchParams({ password: legacyPassword }).toString(),
        redirect: 'manual'
      });

      cookieHeader = mergeCookieHeader(cookieHeader, getSetCookieHeaders(loginResponse.headers));
      const redirectTo = loginResponse.headers.get('location');

      if (redirectTo) {
        const afterLoginUrl = new URL(redirectTo, loginUrl).toString();
        dashboard = await fetchLegacyHtml(afterLoginUrl, cookieHeader);
      } else {
        const postLoginHtml = await loginResponse.text();
        dashboard = {
          html: postLoginHtml,
          status: loginResponse.status,
          cookieHeader
        };
      }

      cookieHeader = dashboard.cookieHeader;
      parsed = parseLegacyAttendancePage(dashboard.html);
      if (parsed.hasPasswordError || (parsed.requiresLogin && !parsed.rows.length)) {
        return {
          ok: false,
          reason: 'Legacy login failed. Check LEGACY_APP_PASSWORD (or ADMIN_PASSWORD) value.',
          rows: []
        };
      }
    }

    const mergedRows = new Map();
    mergeLegacyRows(mergedRows, parsed.rows);

    for (const year of parsed.years) {
      const yearUrl = new URL(rootUrl);
      yearUrl.searchParams.set('year', String(year));

      const yearPage = await fetchLegacyHtml(yearUrl.toString(), cookieHeader);
      cookieHeader = yearPage.cookieHeader;

      const yearParsed = parseLegacyAttendancePage(yearPage.html);
      if (yearParsed.requiresLogin && !yearParsed.rows.length) {
        continue;
      }
      mergeLegacyRows(mergedRows, yearParsed.rows);
    }

    const rows = Array.from(mergedRows.values()).sort((a, b) => {
      if (a.serviceDate === b.serviceDate) {
        return a.serviceType.localeCompare(b.serviceType);
      }
      return a.serviceDate.localeCompare(b.serviceDate);
    });

    if (!rows.length) {
      return {
        ok: false,
        reason: 'No attendance rows were found on the legacy metrics pages.',
        rows: []
      };
    }

    return { ok: true, reason: '', rows };
  } catch (error) {
    return {
      ok: false,
      reason: `Could not scrape legacy app: ${error.message}`,
      rows: []
    };
  }
}

const membershipTypes = ['Prospect', 'Member', 'Voting Member'];
const genderOptions = ['Male', 'Female', 'Unspecified'];
const serviceTypes = ['wed', 'sun_am', 'sun_pm'];
const serviceTypeLabels = {
  wed: 'Wednesday',
  sun_am: 'Sunday AM',
  sun_pm: 'Sunday PM'
};

app.locals.formatDateDMY = formatDateDMY;
app.locals.membershipTypes = membershipTypes;
app.locals.genderOptions = genderOptions;
app.locals.serviceTypeLabels = serviceTypeLabels;

function isAuthEnabled() {
  return Boolean(getAuthPassword());
}

function getAuthPassword() {
  const candidates = [
    process.env.CRM_ADMIN_PASSWORD,
    process.env.ADMIN_PASSWORD,
    process.env.APP_PASSWORD,
    process.env.PASSWORD
  ];

  const first = candidates.find((value) => normalize(value).length > 0);
  return normalize(first);
}

function authPasswordSource() {
  if (normalize(process.env.CRM_ADMIN_PASSWORD)) return 'CRM_ADMIN_PASSWORD';
  if (normalize(process.env.ADMIN_PASSWORD)) return 'ADMIN_PASSWORD';
  if (normalize(process.env.APP_PASSWORD)) return 'APP_PASSWORD';
  if (normalize(process.env.PASSWORD)) return 'PASSWORD';
  return '';
}

function safePasswordCompare(input, expected) {
  const a = Buffer.from(input || '', 'utf-8');
  const b = Buffer.from(expected || '', 'utf-8');
  if (a.length !== b.length) return false;
  return crypto.timingSafeEqual(a, b);
}

function isPublicPath(pathname) {
  return (
    pathname === '/login' ||
    pathname === '/auth-status' ||
    pathname.startsWith('/static/') ||
    pathname.startsWith('/register/')
  );
}

app.use((req, res, next) => {
  res.locals.authEnabled = isAuthEnabled();
  res.locals.isAuthenticated = Boolean(req.session?.isAuthenticated);
  next();
});

app.use((req, res, next) => {
  if (!isAuthEnabled() || isPublicPath(req.path) || req.session?.isAuthenticated) {
    return next();
  }

  const returnTo = encodeURIComponent(req.originalUrl || '/people');
  return res.redirect(`/login?returnTo=${returnTo}`);
});

function mapImportRow(row) {
  const joinedAt = toIsoDate(csvField(row, ['joined_at', 'joined at']));
  const createdAt = toIsoDate(csvField(row, ['created_at', 'created at']));
  const baptismDate = toIsoDate(csvField(row, ['baptism_date', 'baptism date']));
  const membershipType = csvField(row, ['membership_type', 'membership type']);
  const status = csvField(row, ['status']);
  const totalContribution = csvField(row, ['total_contribution', 'total contribution']);
  const baseNotes = csvField(row, ['notes', 'note', 'comments', 'comment']);

  const metaNotes = [
    membershipType ? `Membership Type: ${membershipType}` : '',
    status ? `Status: ${status}` : '',
    baptismDate ? `Baptism Date: ${baptismDate}` : '',
    joinedAt ? `Joined At: ${joinedAt}` : '',
    createdAt ? `Created At: ${createdAt}` : '',
    totalContribution ? `Total Contribution: ${totalContribution}` : ''
  ].filter(Boolean);

  return {
    name:
      csvField(row, ['name', 'full_name', 'full name', 'person', 'display_name']) ||
      [csvField(row, ['first_name', 'first name']), csvField(row, ['last_name', 'last name'])]
        .filter(Boolean)
        .join(' ')
        .trim(),
    phone: csvField(row, ['phone', 'mobile', 'phone_number', 'phone number', 'primary_telephone', 'primary telephone']),
    email: csvField(row, ['email', 'e_mail', 'email_address', 'email address']),
    birthday: toIsoDate(csvField(row, ['birthday', 'birthdate', 'dob', 'date_of_birth'])),
    sectionId: csvField(row, ['section', 'section_id', 'zone', 'area']),
    notes: [baseNotes, metaNotes.join(' | ')].filter(Boolean).join(' | '),
    gender: csvField(row, ['gender']),
    maritalStatus: csvField(row, ['marital_status', 'marital status']),
    address: csvField(row, ['address', 'primary_address', 'primary address']),
    city: csvField(row, ['city']),
    state: csvField(row, ['state']),
    zipCode: csvField(row, ['zip', 'zip_code', 'zip code']),
    membershipType,
    joinedAt,
    createdAt,
    baptismDate,
    totalContribution
  };
}

async function importPeopleRows(rows) {
  const candidates = rows.map(mapImportRow);
  let imported = 0;
  let skipped = 0;

  await updateData((data) => {
    const existingKeys = new Set(
      data.people.map((person) => {
        const keyName = normalize(person.name).toLowerCase();
        const keyPhone = normalizePhone(person.phone);
        const keyEmail = normalize(person.email).toLowerCase();
        return `${keyName}|${keyPhone}|${keyEmail}`;
      })
    );

    candidates.forEach((row) => {
      if (!row.name) {
        skipped += 1;
        return;
      }

      const key = `${row.name.toLowerCase()}|${normalizePhone(row.phone)}|${row.email.toLowerCase()}`;
      if (existingKeys.has(key)) {
        skipped += 1;
        return;
      }

      existingKeys.add(key);
      imported += 1;
      data.people.push({
        id: id(),
        name: row.name,
        phone: row.phone,
        email: row.email,
        birthday: row.birthday,
        sectionId: row.sectionId,
        notes: row.notes,
        gender: row.gender || '',
        ageGroup: '',
        occupation: '',
        language: '',
        maritalStatus: row.maritalStatus || '',
        allergies: '',
        emergencyContact: '',
        medicalNotes: '',
        address: row.address || '',
        city: row.city || '',
        state: row.state || '',
        zipCode: row.zipCode || '',
        membershipType: row.membershipType || '',
        joinedAt: row.joinedAt || '',
        createdAt: row.createdAt || '',
        baptismDate: row.baptismDate || '',
        totalContribution: row.totalContribution || '',
        updatedAt: new Date().toISOString()
      });
    });

    return data;
  });

  return { imported, skipped };
}

app.get('/login', (req, res) => {
  if (!isAuthEnabled()) {
    return res.redirect('/people');
  }

  if (req.session?.isAuthenticated) {
    return res.redirect('/people');
  }

  return res.render('login', {
    returnTo: normalize(req.query.returnTo) || '/people',
    error: ''
  });
});

app.post('/login', (req, res) => {
  if (!isAuthEnabled()) {
    return res.redirect('/people');
  }

  const submitted = normalize(req.body.password);
  const expected = getAuthPassword();
  const returnTo = normalize(req.body.returnTo) || '/people';

  if (!safePasswordCompare(submitted, expected)) {
    return res.status(401).render('login', {
      returnTo,
      error: 'Invalid password'
    });
  }

  req.session.isAuthenticated = true;
  return res.redirect(returnTo);
});

app.get('/auth-status', (req, res) => {
  res.json({
    authEnabled: isAuthEnabled(),
    authPasswordSource: authPasswordSource() || null,
    hasAppSecret: Boolean(normalize(process.env.APP_SECRET)),
    isAuthenticated: Boolean(req.session?.isAuthenticated)
  });
});

app.post('/logout', (req, res) => {
  req.session.destroy(() => {
    res.redirect('/login');
  });
});

app.get('/', (req, res) => {
  res.redirect('/people');
});

app.get('/people', async (req, res, next) => {
  try {
    const data = await readData();
    const q = normalize(req.query.q).toLowerCase();
    const followups = normalize(req.query.followups) || 'all';
    const now = new Date();

    const openFollowUpsByPerson = data.followUps
      .filter((item) => item.status !== 'completed')
      .reduce((acc, item) => {
        acc[item.personId] = (acc[item.personId] || 0) + 1;
        return acc;
      }, {});

    let people = sortByName(data.people).map((person) => ({
      ...enrichPerson(person),
      openFollowUps: openFollowUpsByPerson[person.id] || 0
    }));

    if (q) {
      people = people.filter((person) =>
        [
          person.name,
          person.email,
          person.phone,
          person.sectionId,
          person.membershipType,
          person.notes
        ]
          .join(' ')
          .toLowerCase()
          .includes(q)
      );
    }

    if (followups === 'open') {
      people = people.filter((person) => person.openFollowUps > 0);
    }

    const upcomingBirthdays = sortByName(data.people)
      .map((person) => {
        const nextBirthday = nextBirthdayDate(person.birthday, now);
        if (!nextBirthday) return null;
        return {
          id: person.id,
          name: person.name,
          nextBirthdayIso: nextBirthday.toISOString().slice(0, 10),
          daysUntil: diffDays(nextBirthday, now)
        };
      })
      .filter(Boolean)
      .sort((a, b) => a.daysUntil - b.daysUntil)
      .slice(0, 8);

    res.render('people', {
      activeTab: 'people',
      people,
      q,
      followups,
      upcomingBirthdays
    });
  } catch (err) {
    next(err);
  }
});

app.get('/people/new', async (req, res, next) => {
  try {
    res.render('people-new', {
      activeTab: 'people_new'
    });
  } catch (err) {
    next(err);
  }
});

app.get('/people/:id', async (req, res, next) => {
  try {
    const data = await readData();
    const person = data.people.find((entry) => entry.id === req.params.id);

    if (!person) {
      return res.status(404).send('Person not found');
    }

    const tab = req.query.tab || 'profile';

    const followUps = data.followUps
      .filter((entry) => entry.personId === person.id)
      .sort((a, b) => new Date(a.dueDate || a.createdAt) - new Date(b.dueDate || b.createdAt));

    const visits = data.visits
      .filter((entry) => entry.personId === person.id)
      .sort((a, b) => new Date(b.date) - new Date(a.date));

    res.render('people-profile', {
      activeTab: 'people',
      person: enrichPerson(person),
      tab,
      followUps,
      visits,
      followUpOpenCount: followUps.filter((entry) => entry.status !== 'completed').length
    });
  } catch (err) {
    next(err);
  }
});

app.post('/people', async (req, res, next) => {
  try {
    const person = {
      id: id(),
      name: req.body.name?.trim() || '',
      phone: req.body.phone?.trim() || '',
      email: req.body.email?.trim() || '',
      birthday: toIsoDate(req.body.birthday),
      sectionId: req.body.sectionId || '',
      notes: req.body.notes?.trim() || '',
      gender: req.body.gender || '',
      ageGroup: req.body.ageGroup || '',
      membershipType: req.body.membershipType || '',
      occupation: req.body.occupation || '',
      language: req.body.language || '',
      maritalStatus: req.body.maritalStatus || '',
      allergies: req.body.allergies || '',
      emergencyContact: req.body.emergencyContact || '',
      medicalNotes: req.body.medicalNotes || '',
      address: req.body.address || '',
      city: req.body.city || '',
      state: req.body.state || '',
      zipCode: req.body.zipCode || '',
      updatedAt: new Date().toISOString()
    };

    if (!person.name) {
      return res.redirect('/people/new');
    }

    await updateData((data) => {
      data.people.push(person);
      return data;
    });

    res.redirect(`/people/${person.id}`);
  } catch (err) {
    next(err);
  }
});

app.post('/people/:id', async (req, res, next) => {
  try {
    const returnTo = req.body.returnTo || `/people/${req.params.id}`;

    await updateData((data) => {
      const person = data.people.find((entry) => entry.id === req.params.id);

      if (person) {
        const hasField = (field) => Object.prototype.hasOwnProperty.call(req.body, field);
        const write = (field, mapper = (value) => value) => {
          if (hasField(field)) {
            person[field] = mapper(req.body[field]);
          }
        };

        write('name', (value) => value?.trim() || person.name);
        write('phone', (value) => value?.trim() || '');
        write('email', (value) => value?.trim() || '');
        write('birthday', (value) => toIsoDate(value));
        write('sectionId', (value) => value || '');
        write('notes', (value) => value?.trim() || '');
        write('gender', (value) => value || '');
        write('ageGroup', (value) => value || '');
        write('membershipType', (value) => value || '');
        write('occupation', (value) => value || '');
        write('language', (value) => value || '');
        write('maritalStatus', (value) => value || '');
        write('allergies', (value) => value || '');
        write('emergencyContact', (value) => value || '');
        write('medicalNotes', (value) => value || '');
        write('address', (value) => value || '');
        write('city', (value) => value || '');
        write('state', (value) => value || '');
        write('zipCode', (value) => value || '');
        person.updatedAt = new Date().toISOString();
      }

      return data;
    });

    res.redirect(returnTo);
  } catch (err) {
    next(err);
  }
});

app.post('/people/:id/delete', async (req, res, next) => {
  try {
    await updateData((data) => {
      data.people = data.people.filter((entry) => entry.id !== req.params.id);
      data.followUps = data.followUps.filter((entry) => entry.personId !== req.params.id);
      data.visits = data.visits.filter((entry) => entry.personId !== req.params.id);
      return data;
    });

    res.redirect('/people');
  } catch (err) {
    next(err);
  }
});

app.post('/people/:id/followups', async (req, res, next) => {
  try {
    const title = req.body.title?.trim();

    if (!title) {
      return res.redirect(`/people/${req.params.id}?tab=followups`);
    }

    await updateData((data) => {
      data.followUps.push({
        id: id(),
        personId: req.params.id,
        title,
        dueDate: req.body.dueDate || '',
        notes: req.body.notes?.trim() || '',
        status: 'open',
        createdAt: new Date().toISOString(),
        completedAt: ''
      });

      return data;
    });

    res.redirect(`/people/${req.params.id}?tab=followups`);
  } catch (err) {
    next(err);
  }
});

app.post('/people/:id/followups/:followUpId/complete', async (req, res, next) => {
  try {
    await updateData((data) => {
      const row = data.followUps.find((entry) => entry.id === req.params.followUpId && entry.personId === req.params.id);

      if (row) {
        row.status = 'completed';
        row.completedAt = new Date().toISOString();
      }

      return data;
    });

    res.redirect(`/people/${req.params.id}?tab=followups`);
  } catch (err) {
    next(err);
  }
});

app.post('/people/:id/followups/:followUpId/reopen', async (req, res, next) => {
  try {
    await updateData((data) => {
      const row = data.followUps.find((entry) => entry.id === req.params.followUpId && entry.personId === req.params.id);

      if (row) {
        row.status = 'open';
        row.completedAt = '';
      }

      return data;
    });

    res.redirect(`/people/${req.params.id}?tab=followups`);
  } catch (err) {
    next(err);
  }
});

app.post('/people/:id/visits', async (req, res, next) => {
  try {
    const summary = req.body.summary?.trim();

    if (!summary) {
      return res.redirect(`/people/${req.params.id}?tab=visits`);
    }

    await updateData((data) => {
      data.visits.push({
        id: id(),
        personId: req.params.id,
        date: req.body.date || new Date().toISOString().slice(0, 10),
        summary,
        nextStep: req.body.nextStep?.trim() || '',
        createdAt: new Date().toISOString()
      });

      return data;
    });

    res.redirect(`/people/${req.params.id}?tab=visits`);
  } catch (err) {
    next(err);
  }
});

app.get('/followups', async (req, res, next) => {
  try {
    const data = await readData();
    const status = normalize(req.query.status) || 'open';

    const peopleById = data.people.reduce((acc, person) => {
      acc[person.id] = person;
      return acc;
    }, {});

    let queue = data.followUps
      .map((item) => {
        const person = peopleById[item.personId];
        return {
          ...item,
          personName: person?.name || 'Unknown person',
          personLink: person ? `/people/${person.id}?tab=followups` : '/people'
        };
      })
      .sort((a, b) => {
        const aDate = a.dueDate || a.createdAt;
        const bDate = b.dueDate || b.createdAt;
        return new Date(aDate) - new Date(bDate);
      });

    if (status === 'open') {
      queue = queue.filter((item) => item.status !== 'completed');
    } else if (status === 'completed') {
      queue = queue.filter((item) => item.status === 'completed');
    }

    res.render('followups', {
      activeTab: 'followups',
      queue,
      status
    });
  } catch (err) {
    next(err);
  }
});

app.post('/followups/:followUpId/complete', async (req, res, next) => {
  try {
    await updateData((data) => {
      const row = data.followUps.find((entry) => entry.id === req.params.followUpId);
      if (row) {
        row.status = 'completed';
        row.completedAt = new Date().toISOString();
      }
      return data;
    });

    res.redirect('/followups?status=open');
  } catch (err) {
    next(err);
  }
});

app.post('/followups/:followUpId/reopen', async (req, res, next) => {
  try {
    await updateData((data) => {
      const row = data.followUps.find((entry) => entry.id === req.params.followUpId);
      if (row) {
        row.status = 'open';
        row.completedAt = '';
      }
      return data;
    });

    res.redirect('/followups?status=completed');
  } catch (err) {
    next(err);
  }
});

app.get('/metrics', async (req, res, next) => {
  try {
    const data = await readData();
    const now = new Date();
    const currentYear = now.getFullYear();
    const selectedYear = Number.parseInt(req.query.year, 10);
    const year = Number.isNaN(selectedYear) ? currentYear : selectedYear;
    const totalPeople = data.people.length;
    const metricsAppUrl =
      process.env.METRICS_APP_URL ||
      process.env.ATTENDANCE_APP_URL ||
      data.settings.attendanceAppUrl;

    const birthdaysThisMonth = data.people.filter((person) => {
      if (!person.birthday) return false;
      return new Date(person.birthday).getMonth() === now.getMonth();
    }).length;

    const visitsThisMonth = data.visits.filter((entry) => {
      const date = new Date(entry.date);
      return date.getMonth() === now.getMonth() && date.getFullYear() === now.getFullYear();
    }).length;

    const openFollowUps = data.followUps.filter((entry) => entry.status !== 'completed').length;
    const normalizedAttendance = (data.attendanceRecords || [])
      .map((record) => ({
        id: record.id,
        serviceDate: toIsoDate(record.serviceDate),
        serviceType: normalizeServiceType(record.serviceType),
        headcount: Number(record.headcount) || 0,
        note: normalize(record.note),
        createdAt: record.createdAt || '',
        updatedAt: record.updatedAt || ''
      }))
      .filter((record) => record.serviceDate);

    const yearRecords = normalizedAttendance
      .filter((record) => new Date(record.serviceDate).getFullYear() === year)
      .sort((a, b) => {
        if (a.serviceDate === b.serviceDate) {
          return a.serviceType.localeCompare(b.serviceType);
        }
        return a.serviceDate.localeCompare(b.serviceDate);
      });

    const years = Array.from(
      new Set(normalizedAttendance.map((record) => new Date(record.serviceDate).getFullYear()))
    ).sort((a, b) => b - a);
    if (!years.includes(year)) {
      years.unshift(year);
    }

    const recent = [...normalizedAttendance]
      .sort((a, b) => {
        if (a.serviceDate === b.serviceDate) {
          return a.serviceType.localeCompare(b.serviceType);
        }
        return b.serviceDate.localeCompare(a.serviceDate);
      })
      .slice(0, 15);

    const avgBuckets = serviceTypes.reduce((acc, type) => {
      const rows = yearRecords.filter((record) => record.serviceType === type);
      const total = rows.reduce((sum, row) => sum + row.headcount, 0);
      acc[type] = {
        avg: rows.length ? Math.round(total / rows.length) : null,
        count: rows.length
      };
      return acc;
    }, {});

    const chartData = {
      wed: yearRecords
        .filter((record) => record.serviceType === 'wed')
        .map((record) => ({ x: record.serviceDate, y: record.headcount })),
      sun_am: yearRecords
        .filter((record) => record.serviceType === 'sun_am')
        .map((record) => ({ x: record.serviceDate, y: record.headcount })),
      sun_pm: yearRecords
        .filter((record) => record.serviceType === 'sun_pm')
        .map((record) => ({ x: record.serviceDate, y: record.headcount }))
    };

    const attendanceThisYear = yearRecords.length;
    const attendanceTotalHeadcount = yearRecords.reduce((sum, row) => sum + row.headcount, 0);
    const legacyMessage = normalize(req.query.legacyMessage);
    const legacyImported = normalize(req.query.legacyImported);
    const legacyUpdated = normalize(req.query.legacyUpdated);
    const legacySkipped = normalize(req.query.legacySkipped);

    res.render('metrics', {
      activeTab: 'metrics',
      metricsAppUrl,
      year,
      years,
      chartData,
      recent,
      averages: avgBuckets,
      attendanceThisYear,
      attendanceTotalHeadcount,
      legacyMessage,
      legacyImported,
      legacyUpdated,
      legacySkipped,
      totalPeople,
      birthdaysThisMonth,
      upcomingEvents: data.events.length,
      visitsThisMonth,
      openFollowUps
    });
  } catch (err) {
    next(err);
  }
});

app.get('/metrics/entry', async (req, res, next) => {
  try {
    const today = new Date().toISOString().slice(0, 10);
    const dayOfWeek = new Date().getDay();

    let suggestedType = 'sun_am';
    if (dayOfWeek === 3) suggestedType = 'wed';
    else if (dayOfWeek === 0) suggestedType = 'sun_am';

    res.render('metrics-entry', {
      activeTab: 'metrics',
      record: null,
      date: toIsoDate(req.query.date || today) || today,
      type: normalizeServiceType(req.query.type || suggestedType),
      error: ''
    });
  } catch (err) {
    next(err);
  }
});

app.get('/metrics/entry/:id/edit', async (req, res, next) => {
  try {
    const data = await readData();
    const record = (data.attendanceRecords || []).find((entry) => entry.id === req.params.id);
    if (!record) {
      return res.redirect('/metrics');
    }

    res.render('metrics-entry', {
      activeTab: 'metrics',
      record,
      date: toIsoDate(record.serviceDate),
      type: normalizeServiceType(record.serviceType),
      error: ''
    });
  } catch (err) {
    next(err);
  }
});

app.post('/metrics/entry', async (req, res, next) => {
  try {
    const serviceDate = toIsoDate(req.body.service_date);
    const serviceType = normalizeServiceType(req.body.service_type);
    const headcount = Number.parseInt(req.body.headcount, 10);
    const note = normalize(req.body.note);

    if (!serviceDate || Number.isNaN(headcount) || headcount < 0) {
      return res.status(400).render('metrics-entry', {
        activeTab: 'metrics',
        record: null,
        date: serviceDate || normalize(req.body.service_date),
        type: serviceType,
        error: 'Date, service type, and a non-negative headcount are required.'
      });
    }

    await updateData((data) => {
      const existing = (data.attendanceRecords || []).find(
        (row) => toIsoDate(row.serviceDate) === serviceDate && normalizeServiceType(row.serviceType) === serviceType
      );

      if (existing) {
        existing.headcount = headcount;
        existing.note = note;
        existing.updatedAt = new Date().toISOString();
      } else {
        data.attendanceRecords = data.attendanceRecords || [];
        data.attendanceRecords.push({
          id: id(),
          serviceDate,
          serviceType,
          headcount,
          note,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        });
      }
      return data;
    });

    return res.redirect(`/metrics?year=${new Date(serviceDate).getFullYear()}`);
  } catch (err) {
    next(err);
  }
});

app.post('/metrics/entry/:id/update', async (req, res, next) => {
  try {
    const serviceDate = toIsoDate(req.body.service_date);
    const serviceType = normalizeServiceType(req.body.service_type);
    const headcount = Number.parseInt(req.body.headcount, 10);
    const note = normalize(req.body.note);

    if (!serviceDate || Number.isNaN(headcount) || headcount < 0) {
      return res.status(400).render('metrics-entry', {
        activeTab: 'metrics',
        record: { id: req.params.id, serviceDate, serviceType, headcount: req.body.headcount, note },
        date: serviceDate || normalize(req.body.service_date),
        type: serviceType,
        error: 'Date, service type, and a non-negative headcount are required.'
      });
    }

    await updateData((data) => {
      const row = (data.attendanceRecords || []).find((entry) => entry.id === req.params.id);
      if (row) {
        row.serviceDate = serviceDate;
        row.serviceType = serviceType;
        row.headcount = headcount;
        row.note = note;
        row.updatedAt = new Date().toISOString();
      }
      return data;
    });

    return res.redirect(`/metrics?year=${new Date(serviceDate).getFullYear()}`);
  } catch (err) {
    next(err);
  }
});

app.post('/api/metrics/attendance/:id/delete', async (req, res, next) => {
  try {
    await updateData((data) => {
      data.attendanceRecords = (data.attendanceRecords || []).filter((entry) => entry.id !== req.params.id);
      return data;
    });
    res.json({ success: true });
  } catch (err) {
    next(err);
  }
});

app.post('/metrics/import-legacy', async (req, res, next) => {
  try {
    const legacy = await readLegacyAttendanceRowsByScrape();
    if (!legacy.ok) {
      const params = new URLSearchParams({
        legacyMessage: legacy.reason,
        legacyImported: '0',
        legacyUpdated: '0',
        legacySkipped: '0'
      });
      return res.redirect(`/metrics?${params.toString()}`);
    }

    let imported = 0;
    let updated = 0;
    let skipped = 0;

    await updateData((data) => {
      data.attendanceRecords = data.attendanceRecords || [];
      const keyMap = new Map(
        data.attendanceRecords.map((entry) => {
          const key = `${toIsoDate(entry.serviceDate)}|${normalizeServiceType(entry.serviceType)}`;
          return [key, entry];
        })
      );

      legacy.rows.forEach((row) => {
        const serviceDate = toIsoDate(row.serviceDate);
        const serviceType = normalizeServiceType(row.serviceType);
        const headcount = Number.parseInt(row.headcount, 10);
        const note = normalize(row.note);

        if (!serviceDate || Number.isNaN(headcount) || headcount < 0) {
          skipped += 1;
          return;
        }

        const key = `${serviceDate}|${serviceType}`;
        const existing = keyMap.get(key);

        if (existing) {
          const changed = existing.headcount !== headcount || normalize(existing.note) !== note;
          if (changed) {
            existing.headcount = headcount;
            existing.note = note;
            existing.updatedAt = new Date().toISOString();
            updated += 1;
          } else {
            skipped += 1;
          }
          return;
        }

        const newRecord = {
          id: row.id ? `legacy-att-${row.id}` : id(),
          serviceDate,
          serviceType,
          headcount,
          note,
          createdAt: row.createdAt ? new Date(row.createdAt).toISOString() : new Date().toISOString(),
          updatedAt: new Date().toISOString()
        };
        data.attendanceRecords.push(newRecord);
        keyMap.set(key, newRecord);
        imported += 1;
      });

      return data;
    });

    const params = new URLSearchParams({
      legacyMessage: 'Legacy attendance import complete.',
      legacyImported: String(imported),
      legacyUpdated: String(updated),
      legacySkipped: String(skipped)
    });
    return res.redirect(`/metrics?${params.toString()}`);
  } catch (err) {
    next(err);
  }
});

app.get('/calendar', async (req, res, next) => {
  try {
    res.render('calendar', {
      activeTab: 'calendar'
    });
  } catch (err) {
    next(err);
  }
});

app.get('/api/events', async (req, res, next) => {
  try {
    const data = await readData();
    res.json(data.events);
  } catch (err) {
    next(err);
  }
});

app.get('/api/birthdays', async (req, res, next) => {
  try {
    const data = await readData();
    const start = new Date(req.query.start || `${new Date().getFullYear()}-01-01`);
    const end = new Date(req.query.end || `${new Date().getFullYear() + 1}-01-01`);

    const startDate = Number.isNaN(start.getTime()) ? new Date(new Date().getFullYear(), 0, 1) : start;
    const endDate = Number.isNaN(end.getTime()) ? new Date(new Date().getFullYear() + 1, 0, 1) : end;

    const startYear = startDate.getFullYear() - 1;
    const endYear = endDate.getFullYear() + 1;
    const events = [];

    data.people.forEach((person) => {
      const iso = toIsoDate(person.birthday);
      if (!iso) return;

      const [, monthStr, dayStr] = iso.split('-');
      const month = Number(monthStr);
      const day = Number(dayStr);

      for (let year = startYear; year <= endYear; year += 1) {
        let birthdayDate = new Date(year, month - 1, day);

        if (birthdayDate.getMonth() !== month - 1 || birthdayDate.getDate() !== day) {
          if (month === 2 && day === 29) {
            birthdayDate = new Date(year, 1, 28);
          } else {
            continue;
          }
        }

        if (birthdayDate < startDate || birthdayDate >= endDate) {
          continue;
        }

        events.push({
          id: `birthday-${person.id}-${year}`,
          title: `${person.name} Birthday`,
          start: birthdayDate.toISOString().slice(0, 10),
          allDay: true,
          color: '#f59e0b',
          extendedProps: {
            sourceType: 'birthday',
            personId: person.id
          }
        });
      }
    });

    res.json(events);
  } catch (err) {
    next(err);
  }
});

app.post('/api/events', async (req, res, next) => {
  try {
    const payload = {
      id: req.body.id || id(),
      title: req.body.title?.trim() || '',
      start: req.body.start,
      end: req.body.end || null,
      description: req.body.description?.trim() || ''
    };

    if (!payload.title || !payload.start) {
      return res.status(400).json({ error: 'title and start are required' });
    }

    await updateData((data) => {
      const idx = data.events.findIndex((entry) => entry.id === payload.id);
      if (idx === -1) {
        data.events.push(payload);
      } else {
        data.events[idx] = payload;
      }
      return data;
    });

    res.json({ ok: true, event: payload });
  } catch (err) {
    next(err);
  }
});

app.delete('/api/events/:id', async (req, res, next) => {
  try {
    await updateData((data) => {
      data.events = data.events.filter((entry) => entry.id !== req.params.id);
      return data;
    });

    res.json({ ok: true });
  } catch (err) {
    next(err);
  }
});

app.get('/forms', async (req, res, next) => {
  try {
    const data = await readData();

    res.render('forms', {
      activeTab: 'forms',
      forms: data.forms,
      submissions: data.formSubmissions
    });
  } catch (err) {
    next(err);
  }
});

app.post('/forms', async (req, res, next) => {
  try {
    const name = req.body.name?.trim();
    const description = req.body.description?.trim() || '';

    if (!name) {
      return res.redirect('/forms');
    }

    const fields = (req.body.fields || '')
      .split('\n')
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => {
        const [label, type = 'text', required = 'yes'] = line.split('|').map((value) => value.trim());
        return {
          id: id(),
          label,
          type: type || 'text',
          required: required.toLowerCase() === 'yes'
        };
      })
      .filter((field) => field.label);

    await updateData((data) => {
      data.forms.push({
        id: id(),
        name,
        description,
        fields,
        createdAt: new Date().toISOString()
      });
      return data;
    });

    res.redirect('/forms');
  } catch (err) {
    next(err);
  }
});

app.post('/forms/:id/delete', async (req, res, next) => {
  try {
    await updateData((data) => {
      data.forms = data.forms.filter((entry) => entry.id !== req.params.id);
      data.formSubmissions = data.formSubmissions.filter((entry) => entry.formId !== req.params.id);
      return data;
    });

    res.redirect('/forms');
  } catch (err) {
    next(err);
  }
});

app.get('/forms/:id/submissions', async (req, res, next) => {
  try {
    const data = await readData();
    const form = data.forms.find((entry) => entry.id === req.params.id);

    if (!form) {
      return res.status(404).send('Form not found');
    }

    const rows = data.formSubmissions.filter((entry) => entry.formId === form.id);

    res.render('form-submissions', {
      activeTab: 'forms',
      form,
      rows
    });
  } catch (err) {
    next(err);
  }
});

app.get('/register/:id', async (req, res, next) => {
  try {
    const data = await readData();
    const form = data.forms.find((entry) => entry.id === req.params.id);

    if (!form) {
      return res.status(404).send('Form not found');
    }

    res.render('form-public', { form });
  } catch (err) {
    next(err);
  }
});

app.post('/register/:id', async (req, res, next) => {
  try {
    const data = await readData();
    const form = data.forms.find((entry) => entry.id === req.params.id);

    if (!form) {
      return res.status(404).send('Form not found');
    }

    const answers = {};
    form.fields.forEach((field) => {
      answers[field.label] = req.body[field.id] || '';
    });

    await updateData((latest) => {
      latest.formSubmissions.push({
        id: id(),
        formId: form.id,
        submittedAt: new Date().toISOString(),
        answers
      });

      return latest;
    });

    res.render('form-submitted', { formName: form.name });
  } catch (err) {
    next(err);
  }
});

app.get('/import', async (req, res, next) => {
  try {
    const data = await readData();

    res.render('import', {
      activeTab: 'import',
      totalPeople: data.people.length,
      imported: normalize(req.query.imported),
      skipped: normalize(req.query.skipped),
      message: normalize(req.query.message)
    });
  } catch (err) {
    next(err);
  }
});

app.post('/import/people', async (req, res, next) => {
  try {
    const csvText = normalize(req.body.csvText);

    if (!csvText) {
      return res.redirect('/import?message=Paste+CSV+content+first');
    }

    let rows;
    try {
      const delimiter = detectDelimiter(csvText);
      rows = parse(csvText, {
        columns: true,
        delimiter,
        trim: true,
        skip_empty_lines: true,
        bom: true,
        relax_column_count: true
      });
    } catch {
      return res.redirect('/import?message=CSV+format+could+not+be+parsed');
    }

    if (!rows.length) {
      return res.redirect('/import?message=No+CSV+rows+found');
    }

    const { imported, skipped } = await importPeopleRows(rows);

    const message =
      imported === 0 && skipped > 0
        ? 'No rows imported. Check column headers and delimiter'
        : 'Import complete';
    const params = new URLSearchParams({
      imported: String(imported),
      skipped: String(skipped),
      message
    });

    res.redirect(`/import?${params.toString()}`);
  } catch (err) {
    next(err);
  }
});

app.post('/import/people/file', upload.single('peopleFile'), async (req, res, next) => {
  try {
    if (!req.file) {
      return res.redirect('/import?message=Choose+a+file+first');
    }

    const ext = path.extname(req.file.originalname || '').toLowerCase();
    let rows = [];

    if (ext === '.csv') {
      const csvText = req.file.buffer.toString('utf8');
      const delimiter = detectDelimiter(csvText);

      rows = parse(csvText, {
        columns: true,
        delimiter,
        trim: true,
        skip_empty_lines: true,
        bom: true,
        relax_column_count: true
      });
    } else if (ext === '.xls' || ext === '.xlsx') {
      const workbook = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
      const [firstSheetName] = workbook.SheetNames;

      if (!firstSheetName) {
        return res.redirect('/import?message=No+sheets+found+in+Excel+file');
      }

      const sheet = workbook.Sheets[firstSheetName];
      rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    } else {
      return res.redirect('/import?message=Unsupported+file+type.+Use+CSV,+XLS,+or+XLSX');
    }

    if (!rows.length) {
      return res.redirect('/import?message=No+rows+found+in+uploaded+file');
    }

    const { imported, skipped } = await importPeopleRows(rows);
    const source = ext.replace('.', '').toUpperCase() || 'file';
    const message =
      imported === 0 && skipped > 0
        ? `No rows imported from ${source}. Check column headers`
        : `Import complete from ${source}`;
    const params = new URLSearchParams({
      imported: String(imported),
      skipped: String(skipped),
      message
    });

    res.redirect(`/import?${params.toString()}`);
  } catch (err) {
    next(err);
  }
});

app.get('/visitation', async (req, res, next) => {
  try {
    const data = await readData();

    res.render('visitation', {
      activeTab: 'visitation',
      sections: data.sections
    });
  } catch (err) {
    next(err);
  }
});

app.get('/api/sections', async (req, res, next) => {
  try {
    const data = await readData();
    res.json(data.sections);
  } catch (err) {
    next(err);
  }
});

app.post('/api/sections', async (req, res, next) => {
  try {
    const payload = {
      id: req.body.id || id(),
      name: req.body.name?.trim() || '',
      color: req.body.color || '#0c4a6e',
      lastVisited: req.body.lastVisited || '',
      notes: req.body.notes?.trim() || '',
      geojson: req.body.geojson || null
    };

    if (!payload.name || !payload.geojson) {
      return res.status(400).json({ error: 'name and geojson are required' });
    }

    await updateData((data) => {
      const idx = data.sections.findIndex((entry) => entry.id === payload.id);
      if (idx === -1) {
        data.sections.push(payload);
      } else {
        data.sections[idx] = payload;
      }
      return data;
    });

    res.json({ ok: true, section: payload });
  } catch (err) {
    next(err);
  }
});

app.delete('/api/sections/:id', async (req, res, next) => {
  try {
    await updateData((data) => {
      data.sections = data.sections.filter((entry) => entry.id !== req.params.id);
      return data;
    });

    res.json({ ok: true });
  } catch (err) {
    next(err);
  }
});

app.use((err, req, res, next) => {
  if (req.path === '/import/people/file' && err && err.code === 'LIMIT_FILE_SIZE') {
    return res.redirect('/import?message=File+is+too+large.+Max+size+is+8MB');
  }

  console.error(err);
  res.status(500).send('Something went wrong. Check server logs.');
});

app.listen(PORT, () => {
  console.log(`Church dashboard running on http://localhost:${PORT}`);
});
