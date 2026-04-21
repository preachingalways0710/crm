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
    rolling: true,
    cookie: {
      httpOnly: true,
      sameSite: 'lax',
      secure: 'auto',
      maxAge: 1000 * 60 * 60 * 24 * 30
    }
  })
);
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 8 * 1024 * 1024 }
});
const PEOPLE_DIRECTORY_DEFAULT_PAGE_SIZE = 50;
const PEOPLE_DIRECTORY_PAGE_SIZES = [25, 50, 100];

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

function normalizeCoordinate(value, bounds = {}) {
  const raw = normalize(value).replace(',', '.');
  if (!raw) return '';

  const parsed = Number.parseFloat(raw);
  if (!Number.isFinite(parsed)) return '';

  if (Number.isFinite(bounds.min) && parsed < bounds.min) return '';
  if (Number.isFinite(bounds.max) && parsed > bounds.max) return '';

  return Number.parseFloat(parsed.toFixed(6)).toString();
}

function normalizeLatitude(value) {
  return normalizeCoordinate(value, { min: -90, max: 90 });
}

function normalizeLongitude(value) {
  return normalizeCoordinate(value, { min: -180, max: 180 });
}

function normalizeMapZoom(value, fallback = 16) {
  const parsed = Number.parseInt(normalize(value), 10);
  if (!Number.isInteger(parsed)) return fallback;
  return Math.min(20, Math.max(3, parsed));
}

function mapProfileSummary(person) {
  const addressParts = [person.address, person.city, person.state, person.zipCode]
    .map((part) => normalize(part))
    .filter(Boolean);

  return {
    id: person.id,
    name: normalize(person.name) || 'Unnamed profile',
    lat: normalizeLatitude(person.mapLat),
    lng: normalizeLongitude(person.mapLng),
    address: addressParts.join(', ')
  };
}

function hydrateChurchSettings(settings) {
  const raw = settings && typeof settings === 'object' ? settings : {};

  return {
    name: normalize(raw.name),
    email: normalize(raw.email),
    phone: normalize(raw.phone),
    address: normalize(raw.address),
    city: normalize(raw.city),
    state: normalize(raw.state),
    zipCode: normalize(raw.zipCode),
    mapLat: normalizeLatitude(raw.mapLat || raw.lat),
    mapLng: normalizeLongitude(raw.mapLng || raw.lng)
  };
}

function formatChurchAddress(churchSettings) {
  return [
    normalize(churchSettings?.address),
    normalize(churchSettings?.city),
    normalize(churchSettings?.state),
    normalize(churchSettings?.zipCode)
  ]
    .filter(Boolean)
    .join(', ');
}

function mergeVisitationWithChurchSettings(mapSettings, churchSettings) {
  const churchAddress = formatChurchAddress(churchSettings);

  return {
    mapCenterMode: 'church',
    mapCenterZoom: normalizeMapZoom(mapSettings?.mapCenterZoom, 17),
    profilePersonId: '',
    churchProfile: {
      name: churchSettings.name || normalize(mapSettings?.churchProfile?.name),
      address: churchAddress || normalize(mapSettings?.churchProfile?.address),
      lat: churchSettings.mapLat || normalizeLatitude(mapSettings?.churchProfile?.lat),
      lng: churchSettings.mapLng || normalizeLongitude(mapSettings?.churchProfile?.lng)
    }
  };
}

function hydrateVisitationSettings(settings, people = []) {
  const raw = settings && typeof settings === 'object' ? settings : {};
  const churchProfile = raw.churchProfile && typeof raw.churchProfile === 'object' ? raw.churchProfile : {};

  return {
    mapCenterMode: 'church',
    mapCenterZoom: normalizeMapZoom(raw.mapCenterZoom, 17),
    profilePersonId: '',
    churchProfile: {
      name: normalize(churchProfile.name),
      address: normalize(churchProfile.address),
      lat: normalizeLatitude(churchProfile.lat),
      lng: normalizeLongitude(churchProfile.lng)
    }
  };
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

function parseTagsInput(value) {
  const parts = normalize(value)
    .split(/[,\n;]+/)
    .map((part) => part.trim())
    .filter(Boolean);

  const seen = new Set();
  const tags = [];
  parts.forEach((tag) => {
    const key = tag.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      tags.push(tag);
    }
  });

  return tags.slice(0, 20);
}

function normalizePersonTags(value) {
  if (Array.isArray(value)) {
    return parseTagsInput(value.join(','));
  }
  if (typeof value === 'string') {
    return parseTagsInput(value);
  }
  return [];
}

function normalizePersonRelationIds(value) {
  const list = Array.isArray(value) ? value : typeof value === 'string' ? value.split(',') : [];
  const seen = new Set();
  const ids = [];

  list.forEach((entry) => {
    const item = normalize(entry);
    if (!item || seen.has(item)) return;
    seen.add(item);
    ids.push(item);
  });

  return ids;
}

function ensurePersonRelationships(person) {
  person.spouseIds = normalizePersonRelationIds(person.spouseIds);
  person.parentIds = normalizePersonRelationIds(person.parentIds);
  person.childIds = normalizePersonRelationIds(person.childIds);
  return person;
}

function hydratePersonRelationships(person) {
  return ensurePersonRelationships({ ...(person || {}) });
}

function addUniqueId(list, personId) {
  if (!personId || list.includes(personId)) return false;
  list.push(personId);
  return true;
}

function removeId(list, personId) {
  const before = list.length;
  const next = list.filter((entry) => entry !== personId);
  if (next.length === before) return false;
  list.splice(0, list.length, ...next);
  return true;
}

function normalizeFollowupsFilter(value) {
  return normalize(value) === 'open' ? 'open' : 'all';
}

function normalizeUrlWithScheme(value) {
  const raw = normalize(value);
  if (!raw) return '';
  if (/^https?:\/\//i.test(raw)) return raw;
  return `https://${raw}`;
}

function getPessoasAppUrls() {
  const urls = [];
  const primary = normalize(
    process.env.PESSOAS_APP_URL ||
      process.env.METRICS_APP_URL ||
      process.env.ATTENDANCE_APP_URL ||
      'https://pessoas.meuibbv.com'
  );

  const primaryNormalized = normalizeUrlWithScheme(primary);
  if (primaryNormalized) {
    urls.push(primaryNormalized);
  }

  try {
    const parsed = new URL(primaryNormalized);
    if (parsed.hostname.endsWith('.co')) {
      const corrected = new URL(primaryNormalized);
      corrected.hostname = `${parsed.hostname.slice(0, -3)}.com`;
      urls.push(corrected.toString());
    }
  } catch {
    // Ignore malformed URL here; readPessoasAttendanceRows will return a helpful error.
  }

  return Array.from(new Set(urls));
}

function getPessoasAppPassword() {
  const candidates = [
    process.env.PESSOAS_APP_PASSWORD,
    process.env.LEGACY_APP_PASSWORD,
    process.env.ATTENDANCE_APP_PASSWORD,
    process.env.ADMIN_PASSWORD,
    process.env.CRM_ADMIN_PASSWORD,
    process.env.admin_password
  ];

  const first = candidates.find((value) => normalize(value));
  return normalize(first);
}

function getSetCookie(headers) {
  if (typeof headers.getSetCookie === 'function') {
    return headers.getSetCookie();
  }
  const single = headers.get('set-cookie');
  return single ? [single] : [];
}

function appendCookies(existingHeader, setCookieHeaders) {
  const cookies = new Map();
  normalize(existingHeader)
    .split(';')
    .map((part) => part.trim())
    .filter(Boolean)
    .forEach((entry) => {
      const idx = entry.indexOf('=');
      if (idx > 0) cookies.set(entry.slice(0, idx), entry.slice(idx + 1));
    });

  (setCookieHeaders || []).forEach((header) => {
    const firstPair = normalize(header).split(';')[0];
    const idx = firstPair.indexOf('=');
    if (idx > 0) cookies.set(firstPair.slice(0, idx), firstPair.slice(idx + 1));
  });

  return Array.from(cookies.entries())
    .map(([name, value]) => `${name}=${value}`)
    .join('; ');
}

async function fetchWithTimeout(url, options = {}) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 30000);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(timeout);
  }
}

function stripDiacritics(value) {
  return normalize(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

const brazilStateHints = new Set([
  'ac', 'al', 'am', 'ap', 'ba', 'ce', 'df', 'es', 'go', 'ma', 'mg', 'ms', 'mt', 'pa', 'pb', 'pe', 'pi',
  'pr', 'rj', 'rn', 'ro', 'rr', 'rs', 'sc', 'se', 'sp', 'to',
  'acre', 'alagoas', 'amapa', 'amazonas', 'bahia', 'ceara', 'distrito federal', 'espirito santo', 'goias',
  'maranhao', 'mato grosso', 'mato grosso do sul', 'minas gerais', 'para', 'paraiba', 'parana', 'pernambuco',
  'piaui', 'rio de janeiro', 'rio grande do norte', 'rio grande do sul', 'rondonia', 'roraima', 'santa catarina',
  'sao paulo', 'sergipe', 'tocantins'
]);

function isLikelyBrazilAddress(churchSettings = {}) {
  const stateHint = stripDiacritics(churchSettings.state).toLowerCase();
  const zipDigits = normalize(churchSettings.zipCode).replace(/\D/g, '');
  return brazilStateHints.has(stateHint) || zipDigits.length === 8;
}

function buildGeocodeCandidates(input) {
  if (typeof input === 'string') {
    const query = normalize(input);
    return query ? [{ q: query }] : [];
  }

  const churchSettings = hydrateChurchSettings(input || {});
  const street = normalize(churchSettings.address);
  const city = normalize(churchSettings.city);
  const state = normalize(churchSettings.state);
  const zipCode = normalize(churchSettings.zipCode);
  const formatted = formatChurchAddress(churchSettings);
  const likelyBrazil = isLikelyBrazilAddress(churchSettings);
  const country = likelyBrazil ? 'Brazil' : '';
  const countryCodes = likelyBrazil ? 'br' : '';
  const candidates = [];
  const seen = new Set();

  const addCandidate = (candidate) => {
    const key = JSON.stringify(candidate);
    if (!seen.has(key)) {
      seen.add(key);
      candidates.push(candidate);
    }
  };

  const structured = {};
  if (street) structured.street = street;
  if (city) structured.city = city;
  if (state) structured.state = state;
  if (zipCode) structured.postalcode = zipCode;
  if (country) structured.country = country;
  if (Object.keys(structured).length > 0) {
    addCandidate({ structured, countrycodes: countryCodes });
  }

  if (formatted) {
    addCandidate({
      q: country ? `${formatted}, ${country}` : formatted,
      countrycodes: countryCodes
    });
  }

  const cityStateZip = [city, state, zipCode, country].filter(Boolean).join(', ');
  if (cityStateZip) {
    addCandidate({
      q: cityStateZip,
      countrycodes: countryCodes
    });
  }

  if (zipCode) {
    addCandidate({
      q: country ? `${zipCode}, ${country}` : zipCode,
      countrycodes: countryCodes
    });
  }

  return candidates;
}

async function geocodeAddress(input) {
  const candidates = buildGeocodeCandidates(input);
  if (!candidates.length) return null;

  const endpoint = 'https://nominatim.openstreetmap.org/search';

  for (const candidate of candidates) {
    try {
      const url = new URL(endpoint);
      url.searchParams.set('format', 'jsonv2');
      url.searchParams.set('limit', '1');
      url.searchParams.set('addressdetails', '1');
      if (candidate.countrycodes) {
        url.searchParams.set('countrycodes', candidate.countrycodes);
      }

      if (candidate.q) {
        url.searchParams.set('q', candidate.q);
      } else {
        Object.entries(candidate.structured || {}).forEach(([key, value]) => {
          if (value) {
            url.searchParams.set(key, value);
          }
        });
      }

      const response = await fetchWithTimeout(url.toString(), {
        method: 'GET',
        headers: {
          Accept: 'application/json',
          'Accept-Language': 'pt-BR,en-US;q=0.8',
          'User-Agent': 'ChurchCRM/1.0 (Visitation Map Base)'
        }
      });

      if (!response.ok) continue;

      const rows = await response.json();
      if (!Array.isArray(rows) || rows.length === 0) continue;

      const first = rows[0] || {};
      const lat = normalizeLatitude(first.lat);
      const lng = normalizeLongitude(first.lon);
      if (!lat || !lng) continue;

      return { lat, lng };
    } catch {
      // Try the next candidate.
    }
  }

  return null;
}

function parsePessoasServiceType(value) {
  const raw = normalize(value).toLowerCase();
  if (!raw) return '';
  if (raw.includes('wed')) return 'wed';
  if (raw.includes('sun') && raw.includes('pm')) return 'sun_pm';
  if (raw.includes('sun') && raw.includes('am')) return 'sun_am';
  if (raw === 'wed') return 'wed';
  if (raw === 'sun_pm') return 'sun_pm';
  if (raw === 'sun_am') return 'sun_am';
  return '';
}

function upsertPessoasRow(bucket, row) {
  const serviceDate = toIsoDate(row.serviceDate);
  const serviceType = parsePessoasServiceType(row.serviceType);
  const headcount = Number.parseInt(row.headcount, 10);
  const note = normalize(row.note);

  if (!serviceDate || !serviceType || Number.isNaN(headcount) || headcount < 0) {
    return;
  }

  const key = `${serviceDate}|${serviceType}`;
  bucket.set(key, { serviceDate, serviceType, headcount, note });
}

function parsePessoasDashboardHtml(html) {
  const $ = cheerio.load(html || '');
  const rows = new Map();

  $('table tbody tr').each((_, tr) => {
    const cells = $(tr).find('td');
    if (cells.length < 3) return;

    const classNames = normalize($(cells[1]).find('.badge').attr('class'));
    const classMatch = classNames.match(/badge-(wed|sun_am|sun_pm)/);
    const serviceType = classMatch ? classMatch[1] : parsePessoasServiceType($(cells[1]).text());

    upsertPessoasRow(rows, {
      serviceDate: $(cells[0]).text(),
      serviceType,
      headcount: $(cells[2]).text(),
      note: cells.length > 3 ? $(cells[3]).text() : ''
    });
  });

  const scriptBlob = $('script')
    .map((_, script) => $(script).html() || '')
    .get()
    .find((content) => content.includes('const chartData ='));

  if (scriptBlob) {
    const match = scriptBlob.match(/const\s+chartData\s*=\s*(\{[\s\S]*?\})\s*;/);
    if (match) {
      try {
        const chartData = JSON.parse(match[1]);
        ['wed', 'sun_am', 'sun_pm'].forEach((serviceType) => {
          (chartData[serviceType] || []).forEach((point) => {
            upsertPessoasRow(rows, {
              serviceDate: point.x,
              serviceType,
              headcount: point.y,
              note: ''
            });
          });
        });
      } catch {
        // Keep table rows only if chart parse fails.
      }
    }
  }

  const years = $('#year option')
    .map((_, option) => Number.parseInt(normalize($(option).attr('value') || $(option).text()), 10))
    .get()
    .filter((year) => Number.isInteger(year));

  const loginPageDetected = $('form[action="/login"], form[action*="login"]').length > 0;
  const invalidPasswordDetected = /incorrect password/i.test($.text());

  return {
    rows: Array.from(rows.values()),
    years: Array.from(new Set(years)),
    loginPageDetected,
    invalidPasswordDetected
  };
}

async function fetchPessoasPage(url, cookieHeader = '', redirectsLeft = 5) {
  const response = await fetchWithTimeout(url, {
    method: 'GET',
    headers: {
      Accept: 'text/html,application/xhtml+xml',
      ...(cookieHeader ? { Cookie: cookieHeader } : {})
    },
    redirect: 'manual'
  });

  const updatedCookieHeader = appendCookies(cookieHeader, getSetCookie(response.headers));
  const location = response.headers.get('location');
  if (location && response.status >= 300 && response.status < 400 && redirectsLeft > 0) {
    const nextUrl = new URL(location, url).toString();
    return fetchPessoasPage(nextUrl, updatedCookieHeader, redirectsLeft - 1);
  }

  return {
    html: await response.text(),
    cookieHeader: updatedCookieHeader
  };
}

async function tryReadPessoasAttendanceRows(appUrl) {
  let rootUrl;
  let loginUrl;
  try {
    rootUrl = new URL('/', appUrl).toString();
    loginUrl = new URL('/login', appUrl).toString();
  } catch {
    return { ok: false, reason: `Invalid pessoas URL: ${appUrl}`, rows: [] };
  }

  let cookieHeader = '';
  let page = await fetchPessoasPage(rootUrl, cookieHeader);
  cookieHeader = page.cookieHeader;

  let parsed = parsePessoasDashboardHtml(page.html);
  if (!parsed.rows.length || parsed.loginPageDetected) {
    const password = getPessoasAppPassword();
    if (!password) {
      return {
        ok: false,
        reason: 'Set PESSOAS_APP_PASSWORD (or LEGACY_APP_PASSWORD) so CRM can sign in and import.',
        rows: []
      };
    }

    const loginResponse = await fetchWithTimeout(loginUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        Accept: 'text/html,application/xhtml+xml',
        ...(cookieHeader ? { Cookie: cookieHeader } : {})
      },
      body: new URLSearchParams({ password }).toString(),
      redirect: 'manual'
    });

    cookieHeader = appendCookies(cookieHeader, getSetCookie(loginResponse.headers));
    const destination = loginResponse.headers.get('location')
      ? new URL(loginResponse.headers.get('location'), loginUrl).toString()
      : rootUrl;

    page = await fetchPessoasPage(destination, cookieHeader);
    cookieHeader = page.cookieHeader;
    parsed = parsePessoasDashboardHtml(page.html);

    if (parsed.loginPageDetected || parsed.invalidPasswordDetected) {
      return { ok: false, reason: 'Pessoas login failed. Check PESSOAS_APP_PASSWORD value.', rows: [] };
    }
  }

  const merged = new Map();
  parsed.rows.forEach((row) => upsertPessoasRow(merged, row));

  for (const year of parsed.years) {
    const yearUrl = new URL(rootUrl);
    yearUrl.searchParams.set('year', String(year));
    const yearPage = await fetchPessoasPage(yearUrl.toString(), cookieHeader);
    const yearParsed = parsePessoasDashboardHtml(yearPage.html);
    yearParsed.rows.forEach((row) => upsertPessoasRow(merged, row));
  }

  const rows = Array.from(merged.values()).sort((a, b) => {
    if (a.serviceDate === b.serviceDate) return a.serviceType.localeCompare(b.serviceType);
    return a.serviceDate.localeCompare(b.serviceDate);
  });

  if (!rows.length) {
    return { ok: false, reason: 'No records found on pessoas app.', rows: [] };
  }

  return { ok: true, reason: '', rows };
}

async function readPessoasAttendanceRows() {
  const candidates = getPessoasAppUrls();
  if (!candidates.length) {
    return { ok: false, reason: 'Set METRICS_APP_URL to your pessoas app domain first.', rows: [] };
  }

  let lastError = '';
  for (const appUrl of candidates) {
    try {
      const result = await tryReadPessoasAttendanceRows(appUrl);
      if (result.ok) {
        return result;
      }
      lastError = result.reason || lastError;
    } catch (error) {
      lastError = `Could not import from ${appUrl}: ${error.message}`;
    }
  }

  return {
    ok: false,
    reason: lastError || 'Could not import from pessoas app.',
    rows: []
  };
}

const membershipTypes = ['Prospect', 'Member', 'Voting Member'];
const genderOptions = ['Male', 'Female', 'Unspecified'];
const serviceTypes = ['wed', 'sun_am', 'sun_pm'];
const followUpStages = [
  { value: 'new_visitor', label: 'New Visitor' },
  { value: 'contacted', label: 'Contacted' },
  { value: 'visited', label: 'Visited' },
  { value: 'connected', label: 'Connected' },
  { value: 'member', label: 'Member' }
];
const followUpStageValues = followUpStages.map((stage) => stage.value);
const followUpStageLabels = followUpStages.reduce((acc, stage) => {
  acc[stage.value] = stage.label;
  return acc;
}, {});
const serviceTypeLabels = {
  wed: 'Wednesday',
  sun_am: 'Sunday AM',
  sun_pm: 'Sunday PM'
};
const familyRelationshipTypes = ['spouse', 'parent', 'child'];

function normalizeFamilyRelationshipType(value) {
  const type = normalize(value);
  return familyRelationshipTypes.includes(type) ? type : '';
}

function linkPeopleRelationship(people, sourceId, relationship, targetId) {
  const type = normalizeFamilyRelationshipType(relationship);
  if (!type || sourceId === targetId) return false;

  const source = people.find((entry) => entry.id === sourceId);
  const target = people.find((entry) => entry.id === targetId);
  if (!source || !target) return false;

  ensurePersonRelationships(source);
  ensurePersonRelationships(target);

  let changed = false;
  if (type === 'spouse') {
    changed = addUniqueId(source.spouseIds, targetId) || changed;
    changed = addUniqueId(target.spouseIds, sourceId) || changed;
  }

  if (type === 'parent') {
    changed = addUniqueId(source.parentIds, targetId) || changed;
    changed = addUniqueId(target.childIds, sourceId) || changed;
    changed = removeId(source.childIds, targetId) || changed;
    changed = removeId(target.parentIds, sourceId) || changed;
  }

  if (type === 'child') {
    changed = addUniqueId(source.childIds, targetId) || changed;
    changed = addUniqueId(target.parentIds, sourceId) || changed;
    changed = removeId(source.parentIds, targetId) || changed;
    changed = removeId(target.childIds, sourceId) || changed;
  }

  if (changed) {
    const nowIso = new Date().toISOString();
    source.updatedAt = nowIso;
    target.updatedAt = nowIso;
  }

  return changed;
}

function unlinkPeopleRelationship(people, sourceId, relationship, targetId) {
  const type = normalizeFamilyRelationshipType(relationship);
  if (!type || sourceId === targetId) return false;

  const source = people.find((entry) => entry.id === sourceId);
  const target = people.find((entry) => entry.id === targetId);
  if (!source || !target) return false;

  ensurePersonRelationships(source);
  ensurePersonRelationships(target);

  let changed = false;
  if (type === 'spouse') {
    changed = removeId(source.spouseIds, targetId) || changed;
    changed = removeId(target.spouseIds, sourceId) || changed;
  }

  if (type === 'parent') {
    changed = removeId(source.parentIds, targetId) || changed;
    changed = removeId(target.childIds, sourceId) || changed;
  }

  if (type === 'child') {
    changed = removeId(source.childIds, targetId) || changed;
    changed = removeId(target.parentIds, sourceId) || changed;
  }

  if (changed) {
    const nowIso = new Date().toISOString();
    source.updatedAt = nowIso;
    target.updatedAt = nowIso;
  }

  return changed;
}

function removePersonFromRelationships(people, personId) {
  let changed = false;

  people.forEach((person) => {
    ensurePersonRelationships(person);
    let personChanged = false;
    personChanged = removeId(person.spouseIds, personId) || personChanged;
    personChanged = removeId(person.parentIds, personId) || personChanged;
    personChanged = removeId(person.childIds, personId) || personChanged;

    if (personChanged) {
      person.updatedAt = new Date().toISOString();
    }

    changed = personChanged || changed;
  });

  return changed;
}

function resolveRelationshipPeople(person, field, peopleById) {
  return normalizePersonRelationIds(person[field])
    .map((personId) => peopleById[personId])
    .filter(Boolean)
    .sort((a, b) => normalize(a.name).localeCompare(normalize(b.name)));
}

function normalizeFollowUpStage(value) {
  const stage = normalize(value);
  return followUpStageValues.includes(stage) ? stage : 'new_visitor';
}

function normalizeFollowUpStageFilter(value) {
  const normalized = normalize(value);
  if (!normalized || normalized === 'all') return 'all';
  return normalizeFollowUpStage(normalized);
}

function normalizeMembershipTypeFilter(value) {
  const membershipType = normalize(value);
  return membershipTypes.includes(membershipType) ? membershipType : '';
}

function normalizeSmartPeopleFilter(input = {}) {
  return {
    id: normalize(input.id) || id(),
    name: normalize(input.name),
    q: normalize(input.q),
    followups: normalizeFollowupsFilter(input.followups),
    membershipType: normalizeMembershipTypeFilter(input.membershipType),
    tag: normalize(input.tag),
    createdAt: normalize(input.createdAt) || new Date().toISOString()
  };
}

function buildPersonTimeline(person, followUps, visits) {
  const timeline = [];

  const createdAt = normalize(person.createdAt) || normalize(person.updatedAt);
  if (createdAt) {
    timeline.push({
      id: `person-created-${person.id}`,
      at: createdAt,
      title: 'Profile created',
      detail: 'Person record was added.',
      kind: 'person'
    });
  }

  if (normalize(person.updatedAt) && normalize(person.updatedAt) !== createdAt) {
    timeline.push({
      id: `person-updated-${person.id}`,
      at: person.updatedAt,
      title: 'Profile updated',
      detail: 'Contact or profile details were updated.',
      kind: 'person'
    });
  }

  (followUps || []).forEach((item) => {
    if (normalize(item.createdAt)) {
      timeline.push({
        id: `followup-created-${item.id}`,
        at: item.createdAt,
        title: `Follow-up added: ${item.title}`,
        detail: `Stage: ${followUpStageLabels[normalizeFollowUpStage(item.stage)]}${item.dueDate ? ` · Due: ${item.dueDate}` : ''}`,
        kind: 'followup'
      });
    }

    if (normalize(item.completedAt)) {
      timeline.push({
        id: `followup-completed-${item.id}`,
        at: item.completedAt,
        title: `Follow-up completed: ${item.title}`,
        detail: 'Marked complete.',
        kind: 'followup'
      });
    }
  });

  (visits || []).forEach((visit) => {
    const at = normalize(visit.createdAt) || normalize(visit.date);
    if (!at) return;
    timeline.push({
      id: `visit-${visit.id}`,
      at,
      title: `Visit logged: ${visit.summary}`,
      detail: visit.nextStep ? `Next step: ${visit.nextStep}` : 'No next step provided.',
      kind: 'visit'
    });
  });

  return timeline
    .filter((item) => normalize(item.at))
    .sort((a, b) => new Date(b.at) - new Date(a.at));
}

app.locals.formatDateDMY = formatDateDMY;
app.locals.membershipTypes = membershipTypes;
app.locals.genderOptions = genderOptions;
app.locals.serviceTypeLabels = serviceTypeLabels;
app.locals.followUpStages = followUpStages;
app.locals.followUpStageLabels = followUpStageLabels;

function isAuthEnabled() {
  return Boolean(getAdminAuthPassword() || getUserAuthPassword());
}

function getAdminAuthPassword() {
  const candidates = [
    process.env.CRM_ADMIN_PASSWORD,
    process.env.ADMIN_PASSWORD,
    process.env.APP_PASSWORD,
    process.env.PASSWORD,
    process.env.admin_password,
    process.env.password
  ];

  const first = candidates.find((value) => normalize(value).length > 0);
  return normalize(first);
}

function getUserAuthPassword() {
  return normalize(process.env.CRM_USER_PASSWORD);
}

function authPasswordSource() {
  const sources = [];
  if (normalize(process.env.CRM_ADMIN_PASSWORD)) sources.push('CRM_ADMIN_PASSWORD');
  else if (normalize(process.env.ADMIN_PASSWORD)) sources.push('ADMIN_PASSWORD');
  else if (normalize(process.env.APP_PASSWORD)) sources.push('APP_PASSWORD');
  else if (normalize(process.env.PASSWORD)) sources.push('PASSWORD');
  else if (normalize(process.env.admin_password)) sources.push('admin_password');
  else if (normalize(process.env.password)) sources.push('password');

  if (normalize(process.env.CRM_USER_PASSWORD)) sources.push('CRM_USER_PASSWORD');
  return sources.join('+');
}

function resolveLoginRole(submittedPassword) {
  const adminPassword = getAdminAuthPassword();
  if (adminPassword && safePasswordCompare(submittedPassword, adminPassword)) {
    return 'admin';
  }

  const userPassword = getUserAuthPassword();
  if (userPassword && safePasswordCompare(submittedPassword, userPassword)) {
    return 'user';
  }

  return '';
}

function isSessionAdmin(req) {
  if (!isAuthEnabled()) {
    return true;
  }

  if (req.session?.isAdmin) {
    return true;
  }

  // Backward compatibility: when only one password exists, authenticated users are admins.
  return Boolean(req.session?.isAuthenticated) && !getUserAuthPassword();
}

function requireAdmin(req, res, next) {
  if (isSessionAdmin(req)) {
    return next();
  }

  return res.status(403).send('Admin access required.');
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
  res.locals.isAdmin = isSessionAdmin(req);
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
    tags: parseTagsInput(csvField(row, ['tags', 'labels', 'label'])),
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
        tags: row.tags || [],
        membershipType: row.membershipType || '',
        joinedAt: row.joinedAt || '',
        createdAt: row.createdAt || '',
        baptismDate: row.baptismDate || '',
        totalContribution: row.totalContribution || '',
        spouseIds: [],
        parentIds: [],
        childIds: [],
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
    const returnTo = normalize(req.query.returnTo);
    if (returnTo.startsWith('/')) {
      return res.redirect(returnTo);
    }
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
  const returnTo = normalize(req.body.returnTo) || '/people';
  const role = resolveLoginRole(submitted);

  if (!role) {
    return res.status(401).render('login', {
      returnTo,
      error: 'Invalid password'
    });
  }

  req.session.isAuthenticated = true;
  req.session.role = role;
  req.session.isAdmin = role === 'admin';
  return res.redirect(returnTo);
});

app.get('/auth-status', (req, res) => {
  res.json({
    authEnabled: isAuthEnabled(),
    authPasswordSource: authPasswordSource() || null,
    hasAppSecret: Boolean(normalize(process.env.APP_SECRET)),
    isAuthenticated: Boolean(req.session?.isAuthenticated),
    role: req.session?.role || '',
    isAdmin: isSessionAdmin(req)
  });
});

app.post('/logout', (req, res) => {
  req.session.destroy(() => {
    res.redirect('/login');
  });
});

app.get('/settings/church', async (req, res, next) => {
  try {
    const data = await readData();
    const churchSettings = hydrateChurchSettings((data.settings || {}).church);
    const saveStatus = normalize(req.query.saved) === '1' ? 'saved' : '';
    const geocodeStatus = normalize(req.query.geocode);
    const isReadOnly = !isSessionAdmin(req);

    res.render('church-settings', {
      activeTab: 'church_settings',
      churchSettings,
      saveStatus,
      geocodeStatus,
      isReadOnly
    });
  } catch (err) {
    next(err);
  }
});

app.post('/settings/church', requireAdmin, async (req, res, next) => {
  try {
    let geocodeStatus = 'none';
    const churchSettings = hydrateChurchSettings({
      name: req.body.name,
      email: req.body.email,
      phone: req.body.phone,
      address: req.body.address,
      city: req.body.city,
      state: req.body.state,
      zipCode: req.body.zipCode
    });

    const geocodeQuery = formatChurchAddress(churchSettings);

    if (!geocodeQuery) {
      const params = new URLSearchParams({
        saved: '0',
        geocode: 'missing_address'
      });
      return res.redirect(`/settings/church?${params.toString()}`);
    }

    const geocoded = await geocodeAddress(churchSettings);
    if (geocoded) {
      churchSettings.mapLat = geocoded.lat;
      churchSettings.mapLng = geocoded.lng;
      geocodeStatus = 'success';
    } else {
      geocodeStatus = 'failed';
    }

    await updateData((data) => {
      data.settings = data.settings || {};
      const previousChurchSettings = hydrateChurchSettings((data.settings || {}).church);

      if (churchSettings.mapLat && churchSettings.mapLng) {
        data.settings.church = churchSettings;
      } else if (previousChurchSettings.mapLat && previousChurchSettings.mapLng) {
        // Keep prior resolved coordinates if geocoding is temporarily unavailable.
        data.settings.church = {
          ...churchSettings,
          mapLat: previousChurchSettings.mapLat,
          mapLng: previousChurchSettings.mapLng
        };
        geocodeStatus = 'failed_using_previous';
      } else {
        data.settings.church = churchSettings;
      }

      const currentVisitation = hydrateVisitationSettings((data.settings || {}).visitation, data.people || []);
      const resolvedChurchSettings = hydrateChurchSettings(data.settings.church);
      const resolvedChurchAddress = formatChurchAddress(resolvedChurchSettings);
      data.settings.visitation = {
        ...currentVisitation,
        mapCenterMode: 'church',
        mapCenterZoom: normalizeMapZoom(currentVisitation.mapCenterZoom, 17),
        profilePersonId: '',
        churchProfile: {
          name: resolvedChurchSettings.name,
          address: resolvedChurchAddress,
          lat: resolvedChurchSettings.mapLat,
          lng: resolvedChurchSettings.mapLng
        }
      };

      return data;
    });

    const params = new URLSearchParams({
      saved: '1',
      ...(geocodeStatus !== 'none' ? { geocode: geocodeStatus } : {})
    });

    return res.redirect(`/settings/church?${params.toString()}`);
  } catch (err) {
    return next(err);
  }
});

app.get('/', (req, res) => {
  res.redirect('/people');
});

app.get('/people', async (req, res, next) => {
  try {
    const data = await readData();
    const savedPeopleFilters = ((data.settings || {}).peopleSavedFilters || []).map((entry) =>
      normalizeSmartPeopleFilter(entry)
    );
    const selectedSmartFilterId = normalize(req.query.smartFilter);
    const selectedSmartFilter = savedPeopleFilters.find((entry) => entry.id === selectedSmartFilterId) || null;

    let q = normalize(req.query.q);
    let followups = normalizeFollowupsFilter(req.query.followups);
    let membershipTypeFilter = normalizeMembershipTypeFilter(req.query.membershipType);
    let tagFilter = normalize(req.query.tag);

    if (selectedSmartFilter) {
      q = selectedSmartFilter.q;
      followups = normalizeFollowupsFilter(selectedSmartFilter.followups);
      membershipTypeFilter = normalizeMembershipTypeFilter(selectedSmartFilter.membershipType);
      tagFilter = normalize(selectedSmartFilter.tag);
    }

    const qLower = q.toLowerCase();
    const tagFilterLower = tagFilter.toLowerCase();
    const now = new Date();

    const openFollowUpsByPerson = data.followUps
      .filter((item) => item.status !== 'completed')
      .reduce((acc, item) => {
        acc[item.personId] = (acc[item.personId] || 0) + 1;
        return acc;
      }, {});

    let people = sortByName(data.people).map((person) => ({
      ...enrichPerson(person),
      tags: normalizePersonTags(person.tags),
      openFollowUps: openFollowUpsByPerson[person.id] || 0
    }));

    if (qLower) {
      people = people.filter((person) =>
        [
          person.name,
          person.email,
          person.phone,
          person.sectionId,
          person.membershipType,
          person.notes,
          (person.tags || []).join(' ')
        ]
          .join(' ')
          .toLowerCase()
          .includes(qLower)
      );
    }

    if (followups === 'open') {
      people = people.filter((person) => person.openFollowUps > 0);
    }

    if (membershipTypeFilter) {
      people = people.filter((person) => normalize(person.membershipType) === membershipTypeFilter);
    }

    if (tagFilterLower) {
      people = people.filter((person) =>
        (person.tags || []).some((tag) => normalize(tag).toLowerCase() === tagFilterLower)
      );
    }

    const totalFilteredPeople = people.length;
    const requestedPageSize = Number.parseInt(normalize(req.query.perPage), 10);
    const pageSize = PEOPLE_DIRECTORY_PAGE_SIZES.includes(requestedPageSize)
      ? requestedPageSize
      : PEOPLE_DIRECTORY_DEFAULT_PAGE_SIZE;
    const totalPages = Math.max(1, Math.ceil(totalFilteredPeople / pageSize));
    const requestedPage = Number.parseInt(normalize(req.query.page), 10);
    const currentPage =
      Number.isInteger(requestedPage) && requestedPage > 0 ? Math.min(requestedPage, totalPages) : 1;
    const pageStartIndex = (currentPage - 1) * pageSize;
    const pagedPeople = people.slice(pageStartIndex, pageStartIndex + pageSize);
    const pageRangeStart = totalFilteredPeople ? pageStartIndex + 1 : 0;
    const pageRangeEnd = Math.min(totalFilteredPeople, pageStartIndex + pagedPeople.length);

    const basePaginationParams = new URLSearchParams();
    if (pageSize !== PEOPLE_DIRECTORY_DEFAULT_PAGE_SIZE) {
      basePaginationParams.set('perPage', String(pageSize));
    }
    if (selectedSmartFilter?.id) {
      basePaginationParams.set('smartFilter', selectedSmartFilter.id);
    } else {
      if (q) basePaginationParams.set('q', q);
      if (followups === 'open') basePaginationParams.set('followups', followups);
      if (membershipTypeFilter) basePaginationParams.set('membershipType', membershipTypeFilter);
      if (tagFilter) basePaginationParams.set('tag', tagFilter);
    }

    const buildPeoplePageUrl = (pageNumber) => {
      const params = new URLSearchParams(basePaginationParams.toString());
      if (pageNumber > 1) {
        params.set('page', String(pageNumber));
      } else {
        params.delete('page');
      }
      const query = params.toString();
      return query ? `/people?${query}` : '/people';
    };

    const paginationLinks = [];
    const startPage = Math.max(1, currentPage - 2);
    const endPage = Math.min(totalPages, currentPage + 2);
    for (let pageNumber = startPage; pageNumber <= endPage; pageNumber += 1) {
      paginationLinks.push({
        number: pageNumber,
        isCurrent: pageNumber === currentPage,
        url: buildPeoplePageUrl(pageNumber)
      });
    }

    const tagDisplayMap = new Map();
    data.people.forEach((person) => {
      normalizePersonTags(person.tags).forEach((tag) => {
        const key = tag.toLowerCase();
        if (!tagDisplayMap.has(key)) {
          tagDisplayMap.set(key, tag);
        }
      });
    });
    const availableTags = Array.from(tagDisplayMap.values()).sort((a, b) => a.localeCompare(b));

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
      people: pagedPeople,
      totalFilteredPeople,
      pageSize,
      pageSizeOptions: PEOPLE_DIRECTORY_PAGE_SIZES,
      currentPage,
      totalPages,
      pageRangeStart,
      pageRangeEnd,
      paginationLinks,
      previousPageUrl: currentPage > 1 ? buildPeoplePageUrl(currentPage - 1) : '',
      nextPageUrl: currentPage < totalPages ? buildPeoplePageUrl(currentPage + 1) : '',
      q,
      followups,
      membershipTypeFilter,
      tagFilter,
      availableTags,
      savedPeopleFilters,
      selectedSmartFilterId: selectedSmartFilter?.id || '',
      selectedSmartFilterName: selectedSmartFilter?.name || '',
      upcomingBirthdays
    });
  } catch (err) {
    next(err);
  }
});

app.get('/people/new', async (req, res, next) => {
  try {
    const isDuplicateError = normalize(req.query.error) === 'duplicate';
    const duplicateId = normalize(req.query.duplicateId);
    res.render('people-new', {
      activeTab: 'people_new',
      errorMessage: isDuplicateError
        ? 'Possible duplicate found. Open the existing record or adjust the details before creating.'
        : '',
      duplicateId,
      formData: {
        name: normalize(req.query.name),
        phone: normalize(req.query.phone),
        email: normalize(req.query.email),
        gender: normalize(req.query.gender),
        membershipType: normalize(req.query.membershipType),
        birthday: normalize(req.query.birthday),
        sectionId: normalize(req.query.sectionId),
        tags: normalize(req.query.tags),
        address: normalize(req.query.address),
        city: normalize(req.query.city),
        state: normalize(req.query.state),
        zipCode: normalize(req.query.zipCode),
        mapLat: normalize(req.query.mapLat),
        mapLng: normalize(req.query.mapLng),
        notes: normalize(req.query.notes)
      }
    });
  } catch (err) {
    next(err);
  }
});

app.post('/people/filters', async (req, res, next) => {
  try {
    const filter = normalizeSmartPeopleFilter({
      name: req.body.name,
      q: req.body.q,
      followups: req.body.followups,
      membershipType: req.body.membershipType,
      tag: req.body.tag
    });

    if (!filter.name) {
      return res.redirect('/people');
    }

    await updateData((data) => {
      data.settings = data.settings || {};
      data.settings.peopleSavedFilters = ((data.settings.peopleSavedFilters || [])
        .map((entry) => normalizeSmartPeopleFilter(entry))
        .filter((entry) => entry.name))
        .slice(0, 49);
      data.settings.peopleSavedFilters.push(filter);
      return data;
    });

    return res.redirect(`/people?smartFilter=${encodeURIComponent(filter.id)}`);
  } catch (err) {
    next(err);
  }
});

app.post('/people/filters/:filterId/delete', async (req, res, next) => {
  try {
    await updateData((data) => {
      data.settings = data.settings || {};
      data.settings.peopleSavedFilters = (data.settings.peopleSavedFilters || []).filter(
        (entry) => normalize(entry.id) !== req.params.filterId
      );
      return data;
    });
    return res.redirect('/people');
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

    const tab = normalize(req.query.tab) || 'profile';
    const allowedTabs = new Set(['profile', 'followups', 'visits', 'notes', 'timeline']);
    const currentTab = allowedTabs.has(tab) ? tab : 'profile';

    const followUps = data.followUps
      .filter((entry) => entry.personId === person.id)
      .map((entry) => ({
        ...entry,
        stage: normalizeFollowUpStage(entry.stage)
      }))
      .sort((a, b) => new Date(a.dueDate || a.createdAt) - new Date(b.dueDate || b.createdAt));

    const visits = data.visits
      .filter((entry) => entry.personId === person.id)
      .sort((a, b) => new Date(b.date) - new Date(a.date));

    const peopleById = data.people.reduce((acc, entry) => {
      acc[entry.id] = entry;
      return acc;
    }, {});
    const personWithRelationships = hydratePersonRelationships(person);
    const household = {
      spouses: resolveRelationshipPeople(personWithRelationships, 'spouseIds', peopleById),
      parents: resolveRelationshipPeople(personWithRelationships, 'parentIds', peopleById),
      children: resolveRelationshipPeople(personWithRelationships, 'childIds', peopleById)
    };
    const relationshipCandidates = sortByName(data.people)
      .filter((entry) => entry.id !== person.id)
      .map((entry) => ({
        id: entry.id,
        name: entry.name,
        membershipType: entry.membershipType || 'Prospect'
      }));

    const timeline = buildPersonTimeline(person, followUps, visits);

    res.render('people-profile', {
      activeTab: 'people',
      person: {
        ...enrichPerson(personWithRelationships),
        tags: normalizePersonTags(person.tags)
      },
      tab: currentTab,
      followUps,
      visits,
      timeline,
      household,
      relationshipCandidates,
      followUpOpenCount: followUps.filter((entry) => entry.status !== 'completed').length
    });
  } catch (err) {
    next(err);
  }
});

app.post('/people', async (req, res, next) => {
  try {
    const nowIso = new Date().toISOString();
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
      mapLat: normalizeLatitude(req.body.mapLat),
      mapLng: normalizeLongitude(req.body.mapLng),
      tags: parseTagsInput(req.body.tags),
      spouseIds: [],
      parentIds: [],
      childIds: [],
      createdAt: nowIso,
      updatedAt: nowIso
    };

    if (!person.name) {
      return res.redirect('/people/new');
    }

    let duplicateId = '';
    await updateData((data) => {
      const incomingName = normalize(person.name).toLowerCase();
      const incomingPhone = normalizePhone(person.phone);
      const incomingEmail = normalize(person.email).toLowerCase();

      const duplicate = data.people.find((existing) => {
        const existingName = normalize(existing.name).toLowerCase();
        const existingPhone = normalizePhone(existing.phone);
        const existingEmail = normalize(existing.email).toLowerCase();

        const sameEmail = incomingEmail && existingEmail && incomingEmail === existingEmail;
        const samePhone = incomingPhone && existingPhone && incomingPhone === existingPhone;
        const sameNameAndPhone = incomingName && samePhone && incomingName === existingName;

        return sameEmail || sameNameAndPhone;
      });

      if (duplicate) {
        duplicateId = duplicate.id;
        return data;
      }

      data.people.push(person);
      return data;
    });

    if (duplicateId) {
      const params = new URLSearchParams({
        error: 'duplicate',
        duplicateId,
        name: person.name,
        phone: person.phone,
        email: person.email,
        gender: person.gender,
        membershipType: person.membershipType,
        birthday: person.birthday ? formatDateDMY(person.birthday) : '',
        sectionId: person.sectionId,
        tags: (person.tags || []).join(', '),
        address: person.address,
        city: person.city,
        state: person.state,
        zipCode: person.zipCode,
        mapLat: person.mapLat,
        mapLng: person.mapLng,
        notes: person.notes
      });
      return res.redirect(`/people/new?${params.toString()}`);
    }

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
        write('mapLat', (value) => normalizeLatitude(value));
        write('mapLng', (value) => normalizeLongitude(value));
        write('tags', (value) => parseTagsInput(value));
        person.updatedAt = new Date().toISOString();
      }

      return data;
    });

    res.redirect(returnTo);
  } catch (err) {
    next(err);
  }
});

app.post('/people/:id/relationships', async (req, res, next) => {
  try {
    const fallback = `/people/${req.params.id}?tab=profile`;
    const returnTo = normalize(req.body.returnTo);
    const destination = returnTo.startsWith('/') ? returnTo : fallback;
    const relationship = normalizeFamilyRelationshipType(req.body.relationship);
    const targetId = normalize(req.body.targetId);

    if (!relationship || !targetId) {
      return res.redirect(destination);
    }

    await updateData((data) => {
      linkPeopleRelationship(data.people, req.params.id, relationship, targetId);
      return data;
    });

    return res.redirect(destination);
  } catch (err) {
    return next(err);
  }
});

app.post('/people/:id/relationships/:relationship/:targetId/delete', async (req, res, next) => {
  try {
    const fallback = `/people/${req.params.id}?tab=profile`;
    const returnTo = normalize(req.body.returnTo);
    const destination = returnTo.startsWith('/') ? returnTo : fallback;
    const relationship = normalizeFamilyRelationshipType(req.params.relationship);

    if (!relationship) {
      return res.redirect(destination);
    }

    await updateData((data) => {
      unlinkPeopleRelationship(data.people, req.params.id, relationship, req.params.targetId);
      return data;
    });

    return res.redirect(destination);
  } catch (err) {
    return next(err);
  }
});

app.post('/people/:id/delete', async (req, res, next) => {
  try {
    await updateData((data) => {
      data.people = data.people.filter((entry) => entry.id !== req.params.id);
      removePersonFromRelationships(data.people, req.params.id);
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

    const nowIso = new Date().toISOString();
    await updateData((data) => {
      data.followUps.push({
        id: id(),
        personId: req.params.id,
        title,
        dueDate: req.body.dueDate || '',
        notes: req.body.notes?.trim() || '',
        status: 'open',
        stage: normalizeFollowUpStage(req.body.stage),
        createdAt: nowIso,
        updatedAt: nowIso,
        stageUpdatedAt: nowIso,
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
        row.stage = normalizeFollowUpStage(row.stage);
        row.status = 'completed';
        row.completedAt = new Date().toISOString();
        row.updatedAt = new Date().toISOString();
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
        row.stage = normalizeFollowUpStage(row.stage);
        row.status = 'open';
        row.completedAt = '';
        row.updatedAt = new Date().toISOString();
      }

      return data;
    });

    res.redirect(`/people/${req.params.id}?tab=followups`);
  } catch (err) {
    next(err);
  }
});

app.post('/people/:id/followups/:followUpId/stage', async (req, res, next) => {
  try {
    await updateData((data) => {
      const row = data.followUps.find((entry) => entry.id === req.params.followUpId && entry.personId === req.params.id);
      if (row) {
        row.stage = normalizeFollowUpStage(req.body.stage);
        row.stageUpdatedAt = new Date().toISOString();
        row.updatedAt = new Date().toISOString();
      }
      return data;
    });

    return res.redirect(`/people/${req.params.id}?tab=followups`);
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
    const stage = normalizeFollowUpStageFilter(req.query.stage);
    const view = normalize(req.query.view) === 'board' ? 'board' : 'table';

    const peopleById = data.people.reduce((acc, person) => {
      acc[person.id] = person;
      return acc;
    }, {});

    let queue = data.followUps
      .map((item) => {
        const person = peopleById[item.personId];
        return {
          ...item,
          stage: normalizeFollowUpStage(item.stage),
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

    if (stage !== 'all') {
      queue = queue.filter((item) => item.stage === stage);
    }

    const board = followUpStages.map((stageOption) => ({
      ...stageOption,
      items: queue.filter((item) => item.stage === stageOption.value)
    }));

    res.render('followups', {
      activeTab: 'followups',
      queue,
      status,
      stage,
      view,
      board
    });
  } catch (err) {
    next(err);
  }
});

app.post('/followups/:followUpId/complete', async (req, res, next) => {
  try {
    const returnTo = normalize(req.body.returnTo) || '/followups?status=open';
    await updateData((data) => {
      const row = data.followUps.find((entry) => entry.id === req.params.followUpId);
      if (row) {
        row.stage = normalizeFollowUpStage(row.stage);
        row.status = 'completed';
        row.completedAt = new Date().toISOString();
        row.updatedAt = new Date().toISOString();
      }
      return data;
    });

    res.redirect(returnTo);
  } catch (err) {
    next(err);
  }
});

app.post('/followups/:followUpId/reopen', async (req, res, next) => {
  try {
    const returnTo = normalize(req.body.returnTo) || '/followups?status=completed';
    await updateData((data) => {
      const row = data.followUps.find((entry) => entry.id === req.params.followUpId);
      if (row) {
        row.stage = normalizeFollowUpStage(row.stage);
        row.status = 'open';
        row.completedAt = '';
        row.updatedAt = new Date().toISOString();
      }
      return data;
    });

    res.redirect(returnTo);
  } catch (err) {
    next(err);
  }
});

app.post('/followups/:followUpId/stage', async (req, res, next) => {
  try {
    const returnTo = normalize(req.body.returnTo) || '/followups?status=open';
    await updateData((data) => {
      const row = data.followUps.find((entry) => entry.id === req.params.followUpId);
      if (row) {
        row.stage = normalizeFollowUpStage(req.body.stage);
        row.stageUpdatedAt = new Date().toISOString();
        row.updatedAt = new Date().toISOString();
      }
      return data;
    });

    res.redirect(returnTo);
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
    const recordsForYear = [...yearRecords].sort((a, b) => {
      if (a.serviceDate === b.serviceDate) {
        return a.serviceType.localeCompare(b.serviceType);
      }
      return b.serviceDate.localeCompare(a.serviceDate);
    });
    const attendanceMessage = normalize(req.query.attendanceMessage);
    const attendanceTypeRaw = normalize(req.query.attendanceType);
    const attendanceType = ['success', 'info', 'warning', 'danger'].includes(attendanceTypeRaw)
      ? attendanceTypeRaw
      : 'info';

    res.render('metrics', {
      activeTab: 'metrics',
      year,
      years,
      chartData,
      recordsForYear,
      averages: avgBuckets,
      attendanceThisYear,
      attendanceTotalHeadcount,
      attendanceMessage,
      attendanceType,
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
    const requestedYear = Number.parseInt(req.query.year, 10);
    const returnYear = Number.isNaN(requestedYear) ? new Date().getFullYear() : requestedYear;

    let suggestedType = 'sun_am';
    if (dayOfWeek === 3) suggestedType = 'wed';
    else if (dayOfWeek === 0) suggestedType = 'sun_am';

    res.render('metrics-entry', {
      activeTab: 'metrics',
      record: null,
      date: toIsoDate(req.query.date || today) || today,
      type: normalizeServiceType(req.query.type || suggestedType),
      returnYear,
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
    const requestedYear = Number.parseInt(req.query.year, 10);
    const returnYear = Number.isNaN(requestedYear) ? new Date(record.serviceDate).getFullYear() : requestedYear;

    res.render('metrics-entry', {
      activeTab: 'metrics',
      record,
      date: toIsoDate(record.serviceDate),
      type: normalizeServiceType(record.serviceType),
      returnYear,
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
        returnYear: Number.parseInt(req.body.return_year, 10) || new Date().getFullYear(),
        error: 'Date, service type, and a non-negative headcount are required.'
      });
    }

    let action = 'created';
    await updateData((data) => {
      const existing = (data.attendanceRecords || []).find(
        (row) => toIsoDate(row.serviceDate) === serviceDate && normalizeServiceType(row.serviceType) === serviceType
      );

      if (existing) {
        action = 'updated';
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

    const params = new URLSearchParams({
      year: String(Number.parseInt(req.body.return_year, 10) || new Date(serviceDate).getFullYear()),
      attendanceType: 'success',
      attendanceMessage:
        action === 'created'
          ? 'Attendance record added.'
          : 'Attendance record updated for that same date/service.'
    });
    return res.redirect(`/metrics?${params.toString()}`);
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
        returnYear: Number.parseInt(req.body.return_year, 10) || new Date().getFullYear(),
        error: 'Date, service type, and a non-negative headcount are required.'
      });
    }

    let found = false;
    await updateData((data) => {
      const row = (data.attendanceRecords || []).find((entry) => entry.id === req.params.id);
      if (row) {
        found = true;
        row.serviceDate = serviceDate;
        row.serviceType = serviceType;
        row.headcount = headcount;
        row.note = note;
        row.updatedAt = new Date().toISOString();
      }
      return data;
    });

    const params = new URLSearchParams({
      year: String(Number.parseInt(req.body.return_year, 10) || new Date(serviceDate).getFullYear()),
      attendanceType: found ? 'success' : 'warning',
      attendanceMessage: found ? 'Attendance record updated.' : 'Record not found. It may have been deleted.'
    });
    return res.redirect(`/metrics?${params.toString()}`);
  } catch (err) {
    next(err);
  }
});

app.post('/metrics/import-pessoas', async (req, res, next) => {
  try {
    const source = await readPessoasAttendanceRows();
    if (!source.ok) {
      const params = new URLSearchParams({
        year: String(Number.parseInt(req.body.year, 10) || new Date().getFullYear()),
        attendanceType: 'warning',
        attendanceMessage: source.reason
      });
      return res.redirect(`/metrics?${params.toString()}`);
    }

    let imported = 0;
    let updated = 0;
    let skipped = 0;

    await updateData((data) => {
      data.attendanceRecords = data.attendanceRecords || [];
      const existingByKey = new Map(
        data.attendanceRecords.map((entry) => [
          `${toIsoDate(entry.serviceDate)}|${normalizeServiceType(entry.serviceType)}`,
          entry
        ])
      );

      source.rows.forEach((row) => {
        const serviceDate = toIsoDate(row.serviceDate);
        const serviceType = normalizeServiceType(row.serviceType);
        const headcount = Number.parseInt(row.headcount, 10);
        const note = normalize(row.note);
        if (!serviceDate || Number.isNaN(headcount) || headcount < 0) {
          skipped += 1;
          return;
        }

        const key = `${serviceDate}|${serviceType}`;
        const existing = existingByKey.get(key);
        if (existing) {
          const changed = existing.headcount !== headcount || normalize(existing.note) !== note;
          if (!changed) {
            skipped += 1;
            return;
          }
          existing.headcount = headcount;
          existing.note = note;
          existing.updatedAt = new Date().toISOString();
          updated += 1;
          return;
        }

        const created = {
          id: id(),
          serviceDate,
          serviceType,
          headcount,
          note,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        };
        data.attendanceRecords.push(created);
        existingByKey.set(key, created);
        imported += 1;
      });

      return data;
    });

    const params = new URLSearchParams({
      year: String(Number.parseInt(req.body.year, 10) || new Date().getFullYear()),
      attendanceType: 'success',
      attendanceMessage: `Pessoas import complete. Imported ${imported}, updated ${updated}, skipped ${skipped}.`
    });
    return res.redirect(`/metrics?${params.toString()}`);
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

const sectionStatusValues = new Set(['unclaimed', 'claimed', 'completed']);
const defaultSectionColor = '#0c4a6e';
const defaultFolderColor = '#20c997';

function normalizeSectionStatus(value) {
  const next = normalize(value).toLowerCase();
  return sectionStatusValues.has(next) ? next : 'unclaimed';
}

function parseBoolean(value) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value === 1;
  if (typeof value === 'string') {
    return value.toLowerCase() === 'true';
  }
  return false;
}

function normalizeChecklist(checklist, sectionId = 'section') {
  if (!Array.isArray(checklist)) {
    return [];
  }

  return checklist
    .map((item, index) => {
      if (typeof item === 'string') {
        const label = normalize(item);
        if (!label) return null;
        return {
          id: `${sectionId}-item-${index + 1}`,
          label,
          done: false,
          completedAt: ''
        };
      }

      if (!item || typeof item !== 'object') {
        return null;
      }

      const label = normalize(item.label || item.name || item.text);
      if (!label) {
        return null;
      }

      const done = parseBoolean(item.done);
      return {
        id: normalize(item.id) || `${sectionId}-item-${index + 1}`,
        label,
        done,
        completedAt: done ? normalize(item.completedAt) : ''
      };
    })
    .filter(Boolean);
}

function normalizeSectionStreets(streets, sectionId = 'section') {
  if (!Array.isArray(streets)) {
    return [];
  }

  return streets
    .map((street, index) => {
      if (!street || typeof street !== 'object') {
        return null;
      }

      const geojson = street.geojson || null;
      if (!geojson) {
        return null;
      }

      const name = normalize(street.name || street.label || street.text) || `Street ${index + 1}`;
      const done = parseBoolean(street.done);

      return {
        id: normalize(street.id) || `${sectionId}-street-${index + 1}`,
        name,
        color: normalize(street.color) || '#16a34a',
        done,
        completedAt: done ? normalize(street.completedAt) : '',
        geojson
      };
    })
    .filter(Boolean);
}

function hydrateFolder(folder, index = 0) {
  const folderId = normalize(folder?.id) || `folder-${index + 1}`;

  return {
    id: folderId,
    name: normalize(folder?.name) || `Folder ${index + 1}`,
    color: normalize(folder?.color) || defaultFolderColor,
    notes: normalize(folder?.notes)
  };
}

function hydrateSection(section, index = 0) {
  const sectionId = normalize(section?.id) || `section-${index + 1}`;
  const status = normalizeSectionStatus(section?.status);

  return {
    id: sectionId,
    name: normalize(section?.name),
    folderId: normalize(section?.folderId),
    color: normalize(section?.color) || defaultSectionColor,
    lastVisited: toIsoDate(section?.lastVisited),
    notes: normalize(section?.notes),
    geojson: section?.geojson || null,
    status,
    claimedBy: status === 'unclaimed' ? '' : normalize(section?.claimedBy),
    claimedAt: status === 'unclaimed' ? '' : normalize(section?.claimedAt),
    completedAt: status === 'completed' ? normalize(section?.completedAt) : '',
    checklist: normalizeChecklist(section?.checklist, sectionId),
    streets: normalizeSectionStreets(section?.streets, sectionId)
  };
}

function hydrateMarker(marker, index = 0) {
  const markerId = normalize(marker?.id) || `marker-${index + 1}`;
  const lat = normalizeLatitude(marker?.lat);
  const lng = normalizeLongitude(marker?.lng);

  return {
    id: markerId,
    name: normalize(marker?.name) || `Marker ${index + 1}`,
    notes: normalize(marker?.notes),
    color: normalize(marker?.color) || '#2563eb',
    lat,
    lng
  };
}

app.get('/visitation', async (req, res, next) => {
  try {
    const data = await readData();
    const people = Array.isArray(data.people) ? data.people : [];
    const mapProfiles = sortByName(people).map((entry) => mapProfileSummary(entry));
    const churchSettings = hydrateChurchSettings((data.settings || {}).church);
    const mapSettings = mergeVisitationWithChurchSettings(
      hydrateVisitationSettings((data.settings || {}).visitation, people),
      churchSettings
    );

    res.render('visitation', {
      activeTab: 'visitation',
      sections: (data.sections || []).map((entry, index) => hydrateSection(entry, index)),
      folders: (data.folders || []).map((entry, index) => hydrateFolder(entry, index)),
      markers: (data.markers || []).map((entry, index) => hydrateMarker(entry, index)),
      mapProfiles,
      mapSettings
    });
  } catch (err) {
    next(err);
  }
});

app.get('/api/visitation/settings', async (req, res, next) => {
  try {
    const data = await readData();
    const people = Array.isArray(data.people) ? data.people : [];
    const mapProfiles = sortByName(people).map((entry) => mapProfileSummary(entry));
    const churchSettings = hydrateChurchSettings((data.settings || {}).church);
    const mapSettings = mergeVisitationWithChurchSettings(
      hydrateVisitationSettings((data.settings || {}).visitation, people),
      churchSettings
    );

    res.json({ mapSettings, mapProfiles });
  } catch (err) {
    next(err);
  }
});

app.post('/api/visitation/settings', async (req, res, next) => {
  try {
    const current = await readData();
    const people = Array.isArray(current.people) ? current.people : [];
    const mapProfiles = sortByName(people).map((entry) => mapProfileSummary(entry));
    let churchSettings = hydrateChurchSettings((current.settings || {}).church);
    const currentVisitation = hydrateVisitationSettings((current.settings || {}).visitation, people);
    const nextZoom = normalize(req.body.mapCenterZoom)
      ? normalizeMapZoom(req.body.mapCenterZoom, currentVisitation.mapCenterZoom || 17)
      : currentVisitation.mapCenterZoom || 17;
    let mapSettings = mergeVisitationWithChurchSettings(
      {
        ...currentVisitation,
        mapCenterZoom: nextZoom
      },
      churchSettings
    );

    const churchAddress = formatChurchAddress(churchSettings);
    const hasChurchCoordinates = Boolean(mapSettings.churchProfile.lat && mapSettings.churchProfile.lng);
    if (!hasChurchCoordinates && churchAddress) {
      const geocoded = await geocodeAddress(churchSettings);
      if (geocoded) {
        churchSettings = {
          ...churchSettings,
          mapLat: geocoded.lat,
          mapLng: geocoded.lng
        };
        mapSettings = mergeVisitationWithChurchSettings(mapSettings, churchSettings);
      }
    }

    if (!mapSettings.churchProfile.lat || !mapSettings.churchProfile.lng) {
      return res.status(400).json({
        error: 'Set church address so map base can be located.'
      });
    }

    await updateData((data) => {
      data.settings = data.settings || {};
      if (churchSettings.mapLat && churchSettings.mapLng) {
        const existingChurchSettings = hydrateChurchSettings((data.settings || {}).church);
        data.settings.church = {
          ...existingChurchSettings,
          ...churchSettings
        };
      }
      data.settings.visitation = mapSettings;
      return data;
    });

    res.json({
      ok: true,
      mapSettings,
      mapProfiles
    });
  } catch (err) {
    next(err);
  }
});

app.get('/api/folders', async (req, res, next) => {
  try {
    const data = await readData();
    res.json((data.folders || []).map((entry, index) => hydrateFolder(entry, index)));
  } catch (err) {
    next(err);
  }
});

app.get('/api/markers', async (req, res, next) => {
  try {
    const data = await readData();
    res.json(
      (data.markers || [])
        .map((entry, index) => hydrateMarker(entry, index))
        .filter((entry) => Boolean(entry.lat && entry.lng))
    );
  } catch (err) {
    next(err);
  }
});

app.post('/api/markers', async (req, res, next) => {
  try {
    const payload = hydrateMarker(
      {
        id: req.body.id,
        name: req.body.name,
        notes: req.body.notes,
        color: req.body.color,
        lat: req.body.lat,
        lng: req.body.lng
      },
      0
    );

    if (!payload.lat || !payload.lng) {
      return res.status(400).json({ error: 'lat and lng are required' });
    }

    if (!payload.id || payload.id === 'marker-1') {
      payload.id = id();
    }

    await updateData((data) => {
      data.markers = Array.isArray(data.markers) ? data.markers : [];
      const idx = data.markers.findIndex((entry) => normalize(entry.id) === payload.id);
      if (idx === -1) {
        data.markers.push(payload);
      } else {
        data.markers[idx] = payload;
      }
      return data;
    });

    res.json({ ok: true, marker: payload });
  } catch (err) {
    next(err);
  }
});

app.delete('/api/markers/:id', async (req, res, next) => {
  try {
    const markerId = normalize(req.params.id);
    await updateData((data) => {
      data.markers = (data.markers || []).filter((entry) => normalize(entry.id) !== markerId);
      return data;
    });
    res.json({ ok: true });
  } catch (err) {
    next(err);
  }
});

app.post('/api/folders', async (req, res, next) => {
  try {
    const payload = {
      id: normalize(req.body.id) || id(),
      name: normalize(req.body.name),
      color: normalize(req.body.color) || defaultFolderColor,
      notes: normalize(req.body.notes)
    };

    if (!payload.name) {
      return res.status(400).json({ error: 'name is required' });
    }

    await updateData((data) => {
      data.folders = Array.isArray(data.folders) ? data.folders : [];
      const idx = data.folders.findIndex((entry) => normalize(entry.id) === payload.id);

      if (idx === -1) {
        data.folders.push(payload);
      } else {
        data.folders[idx] = payload;
      }

      return data;
    });

    res.json({ ok: true, folder: payload });
  } catch (err) {
    next(err);
  }
});

app.delete('/api/folders/:id', async (req, res, next) => {
  try {
    const folderId = normalize(req.params.id);

    await updateData((data) => {
      data.folders = (data.folders || []).filter((entry) => normalize(entry.id) !== folderId);
      data.sections = (data.sections || []).map((section) => {
        if (normalize(section.folderId) !== folderId) {
          return section;
        }

        return {
          ...section,
          folderId: ''
        };
      });
      return data;
    });

    res.json({ ok: true });
  } catch (err) {
    next(err);
  }
});

app.get('/api/sections', async (req, res, next) => {
  try {
    const data = await readData();
    res.json((data.sections || []).map((entry, index) => hydrateSection(entry, index)));
  } catch (err) {
    next(err);
  }
});

app.post('/api/sections', async (req, res, next) => {
  try {
    const sectionId = normalize(req.body.id) || id();
    const status = normalizeSectionStatus(req.body.status);
    const nowIso = new Date().toISOString();
    let claimedBy = normalize(req.body.claimedBy);
    let claimedAt = normalize(req.body.claimedAt);
    let completedAt = normalize(req.body.completedAt);

    if (status === 'unclaimed') {
      claimedBy = '';
      claimedAt = '';
      completedAt = '';
    } else if (status === 'claimed') {
      completedAt = '';
      if (claimedBy && !claimedAt) {
        claimedAt = nowIso;
      }
    } else if (status === 'completed') {
      if (claimedBy && !claimedAt) {
        claimedAt = nowIso;
      }
      if (!completedAt) {
        completedAt = nowIso;
      }
    }

    const payload = {
      id: sectionId,
      name: normalize(req.body.name),
      folderId: normalize(req.body.folderId),
      color: normalize(req.body.color) || defaultSectionColor,
      lastVisited: toIsoDate(req.body.lastVisited),
      notes: normalize(req.body.notes),
      geojson: req.body.geojson || null,
      status,
      claimedBy,
      claimedAt,
      completedAt,
      checklist: normalizeChecklist(req.body.checklist, sectionId).map((item) => ({
        ...item,
        completedAt: item.done ? item.completedAt || nowIso : ''
      })),
      streets: normalizeSectionStreets(req.body.streets, sectionId).map((street) => ({
        ...street,
        completedAt: street.done ? street.completedAt || nowIso : ''
      }))
    };

    if (!payload.name || !payload.geojson) {
      return res.status(400).json({ error: 'name and geojson are required' });
    }

    if (payload.folderId) {
      const current = await readData();
      const folderExists = (current.folders || []).some(
        (entry) => normalize(entry.id) === payload.folderId
      );

      if (!folderExists) {
        return res.status(400).json({ error: 'folderId is invalid' });
      }
    }

    await updateData((data) => {
      data.sections = Array.isArray(data.sections) ? data.sections : [];
      const idx = data.sections.findIndex((entry) => normalize(entry.id) === payload.id);
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

app.post('/api/sections/:id/claim', async (req, res, next) => {
  try {
    const sectionId = normalize(req.params.id);
    const claimedBy = normalize(req.body.claimedBy);
    const nowIso = new Date().toISOString();
    let updated = null;

    if (!claimedBy) {
      return res.status(400).json({ error: 'claimedBy is required' });
    }

    await updateData((data) => {
      data.sections = Array.isArray(data.sections) ? data.sections : [];
      const idx = data.sections.findIndex((entry) => normalize(entry.id) === sectionId);

      if (idx === -1) {
        return data;
      }

      const section = hydrateSection(data.sections[idx], idx);
      updated = {
        ...section,
        status: 'claimed',
        claimedBy,
        claimedAt: nowIso,
        completedAt: ''
      };
      data.sections[idx] = updated;
      return data;
    });

    if (!updated) {
      return res.status(404).json({ error: 'section not found' });
    }

    res.json({ ok: true, section: hydrateSection(updated) });
  } catch (err) {
    next(err);
  }
});

app.post('/api/sections/:id/unclaim', async (req, res, next) => {
  try {
    const sectionId = normalize(req.params.id);
    let updated = null;

    await updateData((data) => {
      data.sections = Array.isArray(data.sections) ? data.sections : [];
      const idx = data.sections.findIndex((entry) => normalize(entry.id) === sectionId);

      if (idx === -1) {
        return data;
      }

      const section = hydrateSection(data.sections[idx], idx);
      updated = {
        ...section,
        status: 'unclaimed',
        claimedBy: '',
        claimedAt: '',
        completedAt: ''
      };
      data.sections[idx] = updated;
      return data;
    });

    if (!updated) {
      return res.status(404).json({ error: 'section not found' });
    }

    res.json({ ok: true, section: hydrateSection(updated) });
  } catch (err) {
    next(err);
  }
});

app.post('/api/sections/:id/complete', async (req, res, next) => {
  try {
    const sectionId = normalize(req.params.id);
    const workerName = normalize(req.body.claimedBy);
    const nowIso = new Date().toISOString();
    let updated = null;

    await updateData((data) => {
      data.sections = Array.isArray(data.sections) ? data.sections : [];
      const idx = data.sections.findIndex((entry) => normalize(entry.id) === sectionId);

      if (idx === -1) {
        return data;
      }

      const section = hydrateSection(data.sections[idx], idx);
      const claimedBy = workerName || section.claimedBy;
      updated = {
        ...section,
        status: 'completed',
        claimedBy,
        claimedAt: section.claimedAt || (claimedBy ? nowIso : ''),
        completedAt: nowIso,
        lastVisited: section.lastVisited || nowIso.slice(0, 10)
      };
      data.sections[idx] = updated;
      return data;
    });

    if (!updated) {
      return res.status(404).json({ error: 'section not found' });
    }

    res.json({ ok: true, section: hydrateSection(updated) });
  } catch (err) {
    next(err);
  }
});

app.post('/api/sections/:id/checklist', async (req, res, next) => {
  try {
    const sectionId = normalize(req.params.id);
    const itemId = normalize(req.body.itemId);
    const nowIso = new Date().toISOString();
    let updated = null;
    let itemFound = false;
    let sectionExists = false;

    if (!Array.isArray(req.body.items) && !itemId) {
      return res.status(400).json({ error: 'itemId or items[] is required' });
    }

    await updateData((data) => {
      data.sections = Array.isArray(data.sections) ? data.sections : [];
      const idx = data.sections.findIndex((entry) => normalize(entry.id) === sectionId);

      if (idx === -1) {
        return data;
      }

      sectionExists = true;

      const section = hydrateSection(data.sections[idx], idx);
      let checklist = normalizeChecklist(section.checklist, section.id);

      if (Array.isArray(req.body.items)) {
        checklist = normalizeChecklist(req.body.items, section.id).map((entry) => ({
          ...entry,
          completedAt: entry.done ? entry.completedAt || nowIso : ''
        }));
        itemFound = true;
      } else {
        checklist = checklist.map((entry) => {
          if (entry.id !== itemId) {
            return entry;
          }

          itemFound = true;
          const done = parseBoolean(req.body.done);
          return {
            ...entry,
            done,
            completedAt: done ? entry.completedAt || nowIso : ''
          };
        });
      }

      if (!itemFound) {
        return data;
      }

      updated = {
        ...section,
        checklist
      };
      data.sections[idx] = updated;
      return data;
    });

    if (!sectionExists) {
      return res.status(404).json({ error: 'section not found' });
    }

    if (!itemFound) {
      return res.status(400).json({ error: 'checklist item not found' });
    }

    res.json({ ok: true, section: hydrateSection(updated) });
  } catch (err) {
    next(err);
  }
});

app.post('/api/sections/:id/streets/:streetId', async (req, res, next) => {
  try {
    const sectionId = normalize(req.params.id);
    const streetId = normalize(req.params.streetId);
    const nowIso = new Date().toISOString();
    let updated = null;
    let sectionExists = false;
    let streetFound = false;

    if (!streetId) {
      return res.status(400).json({ error: 'streetId is required' });
    }

    await updateData((data) => {
      data.sections = Array.isArray(data.sections) ? data.sections : [];
      const idx = data.sections.findIndex((entry) => normalize(entry.id) === sectionId);

      if (idx === -1) {
        return data;
      }

      sectionExists = true;
      const section = hydrateSection(data.sections[idx], idx);
      const done = parseBoolean(req.body.done);
      const streets = normalizeSectionStreets(section.streets, section.id).map((street) => {
        if (street.id !== streetId) {
          return street;
        }

        streetFound = true;
        return {
          ...street,
          done,
          completedAt: done ? street.completedAt || nowIso : ''
        };
      });

      if (!streetFound) {
        return data;
      }

      updated = {
        ...section,
        streets
      };
      data.sections[idx] = updated;
      return data;
    });

    if (!sectionExists) {
      return res.status(404).json({ error: 'section not found' });
    }

    if (!streetFound) {
      return res.status(400).json({ error: 'street not found' });
    }

    res.json({ ok: true, section: hydrateSection(updated) });
  } catch (err) {
    next(err);
  }
});

app.delete('/api/sections/:id', async (req, res, next) => {
  try {
    const sectionId = normalize(req.params.id);
    await updateData((data) => {
      data.sections = (data.sections || []).filter((entry) => normalize(entry.id) !== sectionId);
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
  console.log(`Dashboard running on http://localhost:${PORT}`);
});
