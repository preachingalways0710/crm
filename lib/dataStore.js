const fs = require('fs/promises');
const path = require('path');
const mysql = require('mysql2/promise');

const dataDir = path.join(__dirname, '..', 'data');
const dataFile = path.join(dataDir, 'data.json');

const defaultData = {
  people: [],
  events: [],
  attendanceRecords: [],
  folders: [],
  sections: [],
  visits: [],
  followUps: [],
  forms: [],
  formSubmissions: [],
  settings: {
    attendanceAppUrl: '',
    peopleSavedFilters: [],
    church: {
      name: '',
      email: '',
      phone: '',
      address: '',
      city: '',
      state: '',
      zipCode: '',
      mapLat: '',
      mapLng: ''
    },
    visitation: {
      mapCenterMode: 'church',
      mapCenterZoom: 13,
      profilePersonId: '',
      churchProfile: {
        name: '',
        address: '',
        lat: '',
        lng: ''
      }
    }
  }
};

let writeQueue = Promise.resolve();
let poolPromise;
let dbInitPromise;

function useDatabaseStorage() {
  return Boolean(
    process.env.DB_HOST &&
      process.env.DB_NAME &&
      process.env.DB_USER &&
      process.env.DB_PASSWORD
  );
}

function normalizeData(parsed) {
  return {
    ...defaultData,
    ...(parsed || {}),
    settings: {
      ...defaultData.settings,
      ...((parsed || {}).settings || {})
    }
  };
}

async function ensureFile() {
  await fs.mkdir(dataDir, { recursive: true });

  try {
    await fs.access(dataFile);
  } catch {
    await fs.writeFile(dataFile, JSON.stringify(defaultData, null, 2));
  }
}

async function readFileState() {
  await ensureFile();
  const raw = await fs.readFile(dataFile, 'utf-8');
  return normalizeData(JSON.parse(raw));
}

async function getPool() {
  if (!poolPromise) {
    poolPromise = mysql.createPool({
      host: process.env.DB_HOST,
      port: Number(process.env.DB_PORT || 3306),
      user: process.env.DB_USER,
      password: process.env.DB_PASSWORD,
      database: process.env.DB_NAME,
      waitForConnections: true,
      connectionLimit: 10
    });
  }

  return poolPromise;
}

function parsePayload(payload) {
  if (!payload) return { ...defaultData };

  if (typeof payload === 'string') {
    return JSON.parse(payload);
  }

  if (Buffer.isBuffer(payload)) {
    return JSON.parse(payload.toString('utf-8'));
  }

  return payload;
}

async function ensureDatabase() {
  if (!dbInitPromise) {
    dbInitPromise = (async () => {
      const pool = await getPool();

      await pool.query(`
        CREATE TABLE IF NOT EXISTS app_state (
          id TINYINT PRIMARY KEY,
          payload JSON NOT NULL,
          updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
        )
      `);

      const [rows] = await pool.query('SELECT payload FROM app_state WHERE id = 1 LIMIT 1');

      if (!rows.length) {
        const seed = await readFileState();
        await pool.query('INSERT INTO app_state (id, payload) VALUES (1, ?)', [JSON.stringify(seed)]);
      }
    })();
  }

  return dbInitPromise;
}

async function readData() {
  if (!useDatabaseStorage()) {
    return readFileState();
  }

  await ensureDatabase();
  const pool = await getPool();
  const [rows] = await pool.query('SELECT payload FROM app_state WHERE id = 1 LIMIT 1');

  if (!rows.length) {
    return { ...defaultData };
  }

  return normalizeData(parsePayload(rows[0].payload));
}

function writeData(nextData) {
  writeQueue = writeQueue.then(async () => {
    if (!useDatabaseStorage()) {
      await ensureFile();
      await fs.writeFile(dataFile, JSON.stringify(nextData, null, 2), 'utf-8');
      return;
    }

    await ensureDatabase();
    const pool = await getPool();
    await pool.query('UPDATE app_state SET payload = ?, updated_at = NOW() WHERE id = 1', [JSON.stringify(nextData)]);
  });

  return writeQueue;
}

async function updateData(updater) {
  const data = await readData();
  const next = await updater(data);
  await writeData(next || data);
  return next || data;
}

module.exports = {
  readData,
  updateData
};
