const fs = require('fs/promises');
const path = require('path');

const dataDir = path.join(__dirname, '..', 'data');
const dataFile = path.join(dataDir, 'data.json');

const defaultData = {
  people: [],
  events: [],
  sections: [],
  visits: [],
  followUps: [],
  forms: [],
  formSubmissions: [],
  settings: {
    attendanceAppUrl: ''
  }
};

let writeQueue = Promise.resolve();

async function ensureFile() {
  await fs.mkdir(dataDir, { recursive: true });

  try {
    await fs.access(dataFile);
  } catch {
    await fs.writeFile(dataFile, JSON.stringify(defaultData, null, 2));
  }
}

async function readData() {
  await ensureFile();
  const raw = await fs.readFile(dataFile, 'utf-8');
  const parsed = JSON.parse(raw);

  return {
    ...defaultData,
    ...parsed,
    settings: {
      ...defaultData.settings,
      ...(parsed.settings || {})
    }
  };
}

function writeData(nextData) {
  writeQueue = writeQueue.then(() =>
    fs.writeFile(dataFile, JSON.stringify(nextData, null, 2), 'utf-8')
  );

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
