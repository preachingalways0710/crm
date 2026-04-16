const path = require('path');
const crypto = require('crypto');
const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const { parse } = require('csv-parse/sync');
const { readData, updateData } = require('./lib/dataStore');

const app = express();
const PORT = process.env.PORT || 3000;

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use('/static', express.static(path.join(__dirname, 'public')));
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

function calculateAge(birthday) {
  if (!birthday) return null;

  const today = new Date();
  const dob = new Date(birthday);

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

  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    return raw;
  }

  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) {
    return '';
  }

  return parsed.toISOString().slice(0, 10);
}

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

app.get('/', (req, res) => {
  res.redirect('/people');
});

app.get('/people', async (req, res, next) => {
  try {
    const data = await readData();
    const q = normalize(req.query.q).toLowerCase();
    const followups = normalize(req.query.followups) || 'all';

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

    res.render('people', {
      activeTab: 'people',
      people,
      q,
      followups
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
      birthday: req.body.birthday || '',
      sectionId: req.body.sectionId || '',
      notes: req.body.notes?.trim() || '',
      gender: req.body.gender || '',
      ageGroup: req.body.ageGroup || '',
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
      return res.redirect('/people');
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
        write('birthday', (value) => value || '');
        write('sectionId', (value) => value || '');
        write('notes', (value) => value?.trim() || '');
        write('gender', (value) => value || '');
        write('ageGroup', (value) => value || '');
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
    const attendanceAppUrl = process.env.ATTENDANCE_APP_URL || data.settings.attendanceAppUrl;

    const now = new Date();
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

    res.render('metrics', {
      activeTab: 'metrics',
      attendanceAppUrl,
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
