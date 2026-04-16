const path = require('path');
const crypto = require('crypto');
const express = require('express');
const { readData, updateData } = require('./lib/dataStore');

const app = express();
const PORT = process.env.PORT || 3000;

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use('/static', express.static(path.join(__dirname, 'public')));

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

app.get('/', (req, res) => {
  res.redirect('/people');
});

app.get('/people', async (req, res, next) => {
  try {
    const data = await readData();

    const openFollowUpsByPerson = data.followUps
      .filter((item) => item.status !== 'completed')
      .reduce((acc, item) => {
        acc[item.personId] = (acc[item.personId] || 0) + 1;
        return acc;
      }, {});

    const people = sortByName(data.people).map((person) => ({
      ...enrichPerson(person),
      openFollowUps: openFollowUpsByPerson[person.id] || 0
    }));

    res.render('people', {
      activeTab: 'people',
      people
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
  console.error(err);
  res.status(500).send('Something went wrong. Check server logs.');
});

app.listen(PORT, () => {
  console.log(`Church dashboard running on http://localhost:${PORT}`);
});
