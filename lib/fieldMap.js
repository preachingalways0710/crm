const express = require('express');
const path = require('path');
const zones = require('./fieldMapZones.json');
const { readData, updateData } = require('./dataStore');

const router = express.Router();
const uuid = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

function fieldState(data) {
  return {
    stretches: data.fieldMap?.stretches || [],
    completed: data.fieldMap?.completed || [],
    pins: (data.followUps || [])
      .filter((row) => row.fieldPoint)
      .map(({ id, fieldPoint, title, notes, dueDate, status }) => ({
        id,
        ...fieldPoint,
        title,
        notes,
        dueDate,
        status
      }))
  };
}

// Use the same paper coordinates as the drawing. Include the boundary streets.
function inside(point, polygon) {
  if (!Array.isArray(point) || point.length !== 2 || !point.every(Number.isFinite)) return false;
  const [x, y] = point;
  let contained = false;
  for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i++) {
    const [a, b] = polygon[i];
    const [c, d] = polygon[j];
    const dx = c - a,
      dy = d - b;
    const t = Math.max(0, Math.min(1, ((x - a) * dx + (y - b) * dy) / (dx * dx + dy * dy)));
    if (Math.hypot(x - a - t * dx, y - b - t * dy) <= 2) return true;
    if (b > y !== d > y && x < (dx * (y - b)) / dy + a) contained = !contained;
  }
  return contained;
}

router.get('/map', async (req, res) => {
  res.set('Cache-Control', 'no-store').render('field-map', { zones, state: fieldState(await readData()) });
});

router.get('/map/reference.jpg', (_req, res) => {
  res
    .set('Cache-Control', 'private, max-age=86400')
    .sendFile(path.join(__dirname, '../assets/field-map-reference.jpg'));
});

router.put('/api/field-map/zones/:zone', async (req, res) => {
  const zone = zones.find((item) => item.id === Number(req.params.zone));
  const completed = req.body?.completed;
  if (!zone || typeof completed !== 'boolean')
    return res.status(400).json({ error: 'Choose a zone and its completion status.' });
  const data = await updateData((data) => {
    data.fieldMap ||= {};
    const current = new Set(data.fieldMap.completed || []);
    if (completed) current.add(zone.id);
    else current.delete(zone.id);
    data.fieldMap.completed = [...current];
  });
  res.json(fieldState(data));
});

router.put('/api/field-map/marks/:id', async (req, res) => {
  const id = req.params.id;
  const { type, zoneId, start, end, title, notes = '', dueDate = '' } = req.body || {};
  const zone = zones.find((item) => item.id === zoneId);
  if (!zone || !uuid.test(id) || !inside(start, zone.polygon)) {
    return res.status(400).json({ error: 'Tap inside the selected zone.' });
  }
  if (type === 'stretch') {
    if (!inside(end, zone.polygon) || Math.hypot(end[0] - start[0], end[1] - start[1]) < 2) {
      return res.status(400).json({ error: 'Choose a different endpoint inside this zone.' });
    }
    // A straight line can leave a concave zone even when both endpoints are inside.
    const steps = Math.ceil(Math.hypot(end[0] - start[0], end[1] - start[1]) / 2);
    for (let i = 1; i < steps; i++) {
      if (
        !inside(
          start.map((n, axis) => n + ((end[axis] - n) * i) / steps),
          zone.polygon
        )
      ) {
        return res.status(400).json({ error: 'Mark one stretch inside this zone at a time.' });
      }
    }
  } else if (type === 'pin') {
    const date = new Date(dueDate + 'T12:00:00Z');
    if (
      typeof title !== 'string' ||
      !title.trim() ||
      title.length > 160 ||
      typeof notes !== 'string' ||
      notes.length > 1500 ||
      typeof dueDate !== 'string' ||
      (dueDate &&
        (!/^\d{4}-\d{2}-\d{2}$/.test(dueDate) ||
          !Number.isFinite(+date) ||
          date.toISOString().slice(0, 10) !== dueDate))
    ) {
      return res
        .status(400)
        .json({ error: 'Enter a visit label and a valid return date, or leave the date blank.' });
    }
  } else return res.status(400).json({ error: 'Unknown map mark.' });

  const data = await updateData((data) => {
    data.fieldMap ||= {};
    data.fieldMap.stretches ||= [];
    if (data.fieldMap.stretches.some((s) => s.id === id) || data.followUps.some((f) => f.id === id)) return;
    const now = new Date().toISOString();
    if (type === 'stretch') {
      data.fieldMap.stretches.push({ id, zoneId, start, end, createdAt: now });
    } else {
      data.followUps.push({
        id,
        fieldPoint: { zoneId, start },
        title: title.trim(),
        notes: notes.trim(),
        dueDate,
        contactMethod: 'visit',
        stage: 'new_visitor',
        status: 'open',
        createdAt: now,
        updatedAt: now,
        stageUpdatedAt: now,
        completedAt: ''
      });
    }
  });
  res.json(fieldState(data));
});

router.delete('/api/field-map/stretches/:id', async (req, res) => {
  const data = await updateData((data) => {
    if (data.fieldMap?.stretches)
      data.fieldMap.stretches = data.fieldMap.stretches.filter((s) => s.id !== req.params.id);
  });
  res.json(fieldState(data));
});

module.exports = router;
