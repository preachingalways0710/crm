let map;
let drawnItems;
let activeLayer = null;

async function api(path, options = {}) {
  const res = await fetch(path, {
    headers: { 'Content-Type': 'application/json' },
    ...options
  });

  if (!res.ok) {
    throw new Error('API request failed');
  }

  return res.json();
}

function setForm(section) {
  document.getElementById('sectionId').value = section?.id || '';
  document.getElementById('sectionName').value = section?.name || '';
  document.getElementById('sectionColor').value = section?.color || '#0c4a6e';
  document.getElementById('sectionDate').value = section?.lastVisited || '';
  document.getElementById('sectionNotes').value = section?.notes || '';
}

function clearDrawing() {
  if (activeLayer) {
    drawnItems.removeLayer(activeLayer);
    activeLayer = null;
  }
}

function renderSectionsList(sections) {
  const root = document.getElementById('sectionsList');
  root.innerHTML = '';

  if (!sections.length) {
    root.innerHTML = '<p class="text-secondary mb-0">No sections saved yet.</p>';
    return;
  }

  sections.forEach((section) => {
    const row = document.createElement('div');
    row.className = 'section-row';
    row.innerHTML = `
      <div>
        <strong>${section.name}</strong>
        <p class="text-secondary mb-0">Last visited: ${section.lastVisited || 'n/a'}</p>
      </div>
      <div class="d-flex gap-2">
        <button class="btn btn-outline-primary btn-sm" data-action="edit">Edit</button>
        <button class="btn btn-outline-danger btn-sm" data-action="delete">Delete</button>
      </div>
    `;

    row.querySelector('[data-action="edit"]').addEventListener('click', () => {
      setForm(section);
      clearDrawing();
      const layer = L.geoJSON(section.geojson, {
        style: { color: section.color || '#0c4a6e' }
      });
      layer.eachLayer((child) => {
        drawnItems.addLayer(child);
        activeLayer = child;
      });
      map.fitBounds(layer.getBounds());
    });

    row.querySelector('[data-action="delete"]').addEventListener('click', async () => {
      const ok = window.confirm(`Delete ${section.name}?`);
      if (!ok) return;
      await api(`/api/sections/${section.id}`, { method: 'DELETE' });
      await init();
    });

    root.appendChild(row);
  });
}

async function init() {
  if (!map) {
    map = L.map('map').setView([47.6062, -122.3321], 11);

    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      maxZoom: 19,
      attribution: '&copy; OpenStreetMap contributors'
    }).addTo(map);

    drawnItems = new L.FeatureGroup();
    map.addLayer(drawnItems);

    const drawControl = new L.Control.Draw({
      draw: {
        marker: false,
        polyline: false,
        circle: false,
        rectangle: false,
        circlemarker: false,
        polygon: {
          allowIntersection: false,
          showArea: true
        }
      },
      edit: {
        featureGroup: drawnItems
      }
    });

    map.addControl(drawControl);

    map.on(L.Draw.Event.CREATED, (event) => {
      clearDrawing();
      activeLayer = event.layer;
      drawnItems.addLayer(activeLayer);
    });
  }

  const sections = await api('/api/sections');
  renderSectionsList(sections);
}

document.addEventListener('DOMContentLoaded', async () => {
  await init();

  document.getElementById('clearDrawing').addEventListener('click', () => {
    clearDrawing();
    setForm(null);
  });

  document.getElementById('sectionForm').addEventListener('submit', async (event) => {
    event.preventDefault();

    if (!activeLayer) {
      window.alert('Draw a section polygon on the map first.');
      return;
    }

    const payload = {
      id: document.getElementById('sectionId').value || undefined,
      name: document.getElementById('sectionName').value,
      color: document.getElementById('sectionColor').value,
      lastVisited: document.getElementById('sectionDate').value,
      notes: document.getElementById('sectionNotes').value,
      geojson: activeLayer.toGeoJSON()
    };

    await api('/api/sections', {
      method: 'POST',
      body: JSON.stringify(payload)
    });

    setForm(null);
    clearDrawing();
    await init();
  });
});
