let map;
let markerLayer;
let searchMarker;
let baseLayer;
let mapLabelLayer;
let churchMarker;

const boot = window.__PEOPLE_MAP_BOOTSTRAP__ || {
  church: { name: '', lat: '', lng: '' },
  items: [],
  filters: { tags: [], membershipTypes: [] },
  initialSelectionId: '',
  clubKids: { connected: false, totalKids: 0 }
};

const state = {
  items: Array.isArray(boot.items) ? boot.items : [],
  selectedId: safeText(boot.initialSelectionId),
  searchText: '',
  membershipFilter: '',
  tagFilter: '',
  placingMode: false,
  markersById: new Map(),
  mapStyle: 'satellite',
  bearing: -28
};

function safeText(value) {
  return (value || '').toString().trim();
}

function escapeHtml(value) {
  return safeText(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function parseCoordinate(value, min, max) {
  const parsed = Number.parseFloat(safeText(value).replace(',', '.'));
  if (!Number.isFinite(parsed)) return null;
  if (parsed < min || parsed > max) return null;
  return parsed;
}

async function api(path, options = {}) {
  const response = await fetch(path, {
    headers: { 'Content-Type': 'application/json' },
    ...options
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload?.error || 'Request failed.');
  }
  return payload;
}

function normalizeItem(item) {
  return {
    ...item,
    id: safeText(item?.id),
    name: safeText(item?.name) || 'Unnamed',
    membershipType: safeText(item?.membershipType),
    tags: Array.isArray(item?.tags) ? item.tags.map((tag) => safeText(tag)).filter(Boolean) : [],
    address: safeText(item?.address),
    lat: safeText(item?.lat),
    lng: safeText(item?.lng),
    phone: safeText(item?.phone),
    phoneDigits: safeText(item?.phoneDigits),
    whatsappUrl: safeText(item?.whatsappUrl),
    notes: safeText(item?.notes),
    sectionId: safeText(item?.sectionId),
    email: safeText(item?.email),
    photoUrl: safeText(item?.photoUrl),
    birthday: safeText(item?.birthday),
    archivedAt: safeText(item?.archivedAt),
    openFollowUps: Number.isFinite(Number(item?.openFollowUps)) ? Number(item.openFollowUps) : 0,
    latestVisit: item?.latestVisit || null,
    nextFollowUp: item?.nextFollowUp || null,
    profileUrl: safeText(item?.profileUrl),
    visitsUrl: safeText(item?.visitsUrl),
    followUpsUrl: safeText(item?.followUpsUrl),
    clubKids: item?.clubKids
      ? {
          id: safeText(item.clubKids.id),
          birthday: safeText(item.clubKids.birthday),
          photoUrl: safeText(item.clubKids.photoUrl),
          points: Number.isFinite(Number(item.clubKids.points)) ? Number(item.clubKids.points) : 0
        }
      : null
  };
}

function formatBirthday(value) {
  const raw = safeText(value);
  if (!raw) return '';
  const date = new Date(`${raw}T00:00:00`);
  if (Number.isNaN(date.getTime())) return raw;
  return date.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  });
}

function calculateAge(value) {
  const raw = safeText(value);
  if (!raw) return null;
  const birthday = new Date(`${raw}T00:00:00`);
  if (Number.isNaN(birthday.getTime())) return null;
  const now = new Date();
  let age = now.getFullYear() - birthday.getFullYear();
  const hasHadBirthday =
    now.getMonth() > birthday.getMonth() ||
    (now.getMonth() === birthday.getMonth() && now.getDate() >= birthday.getDate());
  if (!hasHadBirthday) age -= 1;
  return age >= 0 ? age : null;
}

function renderAvatar(item, size = 'lg') {
  const classes = size === 'sm' ? 'people-map-avatar people-map-avatar-sm' : 'people-map-avatar';
  if (item.photoUrl) {
    return `<img class="${classes}" src="${escapeHtml(item.photoUrl)}" alt="${escapeHtml(item.name)}" loading="lazy" />`;
  }
  const initial = escapeHtml((item.name || '?').charAt(0).toUpperCase() || '?');
  return `<span class="${classes} people-map-avatar-fallback">${initial}</span>`;
}

function getFilteredItems() {
  const search = state.searchText.toLowerCase();
  return state.items.filter((item) => {
    if (state.membershipFilter && item.membershipType !== state.membershipFilter) return false;
    if (state.tagFilter && !item.tags.includes(state.tagFilter)) return false;
    if (!search) return true;
    return [
      item.name,
      item.phone,
      item.email,
      item.address,
      item.membershipType,
      item.sectionId,
      item.notes,
      item.tags.join(' ')
    ]
      .join(' ')
      .toLowerCase()
      .includes(search);
  });
}

function badgeColorForItem(item) {
  if (item.archivedAt) return '#6b7280';
  if (item.openFollowUps > 0) return '#f59e0b';
  if (item.tags.includes('children') || item.tags.includes('childrens') || item.tags.includes('kids')) return '#2563eb';
  if (item.membershipType.toLowerCase().includes('convid')) return '#14b8a6';
  if (item.membershipType.toLowerCase().includes('membro')) return '#16a34a';
  return '#dc2626';
}

function statusText(message) {
  const el = document.getElementById('peopleMapStatus');
  if (el) {
    el.textContent = message;
  }
}

function createBaseLayer(mode) {
  if (mode === 'standard') {
    return L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      maxZoom: 19,
      attribution: '&copy; OpenStreetMap contributors'
    });
  }

  return L.tileLayer(
    'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
    {
      maxZoom: 19,
      attribution: 'Tiles &copy; Esri'
    }
  );
}

function createLabelLayer() {
  return L.tileLayer('https://{s}.basemaps.cartocdn.com/light_only_labels/{z}/{x}/{y}{r}.png', {
    maxZoom: 20,
    subdomains: 'abcd',
    opacity: 0.9,
    attribution: '&copy; OpenStreetMap contributors &copy; CARTO'
  });
}

function updateMapStyleControls() {
  const standardBtn = document.getElementById('peopleMapStandardBtn');
  const satelliteBtn = document.getElementById('peopleMapSatelliteBtn');
  if (standardBtn) {
    standardBtn.classList.toggle('active', state.mapStyle === 'standard');
  }
  if (satelliteBtn) {
    satelliteBtn.classList.toggle('active', state.mapStyle === 'satellite');
  }
}

function applyMapStyle(mode) {
  if (!map) return;
  state.mapStyle = mode === 'standard' ? 'standard' : 'satellite';

  if (baseLayer && map.hasLayer(baseLayer)) {
    map.removeLayer(baseLayer);
  }
  if (mapLabelLayer && map.hasLayer(mapLabelLayer)) {
    map.removeLayer(mapLabelLayer);
  }

  baseLayer = createBaseLayer(state.mapStyle);
  baseLayer.addTo(map);

  if (state.mapStyle === 'satellite') {
    mapLabelLayer = createLabelLayer();
    mapLabelLayer.addTo(map);
  } else {
    mapLabelLayer = null;
  }

  updateMapStyleControls();
}

function setMapBearing(value) {
  state.bearing = Number.isFinite(value) ? value : 0;
  if (map && typeof map.setBearing === 'function') {
    map.setBearing(state.bearing);
  }
}

function rotateMapBy(deltaDegrees) {
  setMapBearing((state.bearing || 0) + deltaDegrees);
}

function renderChipGroup(rootId, items, activeValue, onPick) {
  const root = document.getElementById(rootId);
  if (!root) return;
  root.innerHTML = '';

  const allButton = document.createElement('button');
  allButton.type = 'button';
  allButton.className = `btn btn-sm ${activeValue ? 'btn-outline-secondary' : 'btn-secondary'}`;
  allButton.textContent = 'All';
  allButton.addEventListener('click', () => onPick(''));
  root.appendChild(allButton);

  items.forEach((value) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = `btn btn-sm ${activeValue === value ? 'btn-primary' : 'btn-outline-primary'}`;
    button.textContent = value;
    button.addEventListener('click', () => onPick(value));
    root.appendChild(button);
  });
}

function renderPeopleList() {
  const root = document.getElementById('peopleMapList');
  const badge = document.getElementById('peopleMapCountBadge');
  const filtered = getFilteredItems();

  if (badge) {
    badge.textContent = `${filtered.length} shown`;
  }
  if (!root) return;

  root.innerHTML = '';
  if (!filtered.length) {
    root.innerHTML = '<div class="text-secondary">No people match this filter.</div>';
    return;
  }

  filtered.forEach((item) => {
    const row = document.createElement('button');
    row.type = 'button';
    row.className = `people-map-person ${state.selectedId === item.id ? 'active' : ''}`;
    const birthdayLabel = formatBirthday(item.birthday);
    row.innerHTML = `
      <span class="people-map-person-accent" style="background:${badgeColorForItem(item)};"></span>
      ${renderAvatar(item, 'sm')}
      <span class="people-map-person-main">
        <strong>${escapeHtml(item.name)}</strong>
        <small>${escapeHtml(item.address || item.membershipType || 'No address yet')}</small>
        ${
          birthdayLabel
            ? `<small class="people-map-person-meta">Birthday: ${escapeHtml(birthdayLabel)}</small>`
            : ''
        }
      </span>
      <span class="people-map-person-badges">
        ${item.clubKids ? '<span class="badge bg-orange-lt">ClubKids</span>' : ''}
        <span class="badge ${item.lat && item.lng ? 'bg-green-lt' : 'bg-secondary-lt'}">${item.lat && item.lng ? 'Pinned' : 'Needs pin'}</span>
      </span>
    `;
    row.addEventListener('click', () => {
      selectPerson(item.id, { focus: true });
    });
    root.appendChild(row);
  });
}

function renderDetail() {
  const root = document.getElementById('peopleMapDetail');
  if (!root) return;
  const item = state.items.find((entry) => entry.id === state.selectedId);
  const placeBtn = document.getElementById('peopleMapPlacePinBtn');

  if (placeBtn) {
    placeBtn.disabled = !item;
    placeBtn.textContent = state.placingMode ? 'Click Map To Save Pin' : 'Place / Update Pin';
  }

  if (!item) {
    root.innerHTML = '<div class="people-map-detail-empty">Click a pin or choose a person from the list to open their contact card.</div>';
    return;
  }

  const latestVisit = item.latestVisit
    ? `<div><strong>Last visit:</strong> ${escapeHtml(item.latestVisit.date || '-')}<br>${escapeHtml(item.latestVisit.summary || '')}</div>`
    : '<div><strong>Last visit:</strong> none yet</div>';
  const nextFollowUp = item.nextFollowUp
    ? `<div><strong>Next follow-up:</strong> ${escapeHtml(item.nextFollowUp.title || 'Open follow-up')}${item.nextFollowUp.dueDate ? ` · ${escapeHtml(item.nextFollowUp.dueDate)}` : ''}</div>`
    : '<div><strong>Next follow-up:</strong> none open</div>';
  const directionsUrl = item.lat && item.lng ? `https://www.google.com/maps?q=${item.lat},${item.lng}` : '';
  const safeTags = item.tags.map((tag) => `<span class="badge bg-secondary-lt">${escapeHtml(tag)}</span>`).join(' ');
  const birthdayLabel = formatBirthday(item.birthday);
  const age = calculateAge(item.birthday);
  const birthdayLine = birthdayLabel
    ? `<div><strong>Birthday:</strong> ${escapeHtml(birthdayLabel)}${age !== null ? ` · Age ${age}` : ''}</div>`
    : '<div><strong>Birthday:</strong> not saved yet</div>';
  const clubKidsLine = item.clubKids
    ? `<div class="people-map-source-card"><strong>ClubKids:</strong> linked child profile${item.clubKids.points ? ` · ${item.clubKids.points} points` : ''}</div>`
    : '';

  root.innerHTML = `
    <div class="people-map-detail-card">
      <div class="people-map-detail-header">
        <div class="people-map-detail-identity">
          ${renderAvatar(item)}
          <div>
            <h4>${escapeHtml(item.name)}</h4>
            <p>${escapeHtml(item.membershipType || 'No membership type')}${item.sectionId ? ` · ${escapeHtml(item.sectionId)}` : ''}</p>
          </div>
        </div>
        <div class="people-map-detail-header-meta">
          ${item.clubKids ? '<span class="badge bg-orange-lt">ClubKids Sync</span>' : ''}
          <span class="badge ${item.archivedAt ? 'bg-yellow-lt' : 'bg-azure-lt'}">${item.archivedAt ? 'Archived' : `${item.openFollowUps} open follow-ups`}</span>
        </div>
      </div>
      <div class="people-map-detail-body">
        <div>${escapeHtml(item.address || 'No address saved yet.')}</div>
        <div>${escapeHtml(item.phone || 'No phone')}${item.email ? ` · ${escapeHtml(item.email)}` : ''}</div>
        ${birthdayLine}
        ${clubKidsLine}
        <div class="people-map-tag-row">${safeTags || '<span class="text-secondary">No tags yet</span>'}</div>
        ${latestVisit}
        ${nextFollowUp}
        <div>${item.notes ? escapeHtml(item.notes) : '<span class="text-secondary">No notes yet.</span>'}</div>
      </div>
      <div class="people-map-detail-actions">
        <a class="btn btn-primary btn-sm" href="${item.profileUrl}">Open CRM Profile</a>
        <a class="btn btn-outline-secondary btn-sm ${item.whatsappUrl ? '' : 'disabled'}" href="${item.whatsappUrl || '#'}" target="_blank" rel="noreferrer">WhatsApp</a>
        <a class="btn btn-outline-secondary btn-sm ${directionsUrl ? '' : 'disabled'}" href="${directionsUrl || '#'}" target="_blank" rel="noreferrer">Open Route</a>
      </div>
      <form id="peopleMapVisitForm" class="people-map-visit-form">
        <div class="row g-2">
          <div class="col-12 col-md-3">
            <label class="form-label">Visit Date</label>
            <input class="form-control" name="date" type="date" value="${new Date().toISOString().slice(0, 10)}" />
          </div>
          <div class="col-12 col-md-9">
            <label class="form-label">Visit Summary</label>
            <input class="form-control" name="summary" placeholder="Visited mother, invited children to ministry..." />
          </div>
          <div class="col-12">
            <label class="form-label">Next Step</label>
            <input class="form-control" name="nextStep" placeholder="Return Saturday, bring transport info..." />
          </div>
          <div class="col-12">
            <button class="btn btn-success" type="submit">Log Visit From Map</button>
          </div>
        </div>
      </form>
    </div>
  `;

  const form = document.getElementById('peopleMapVisitForm');
  if (form) {
    form.addEventListener('submit', async (event) => {
      event.preventDefault();
      try {
        const formData = new FormData(form);
        const payload = {
          date: safeText(formData.get('date')),
          summary: safeText(formData.get('summary')),
          nextStep: safeText(formData.get('nextStep'))
        };
        const response = await api(`/api/people-map/people/${item.id}/visit`, {
          method: 'POST',
          body: JSON.stringify(payload)
        });
        item.latestVisit = response.visit || null;
        renderDetail();
        statusText(`Visit logged for "${item.name}".`);
      } catch (error) {
        window.alert(error.message || 'Could not save visit.');
      }
    });
  }
}

function renderMarkers() {
  if (!markerLayer) return;
  markerLayer.clearLayers();
  state.markersById.clear();

  getFilteredItems().forEach((item) => {
    const lat = parseCoordinate(item.lat, -90, 90);
    const lng = parseCoordinate(item.lng, -180, 180);
    if (lat === null || lng === null) return;

    const marker = L.circleMarker([lat, lng], {
      radius: state.selectedId === item.id ? 10 : 8,
      color: '#ffffff',
      weight: 2,
      fillColor: badgeColorForItem(item),
      fillOpacity: 0.95
    });
    marker.bindTooltip(item.name, {
      permanent: true,
      direction: 'top',
      offset: [0, -8],
      className: 'people-map-marker-label'
    });
    marker.on('click', () => selectPerson(item.id, { focus: false }));
    marker.addTo(markerLayer);
    state.markersById.set(item.id, marker);
  });
}

function focusSelectedPerson() {
  const marker = state.markersById.get(state.selectedId);
  if (!marker || !map) return;
  map.setView(marker.getLatLng(), Math.max(map.getZoom(), 16));
  marker.openTooltip();
}

function selectPerson(personId, options = {}) {
  state.selectedId = safeText(personId);
  state.placingMode = false;
  renderPeopleList();
  renderMarkers();
  renderDetail();
  if (options.focus) {
    focusSelectedPerson();
  }
}

function initializeMap() {
  const churchLat = parseCoordinate(boot.church?.lat, -90, 90);
  const churchLng = parseCoordinate(boot.church?.lng, -180, 180);
  const firstMapped = state.items.find((item) => parseCoordinate(item.lat, -90, 90) !== null && parseCoordinate(item.lng, -180, 180) !== null);
  const initialLat = churchLat ?? parseCoordinate(firstMapped?.lat, -90, 90) ?? -8.7619;
  const initialLng = churchLng ?? parseCoordinate(firstMapped?.lng, -180, 180) ?? -63.9039;

  map = L.map('peopleMapCanvas', {
    rotate: true,
    touchRotate: true,
    rotateControl: false,
    bearing: state.bearing
  }).setView([initialLat, initialLng], 15);

  applyMapStyle(state.mapStyle);

  markerLayer = L.layerGroup().addTo(map);

  if (churchLat !== null && churchLng !== null) {
    churchMarker = L.circleMarker([churchLat, churchLng], {
      radius: 7,
      color: '#0f172a',
      fillColor: '#10b981',
      fillOpacity: 0.9,
      weight: 2
    })
      .bindTooltip(boot.church?.name || 'Church Base', { permanent: true, direction: 'right', className: 'people-map-marker-label' })
      .addTo(map);
  }

  map.on('click', async (event) => {
    if (!state.placingMode || !state.selectedId) return;
    const item = state.items.find((entry) => entry.id === state.selectedId);
    if (!item) return;
    try {
      const payload = { lat: event.latlng.lat, lng: event.latlng.lng };
      const response = await api(`/api/people-map/people/${item.id}/location`, {
        method: 'POST',
        body: JSON.stringify(payload)
      });
      item.lat = response.person?.lat || item.lat;
      item.lng = response.person?.lng || item.lng;
      state.placingMode = false;
      renderPeopleList();
      renderMarkers();
      renderDetail();
      focusSelectedPerson();
      statusText(`Pin updated for "${item.name}".`);
    } catch (error) {
      window.alert(error.message || 'Could not update map location.');
    }
  });
}

async function searchLocation() {
  const input = document.getElementById('peopleMapLocationSearch');
  const query = safeText(input?.value);
  if (!query || !map) return;

  const url = new URL('https://nominatim.openstreetmap.org/search');
  url.searchParams.set('q', query);
  url.searchParams.set('format', 'jsonv2');
  url.searchParams.set('limit', '1');

  const response = await fetch(url.toString(), {
    headers: { Accept: 'application/json' }
  });
  if (!response.ok) throw new Error('Could not search location right now.');
  const rows = await response.json();
  if (!Array.isArray(rows) || !rows.length) throw new Error('Location not found.');

  const lat = Number.parseFloat(rows[0].lat);
  const lng = Number.parseFloat(rows[0].lon);
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) throw new Error('Location not found.');

  map.setView([lat, lng], 17);
  if (searchMarker) {
    map.removeLayer(searchMarker);
  }
  searchMarker = L.marker([lat, lng]).addTo(map);
}

function wireUi() {
  document.getElementById('peopleMapStandardBtn')?.addEventListener('click', () => {
    applyMapStyle('standard');
  });

  document.getElementById('peopleMapSatelliteBtn')?.addEventListener('click', () => {
    applyMapStyle('satellite');
  });

  document.getElementById('peopleMapRotateLeftBtn')?.addEventListener('click', () => {
    rotateMapBy(-15);
  });

  document.getElementById('peopleMapRotateResetBtn')?.addEventListener('click', () => {
    setMapBearing(0);
  });

  document.getElementById('peopleMapRotateRightBtn')?.addEventListener('click', () => {
    rotateMapBy(15);
  });

  document.getElementById('peopleMapSearchInput')?.addEventListener('input', (event) => {
    state.searchText = safeText(event.target.value);
    renderPeopleList();
    renderMarkers();
  });

  document.getElementById('peopleMapPlacePinBtn')?.addEventListener('click', () => {
    if (!state.selectedId) return;
    state.placingMode = !state.placingMode;
    renderDetail();
    statusText(
      state.placingMode
        ? 'Click the map where this person should be pinned.'
        : 'Select a person to focus their pin, move it, or log a visit from the map.'
    );
  });

  document.getElementById('peopleMapClearSelectionBtn')?.addEventListener('click', () => {
    state.selectedId = '';
    state.placingMode = false;
    renderPeopleList();
    renderMarkers();
    renderDetail();
    statusText('Selection cleared.');
  });

  document.getElementById('peopleMapSearchBtn')?.addEventListener('click', async () => {
    try {
      await searchLocation();
    } catch (error) {
      window.alert(error.message || 'Could not search location.');
    }
  });

  document.getElementById('peopleMapLocationSearch')?.addEventListener('keydown', async (event) => {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    try {
      await searchLocation();
    } catch (error) {
      window.alert(error.message || 'Could not search location.');
    }
  });
}

function init() {
  state.items = state.items.map((item) => normalizeItem(item));
  initializeMap();
  renderChipGroup('peopleMapMembershipFilters', boot.filters?.membershipTypes || [], state.membershipFilter, (value) => {
    state.membershipFilter = value;
    renderPeopleList();
    renderMarkers();
  });
  renderChipGroup('peopleMapTagFilters', boot.filters?.tags || [], state.tagFilter, (value) => {
    state.tagFilter = value;
    renderPeopleList();
    renderMarkers();
  });
  wireUi();
  renderPeopleList();
  renderMarkers();
  renderDetail();

  if (state.selectedId) {
    focusSelectedPerson();
  }
}

document.addEventListener('DOMContentLoaded', init);
