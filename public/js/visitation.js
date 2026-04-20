let map;
let drawnItems;
let savedSectionsLayer;
let activeLayer = null;
let locationMarker = null;
let didFitBounds = false;

const boot = window.__VISITATION_BOOTSTRAP__ || { sections: [], folders: [], mapSettings: {}, mapProfiles: [] };
const state = {
  sections: Array.isArray(boot.sections) ? boot.sections : [],
  folders: Array.isArray(boot.folders) ? boot.folders : [],
  mapProfiles: Array.isArray(boot.mapProfiles) ? boot.mapProfiles : [],
  mapSettings: boot.mapSettings && typeof boot.mapSettings === 'object' ? boot.mapSettings : {},
  selectedFolderId: '',
  selectedSectionId: '',
  mode: 'map',
  currentLocation: null,
  hasConfiguredMapBase: false
};

const statusLabels = {
  unclaimed: 'Unclaimed',
  claimed: 'Claimed',
  completed: 'Completed'
};

async function api(path, options = {}) {
  const res = await fetch(path, {
    headers: { 'Content-Type': 'application/json' },
    ...options
  });

  if (!res.ok) {
    let message = 'API request failed';
    try {
      const payload = await res.json();
      if (payload?.error) {
        message = payload.error;
      }
    } catch {
      // no-op
    }
    throw new Error(message);
  }

  return res.json();
}

function safeText(value) {
  return (value || '').toString();
}

function normalizeSection(section) {
  const status = ['unclaimed', 'claimed', 'completed'].includes(section?.status)
    ? section.status
    : 'unclaimed';

  return {
    id: safeText(section?.id),
    name: safeText(section?.name),
    folderId: safeText(section?.folderId),
    color: safeText(section?.color) || '#0c4a6e',
    lastVisited: safeText(section?.lastVisited),
    notes: safeText(section?.notes),
    geojson: section?.geojson || null,
    status,
    claimedBy: status === 'unclaimed' ? '' : safeText(section?.claimedBy),
    claimedAt: status === 'unclaimed' ? '' : safeText(section?.claimedAt),
    completedAt: status === 'completed' ? safeText(section?.completedAt) : '',
    checklist: Array.isArray(section?.checklist)
      ? section.checklist
          .map((item, index) => {
            const label = safeText(item?.label);
            if (!label) return null;
            const done = Boolean(item?.done);
            return {
              id: safeText(item?.id) || `${safeText(section?.id) || 'section'}-item-${index + 1}`,
              label,
              done,
              completedAt: done ? safeText(item?.completedAt) : ''
            };
          })
          .filter(Boolean)
      : []
  };
}

function normalizeFolder(folder) {
  return {
    id: safeText(folder?.id),
    name: safeText(folder?.name),
    color: safeText(folder?.color) || '#20c997',
    notes: safeText(folder?.notes)
  };
}

function parseCoordinate(value, min, max) {
  const parsed = Number.parseFloat(safeText(value).replace(',', '.'));
  if (!Number.isFinite(parsed)) return null;
  if (parsed < min || parsed > max) return null;
  return parsed;
}

function normalizeMapProfile(profile) {
  return {
    id: safeText(profile?.id),
    name: safeText(profile?.name),
    lat: safeText(profile?.lat),
    lng: safeText(profile?.lng),
    address: safeText(profile?.address)
  };
}

function normalizeMapSettings(settings) {
  const mode = safeText(settings?.mapCenterMode) === 'profile' ? 'profile' : 'church';
  const zoomRaw = Number.parseInt(safeText(settings?.mapCenterZoom), 10);

  return {
    mapCenterMode: mode,
    mapCenterZoom: Number.isInteger(zoomRaw) ? Math.min(20, Math.max(3, zoomRaw)) : 13,
    profilePersonId: safeText(settings?.profilePersonId),
    churchProfile: {
      name: safeText(settings?.churchProfile?.name),
      address: safeText(settings?.churchProfile?.address),
      lat: safeText(settings?.churchProfile?.lat),
      lng: safeText(settings?.churchProfile?.lng)
    }
  };
}

function resolveConfiguredMapBase() {
  const settings = normalizeMapSettings(state.mapSettings);
  const zoom = settings.mapCenterZoom || 13;

  if (settings.mapCenterMode === 'profile') {
    const profile = state.mapProfiles.find((entry) => entry.id === settings.profilePersonId);
    const lat = parseCoordinate(profile?.lat, -90, 90);
    const lng = parseCoordinate(profile?.lng, -180, 180);
    if (lat !== null && lng !== null) {
      return { lat, lng, zoom, hasConfiguredMapBase: true };
    }
  }

  const churchLat = parseCoordinate(settings.churchProfile.lat, -90, 90);
  const churchLng = parseCoordinate(settings.churchProfile.lng, -180, 180);
  if (churchLat !== null && churchLng !== null) {
    return { lat: churchLat, lng: churchLng, zoom, hasConfiguredMapBase: true };
  }

  return { lat: 47.6062, lng: -122.3321, zoom: 11, hasConfiguredMapBase: false };
}

function statusBadgeClass(status) {
  if (status === 'claimed') return 'bg-yellow-lt text-yellow';
  if (status === 'completed') return 'bg-green-lt text-green';
  return 'bg-blue-lt text-blue';
}

function checklistToLines(checklist = []) {
  return checklist.map((item) => `${item.done ? '[x]' : '[ ]'} ${item.label}`).join('\n');
}

function parseChecklistInput(rawText, sectionId) {
  return (rawText || '')
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line, index) => {
      const checkedMatch = line.match(/^\[(x|X|\s)\]\s*(.+)$/);
      const done = Boolean(checkedMatch && checkedMatch[1].toLowerCase() === 'x');
      const label = checkedMatch ? checkedMatch[2].trim() : line;
      return {
        id: `${sectionId || 'new'}-item-${index + 1}`,
        label,
        done,
        completedAt: done ? new Date().toISOString() : ''
      };
    })
    .filter((item) => item.label.length > 0);
}

function getFolderName(folderId) {
  if (!folderId) return 'Unassigned';
  const folder = state.folders.find((entry) => entry.id === folderId);
  return folder ? folder.name : 'Unassigned';
}

function getFilteredSections() {
  if (!state.selectedFolderId) {
    return state.sections;
  }

  return state.sections.filter((section) => section.folderId === state.selectedFolderId);
}

function setForm(section) {
  document.getElementById('sectionId').value = section?.id || '';
  document.getElementById('sectionName').value = section?.name || '';
  document.getElementById('sectionFolder').value = section?.folderId || '';
  document.getElementById('sectionColor').value = section?.color || '#0c4a6e';
  document.getElementById('sectionStatus').value = section?.status || 'unclaimed';
  document.getElementById('sectionDate').value = section?.lastVisited || '';
  document.getElementById('sectionClaimedBy').value = section?.claimedBy || '';
  document.getElementById('sectionChecklist').value = checklistToLines(section?.checklist || []);
  document.getElementById('sectionNotes').value = section?.notes || '';
}

function clearDrawing() {
  if (!drawnItems) return;

  if (activeLayer) {
    drawnItems.removeLayer(activeLayer);
    activeLayer = null;
  }
}

function clearFormAndSelection() {
  state.selectedSectionId = '';
  setForm(null);
  clearDrawing();
  renderBoard();
}

function setMode(mode) {
  state.mode = mode === 'list' ? 'list' : 'map';

  document.querySelectorAll('[data-mode]').forEach((button) => {
    const isActive = button.getAttribute('data-mode') === state.mode;
    button.classList.toggle('btn-primary', isActive);
    button.classList.toggle('btn-outline-primary', !isActive);
  });

  document.getElementById('mapWrap').classList.toggle('d-none', state.mode !== 'map');
  document.getElementById('listWrap').classList.toggle('d-none', state.mode !== 'list');

  if (state.mode === 'map' && map) {
    setTimeout(() => map.invalidateSize(), 120);
  }
}

function setActiveFolder(folderId) {
  state.selectedFolderId = folderId || '';
  renderBoard();
}

function renderFolderSelectOptions() {
  const select = document.getElementById('sectionFolder');
  if (!select) return;

  const currentValue = select.value;
  select.innerHTML = '<option value="">Unassigned</option>';

  state.folders.forEach((folder) => {
    const option = document.createElement('option');
    option.value = folder.id;
    option.textContent = folder.name;
    select.appendChild(option);
  });

  select.value = state.folders.some((folder) => folder.id === currentValue) ? currentValue : '';
}

function renderMapSettingsForm() {
  const modeInput = document.getElementById('mapCenterMode');
  if (!modeInput) return;

  const settings = normalizeMapSettings(state.mapSettings);
  const profileSelect = document.getElementById('mapProfilePersonId');
  const churchNameInput = document.getElementById('churchName');
  const churchAddressInput = document.getElementById('churchAddress');
  const churchLatInput = document.getElementById('churchLat');
  const churchLngInput = document.getElementById('churchLng');
  const zoomInput = document.getElementById('mapCenterZoom');

  modeInput.value = settings.mapCenterMode;
  zoomInput.value = settings.mapCenterZoom || 13;
  churchNameInput.value = settings.churchProfile.name || '';
  churchAddressInput.value = settings.churchProfile.address || '';
  churchLatInput.value = settings.churchProfile.lat || '';
  churchLngInput.value = settings.churchProfile.lng || '';

  const currentProfileId = settings.profilePersonId;
  profileSelect.innerHTML = '<option value="">Select profile</option>';
  state.mapProfiles.forEach((profile) => {
    const option = document.createElement('option');
    option.value = profile.id;
    option.textContent = `${profile.name}${profile.lat && profile.lng ? '' : ' (no coords)'}`;
    profileSelect.appendChild(option);
  });
  profileSelect.value = state.mapProfiles.some((entry) => entry.id === currentProfileId) ? currentProfileId : '';
  profileSelect.disabled = settings.mapCenterMode !== 'profile';
}

function renderFolderSidebar() {
  const root = document.getElementById('foldersList');
  root.innerHTML = '';

  const allButton = document.getElementById('showAllFolders');
  allButton.classList.toggle('btn-secondary', !state.selectedFolderId);
  allButton.classList.toggle('btn-outline-secondary', Boolean(state.selectedFolderId));

  if (!state.folders.length) {
    root.innerHTML = '<p class="text-secondary mb-0">No folders yet.</p>';
    return;
  }

  state.folders.forEach((folder) => {
    const row = document.createElement('div');
    row.className = `folder-row ${state.selectedFolderId === folder.id ? 'active' : ''}`;
    row.innerHTML = `
      <button type="button" class="btn btn-sm btn-ghost-secondary folder-pick">
        <span class="folder-dot" style="background:${folder.color};"></span>${folder.name}
      </button>
      <button type="button" class="btn btn-sm btn-outline-danger folder-delete">
        <i class="bi bi-trash"></i>
      </button>
    `;

    row.querySelector('.folder-pick').addEventListener('click', () => {
      setActiveFolder(folder.id);
    });

    row.querySelector('.folder-delete').addEventListener('click', async () => {
      const ok = window.confirm(`Delete folder "${folder.name}"? Sections will stay unassigned.`);
      if (!ok) return;
      await api(`/api/folders/${folder.id}`, { method: 'DELETE' });
      await refreshData();
      if (state.selectedFolderId === folder.id) {
        setActiveFolder('');
      } else {
        renderBoard();
      }
    });

    root.appendChild(row);
  });
}

function getSectionStyle(section) {
  const isSelected = state.selectedSectionId === section.id;
  const isCompleted = section.status === 'completed';
  const isClaimed = section.status === 'claimed';

  return {
    color: section.color || '#0c4a6e',
    weight: isSelected ? 4 : 2,
    fillOpacity: isCompleted ? 0.08 : isClaimed ? 0.16 : 0.24,
    dashArray: isCompleted ? '8 6' : ''
  };
}

function renderMapSections() {
  if (!savedSectionsLayer) return;

  savedSectionsLayer.clearLayers();

  const filtered = getFilteredSections();
  filtered.forEach((section) => {
    if (!section.geojson) return;

    const layer = L.geoJSON(section.geojson, {
      style: getSectionStyle(section)
    });

    layer.eachLayer((child) => {
      child.on('click', () => focusSection(section.id, { zoom: false }));
      child.bindTooltip(`${section.name} (${statusLabels[section.status]})`, { sticky: true });
    });

    savedSectionsLayer.addLayer(layer);
  });

  if (!didFitBounds && !state.hasConfiguredMapBase && filtered.length && savedSectionsLayer.getLayers().length) {
    try {
      map.fitBounds(savedSectionsLayer.getBounds(), { padding: [24, 24], maxZoom: 15 });
      didFitBounds = true;
    } catch {
      // no-op
    }
  }
}

function renderSectionsList() {
  const root = document.getElementById('sectionsList');
  root.innerHTML = '';

  const filtered = getFilteredSections();
  if (!filtered.length) {
    root.innerHTML = '<p class="text-secondary mb-0">No sections in this view yet.</p>';
    return;
  }

  filtered.forEach((section) => {
    const card = document.createElement('div');
    card.className = `territory-card ${state.selectedSectionId === section.id ? 'active' : ''}`;

    const checklistHtml = section.checklist.length
      ? section.checklist
          .map(
            (item) => `
              <label class="form-check territory-check-item">
                <input class="form-check-input checklist-toggle" type="checkbox" data-item-id="${item.id}" ${
                  item.done ? 'checked' : ''
                } />
                <span class="form-check-label">${item.label}</span>
              </label>
            `
          )
          .join('')
      : '<p class="text-secondary mb-0">No checklist items yet.</p>';

    card.innerHTML = `
      <div class="d-flex justify-content-between align-items-start flex-wrap gap-2">
        <div>
          <h4 class="mb-1">${section.name}</h4>
          <p class="text-secondary mb-0">
            ${getFolderName(section.folderId)} · Last visited: ${section.lastVisited || 'n/a'}
          </p>
          ${
            section.claimedBy
              ? `<p class="text-secondary mb-0">Claimed by: ${section.claimedBy}</p>`
              : '<p class="text-secondary mb-0">Unclaimed</p>'
          }
        </div>
        <span class="badge ${statusBadgeClass(section.status)}">${statusLabels[section.status]}</span>
      </div>
      <div class="territory-checklist mt-2">${checklistHtml}</div>
      <div class="d-flex gap-2 flex-wrap mt-3">
        <button type="button" class="btn btn-outline-primary btn-sm action-focus">Focus</button>
        <button type="button" class="btn btn-outline-secondary btn-sm action-edit">Edit</button>
        ${
          section.status === 'unclaimed'
            ? '<button type="button" class="btn btn-warning btn-sm action-claim">Claim</button>'
            : '<button type="button" class="btn btn-outline-warning btn-sm action-unclaim">Unclaim</button>'
        }
        ${
          section.status === 'completed'
            ? ''
            : '<button type="button" class="btn btn-success btn-sm action-complete">Complete</button>'
        }
        <button type="button" class="btn btn-outline-danger btn-sm action-delete">Delete</button>
      </div>
    `;

    card.querySelector('.action-focus').addEventListener('click', () => focusSection(section.id, { zoom: true }));
    card.querySelector('.action-edit').addEventListener('click', () => startEditingSection(section.id));

    const claimButton = card.querySelector('.action-claim');
    if (claimButton) {
      claimButton.addEventListener('click', async () => {
        await claimSection(section.id);
      });
    }

    const unclaimButton = card.querySelector('.action-unclaim');
    if (unclaimButton) {
      unclaimButton.addEventListener('click', async () => {
        await unclaimSection(section.id);
      });
    }

    const completeButton = card.querySelector('.action-complete');
    if (completeButton) {
      completeButton.addEventListener('click', async () => {
        await completeSection(section.id);
      });
    }

    card.querySelector('.action-delete').addEventListener('click', async () => {
      const ok = window.confirm(`Delete "${section.name}"?`);
      if (!ok) return;
      await api(`/api/sections/${section.id}`, { method: 'DELETE' });
      await refreshData();
      if (state.selectedSectionId === section.id) {
        clearFormAndSelection();
      }
    });

    card.querySelectorAll('.checklist-toggle').forEach((input) => {
      input.addEventListener('change', async (event) => {
        await api(`/api/sections/${section.id}/checklist`, {
          method: 'POST',
          body: JSON.stringify({
            itemId: event.target.getAttribute('data-item-id'),
            done: event.target.checked
          })
        });
        await refreshData();
      });
    });

    root.appendChild(card);
  });
}

function renderActiveFolderLabel() {
  const badge = document.getElementById('activeFolderLabel');
  badge.textContent = state.selectedFolderId ? getFolderName(state.selectedFolderId) : 'All Folders';
}

function renderBoard() {
  renderMapSettingsForm();
  renderFolderSelectOptions();
  renderFolderSidebar();
  renderActiveFolderLabel();
  renderMapSections();
  renderSectionsList();
}

function sectionLayerBounds(section) {
  if (!section?.geojson) return null;
  try {
    const layer = L.geoJSON(section.geojson);
    return layer.getBounds();
  } catch {
    return null;
  }
}

function focusSection(sectionId, options = {}) {
  state.selectedSectionId = sectionId;
  renderBoard();

  if (!map || options.zoom === false) return;

  const section = state.sections.find((entry) => entry.id === sectionId);
  const bounds = sectionLayerBounds(section);
  if (bounds) {
    map.fitBounds(bounds, { padding: [24, 24], maxZoom: 17 });
  }
}

function startEditingSection(sectionId) {
  const section = state.sections.find((entry) => entry.id === sectionId);
  if (!section) return;

  state.selectedSectionId = sectionId;
  setForm(section);

  clearDrawing();

  if (section.geojson) {
    const layer = L.geoJSON(section.geojson);
    layer.eachLayer((child) => {
      drawnItems.addLayer(child);
      activeLayer = child;
    });
  }

  focusSection(sectionId, { zoom: true });
  setMode('map');
}

async function refreshData() {
  const [sections, folders, visitationSettings] = await Promise.all([
    api('/api/sections'),
    api('/api/folders'),
    api('/api/visitation/settings')
  ]);
  state.sections = (sections || []).map((entry) => normalizeSection(entry));
  state.folders = (folders || []).map((entry) => normalizeFolder(entry));
  state.mapProfiles = (visitationSettings?.mapProfiles || []).map((entry) => normalizeMapProfile(entry));
  state.mapSettings = normalizeMapSettings(visitationSettings?.mapSettings || {});

  if (
    state.selectedFolderId &&
    !state.folders.some((entry) => entry.id === state.selectedFolderId)
  ) {
    state.selectedFolderId = '';
  }

  if (
    state.selectedSectionId &&
    !state.sections.some((entry) => entry.id === state.selectedSectionId)
  ) {
    state.selectedSectionId = '';
  }

  renderBoard();
}

function extractGeometry(section) {
  const raw = section?.geojson;
  if (!raw) return null;
  if (raw.type === 'Feature') return raw.geometry || null;
  return raw;
}

function extractPolygons(section) {
  const geometry = extractGeometry(section);
  if (!geometry) return [];

  if (geometry.type === 'Polygon') {
    return [geometry.coordinates];
  }

  if (geometry.type === 'MultiPolygon') {
    return geometry.coordinates;
  }

  return [];
}

function pointInRing(point, ring) {
  let inside = false;
  const [px, py] = point;

  for (let i = 0, j = ring.length - 1; i < ring.length; j = i, i += 1) {
    const [xi, yi] = ring[i];
    const [xj, yj] = ring[j];

    const intersect =
      yi > py !== yj > py &&
      px < ((xj - xi) * (py - yi)) / (yj - yi + Number.EPSILON) + xi;

    if (intersect) {
      inside = !inside;
    }
  }

  return inside;
}

function sectionContainsPoint(section, lat, lng) {
  const point = [lng, lat];
  const polygons = extractPolygons(section);

  return polygons.some((polygon) => {
    const [outerRing, ...holes] = polygon;
    if (!outerRing || !outerRing.length) return false;
    if (!pointInRing(point, outerRing)) return false;

    return !holes.some((hole) => pointInRing(point, hole));
  });
}

function sectionCentroid(section) {
  const polygons = extractPolygons(section);
  if (!polygons.length) return null;

  const ring = polygons[0][0];
  if (!ring || !ring.length) return null;

  const total = ring.reduce(
    (acc, [lng, lat]) => {
      return { lng: acc.lng + lng, lat: acc.lat + lat };
    },
    { lng: 0, lat: 0 }
  );

  return {
    lat: total.lat / ring.length,
    lng: total.lng / ring.length
  };
}

function haversineMeters(aLat, aLng, bLat, bLng) {
  const toRad = (value) => (value * Math.PI) / 180;
  const earthRadius = 6371000;
  const dLat = toRad(bLat - aLat);
  const dLng = toRad(bLng - aLng);
  const x =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(aLat)) * Math.cos(toRad(bLat)) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  const y = 2 * Math.atan2(Math.sqrt(x), Math.sqrt(1 - x));
  return earthRadius * y;
}

function findBestSection(lat, lng) {
  const candidates = getFilteredSections().filter((section) => Boolean(section.geojson));
  if (!candidates.length) return null;

  const containing = candidates.find((section) => sectionContainsPoint(section, lat, lng));
  if (containing) {
    return {
      section: containing,
      inside: true,
      distanceMeters: 0
    };
  }

  const nearest = candidates.reduce((best, section) => {
    const centroid = sectionCentroid(section);
    if (!centroid) return best;

    const distanceMeters = haversineMeters(lat, lng, centroid.lat, centroid.lng);
    if (!best || distanceMeters < best.distanceMeters) {
      return { section, distanceMeters, inside: false };
    }
    return best;
  }, null);

  return nearest;
}

function workerNameFromUi() {
  const fromQuickClaim = safeText(document.getElementById('claimWorkerName').value).trim();
  if (fromQuickClaim) return fromQuickClaim;
  const fromForm = safeText(document.getElementById('sectionClaimedBy').value).trim();
  return fromForm;
}

async function claimSection(sectionId) {
  let claimedBy = workerNameFromUi();
  if (!claimedBy) {
    claimedBy = window.prompt('Who is claiming this section?') || '';
    claimedBy = claimedBy.trim();
  }

  if (!claimedBy) {
    window.alert('Worker name is required to claim a section.');
    return;
  }

  await api(`/api/sections/${sectionId}/claim`, {
    method: 'POST',
    body: JSON.stringify({ claimedBy })
  });

  await refreshData();
  focusSection(sectionId, { zoom: false });
}

async function unclaimSection(sectionId) {
  await api(`/api/sections/${sectionId}/unclaim`, { method: 'POST', body: '{}' });
  await refreshData();
  focusSection(sectionId, { zoom: false });
}

async function completeSection(sectionId) {
  const claimedBy = workerNameFromUi();
  await api(`/api/sections/${sectionId}/complete`, {
    method: 'POST',
    body: JSON.stringify({ claimedBy })
  });
  await refreshData();
  focusSection(sectionId, { zoom: false });
}

function requestCurrentPosition() {
  return new Promise((resolve, reject) => {
    if (!navigator.geolocation) {
      reject(new Error('Geolocation is not supported on this device.'));
      return;
    }

    navigator.geolocation.getCurrentPosition(
      (position) => resolve(position.coords),
      (error) => reject(error),
      { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }
    );
  });
}

function setLocationHint(text) {
  document.getElementById('locationHint').textContent = text;
}

async function locateUser() {
  const coords = await requestCurrentPosition();
  state.currentLocation = { lat: coords.latitude, lng: coords.longitude };

  if (map) {
    map.setView([coords.latitude, coords.longitude], 16);

    if (locationMarker) {
      map.removeLayer(locationMarker);
    }

    locationMarker = L.marker([coords.latitude, coords.longitude]).addTo(map);
    locationMarker.bindPopup('Your location').openPopup();
  }

  const match = findBestSection(coords.latitude, coords.longitude);
  if (match) {
    if (match.inside) {
      setLocationHint(`You are inside "${match.section.name}".`);
    } else {
      setLocationHint(
        `Nearest section: "${match.section.name}" (${Math.round(match.distanceMeters)}m away).`
      );
    }
    focusSection(match.section.id, { zoom: false });
  } else {
    setLocationHint('Location found, but no mapped sections are available in this filter.');
  }

  return match;
}

async function claimNearestSection() {
  let match = null;

  if (!state.currentLocation) {
    match = await locateUser();
  } else {
    match = findBestSection(state.currentLocation.lat, state.currentLocation.lng);
  }

  if (!match?.section) {
    window.alert('No section available to claim.');
    return;
  }

  await claimSection(match.section.id);
  setLocationHint(`Claimed "${match.section.name}".`);
}

function initializeMap() {
  if (map) return;

  const initialView = resolveConfiguredMapBase();
  state.hasConfiguredMapBase = Boolean(initialView.hasConfiguredMapBase);
  map = L.map('map').setView([initialView.lat, initialView.lng], initialView.zoom);

  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19,
    attribution: '&copy; OpenStreetMap contributors'
  }).addTo(map);

  drawnItems = new L.FeatureGroup();
  savedSectionsLayer = new L.FeatureGroup();
  map.addLayer(savedSectionsLayer);
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

function applyConfiguredMapBaseToMap() {
  const view = resolveConfiguredMapBase();
  state.hasConfiguredMapBase = Boolean(view.hasConfiguredMapBase);
  didFitBounds = state.hasConfiguredMapBase;

  if (!map) return;
  map.setView([view.lat, view.lng], view.zoom);
}

async function saveSection(event) {
  event.preventDefault();

  const sectionId = safeText(document.getElementById('sectionId').value).trim() || undefined;
  const existing = state.sections.find((entry) => entry.id === sectionId);
  const geometry = activeLayer ? activeLayer.toGeoJSON() : existing?.geojson;

  if (!geometry) {
    window.alert('Draw a section polygon on the map first.');
    return;
  }

  const payload = {
    id: sectionId,
    name: safeText(document.getElementById('sectionName').value).trim(),
    folderId: safeText(document.getElementById('sectionFolder').value).trim(),
    color: document.getElementById('sectionColor').value,
    status: document.getElementById('sectionStatus').value,
    lastVisited: document.getElementById('sectionDate').value,
    claimedBy: safeText(document.getElementById('sectionClaimedBy').value).trim(),
    notes: safeText(document.getElementById('sectionNotes').value).trim(),
    checklist: parseChecklistInput(
      document.getElementById('sectionChecklist').value,
      sectionId || 'new'
    ),
    geojson: geometry
  };

  await api('/api/sections', {
    method: 'POST',
    body: JSON.stringify(payload)
  });

  clearFormAndSelection();
  await refreshData();
}

async function init() {
  state.sections = state.sections.map((entry) => normalizeSection(entry));
  state.folders = state.folders.map((entry) => normalizeFolder(entry));
  state.mapProfiles = state.mapProfiles.map((entry) => normalizeMapProfile(entry));
  state.mapSettings = normalizeMapSettings(state.mapSettings);

  initializeMap();

  renderBoard();
  setMode('map');

  await refreshData();
}

function onError(error) {
  window.alert(error?.message || 'Something went wrong.');
}

document.addEventListener('DOMContentLoaded', async () => {
  try {
    await init();

    const mapCenterModeInput = document.getElementById('mapCenterMode');
    const mapSettingsForm = document.getElementById('mapSettingsForm');

    if (mapCenterModeInput) {
      mapCenterModeInput.addEventListener('change', (event) => {
        state.mapSettings = normalizeMapSettings({
          ...state.mapSettings,
          mapCenterMode: event.target.value
        });
        renderMapSettingsForm();
      });
    }

    if (mapSettingsForm) {
      mapSettingsForm.addEventListener('submit', async (event) => {
        try {
          event.preventDefault();

          const payload = {
            mapCenterMode: safeText(document.getElementById('mapCenterMode').value),
            profilePersonId: safeText(document.getElementById('mapProfilePersonId').value),
            mapCenterZoom: safeText(document.getElementById('mapCenterZoom').value),
            churchName: safeText(document.getElementById('churchName').value),
            churchAddress: safeText(document.getElementById('churchAddress').value),
            churchLat: safeText(document.getElementById('churchLat').value),
            churchLng: safeText(document.getElementById('churchLng').value)
          };

          const saved = await api('/api/visitation/settings', {
            method: 'POST',
            body: JSON.stringify(payload)
          });

          state.mapSettings = normalizeMapSettings(saved?.mapSettings || {});
          state.mapProfiles = (saved?.mapProfiles || []).map((entry) => normalizeMapProfile(entry));
          renderBoard();
          applyConfiguredMapBaseToMap();
        } catch (error) {
          onError(error);
        }
      });
    }

    document.getElementById('showAllFolders').addEventListener('click', () => setActiveFolder(''));

    document.getElementById('folderForm').addEventListener('submit', async (event) => {
      event.preventDefault();
      const name = safeText(document.getElementById('folderName').value).trim();
      const color = document.getElementById('folderColor').value;

      if (!name) {
        window.alert('Folder name is required.');
        return;
      }

      await api('/api/folders', {
        method: 'POST',
        body: JSON.stringify({ name, color })
      });

      document.getElementById('folderName').value = '';
      await refreshData();
    });

    document.querySelectorAll('[data-mode]').forEach((button) => {
      button.addEventListener('click', () => setMode(button.getAttribute('data-mode')));
    });

    document.getElementById('clearDrawing').addEventListener('click', () => {
      clearFormAndSelection();
    });

    document.getElementById('sectionForm').addEventListener('submit', async (event) => {
      await saveSection(event);
    });

    document.getElementById('locateMe').addEventListener('click', async () => {
      try {
        await locateUser();
      } catch (error) {
        onError(error);
      }
    });

    document.getElementById('claimNearest').addEventListener('click', async () => {
      try {
        await claimNearestSection();
      } catch (error) {
        onError(error);
      }
    });
  } catch (error) {
    onError(error);
  }
});
