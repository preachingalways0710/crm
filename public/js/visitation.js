let map;
let drawnItems;
let savedSectionsLayer;
let savedMarkersLayer;
let activeLayer = null;
let didFitBounds = false;
let draftStreetCounter = 0;
let mapBaseLayer;
let mapLabelLayer;
let mapLayerMode = 'satellite';

const boot = window.__VISITATION_BOOTSTRAP__ || {
  sections: [],
  folders: [],
  markers: [],
  mapSettings: {},
  mapProfiles: []
};
const state = {
  sections: Array.isArray(boot.sections) ? boot.sections : [],
  folders: Array.isArray(boot.folders) ? boot.folders : [],
  markers: Array.isArray(boot.markers) ? boot.markers : [],
  mapProfiles: Array.isArray(boot.mapProfiles) ? boot.mapProfiles : [],
  mapSettings: boot.mapSettings && typeof boot.mapSettings === 'object' ? boot.mapSettings : {},
  selectedFolderId: '',
  selectedSectionId: '',
  mode: 'map',
  drawingTarget: 'section',
  draftStreets: [],
  hasConfiguredMapBase: false,
  showMapNames: true,
  foldersVisible: true
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

function normalizeStreet(street, index = 0, sectionId = 'section') {
  return {
    id: safeText(street?.id) || `${sectionId}-street-${index + 1}`,
    name: safeText(street?.name) || `Street ${index + 1}`,
    color: safeText(street?.color) || '#16a34a',
    done: Boolean(street?.done),
    completedAt: street?.done ? safeText(street?.completedAt) : '',
    geojson: street?.geojson || null
  };
}

function normalizeSection(section) {
  const status = ['unclaimed', 'claimed', 'completed'].includes(section?.status)
    ? section.status
    : 'unclaimed';
  const sectionId = safeText(section?.id);

  return {
    id: sectionId,
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
      : [],
    streets: Array.isArray(section?.streets)
      ? section.streets
          .map((street, index) => normalizeStreet(street, index, sectionId || 'section'))
          .filter((street) => Boolean(street.geojson))
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

function normalizeMarker(marker, index = 0) {
  return {
    id: safeText(marker?.id) || `marker-${index + 1}`,
    name: safeText(marker?.name) || `Marker ${index + 1}`,
    notes: safeText(marker?.notes),
    color: safeText(marker?.color) || '#2563eb',
    lat: safeText(marker?.lat),
    lng: safeText(marker?.lng)
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

function buildMapBaseLayer(mode) {
  if (mode === 'map') {
    return L.tileLayer('https://{s}.basemaps.cartocdn.com/light_nolabels/{z}/{x}/{y}{r}.png', {
      maxZoom: 20,
      subdomains: 'abcd',
      attribution: '&copy; OpenStreetMap contributors &copy; CARTO'
    });
  }

  return L.tileLayer(
    'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
    {
      maxZoom: 20,
      attribution: 'Tiles &copy; Esri'
    }
  );
}

function buildMapLabelLayer() {
  return L.tileLayer('https://{s}.basemaps.cartocdn.com/light_only_labels/{z}/{x}/{y}{r}.png', {
    maxZoom: 20,
    subdomains: 'abcd',
    attribution: '&copy; OpenStreetMap contributors &copy; CARTO'
  });
}

function applyMapStyle(mode = mapLayerMode) {
  mapLayerMode = mode === 'map' ? 'map' : 'satellite';
  if (!map) return;

  if (mapBaseLayer && map.hasLayer(mapBaseLayer)) {
    map.removeLayer(mapBaseLayer);
  }
  if (mapLabelLayer && map.hasLayer(mapLabelLayer)) {
    map.removeLayer(mapLabelLayer);
  }

  mapBaseLayer = buildMapBaseLayer(mapLayerMode);
  mapBaseLayer.addTo(map);

  if (state.showMapNames) {
    mapLabelLayer = buildMapLabelLayer();
    mapLabelLayer.addTo(map);
  }

  syncTopControlState();
}

function syncTopControlState() {
  const mapBtn = document.getElementById('mapStyleMapBtn');
  const satelliteBtn = document.getElementById('mapStyleSatelliteBtn');
  const namesBtn = document.getElementById('toggleMapNamesBtn');
  const foldersBtn = document.getElementById('toggleFoldersBtn');
  if (!mapBtn || !satelliteBtn || !namesBtn || !foldersBtn) return;

  mapBtn.classList.toggle('active', mapLayerMode === 'map');
  satelliteBtn.classList.toggle('active', mapLayerMode === 'satellite');
  namesBtn.classList.toggle('active', state.showMapNames);
  foldersBtn.classList.toggle('active', state.foldersVisible);
}

function toggleFoldersPanel() {
  const sidebar = document.getElementById('visitationSidebarCol');
  const main = document.getElementById('visitationMainCol');
  if (!sidebar || !main) return;

  state.foldersVisible = !state.foldersVisible;
  sidebar.classList.toggle('d-none', !state.foldersVisible);
  main.classList.toggle('col-lg-9', state.foldersVisible);
  main.classList.toggle('col-lg-12', !state.foldersVisible);
  syncTopControlState();

  if (map) {
    setTimeout(() => map.invalidateSize(), 120);
  }
}

function startAddMarkerMode() {
  if (!map) return;
  const drawMarker = new L.Draw.Marker(map);
  drawMarker.enable();
}

function startAddMapMode() {
  state.drawingTarget = 'section';
  syncDrawingTargetInput();
  if (!map) return;
  const drawPolygon = new L.Draw.Polygon(map, {
    allowIntersection: false,
    showArea: true
  });
  drawPolygon.enable();
}

function toggleMapNames() {
  state.showMapNames = !state.showMapNames;
  applyMapStyle(mapLayerMode);
  renderBoard();
}

async function searchLocation(query) {
  const q = safeText(query).trim();
  if (!q) return;

  const url = new URL('https://nominatim.openstreetmap.org/search');
  url.searchParams.set('q', q);
  url.searchParams.set('format', 'jsonv2');
  url.searchParams.set('limit', '1');

  const response = await fetch(url.toString(), {
    headers: {
      Accept: 'application/json'
    }
  });
  if (!response.ok) {
    throw new Error('Could not search this location right now.');
  }

  const rows = await response.json();
  if (!Array.isArray(rows) || !rows.length) {
    throw new Error('Location not found.');
  }

  const first = rows[0];
  const lat = Number.parseFloat(first.lat);
  const lng = Number.parseFloat(first.lon);
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) {
    throw new Error('Location not found.');
  }

  map.setView([lat, lng], 16);
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
  document.getElementById('drawingTarget').value = 'section';
  document.getElementById('streetName').value = '';
  state.drawingTarget = 'section';
}

function clearDrawing() {
  if (!drawnItems) return;

  if (activeLayer) {
    drawnItems.removeLayer(activeLayer);
    activeLayer = null;
  }

  state.draftStreets.forEach((street) => {
    if (street.layer && drawnItems.hasLayer(street.layer)) {
      drawnItems.removeLayer(street.layer);
    }
  });
  state.draftStreets = [];
  renderStreetDraftList();
}

function clearFormAndSelection() {
  state.selectedSectionId = '';
  setForm(null);
  clearDrawing();
  renderBoard();
  updateClaimHintFromMapCenter();
}

function getStreetStyle(street, isSelectedSection) {
  const done = Boolean(street?.done);
  return {
    color: street?.color || '#16a34a',
    weight: isSelectedSection ? 3 : 2,
    fillOpacity: done ? 0.34 : 0.17,
    dashArray: done ? '' : '6 4'
  };
}

function syncDrawingTargetInput() {
  const input = document.getElementById('drawingTarget');
  if (!input) return;
  input.value = state.drawingTarget === 'street' ? 'street' : 'section';
}

function renderStreetDraftList() {
  const root = document.getElementById('streetDraftList');
  if (!root) return;

  if (!state.draftStreets.length) {
    root.innerHTML =
      '<p class="text-secondary mb-0">Switch Draw Mode to Street Layer and draw green areas inside the section.</p>';
    return;
  }

  root.innerHTML = '';
  state.draftStreets.forEach((street, index) => {
    const row = document.createElement('div');
    row.className = 'street-draft-item';
    row.innerHTML = `
      <input class="form-control form-control-sm street-draft-name" value="${street.name}" />
      <label class="form-check mb-0 ms-2">
        <input class="form-check-input street-draft-done" type="checkbox" ${street.done ? 'checked' : ''} />
        <span class="form-check-label">Done</span>
      </label>
      <button type="button" class="btn btn-outline-danger btn-sm street-draft-remove">Remove</button>
    `;

    row.querySelector('.street-draft-name').addEventListener('input', (event) => {
      const value = safeText(event.target.value).trim();
      state.draftStreets[index].name = value || `Street ${index + 1}`;
    });

    row.querySelector('.street-draft-done').addEventListener('change', (event) => {
      state.draftStreets[index].done = event.target.checked;
      if (state.draftStreets[index].layer) {
        state.draftStreets[index].layer.setStyle(getStreetStyle(state.draftStreets[index], true));
      }
    });

    row.querySelector('.street-draft-remove').addEventListener('click', () => {
      const target = state.draftStreets[index];
      if (target?.layer && drawnItems?.hasLayer(target.layer)) {
        drawnItems.removeLayer(target.layer);
      }
      state.draftStreets.splice(index, 1);
      renderStreetDraftList();
    });

    root.appendChild(row);
  });
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
  updateClaimHintFromMapCenter();
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

    const sectionLayer = L.geoJSON(section.geojson, {
      style: getSectionStyle(section)
    });

    sectionLayer.eachLayer((child) => {
      child.on('click', () => focusSection(section.id, { zoom: false }));
      child.bindTooltip(`${section.name} (${statusLabels[section.status]})`, { sticky: true });
    });

    savedSectionsLayer.addLayer(sectionLayer);

    (section.streets || []).forEach((street, index) => {
      if (!street.geojson) return;
      const streetLayer = L.geoJSON(street.geojson, {
        style: getStreetStyle(street, state.selectedSectionId === section.id)
      });

      streetLayer.eachLayer((child) => {
        child.on('click', () => focusSection(section.id, { zoom: false }));
        child.bindTooltip(`${section.name} · ${street.name || `Street ${index + 1}`}`, { sticky: true });
      });

      savedSectionsLayer.addLayer(streetLayer);
    });
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

function renderMapMarkers() {
  if (!savedMarkersLayer) return;
  savedMarkersLayer.clearLayers();

  state.markers.forEach((marker, index) => {
    const lat = parseCoordinate(marker.lat, -90, 90);
    const lng = parseCoordinate(marker.lng, -180, 180);
    if (lat === null || lng === null) return;

    const layer = L.circleMarker([lat, lng], {
      radius: 7,
      color: marker.color || '#2563eb',
      fillColor: marker.color || '#2563eb',
      fillOpacity: 0.95,
      weight: 2
    });
    const markerName = marker.name || `Marker ${index + 1}`;
    const markerNotes = marker.notes ? `<p class="mb-2">${marker.notes}</p>` : '';
    layer.bindPopup(
      `<strong>${markerName}</strong>${markerNotes}<button type="button" class="btn btn-sm btn-outline-danger marker-delete-btn" data-marker-id="${marker.id}">Delete Marker</button>`
    );

    layer.on('popupopen', (event) => {
      const popupRoot = event.popup.getElement();
      const deleteButton = popupRoot?.querySelector('.marker-delete-btn');
      if (!deleteButton) return;
      deleteButton.addEventListener('click', async () => {
        const ok = window.confirm(`Delete marker "${markerName}"?`);
        if (!ok) return;
        await api(`/api/markers/${marker.id}`, { method: 'DELETE' });
        await refreshData();
      });
    });

    if (state.showMapNames) {
      layer.bindTooltip(markerName, { permanent: false, direction: 'top', offset: [0, -6] });
    }

    savedMarkersLayer.addLayer(layer);
  });
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
    const streetsHtml = section.streets.length
      ? section.streets
          .map(
            (street) => `
              <label class="form-check territory-check-item">
                <input class="form-check-input street-toggle" type="checkbox" data-street-id="${street.id}" ${
                  street.done ? 'checked' : ''
                } />
                <span class="form-check-label">${street.name}</span>
              </label>
            `
          )
          .join('')
      : '<p class="text-secondary mb-0">No streets mapped yet.</p>';

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
      <div class="territory-checklist mt-2">
        <p class="text-secondary small mb-1">Street Layer</p>
        ${streetsHtml}
      </div>
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

    card.querySelectorAll('.street-toggle').forEach((input) => {
      input.addEventListener('change', async (event) => {
        await api(
          `/api/sections/${section.id}/streets/${encodeURIComponent(
            event.target.getAttribute('data-street-id')
          )}`,
          {
            method: 'POST',
            body: JSON.stringify({
              done: event.target.checked
            })
          }
        );
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
  renderMapMarkers();
  renderSectionsList();
  syncTopControlState();
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
  updateClaimHintFromMapCenter();

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
      child.setStyle(getSectionStyle(section));
      drawnItems.addLayer(child);
      activeLayer = child;
    });
  }

  state.draftStreets = [];
  (section.streets || []).forEach((street, index) => {
    if (!street.geojson) return;
    const group = L.geoJSON(street.geojson);
    group.eachLayer((child) => {
      child.setStyle(getStreetStyle(street, true));
      drawnItems.addLayer(child);
      state.draftStreets.push({
        ...normalizeStreet(street, index, section.id),
        layer: child
      });
    });
  });
  renderStreetDraftList();
  syncDrawingTargetInput();

  focusSection(sectionId, { zoom: true });
  setMode('map');
}

async function refreshData() {
  const [sections, folders, markers, visitationSettings] = await Promise.all([
    api('/api/sections'),
    api('/api/folders'),
    api('/api/markers'),
    api('/api/visitation/settings')
  ]);
  state.sections = (sections || []).map((entry) => normalizeSection(entry));
  state.folders = (folders || []).map((entry) => normalizeFolder(entry));
  state.markers = (markers || []).map((entry, index) => normalizeMarker(entry, index));
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
  updateClaimHintFromMapCenter();
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

function findBestSection(lat, lng, sourceSections) {
  const candidates = Array.isArray(sourceSections)
    ? sourceSections.filter((section) => Boolean(section.geojson))
    : getFilteredSections().filter((section) => Boolean(section.geojson));
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

function setLocationHint(text) {
  const hint = document.getElementById('locationHint');
  if (hint) {
    hint.textContent = text;
  }
}

function claimReferencePoint() {
  if (map) {
    const center = map.getCenter();
    return { lat: center.lat, lng: center.lng, source: 'map' };
  }

  const fallback = resolveConfiguredMapBase();
  return { lat: fallback.lat, lng: fallback.lng, source: 'church' };
}

function updateClaimHintFromMapCenter() {
  const reference = claimReferencePoint();
  const filtered = getFilteredSections().filter((section) => Boolean(section.geojson));
  const openSections = filtered.filter((section) => section.status !== 'completed');
  const source = openSections.length ? openSections : filtered;
  const match = findBestSection(reference.lat, reference.lng, source);

  if (!match?.section) {
    setLocationHint('No mapped sections available to claim in this filter.');
    return;
  }

  if (match.inside) {
    setLocationHint(`Map center is inside "${match.section.name}".`);
    return;
  }

  setLocationHint(`Nearest section from map center: "${match.section.name}" (${Math.round(match.distanceMeters)}m).`);
}

async function claimNearestSection() {
  const selected = state.sections.find((entry) => entry.id === state.selectedSectionId);
  if (selected?.id && selected.status !== 'completed') {
    await claimSection(selected.id);
    setLocationHint(`Claimed selected section "${selected.name}".`);
    return;
  }

  const reference = claimReferencePoint();
  const filtered = getFilteredSections().filter((section) => Boolean(section.geojson));
  const preferred = filtered.filter((section) => section.status === 'unclaimed');
  const source = preferred.length ? preferred : filtered;
  const match = findBestSection(reference.lat, reference.lng, source);

  if (!match?.section) {
    window.alert('No section available to claim.');
    return;
  }

  await claimSection(match.section.id);
  const distanceText = match.inside ? 'inside current map center area' : `${Math.round(match.distanceMeters)}m from map center`;
  setLocationHint(`Claimed "${match.section.name}" (${distanceText}).`);
}

function initializeMap() {
  if (map) return;

  const initialView = resolveConfiguredMapBase();
  state.hasConfiguredMapBase = Boolean(initialView.hasConfiguredMapBase);
  map = L.map('map').setView([initialView.lat, initialView.lng], initialView.zoom);

  drawnItems = new L.FeatureGroup();
  savedSectionsLayer = new L.FeatureGroup();
  savedMarkersLayer = new L.FeatureGroup();
  applyMapStyle('satellite');
  map.addLayer(savedSectionsLayer);
  map.addLayer(savedMarkersLayer);
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

  map.on(L.Draw.Event.CREATED, async (event) => {
    if (event.layerType === 'marker') {
      try {
        const latLng = event.layer.getLatLng();
        let markerName =
          window.prompt('Marker name', `Marker ${state.markers.length + 1}`) || '';
        markerName = markerName.trim() || `Marker ${state.markers.length + 1}`;
        await api('/api/markers', {
          method: 'POST',
          body: JSON.stringify({
            name: markerName,
            lat: latLng.lat,
            lng: latLng.lng
          })
        });
        await refreshData();
        setLocationHint(`Marker "${markerName}" added.`);
      } catch (error) {
        onError(error);
      }
      return;
    }

    const target = state.drawingTarget === 'street' ? 'street' : 'section';

    if (target === 'section') {
      if (activeLayer && drawnItems.hasLayer(activeLayer)) {
        drawnItems.removeLayer(activeLayer);
      }

      activeLayer = event.layer;
      const sectionColor = safeText(document.getElementById('sectionColor')?.value) || '#0c4a6e';
      activeLayer.setStyle({
        color: sectionColor,
        weight: 3,
        fillOpacity: 0.24
      });
      drawnItems.addLayer(activeLayer);
      return;
    }

    if (!activeLayer) {
      window.alert('Draw the red section outline first, then add streets inside it.');
      return;
    }

    draftStreetCounter += 1;
    const customStreetName = safeText(document.getElementById('streetName')?.value).trim();
    const streetDraft = normalizeStreet({
      id: `draft-street-${Date.now()}-${draftStreetCounter}`,
      name: customStreetName || `Street ${state.draftStreets.length + 1}`,
      done: false,
      color: '#16a34a',
      geojson: event.layer.toGeoJSON()
    });
    event.layer.setStyle(getStreetStyle(streetDraft, true));
    drawnItems.addLayer(event.layer);

    state.draftStreets.push({
      ...streetDraft,
      layer: event.layer
    });
    document.getElementById('streetName').value = '';
    renderStreetDraftList();
  });

  map.on('moveend', () => {
    updateClaimHintFromMapCenter();
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

  const streetsSource = activeLayer || state.draftStreets.length ? state.draftStreets : existing?.streets || [];
  const streets = streetsSource
    .map((street, index) => {
      const normalized = normalizeStreet(
        {
          ...street,
          geojson: street?.layer ? street.layer.toGeoJSON() : street?.geojson
        },
        index,
        sectionId || 'new'
      );
      return normalized.geojson ? normalized : null;
    })
    .filter(Boolean);

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
    geojson: geometry,
    streets
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
  state.markers = state.markers.map((entry, index) => normalizeMarker(entry, index));
  state.mapProfiles = state.mapProfiles.map((entry) => normalizeMapProfile(entry));
  state.mapSettings = normalizeMapSettings(state.mapSettings);

  initializeMap();
  setForm(null);
  syncDrawingTargetInput();
  renderStreetDraftList();

  renderBoard();
  setMode('map');

  await refreshData();
  updateClaimHintFromMapCenter();
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

    document.getElementById('drawingTarget').addEventListener('change', (event) => {
      state.drawingTarget = safeText(event.target.value) === 'street' ? 'street' : 'section';
      syncDrawingTargetInput();
    });

    document.getElementById('sectionColor').addEventListener('change', (event) => {
      if (!activeLayer) return;
      activeLayer.setStyle({
        color: safeText(event.target.value) || '#0c4a6e'
      });
    });

    document.getElementById('claimNearest').addEventListener('click', async () => {
      try {
        await claimNearestSection();
      } catch (error) {
        onError(error);
      }
    });

    document.getElementById('mapStyleMapBtn').addEventListener('click', () => {
      applyMapStyle('map');
    });

    document.getElementById('mapStyleSatelliteBtn').addEventListener('click', () => {
      applyMapStyle('satellite');
    });

    document.getElementById('toggleFoldersBtn').addEventListener('click', () => {
      toggleFoldersPanel();
    });

    document.getElementById('addMarkerBtn').addEventListener('click', () => {
      startAddMarkerMode();
    });

    document.getElementById('addMapBtn').addEventListener('click', () => {
      startAddMapMode();
    });

    document.getElementById('toggleMapNamesBtn').addEventListener('click', () => {
      toggleMapNames();
    });

    document.getElementById('mapSearchInput').addEventListener('keydown', async (event) => {
      if (event.key !== 'Enter') return;
      try {
        event.preventDefault();
        await searchLocation(event.target.value);
      } catch (error) {
        onError(error);
      }
    });
  } catch (error) {
    onError(error);
  }
});
