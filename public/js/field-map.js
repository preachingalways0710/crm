(() => {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const { zones, state: loaded } = JSON.parse($('bootstrap').textContent);
  const town = $('town');
  const housingLabels = { sparse: 'Few or no homes', none: 'No homes', 'one-two': '1–2 homes' };
  const overview = [45, 99, 1310, 976];
  let state = loaded;
  let zone,
    mode = '',
    start,
    id,
    busy = false,
    selectedMark = '';
  let view = overview.slice();
  const pointers = new Map();
  let gesture,
    dragged = false;

  function node(name, text, className) {
    const element = document.createElement(name);
    if (text !== undefined) element.textContent = text;
    if (className) element.className = className;
    return element;
  }

  function svg(parent, name, attributes, text) {
    const element = document.createElementNS('http://www.w3.org/2000/svg', name);
    for (const [key, value] of Object.entries(attributes)) element.setAttribute(key, value);
    if (text !== undefined) element.textContent = text;
    parent.append(element);
    return element;
  }

  const streetLabels = [
    [299, 247, 'Av. Mal. Rondon', 8],
    [333, 516, 'Rua José Silvestre'],
    [294, 864, 'Av. Presidente Médici'],
    [219, 1032, 'BR-364'],
    [469, 725, 'Bodega', -84],
    [575, 705, 'Recife', -86],
    [605, 945, 'João Pessoa', -86],
    [708, 699, 'Airton Senna', -86],
    [744, 700, 'Moisés Freitas', -86],
    [794, 688, 'Goiânia', -88],
    [837, 581, 'Rio Claro', -88],
    [870, 704, '13 de Maio', -88],
    [905, 673, 'Sete de Setembro', -90],
    [811, 617, 'Rua Getúlio Vargas'],
    [1007, 529, 'José Silvestre'],
    [1016, 576, 'Rua Brasília'],
    [1034, 662, 'Rua Jacy Paraná'],
    [1091, 454, 'Rua Jorge Couto Alves'],
    [1148, 877, 'Av. Presidente Médici'],
    [1143, 902, 'Rua Venezuela'],
    [1100, 1037, 'BR-364'],
    [786, 987, 'Av. Costa e Silva'],
    [459, 476, 'Rua Macapá', -84],
    [1001, 734, 'Fortaleza Abunã'],
    [815, 435, 'Rua Jorge Couto Alves'],
    [772, 310, 'Moisés', -89]
  ];
  const landmarks = [
    [771, 263, 'CREAS'],
    [671, 487, 'Cemitério'],
    [1001, 395, 'Campo de futebol'],
    [791, 733, 'Prefeitura'],
    [728, 943, 'Praça Central'],
    [428, 906, 'Hotel Itapuã'],
    [1184, 996, 'Hotel Garcia']
  ];
  for (const [x, y, name, angle = 0] of streetLabels) {
    svg($('labels'), 'text', { transform: 'translate(' + x + ' ' + y + ') rotate(' + angle + ')' }, name);
  }
  for (const [x, y, name] of landmarks) {
    svg($('labels'), 'circle', { cx: x, cy: y - 12, r: 4, fill: '#5e815f', stroke: 'white' });
    svg($('labels'), 'text', { x, y: y + 5, class: 'poi' }, name);
  }

  function message(text = '', error = false) {
    $('message').textContent = text;
    $('message').classList.toggle('error', error);
  }

  function setView(next) {
    view = next;
    town.setAttribute('viewBox', view.join(' '));
  }

  function fit() {
    if (!zone) return setView(overview.slice());
    const xs = zone.polygon.map((p) => p[0]),
      ys = zone.polygon.map((p) => p[1]);
    const x = Math.min(...xs),
      y = Math.min(...ys);
    setView([x - 22, y - 22, Math.max(...xs) - x + 44, Math.max(...ys) - y + 44]);
  }

  function zoom(factor, at = [view[0] + view[2] / 2, view[1] + view[3] / 2]) {
    const width = Math.min(2200, Math.max(65, view[2] * factor));
    const scale = width / view[2];
    setView([at[0] + (view[0] - at[0]) * scale, at[1] + (view[1] - at[1]) * scale, width, view[3] * scale]);
  }

  function mapPoint(x, y) {
    const point = new DOMPoint(x, y).matrixTransform(town.getScreenCTM().inverse());
    return [Math.round(point.x * 10) / 10, Math.round(point.y * 10) / 10];
  }

  function newId() {
    const bytes = crypto.getRandomValues(new Uint8Array(16));
    bytes[6] = (bytes[6] & 15) | 64;
    bytes[8] = (bytes[8] & 63) | 128;
    const hex = Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('');
    return [hex.slice(0, 8), hex.slice(8, 12), hex.slice(12, 16), hex.slice(16, 20), hex.slice(20)].join('-');
  }

  function edit(next = '') {
    mode = next;
    start = null;
    id = next ? newId() : '';
    $('draft').replaceChildren();
    updateTools();
  }

  function updateTools() {
    $('tools').hidden = !zone;
    $('walk').hidden = $('visit').hidden = $('housingTool').hidden = Boolean(mode);
    $('housingChoice').hidden = mode !== 'housing';
    $('cancel').hidden = !mode;
    $('instruction').hidden = !mode;
    $('instruction').textContent =
      mode === 'pin'
        ? 'Tap the house or landmark to revisit.'
        : start
          ? '2 of 2 · Tap the end point.'
          : '1 of 2 · Tap the starting point.';
    town.style.cursor = mode ? 'crosshair' : '';
  }

  function choose(id, move = true) {
    zone = zones.find((item) => item.id === id);
    selectedMark = '';
    message();
    edit();
    fit();
    render();
    const url = new URL(location.href);
    url.searchParams.delete('pin');
    if (zone) url.searchParams.set('zone', zone.id);
    else url.searchParams.delete('zone');
    history.replaceState(null, '', url);
    if (move && window.matchMedia('(max-width:720px)').matches) {
      document.querySelector('.map-card').scrollIntoView({ block: 'start' });
    }
  }

  function openPins(z) {
    return state.pins.filter((p) => p.zoneId === z && p.status !== 'completed');
  }
  function stretches(z) {
    return state.stretches.filter((s) => s.zoneId === z);
  }
  function count(n, singular, plural) {
    return n + ' ' + (n === 1 ? singular : plural);
  }
  function dateLabel(value) {
    return new Date(value.length === 10 ? value + 'T12:00:00' : value).toLocaleDateString(undefined, {
      month: 'short',
      day: 'numeric',
      year: 'numeric'
    });
  }

  function showMark(mark) {
    if (busy || mode) return;
    if (zone?.id !== mark.zoneId) choose(mark.zoneId, false);
    selectedMark = mark.id;
    render();
    if (mark.housing) $('housingPanel').open = true;
    $('item-' + mark.id)?.scrollIntoView({ block: 'nearest' });
  }

  function drawMarks() {
    const root = $('marks');
    root.replaceChildren();
    const scale = Math.min(1.7, Math.max(0.34, 1 / town.getScreenCTM().a));
    for (const s of [...state.stretches, ...(state.housing || [])]) {
      const group = svg(root, 'g', { 'data-mark': s.id, opacity: !zone || s.zoneId === zone.id ? 1 : 0.45 });
      const attrs = {
        x1: s.start[0],
        y1: s.start[1],
        x2: s.end[0],
        y2: s.end[1],
        'stroke-linecap': 'round',
        'vector-effect': 'non-scaling-stroke'
      };
      if (s.housing) {
        // Leave gaps so walked coverage underneath stays visible.
        svg(group, 'line', {
          ...attrs,
          stroke: s.id === selectedMark ? '#801722' : '#b1353d',
          'stroke-width': 3,
          'stroke-dasharray': s.housing === 'none' ? '1 6' : '7 6'
        });
        svg(group, 'title', {}, housingLabels[s.housing]);
      } else {
        svg(group, 'line', { ...attrs, stroke: 'white', 'stroke-width': 8 });
        svg(group, 'line', {
          ...attrs,
          stroke: s.id === selectedMark ? '#184632' : '#26855e',
          'stroke-width': 5
        });
      }
      svg(group, 'line', { ...attrs, stroke: 'transparent', 'stroke-width': 18 });
    }
    for (const pin of state.pins) {
      if (pin.status === 'completed' && pin.id !== selectedMark) continue;
      const group = svg(root, 'g', {
        'data-mark': pin.id,
        transform: 'translate(' + pin.start.join(' ') + ') scale(' + scale + ')'
      });
      svg(group, 'path', {
        d: 'M0 0C-2-4-9-10-9-16a9 9 0 1 1 18 0C9-10 2-4 0 0Z',
        fill: pin.status === 'completed' ? '#53695a' : '#ae4b1b',
        stroke: 'white',
        'stroke-width': 1.6
      });
      svg(group, 'circle', { cx: 0, cy: -16, r: 3, fill: 'white' });
      svg(group, 'title', {}, pin.title);
    }
  }

  function renderEntries() {
    const pins = state.pins
      .filter((p) => p.zoneId === zone.id)
      .sort(
        (a, b) =>
          Number(a.status === 'completed') - Number(b.status === 'completed') ||
          (a.dueDate || '9999').localeCompare(b.dueDate || '9999')
      );
    const list = $('returnList');
    list.replaceChildren();
    if (!pins.length) list.append(node('p', 'No return visits pinned.', 'empty'));
    for (const pin of pins) {
      const row = node('div', undefined, 'entry');
      row.id = 'item-' + pin.id;
      row.classList.toggle('highlight', pin.id === selectedMark);
      row.append(
        node('strong', pin.title),
        node(
          'p',
          pin.status === 'completed'
            ? 'Completed'
            : pin.dueDate
              ? 'Return ' + dateLabel(pin.dueDate)
              : 'Open visit',
          'item-meta'
        )
      );
      if (pin.notes) row.append(node('p', pin.notes));
      const form = node('form');
      form.method = 'post';
      form.action =
        '/followups/' +
        encodeURIComponent(pin.id) +
        '/' +
        (pin.status === 'completed' ? 'reopen' : 'complete');
      const ret = node('input');
      ret.type = 'hidden';
      ret.name = 'returnTo';
      ret.value = '/map?zone=' + zone.id + '&pin=' + pin.id;
      const locate = node('button', 'Show on map');
      locate.type = 'button';
      locate.onclick = () => {
        selectedMark = pin.id;
        const width = Math.min(view[2], 260);
        setView([pin.start[0] - width / 2, pin.start[1] - width / 2, width, width]);
        render();
        document.querySelector('.map-card').scrollIntoView({ block: 'nearest' });
      };
      form.append(
        ret,
        locate,
        ' ',
        node('button', pin.status === 'completed' ? 'Reopen visit' : 'Mark visited')
      );
      row.append(form);
      list.append(row);
    }

    renderLines(stretches(zone.id), $('stretchList'), 'No stretches marked yet.');
    const housing = (state.housing || []).filter((s) => s.zoneId === zone.id);
    $('housingCount').textContent = housing.length;
    renderLines(housing, $('housingList'), 'No housing marks yet.');
  }

  function renderLines(lines, list, emptyMessage) {
    list.replaceChildren();
    if (!lines.length) list.append(node('p', emptyMessage, 'empty'));
    lines
      .slice()
      .reverse()
      .forEach((s, i) => {
        const row = node('div', undefined, 'entry');
        row.id = 'item-' + s.id;
        row.classList.toggle('highlight', s.id === selectedMark);
        const label = s.housing ? housingLabels[s.housing] : 'Stretch';
        const show = node('button', label + ' ' + (s.housing ? '· ' : '') + (lines.length - i));
        show.type = 'button';
        show.onclick = () => {
          selectedMark = s.id;
          fit();
          zoom(0.55, [(s.start[0] + s.end[0]) / 2, (s.start[1] + s.end[1]) / 2]);
          render();
          document.querySelector('.map-card').scrollIntoView({ block: 'nearest' });
        };
        const del = node('button', 'Remove', 'danger');
        del.onclick = () => {
          const collection = s.housing ? 'housing' : 'stretches';
          if (confirm(s.housing ? 'Remove this housing mark?' : 'Remove this walked stretch?'))
            save('/api/field-map/' + collection + '/' + s.id, 'DELETE', null, 'Mark removed.');
        };
        row.append(
          node('p', s.createdAt ? dateLabel(s.createdAt) : 'From the paper map', 'item-meta'),
          show,
          ' ',
          del
        );
        list.append(row);
      });
  }

  function render() {
    $('totalDone').textContent = state.completed.length + ' / 9';
    $('progress').value = state.completed.length;
    $('zoneList').replaceChildren();
    for (const z of zones) {
      const done = state.completed.includes(z.id);
      const walked = stretches(z.id).length;
      const row = node('button', undefined, 'zone-row' + (done ? ' done' : ''));
      const main = node('span', undefined, 'row-main');
      main.append(
        node('strong', 'Zone ' + z.id),
        node(
          'small',
          done ? 'Completed' : walked ? count(walked, 'stretch', 'stretches') + ' walked' : z.landmark
        )
      );
      row.append(node('span', String(z.id), 'number'), main, node('span', '›', 'chevron'));
      row.onclick = () => {
        if (!busy) choose(z.id);
      };
      $('zoneList').append(row);
    }
    town.querySelectorAll('[data-zone]').forEach((el) => {
      const n = Number(el.dataset.zone);
      el.classList.toggle('done', state.completed.includes(n));
      el.classList.toggle('selected', zone?.id === n);
      el.classList.toggle('muted', Boolean(zone && zone.id !== n));
    });
    $('overview').hidden = Boolean(zone);
    $('detail').hidden = !zone;
    $('mapTitle').textContent = zone ? 'Zone ' + zone.id + ' map' : 'Town overview';
    if (zone) {
      $('detailTitle').textContent = 'Zone ' + zone.id;
      $('landmark').textContent = zone.landmark;
      $('counts').textContent =
        count(stretches(zone.id).length, 'stretch', 'stretches') +
        ' walked · ' +
        count(openPins(zone.id).length, 'open visit', 'open visits');
      const completed = state.completed.includes(zone.id);
      $('completionHint').textContent = completed
        ? 'This zone is marked complete.'
        : 'Mark this zone complete once you’ve covered the homes on each street. Return visits can stay open.';
      $('zoneStatus').textContent = completed ? 'Reopen zone' : 'Mark zone complete';
      renderEntries();
    }
    drawMarks();
    updateTools();
  }

  async function save(path, method, body, success) {
    if (busy) return;
    busy = true;
    document
      .querySelectorAll('#tools button,#zoneStatus,#allZones,#visitForm button')
      .forEach((b) => (b.disabled = true));
    $('formError').textContent = '';
    try {
      const res = await fetch(path, {
        method,
        headers: { 'Content-Type': 'application/json' },
        body: body ? JSON.stringify(body) : undefined
      });
      if (!res.headers.get('content-type')?.includes('application/json'))
        throw new Error('Could not save. Check your connection and CRM login, then try again.');
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Could not save.');
      state = data;
      if ($('visitDialog').open) $('visitDialog').close();
      edit();
      render();
      message(success);
    } catch (err) {
      const text =
        err instanceof TypeError ? 'Could not reach the CRM. Reconnect and try again.' : err.message;
      if ($('visitDialog').open) $('formError').textContent = text;
      else if (mode) $('instruction').textContent = text;
      message(text, true);
    } finally {
      busy = false;
      document
        .querySelectorAll('#tools button,#zoneStatus,#allZones,#visitForm button')
        .forEach((b) => (b.disabled = false));
    }
  }

  function tap(event) {
    if (busy) return;
    const element = document.elementFromPoint(event.clientX, event.clientY);
    const point = mapPoint(event.clientX, event.clientY);
    if (mode) {
      const shape = $('zones').querySelector('[data-zone="' + zone.id + '"]');
      const p = new DOMPoint(...point);
      if (!shape.isPointInFill(p) && !shape.isPointInStroke(p)) {
        $('instruction').textContent = 'Tap inside Zone ' + zone.id + '.';
        return;
      }
      if (mode !== 'pin' && start) {
        save(
          '/api/field-map/marks/' + id,
          'PUT',
          {
            type: mode,
            zoneId: zone.id,
            start,
            end: point,
            housing: mode === 'housing' ? $('housingKind').value : undefined
          },
          mode === 'housing' ? 'Street housing saved.' : 'Walked stretch saved.'
        );
      } else {
        start = point;
        $('draft').replaceChildren();
        svg($('draft'), 'circle', {
          cx: point[0],
          cy: point[1],
          r: 5 / town.getScreenCTM().a,
          fill: mode === 'housing' ? '#b1353d' : '#297854',
          stroke: '#fff',
          'stroke-width': 2,
          'vector-effect': 'non-scaling-stroke'
        });
        if (mode === 'pin') {
          $('visitForm').reset();
          $('formError').textContent = '';
          $('dialogZone').textContent = 'Zone ' + zone.id + ' · ' + zone.landmark;
          $('visitDialog').showModal();
        }
        updateTools();
      }
      return;
    }
    const markId = element?.closest('[data-mark]')?.dataset.mark;
    const mark =
      markId && [...state.pins, ...state.stretches, ...(state.housing || [])].find((m) => m.id === markId);
    if (mark) return showMark(mark);
    const n = Number(element?.closest('[data-zone]')?.dataset.zone);
    if (n && n !== zone?.id) choose(n);
  }

  town.addEventListener('pointerdown', (e) => {
    if (e.button !== 0) return;
    if (!pointers.size) dragged = false;
    pointers.set(e.pointerId, [e.clientX, e.clientY]);
    town.setPointerCapture(e.pointerId);
    const pts = [...pointers.values()];
    gesture = { x: e.clientX, y: e.clientY, view: view.slice(), scale: town.getScreenCTM().a };
    if (pts.length === 2) {
      dragged = true;
      gesture.distance = Math.hypot(pts[0][0] - pts[1][0], pts[0][1] - pts[1][1]);
      gesture.at = mapPoint((pts[0][0] + pts[1][0]) / 2, (pts[0][1] + pts[1][1]) / 2);
    }
  });
  town.addEventListener('pointermove', (e) => {
    if (!gesture || !pointers.has(e.pointerId)) return;
    pointers.set(e.pointerId, [e.clientX, e.clientY]);
    const pts = [...pointers.values()];
    if (pts.length === 2 && gesture.distance) {
      const distance = Math.hypot(pts[0][0] - pts[1][0], pts[0][1] - pts[1][1]);
      if (distance > 0) {
        setView(gesture.view);
        zoom(gesture.distance / distance, gesture.at);
        drawMarks();
      }
    } else if (pts.length === 1) {
      const dx = e.clientX - gesture.x,
        dy = e.clientY - gesture.y;
      if (Math.hypot(dx, dy) > 6) dragged = true;
      if (dragged)
        setView([
          gesture.view[0] - dx / gesture.scale,
          gesture.view[1] - dy / gesture.scale,
          ...gesture.view.slice(2)
        ]);
    }
  });
  function release(e) {
    if (!pointers.has(e.pointerId)) return;
    pointers.delete(e.pointerId);
    if (!pointers.size && !dragged && e.type === 'pointerup') tap(e);
    gesture = null;
  }
  town.addEventListener('pointerup', release);
  town.addEventListener('pointercancel', release);
  $('allZones').onclick = () => choose();
  $('source').onchange = () => {
    const paper = $('source').value === 'paper';
    if (paper) $('paper').setAttribute('href', '/map/reference.jpg');
    $('paper').setAttribute('visibility', paper ? 'visible' : 'hidden');
    $('drawing').setAttribute('visibility', paper ? 'hidden' : 'visible');
  };
  $('plus').onclick = () => {
    zoom(0.72);
    drawMarks();
  };
  $('minus').onclick = () => {
    zoom(1 / 0.72);
    drawMarks();
  };
  $('fit').onclick = () => {
    fit();
    drawMarks();
  };
  $('housingTool').onclick = () => edit('housing');
  $('walk').onclick = () => edit('stretch');
  $('visit').onclick = () => edit('pin');
  $('cancel').onclick = () => edit();
  $('dismiss').onclick = () => $('visitDialog').close();
  $('visitDialog').addEventListener('cancel', (e) => {
    if (busy) e.preventDefault();
  });
  $('visitDialog').addEventListener('close', () => edit());
  $('visitForm').onsubmit = (e) => {
    e.preventDefault();
    const fields = Object.fromEntries(new FormData(e.target));
    save(
      '/api/field-map/marks/' + id,
      'PUT',
      { type: 'pin', zoneId: zone.id, start, ...fields },
      'Return visit saved.'
    );
  };
  $('zoneStatus').onclick = () => {
    const completed = !state.completed.includes(zone.id);
    save(
      '/api/field-map/zones/' + zone.id,
      'PUT',
      { completed },
      completed ? 'Zone marked complete.' : 'Zone reopened.'
    );
  };
  const query = new URLSearchParams(location.search);
  choose(Number(query.get('zone')), false);
  const pin = state.pins.find((p) => p.id === query.get('pin'));
  if (pin) showMark(pin);
  new ResizeObserver(drawMarks).observe(town);
})();
