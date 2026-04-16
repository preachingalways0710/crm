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

document.addEventListener('DOMContentLoaded', async () => {
  const calendarEl = document.getElementById('calendar');
  const showBirthdaysToggle = document.getElementById('showBirthdaysToggle');

  const standardEventsSource = {
    id: 'church-events',
    url: '/api/events'
  };

  const birthdaysSource = {
    id: 'birthday-events',
    url: '/api/birthdays'
  };

  const calendar = new FullCalendar.Calendar(calendarEl, {
    initialView: 'dayGridMonth',
    editable: false,
    eventSources: [standardEventsSource, birthdaysSource],
    dateClick: async (info) => {
      const title = window.prompt('Event title');
      if (!title) return;
      const description = window.prompt('Description (optional)') || '';

      await api('/api/events', {
        method: 'POST',
        body: JSON.stringify({ title, start: info.dateStr, description })
      });

      calendar.refetchEvents();
    },
    eventClick: async (info) => {
      if (info.event.extendedProps.sourceType === 'birthday') {
        return;
      }

      const nextTitle = window.prompt('Update event title (or leave empty to delete):', info.event.title);
      if (nextTitle === null) return;

      if (!nextTitle.trim()) {
        const ok = window.confirm('Delete this event?');
        if (!ok) return;

        await api(`/api/events/${info.event.id}`, { method: 'DELETE' });
        info.event.remove();
        return;
      }

      await api('/api/events', {
        method: 'POST',
        body: JSON.stringify({
          id: info.event.id,
          title: nextTitle,
          start: info.event.startStr,
          end: info.event.endStr,
          description: info.event.extendedProps.description || ''
        })
      });

      info.event.setProp('title', nextTitle);
    }
  });

  if (showBirthdaysToggle) {
    showBirthdaysToggle.addEventListener('change', () => {
      const existing = calendar.getEventSourceById('birthday-events');

      if (showBirthdaysToggle.checked) {
        if (!existing) {
          calendar.addEventSource(birthdaysSource);
        }
      } else if (existing) {
        existing.remove();
      }
    });
  }

  calendar.render();
});
