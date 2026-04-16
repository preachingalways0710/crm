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
  const eventData = await api('/api/events');

  const calendar = new FullCalendar.Calendar(calendarEl, {
    initialView: 'dayGridMonth',
    editable: false,
    events: eventData,
    dateClick: async (info) => {
      const title = window.prompt('Event title');
      if (!title) return;
      const description = window.prompt('Description (optional)') || '';

      await api('/api/events', {
        method: 'POST',
        body: JSON.stringify({ title, start: info.dateStr, description })
      });

      calendar.refetchEvents();
      window.location.reload();
    },
    eventClick: async (info) => {
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

  calendar.render();
});
