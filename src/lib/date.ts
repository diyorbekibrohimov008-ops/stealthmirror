export function getUtcDayKey(isoOrDate: string | Date): string {
  const dt = typeof isoOrDate === 'string' ? new Date(isoOrDate) : isoOrDate;
  return dt.toISOString().slice(0, 10);
}

export function getUtcTodayAndTomorrowKeys(): {
  todayKey: string;
  tomorrowKey: string;
  todayLabel: string;
  tomorrowLabel: string;
} {
  const now = new Date();
  const today = new Date(
    Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate())
  );
  const tomorrow = new Date(
    Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate() + 1)
  );
  const fmt = new Intl.DateTimeFormat('en-US', {
    weekday: 'short',
    month: 'short',
    day: 'numeric',
    timeZone: 'UTC',
  });

  return {
    todayKey: today.toISOString().slice(0, 10),
    tomorrowKey: tomorrow.toISOString().slice(0, 10),
    todayLabel: `Today · ${fmt.format(today)}`,
    tomorrowLabel: `Tomorrow · ${fmt.format(tomorrow)}`,
  };
}

export function formatUtcTimeRange(startIso: string, endIso: string, allDay: boolean): string {
  if (allDay) return 'All day (UTC)';
  const start = new Date(startIso);
  const end = new Date(endIso);
  const fmt = new Intl.DateTimeFormat('en-US', {
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
    timeZone: 'UTC',
  });
  return `${fmt.format(start)}–${fmt.format(end)} UTC`;
}

export function shiftIsoHours(iso: string, hours: number): string {
  const dt = new Date(iso);
  dt.setUTCHours(dt.getUTCHours() + hours);
  return dt.toISOString();
}

export function buildUtcIsoForDayHour(dayKey: string, hour: number, minute = 0): string {
  const [y, m, d] = dayKey.split('-').map((n) => Number.parseInt(n, 10));
  return new Date(Date.UTC(y, m - 1, d, hour, minute, 0, 0)).toISOString();
}
