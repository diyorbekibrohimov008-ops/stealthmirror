import { buildGoogleBusyPayload } from '../../../shared/sanitizer';
import type {
  CalendarAccount,
  CalendarEvent,
  EventResponseStatus,
  SyncRule,
} from '../../../shared/types';

export interface GoogleRawEvent {
  id: string;
  summary?: string;
  description?: string;
  location?: string;
  status?: string;
  transparency?: string;
  visibility?: string;
  locked?: boolean;
  start?: { dateTime?: string; date?: string };
  end?: { dateTime?: string; date?: string };
  attendees?: Array<{ email?: string; responseStatus?: string; self?: boolean }>;
  extendedProperties?: {
    private?: Record<string, string>;
  };
  updated?: string;
}

export function buildGoogleAuthUrl(params: {
  clientId: string;
  redirectUri: string;
  state: string;
}): string {
  const url = new URL('https://accounts.google.com/o/oauth2/v2/auth');
  url.searchParams.set('client_id', params.clientId);
  url.searchParams.set('redirect_uri', params.redirectUri);
  url.searchParams.set('response_type', 'code');
  url.searchParams.set(
    'scope',
    [
      'https://www.googleapis.com/auth/calendar.events',
      'https://www.googleapis.com/auth/userinfo.email',
      'https://www.googleapis.com/auth/userinfo.profile',
    ].join(' ')
  );
  url.searchParams.set('access_type', 'offline');
  url.searchParams.set('prompt', 'consent');
  url.searchParams.set('state', params.state);
  return url.toString();
}

export async function exchangeGoogleAuthCode(params: {
  code: string;
  clientId: string;
  clientSecret: string;
  redirectUri: string;
}): Promise<{ accessToken: string; refreshToken?: string; expiresIn: number }> {
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      code: params.code,
      client_id: params.clientId,
      client_secret: params.clientSecret,
      redirect_uri: params.redirectUri,
      grant_type: 'authorization_code',
    }),
  });
  if (!res.ok) {
    throw new Error(`Google token exchange failed (${res.status})`);
  }
  const data = (await res.json()) as {
    access_token: string;
    refresh_token?: string;
    expires_in: number;
  };
  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token,
    expiresIn: data.expires_in,
  };
}

export async function refreshGoogleAccessToken(params: {
  refreshToken: string;
  clientId: string;
  clientSecret: string;
}): Promise<{ accessToken: string; expiresIn: number }> {
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      refresh_token: params.refreshToken,
      client_id: params.clientId,
      client_secret: params.clientSecret,
      grant_type: 'refresh_token',
    }),
  });
  if (!res.ok) {
    throw new Error(`Google token refresh failed (${res.status})`);
  }
  const data = (await res.json()) as { access_token: string; expires_in: number };
  return { accessToken: data.access_token, expiresIn: data.expires_in };
}

export async function fetchGoogleUserProfile(
  accessToken: string
): Promise<{ email: string; name?: string }> {
  const res = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
    headers: { authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error(`Google userinfo failed (${res.status})`);
  const data = (await res.json()) as { email: string; name?: string };
  return data;
}

export function parseGoogleWebhookHeaders(headers: Headers): {
  channelId: string | null;
  resourceId: string | null;
  resourceState: string | null;
  channelToken: string | null;
} {
  return {
    channelId: headers.get('x-goog-channel-id'),
    resourceId: headers.get('x-goog-resource-id'),
    resourceState: headers.get('x-goog-resource-state'),
    channelToken: headers.get('x-goog-channel-token'),
  };
}

export function googleEventToCanonical(
  accountId: string,
  raw: GoogleRawEvent
): CalendarEvent {
  const priv = raw.extendedProperties?.private;
  const managed = priv?.stealthMirrorManaged === 'true';
  const allDay = Boolean(raw.start?.date && !raw.start?.dateTime);

  const startIso = raw.start?.dateTime
    ? new Date(raw.start.dateTime).toISOString()
    : raw.start?.date
      ? `${raw.start.date}T00:00:00.000Z`
      : new Date().toISOString();
  const endIso = raw.end?.dateTime
    ? new Date(raw.end.dateTime).toISOString()
    : raw.end?.date
      ? `${raw.end.date}T00:00:00.000Z`
      : startIso;

  const selfAttendee = raw.attendees?.find((a) => a.self);
  const responseStatus = (selfAttendee?.responseStatus as EventResponseStatus) ?? 'accepted';

  return {
    id: `evt_${accountId}_${raw.id}`,
    accountId,
    externalEventId: raw.id,
    title: raw.summary ?? '(No title)',
    description: raw.description ?? '',
    location: raw.location ?? '',
    attendees: (raw.attendees ?? []).map((a) => a.email ?? '').filter(Boolean),
    startIso,
    endIso,
    allDay,
    showAs: raw.transparency === 'transparent' ? 'free' : 'busy',
    visibility:
      raw.visibility === 'private' ||
      raw.visibility === 'public' ||
      raw.visibility === 'confidential'
        ? raw.visibility
        : 'default',
    responseStatus,
    locked: Boolean(raw.locked),
    stealthMirrorMeta: managed
      ? {
          managed: true,
          originAccountId: priv?.stealthMirrorOriginAccountId ?? '',
          originEventId: priv?.stealthMirrorOriginEventId ?? '',
          ruleId: priv?.stealthMirrorRuleId ?? '',
          fingerprint: priv?.stealthMirrorFingerprint ?? '',
        }
      : undefined,
    updatedAt: raw.updated ?? new Date().toISOString(),
    deleted: raw.status === 'cancelled',
  };
}

export async function watchGoogleCalendar(params: {
  accessToken: string;
  calendarId: string;
  channelId: string;
  webhookUrl: string;
  token: string;
}): Promise<{ resourceId: string; expirationIso: string }> {
  const res = await fetch(
    `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(params.calendarId)}/events/watch`,
    {
      method: 'POST',
      headers: {
        authorization: `Bearer ${params.accessToken}`,
        'content-type': 'application/json',
      },
      body: JSON.stringify({
        id: params.channelId,
        type: 'web_hook',
        address: params.webhookUrl,
        token: params.token,
      }),
    }
  );
  if (!res.ok) {
    throw new Error(`Google Calendar watch failed (${res.status})`);
  }
  const data = (await res.json()) as { resourceId: string; expiration: string };
  return {
    resourceId: data.resourceId,
    expirationIso: new Date(Number(data.expiration)).toISOString(),
  };
}

export async function fetchGoogleIncrementalChanges(params: {
  accessToken: string;
  calendarId: string;
  syncToken?: string;
}): Promise<{ items: GoogleRawEvent[]; nextSyncToken?: string; syncResetRequired?: boolean }> {
  const url = new URL(
    `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(params.calendarId)}/events`
  );
  url.searchParams.set('singleEvents', 'true');
  url.searchParams.set('showDeleted', 'true');
  if (params.syncToken) {
    url.searchParams.set('syncToken', params.syncToken);
  } else {
    url.searchParams.set('timeMin', new Date(Date.now() - 24 * 3600_000).toISOString());
    url.searchParams.set('maxResults', '100');
  }

  const res = await fetch(url.toString(), {
    headers: { authorization: `Bearer ${params.accessToken}` },
  });
  if (res.status === 410) {
    return { items: [], syncResetRequired: true };
  }
  if (!res.ok) {
    throw new Error(`Google Calendar list events failed (${res.status})`);
  }
  const data = (await res.json()) as { items?: GoogleRawEvent[]; nextSyncToken?: string };
  return {
    items: data.items ?? [],
    nextSyncToken: data.nextSyncToken,
  };
}

export async function upsertGoogleBusyEvent(
  targetAccount: CalendarAccount,
  rule: SyncRule,
  sourceEvent: CalendarEvent,
  existingTargetExternalId?: string
): Promise<{ targetExternalEventId: string; payloadPreview: Record<string, unknown> }> {
  const payload = buildGoogleBusyPayload(sourceEvent, rule);
  if (!targetAccount.accessToken) {
    throw new Error('Google account missing accessToken');
  }
  const baseUrl = `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(targetAccount.calendarId)}/events`;
  const endpoint = existingTargetExternalId
    ? `${baseUrl}/${encodeURIComponent(existingTargetExternalId)}`
    : baseUrl;

  const res = await fetch(endpoint, {
    method: existingTargetExternalId ? 'PATCH' : 'POST',
    headers: {
      authorization: `Bearer ${targetAccount.accessToken}`,
      'content-type': 'application/json',
    },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    throw new Error(`Google upsert Busy event failed (${res.status})`);
  }
  const created = (await res.json()) as { id: string };
  return {
    targetExternalEventId: created.id,
    payloadPreview: payload as unknown as Record<string, unknown>,
  };
}

export async function deleteGoogleBusyEvent(
  targetAccount: CalendarAccount,
  targetExternalEventId: string
): Promise<void> {
  if (!targetAccount.accessToken) return;
  const endpoint = `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(targetAccount.calendarId)}/events/${encodeURIComponent(targetExternalEventId)}`;
  const res = await fetch(endpoint, {
    method: 'DELETE',
    headers: { authorization: `Bearer ${targetAccount.accessToken}` },
  });
  if (!res.ok && res.status !== 404 && res.status !== 410) {
    throw new Error(`Google delete Busy event failed (${res.status})`);
  }
}
