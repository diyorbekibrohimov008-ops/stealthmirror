import {
  buildOutlookBusyPayload,
  STEALTH_OUTLOOK_EXT_PROP_ID,
  STEALTH_OUTLOOK_MANAGED_PROP_ID,
} from '../../../shared/sanitizer';
import type {
  CalendarAccount,
  CalendarEvent,
  EventResponseStatus,
  EventShowAs,
  SyncRule,
} from '../../../shared/types';

export { STEALTH_OUTLOOK_EXT_PROP_ID, STEALTH_OUTLOOK_MANAGED_PROP_ID };

export interface OutlookRawEvent {
  id: string;
  subject?: string;
  bodyPreview?: string;
  location?: { displayName?: string };
  showAs?: string;
  sensitivity?: string;
  isAllDay?: boolean;
  isCancelled?: boolean;
  responseStatus?: { response?: string };
  start?: { dateTime: string; timeZone?: string };
  end?: { dateTime: string; timeZone?: string };
  attendees?: Array<{ emailAddress?: { address?: string } }>;
  singleValueExtendedProperties?: Array<{ id: string; value: string }>;
  lastModifiedDateTime?: string;
  '@removed'?: { reason?: string };
}

export function buildMicrosoftAuthUrl(params: {
  clientId: string;
  redirectUri: string;
  state: string;
}): string {
  const url = new URL('https://login.microsoftonline.com/common/oauth2/v2.0/authorize');
  url.searchParams.set('client_id', params.clientId);
  url.searchParams.set('redirect_uri', params.redirectUri);
  url.searchParams.set('response_type', 'code');
  url.searchParams.set('scope', 'Calendars.ReadWrite offline_access User.Read');
  url.searchParams.set('response_mode', 'query');
  url.searchParams.set('state', params.state);
  return url.toString();
}

export async function exchangeMicrosoftAuthCode(params: {
  code: string;
  clientId: string;
  clientSecret: string;
  redirectUri: string;
}): Promise<{ accessToken: string; refreshToken?: string; expiresIn: number }> {
  const res = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: params.clientId,
      client_secret: params.clientSecret,
      code: params.code,
      redirect_uri: params.redirectUri,
      grant_type: 'authorization_code',
      scope: 'Calendars.ReadWrite offline_access User.Read',
    }),
  });
  if (!res.ok) {
    throw new Error(`Microsoft token exchange failed (${res.status})`);
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

export async function refreshMicrosoftAccessToken(params: {
  refreshToken: string;
  clientId: string;
  clientSecret: string;
}): Promise<{ accessToken: string; refreshToken?: string; expiresIn: number }> {
  const res = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: params.clientId,
      client_secret: params.clientSecret,
      refresh_token: params.refreshToken,
      grant_type: 'refresh_token',
      scope: 'Calendars.ReadWrite offline_access User.Read',
    }),
  });
  if (!res.ok) {
    throw new Error(`Microsoft token refresh failed (${res.status})`);
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

export async function fetchMicrosoftUserProfile(
  accessToken: string
): Promise<{ mail?: string; userPrincipalName: string; displayName?: string }> {
  const res = await fetch('https://graph.microsoft.com/v1.0/me', {
    headers: { authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error(`Microsoft Graph /me failed (${res.status})`);
  return (await res.json()) as {
    mail?: string;
    userPrincipalName: string;
    displayName?: string;
  };
}

export function handleMicrosoftValidationHandshake(req: Request): Response | null {
  const url = new URL(req.url);
  const validationToken = url.searchParams.get('validationToken');
  if (!validationToken) return null;
  return new Response(validationToken, {
    status: 200,
    headers: { 'content-type': 'text/plain; charset=utf-8' },
  });
}

export function parseMicrosoftWebhookNotifications(body: {
  value?: Array<{
    subscriptionId?: string;
    clientState?: string;
    changeType?: string;
    resourceData?: { id?: string };
  }>;
}): Array<{
  subscriptionId: string;
  clientState: string;
  changeType: string;
  eventId?: string;
}> {
  return (body.value ?? []).map((item) => ({
    subscriptionId: item.subscriptionId ?? '',
    clientState: item.clientState ?? '',
    changeType: item.changeType ?? 'updated',
    eventId: item.resourceData?.id,
  }));
}

function mapOutlookResponseStatus(raw?: string): EventResponseStatus {
  switch (raw) {
    case 'declined':
      return 'declined';
    case 'tentativelyAccepted':
      return 'tentative';
    case 'organizer':
      return 'organizer';
    case 'notResponded':
    case 'none':
      return 'needsAction';
    default:
      return 'accepted';
  }
}

export function outlookEventToCanonical(
  accountId: string,
  raw: OutlookRawEvent
): CalendarEvent {
  const props = raw.singleValueExtendedProperties ?? [];
  const metaProp = props.find((p) => p.id === STEALTH_OUTLOOK_EXT_PROP_ID);
  const managedFlag = props.find((p) => p.id === STEALTH_OUTLOOK_MANAGED_PROP_ID);

  let managed = managedFlag?.value === 'true';
  let originAccountId = '';
  let originEventId = '';
  let ruleId = '';
  let fingerprint = '';

  if (metaProp?.value) {
    const parts = metaProp.value.split('|');
    if (parts[0] === 'true') {
      managed = true;
      originAccountId = parts[1] ?? '';
      originEventId = parts[2] ?? '';
      ruleId = parts[3] ?? '';
      fingerprint = parts[4] ?? '';
    }
  }

  const showAs: EventShowAs =
    raw.showAs === 'free'
      ? 'free'
      : raw.showAs === 'tentative'
        ? 'tentative'
        : raw.showAs === 'oof'
          ? 'oof'
          : 'busy';

  return {
    id: `evt_${accountId}_${raw.id}`,
    accountId,
    externalEventId: raw.id,
    title: raw.subject ?? '(No subject)',
    description: raw.bodyPreview ?? '',
    location: raw.location?.displayName ?? '',
    attendees: (raw.attendees ?? [])
      .map((a) => a.emailAddress?.address ?? '')
      .filter(Boolean),
    startIso: raw.start?.dateTime
      ? new Date(raw.start.dateTime).toISOString()
      : new Date().toISOString(),
    endIso: raw.end?.dateTime
      ? new Date(raw.end.dateTime).toISOString()
      : new Date().toISOString(),
    allDay: Boolean(raw.isAllDay),
    showAs,
    visibility:
      raw.sensitivity === 'private' || raw.sensitivity === 'confidential'
        ? raw.sensitivity
        : 'default',
    responseStatus: mapOutlookResponseStatus(raw.responseStatus?.response),
    locked: managed,
    stealthMirrorMeta: managed
      ? {
          managed: true,
          originAccountId,
          originEventId,
          ruleId,
          fingerprint,
        }
      : undefined,
    updatedAt: raw.lastModifiedDateTime ?? new Date().toISOString(),
    deleted: Boolean(raw.isCancelled || raw['@removed']),
  };
}

export async function createMicrosoftGraphSubscription(params: {
  accessToken: string;
  notificationUrl: string;
  clientState: string;
}): Promise<{ subscriptionId: string; expirationDateTime: string }> {
  const expirationDateTime = new Date(Date.now() + 2 * 24 * 3600_000).toISOString();
  const res = await fetch('https://graph.microsoft.com/v1.0/subscriptions', {
    method: 'POST',
    headers: {
      authorization: `Bearer ${params.accessToken}`,
      'content-type': 'application/json',
    },
    body: JSON.stringify({
      changeType: 'created,updated,deleted',
      notificationUrl: params.notificationUrl,
      resource: 'me/events',
      expirationDateTime,
      clientState: params.clientState,
    }),
  });
  if (!res.ok) {
    throw new Error(`Microsoft Graph subscription failed (${res.status})`);
  }
  const data = (await res.json()) as { id: string; expirationDateTime: string };
  return {
    subscriptionId: data.id,
    expirationDateTime: data.expirationDateTime,
  };
}

export async function renewMicrosoftGraphSubscription(params: {
  accessToken: string;
  subscriptionId: string;
}): Promise<{ expirationDateTime: string }> {
  const expirationDateTime = new Date(Date.now() + 2 * 24 * 3600_000).toISOString();
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/subscriptions/${encodeURIComponent(params.subscriptionId)}`,
    {
      method: 'PATCH',
      headers: {
        authorization: `Bearer ${params.accessToken}`,
        'content-type': 'application/json',
      },
      body: JSON.stringify({ expirationDateTime }),
    }
  );
  if (!res.ok) {
    throw new Error(`Microsoft Graph subscription renewal failed (${res.status})`);
  }
  const data = (await res.json()) as { expirationDateTime: string };
  return { expirationDateTime: data.expirationDateTime };
}

export async function fetchMicrosoftDeltaChanges(params: {
  accessToken: string;
  deltaLink?: string;
}): Promise<{ items: OutlookRawEvent[]; nextDeltaLink?: string }> {
  const expandFilter = encodeURIComponent(
    `id eq '${STEALTH_OUTLOOK_EXT_PROP_ID}' or id eq '${STEALTH_OUTLOOK_MANAGED_PROP_ID}'`
  );
  const startIso = new Date(Date.now() - 24 * 3600_000).toISOString();
  const endIso = new Date(Date.now() + 60 * 24 * 3600_000).toISOString();
  const url =
    params.deltaLink ??
    `https://graph.microsoft.com/v1.0/me/calendarView/delta?startDateTime=${encodeURIComponent(startIso)}&endDateTime=${encodeURIComponent(endIso)}&$expand=singleValueExtendedProperties($filter=${expandFilter})`;

  const res = await fetch(url, {
    headers: { authorization: `Bearer ${params.accessToken}` },
  });
  if (!res.ok) {
    throw new Error(`Microsoft Graph delta fetch failed (${res.status})`);
  }
  const data = (await res.json()) as {
    value?: OutlookRawEvent[];
    '@odata.deltaLink'?: string;
    '@odata.nextLink'?: string;
  };
  return {
    items: data.value ?? [],
    nextDeltaLink: data['@odata.deltaLink'] ?? data['@odata.nextLink'],
  };
}

export async function upsertMicrosoftBusyEvent(
  targetAccount: CalendarAccount,
  rule: SyncRule,
  sourceEvent: CalendarEvent,
  existingTargetExternalId?: string
): Promise<{ targetExternalEventId: string; payloadPreview: Record<string, unknown> }> {
  const payload = buildOutlookBusyPayload(sourceEvent, rule);
  if (!targetAccount.accessToken) {
    throw new Error('Outlook account missing accessToken');
  }
  const endpoint = existingTargetExternalId
    ? `https://graph.microsoft.com/v1.0/me/events/${encodeURIComponent(existingTargetExternalId)}`
    : 'https://graph.microsoft.com/v1.0/me/events';

  const res = await fetch(endpoint, {
    method: existingTargetExternalId ? 'PATCH' : 'POST',
    headers: {
      authorization: `Bearer ${targetAccount.accessToken}`,
      'content-type': 'application/json',
    },
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    throw new Error(`Microsoft Graph upsert Busy event failed (${res.status})`);
  }
  const created = (await res.json()) as { id: string };
  return {
    targetExternalEventId: created.id,
    payloadPreview: payload as unknown as Record<string, unknown>,
  };
}

export async function deleteMicrosoftBusyEvent(
  targetAccount: CalendarAccount,
  targetExternalEventId: string
): Promise<void> {
  if (!targetAccount.accessToken) return;
  const endpoint = `https://graph.microsoft.com/v1.0/me/events/${encodeURIComponent(targetExternalEventId)}`;
  const res = await fetch(endpoint, {
    method: 'DELETE',
    headers: { authorization: `Bearer ${targetAccount.accessToken}` },
  });
  if (!res.ok && res.status !== 404) {
    throw new Error(`Microsoft Graph delete Busy event failed (${res.status})`);
  }
}
