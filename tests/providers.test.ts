import { describe, expect, it } from 'vitest';
import {
  buildGoogleAuthUrl,
  googleEventToCanonical,
  parseGoogleWebhookHeaders,
} from '../netlify/lib/providers/google';
import {
  buildMicrosoftAuthUrl,
  handleMicrosoftValidationHandshake,
  outlookEventToCanonical,
  parseMicrosoftWebhookNotifications,
  STEALTH_OUTLOOK_EXT_PROP_ID,
} from '../netlify/lib/providers/microsoft';

describe('Google provider adapter', () => {
  it('builds a least-privilege Google Calendar OAuth URL with state and offline access', () => {
    const url = new URL(
      buildGoogleAuthUrl({
        clientId: 'goog-client-123',
        redirectUri: 'https://stealthmirror.app/api/oauth/google/callback',
        state: 'ws_demo:nonce123',
      })
    );

    expect(url.origin + url.pathname).toBe('https://accounts.google.com/o/oauth2/v2/auth');
    expect(url.searchParams.get('client_id')).toBe('goog-client-123');
    expect(url.searchParams.get('access_type')).toBe('offline');
    expect(url.searchParams.get('prompt')).toBe('consent');
    expect(url.searchParams.get('state')).toBe('ws_demo:nonce123');
    expect(url.searchParams.get('scope')).toContain('calendar');
  });

  it('parses Google Calendar push channel headers accurately', () => {
    const headers = new Headers({
      'x-goog-channel-id': 'chan_abc',
      'x-goog-resource-id': 'res_xyz',
      'x-goog-resource-state': 'exists',
      'x-goog-channel-token': 'ws_1:acc_1',
    });

    const parsed = parseGoogleWebhookHeaders(headers);
    expect(parsed).toEqual({
      channelId: 'chan_abc',
      resourceId: 'res_xyz',
      resourceState: 'exists',
      channelToken: 'ws_1:acc_1',
    });
  });

  it('converts a Google API event into canonical form and preserves stealthMirrorManaged metadata', () => {
    const canonical = googleEventToCanonical('acc_g1', {
      id: 'g_evt_9',
      summary: 'M&A Review',
      description: 'Top secret doc',
      location: 'SF HQ',
      start: { dateTime: '2026-11-12T17:00:00Z' },
      end: { dateTime: '2026-11-12T18:00:00Z' },
      status: 'confirmed',
      transparency: 'opaque',
      visibility: 'private',
      locked: true,
      attendees: [{ email: 'partner@firm.com', responseStatus: 'accepted', self: true }],
      extendedProperties: {
        private: {
          stealthMirrorManaged: 'true',
          stealthMirrorOriginAccountId: 'acc_o1',
          stealthMirrorOriginEventId: 'o_evt_1',
          stealthMirrorRuleId: 'rule_1',
          stealthMirrorFingerprint: 'fp_1',
        },
      },
      updated: '2026-11-11T10:00:00Z',
    });

    expect(canonical.externalEventId).toBe('g_evt_9');
    expect(canonical.title).toBe('M&A Review');
    expect(canonical.stealthMirrorMeta?.managed).toBe(true);
    expect(canonical.stealthMirrorMeta?.originAccountId).toBe('acc_o1');
  });
});

describe('Microsoft Outlook / Graph adapter', () => {
  it('builds a Microsoft Graph OAuth URL requesting Calendars.ReadWrite and offline_access', () => {
    const url = new URL(
      buildMicrosoftAuthUrl({
        clientId: 'ms-client-999',
        redirectUri: 'https://stealthmirror.app/api/oauth/microsoft/callback',
        state: 'ws_demo:nonce456',
      })
    );

    expect(url.origin + url.pathname).toBe(
      'https://login.microsoftonline.com/common/oauth2/v2.0/authorize'
    );
    expect(url.searchParams.get('client_id')).toBe('ms-client-999');
    expect(url.searchParams.get('scope')).toContain('Calendars.ReadWrite');
    expect(url.searchParams.get('scope')).toContain('offline_access');
  });

  it('responds immediately with plain-text validationToken during Graph subscription handshake', async () => {
    const req = new Request(
      'https://stealthmirror.app/api/webhooks/microsoft?validationToken=hello-graph-token',
      { method: 'POST' }
    );
    const res = handleMicrosoftValidationHandshake(req);
    expect(res).not.toBeNull();
    expect(res?.status).toBe(200);
    expect(res?.headers.get('content-type')).toContain('text/plain');
    expect(await res?.text()).toBe('hello-graph-token');
  });

  it('parses Graph change notifications and Outlook extended properties for loop detection', () => {
    const notifications = parseMicrosoftWebhookNotifications({
      value: [
        {
          subscriptionId: 'sub_1',
          clientState: 'ws_1:acc_o1',
          changeType: 'updated',
          resourceData: { id: 'o_evt_77' },
        },
      ],
    });
    expect(notifications).toHaveLength(1);
    expect(notifications[0]).toMatchObject({
      subscriptionId: 'sub_1',
      clientState: 'ws_1:acc_o1',
      changeType: 'updated',
      eventId: 'o_evt_77',
    });

    const canonical = outlookEventToCanonical('acc_o1', {
      id: 'o_evt_77',
      subject: 'Busy',
      bodyPreview: '',
      showAs: 'busy',
      sensitivity: 'private',
      isAllDay: false,
      isCancelled: false,
      start: { dateTime: '2026-11-12T17:00:00Z', timeZone: 'UTC' },
      end: { dateTime: '2026-11-12T18:00:00Z', timeZone: 'UTC' },
      lastModifiedDateTime: '2026-11-11T10:00:00Z',
      singleValueExtendedProperties: [
        { id: STEALTH_OUTLOOK_EXT_PROP_ID, value: 'true|acc_g1|g_evt_2|rule_5|fp_9' },
      ],
    });

    expect(canonical.stealthMirrorMeta?.managed).toBe(true);
    expect(canonical.stealthMirrorMeta?.originAccountId).toBe('acc_g1');
    expect(canonical.stealthMirrorMeta?.originEventId).toBe('g_evt_2');
  });
});
