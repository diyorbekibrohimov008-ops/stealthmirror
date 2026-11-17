import { describe, expect, it } from 'vitest';
import {
  buildGoogleBusyPayload,
  buildOutlookBusyPayload,
  computeEventFingerprint,
  detectLegacySyncClutter,
  isStealthManagedEvent,
  sanitizeEventForMirror,
} from '../shared/sanitizer';
import type { CalendarEvent, SyncRule } from '../shared/types';

const baseEvent: CalendarEvent = {
  id: 'evt_1',
  accountId: 'acc_work',
  externalEventId: 'g_ext_123',
  title: 'Confidential Acquisition Call w/ Stripe',
  description: 'Zoom link: https://zoom.us/j/999888777 Passcode: secret123',
  location: 'Boardroom 4B / Private',
  attendees: ['ceo@stripe.com', 'counsel@acme.com'],
  startIso: '2026-11-12T15:00:00.000Z',
  endIso: '2026-11-12T16:00:00.000Z',
  allDay: false,
  showAs: 'busy',
  visibility: 'default',
  responseStatus: 'accepted',
  locked: false,
  updatedAt: '2026-11-11T10:00:00.000Z',
};

const baseRule: SyncRule = {
  id: 'rule_1',
  workspaceId: 'ws_1',
  sourceAccountId: 'acc_work',
  targetAccountId: 'acc_personal',
  enabled: true,
  busyTitle: 'Busy',
  lockMirrored: true,
  skipFreeEvents: true,
  skipDeclinedEvents: true,
  skipAllDayEvents: true,
  createdAt: '2026-11-11T09:00:00.000Z',
};

describe('sanitizer', () => {
  it('strips all sensitive title, description, location, and attendee metadata', () => {
    const sanitized = sanitizeEventForMirror(baseEvent, baseRule);

    expect(sanitized.title).toBe('Busy');
    expect(sanitized.description).toBe('');
    expect(sanitized.location).toBe('');
    expect(sanitized.attendees).toEqual([]);
    expect(sanitized.visibility).toBe('private');
    expect(sanitized.showAs).toBe('busy');
    expect(sanitized.locked).toBe(true);
    expect(sanitized.stealthMirrorMeta).toMatchObject({
      managed: true,
      originAccountId: 'acc_work',
      originEventId: 'g_ext_123',
      ruleId: 'rule_1',
    });

    const serialized = JSON.stringify(sanitized);
    expect(serialized).not.toContain('Acquisition');
    expect(serialized).not.toContain('zoom.us');
    expect(serialized).not.toContain('Boardroom');
    expect(serialized).not.toContain('ceo@stripe.com');
  });

  it('builds leak-proof Google Calendar API payloads with locked flags and private extendedProperties', () => {
    const payload = buildGoogleBusyPayload(baseEvent, baseRule);

    expect(payload.summary).toBe('Busy');
    expect(payload.description).toBe('');
    expect(payload.location).toBe('');
    expect(payload.visibility).toBe('private');
    expect(payload.transparency).toBe('opaque');
    expect(payload.locked).toBe(true);
    expect(payload.guestsCanModify).toBe(false);
    expect(payload.guestsCanInviteOthers).toBe(false);
    expect(payload.guestsCanSeeOtherGuests).toBe(false);
    expect(payload.extendedProperties.private.stealthMirrorManaged).toBe('true');
    expect(payload.extendedProperties.private.stealthMirrorOriginEventId).toBe('g_ext_123');
    expect(JSON.stringify(payload)).not.toContain('Stripe');
  });

  it('builds leak-proof Microsoft Graph payloads with sensitivity=private and extended properties', () => {
    const payload = buildOutlookBusyPayload(baseEvent, baseRule);

    expect(payload.subject).toBe('Busy');
    expect(payload.body.content).toBe('');
    expect(payload.location.displayName).toBe('');
    expect(payload.attendees).toEqual([]);
    expect(payload.sensitivity).toBe('private');
    expect(payload.showAs).toBe('busy');
    expect(payload.isReminderOn).toBe(false);
    expect(payload.hideAttendees).toBe(true);
    expect(
      payload.singleValueExtendedProperties.some((p) => p.value === 'true')
    ).toBe(true);
    expect(JSON.stringify(payload)).not.toContain('secret123');
  });

  it('produces identical fingerprints for equal timing/busy state and changes when time moves', () => {
    const f1 = computeEventFingerprint(baseEvent, baseRule);
    const f2 = computeEventFingerprint(
      { ...baseEvent, title: 'Renamed Title That Should Not Affect Fingerprint' },
      baseRule
    );
    const f3 = computeEventFingerprint(
      { ...baseEvent, endIso: '2026-11-12T16:30:00.000Z' },
      baseRule
    );

    expect(f1).toBe(f2);
    expect(f1).not.toBe(f3);
  });

  it('detects managed events cleanly to prevent sync loops', () => {
    const mirrored = sanitizeEventForMirror(baseEvent, baseRule);
    expect(isStealthManagedEvent(baseEvent)).toBe(false);
    expect(isStealthManagedEvent(mirrored)).toBe(true);
  });

  it('detects legacy Reclaim/SyncGene clutter without false-positiving normal meetings', () => {
    expect(
      detectLegacySyncClutter({
        ...baseEvent,
        title: '🛡️ Lunch [Reclaim.ai]',
        description: 'Managed by Reclaim',
      })
    ).toBe('reclaim');
    expect(
      detectLegacySyncClutter({
        ...baseEvent,
        title: 'Busy (Synced via SyncGene)',
        description: '',
      })
    ).toBe('syncgene');
    expect(detectLegacySyncClutter(baseEvent)).toBeNull();
  });
});
