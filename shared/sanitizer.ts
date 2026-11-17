import type {
  CalendarEvent,
  GoogleBusyEventPayload,
  OutlookBusyEventPayload,
  SyncRule,
} from './types';

export const STEALTH_OUTLOOK_MANAGED_PROP_ID =
  'String {66f5a359-4659-4830-9070-00047ec6ac6e} Name stealthMirrorManaged';
export const STEALTH_OUTLOOK_EXT_PROP_ID =
  'String {66f5a359-4659-4830-9070-00047ec6ac6e} Name stealthMirrorMeta';

const SAFE_BUSY_LABELS = new Set(['Busy', 'Busy — Blocked', 'Unavailable', 'Private Block']);

/**
 * Ensures mirrored titles never leak source metadata by falling back to "Busy"
 * unless the rule uses a short, approved static label.
 */
export function resolveSafeBusyTitle(rawTitle?: string): string {
  const trimmed = rawTitle?.trim() ?? '';
  if (!trimmed) return 'Busy';
  if (SAFE_BUSY_LABELS.has(trimmed)) return trimmed;
  return trimmed.slice(0, 24);
}

/**
 * Deterministic fingerprint covering only timing and busy state—never the source
 * title, description, location, or attendees.
 */
export function computeEventFingerprint(event: CalendarEvent, rule: SyncRule): string {
  const label = resolveSafeBusyTitle(rule.busyTitle);
  return [
    event.startIso,
    event.endIso,
    event.allDay ? '1' : '0',
    label,
    rule.lockMirrored ? 'L' : 'U',
  ].join('|');
}

export function isStealthManagedEvent(event: CalendarEvent): boolean {
  return event.stealthMirrorMeta?.managed === true;
}

export function sanitizeEventForMirror(
  sourceEvent: CalendarEvent,
  rule: SyncRule,
  existingId?: string,
  existingExternalId?: string
): CalendarEvent {
  const safeTitle = resolveSafeBusyTitle(rule.busyTitle);
  const fingerprint = computeEventFingerprint(sourceEvent, rule);
  const externalEventId =
    existingExternalId ?? `sm_${rule.targetAccountId}_${sourceEvent.externalEventId}`;

  return {
    id: existingId ?? `evt_${externalEventId}`,
    accountId: rule.targetAccountId,
    externalEventId,
    title: safeTitle,
    description: '',
    location: '',
    attendees: [],
    startIso: sourceEvent.startIso,
    endIso: sourceEvent.endIso,
    allDay: sourceEvent.allDay,
    showAs: 'busy',
    visibility: 'private',
    responseStatus: 'accepted',
    locked: rule.lockMirrored,
    stealthMirrorMeta: {
      managed: true,
      originAccountId: sourceEvent.accountId,
      originEventId: sourceEvent.externalEventId,
      ruleId: rule.id,
      fingerprint,
    },
    updatedAt: new Date().toISOString(),
  };
}

export function buildGoogleBusyPayload(
  sourceEvent: CalendarEvent,
  rule: SyncRule
): GoogleBusyEventPayload {
  const safeTitle = resolveSafeBusyTitle(rule.busyTitle);
  const fingerprint = computeEventFingerprint(sourceEvent, rule);

  return {
    summary: safeTitle,
    description: '',
    location: '',
    visibility: 'private',
    transparency: 'opaque',
    locked: rule.lockMirrored,
    guestsCanModify: false,
    guestsCanInviteOthers: false,
    guestsCanSeeOtherGuests: false,
    start: sourceEvent.allDay
      ? { date: sourceEvent.startIso.slice(0, 10) }
      : { dateTime: sourceEvent.startIso },
    end: sourceEvent.allDay
      ? { date: sourceEvent.endIso.slice(0, 10) }
      : { dateTime: sourceEvent.endIso },
    extendedProperties: {
      private: {
        stealthMirrorManaged: 'true',
        stealthMirrorOriginAccountId: sourceEvent.accountId,
        stealthMirrorOriginEventId: sourceEvent.externalEventId,
        stealthMirrorRuleId: rule.id,
        stealthMirrorFingerprint: fingerprint,
      },
    },
  };
}

export function buildOutlookBusyPayload(
  sourceEvent: CalendarEvent,
  rule: SyncRule
): OutlookBusyEventPayload {
  const safeTitle = resolveSafeBusyTitle(rule.busyTitle);
  const fingerprint = computeEventFingerprint(sourceEvent, rule);
  const packedMeta = [
    'true',
    sourceEvent.accountId,
    sourceEvent.externalEventId,
    rule.id,
    fingerprint,
  ].join('|');

  return {
    subject: safeTitle,
    body: { contentType: 'text', content: '' },
    location: { displayName: '' },
    attendees: [],
    sensitivity: 'private',
    showAs: 'busy',
    isAllDay: sourceEvent.allDay,
    isReminderOn: false,
    hideAttendees: true,
    start: { dateTime: sourceEvent.startIso, timeZone: 'UTC' },
    end: { dateTime: sourceEvent.endIso, timeZone: 'UTC' },
    singleValueExtendedProperties: [
      {
        id: STEALTH_OUTLOOK_MANAGED_PROP_ID,
        value: 'true',
      },
      {
        id: STEALTH_OUTLOOK_EXT_PROP_ID,
        value: packedMeta,
      },
    ],
  };
}

export function detectLegacySyncClutter(
  event: CalendarEvent
): 'reclaim' | 'syncgene' | null {
  if (isStealthManagedEvent(event)) return null;
  const hay = `${event.title}\n${event.description}`.toLowerCase();
  if (
    hay.includes('[reclaim') ||
    hay.includes('reclaim.ai') ||
    hay.includes('managed by reclaim')
  ) {
    return 'reclaim';
  }
  if (hay.includes('syncgene') || hay.includes('synced via syncgene')) {
    return 'syncgene';
  }
  return null;
}
