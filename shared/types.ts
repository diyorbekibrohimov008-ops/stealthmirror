export type CalendarProvider = 'google' | 'outlook';

export type EventVisibility = 'default' | 'public' | 'private' | 'confidential';
export type EventShowAs = 'busy' | 'free' | 'tentative' | 'oof';
export type EventResponseStatus =
  | 'accepted'
  | 'tentative'
  | 'declined'
  | 'needsAction'
  | 'organizer';

export interface StealthMirrorMeta {
  managed: boolean;
  originAccountId: string;
  originEventId: string;
  ruleId: string;
  fingerprint: string;
}

export interface CalendarEvent {
  id: string;
  accountId: string;
  externalEventId: string;
  title: string;
  description: string;
  location: string;
  attendees: string[];
  startIso: string;
  endIso: string;
  allDay: boolean;
  showAs: EventShowAs;
  visibility: EventVisibility;
  responseStatus: EventResponseStatus;
  locked: boolean;
  stealthMirrorMeta?: StealthMirrorMeta;
  updatedAt: string;
  deleted?: boolean;
}

export interface CalendarAccount {
  id: string;
  workspaceId: string;
  label: string;
  email: string;
  provider: CalendarProvider;
  calendarId: string;
  calendarName: string;
  color: string;
  mode: 'oauth' | 'sandbox';
  status: 'connected' | 'reauth_needed';
  syncToken?: string;
  deltaLink?: string;
  watchChannelId?: string;
  watchResourceId?: string;
  watchExpiresAt?: string;
  accessToken?: string;
  refreshToken?: string;
  tokenExpiresAt?: string;
  createdAt: string;
}

export interface SyncRule {
  id: string;
  workspaceId: string;
  sourceAccountId: string;
  targetAccountId: string;
  enabled: boolean;
  busyTitle: string;
  lockMirrored: boolean;
  skipFreeEvents: boolean;
  skipDeclinedEvents: boolean;
  skipAllDayEvents: boolean;
  createdAt: string;
}

export interface MirrorMapping {
  id: string;
  workspaceId: string;
  ruleId: string;
  sourceAccountId: string;
  sourceExternalEventId: string;
  targetAccountId: string;
  targetExternalEventId: string;
  fingerprint: string;
  updatedAt: string;
}

export type SyncAction =
  | 'created'
  | 'updated'
  | 'deleted'
  | 'skipped_loop'
  | 'skipped_policy'
  | 'error';

export interface SyncLogEntry {
  id: string;
  workspaceId: string;
  ruleId?: string;
  sourceAccountId: string;
  targetAccountId?: string;
  sourceEventId: string;
  action: SyncAction;
  reason: string;
  latencyMs: number;
  payloadPreview?: Record<string, unknown>;
  createdAt: string;
}

export type BillingPlan = 'trial' | 'monthly_5' | 'lifetime_39';

export interface Workspace {
  id: string;
  name: string;
  slug: string;
  plan: BillingPlan;
  createdAt: string;
}

export interface GoogleBusyEventPayload {
  summary: string;
  description: '';
  location: '';
  visibility: 'private';
  transparency: 'opaque';
  locked: boolean;
  guestsCanModify: false;
  guestsCanInviteOthers: false;
  guestsCanSeeOtherGuests: false;
  start: { dateTime: string } | { date: string };
  end: { dateTime: string } | { date: string };
  extendedProperties: {
    private: {
      stealthMirrorManaged: 'true';
      stealthMirrorOriginAccountId: string;
      stealthMirrorOriginEventId: string;
      stealthMirrorRuleId: string;
      stealthMirrorFingerprint: string;
    };
  };
}

export interface OutlookBusyEventPayload {
  subject: string;
  body: { contentType: 'text'; content: '' };
  location: { displayName: '' };
  attendees: [];
  sensitivity: 'private';
  showAs: 'busy';
  isAllDay: boolean;
  isReminderOn: false;
  hideAttendees: true;
  start: { dateTime: string; timeZone: 'UTC' };
  end: { dateTime: string; timeZone: 'UTC' };
  singleValueExtendedProperties: Array<{
    id: string;
    value: string;
  }>;
}
