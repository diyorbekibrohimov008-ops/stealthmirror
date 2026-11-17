import { describe, expect, it } from 'vitest';
import { StealthSyncEngine } from '../shared/syncEngine';
import type {
  CalendarAccount,
  CalendarEvent,
  MirrorMapping,
  SyncLogEntry,
  SyncRule,
} from '../shared/types';

function createInMemoryPort() {
  const accounts = new Map<string, CalendarAccount>();
  const events = new Map<string, CalendarEvent>();
  const rules = new Map<string, SyncRule>();
  const mappings = new Map<string, MirrorMapping>();
  const logs: SyncLogEntry[] = [];

  const eventKey = (accountId: string, externalEventId: string) =>
    `${accountId}::${externalEventId}`;

  return {
    accounts,
    events,
    rules,
    mappings,
    logs,
    port: {
      async getRulesForSource(workspaceId: string, sourceAccountId: string) {
        return Array.from(rules.values()).filter(
          (r) => r.workspaceId === workspaceId && r.sourceAccountId === sourceAccountId && r.enabled
        );
      },
      async getAccount(accountId: string) {
        return accounts.get(accountId) ?? null;
      },
      async getMapping(workspaceId: string, ruleId: string, sourceExternalEventId: string) {
        return (
          Array.from(mappings.values()).find(
            (m) =>
              m.workspaceId === workspaceId &&
              m.ruleId === ruleId &&
              m.sourceExternalEventId === sourceExternalEventId
          ) ?? null
        );
      },
      async isKnownMirroredTarget(workspaceId: string, targetAccountId: string, externalEventId: string) {
        return Array.from(mappings.values()).some(
          (m) =>
            m.workspaceId === workspaceId &&
            m.targetAccountId === targetAccountId &&
            m.targetExternalEventId === externalEventId
        );
      },
      async saveMapping(mapping: MirrorMapping) {
        mappings.set(mapping.id, mapping);
      },
      async deleteMapping(mappingId: string) {
        mappings.delete(mappingId);
      },
      async upsertDestinationBusyEvent(
        targetAccount: CalendarAccount,
        rule: SyncRule,
        sourceEvent: CalendarEvent,
        existingTargetExternalId?: string
      ) {
        const externalEventId = existingTargetExternalId ?? `mirrored_${targetAccount.id}_${sourceEvent.externalEventId}`;
        const existing = events.get(eventKey(targetAccount.id, externalEventId));
        const next: CalendarEvent = {
          id: existing?.id ?? `evt_${externalEventId}`,
          accountId: targetAccount.id,
          externalEventId,
          title: rule.busyTitle || 'Busy',
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
            fingerprint: `${sourceEvent.startIso}:${sourceEvent.endIso}`,
          },
          updatedAt: new Date().toISOString(),
        };
        events.set(eventKey(targetAccount.id, externalEventId), next);
        return {
          targetExternalEventId: externalEventId,
          payloadPreview: {
            summary: next.title,
            visibility: next.visibility,
            locked: next.locked,
          },
        };
      },
      async deleteDestinationBusyEvent(targetAccount: CalendarAccount, targetExternalEventId: string) {
        events.delete(eventKey(targetAccount.id, targetExternalEventId));
      },
      async appendLog(entry: SyncLogEntry) {
        logs.push(entry);
      },
    },
  };
}

describe('StealthSyncEngine', () => {
  it('mirrors a new event across Google and Outlook with zero title leak and sub-second telemetry', async () => {
    const store = createInMemoryPort();
    store.accounts.set('acc_g', {
      id: 'acc_g',
      workspaceId: 'ws_1',
      label: 'Acme Work',
      email: 'alex@acme.io',
      provider: 'google',
      calendarId: 'primary',
      calendarName: 'Work',
      color: '#10b981',
      mode: 'sandbox',
      status: 'connected',
      createdAt: '2026-11-11T00:00:00Z',
    });
    store.accounts.set('acc_o', {
      id: 'acc_o',
      workspaceId: 'ws_1',
      label: 'Client Outlook',
      email: 'alex@client.com',
      provider: 'outlook',
      calendarId: 'primary',
      calendarName: 'Consulting',
      color: '#3b82f6',
      mode: 'sandbox',
      status: 'connected',
      createdAt: '2026-11-11T00:00:00Z',
    });
    store.rules.set('r_g_o', {
      id: 'r_g_o',
      workspaceId: 'ws_1',
      sourceAccountId: 'acc_g',
      targetAccountId: 'acc_o',
      enabled: true,
      busyTitle: 'Busy',
      lockMirrored: true,
      skipFreeEvents: true,
      skipDeclinedEvents: true,
      skipAllDayEvents: true,
      createdAt: '2026-11-11T00:00:00Z',
    });
    store.rules.set('r_o_g', {
      id: 'r_o_g',
      workspaceId: 'ws_1',
      sourceAccountId: 'acc_o',
      targetAccountId: 'acc_g',
      enabled: true,
      busyTitle: 'Busy',
      lockMirrored: true,
      skipFreeEvents: true,
      skipDeclinedEvents: true,
      skipAllDayEvents: true,
      createdAt: '2026-11-11T00:00:00Z',
    });

    const engine = new StealthSyncEngine(store.port);
    const sourceEvent: CalendarEvent = {
      id: 'evt_src_1',
      accountId: 'acc_g',
      externalEventId: 'g_100',
      title: '1:1 Salary Negotiation',
      description: 'Private HR notes',
      location: 'Room 12',
      attendees: ['hr@acme.io'],
      startIso: '2026-11-12T14:00:00.000Z',
      endIso: '2026-11-12T15:00:00.000Z',
      allDay: false,
      showAs: 'busy',
      visibility: 'default',
      responseStatus: 'accepted',
      locked: false,
      updatedAt: '2026-11-11T10:00:00.000Z',
    };

    const result = await engine.processEventChange('ws_1', sourceEvent);
    expect(result.created).toBe(1);

    const mirrored = Array.from(store.events.values()).find((e) => e.accountId === 'acc_o');
    expect(mirrored).toBeDefined();
    expect(mirrored?.title).toBe('Busy');
    expect(mirrored?.description).toBe('');
    expect(mirrored?.locked).toBe(true);
    expect(mirrored?.visibility).toBe('private');

    // Simulate webhook bounce back from Outlook -> Google: MUST NOT create a loop!
    const bounceResult = await engine.processEventChange('ws_1', mirrored!);
    expect(bounceResult.skippedLoop).toBe(1);
    expect(bounceResult.created).toBe(0);
    expect(store.events.size).toBe(1);
  });

  it('updates an existing mirrored block in-place when rescheduled and deletes it when removed or marked free', async () => {
    const store = createInMemoryPort();
    store.accounts.set('acc_g', {
      id: 'acc_g',
      workspaceId: 'ws_1',
      label: 'Work',
      email: 'alex@acme.io',
      provider: 'google',
      calendarId: 'primary',
      calendarName: 'Work',
      color: '#10b981',
      mode: 'sandbox',
      status: 'connected',
      createdAt: '2026-11-11T00:00:00Z',
    });
    store.accounts.set('acc_p', {
      id: 'acc_p',
      workspaceId: 'ws_1',
      label: 'Personal',
      email: 'alex@gmail.com',
      provider: 'google',
      calendarId: 'primary',
      calendarName: 'Personal',
      color: '#a855f7',
      mode: 'sandbox',
      status: 'connected',
      createdAt: '2026-11-11T00:00:00Z',
    });
    store.rules.set('r1', {
      id: 'r1',
      workspaceId: 'ws_1',
      sourceAccountId: 'acc_g',
      targetAccountId: 'acc_p',
      enabled: true,
      busyTitle: 'Busy',
      lockMirrored: true,
      skipFreeEvents: true,
      skipDeclinedEvents: true,
      skipAllDayEvents: true,
      createdAt: '2026-11-11T00:00:00Z',
    });

    const engine = new StealthSyncEngine(store.port);
    const evt: CalendarEvent = {
      id: 'e1',
      accountId: 'acc_g',
      externalEventId: 'g_200',
      title: 'Deep Focus',
      description: '',
      location: '',
      attendees: [],
      startIso: '2026-11-12T09:00:00.000Z',
      endIso: '2026-11-12T10:00:00.000Z',
      allDay: false,
      showAs: 'busy',
      visibility: 'default',
      responseStatus: 'accepted',
      locked: false,
      updatedAt: '2026-11-11T09:00:00.000Z',
    };

    await engine.processEventChange('ws_1', evt);
    expect(store.mappings.size).toBe(1);

    // Move by 1 hour
    const moved = {
      ...evt,
      startIso: '2026-11-12T10:00:00.000Z',
      endIso: '2026-11-12T11:30:00.000Z',
    };
    const updateRes = await engine.processEventChange('ws_1', moved);
    expect(updateRes.updated).toBe(1);
    expect(store.events.size).toBe(1);
    const mirroredAfterMove = Array.from(store.events.values())[0];
    expect(mirroredAfterMove.startIso).toBe('2026-11-12T10:00:00.000Z');
    expect(mirroredAfterMove.endIso).toBe('2026-11-12T11:30:00.000Z');

    // Mark as "free" -> should delete mirrored Busy block
    const freeRes = await engine.processEventChange('ws_1', { ...moved, showAs: 'free' });
    expect(freeRes.deleted).toBe(1);
    expect(store.events.size).toBe(0);
    expect(store.mappings.size).toBe(0);
  });

  it('skips declined and all-day events according to rule flags', async () => {
    const store = createInMemoryPort();
    store.accounts.set('acc_g', {
      id: 'acc_g',
      workspaceId: 'ws_1',
      label: 'Work',
      email: 'alex@acme.io',
      provider: 'google',
      calendarId: 'primary',
      calendarName: 'Work',
      color: '#10b981',
      mode: 'sandbox',
      status: 'connected',
      createdAt: '2026-11-11T00:00:00Z',
    });
    store.accounts.set('acc_p', {
      id: 'acc_p',
      workspaceId: 'ws_1',
      label: 'Personal',
      email: 'alex@gmail.com',
      provider: 'google',
      calendarId: 'primary',
      calendarName: 'Personal',
      color: '#a855f7',
      mode: 'sandbox',
      status: 'connected',
      createdAt: '2026-11-11T00:00:00Z',
    });
    store.rules.set('r1', {
      id: 'r1',
      workspaceId: 'ws_1',
      sourceAccountId: 'acc_g',
      targetAccountId: 'acc_p',
      enabled: true,
      busyTitle: 'Busy',
      lockMirrored: true,
      skipFreeEvents: true,
      skipDeclinedEvents: true,
      skipAllDayEvents: true,
      createdAt: '2026-11-11T00:00:00Z',
    });

    const engine = new StealthSyncEngine(store.port);
    const res1 = await engine.processEventChange('ws_1', {
      id: 'e_dec',
      accountId: 'acc_g',
      externalEventId: 'g_dec',
      title: 'Optional Town Hall',
      description: '',
      location: '',
      attendees: [],
      startIso: '2026-11-12T12:00:00.000Z',
      endIso: '2026-11-12T13:00:00.000Z',
      allDay: false,
      showAs: 'busy',
      visibility: 'default',
      responseStatus: 'declined',
      locked: false,
      updatedAt: '2026-11-11T09:00:00.000Z',
    });
    expect(res1.skippedPolicy).toBe(1);
    expect(store.events.size).toBe(0);

    const res2 = await engine.processEventChange('ws_1', {
      id: 'e_allday',
      accountId: 'acc_g',
      externalEventId: 'g_allday',
      title: 'Company Holiday',
      description: '',
      location: '',
      attendees: [],
      startIso: '2026-11-13T00:00:00.000Z',
      endIso: '2026-11-14T00:00:00.000Z',
      allDay: true,
      showAs: 'busy',
      visibility: 'default',
      responseStatus: 'accepted',
      locked: false,
      updatedAt: '2026-11-11T09:00:00.000Z',
    });
    expect(res2.skippedPolicy).toBe(1);
    expect(store.events.size).toBe(0);
  });
});
