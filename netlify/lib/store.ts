import { getStore } from '@netlify/blobs';
import { del as vercelBlobDel, get as vercelBlobGet, put as vercelBlobPut } from '@vercel/blob';
import {
  buildGoogleBusyPayload,
  buildOutlookBusyPayload,
  sanitizeEventForMirror,
} from '../../shared/sanitizer';
import { StealthSyncEngine, type SyncStoragePort } from '../../shared/syncEngine';
import type {
  CalendarAccount,
  CalendarEvent,
  CalendarProvider,
  MirrorMapping,
  SyncLogEntry,
  SyncRule,
  Workspace,
} from '../../shared/types';
import {
  deleteGoogleBusyEvent,
  upsertGoogleBusyEvent,
} from './providers/google';
import {
  deleteMicrosoftBusyEvent,
  upsertMicrosoftBusyEvent,
} from './providers/microsoft';

export interface WorkspaceState {
  workspace: Workspace;
  accounts: CalendarAccount[];
  rules: SyncRule[];
  events: CalendarEvent[];
  mappings: MirrorMapping[];
  logs: SyncLogEntry[];
}

export interface OAuthStateRecord {
  workspaceId: string;
  provider: CalendarProvider;
  label?: string;
  color?: string;
  origin: string;
  createdAt: string;
}

export interface ChannelLookupRecord {
  workspaceId: string;
  accountId: string;
}

export interface StoreBackend {
  getJSON<T>(key: string): Promise<T | null>;
  setJSON(key: string, value: unknown): Promise<void>;
  deleteKey(key: string): Promise<void>;
}

class MemoryStoreBackend implements StoreBackend {
  private readonly map = new Map<string, unknown>();

  async getJSON<T>(key: string): Promise<T | null> {
    if (!this.map.has(key)) return null;
    return structuredClone(this.map.get(key)) as T;
  }

  async setJSON(key: string, value: unknown): Promise<void> {
    this.map.set(key, structuredClone(value));
  }

  async deleteKey(key: string): Promise<void> {
    this.map.delete(key);
  }
}

class VercelBlobStoreBackend implements StoreBackend {
  private readonly fallback = new MemoryStoreBackend();

  private toPath(key: string): string {
    return `stealthmirror/${key.replace(/^\/+/, '')}.json`;
  }

  async getJSON<T>(key: string): Promise<T | null> {
    try {
      const result = await vercelBlobGet(this.toPath(key), {
        access: 'private',
        useCache: false,
      });
      if (!result || result.statusCode !== 200) {
        return null;
      }
      const raw = await new Response(result.stream).text();
      return JSON.parse(raw) as T;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      if (msg.toLowerCase().includes('not found') || msg.includes('404')) {
        return null;
      }
      console.error('VercelBlobStoreBackend.getJSON error, using fallback:', err);
      return this.fallback.getJSON<T>(key);
    }
  }

  async setJSON(key: string, value: unknown): Promise<void> {
    try {
      await vercelBlobPut(this.toPath(key), JSON.stringify(value), {
        access: 'private',
        addRandomSuffix: false,
        allowOverwrite: true,
        contentType: 'application/json',
      });
      await this.fallback.setJSON(key, value);
    } catch (err) {
      console.error('VercelBlobStoreBackend.setJSON error, using fallback:', err);
      await this.fallback.setJSON(key, value);
    }
  }

  async deleteKey(key: string): Promise<void> {
    try {
      await vercelBlobDel(this.toPath(key));
    } catch (err) {
      console.error('VercelBlobStoreBackend.deleteKey error:', err);
    }
    await this.fallback.deleteKey(key);
  }
}

class NetlifyBlobStoreBackend implements StoreBackend {
  private readonly fallback = new MemoryStoreBackend();

  private getBlobStore() {
    return getStore({ name: 'stealthmirror-state', consistency: 'strong' });
  }

  async getJSON<T>(key: string): Promise<T | null> {
    try {
      const val = await this.getBlobStore().get(key, { type: 'json' });
      return (val as T) ?? null;
    } catch {
      return this.fallback.getJSON<T>(key);
    }
  }

  async setJSON(key: string, value: unknown): Promise<void> {
    try {
      await this.getBlobStore().setJSON(key, value);
    } catch {
      await this.fallback.setJSON(key, value);
    }
  }

  async deleteKey(key: string): Promise<void> {
    try {
      await this.getBlobStore().delete(key);
    } catch {
      await this.fallback.deleteKey(key);
    }
  }
}

function utcTodayAt(hour: number, minute = 0, dayOffset = 0): string {
  const now = new Date();
  const dt = new Date(
    Date.UTC(
      now.getUTCFullYear(),
      now.getUTCMonth(),
      now.getUTCDate() + dayOffset,
      hour,
      minute,
      0,
      0
    )
  );
  return dt.toISOString();
}

export class StealthMirrorStore {
  constructor(private readonly backend: StoreBackend) {}

  private wsKey(workspaceId: string): string {
    return `ws/${workspaceId}/state`;
  }

  async getWorkspaceState(workspaceId: string): Promise<WorkspaceState | null> {
    return this.backend.getJSON<WorkspaceState>(this.wsKey(workspaceId));
  }

  async saveWorkspaceState(state: WorkspaceState): Promise<void> {
    if (state.logs.length > 150) {
      state.logs = state.logs.slice(0, 150);
    }
    await this.backend.setJSON(this.wsKey(state.workspace.id), state);
    await this.registerWorkspaceId(state.workspace.id);
  }

  async registerWorkspaceId(workspaceId: string): Promise<void> {
    const existing = (await this.backend.getJSON<string[]>('index/workspaces')) ?? [];
    if (!existing.includes(workspaceId)) {
      existing.push(workspaceId);
      await this.backend.setJSON('index/workspaces', existing);
    }
  }

  async listWorkspaceIds(): Promise<string[]> {
    return (await this.backend.getJSON<string[]>('index/workspaces')) ?? [];
  }

  async saveOAuthState(nonce: string, record: OAuthStateRecord): Promise<void> {
    await this.backend.setJSON(`oauth/${nonce}`, record);
  }

  async consumeOAuthState(nonce: string): Promise<OAuthStateRecord | null> {
    const key = `oauth/${nonce}`;
    const record = await this.backend.getJSON<OAuthStateRecord>(key);
    if (record) await this.backend.deleteKey(key);
    return record;
  }

  async saveChannelLookup(
    channelOrSubId: string,
    record: ChannelLookupRecord
  ): Promise<void> {
    await this.backend.setJSON(`channel/${channelOrSubId}`, record);
  }

  async getChannelLookup(
    channelOrSubId: string
  ): Promise<ChannelLookupRecord | null> {
    return this.backend.getJSON<ChannelLookupRecord>(`channel/${channelOrSubId}`);
  }

  async getOrSeedWorkspace(workspaceId: string): Promise<WorkspaceState> {
    const existing = await this.getWorkspaceState(workspaceId);
    if (existing) return existing;

    if (workspaceId === 'demo') {
      const seeded = await this.buildDemoState();
      await this.saveWorkspaceState(seeded);
      return seeded;
    }

    const fresh: WorkspaceState = {
      workspace: {
        id: workspaceId,
        name: workspaceId === 'personal' ? 'Personal Workspace' : `Workspace ${workspaceId}`,
        slug: workspaceId,
        plan: 'monthly_5',
        createdAt: new Date().toISOString(),
      },
      accounts: [],
      rules: [],
      events: [],
      mappings: [],
      logs: [],
    };
    await this.saveWorkspaceState(fresh);
    return fresh;
  }

  createSyncPort(state: WorkspaceState): SyncStoragePort {
    return {
      async getRulesForSource(workspaceId: string, sourceAccountId: string) {
        return state.rules.filter(
          (r) =>
            r.workspaceId === workspaceId &&
            r.sourceAccountId === sourceAccountId &&
            r.enabled
        );
      },
      async getAccount(accountId: string) {
        return state.accounts.find((a) => a.id === accountId) ?? null;
      },
      async getMapping(
        workspaceId: string,
        ruleId: string,
        sourceExternalEventId: string
      ) {
        return (
          state.mappings.find(
            (m) =>
              m.workspaceId === workspaceId &&
              m.ruleId === ruleId &&
              m.sourceExternalEventId === sourceExternalEventId
          ) ?? null
        );
      },
      async isKnownMirroredTarget(
        workspaceId: string,
        targetAccountId: string,
        externalEventId: string
      ) {
        return state.mappings.some(
          (m) =>
            m.workspaceId === workspaceId &&
            m.targetAccountId === targetAccountId &&
            m.targetExternalEventId === externalEventId
        );
      },
      async saveMapping(mapping: MirrorMapping) {
        const idx = state.mappings.findIndex((m) => m.id === mapping.id);
        if (idx >= 0) state.mappings[idx] = mapping;
        else state.mappings.push(mapping);
      },
      async deleteMapping(mappingId: string) {
        state.mappings = state.mappings.filter((m) => m.id !== mappingId);
      },
      async upsertDestinationBusyEvent(
        targetAccount: CalendarAccount,
        rule: SyncRule,
        sourceEvent: CalendarEvent,
        existingTargetExternalId?: string
      ) {
        let targetExternalEventId = existingTargetExternalId;
        let payloadPreview: Record<string, unknown>;

        if (targetAccount.mode === 'oauth' && targetAccount.accessToken) {
          if (targetAccount.provider === 'google') {
            const res = await upsertGoogleBusyEvent(
              targetAccount,
              rule,
              sourceEvent,
              existingTargetExternalId
            );
            targetExternalEventId = res.targetExternalEventId;
            payloadPreview = res.payloadPreview;
          } else {
            const res = await upsertMicrosoftBusyEvent(
              targetAccount,
              rule,
              sourceEvent,
              existingTargetExternalId
            );
            targetExternalEventId = res.targetExternalEventId;
            payloadPreview = res.payloadPreview;
          }
        } else {
          payloadPreview =
            targetAccount.provider === 'google'
              ? (buildGoogleBusyPayload(sourceEvent, rule) as unknown as Record<
                  string,
                  unknown
                >)
              : (buildOutlookBusyPayload(sourceEvent, rule) as unknown as Record<
                  string,
                  unknown
                >);
          targetExternalEventId =
            existingTargetExternalId ??
            `sm_${targetAccount.id}_${sourceEvent.externalEventId}`;
        }

        const existingIdx = state.events.findIndex(
          (e) =>
            e.accountId === targetAccount.id &&
            e.externalEventId === targetExternalEventId
        );
        const existingEvt = existingIdx >= 0 ? state.events[existingIdx] : undefined;
        const mirroredEvent = sanitizeEventForMirror(
          sourceEvent,
          rule,
          existingEvt?.id,
          targetExternalEventId
        );

        if (existingIdx >= 0) {
          state.events[existingIdx] = mirroredEvent;
        } else {
          state.events.push(mirroredEvent);
        }

        return {
          targetExternalEventId,
          payloadPreview,
        };
      },
      async deleteDestinationBusyEvent(
        targetAccount: CalendarAccount,
        targetExternalEventId: string
      ) {
        if (targetAccount.mode === 'oauth' && targetAccount.accessToken) {
          if (targetAccount.provider === 'google') {
            await deleteGoogleBusyEvent(targetAccount, targetExternalEventId);
          } else {
            await deleteMicrosoftBusyEvent(targetAccount, targetExternalEventId);
          }
        }
        state.events = state.events.filter(
          (e) =>
            !(
              e.accountId === targetAccount.id &&
              e.externalEventId === targetExternalEventId
            )
        );
      },
      async appendLog(entry: SyncLogEntry) {
        state.logs.unshift(entry);
        if (state.logs.length > 150) {
          state.logs = state.logs.slice(0, 150);
        }
      },
    };
  }

  private async buildDemoState(): Promise<WorkspaceState> {
    const nowIso = new Date().toISOString();
    const accounts: CalendarAccount[] = [
      {
        id: 'acc_work',
        workspaceId: 'demo',
        label: 'Acme Work',
        email: 'alex@acme.io',
        provider: 'google',
        calendarId: 'primary',
        calendarName: 'Primary Work',
        color: '#0EA5E9',
        mode: 'sandbox',
        status: 'connected',
        watchChannelId: 'chan_demo_work',
        watchExpiresAt: new Date(Date.now() + 6 * 24 * 3600_000).toISOString(),
        createdAt: nowIso,
      },
      {
        id: 'acc_client',
        workspaceId: 'demo',
        label: 'Client Advisory',
        email: 'alex@advisory.co',
        provider: 'outlook',
        calendarId: 'primary',
        calendarName: 'Microsoft 365',
        color: '#0F766E',
        mode: 'sandbox',
        status: 'connected',
        watchChannelId: 'sub_demo_client',
        watchExpiresAt: new Date(Date.now() + 2 * 24 * 3600_000).toISOString(),
        createdAt: nowIso,
      },
      {
        id: 'acc_personal',
        workspaceId: 'demo',
        label: 'Personal Gmail',
        email: 'alex.personal@gmail.com',
        provider: 'google',
        calendarId: 'primary',
        calendarName: 'Personal',
        color: '#10B981',
        mode: 'sandbox',
        status: 'connected',
        watchChannelId: 'chan_demo_personal',
        watchExpiresAt: new Date(Date.now() + 6 * 24 * 3600_000).toISOString(),
        createdAt: nowIso,
      },
    ];

    const rules: SyncRule[] = [
      {
        id: 'rule_work_client',
        workspaceId: 'demo',
        sourceAccountId: 'acc_work',
        targetAccountId: 'acc_client',
        enabled: true,
        busyTitle: 'Busy',
        lockMirrored: true,
        skipFreeEvents: true,
        skipDeclinedEvents: true,
        skipAllDayEvents: true,
        createdAt: nowIso,
      },
      {
        id: 'rule_work_personal',
        workspaceId: 'demo',
        sourceAccountId: 'acc_work',
        targetAccountId: 'acc_personal',
        enabled: true,
        busyTitle: 'Busy',
        lockMirrored: true,
        skipFreeEvents: true,
        skipDeclinedEvents: true,
        skipAllDayEvents: true,
        createdAt: nowIso,
      },
      {
        id: 'rule_client_work',
        workspaceId: 'demo',
        sourceAccountId: 'acc_client',
        targetAccountId: 'acc_work',
        enabled: true,
        busyTitle: 'Busy',
        lockMirrored: true,
        skipFreeEvents: true,
        skipDeclinedEvents: true,
        skipAllDayEvents: true,
        createdAt: nowIso,
      },
      {
        id: 'rule_personal_work',
        workspaceId: 'demo',
        sourceAccountId: 'acc_personal',
        targetAccountId: 'acc_work',
        enabled: true,
        busyTitle: 'Busy',
        lockMirrored: true,
        skipFreeEvents: true,
        skipDeclinedEvents: true,
        skipAllDayEvents: true,
        createdAt: nowIso,
      },
    ];

    const seedEvents: CalendarEvent[] = [
      {
        id: 'evt_seed_1',
        accountId: 'acc_work',
        externalEventId: 'g_seed_board_prep',
        title: 'Q4 Board Deck & Runway Review',
        description: 'Confidential doc: https://docs.google.com/secret-board-deck',
        location: 'HQ Glass Room 4A',
        attendees: ['cfo@acme.io', 'founder@acme.io'],
        startIso: utcTodayAt(9, 30, 0),
        endIso: utcTodayAt(10, 30, 0),
        allDay: false,
        showAs: 'busy',
        visibility: 'default',
        responseStatus: 'accepted',
        locked: false,
        updatedAt: nowIso,
      },
      {
        id: 'evt_seed_2',
        accountId: 'acc_client',
        externalEventId: 'o_seed_arch_review',
        title: 'Fintech SOC2 Architecture Deep Dive',
        description: 'Teams link: https://teams.microsoft.com/l/meetup-join/xyz',
        location: 'Microsoft Teams',
        attendees: ['cto@fintechclient.com'],
        startIso: utcTodayAt(13, 0, 0),
        endIso: utcTodayAt(14, 0, 0),
        allDay: false,
        showAs: 'busy',
        visibility: 'default',
        responseStatus: 'accepted',
        locked: false,
        updatedAt: nowIso,
      },
      {
        id: 'evt_seed_3',
        accountId: 'acc_personal',
        externalEventId: 'g_seed_dentist',
        title: 'Dentist Appointment — Dr. Sarah Chen',
        description: 'Bring dental insurance card',
        location: '450 Sutter St Suite 900',
        attendees: [],
        startIso: utcTodayAt(16, 0, 0),
        endIso: utcTodayAt(17, 0, 0),
        allDay: false,
        showAs: 'busy',
        visibility: 'private',
        responseStatus: 'accepted',
        locked: false,
        updatedAt: nowIso,
      },
      {
        id: 'evt_seed_4',
        accountId: 'acc_work',
        externalEventId: 'g_seed_investor',
        title: 'Series A Partner Pitch (Sequoia)',
        description: 'Private deck & cap table walkthrough',
        location: 'Sand Hill Rd / Zoom',
        attendees: ['partner@sequoia.com'],
        startIso: utcTodayAt(11, 0, 1),
        endIso: utcTodayAt(12, 0, 1),
        allDay: false,
        showAs: 'busy',
        visibility: 'default',
        responseStatus: 'accepted',
        locked: false,
        updatedAt: nowIso,
      },
      {
        id: 'evt_legacy_reclaim',
        accountId: 'acc_work',
        externalEventId: 'g_legacy_reclaim_1',
        title: 'Decompress / Smart Habit [Reclaim.ai]',
        description: 'Auto-scheduled by Reclaim.ai — click to reschedule your day',
        location: '',
        attendees: [],
        startIso: utcTodayAt(12, 0, 0),
        endIso: utcTodayAt(12, 30, 0),
        allDay: false,
        showAs: 'free',
        visibility: 'default',
        responseStatus: 'accepted',
        locked: false,
        updatedAt: nowIso,
      },
      {
        id: 'evt_legacy_syncgene',
        accountId: 'acc_personal',
        externalEventId: 'g_legacy_syncgene_1',
        title: 'Busy (Synced via SyncGene)',
        description: 'Synced 24 minutes ago via SyncGene pair #2',
        location: '',
        attendees: [],
        startIso: utcTodayAt(15, 0, 1),
        endIso: utcTodayAt(15, 30, 1),
        allDay: false,
        showAs: 'busy',
        visibility: 'default',
        responseStatus: 'accepted',
        locked: false,
        updatedAt: nowIso,
      },
    ];

    const state: WorkspaceState = {
      workspace: {
        id: 'demo',
        name: 'StealthMirror Live Workspace',
        slug: 'demo',
        plan: 'monthly_5',
        createdAt: nowIso,
      },
      accounts,
      rules,
      events: [...seedEvents],
      mappings: [],
      logs: [],
    };

    const engine = new StealthSyncEngine(this.createSyncPort(state));
    for (const evt of seedEvents) {
      if (evt.id.startsWith('evt_seed_')) {
        await engine.processEventChange('demo', evt);
      }
    }

    return state;
  }
}

export function createMemoryStore(): StealthMirrorStore {
  return new StealthMirrorStore(new MemoryStoreBackend());
}

let sharedStore: StealthMirrorStore | null = null;
export function getStoreClient(): StealthMirrorStore {
  if (!sharedStore) {
    const useVercelBlob = Boolean(
      typeof process !== 'undefined' &&
        (process.env.BLOB_READ_WRITE_TOKEN || process.env.BLOB_STORE_ID)
    );
    sharedStore = new StealthMirrorStore(
      useVercelBlob ? new VercelBlobStoreBackend() : new NetlifyBlobStoreBackend()
    );
  }
  return sharedStore;
}
