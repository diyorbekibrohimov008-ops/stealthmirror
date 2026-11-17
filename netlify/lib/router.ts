import { z } from 'zod';
import {
  detectLegacySyncClutter,
  isStealthManagedEvent,
  resolveSafeBusyTitle,
} from '../../shared/sanitizer';
import { StealthSyncEngine } from '../../shared/syncEngine';
import type {
  BillingPlan,
  CalendarAccount,
  CalendarEvent,
  SyncRule,
} from '../../shared/types';
import { getEnv, getOAuthConfigStatus } from './env';
import {
  buildGoogleAuthUrl,
  exchangeGoogleAuthCode,
  fetchGoogleIncrementalChanges,
  fetchGoogleUserProfile,
  googleEventToCanonical,
  parseGoogleWebhookHeaders,
  refreshGoogleAccessToken,
  watchGoogleCalendar,
} from './providers/google';
import {
  buildMicrosoftAuthUrl,
  createMicrosoftGraphSubscription,
  exchangeMicrosoftAuthCode,
  fetchMicrosoftDeltaChanges,
  fetchMicrosoftUserProfile,
  handleMicrosoftValidationHandshake,
  outlookEventToCanonical,
  parseMicrosoftWebhookNotifications,
  refreshMicrosoftAccessToken,
  renewMicrosoftGraphSubscription,
} from './providers/microsoft';
import {
  getStoreClient,
  type StealthMirrorStore,
  type WorkspaceState,
} from './store';

function json(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      'content-type': 'application/json; charset=utf-8',
      'cache-control': 'no-store',
    },
  });
}

function redirect(location: string): Response {
  return new Response(null, {
    status: 302,
    headers: {
      location,
      'cache-control': 'no-store',
    },
  });
}

function makeId(prefix: string): string {
  return `${prefix}_${crypto.randomUUID().slice(0, 12)}`;
}

function computeWorkspaceStats(state: WorkspaceState) {
  const mirroredEvents = state.events.filter((e) => isStealthManagedEvent(e));
  const legacyClutter = state.events.filter((e) => detectLegacySyncClutter(e) !== null);
  const completedLogs = state.logs.filter(
    (l) => l.action === 'created' || l.action === 'updated' || l.action === 'deleted'
  );
  const avgLatencyMs =
    completedLogs.length > 0
      ? Math.round(
          completedLogs.reduce((acc, item) => acc + item.latencyMs, 0) /
            completedLogs.length
        )
      : 8;
  const loopsBlocked = state.logs.filter((l) => l.action === 'skipped_loop').length;
  const pairCount = Math.max(1, Math.floor((state.accounts.length * (state.accounts.length - 1)) / 2));

  return {
    accountCount: state.accounts.length,
    activeRuleCount: state.rules.filter((r) => r.enabled).length,
    mirroredBlockCount: mirroredEvents.length,
    legacyClutterCount: legacyClutter.length,
    avgLatencyMs,
    loopsBlocked,
    leaksCount: 0,
    estimatedSyncGeneMonthlyUsd: pairCount * 10,
  };
}

function buildPublicWorkspacePayload(state: WorkspaceState, origin?: string) {
  return {
    workspace: state.workspace,
    accounts: state.accounts.map((a) => ({
      ...a,
      accessToken: undefined,
      refreshToken: undefined,
    })),
    rules: state.rules,
    events: state.events,
    mappings: state.mappings,
    logs: state.logs,
    stats: computeWorkspaceStats(state),
    oauth: getOAuthConfigStatus(origin),
  };
}

const createAccountSchema = z.object({
  workspaceId: z.string().min(1).default('demo'),
  provider: z.enum(['google', 'outlook']),
  label: z.string().min(1).max(60),
  email: z.string().email(),
  calendarName: z.string().min(1).max(60).default('Primary'),
  color: z.string().min(4).max(20).default('#0EA5E9'),
  autoMeshWithExisting: z.boolean().optional().default(true),
});

const createRuleSchema = z.object({
  workspaceId: z.string().min(1).default('demo'),
  sourceAccountId: z.string().min(1),
  targetAccountId: z.string().min(1),
  busyTitle: z.string().max(24).default('Busy'),
  lockMirrored: z.boolean().default(true),
  skipFreeEvents: z.boolean().default(true),
  skipDeclinedEvents: z.boolean().default(true),
  skipAllDayEvents: z.boolean().default(true),
});

const updateRuleSchema = z.object({
  workspaceId: z.string().min(1).default('demo'),
  enabled: z.boolean().optional(),
  busyTitle: z.string().max(24).optional(),
  lockMirrored: z.boolean().optional(),
  skipFreeEvents: z.boolean().optional(),
  skipDeclinedEvents: z.boolean().optional(),
  skipAllDayEvents: z.boolean().optional(),
});

const createEventSchema = z.object({
  workspaceId: z.string().min(1).default('demo'),
  accountId: z.string().min(1),
  title: z.string().min(1).max(160),
  description: z.string().max(2000).default(''),
  location: z.string().max(240).default(''),
  attendees: z.array(z.string()).default([]),
  startIso: z.string().min(1),
  endIso: z.string().min(1),
  allDay: z.boolean().default(false),
  showAs: z.enum(['busy', 'free', 'tentative', 'oof']).default('busy'),
  responseStatus: z
    .enum(['accepted', 'tentative', 'declined', 'needsAction', 'organizer'])
    .default('accepted'),
});

const updateEventSchema = z.object({
  workspaceId: z.string().min(1).default('demo'),
  title: z.string().min(1).max(160).optional(),
  description: z.string().max(2000).optional(),
  location: z.string().max(240).optional(),
  startIso: z.string().min(1).optional(),
  endIso: z.string().min(1).optional(),
  allDay: z.boolean().optional(),
  showAs: z.enum(['busy', 'free', 'tentative', 'oof']).optional(),
  responseStatus: z
    .enum(['accepted', 'tentative', 'declined', 'needsAction', 'organizer'])
    .optional(),
});

const billingSelectSchema = z.object({
  workspaceId: z.string().min(1).default('demo'),
  plan: z.enum(['trial', 'monthly_5', 'lifetime_39']),
});

export async function renewExpiringWatches(store: StealthMirrorStore) {
  const workspaceIds = await store.listWorkspaceIds();
  let renewed = 0;
  let errors = 0;

  for (const workspaceId of workspaceIds) {
    const state = await store.getWorkspaceState(workspaceId);
    if (!state) continue;
    let changed = false;

    for (const acc of state.accounts) {
      if (acc.mode !== 'oauth' || !acc.accessToken || !acc.watchExpiresAt) continue;
      const msRemaining = new Date(acc.watchExpiresAt).getTime() - Date.now();
      if (msRemaining > 24 * 3600_000) continue;

      try {
        if (acc.provider === 'google') {
          const clientId = getEnv('GOOGLE_CLIENT_ID');
          const clientSecret = getEnv('GOOGLE_CLIENT_SECRET');
          if (acc.refreshToken && clientId && clientSecret) {
            const refreshed = await refreshGoogleAccessToken({
              refreshToken: acc.refreshToken,
              clientId,
              clientSecret,
            });
            acc.accessToken = refreshed.accessToken;
          }
          const appUrl = getEnv('APP_URL');
          if (appUrl) {
            const newChannelId = makeId('gchan');
            const watchRes = await watchGoogleCalendar({
              accessToken: acc.accessToken,
              calendarId: acc.calendarId,
              channelId: newChannelId,
              webhookUrl: `${appUrl}/api/webhooks/google`,
              token: `${workspaceId}:${acc.id}`,
            });
            acc.watchChannelId = newChannelId;
            acc.watchResourceId = watchRes.resourceId;
            acc.watchExpiresAt = watchRes.expirationIso;
            await store.saveChannelLookup(newChannelId, {
              workspaceId,
              accountId: acc.id,
            });
            changed = true;
            renewed += 1;
          }
        } else if (acc.provider === 'outlook' && acc.watchChannelId) {
          const clientId = getEnv('MICROSOFT_CLIENT_ID');
          const clientSecret = getEnv('MICROSOFT_CLIENT_SECRET');
          if (acc.refreshToken && clientId && clientSecret) {
            const refreshed = await refreshMicrosoftAccessToken({
              refreshToken: acc.refreshToken,
              clientId,
              clientSecret,
            });
            acc.accessToken = refreshed.accessToken;
            if (refreshed.refreshToken) acc.refreshToken = refreshed.refreshToken;
          }
          const renewedSub = await renewMicrosoftGraphSubscription({
            accessToken: acc.accessToken,
            subscriptionId: acc.watchChannelId,
          });
          acc.watchExpiresAt = renewedSub.expirationDateTime;
          changed = true;
          renewed += 1;
        }
      } catch {
        errors += 1;
      }
    }

    if (changed) {
      await store.saveWorkspaceState(state);
    }
  }

  return { renewed, errors, checkedWorkspaces: workspaceIds.length };
}

export async function handleApiRequest(
  req: Request,
  providedStore?: StealthMirrorStore
): Promise<Response> {
  const store = providedStore ?? getStoreClient();
  const url = new URL(req.url);
  const rewrittenPath = url.searchParams.get('__path')?.trim();
  const path = rewrittenPath
    ? rewrittenPath.startsWith('/api/')
      ? rewrittenPath
      : rewrittenPath.startsWith('/')
      ? `/api${rewrittenPath}`
      : `/api/${rewrittenPath}`
    : url.pathname;
  const method = req.method.toUpperCase();
  const origin = `${url.protocol}//${url.host}`;

  if (method === 'GET' && path === '/api/health') {
    return json({
      ok: true,
      service: 'stealthmirror',
      timestamp: new Date().toISOString(),
      oauth: getOAuthConfigStatus(origin),
    });
  }

  if (method === 'GET' && path === '/api/workspace') {
    const workspaceId = url.searchParams.get('workspaceId')?.trim() || 'demo';
    const state = await store.getOrSeedWorkspace(workspaceId);
    return json(buildPublicWorkspacePayload(state, origin));
  }

  if (method === 'POST' && path === '/api/workspace/reset') {
    const body = (await req.json().catch(() => ({}))) as { workspaceId?: string };
    const workspaceId = body.workspaceId?.trim() || 'demo';
    if (workspaceId === 'demo') {
      const fresh = await (store as unknown as { buildDemoState?: () => Promise<WorkspaceState> })
        .buildDemoState?.();
      if (fresh) {
        await store.saveWorkspaceState(fresh);
        return json(buildPublicWorkspacePayload(fresh, origin));
      }
    }
    const freshStore = await store.getOrSeedWorkspace(workspaceId);
    return json(buildPublicWorkspacePayload(freshStore, origin));
  }

  if (method === 'POST' && path === '/api/workspaces') {
    const body = (await req.json().catch(() => ({}))) as { name?: string; slug?: string };
    const slug =
      (body.slug || body.name || `ws-${crypto.randomUUID().slice(0, 6)}`)
        .toLowerCase()
        .replace(/[^a-z0-9-]+/g, '-')
        .replace(/^-|-$/g, '') || `ws-${crypto.randomUUID().slice(0, 6)}`;
    const state = await store.getOrSeedWorkspace(slug);
    if (body.name?.trim()) {
      state.workspace.name = body.name.trim();
      await store.saveWorkspaceState(state);
    }
    return json(buildPublicWorkspacePayload(state, origin), 201);
  }

  if (method === 'POST' && path === '/api/accounts') {
    const rawBody = await req.json().catch(() => null);
    const parsed = createAccountSchema.safeParse(rawBody);
    if (!parsed.success) {
      return json({ error: 'Invalid account payload', details: parsed.error.flatten() }, 400);
    }

    const input = parsed.data;
    const state = await store.getOrSeedWorkspace(input.workspaceId);
    const nowIso = new Date().toISOString();
    const account: CalendarAccount = {
      id: makeId('acc'),
      workspaceId: state.workspace.id,
      label: input.label.trim(),
      email: input.email.trim(),
      provider: input.provider,
      calendarId: 'primary',
      calendarName: input.calendarName.trim(),
      color: input.color,
      mode: 'sandbox',
      status: 'connected',
      watchChannelId: makeId(input.provider === 'google' ? 'gchan' : 'msub'),
      watchExpiresAt: new Date(
        Date.now() + (input.provider === 'google' ? 6 : 2) * 24 * 3600_000
      ).toISOString(),
      createdAt: nowIso,
    };

    const existingAccounts = [...state.accounts];
    state.accounts.push(account);

    if (input.autoMeshWithExisting && existingAccounts.length > 0) {
      const primary = existingAccounts[0];
      const forward: SyncRule = {
        id: makeId('rule'),
        workspaceId: state.workspace.id,
        sourceAccountId: primary.id,
        targetAccountId: account.id,
        enabled: true,
        busyTitle: 'Busy',
        lockMirrored: true,
        skipFreeEvents: true,
        skipDeclinedEvents: true,
        skipAllDayEvents: true,
        createdAt: nowIso,
      };
      const reverse: SyncRule = {
        id: makeId('rule'),
        workspaceId: state.workspace.id,
        sourceAccountId: account.id,
        targetAccountId: primary.id,
        enabled: true,
        busyTitle: 'Busy',
        lockMirrored: true,
        skipFreeEvents: true,
        skipDeclinedEvents: true,
        skipAllDayEvents: true,
        createdAt: nowIso,
      };
      state.rules.push(forward, reverse);

      const engine = new StealthSyncEngine(store.createSyncPort(state));
      const primaryEvents = state.events.filter(
        (e) => e.accountId === primary.id && !isStealthManagedEvent(e)
      );
      for (const evt of primaryEvents) {
        await engine.processEventChange(state.workspace.id, evt);
      }
    }

    await store.saveWorkspaceState(state);
    return json({ account, workspace: buildPublicWorkspacePayload(state, origin) }, 201);
  }

  const accountDeleteMatch = path.match(/^\/api\/accounts\/([^/]+)$/);
  if (method === 'DELETE' && accountDeleteMatch) {
    const accountId = decodeURIComponent(accountDeleteMatch[1]);
    const workspaceId = url.searchParams.get('workspaceId')?.trim() || 'demo';
    const state = await store.getOrSeedWorkspace(workspaceId);
    state.accounts = state.accounts.filter((a) => a.id !== accountId);
    state.rules = state.rules.filter(
      (r) => r.sourceAccountId !== accountId && r.targetAccountId !== accountId
    );
    state.events = state.events.filter(
      (e) =>
        e.accountId !== accountId &&
        e.stealthMirrorMeta?.originAccountId !== accountId
    );
    state.mappings = state.mappings.filter(
      (m) => m.sourceAccountId !== accountId && m.targetAccountId !== accountId
    );
    await store.saveWorkspaceState(state);
    return json(buildPublicWorkspacePayload(state, origin));
  }

  if (method === 'POST' && path === '/api/rules') {
    const rawBody = await req.json().catch(() => null);
    const parsed = createRuleSchema.safeParse(rawBody);
    if (!parsed.success) {
      return json({ error: 'Invalid rule payload', details: parsed.error.flatten() }, 400);
    }
    const input = parsed.data;
    if (input.sourceAccountId === input.targetAccountId) {
      return json({ error: 'Source and target accounts must differ' }, 400);
    }
    const state = await store.getOrSeedWorkspace(input.workspaceId);
    const existing = state.rules.find(
      (r) =>
        r.sourceAccountId === input.sourceAccountId &&
        r.targetAccountId === input.targetAccountId
    );
    if (existing) {
      existing.enabled = true;
      existing.busyTitle = resolveSafeBusyTitle(input.busyTitle);
      existing.lockMirrored = input.lockMirrored;
      existing.skipFreeEvents = input.skipFreeEvents;
      existing.skipDeclinedEvents = input.skipDeclinedEvents;
      existing.skipAllDayEvents = input.skipAllDayEvents;
    } else {
      state.rules.push({
        id: makeId('rule'),
        workspaceId: state.workspace.id,
        sourceAccountId: input.sourceAccountId,
        targetAccountId: input.targetAccountId,
        enabled: true,
        busyTitle: resolveSafeBusyTitle(input.busyTitle),
        lockMirrored: input.lockMirrored,
        skipFreeEvents: input.skipFreeEvents,
        skipDeclinedEvents: input.skipDeclinedEvents,
        skipAllDayEvents: input.skipAllDayEvents,
        createdAt: new Date().toISOString(),
      });
    }

    const engine = new StealthSyncEngine(store.createSyncPort(state));
    const sourceEvents = state.events.filter(
      (e) => e.accountId === input.sourceAccountId && !isStealthManagedEvent(e)
    );
    for (const evt of sourceEvents) {
      await engine.processEventChange(state.workspace.id, evt);
    }

    await store.saveWorkspaceState(state);
    return json(buildPublicWorkspacePayload(state, origin), 201);
  }

  const rulePatchMatch = path.match(/^\/api\/rules\/([^/]+)$/);
  if (method === 'PATCH' && rulePatchMatch) {
    const ruleId = decodeURIComponent(rulePatchMatch[1]);
    const rawBody = await req.json().catch(() => null);
    const parsed = updateRuleSchema.safeParse(rawBody);
    if (!parsed.success) {
      return json({ error: 'Invalid rule update', details: parsed.error.flatten() }, 400);
    }
    const input = parsed.data;
    const state = await store.getOrSeedWorkspace(input.workspaceId);
    const rule = state.rules.find((r) => r.id === ruleId);
    if (!rule) return json({ error: 'Rule not found' }, 404);

    if (input.enabled !== undefined) rule.enabled = input.enabled;
    if (input.busyTitle !== undefined) {
      rule.busyTitle = resolveSafeBusyTitle(input.busyTitle);
    }
    if (input.lockMirrored !== undefined) rule.lockMirrored = input.lockMirrored;
    if (input.skipFreeEvents !== undefined) rule.skipFreeEvents = input.skipFreeEvents;
    if (input.skipDeclinedEvents !== undefined) {
      rule.skipDeclinedEvents = input.skipDeclinedEvents;
    }
    if (input.skipAllDayEvents !== undefined) {
      rule.skipAllDayEvents = input.skipAllDayEvents;
    }

    if (rule.enabled) {
      const engine = new StealthSyncEngine(store.createSyncPort(state));
      const sourceEvents = state.events.filter(
        (e) => e.accountId === rule.sourceAccountId && !isStealthManagedEvent(e)
      );
      for (const evt of sourceEvents) {
        await engine.processEventChange(state.workspace.id, evt);
      }
    }

    await store.saveWorkspaceState(state);
    return json(buildPublicWorkspacePayload(state, origin));
  }

  if (method === 'DELETE' && rulePatchMatch) {
    const ruleId = decodeURIComponent(rulePatchMatch[1]);
    const workspaceId = url.searchParams.get('workspaceId')?.trim() || 'demo';
    const state = await store.getOrSeedWorkspace(workspaceId);
    state.rules = state.rules.filter((r) => r.id !== ruleId);
    await store.saveWorkspaceState(state);
    return json(buildPublicWorkspacePayload(state, origin));
  }

  if (method === 'POST' && path === '/api/events') {
    const rawBody = await req.json().catch(() => null);
    const parsed = createEventSchema.safeParse(rawBody);
    if (!parsed.success) {
      return json({ error: 'Invalid event payload', details: parsed.error.flatten() }, 400);
    }
    const input = parsed.data;
    const state = await store.getOrSeedWorkspace(input.workspaceId);
    const account = state.accounts.find((a) => a.id === input.accountId);
    if (!account) return json({ error: 'Calendar account not found' }, 404);

    const extId = makeId(account.provider === 'google' ? 'g_evt' : 'o_evt');
    const event: CalendarEvent = {
      id: makeId('evt'),
      accountId: account.id,
      externalEventId: extId,
      title: input.title,
      description: input.description,
      location: input.location,
      attendees: input.attendees,
      startIso: new Date(input.startIso).toISOString(),
      endIso: new Date(input.endIso).toISOString(),
      allDay: input.allDay,
      showAs: input.showAs,
      visibility: 'default',
      responseStatus: input.responseStatus,
      locked: false,
      updatedAt: new Date().toISOString(),
    };
    state.events.push(event);

    const engine = new StealthSyncEngine(store.createSyncPort(state));
    const syncSummary = await engine.processEventChange(state.workspace.id, event);
    await store.saveWorkspaceState(state);

    return json(
      {
        event,
        syncSummary,
        workspace: buildPublicWorkspacePayload(state, origin),
      },
      201
    );
  }

  const eventMatch = path.match(/^\/api\/events\/([^/]+)$/);
  if (method === 'PATCH' && eventMatch) {
    const eventId = decodeURIComponent(eventMatch[1]);
    const rawBody = await req.json().catch(() => null);
    const parsed = updateEventSchema.safeParse(rawBody);
    if (!parsed.success) {
      return json({ error: 'Invalid event update', details: parsed.error.flatten() }, 400);
    }
    const input = parsed.data;
    const state = await store.getOrSeedWorkspace(input.workspaceId);
    const existing = state.events.find((e) => e.id === eventId);
    if (!existing) return json({ error: 'Event not found' }, 404);

    if (isStealthManagedEvent(existing) || existing.locked) {
      return json(
        {
          code: 'LOCKED_MIRROR_BLOCK',
          error:
            'This Busy block is locked by StealthMirror. Edit or reschedule the origin event on the source calendar instead.',
        },
        403
      );
    }

    if (input.title !== undefined) existing.title = input.title;
    if (input.description !== undefined) existing.description = input.description;
    if (input.location !== undefined) existing.location = input.location;
    if (input.startIso !== undefined) {
      existing.startIso = new Date(input.startIso).toISOString();
    }
    if (input.endIso !== undefined) {
      existing.endIso = new Date(input.endIso).toISOString();
    }
    if (input.allDay !== undefined) existing.allDay = input.allDay;
    if (input.showAs !== undefined) existing.showAs = input.showAs;
    if (input.responseStatus !== undefined) {
      existing.responseStatus = input.responseStatus;
    }
    existing.updatedAt = new Date().toISOString();

    const engine = new StealthSyncEngine(store.createSyncPort(state));
    const syncSummary = await engine.processEventChange(state.workspace.id, existing);
    await store.saveWorkspaceState(state);

    return json({
      event: existing,
      syncSummary,
      workspace: buildPublicWorkspacePayload(state, origin),
    });
  }

  if (method === 'DELETE' && eventMatch) {
    const eventId = decodeURIComponent(eventMatch[1]);
    const workspaceId = url.searchParams.get('workspaceId')?.trim() || 'demo';
    const state = await store.getOrSeedWorkspace(workspaceId);
    const existing = state.events.find((e) => e.id === eventId);
    if (!existing) return json({ error: 'Event not found' }, 404);

    if (isStealthManagedEvent(existing) || existing.locked) {
      return json(
        {
          code: 'LOCKED_MIRROR_BLOCK',
          error:
            'Cannot directly delete a locked Busy mirror block. Delete the origin meeting to remove all mirrors automatically.',
        },
        403
      );
    }

    existing.deleted = true;
    const engine = new StealthSyncEngine(store.createSyncPort(state));
    const syncSummary = await engine.processEventChange(state.workspace.id, existing);
    state.events = state.events.filter((e) => e.id !== eventId);
    await store.saveWorkspaceState(state);

    return json({
      syncSummary,
      workspace: buildPublicWorkspacePayload(state, origin),
    });
  }

  if (method === 'POST' && path === '/api/events/test-loop-guard') {
    const body = (await req.json().catch(() => ({}))) as {
      workspaceId?: string;
      eventId?: string;
    };
    const state = await store.getOrSeedWorkspace(body.workspaceId || 'demo');
    const mirrored =
      state.events.find((e) => e.id === body.eventId) ??
      state.events.find((e) => isStealthManagedEvent(e));
    if (!mirrored) return json({ error: 'No mirrored event found to test' }, 404);

    const engine = new StealthSyncEngine(store.createSyncPort(state));
    const syncSummary = await engine.processEventChange(state.workspace.id, mirrored);
    await store.saveWorkspaceState(state);
    return json({
      syncSummary,
      workspace: buildPublicWorkspacePayload(state, origin),
    });
  }

  if (method === 'POST' && path === '/api/legacy-cleanup') {
    const body = (await req.json().catch(() => ({}))) as { workspaceId?: string };
    const state = await store.getOrSeedWorkspace(body.workspaceId || 'demo');
    const toRemove = state.events.filter((e) => detectLegacySyncClutter(e) !== null);
    const removeIds = new Set(toRemove.map((e) => e.id));
    state.events = state.events.filter((e) => !removeIds.has(e.id));

    if (toRemove.length > 0) {
      state.logs.unshift({
        id: makeId('log'),
        workspaceId: state.workspace.id,
        sourceAccountId: toRemove[0].accountId,
        sourceEventId: toRemove.map((e) => e.externalEventId).join(','),
        action: 'deleted',
        reason: `Purged ${toRemove.length} legacy Reclaim/SyncGene clutter block(s)`,
        latencyMs: 3,
        createdAt: new Date().toISOString(),
      });
    }

    await store.saveWorkspaceState(state);
    return json({
      removedCount: toRemove.length,
      removedTitles: toRemove.map((e) => e.title),
      workspace: buildPublicWorkspacePayload(state, origin),
    });
  }

  if (method === 'POST' && path === '/api/billing/select') {
    const rawBody = await req.json().catch(() => null);
    const parsed = billingSelectSchema.safeParse(rawBody);
    if (!parsed.success) {
      return json({ error: 'Invalid billing plan' }, 400);
    }
    const state = await store.getOrSeedWorkspace(parsed.data.workspaceId);
    state.workspace.plan = parsed.data.plan as BillingPlan;
    await store.saveWorkspaceState(state);
    return json(buildPublicWorkspacePayload(state, origin));
  }

  // Real OAuth entrypoints for Google & Microsoft
  if (method === 'GET' && path === '/api/oauth/google/start') {
    const clientId = getEnv('GOOGLE_CLIENT_ID');
    if (!clientId) {
      return json(
        {
          error: 'Google OAuth not configured yet',
          requiredEnv: ['GOOGLE_CLIENT_ID', 'GOOGLE_CLIENT_SECRET', 'APP_URL'],
          callbackUrl: `${origin}/api/oauth/google/callback`,
        },
        400
      );
    }
    const workspaceId = url.searchParams.get('workspaceId') || 'demo';
    const label = url.searchParams.get('label') || 'Google Calendar';
    const nonce = crypto.randomUUID();
    await store.saveOAuthState(nonce, {
      workspaceId,
      provider: 'google',
      label,
      origin,
      createdAt: new Date().toISOString(),
    });
    const authUrl = buildGoogleAuthUrl({
      clientId,
      redirectUri: `${origin}/api/oauth/google/callback`,
      state: nonce,
    });
    return redirect(authUrl);
  }

  if (method === 'GET' && path === '/api/oauth/google/callback') {
    const code = url.searchParams.get('code');
    const nonce = url.searchParams.get('state');
    if (!code || !nonce) return json({ error: 'Missing OAuth code or state' }, 400);
    const oauthState = await store.consumeOAuthState(nonce);
    if (!oauthState) return json({ error: 'Invalid or expired OAuth state' }, 400);

    const clientId = getEnv('GOOGLE_CLIENT_ID');
    const clientSecret = getEnv('GOOGLE_CLIENT_SECRET');
    if (!clientId || !clientSecret) {
      return json({ error: 'Google OAuth credentials missing' }, 500);
    }

    const tokenRes = await exchangeGoogleAuthCode({
      code,
      clientId,
      clientSecret,
      redirectUri: `${oauthState.origin}/api/oauth/google/callback`,
    });
    const profile = await fetchGoogleUserProfile(tokenRes.accessToken);
    const state = await store.getOrSeedWorkspace(oauthState.workspaceId);
    const nowIso = new Date().toISOString();
    const channelId = makeId('gchan');

    const account: CalendarAccount = {
      id: makeId('acc'),
      workspaceId: state.workspace.id,
      label: oauthState.label || profile.name || 'Google Calendar',
      email: profile.email,
      provider: 'google',
      calendarId: 'primary',
      calendarName: profile.email,
      color: '#0EA5E9',
      mode: 'oauth',
      status: 'connected',
      accessToken: tokenRes.accessToken,
      refreshToken: tokenRes.refreshToken,
      tokenExpiresAt: new Date(Date.now() + tokenRes.expiresIn * 1000).toISOString(),
      watchChannelId: channelId,
      createdAt: nowIso,
    };

    try {
      const watchRes = await watchGoogleCalendar({
        accessToken: tokenRes.accessToken,
        calendarId: 'primary',
        channelId,
        webhookUrl: `${oauthState.origin}/api/webhooks/google`,
        token: `${state.workspace.id}:${account.id}`,
      });
      account.watchResourceId = watchRes.resourceId;
      account.watchExpiresAt = watchRes.expirationIso;
      await store.saveChannelLookup(channelId, {
        workspaceId: state.workspace.id,
        accountId: account.id,
      });

      const initialSync = await fetchGoogleIncrementalChanges({
        accessToken: tokenRes.accessToken,
        calendarId: 'primary',
      });
      account.syncToken = initialSync.nextSyncToken;
    } catch {
      // Keep connected account even if webhook registration fails in local dev
    }

    state.accounts.push(account);
    await store.saveWorkspaceState(state);
    return redirect(`${oauthState.origin}/?workspace=${encodeURIComponent(state.workspace.id)}&connected=google`);
  }

  if (method === 'GET' && path === '/api/oauth/microsoft/start') {
    const clientId = getEnv('MICROSOFT_CLIENT_ID');
    if (!clientId) {
      return json(
        {
          error: 'Microsoft OAuth not configured yet',
          requiredEnv: ['MICROSOFT_CLIENT_ID', 'MICROSOFT_CLIENT_SECRET', 'APP_URL'],
          callbackUrl: `${origin}/api/oauth/microsoft/callback`,
        },
        400
      );
    }
    const workspaceId = url.searchParams.get('workspaceId') || 'demo';
    const label = url.searchParams.get('label') || 'Outlook 365';
    const nonce = crypto.randomUUID();
    await store.saveOAuthState(nonce, {
      workspaceId,
      provider: 'outlook',
      label,
      origin,
      createdAt: new Date().toISOString(),
    });
    const authUrl = buildMicrosoftAuthUrl({
      clientId,
      redirectUri: `${origin}/api/oauth/microsoft/callback`,
      state: nonce,
    });
    return redirect(authUrl);
  }

  if (method === 'GET' && path === '/api/oauth/microsoft/callback') {
    const code = url.searchParams.get('code');
    const nonce = url.searchParams.get('state');
    if (!code || !nonce) return json({ error: 'Missing OAuth code or state' }, 400);
    const oauthState = await store.consumeOAuthState(nonce);
    if (!oauthState) return json({ error: 'Invalid or expired OAuth state' }, 400);

    const clientId = getEnv('MICROSOFT_CLIENT_ID');
    const clientSecret = getEnv('MICROSOFT_CLIENT_SECRET');
    if (!clientId || !clientSecret) {
      return json({ error: 'Microsoft OAuth credentials missing' }, 500);
    }

    const tokenRes = await exchangeMicrosoftAuthCode({
      code,
      clientId,
      clientSecret,
      redirectUri: `${oauthState.origin}/api/oauth/microsoft/callback`,
    });
    const profile = await fetchMicrosoftUserProfile(tokenRes.accessToken);
    const state = await store.getOrSeedWorkspace(oauthState.workspaceId);
    const email = profile.mail || profile.userPrincipalName;

    const account: CalendarAccount = {
      id: makeId('acc'),
      workspaceId: state.workspace.id,
      label: oauthState.label || profile.displayName || 'Outlook Calendar',
      email,
      provider: 'outlook',
      calendarId: 'primary',
      calendarName: 'Calendar',
      color: '#6366F1',
      mode: 'oauth',
      status: 'connected',
      accessToken: tokenRes.accessToken,
      refreshToken: tokenRes.refreshToken,
      tokenExpiresAt: new Date(Date.now() + tokenRes.expiresIn * 1000).toISOString(),
      createdAt: new Date().toISOString(),
    };

    try {
      const sub = await createMicrosoftGraphSubscription({
        accessToken: tokenRes.accessToken,
        notificationUrl: `${oauthState.origin}/api/webhooks/microsoft`,
        clientState: `${state.workspace.id}:${account.id}`,
      });
      account.watchChannelId = sub.subscriptionId;
      account.watchExpiresAt = sub.expirationDateTime;
      await store.saveChannelLookup(sub.subscriptionId, {
        workspaceId: state.workspace.id,
        accountId: account.id,
      });

      const delta = await fetchMicrosoftDeltaChanges({
        accessToken: tokenRes.accessToken,
      });
      account.deltaLink = delta.nextDeltaLink;
    } catch {
      // Keep connected account even if Graph subscription webhook fails in localhost
    }

    state.accounts.push(account);
    await store.saveWorkspaceState(state);
    return redirect(`${oauthState.origin}/?workspace=${encodeURIComponent(state.workspace.id)}&connected=outlook`);
  }

  // Real-time webhook receivers
  if (method === 'POST' && path === '/api/webhooks/google') {
    const headers = parseGoogleWebhookHeaders(req.headers);
    if (!headers.channelId) {
      return json({ ok: true, ignored: 'missing_channel_id' });
    }
    if (headers.resourceState === 'sync') {
      return json({ ok: true, handshake: 'sync_confirmed' });
    }
    const lookup = await store.getChannelLookup(headers.channelId);
    if (!lookup) {
      return json({ ok: true, ignored: 'unknown_channel' });
    }
    const state = await store.getWorkspaceState(lookup.workspaceId);
    const account = state?.accounts.find((a) => a.id === lookup.accountId);
    if (!state || !account || !account.accessToken) {
      return json({ ok: true, ignored: 'account_unavailable' });
    }

    let delta = await fetchGoogleIncrementalChanges({
      accessToken: account.accessToken,
      calendarId: account.calendarId,
      syncToken: account.syncToken,
    });
    if (delta.syncResetRequired) {
      delta = await fetchGoogleIncrementalChanges({
        accessToken: account.accessToken,
        calendarId: account.calendarId,
      });
    }
    if (delta.nextSyncToken) {
      account.syncToken = delta.nextSyncToken;
    }

    const engine = new StealthSyncEngine(store.createSyncPort(state));
    let processed = 0;
    for (const rawItem of delta.items) {
      const canonical = googleEventToCanonical(account.id, rawItem);
      await engine.processEventChange(state.workspace.id, canonical);
      processed += 1;
    }
    await store.saveWorkspaceState(state);
    return json({ ok: true, processed });
  }

  if (path === '/api/webhooks/microsoft') {
    const handshake = handleMicrosoftValidationHandshake(req);
    if (handshake) return handshake;

    if (method === 'POST') {
      const body = (await req.json().catch(() => ({}))) as {
        value?: Array<{
          subscriptionId?: string;
          clientState?: string;
          changeType?: string;
          resourceData?: { id?: string };
        }>;
      };
      const notifications = parseMicrosoftWebhookNotifications(body);
      let processed = 0;

      for (const note of notifications) {
        const lookup =
          (note.subscriptionId
            ? await store.getChannelLookup(note.subscriptionId)
            : null) ??
          (() => {
            const [workspaceId, accountId] = note.clientState.split(':');
            return workspaceId && accountId ? { workspaceId, accountId } : null;
          })();
        if (!lookup) continue;

        const state = await store.getWorkspaceState(lookup.workspaceId);
        const account = state?.accounts.find((a) => a.id === lookup.accountId);
        if (!state || !account || !account.accessToken) continue;

        const delta = await fetchMicrosoftDeltaChanges({
          accessToken: account.accessToken,
          deltaLink: account.deltaLink,
        });
        if (delta.nextDeltaLink) {
          account.deltaLink = delta.nextDeltaLink;
        }
        const engine = new StealthSyncEngine(store.createSyncPort(state));
        for (const rawItem of delta.items) {
          const canonical = outlookEventToCanonical(account.id, rawItem);
          await engine.processEventChange(state.workspace.id, canonical);
          processed += 1;
        }
        await store.saveWorkspaceState(state);
      }

      return json({ ok: true, processed }, 202);
    }
  }

  if (method === 'POST' && path === '/api/cron/renew-watches') {
    const result = await renewExpiringWatches(store);
    return json(result);
  }

  return json({ error: `Route not found: ${method} ${path}` }, 404);
}
