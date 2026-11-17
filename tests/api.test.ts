import { describe, expect, it } from 'vitest';
import { createMemoryStore } from '../netlify/lib/store';
import { handleApiRequest } from '../netlify/lib/router';

function jsonReq(path: string, method = 'GET', body?: unknown) {
  return new Request(`http://localhost${path}`, {
    method,
    headers: body ? { 'content-type': 'application/json' } : undefined,
    body: body ? JSON.stringify(body) : undefined,
  });
}

describe('StealthMirror API router', () => {
  it('seeds a realistic demo workspace and supports unlimited accounts with $0 add-on fees', async () => {
    const store = createMemoryStore();
    const res = await handleApiRequest(jsonReq('/api/workspace?workspaceId=demo'), store);
    expect(res.status).toBe(200);
    const data = (await res.json()) as {
      workspace: { id: string; plan: string };
      accounts: Array<{ id: string }>;
      rules: Array<{ id: string }>;
      events: Array<{ id: string }>;
    };

    expect(data.workspace.id).toBe('demo');
    expect(data.accounts.length).toBeGreaterThanOrEqual(3);
    expect(data.rules.length).toBeGreaterThanOrEqual(2);
    expect(data.events.length).toBeGreaterThan(0);

    // Add 4th calendar without any per-pair surcharge
    const addRes = await handleApiRequest(
      jsonReq('/api/accounts', 'POST', {
        workspaceId: 'demo',
        provider: 'outlook',
        label: 'Board Advisory',
        email: 'alex@board.org',
        calendarName: 'Advisory',
        color: '#f59e0b',
      }),
      store
    );
    expect(addRes.status).toBe(201);
    const created = (await addRes.json()) as { account: { id: string } };
    expect(created.account.id).toBeTruthy();
  });

  it('creates an event and immediately mirrors an unmovable private Busy block with zero title leak', async () => {
    const store = createMemoryStore();
    const initialRes = await handleApiRequest(jsonReq('/api/workspace?workspaceId=demo'), store);
    const initial = (await initialRes.json()) as {
      accounts: Array<{ id: string; provider: string }>;
    };
    const workAcc = initial.accounts[0];

    const createEvtRes = await handleApiRequest(
      jsonReq('/api/events', 'POST', {
        workspaceId: 'demo',
        accountId: workAcc.id,
        title: 'Secret Series B Term Sheet Review',
        description: 'CONFIDENTIAL: https://meet.google.com/secret-room',
        location: 'Executive Suite 9',
        attendees: ['vc@sequoia.com'],
        startIso: '2026-11-12T18:00:00.000Z',
        endIso: '2026-11-12T19:00:00.000Z',
        showAs: 'busy',
        allDay: false,
      }),
      store
    );
    expect(createEvtRes.status).toBe(201);
    const created = (await createEvtRes.json()) as {
      syncSummary: { created: number; skippedLoop: number };
      event: { id: string; externalEventId: string };
    };
    expect(created.syncSummary.created).toBeGreaterThanOrEqual(1);

    // Verify mirrored blocks exist on target accounts without leaking any title/description/location
    const afterRes = await handleApiRequest(jsonReq('/api/workspace?workspaceId=demo'), store);
    const after = (await afterRes.json()) as {
      events: Array<{
        id: string;
        accountId: string;
        title: string;
        description: string;
        location: string;
        attendees: string[];
        locked: boolean;
        visibility: string;
        stealthMirrorMeta?: { managed: boolean; originEventId: string };
      }>;
      logs: Array<{ action: string; latencyMs: number; payloadPreview?: Record<string, unknown> }>;
    };

    const mirroredBlocks = after.events.filter(
      (e) => e.stealthMirrorMeta?.managed && e.stealthMirrorMeta.originEventId === created.event.externalEventId
    );
    expect(mirroredBlocks.length).toBeGreaterThanOrEqual(1);
    for (const block of mirroredBlocks) {
      expect(block.title).toBe('Busy');
      expect(block.description).toBe('');
      expect(block.location).toBe('');
      expect(block.attendees).toEqual([]);
      expect(block.locked).toBe(true);
      expect(block.visibility).toBe('private');
    }

    // Attempt to directly modify a locked mirrored block -> rejected with 403
    const lockTestRes = await handleApiRequest(
      jsonReq(`/api/events/${mirroredBlocks[0].id}`, 'PATCH', {
        workspaceId: 'demo',
        startIso: '2026-11-12T20:00:00.000Z',
        endIso: '2026-11-12T21:00:00.000Z',
      }),
      store
    );
    expect(lockTestRes.status).toBe(403);
  });

  it('cleans up legacy Reclaim/SyncGene clutter in 1 click without deleting real meetings', async () => {
    const store = createMemoryStore();
    await handleApiRequest(jsonReq('/api/workspace?workspaceId=demo'), store);

    const cleanupRes = await handleApiRequest(
      jsonReq('/api/legacy-cleanup', 'POST', { workspaceId: 'demo' }),
      store
    );
    expect(cleanupRes.status).toBe(200);
    const body = (await cleanupRes.json()) as { removedCount: number };
    expect(body.removedCount).toBeGreaterThanOrEqual(1);
  });

  it('switches billing plans between $5/mo flat and $39 lifetime', async () => {
    const store = createMemoryStore();
    await handleApiRequest(jsonReq('/api/workspace?workspaceId=demo'), store);

    const res = await handleApiRequest(
      jsonReq('/api/billing/select', 'POST', {
        workspaceId: 'demo',
        plan: 'lifetime_39',
      }),
      store
    );
    expect(res.status).toBe(200);
    const data = (await res.json()) as { workspace: { plan: string } };
    expect(data.workspace.plan).toBe('lifetime_39');
  });

  it('removes mirrored blocks when deleting an origin event and blocks loop webhooks', async () => {
    const store = createMemoryStore();
    const initialRes = await handleApiRequest(jsonReq('/api/workspace?workspaceId=demo'), store);
    const initial = (await initialRes.json()) as {
      accounts: Array<{ id: string }>;
    };

    const createRes = await handleApiRequest(
      jsonReq('/api/events', 'POST', {
        workspaceId: 'demo',
        accountId: initial.accounts[0].id,
        title: 'Confidential Merger Call',
        description: 'Zoom 123',
        location: 'Room 1',
        attendees: ['a@b.com'],
        startIso: '2026-11-12T21:00:00.000Z',
        endIso: '2026-11-12T22:00:00.000Z',
        showAs: 'busy',
        allDay: false,
      }),
      store
    );
    const created = (await createRes.json()) as {
      event: { id: string; externalEventId: string };
    };

    const loopRes = await handleApiRequest(
      jsonReq('/api/events/test-loop-guard', 'POST', { workspaceId: 'demo' }),
      store
    );
    expect(loopRes.status).toBe(200);
    const loopBody = (await loopRes.json()) as {
      syncSummary: { skippedLoop: number; created: number };
    };
    expect(loopBody.syncSummary.skippedLoop).toBe(1);
    expect(loopBody.syncSummary.created).toBe(0);

    const delRes = await handleApiRequest(
      jsonReq(`/api/events/${created.event.id}?workspaceId=demo`, 'DELETE'),
      store
    );
    expect(delRes.status).toBe(200);
    const delBody = (await delRes.json()) as {
      syncSummary: { deleted: number };
      workspace: { events: Array<{ stealthMirrorMeta?: { originEventId: string } }> };
    };
    expect(delBody.syncSummary.deleted).toBeGreaterThanOrEqual(1);
    expect(
      delBody.workspace.events.some(
        (e) => e.stealthMirrorMeta?.originEventId === created.event.externalEventId
      )
    ).toBe(false);
  });

  it('handles Microsoft Graph webhook validation handshake directly on /api/webhooks/microsoft', async () => {
    const store = createMemoryStore();
    const res = await handleApiRequest(
      new Request('http://localhost/api/webhooks/microsoft?validationToken=abc-xyz-123', {
        method: 'POST',
      }),
      store
    );
    expect(res.status).toBe(200);
    expect(await res.text()).toBe('abc-xyz-123');
  });

  it('supports Vercel rewrite __path query routing cleanly', async () => {
    const store = createMemoryStore();
    const res = await handleApiRequest(
      jsonReq('/api/index?__path=/api/workspace&workspaceId=demo'),
      store
    );
    expect(res.status).toBe(200);
    const body = (await res.json()) as { workspace: { id: string } };
    expect(body.workspace.id).toBe('demo');
  });
});
