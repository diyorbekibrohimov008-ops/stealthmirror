import type {
  BillingPlan,
  CalendarAccount,
  CalendarEvent,
  CalendarProvider,
  EventResponseStatus,
  EventShowAs,
  MirrorMapping,
  SyncLogEntry,
  SyncRule,
  Workspace,
} from '../../shared/types';

export interface WorkspaceStats {
  accountCount: number;
  activeRuleCount: number;
  mirroredBlockCount: number;
  legacyClutterCount: number;
  avgLatencyMs: number;
  loopsBlocked: number;
  leaksCount: number;
  estimatedSyncGeneMonthlyUsd: number;
}

export interface OAuthStatus {
  googleConfigured: boolean;
  microsoftConfigured: boolean;
  appUrl: string;
  googleCallbackUrl: string;
  microsoftCallbackUrl: string;
  googleWebhookUrl: string;
  microsoftWebhookUrl: string;
}

export interface PublicWorkspacePayload {
  workspace: Workspace;
  accounts: CalendarAccount[];
  rules: SyncRule[];
  events: CalendarEvent[];
  mappings: MirrorMapping[];
  logs: SyncLogEntry[];
  stats: WorkspaceStats;
  oauth: OAuthStatus;
}

export interface SyncChangeResponse {
  event?: CalendarEvent;
  syncSummary: {
    created: number;
    updated: number;
    deleted: number;
    skippedLoop: number;
    skippedPolicy: number;
  };
  workspace: PublicWorkspacePayload;
}

export interface LegacyCleanupResponse {
  removedCount: number;
  removedTitles: string[];
  workspace: PublicWorkspacePayload;
}

export interface RenewWatchesResponse {
  renewed: number;
  errors: number;
  checkedWorkspaces: number;
}

export class ApiClientError extends Error {
  readonly status: number;
  readonly code?: string;
  readonly details?: Record<string, unknown>;

  constructor(
    message: string,
    status: number,
    code?: string,
    details?: Record<string, unknown>
  ) {
    super(message);
    this.name = 'ApiClientError';
    this.status = status;
    this.code = code;
    this.details = details;
  }
}

async function requestJson<T>(url: string, init?: RequestInit): Promise<T> {
  const res = await fetch(url, {
    headers: {
      'content-type': 'application/json',
      ...(init?.headers ?? {}),
    },
    ...init,
  });

  const payload = (await res.json().catch(() => ({}))) as Record<string, unknown>;
  if (!res.ok) {
    const errorMsg =
      typeof payload.error === 'string'
        ? payload.error
        : `Request failed with status ${res.status}`;
    const code = typeof payload.code === 'string' ? payload.code : undefined;
    throw new ApiClientError(errorMsg, res.status, code, payload);
  }
  return payload as T;
}

export const api = {
  getWorkspace(workspaceId = 'demo'): Promise<PublicWorkspacePayload> {
    return requestJson<PublicWorkspacePayload>(
      `/api/workspace?workspaceId=${encodeURIComponent(workspaceId)}`
    );
  },

  resetWorkspace(workspaceId = 'demo'): Promise<PublicWorkspacePayload> {
    return requestJson<PublicWorkspacePayload>('/api/workspace/reset', {
      method: 'POST',
      body: JSON.stringify({ workspaceId }),
    });
  },

  createWorkspace(input: { name: string; slug?: string }): Promise<PublicWorkspacePayload> {
    return requestJson<PublicWorkspacePayload>('/api/workspaces', {
      method: 'POST',
      body: JSON.stringify(input),
    });
  },

  createAccount(input: {
    workspaceId: string;
    provider: CalendarProvider;
    label: string;
    email: string;
    calendarName?: string;
    color?: string;
    autoMeshWithExisting?: boolean;
  }): Promise<{ account: CalendarAccount; workspace: PublicWorkspacePayload }> {
    return requestJson<{ account: CalendarAccount; workspace: PublicWorkspacePayload }>(
      '/api/accounts',
      {
        method: 'POST',
        body: JSON.stringify(input),
      }
    );
  },

  deleteAccount(accountId: string, workspaceId: string): Promise<PublicWorkspacePayload> {
    return requestJson<PublicWorkspacePayload>(
      `/api/accounts/${encodeURIComponent(accountId)}?workspaceId=${encodeURIComponent(workspaceId)}`,
      { method: 'DELETE' }
    );
  },

  createRule(input: {
    workspaceId: string;
    sourceAccountId: string;
    targetAccountId: string;
    busyTitle?: string;
    lockMirrored?: boolean;
    skipFreeEvents?: boolean;
    skipDeclinedEvents?: boolean;
    skipAllDayEvents?: boolean;
  }): Promise<PublicWorkspacePayload> {
    return requestJson<PublicWorkspacePayload>('/api/rules', {
      method: 'POST',
      body: JSON.stringify(input),
    });
  },

  updateRule(
    ruleId: string,
    input: {
      workspaceId: string;
      enabled?: boolean;
      busyTitle?: string;
      lockMirrored?: boolean;
      skipFreeEvents?: boolean;
      skipDeclinedEvents?: boolean;
      skipAllDayEvents?: boolean;
    }
  ): Promise<PublicWorkspacePayload> {
    return requestJson<PublicWorkspacePayload>(`/api/rules/${encodeURIComponent(ruleId)}`, {
      method: 'PATCH',
      body: JSON.stringify(input),
    });
  },

  deleteRule(ruleId: string, workspaceId: string): Promise<PublicWorkspacePayload> {
    return requestJson<PublicWorkspacePayload>(
      `/api/rules/${encodeURIComponent(ruleId)}?workspaceId=${encodeURIComponent(workspaceId)}`,
      { method: 'DELETE' }
    );
  },

  createEvent(input: {
    workspaceId: string;
    accountId: string;
    title: string;
    description?: string;
    location?: string;
    attendees?: string[];
    startIso: string;
    endIso: string;
    allDay?: boolean;
    showAs?: EventShowAs;
    responseStatus?: EventResponseStatus;
  }): Promise<SyncChangeResponse> {
    return requestJson<SyncChangeResponse>('/api/events', {
      method: 'POST',
      body: JSON.stringify(input),
    });
  },

  updateEvent(
    eventId: string,
    input: {
      workspaceId: string;
      title?: string;
      description?: string;
      location?: string;
      startIso?: string;
      endIso?: string;
      allDay?: boolean;
      showAs?: EventShowAs;
      responseStatus?: EventResponseStatus;
    }
  ): Promise<SyncChangeResponse> {
    return requestJson<SyncChangeResponse>(`/api/events/${encodeURIComponent(eventId)}`, {
      method: 'PATCH',
      body: JSON.stringify(input),
    });
  },

  deleteEvent(eventId: string, workspaceId: string): Promise<SyncChangeResponse> {
    return requestJson<SyncChangeResponse>(
      `/api/events/${encodeURIComponent(eventId)}?workspaceId=${encodeURIComponent(workspaceId)}`,
      { method: 'DELETE' }
    );
  },

  testLoopGuard(workspaceId: string, eventId?: string): Promise<SyncChangeResponse> {
    return requestJson<SyncChangeResponse>('/api/events/test-loop-guard', {
      method: 'POST',
      body: JSON.stringify({ workspaceId, eventId }),
    });
  },

  cleanupLegacyClutter(workspaceId: string): Promise<LegacyCleanupResponse> {
    return requestJson<LegacyCleanupResponse>('/api/legacy-cleanup', {
      method: 'POST',
      body: JSON.stringify({ workspaceId }),
    });
  },

  selectBillingPlan(workspaceId: string, plan: BillingPlan): Promise<PublicWorkspacePayload> {
    return requestJson<PublicWorkspacePayload>('/api/billing/select', {
      method: 'POST',
      body: JSON.stringify({ workspaceId, plan }),
    });
  },

  renewWatches(): Promise<RenewWatchesResponse> {
    return requestJson<RenewWatchesResponse>('/api/cron/renew-watches', {
      method: 'POST',
    });
  },
};
