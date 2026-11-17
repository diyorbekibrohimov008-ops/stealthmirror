import { useCallback, useEffect, useState } from 'react';
import { CheckCircle2, Lock, ShieldAlert, X } from 'lucide-react';
import { MigrationAndPricing } from './components/MigrationAndPricing';
import { Navbar, type ActiveTab } from './components/Navbar';
import { SplitCalendarCanvas } from './components/SplitCalendarCanvas';
import { SyncMatrixPanel } from './components/SyncMatrixPanel';
import { api, ApiClientError, type PublicWorkspacePayload } from './lib/api';

interface ToastBanner {
  id: number;
  kind: 'success' | 'lock' | 'error';
  title: string;
  detail?: string;
}

function getInitialWorkspaceId(): string {
  if (typeof window === 'undefined') return 'demo';
  const params = new URLSearchParams(window.location.search);
  return params.get('workspace')?.trim() || 'demo';
}

export function App() {
  const [workspaceId, setWorkspaceId] = useState<string>(getInitialWorkspaceId);
  const [workspaceIds, setWorkspaceIds] = useState<string[]>(() => {
    const initial = getInitialWorkspaceId();
    return initial === 'demo' ? ['demo'] : ['demo', initial];
  });
  const [data, setData] = useState<PublicWorkspacePayload | null>(null);
  const [loading, setLoading] = useState(true);
  const [busyAction, setBusyAction] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<ActiveTab>('workspace');
  const [toast, setToast] = useState<ToastBanner | null>(null);

  const pushToast = useCallback(
    (kind: ToastBanner['kind'], title: string, detail?: string) => {
      setToast({ id: Date.now(), kind, title, detail });
    },
    []
  );

  useEffect(() => {
    if (!toast) return;
    const timer = setTimeout(() => setToast(null), 5000);
    return () => clearTimeout(timer);
  }, [toast]);

  const loadWorkspace = useCallback(
    async (targetId: string) => {
      setLoading(true);
      try {
        const res = await api.getWorkspace(targetId);
        setData(res);
        const params = new URLSearchParams(window.location.search);
        const connected = params.get('connected');
        if (connected) {
          pushToast(
            'success',
            `${connected === 'google' ? 'Google Calendar' : 'Outlook 365'} connected via OAuth`,
            'Real-time webhook watch & delta token active.'
          );
        }
      } catch (err) {
        const msg = err instanceof Error ? err.message : 'Unable to reach StealthMirror API';
        pushToast('error', 'Failed to load workspace', msg);
      } finally {
        setLoading(false);
      }
    },
    [pushToast]
  );

  useEffect(() => {
    void loadWorkspace(workspaceId);
  }, [workspaceId, loadWorkspace]);

  async function runAction<T>(
    actionKey: string,
    fn: () => Promise<T>,
    onSuccess?: (result: T) => void
  ) {
    setBusyAction(actionKey);
    try {
      const result = await fn();
      onSuccess?.(result);
    } catch (err) {
      if (err instanceof ApiClientError && err.code === 'LOCKED_MIRROR_BLOCK') {
        pushToast(
          'lock',
          'Edit Blocked by StealthMirror Lock',
          err.message
        );
      } else {
        const msg = err instanceof Error ? err.message : 'Operation failed';
        pushToast('error', 'Action Failed', msg);
      }
    } finally {
      setBusyAction(null);
    }
  }

  async function handleCreateWorkspace() {
    const randomSuffix = Math.random().toString(36).slice(2, 6);
    const slug = `team-${randomSuffix}`;
    await runAction(
      'create-workspace',
      () => api.createWorkspace({ name: `Team ${randomSuffix.toUpperCase()}`, slug }),
      (payload) => {
        setWorkspaceIds((prev) =>
          prev.includes(payload.workspace.id) ? prev : [...prev, payload.workspace.id]
        );
        setWorkspaceId(payload.workspace.id);
        setData(payload);
        setActiveTab('matrix');
        pushToast(
          'success',
          `Created blank workspace "${payload.workspace.name}"`,
          'Connect your own accounts or sandbox calendars.'
        );
      }
    );
  }

  return (
    <div className="min-h-screen bg-[#F8FAFC] text-slate-900">
      {data && (
        <Navbar
          workspace={data.workspace}
          stats={data.stats}
          activeTab={activeTab}
          onSelectTab={setActiveTab}
          workspaceIds={workspaceIds}
          onSwitchWorkspace={(id) => setWorkspaceId(id)}
          onCreateWorkspace={() => void handleCreateWorkspace()}
          onResetDemo={() =>
            void runAction(
              'reset-demo',
              () => api.resetWorkspace('demo'),
              (payload) => {
                setData(payload);
                pushToast(
                  'success',
                  'Demo workspace reset',
                  'Restored 3 accounts, 4 bidirectional rules, and sample legacy clutter.'
                );
              }
            )
          }
          busyAction={busyAction}
        />
      )}

      {toast && (
        <div className="fixed right-4 bottom-4 z-50 max-w-md">
          <div
            className={`flex items-start gap-2.5 rounded-xl border px-4 py-3 text-xs shadow-lg ${
              toast.kind === 'lock'
                ? 'border-slate-700 bg-slate-900 text-white'
                : toast.kind === 'error'
                ? 'border-rose-300 bg-rose-50 text-rose-950'
                : 'border-emerald-200 bg-white text-slate-900'
            }`}
          >
            {toast.kind === 'lock' ? (
              <Lock className="mt-0.5 h-4 w-4 shrink-0 text-amber-400" />
            ) : toast.kind === 'error' ? (
              <ShieldAlert className="mt-0.5 h-4 w-4 shrink-0 text-rose-600" />
            ) : (
              <CheckCircle2 className="mt-0.5 h-4 w-4 shrink-0 text-emerald-600" />
            )}
            <div className="flex-1">
              <p className="font-semibold">{toast.title}</p>
              {toast.detail && (
                <p
                  className={`mt-0.5 leading-relaxed ${
                    toast.kind === 'lock' ? 'text-slate-300' : 'text-slate-600'
                  }`}
                >
                  {toast.detail}
                </p>
              )}
            </div>
            <button
              type="button"
              onClick={() => setToast(null)}
              aria-label="Dismiss notification"
              className="rounded p-0.5 opacity-70 hover:opacity-100"
            >
              <X className="h-3.5 w-3.5" />
            </button>
          </div>
        </div>
      )}

      <main className="mx-auto max-w-7xl px-4 py-6 sm:px-6">
        {loading && !data ? (
          <div className="grid grid-cols-1 gap-4 lg:grid-cols-3">
            {[1, 2, 3].map((idx) => (
              <div
                key={idx}
                className="h-72 animate-pulse rounded-xl border border-slate-200 bg-white"
              />
            ))}
          </div>
        ) : !data ? (
          <div className="rounded-xl border border-rose-200 bg-white p-8 text-center">
            <p className="text-sm font-semibold text-slate-900">
              Could not load workspace data.
            </p>
            <button
              type="button"
              onClick={() => void loadWorkspace(workspaceId)}
              className="mt-3 rounded-lg bg-slate-900 px-4 py-2 text-xs font-semibold text-white"
            >
              Retry
            </button>
          </div>
        ) : activeTab === 'workspace' ? (
          <SplitCalendarCanvas
            accounts={data.accounts}
            rules={data.rules}
            events={data.events}
            logs={data.logs}
            stats={data.stats}
            oauth={data.oauth}
            busyAction={busyAction}
            onCreateEvent={async (input) => {
              await runAction(
                'create-event',
                () => api.createEvent({ workspaceId: data.workspace.id, ...input }),
                (res) => {
                  setData(res.workspace);
                  const created = res.syncSummary.created;
                  const skippedPolicy = res.syncSummary.skippedPolicy;
                  pushToast(
                    'success',
                    `"${input.title}" saved`,
                    created > 0
                      ? `Mirrored ${created} private "Busy" block(s) with 0 title leaks.`
                      : skippedPolicy > 0
                      ? `Skipped cleanly per Free/Declined/All-Day rule (${skippedPolicy}).`
                      : 'Saved without outbound mirrors.'
                  );
                }
              );
            }}
            onUpdateEvent={async (eventId, patch) => {
              await runAction(
                `patch-${eventId}`,
                () =>
                  api.updateEvent(eventId, {
                    workspaceId: data.workspace.id,
                    ...patch,
                  }),
                (res) => {
                  setData(res.workspace);
                  pushToast(
                    'success',
                    'Source event updated & mirrors resynced',
                    `${res.syncSummary.updated} updated · ${res.syncSummary.deleted} deleted.`
                  );
                }
              );
            }}
            onDeleteEvent={async (eventId) => {
              await runAction(
                `delete-${eventId}`,
                () => api.deleteEvent(eventId, data.workspace.id),
                (res) => {
                  setData(res.workspace);
                  pushToast(
                    'success',
                    'Origin event & mirrored Busy blocks deleted',
                    `${res.syncSummary.deleted} mirror block(s) removed across targets.`
                  );
                }
              );
            }}
            onTestLoopGuard={async (eventId) => {
              await runAction(
                'loop-guard',
                () => api.testLoopGuard(data.workspace.id, eventId),
                (res) => {
                  setData(res.workspace);
                  pushToast(
                    'success',
                    'Webhook Loop Guard Verified',
                    `Blocked ${res.syncSummary.skippedLoop} echo webhook on mirrored Busy block.`
                  );
                }
              );
            }}
            onCleanupLegacy={async () => {
              await runAction(
                'legacy-cleanup',
                () => api.cleanupLegacyClutter(data.workspace.id),
                (res) => {
                  setData(res.workspace);
                  pushToast(
                    'success',
                    `Removed ${res.removedCount} legacy clutter block(s)`,
                    'Reclaim.ai & SyncGene injected entries purged.'
                  );
                }
              );
            }}
            onRenewWatches={async () => {
              await runAction(
                'renew',
                () => api.renewWatches(),
                (res) => {
                  pushToast(
                    'success',
                    'Push Channel Renewal Check Complete',
                    `Checked ${res.checkedWorkspaces} workspace(s) · renewed ${res.renewed}.`
                  );
                }
              );
            }}
            onOpenConnectTab={() => setActiveTab('matrix')}
          />
        ) : activeTab === 'matrix' ? (
          <SyncMatrixPanel
            workspaceId={data.workspace.id}
            accounts={data.accounts}
            rules={data.rules}
            oauth={data.oauth}
            busyAction={busyAction}
            onCreateAccount={async (input) => {
              await runAction(
                'create-account',
                () => api.createAccount({ workspaceId: data.workspace.id, ...input }),
                (res) => {
                  setData(res.workspace);
                  pushToast(
                    'success',
                    `Connected "${res.account.label}"`,
                    'Included free in unlimited flat plan ($0 extra).'
                  );
                }
              );
            }}
            onDeleteAccount={async (accountId) => {
              await runAction(
                `delete-acc-${accountId}`,
                () => api.deleteAccount(accountId, data.workspace.id),
                (payload) => {
                  setData(payload);
                  pushToast('success', 'Account disconnected & orphaned mirrors cleaned');
                }
              );
            }}
            onCreateRule={async (input) => {
              await runAction(
                'create-rule',
                () => api.createRule({ workspaceId: data.workspace.id, ...input }),
                (payload) => {
                  setData(payload);
                  pushToast('success', 'Sync route added & initial mirror pass complete');
                }
              );
            }}
            onUpdateRule={async (ruleId, patch) => {
              await runAction(
                `patch-rule-${ruleId}`,
                () => api.updateRule(ruleId, { workspaceId: data.workspace.id, ...patch }),
                (payload) => {
                  setData(payload);
                  pushToast('success', 'Sync rule updated');
                }
              );
            }}
            onDeleteRule={async (ruleId) => {
              await runAction(
                `delete-rule-${ruleId}`,
                () => api.deleteRule(ruleId, data.workspace.id),
                (payload) => {
                  setData(payload);
                  pushToast('success', 'Sync rule removed');
                }
              );
            }}
          />
        ) : (
          <MigrationAndPricing
            events={data.events}
            currentPlan={data.workspace.plan}
            busyAction={busyAction}
            onCleanupLegacy={async () => {
              await runAction(
                'legacy-cleanup',
                () => api.cleanupLegacyClutter(data.workspace.id),
                (res) => {
                  setData(res.workspace);
                  pushToast(
                    'success',
                    `Purged ${res.removedCount} legacy Reclaim/SyncGene block(s)`
                  );
                }
              );
            }}
            onSelectPlan={async (plan) => {
              await runAction(
                'select-plan',
                () => api.selectBillingPlan(data.workspace.id, plan),
                (payload) => {
                  setData(payload);
                  pushToast(
                    'success',
                    plan === 'lifetime_39'
                      ? 'Activated $39 Lifetime Unlimited Plan'
                      : 'Activated $5/mo Flat Unlimited Plan'
                  );
                }
              );
            }}
          />
        )}
      </main>
    </div>
  );
}

export default App;
