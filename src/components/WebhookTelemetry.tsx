import { useState } from 'react';
import {
  Activity,
  ChevronDown,
  ChevronUp,
  RefreshCw,
  ShieldAlert,
  ShieldCheck,
  Zap,
} from 'lucide-react';
import type { CalendarAccount, SyncLogEntry } from '../../shared/types';

interface WebhookTelemetryProps {
  accounts: CalendarAccount[];
  logs: SyncLogEntry[];
  onRenewWatches: () => void;
  busyAction: string | null;
}

const ACTION_BADGE_MAP: Record<SyncLogEntry['action'], { label: string; classes: string }> = {
  created: { label: 'CREATED', classes: 'bg-emerald-50 text-emerald-700 ring-emerald-200' },
  updated: { label: 'UPDATED', classes: 'bg-sky-50 text-sky-700 ring-sky-200' },
  deleted: { label: 'DELETED', classes: 'bg-slate-100 text-slate-700 ring-slate-200' },
  skipped_loop: {
    label: 'LOOP BLOCKED',
    classes: 'bg-violet-50 text-violet-700 ring-violet-200',
  },
  skipped_policy: {
    label: 'POLICY SKIP',
    classes: 'bg-amber-50 text-amber-700 ring-amber-200',
  },
  error: { label: 'ERROR', classes: 'bg-rose-50 text-rose-700 ring-rose-200' },
};

export function WebhookTelemetry({
  accounts,
  logs,
  onRenewWatches,
  busyAction,
}: WebhookTelemetryProps) {
  const [expandedLogId, setExpandedLogId] = useState<string | null>(null);
  const accountMap = new Map(accounts.map((a) => [a.id, a]));

  return (
    <section className="flex flex-col rounded-xl border border-slate-200 bg-white shadow-xs">
      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 px-4 py-3">
        <div className="flex items-center gap-2">
          <div className="flex h-7 w-7 items-center justify-center rounded-md bg-emerald-50 text-emerald-700">
            <Activity className="h-4 w-4" />
          </div>
          <div>
            <h3 className="text-sm font-semibold text-slate-900">
              Sub-Second Webhook &amp; Delta Stream
            </h3>
            <p className="text-xs text-slate-500">
              Real-time change feed · fingerprint dedup &amp; loop guard active
            </p>
          </div>
        </div>

        <button
          type="button"
          onClick={onRenewWatches}
          disabled={busyAction !== null}
          className="inline-flex items-center gap-1.5 rounded-md border border-slate-200 bg-slate-50 px-2.5 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-100 disabled:opacity-50"
        >
          <RefreshCw
            className={`h-3.5 w-3.5 ${busyAction === 'renew' ? 'animate-spin' : ''}`}
          />
          <span>Renew Watches</span>
        </button>
      </div>

      <div className="max-h-96 divide-y divide-slate-100 overflow-y-auto">
        {logs.length === 0 ? (
          <div className="p-6 text-center text-xs text-slate-500">
            No sync telemetry yet. Create, move, or delete any event to watch sub-second
            propagation.
          </div>
        ) : (
          logs.slice(0, 20).map((log) => {
            const sourceAcc = accountMap.get(log.sourceAccountId);
            const targetAcc = log.targetAccountId
              ? accountMap.get(log.targetAccountId)
              : undefined;
            const badge = ACTION_BADGE_MAP[log.action];
            const isExpanded = expandedLogId === log.id;

            return (
              <div key={log.id} className="px-4 py-2.5 text-xs transition hover:bg-slate-50/80">
                <div className="flex items-start justify-between gap-2">
                  <div className="min-w-0 flex-1 space-y-1">
                    <div className="flex flex-wrap items-center gap-1.5">
                      <span
                        className={`inline-flex items-center rounded px-1.5 py-0.5 font-mono text-[10px] font-semibold ring-1 ${badge.classes}`}
                      >
                        {badge.label}
                      </span>
                      <span className="inline-flex items-center gap-1 font-mono text-[11px] font-semibold text-emerald-700">
                        <Zap className="h-3 w-3" />
                        {log.latencyMs}ms
                      </span>
                      <span className="truncate font-medium text-slate-800">
                        {sourceAcc?.label ?? log.sourceAccountId}
                        {targetAcc ? ` → ${targetAcc.label}` : ''}
                      </span>
                    </div>

                    <p className="text-[11px] leading-relaxed text-slate-600">{log.reason}</p>
                  </div>

                  <div className="flex shrink-0 items-center gap-2">
                    {log.action === 'skipped_loop' ? (
                      <ShieldAlert className="h-4 w-4 text-violet-600" />
                    ) : (
                      <ShieldCheck className="h-4 w-4 text-emerald-600" />
                    )}
                    {log.payloadPreview && (
                      <button
                        type="button"
                        onClick={() =>
                          setExpandedLogId(isExpanded ? null : log.id)
                        }
                        className="inline-flex items-center gap-1 rounded border border-slate-200 px-1.5 py-0.5 text-[10px] font-medium text-slate-600 hover:bg-slate-100"
                      >
                        <span>JSON</span>
                        {isExpanded ? (
                          <ChevronUp className="h-3 w-3" />
                        ) : (
                          <ChevronDown className="h-3 w-3" />
                        )}
                      </button>
                    )}
                  </div>
                </div>

                {isExpanded && log.payloadPreview && (
                  <pre className="mt-2 max-h-48 overflow-auto rounded-lg bg-slate-900 p-2.5 font-mono text-[10px] text-emerald-300">
                    {JSON.stringify(log.payloadPreview, null, 2)}
                  </pre>
                )}
              </div>
            );
          })
        )}
      </div>
    </section>
  );
}
