import { useMemo, useState } from 'react';
import { Check, CheckCircle2, Shield, Sparkles, Trash2 } from 'lucide-react';
import { detectLegacySyncClutter } from '../../shared/sanitizer';
import type { BillingPlan, CalendarEvent } from '../../shared/types';

interface MigrationAndPricingProps {
  events: CalendarEvent[];
  currentPlan: BillingPlan;
  busyAction: string | null;
  onCleanupLegacy: () => Promise<void>;
  onSelectPlan: (plan: BillingPlan) => Promise<void>;
}

const COMPARISON_ROWS = [
  {
    feature: 'Core Behavior',
    stealth: 'Does one thing cleanly: mirrors unmovable "Busy" blocks across accounts',
    reclaim: 'Aggressive AI calendar takeover: dynamically shuffles your tasks and habits',
    syncgene: 'Legacy folder/account pair bridge with multi-step setup wizard',
  },
  {
    feature: 'Sync Speed',
    stealth: 'Sub-second real-time webhooks (Google Watch + Microsoft Graph Delta)',
    reclaim: 'Queue-based background re-scheduling cycles',
    syncgene: 'Up to 15–30 minute polling delay before updates appear',
  },
  {
    feature: 'Title & Metadata Privacy',
    stealth: 'Zero title/description/location/attendee leaks, private visibility enforced',
    reclaim: 'Injects habit blocks and smart annotations directly on your schedule',
    syncgene: 'Exposes sync tags and clutters titles across bridged calendars',
  },
  {
    feature: 'Connected Accounts',
    stealth: 'Unlimited Google & Outlook calendars included at one flat rate',
    reclaim: 'Per-seat tiered subscriptions for advanced calendar features',
    syncgene: 'Charges extra per connected account pair',
  },
];

export function MigrationAndPricing({
  events,
  currentPlan,
  busyAction,
  onCleanupLegacy,
  onSelectPlan,
}: MigrationAndPricingProps) {
  const [calendarCount, setCalendarCount] = useState(4);

  const clutterEvents = useMemo(
    () =>
      events
        .map((evt) => ({ evt, source: detectLegacySyncClutter(evt) }))
        .filter((item): item is { evt: CalendarEvent; source: 'reclaim' | 'syncgene' } =>
          item.source !== null
        ),
    [events]
  );

  const pairCount = Math.floor((calendarCount * (calendarCount - 1)) / 2);
  const estimatedSyncGeneMonthly = pairCount * 10;

  return (
    <div className="space-y-6">
      {/* Reclaim & SyncGene Migration Cleaner */}
      <section className="rounded-xl border border-slate-200 bg-white p-5 shadow-xs">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <h2 className="text-base font-bold text-slate-900">
              1-Click Migration &amp; Clutter Exit Cleaner
            </h2>
            <p className="mt-0.5 text-xs text-slate-500">
              Switching from Reclaim.ai or SyncGene? Scan connected calendars and strip legacy
              injected habit/sync clutter safely.
            </p>
          </div>

          <button
            type="button"
            onClick={() => void onCleanupLegacy()}
            disabled={busyAction !== null || clutterEvents.length === 0}
            className="inline-flex items-center gap-1.5 rounded-lg bg-slate-900 px-3.5 py-2 text-xs font-semibold text-white hover:bg-slate-800 disabled:opacity-40"
          >
            <Trash2 className="h-3.5 w-3.5 text-amber-400" />
            <span>
              {clutterEvents.length > 0
                ? `Purge ${clutterEvents.length} Clutter Item(s)`
                : 'Workspace Clean (0 Legacy Items)'}
            </span>
          </button>
        </div>

        {clutterEvents.length > 0 ? (
          <div className="mt-3 grid grid-cols-1 gap-2 sm:grid-cols-2">
            {clutterEvents.map(({ evt, source }) => (
              <div
                key={evt.id}
                className="rounded-lg border border-amber-200 bg-amber-50/70 p-3 text-xs"
              >
                <div className="flex items-center justify-between gap-2">
                  <span className="font-semibold text-amber-950">{evt.title}</span>
                  <span className="rounded bg-amber-200/70 px-1.5 py-0.5 text-[10px] font-bold uppercase text-amber-900">
                    {source}
                  </span>
                </div>
                <p className="mt-1 text-[11px] text-amber-800">{evt.description}</p>
              </div>
            ))}
          </div>
        ) : (
          <div className="mt-3 flex items-center gap-2 rounded-lg border border-emerald-200 bg-emerald-50 p-3 text-xs text-emerald-800">
            <CheckCircle2 className="h-4 w-4 text-emerald-600" />
            <span>Zero Reclaim.ai or SyncGene clutter blocks remain on any account.</span>
          </div>
        )}
      </section>

      {/* Transparent Flat Pricing & Calculator */}
      <div className="grid grid-cols-1 gap-5 lg:grid-cols-3">
        {/* Flat Monthly $5 */}
        <section
          className={`flex flex-col justify-between rounded-xl border p-5 shadow-xs ${
            currentPlan === 'monthly_5'
              ? 'border-slate-900 bg-white ring-2 ring-slate-900'
              : 'border-slate-200 bg-white'
          }`}
        >
          <div>
            <div className="flex items-center justify-between">
              <h3 className="text-sm font-bold text-slate-900">Flat Monthly</h3>
              {currentPlan === 'monthly_5' && (
                <span className="rounded-full bg-slate-900 px-2 py-0.5 text-[10px] font-semibold text-white">
                  ACTIVE PLAN
                </span>
              )}
            </div>
            <div className="mt-3 flex items-baseline gap-1">
              <span className="font-mono text-3xl font-extrabold text-slate-900">$5</span>
              <span className="text-xs text-slate-500">/month · unlimited calendars</span>
            </div>
            <ul className="mt-4 space-y-2 text-xs text-slate-600">
              {[
                'Unlimited Google Workspace & Outlook 365 accounts',
                'Sub-second push sync + automatic watch renewal',
                'Zero-leak private "Busy" blocks & edit locking',
                '1-click Reclaim / SyncGene clutter exit cleaner',
              ].map((perk) => (
                <li key={perk} className="flex items-start gap-2">
                  <Check className="mt-0.5 h-3.5 w-3.5 shrink-0 text-emerald-600" />
                  <span>{perk}</span>
                </li>
              ))}
            </ul>
          </div>

          <button
            type="button"
            onClick={() => void onSelectPlan('monthly_5')}
            disabled={busyAction !== null || currentPlan === 'monthly_5'}
            className="mt-5 w-full rounded-lg bg-slate-900 px-4 py-2 text-xs font-semibold text-white hover:bg-slate-800 disabled:opacity-50"
          >
            {currentPlan === 'monthly_5' ? 'Selected ($5/mo)' : 'Switch to $5/mo Flat'}
          </button>
        </section>

        {/* Lifetime $39 */}
        <section
          className={`flex flex-col justify-between rounded-xl border p-5 shadow-xs ${
            currentPlan === 'lifetime_39'
              ? 'border-emerald-600 bg-slate-900 text-white ring-2 ring-emerald-500'
              : 'border-slate-800 bg-slate-900 text-white'
          }`}
        >
          <div>
            <div className="flex items-center justify-between">
              <span className="inline-flex items-center gap-1 text-sm font-bold text-white">
                <Sparkles className="h-4 w-4 text-emerald-400" />
                Lifetime Deal
              </span>
              {currentPlan === 'lifetime_39' ? (
                <span className="rounded-full bg-emerald-500 px-2 py-0.5 text-[10px] font-bold text-slate-950">
                  UNLOCKED FOREVER
                </span>
              ) : (
                <span className="rounded-full bg-slate-800 px-2 py-0.5 text-[10px] font-semibold text-emerald-300">
                  MOST POPULAR
                </span>
              )}
            </div>
            <div className="mt-3 flex items-baseline gap-1">
              <span className="font-mono text-3xl font-extrabold text-white">$39</span>
              <span className="text-xs text-slate-300">one-time · forever</span>
            </div>
            <ul className="mt-4 space-y-2 text-xs text-slate-300">
              {[
                'Pay once, sync unlimited work/client/personal calendars',
                'Pays for itself vs monthly per-pair tools in months',
                'Real-time loop guard & fingerprint deduplication',
                'Includes all future Microsoft Graph & Google updates',
              ].map((perk) => (
                <li key={perk} className="flex items-start gap-2">
                  <Check className="mt-0.5 h-3.5 w-3.5 shrink-0 text-emerald-400" />
                  <span>{perk}</span>
                </li>
              ))}
            </ul>
          </div>

          <button
            type="button"
            onClick={() => void onSelectPlan('lifetime_39')}
            disabled={busyAction !== null || currentPlan === 'lifetime_39'}
            className="mt-5 w-full rounded-lg bg-emerald-500 px-4 py-2 text-xs font-bold text-slate-950 hover:bg-emerald-400 disabled:opacity-60"
          >
            {currentPlan === 'lifetime_39'
              ? 'Active ($39 Lifetime)'
              : 'Select $39 Lifetime'}
          </button>
        </section>

        {/* Interactive Cost Calculator */}
        <section className="flex flex-col justify-between rounded-xl border border-slate-200 bg-white p-5 shadow-xs">
          <div>
            <div className="flex items-center gap-2 text-sm font-bold text-slate-900">
              <Shield className="h-4 w-4 text-emerald-600" />
              <span>Per-Pair Fee Calculator</span>
            </div>
            <p className="mt-1 text-xs text-slate-500">
              Full mesh sync across {calendarCount} calendars requires{' '}
              <span className="font-mono font-semibold text-slate-900">{pairCount}</span>{' '}
              bidirectional pairs.
            </p>

            <div className="mt-4">
              <div className="flex items-center justify-between text-xs font-medium text-slate-700">
                <span>Connected Calendars:</span>
                <span className="font-mono text-sm font-bold text-slate-900">
                  {calendarCount} calendars
                </span>
              </div>
              <input
                type="range"
                min={2}
                max={10}
                value={calendarCount}
                onChange={(e) => setCalendarCount(Number(e.target.value))}
                className="mt-2 w-full accent-slate-900"
              />
            </div>

            <div className="mt-4 space-y-2 rounded-lg bg-slate-50 p-3 text-xs">
              <div className="flex items-center justify-between">
                <span className="text-slate-600">Est. SyncGene per-pair ({pairCount} pairs)</span>
                <span className="font-mono font-bold text-rose-600">
                  ${estimatedSyncGeneMonthly}/mo
                </span>
              </div>
              <div className="flex items-center justify-between">
                <span className="text-slate-600">StealthMirror Unlimited</span>
                <span className="font-mono font-bold text-emerald-700">$5/mo</span>
              </div>
              <div className="border-t border-slate-200 pt-2 text-[11px] font-semibold text-slate-800">
                Annual savings vs per-pair pricing:{' '}
                <span className="font-mono text-emerald-700">
                  ${Math.max(0, (estimatedSyncGeneMonthly - 5) * 12)}/yr
                </span>
              </div>
            </div>
          </div>

          <p className="mt-3 text-[11px] text-slate-400">
            Example illustration based on $10/pair/month scaling for bidirectional mesh links.
          </p>
        </section>
      </div>

      {/* Competitor Matrix */}
      <section className="overflow-hidden rounded-xl border border-slate-200 bg-white shadow-xs">
        <div className="border-b border-slate-200 px-5 py-3.5">
          <h3 className="text-sm font-bold text-slate-900">
            Why Teams Switch to StealthMirror
          </h3>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[700px] border-collapse text-left text-xs">
            <thead>
              <tr className="border-b border-slate-200 bg-slate-50 text-[11px] uppercase tracking-wide text-slate-500">
                <th className="px-4 py-2.5 font-medium">Dimension</th>
                <th className="px-4 py-2.5 font-bold text-slate-900">StealthMirror</th>
                <th className="px-4 py-2.5 font-medium">Reclaim.ai</th>
                <th className="px-4 py-2.5 font-medium">SyncGene</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {COMPARISON_ROWS.map((row) => (
                <tr key={row.feature} className="align-top">
                  <td className="px-4 py-3 font-semibold text-slate-800">{row.feature}</td>
                  <td className="bg-emerald-50/35 px-4 py-3 font-medium text-slate-900">
                    {row.stealth}
                  </td>
                  <td className="px-4 py-3 text-slate-600">{row.reclaim}</td>
                  <td className="px-4 py-3 text-slate-600">{row.syncgene}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}
