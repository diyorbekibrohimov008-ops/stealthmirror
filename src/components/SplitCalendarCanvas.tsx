import { useMemo, useState } from 'react';
import {
  AlertTriangle,
  ArrowRight,
  CalendarPlus,
  Clock,
  Eye,
  Lock,
  Plus,
  ShieldCheck,
  Sparkles,
  Trash2,
  Zap,
} from 'lucide-react';
import {
  buildGoogleBusyPayload,
  buildOutlookBusyPayload,
  detectLegacySyncClutter,
  isStealthManagedEvent,
} from '../../shared/sanitizer';
import type {
  CalendarAccount,
  CalendarEvent,
  EventResponseStatus,
  EventShowAs,
  SyncLogEntry,
  SyncRule,
} from '../../shared/types';
import type { OAuthStatus, WorkspaceStats } from '../lib/api';
import {
  buildUtcIsoForDayHour,
  formatUtcTimeRange,
  getUtcDayKey,
  getUtcTodayAndTomorrowKeys,
  shiftIsoHours,
} from '../lib/date';
import { WebhookTelemetry } from './WebhookTelemetry';

interface SplitCalendarCanvasProps {
  accounts: CalendarAccount[];
  rules: SyncRule[];
  events: CalendarEvent[];
  logs: SyncLogEntry[];
  stats: WorkspaceStats;
  oauth: OAuthStatus;
  busyAction: string | null;
  onCreateEvent: (input: {
    accountId: string;
    title: string;
    description: string;
    location: string;
    attendees: string[];
    startIso: string;
    endIso: string;
    allDay: boolean;
    showAs: EventShowAs;
    responseStatus: EventResponseStatus;
  }) => Promise<void>;
  onUpdateEvent: (
    eventId: string,
    patch: {
      startIso?: string;
      endIso?: string;
      showAs?: EventShowAs;
      responseStatus?: EventResponseStatus;
    }
  ) => Promise<void>;
  onDeleteEvent: (eventId: string) => Promise<void>;
  onTestLoopGuard: (eventId?: string) => Promise<void>;
  onCleanupLegacy: () => Promise<void>;
  onRenewWatches: () => Promise<void>;
  onOpenConnectTab: () => void;
}

type DayFilter = 'all' | 'today' | 'tomorrow';

function ProviderPill({ provider }: { provider: CalendarAccount['provider'] }) {
  return provider === 'google' ? (
    <span className="inline-flex items-center rounded bg-sky-50 px-1.5 py-0.5 text-[10px] font-semibold text-sky-700 ring-1 ring-sky-200">
      Google
    </span>
  ) : (
    <span className="inline-flex items-center rounded bg-indigo-50 px-1.5 py-0.5 text-[10px] font-semibold text-indigo-700 ring-1 ring-indigo-200">
      Outlook 365
    </span>
  );
}

export function SplitCalendarCanvas({
  accounts,
  rules,
  events,
  logs,
  stats,
  busyAction,
  onCreateEvent,
  onUpdateEvent,
  onDeleteEvent,
  onTestLoopGuard,
  onCleanupLegacy,
  onRenewWatches,
  onOpenConnectTab,
}: SplitCalendarCanvasProps) {
  const days = useMemo(() => getUtcTodayAndTomorrowKeys(), []);
  const [dayFilter, setDayFilter] = useState<DayFilter>('all');
  const [showComposer, setShowComposer] = useState(false);

  const [accountId, setAccountId] = useState<string>(accounts[0]?.id ?? '');
  const [title, setTitle] = useState('Acquisition Term Sheet Review');
  const [description, setDescription] = useState(
    'Confidential folder & valuation spreadsheet — do not share with external contractors.'
  );
  const [location, setLocation] = useState('Executive Boardroom / Zoom');
  const [attendeesRaw, setAttendeesRaw] = useState('ceo@acme.io, counsel@firm.com');
  const [dayKey, setDayKey] = useState(days.todayKey);
  const [startHour, setStartHour] = useState(14);
  const [durationHours, setDurationHours] = useState(1);
  const [showAs, setShowAs] = useState<EventShowAs>('busy');
  const [responseStatus, setResponseStatus] = useState<EventResponseStatus>('accepted');
  const [allDay, setAllDay] = useState(false);

  const [inspectedEventId, setInspectedEventId] = useState<string | null>(null);

  const selectedAccountId = accounts.some((a) => a.id === accountId)
    ? accountId
    : accounts[0]?.id ?? '';

  const accountMap = useMemo(
    () => new Map(accounts.map((a) => [a.id, a])),
    [accounts]
  );

  const inspectedPair = useMemo(() => {
    const fallbackMirror =
      events.find((e) => isStealthManagedEvent(e)) ?? events[0] ?? null;
    const chosen =
      events.find((e) => e.id === inspectedEventId) ?? fallbackMirror;
    if (!chosen) return null;

    if (isStealthManagedEvent(chosen) && chosen.stealthMirrorMeta) {
      const originEvent = events.find(
        (e) =>
          e.accountId === chosen.stealthMirrorMeta?.originAccountId &&
          e.externalEventId === chosen.stealthMirrorMeta?.originEventId
      );
      const targetAccount = accountMap.get(chosen.accountId);
      const rule = rules.find((r) => r.id === chosen.stealthMirrorMeta?.ruleId);
      const samplePayload =
        originEvent && rule && targetAccount
          ? targetAccount.provider === 'google'
            ? buildGoogleBusyPayload(originEvent, rule)
            : buildOutlookBusyPayload(originEvent, rule)
          : undefined;
      return { sourceEvent: originEvent ?? chosen, mirrorEvent: chosen, targetAccount, samplePayload };
    }

    const mirrorChild = events.find(
      (e) =>
        isStealthManagedEvent(e) &&
        e.stealthMirrorMeta?.originAccountId === chosen.accountId &&
        e.stealthMirrorMeta?.originEventId === chosen.externalEventId
    );
    const targetAccount = mirrorChild
      ? accountMap.get(mirrorChild.accountId)
      : accountMap.get(chosen.accountId);
    const rule = mirrorChild
      ? rules.find((r) => r.id === mirrorChild.stealthMirrorMeta?.ruleId)
      : rules.find((r) => r.sourceAccountId === chosen.accountId);
    const samplePayload =
      rule && targetAccount
        ? targetAccount.provider === 'google'
          ? buildGoogleBusyPayload(chosen, rule)
          : buildOutlookBusyPayload(chosen, rule)
        : undefined;

    return {
      sourceEvent: chosen,
      mirrorEvent: mirrorChild ?? null,
      targetAccount,
      samplePayload,
    };
  }, [events, inspectedEventId, accountMap, rules]);

  async function handleComposerSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!selectedAccountId || !title.trim()) return;
    const startIso = allDay
      ? buildUtcIsoForDayHour(dayKey, 0, 0)
      : buildUtcIsoForDayHour(dayKey, startHour, 0);
    const endIso = allDay
      ? buildUtcIsoForDayHour(dayKey, 23, 59)
      : buildUtcIsoForDayHour(dayKey, startHour + durationHours, 0);

    await onCreateEvent({
      accountId: selectedAccountId,
      title: title.trim(),
      description: description.trim(),
      location: location.trim(),
      attendees: attendeesRaw
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean),
      startIso,
      endIso,
      allDay,
      showAs,
      responseStatus,
    });
    setShowComposer(false);
  }

  async function handleQuickConfidentialEvent() {
    if (!selectedAccountId) return;
    await onCreateEvent({
      accountId: selectedAccountId,
      title: 'Confidential M&A LOI Negotiation',
      description: 'Internal link: https://vault.acme.io/loi-term-sheet — strict NDA',
      location: 'Private Suite 1200',
      attendees: ['cfo@acme.io', 'partner@lawfirm.com'],
      startIso: buildUtcIsoForDayHour(days.todayKey, 15, 0),
      endIso: buildUtcIsoForDayHour(days.todayKey, 16, 0),
      allDay: false,
      showAs: 'busy',
      responseStatus: 'accepted',
    });
  }

  function filterEventsForAccount(accId: string): CalendarEvent[] {
    return events
      .filter((evt) => {
        if (evt.accountId !== accId) return false;
        const evtDay = getUtcDayKey(evt.startIso);
        if (dayFilter === 'today') return evtDay === days.todayKey;
        if (dayFilter === 'tomorrow') return evtDay === days.tomorrowKey;
        return true;
      })
      .sort((a, b) => a.startIso.localeCompare(b.startIso));
  }

  return (
    <div className="space-y-5">
      {/* Positioning + KPI Strip */}
      <section className="grid grid-cols-2 gap-3 lg:grid-cols-4">
        <div className="rounded-xl border border-slate-200 bg-white p-3.5 shadow-xs">
          <div className="flex items-center justify-between text-xs text-slate-500">
            <span>Push Latency</span>
            <Zap className="h-3.5 w-3.5 text-emerald-600" />
          </div>
          <div className="mt-1 flex items-baseline gap-2">
            <span className="font-mono text-2xl font-bold text-slate-900">
              {stats.avgLatencyMs}ms
            </span>
            <span className="text-[11px] font-medium text-emerald-700">
              vs 15–30m SyncGene
            </span>
          </div>
        </div>

        <div className="rounded-xl border border-slate-200 bg-white p-3.5 shadow-xs">
          <div className="flex items-center justify-between text-xs text-slate-500">
            <span>Metadata Protection</span>
            <Lock className="h-3.5 w-3.5 text-slate-800" />
          </div>
          <div className="mt-1 flex items-baseline gap-2">
            <span className="font-mono text-2xl font-bold text-slate-900">
              {stats.leaksCount} Leaks
            </span>
            <span className="text-[11px] text-slate-500">
              {stats.mirroredBlockCount} Busy mirrors locked
            </span>
          </div>
        </div>

        <div className="rounded-xl border border-slate-200 bg-white p-3.5 shadow-xs">
          <div className="flex items-center justify-between text-xs text-slate-500">
            <span>Loop Guard Engine</span>
            <ShieldCheck className="h-3.5 w-3.5 text-violet-600" />
          </div>
          <div className="mt-1 flex items-baseline justify-between gap-2">
            <span className="font-mono text-2xl font-bold text-slate-900">
              {stats.loopsBlocked} Blocked
            </span>
            <button
              type="button"
              onClick={() => void onTestLoopGuard()}
              disabled={busyAction !== null || stats.mirroredBlockCount === 0}
              className="rounded bg-violet-50 px-2 py-0.5 text-[11px] font-semibold text-violet-700 ring-1 ring-violet-200 hover:bg-violet-100 disabled:opacity-40"
            >
              Test Loop
            </button>
          </div>
        </div>

        <div className="rounded-xl border border-slate-200 bg-white p-3.5 shadow-xs">
          <div className="flex items-center justify-between text-xs text-slate-500">
            <span>Connected Accounts</span>
            <Sparkles className="h-3.5 w-3.5 text-sky-600" />
          </div>
          <div className="mt-1 flex items-baseline justify-between gap-2">
            <span className="font-mono text-2xl font-bold text-slate-900">
              {stats.accountCount} / &infin;
            </span>
            <span className="text-[11px] font-medium text-emerald-700">
              $5 flat (saves ~${Math.max(0, stats.estimatedSyncGeneMonthlyUsd - 5)}/mo)
            </span>
          </div>
        </div>
      </section>

      {/* Legacy clutter banner if Reclaim or SyncGene clutter exists */}
      {stats.legacyClutterCount > 0 && (
        <div className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-amber-300 bg-amber-50/90 px-4 py-3 text-xs text-amber-950">
          <div className="flex items-center gap-2.5">
            <AlertTriangle className="h-4 w-4 shrink-0 text-amber-600" />
            <div>
              <span className="font-semibold">
                {stats.legacyClutterCount} legacy Reclaim.ai / SyncGene clutter event(s)
              </span>{' '}
              still sit on your calendars. Purge them cleanly without touching real meetings.
            </div>
          </div>
          <button
            type="button"
            onClick={() => void onCleanupLegacy()}
            disabled={busyAction !== null}
            className="inline-flex items-center gap-1.5 rounded-lg bg-amber-600 px-3 py-1.5 font-semibold text-white shadow-xs hover:bg-amber-700 disabled:opacity-50"
          >
            <Trash2 className="h-3.5 w-3.5" />
            <span>Purge Legacy Bloat</span>
          </button>
        </div>
      )}

      {/* Action Toolbar */}
      <div className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-slate-200 bg-white px-4 py-3 shadow-xs">
        <div className="flex flex-wrap items-center gap-1.5">
          {(
            [
              { id: 'all', label: '2-Day Split (Today + Tomorrow)' },
              { id: 'today', label: days.todayLabel },
              { id: 'tomorrow', label: days.tomorrowLabel },
            ] as const
          ).map((option) => (
            <button
              key={option.id}
              type="button"
              onClick={() => setDayFilter(option.id)}
              className={`rounded-md px-2.5 py-1.5 text-xs font-medium transition ${
                dayFilter === option.id
                  ? 'bg-slate-900 text-white'
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              {option.label}
            </button>
          ))}
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={() => void handleQuickConfidentialEvent()}
            disabled={busyAction !== null || accounts.length === 0}
            className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-slate-50 px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-100 disabled:opacity-50"
          >
            <Zap className="h-3.5 w-3.5 text-emerald-600" />
            <span>1-Click M&amp;A Sync Test</span>
          </button>

          <button
            type="button"
            onClick={() => setShowComposer((prev) => !prev)}
            disabled={accounts.length === 0}
            className="inline-flex items-center gap-1.5 rounded-lg bg-slate-900 px-3.5 py-1.5 text-xs font-semibold text-white shadow-xs hover:bg-slate-800 disabled:opacity-50"
          >
            <CalendarPlus className="h-3.5 w-3.5 text-emerald-400" />
            <span>+ New Event</span>
          </button>

          <button
            type="button"
            onClick={onOpenConnectTab}
            className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50"
          >
            <Plus className="h-3.5 w-3.5" />
            <span>Add Calendar</span>
          </button>
        </div>
      </div>

      {/* Inline Event Composer */}
      {showComposer && (
        <form
          onSubmit={(e) => void handleComposerSubmit(e)}
          className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm"
        >
          <div className="mb-3 flex items-center justify-between">
            <h3 className="text-sm font-semibold text-slate-900">
              Create Source Meeting (Watch Instant Stealth Mirror)
            </h3>
            <button
              type="button"
              onClick={() => setShowComposer(false)}
              className="text-xs text-slate-500 hover:text-slate-900"
            >
              Close
            </button>
          </div>

          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4">
            <label className="text-xs">
              <span className="mb-1 block font-medium text-slate-700">Source Calendar</span>
              <select
                value={selectedAccountId}
                onChange={(e) => setAccountId(e.target.value)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs text-slate-900 outline-none focus:border-slate-900"
              >
                {accounts.map((acc) => (
                  <option key={acc.id} value={acc.id}>
                    {acc.label} ({acc.provider})
                  </option>
                ))}
              </select>
            </label>

            <label className="text-xs sm:col-span-2">
              <span className="mb-1 block font-medium text-slate-700">
                Confidential Meeting Title (Never Mirrored)
              </span>
              <input
                type="text"
                required
                value={title}
                onChange={(e) => setTitle(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-2.5 py-2 text-xs text-slate-900 outline-none focus:border-slate-900"
                placeholder="e.g., Executive Comp Review"
              />
            </label>

            <label className="text-xs">
              <span className="mb-1 block font-medium text-slate-700">Day (UTC)</span>
              <select
                value={dayKey}
                onChange={(e) => setDayKey(e.target.value)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs text-slate-900 outline-none focus:border-slate-900"
              >
                <option value={days.todayKey}>{days.todayLabel}</option>
                <option value={days.tomorrowKey}>{days.tomorrowLabel}</option>
              </select>
            </label>

            <label className="text-xs">
              <span className="mb-1 block font-medium text-slate-700">Start Hour (UTC)</span>
              <select
                value={startHour}
                disabled={allDay}
                onChange={(e) => setStartHour(Number(e.target.value))}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs text-slate-900 outline-none focus:border-slate-900 disabled:bg-slate-100"
              >
                {Array.from({ length: 14 }, (_, idx) => idx + 7).map((hr) => (
                  <option key={hr} value={hr}>
                    {String(hr).padStart(2, '0')}:00 UTC
                  </option>
                ))}
              </select>
            </label>

            <label className="text-xs">
              <span className="mb-1 block font-medium text-slate-700">Duration</span>
              <select
                value={durationHours}
                disabled={allDay}
                onChange={(e) => setDurationHours(Number(e.target.value))}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs text-slate-900 outline-none focus:border-slate-900 disabled:bg-slate-100"
              >
                <option value={1}>60 min</option>
                <option value={2}>120 min</option>
              </select>
            </label>

            <label className="text-xs">
              <span className="mb-1 block font-medium text-slate-700">Show As Policy</span>
              <select
                value={showAs}
                onChange={(e) => setShowAs(e.target.value as EventShowAs)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs text-slate-900 outline-none focus:border-slate-900"
              >
                <option value="busy">Busy (Mirrors to targets)</option>
                <option value="oof">Out of Office (Mirrors)</option>
                <option value="free">Free (Skipped cleanly)</option>
                <option value="tentative">Tentative (Mirrors)</option>
              </select>
            </label>

            <label className="text-xs">
              <span className="mb-1 block font-medium text-slate-700">RSVP Status</span>
              <select
                value={responseStatus}
                onChange={(e) =>
                  setResponseStatus(e.target.value as EventResponseStatus)
                }
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs text-slate-900 outline-none focus:border-slate-900"
              >
                <option value="accepted">Accepted</option>
                <option value="declined">Declined (Skipped cleanly)</option>
              </select>
            </label>

            <label className="text-xs sm:col-span-2">
              <span className="mb-1 block font-medium text-slate-700">
                Secret Notes / URL (Stripped on Mirror)
              </span>
              <input
                type="text"
                value={description}
                onChange={(e) => setDescription(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-2.5 py-2 text-xs text-slate-900 outline-none focus:border-slate-900"
              />
            </label>

            <label className="text-xs">
              <span className="mb-1 block font-medium text-slate-700">
                Location (Stripped)
              </span>
              <input
                type="text"
                value={location}
                onChange={(e) => setLocation(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-2.5 py-2 text-xs text-slate-900 outline-none focus:border-slate-900"
              />
            </label>

            <label className="text-xs">
              <span className="mb-1 block font-medium text-slate-700">
                Attendees CSV (Stripped)
              </span>
              <input
                type="text"
                value={attendeesRaw}
                onChange={(e) => setAttendeesRaw(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-2.5 py-2 text-xs text-slate-900 outline-none focus:border-slate-900"
              />
            </label>
          </div>

          <div className="mt-3 flex flex-wrap items-center justify-between gap-3 border-t border-slate-100 pt-3">
            <label className="inline-flex items-center gap-2 text-xs text-slate-700">
              <input
                type="checkbox"
                checked={allDay}
                onChange={(e) => setAllDay(e.target.checked)}
                className="h-4 w-4 rounded border-slate-300 text-slate-900"
              />
              <span>All-Day event (skipped by default per All-Day policy)</span>
            </label>

            <button
              type="submit"
              disabled={busyAction !== null}
              className="inline-flex items-center gap-1.5 rounded-lg bg-slate-900 px-4 py-2 text-xs font-semibold text-white hover:bg-slate-800 disabled:opacity-50"
            >
              <Zap className="h-3.5 w-3.5 text-emerald-400" />
              <span>Save &amp; Sync Across Accounts</span>
            </button>
          </div>
        </form>
      )}

      {/* Multi-Account Side-by-Side Calendar Grid */}
      {accounts.length === 0 ? (
        <div className="rounded-xl border border-dashed border-slate-300 bg-white p-10 text-center">
          <p className="text-sm font-semibold text-slate-900">
            No calendars connected in this workspace yet.
          </p>
          <p className="mt-1 text-xs text-slate-500">
            Add unlimited Google Workspace or Microsoft Outlook calendars to start sub-second
            mirroring.
          </p>
          <button
            type="button"
            onClick={onOpenConnectTab}
            className="mt-4 inline-flex items-center gap-1.5 rounded-lg bg-slate-900 px-4 py-2 text-xs font-semibold text-white hover:bg-slate-800"
          >
            <Plus className="h-3.5 w-3.5" />
            <span>Connect First Calendar</span>
          </button>
        </div>
      ) : (
        <div className="grid grid-cols-1 gap-4 lg:grid-cols-3">
          {accounts.map((acc) => {
            const accEvents = filterEventsForAccount(acc.id);
            const outboundRules = rules.filter(
              (r) => r.sourceAccountId === acc.id && r.enabled
            );

            return (
              <div
                key={acc.id}
                className="flex flex-col rounded-xl border border-slate-200 bg-white shadow-xs"
              >
                {/* Account Header */}
                <div className="border-b border-slate-200 px-4 py-3">
                  <div className="flex items-start justify-between gap-2">
                    <div className="min-w-0">
                      <div className="flex items-center gap-2">
                        <span
                          className="h-2.5 w-2.5 shrink-0 rounded-full"
                          style={{ backgroundColor: acc.color }}
                        />
                        <h3 className="truncate text-sm font-semibold text-slate-900">
                          {acc.label}
                        </h3>
                        <ProviderPill provider={acc.provider} />
                      </div>
                      <p className="mt-0.5 truncate text-xs text-slate-500">{acc.email}</p>
                    </div>

                    <span className="shrink-0 rounded-full bg-emerald-50 px-2 py-0.5 text-[10px] font-semibold text-emerald-700 ring-1 ring-emerald-200">
                      {acc.mode === 'oauth' ? 'Live OAuth' : 'Live Sandbox'}
                    </span>
                  </div>

                  <div className="mt-2 flex items-center justify-between text-[11px] text-slate-500">
                    <span>Mirrors out to {outboundRules.length} target(s)</span>
                    <span className="font-mono text-[10px] text-slate-400">
                      {accEvents.length} event(s)
                    </span>
                  </div>
                </div>

                {/* Events Column */}
                <div className="flex-1 space-y-2.5 p-3">
                  {accEvents.length === 0 ? (
                    <div className="rounded-lg border border-dashed border-slate-200 p-5 text-center text-xs text-slate-400">
                      No events in this date window.
                    </div>
                  ) : (
                    accEvents.map((evt) => {
                      const isMirror = isStealthManagedEvent(evt);
                      const legacySource = detectLegacySyncClutter(evt);
                      const dayTag =
                        getUtcDayKey(evt.startIso) === days.todayKey
                          ? 'Today'
                          : getUtcDayKey(evt.startIso) === days.tomorrowKey
                          ? 'Tomorrow'
                          : getUtcDayKey(evt.startIso);
                      const timeText = formatUtcTimeRange(
                        evt.startIso,
                        evt.endIso,
                        evt.allDay
                      );

                      if (isMirror) {
                        const originAcc = evt.stealthMirrorMeta
                          ? accountMap.get(evt.stealthMirrorMeta.originAccountId)
                          : undefined;

                        return (
                          <div
                            key={evt.id}
                            className="stealth-hatch rounded-lg border border-slate-800 bg-slate-900 p-3 text-white shadow-xs"
                          >
                            <div className="flex items-start justify-between gap-2">
                              <div className="flex items-center gap-1.5">
                                <Lock className="h-3.5 w-3.5 text-emerald-400" />
                                <span className="text-xs font-bold tracking-wide">
                                  {evt.title}
                                </span>
                              </div>
                              <span className="rounded bg-slate-800 px-1.5 py-0.5 font-mono text-[10px] text-emerald-300 ring-1 ring-slate-700">
                                LOCKED · PRIVATE
                              </span>
                            </div>

                            <div className="mt-1.5 flex flex-wrap items-center gap-2 text-[11px] text-slate-300">
                              <span className="inline-flex items-center gap-1">
                                <Clock className="h-3 w-3 text-slate-400" />
                                {dayTag} · {timeText}
                              </span>
                              {originAcc && (
                                <span className="rounded bg-slate-800 px-1.5 py-0.5 text-[10px] text-slate-300">
                                  Origin: {originAcc.label} (Zero Title Leak)
                                </span>
                              )}
                            </div>

                            <div className="mt-2.5 flex flex-wrap items-center gap-1.5">
                              <button
                                type="button"
                                onClick={() =>
                                  void onUpdateEvent(evt.id, {
                                    startIso: shiftIsoHours(evt.startIso, 1),
                                    endIso: shiftIsoHours(evt.endIso, 1),
                                  })
                                }
                                disabled={busyAction !== null}
                                className="rounded bg-slate-800 px-2 py-1 text-[10px] font-medium text-amber-300 ring-1 ring-slate-700 hover:bg-slate-700"
                              >
                                Try Edit (+1h Lock Test)
                              </button>

                              <button
                                type="button"
                                onClick={() => void onTestLoopGuard(evt.id)}
                                disabled={busyAction !== null}
                                className="rounded bg-slate-800 px-2 py-1 text-[10px] font-medium text-violet-300 ring-1 ring-slate-700 hover:bg-slate-700"
                              >
                                Test Loop Guard
                              </button>

                              <button
                                type="button"
                                onClick={() => setInspectedEventId(evt.id)}
                                className="inline-flex items-center gap-1 rounded bg-slate-800 px-2 py-1 text-[10px] font-medium text-emerald-300 ring-1 ring-slate-700 hover:bg-slate-700"
                              >
                                <Eye className="h-3 w-3" />
                                <span>JSON</span>
                              </button>
                            </div>
                          </div>
                        );
                      }

                      if (legacySource) {
                        return (
                          <div
                            key={evt.id}
                            className="rounded-lg border border-dashed border-amber-300 bg-amber-50/70 p-3 text-xs"
                          >
                            <div className="flex items-start justify-between gap-2">
                              <span className="font-semibold text-amber-950">{evt.title}</span>
                              <span className="shrink-0 rounded bg-amber-100 px-1.5 py-0.5 text-[10px] font-semibold uppercase text-amber-800">
                                {legacySource} bloat
                              </span>
                            </div>
                            <p className="mt-1 text-[11px] text-amber-800">
                              {dayTag} · {timeText}
                            </p>
                            {evt.description && (
                              <p className="mt-1 text-[11px] text-amber-700/90">
                                {evt.description}
                              </p>
                            )}
                            <div className="mt-2">
                              <button
                                type="button"
                                onClick={() => void onCleanupLegacy()}
                                disabled={busyAction !== null}
                                className="inline-flex items-center gap-1 rounded bg-amber-600 px-2 py-1 text-[10px] font-semibold text-white hover:bg-amber-700"
                              >
                                <Trash2 className="h-3 w-3" />
                                <span>Purge Legacy Bloat</span>
                              </button>
                            </div>
                          </div>
                        );
                      }

                      return (
                        <div
                          key={evt.id}
                          className="rounded-lg border border-slate-200 bg-white p-3 shadow-2xs transition hover:border-slate-300"
                          style={{ borderLeftWidth: '4px', borderLeftColor: acc.color }}
                        >
                          <div className="flex items-start justify-between gap-2">
                            <span className="text-xs font-semibold text-slate-900">
                              {evt.title}
                            </span>
                            <span
                              className={`shrink-0 rounded px-1.5 py-0.5 text-[10px] font-semibold uppercase ${
                                evt.showAs === 'free'
                                  ? 'bg-slate-100 text-slate-600'
                                  : 'bg-sky-50 text-sky-700'
                              }`}
                            >
                              {evt.showAs}
                            </span>
                          </div>

                          <p className="mt-1 flex items-center gap-1 text-[11px] text-slate-500">
                            <Clock className="h-3 w-3" />
                            <span>
                              {dayTag} · {timeText}
                            </span>
                          </p>

                          {(evt.location || evt.attendees.length > 0) && (
                            <p className="mt-1 truncate text-[11px] text-slate-500">
                              {evt.location ? `📍 ${evt.location}` : ''}
                              {evt.location && evt.attendees.length > 0 ? ' · ' : ''}
                              {evt.attendees.length > 0
                                ? `👥 ${evt.attendees.length} attendee(s)`
                                : ''}
                            </p>
                          )}

                          <div className="mt-2.5 flex flex-wrap items-center gap-1.5">
                            <button
                              type="button"
                              onClick={() =>
                                void onUpdateEvent(evt.id, {
                                  startIso: shiftIsoHours(evt.startIso, 1),
                                  endIso: shiftIsoHours(evt.endIso, 1),
                                })
                              }
                              disabled={busyAction !== null}
                              className="rounded border border-slate-200 bg-slate-50 px-2 py-1 text-[10px] font-medium text-slate-700 hover:bg-slate-100"
                            >
                              +1h Move
                            </button>

                            <button
                              type="button"
                              onClick={() =>
                                void onUpdateEvent(evt.id, {
                                  showAs: evt.showAs === 'free' ? 'busy' : 'free',
                                })
                              }
                              disabled={busyAction !== null}
                              className="rounded border border-slate-200 bg-slate-50 px-2 py-1 text-[10px] font-medium text-slate-700 hover:bg-slate-100"
                            >
                              {evt.showAs === 'free' ? 'Mark Busy' : 'Mark Free'}
                            </button>

                            <button
                              type="button"
                              onClick={() => setInspectedEventId(evt.id)}
                              className="inline-flex items-center gap-1 rounded border border-slate-200 bg-slate-50 px-2 py-1 text-[10px] font-medium text-slate-700 hover:bg-slate-100"
                            >
                              <Eye className="h-3 w-3" />
                              <span>Inspect</span>
                            </button>

                            <button
                              type="button"
                              onClick={() => void onDeleteEvent(evt.id)}
                              disabled={busyAction !== null}
                              className="ml-auto inline-flex items-center gap-1 rounded border border-rose-200 bg-rose-50 px-2 py-1 text-[10px] font-medium text-rose-700 hover:bg-rose-100"
                            >
                              <Trash2 className="h-3 w-3" />
                            </button>
                          </div>
                        </div>
                      );
                    })
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* Bottom Split: Leak-Proof Wire Inspector + Webhook Stream */}
      <div className="grid grid-cols-1 gap-4 xl:grid-cols-2">
        <section className="rounded-xl border border-slate-200 bg-white shadow-xs">
          <div className="flex items-center justify-between border-b border-slate-200 px-4 py-3">
            <div className="flex items-center gap-2">
              <div className="flex h-7 w-7 items-center justify-center rounded-md bg-slate-900 text-emerald-400">
                <Lock className="h-4 w-4" />
              </div>
              <div>
                <h3 className="text-sm font-semibold text-slate-900">
                  Leak-Proof API Payload Inspector
                </h3>
                <p className="text-xs text-slate-500">
                  Source confidential details vs exact sanitized destination payload
                </p>
              </div>
            </div>
            <span className="rounded-full bg-emerald-50 px-2 py-0.5 text-[10px] font-semibold text-emerald-700 ring-1 ring-emerald-200">
              PII STRIPPED
            </span>
          </div>

          {inspectedPair ? (
            <div className="grid grid-cols-1 gap-3 p-4 md:grid-cols-2">
              <div className="rounded-lg border border-slate-200 bg-slate-50 p-3 text-xs">
                <div className="mb-1.5 flex items-center justify-between text-[11px] font-semibold text-slate-600">
                  <span>ORIGIN EVENT (NEVER LEAKED)</span>
                  <ArrowRight className="h-3.5 w-3.5 text-slate-400" />
                </div>
                <pre className="max-h-64 overflow-auto font-mono text-[10px] leading-relaxed text-slate-800">
                  {JSON.stringify(
                    {
                      title: inspectedPair.sourceEvent.title,
                      description: inspectedPair.sourceEvent.description,
                      location: inspectedPair.sourceEvent.location,
                      attendees: inspectedPair.sourceEvent.attendees,
                      showAs: inspectedPair.sourceEvent.showAs,
                      startIso: inspectedPair.sourceEvent.startIso,
                      endIso: inspectedPair.sourceEvent.endIso,
                    },
                    null,
                    2
                  )}
                </pre>
              </div>

              <div className="rounded-lg border border-slate-800 bg-slate-900 p-3 text-xs text-white">
                <div className="mb-1.5 flex items-center justify-between text-[11px] font-semibold text-emerald-400">
                  <span>
                    WIRE PAYLOAD →{' '}
                    {inspectedPair.targetAccount?.provider === 'outlook'
                      ? 'MICROSOFT GRAPH'
                      : 'GOOGLE CALENDAR API'}
                  </span>
                </div>
                <pre className="max-h-64 overflow-auto font-mono text-[10px] leading-relaxed text-emerald-300">
                  {JSON.stringify(
                    inspectedPair.samplePayload ?? {
                      summary: inspectedPair.mirrorEvent?.title ?? 'Busy',
                      description: '',
                      location: '',
                      attendees: [],
                      visibility: 'private',
                      locked: true,
                    },
                    null,
                    2
                  )}
                </pre>
              </div>
            </div>
          ) : (
            <div className="p-6 text-center text-xs text-slate-500">
              Select Inspect or JSON on any calendar card to audit exact API payloads.
            </div>
          )}
        </section>

        <WebhookTelemetry
          accounts={accounts}
          logs={logs}
          onRenewWatches={() => void onRenewWatches()}
          busyAction={busyAction}
        />
      </div>
    </div>
  );
}
