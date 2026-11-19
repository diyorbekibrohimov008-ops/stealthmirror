import { useState } from 'react';
import {
  ArrowRight,
  CheckCircle2,
  Copy,
  ExternalLink,
  GitBranch,
  KeyRound,
  Lock,
  Plus,
  Trash2,
} from 'lucide-react';
import type { CalendarAccount, CalendarProvider, SyncRule } from '../../shared/types';
import type { OAuthStatus } from '../lib/api';

interface SyncMatrixPanelProps {
  workspaceId: string;
  accounts: CalendarAccount[];
  rules: SyncRule[];
  oauth: OAuthStatus;
  busyAction: string | null;
  onCreateAccount: (input: {
    provider: CalendarProvider;
    label: string;
    email: string;
    calendarName: string;
    color: string;
    autoMeshWithExisting: boolean;
  }) => Promise<void>;
  onDeleteAccount: (accountId: string) => Promise<void>;
  onCreateRule: (input: {
    sourceAccountId: string;
    targetAccountId: string;
    busyTitle: string;
    lockMirrored: boolean;
    skipFreeEvents: boolean;
    skipDeclinedEvents: boolean;
    skipAllDayEvents: boolean;
  }) => Promise<void>;
  onUpdateRule: (
    ruleId: string,
    patch: {
      enabled?: boolean;
      busyTitle?: string;
      lockMirrored?: boolean;
      skipFreeEvents?: boolean;
      skipDeclinedEvents?: boolean;
      skipAllDayEvents?: boolean;
    }
  ) => Promise<void>;
  onDeleteRule: (ruleId: string) => Promise<void>;
}

const COLOR_SWATCHES = ['#0EA5E9', '#0F766E', '#10B981', '#F59E0B', '#EC4899', '#0891B2'];
const SAFE_BUSY_TITLES = ['Busy', 'Busy — Blocked', 'Unavailable', 'Private Block'];

export function SyncMatrixPanel({
  workspaceId,
  accounts,
  rules,
  oauth,
  busyAction,
  onCreateAccount,
  onDeleteAccount,
  onCreateRule,
  onUpdateRule,
  onDeleteRule,
}: SyncMatrixPanelProps) {
  const [provider, setProvider] = useState<CalendarProvider>('google');
  const [label, setLabel] = useState('Advisory Board Gmail');
  const [email, setEmail] = useState('alex.advisory@gmail.com');
  const [color, setColor] = useState('#8B5CF6');
  const [autoMesh, setAutoMesh] = useState(true);

  const [sourceAccountId, setSourceAccountId] = useState(accounts[0]?.id ?? '');
  const [targetAccountId, setTargetAccountId] = useState(accounts[1]?.id ?? '');
  const [newRuleBusyTitle, setNewRuleBusyTitle] = useState('Busy');
  const [copiedUrl, setCopiedUrl] = useState<string | null>(null);

  const resolvedSourceId = accounts.some((a) => a.id === sourceAccountId)
    ? sourceAccountId
    : accounts[0]?.id ?? '';
  const resolvedTargetId = accounts.some(
    (a) => a.id === targetAccountId && a.id !== resolvedSourceId
  )
    ? targetAccountId
    : accounts.find((a) => a.id !== resolvedSourceId)?.id ?? '';

  const accountMap = new Map(accounts.map((a) => [a.id, a]));

  async function handleAddSandboxAccount(e: React.FormEvent) {
    e.preventDefault();
    if (!label.trim() || !email.trim()) return;
    await onCreateAccount({
      provider,
      label: label.trim(),
      email: email.trim(),
      calendarName: 'Primary',
      color,
      autoMeshWithExisting: autoMesh,
    });
  }

  async function handleCopy(text: string) {
    try {
      await navigator.clipboard.writeText(text);
      setCopiedUrl(text);
      setTimeout(() => setCopiedUrl(null), 1500);
    } catch {
      // Clipboard API unavailable
    }
  }

  function handleOAuthClick(oauthProvider: CalendarProvider) {
    const configured =
      oauthProvider === 'google' ? oauth.googleConfigured : oauth.microsoftConfigured;
    if (!configured) {
      return;
    }
    const url = `/api/oauth/${
      oauthProvider === 'google' ? 'google' : 'microsoft'
    }/start?workspaceId=${encodeURIComponent(workspaceId)}&label=${encodeURIComponent(
      label.trim() || (oauthProvider === 'google' ? 'Google Calendar' : 'Outlook 365')
    )}`;
    window.location.href = url;
  }

  return (
    <div className="space-y-6">
      {/* Connected Accounts + Instant Sandbox & Real OAuth */}
      <div className="grid grid-cols-1 gap-5 lg:grid-cols-5">
        <section className="rounded-xl border border-slate-200 bg-white p-4 shadow-xs lg:col-span-3">
          <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 pb-3">
            <div>
              <h2 className="text-sm font-semibold text-slate-900">
                Connected Calendars ({accounts.length} Connected · Unlimited Included)
              </h2>
              <p className="text-xs text-slate-500">
                Flat $5/mo or $39 lifetime — connect 2 or 20 accounts with $0 per-pair fees.
              </p>
            </div>
            <span className="rounded-full bg-emerald-50 px-2.5 py-0.5 text-[11px] font-semibold text-emerald-700 ring-1 ring-emerald-200">
              $0 / extra account
            </span>
          </div>

          <div className="mt-3 space-y-2.5">
            {accounts.length === 0 ? (
              <p className="rounded-lg border border-dashed border-slate-200 p-5 text-center text-xs text-slate-500">
                No accounts connected yet. Use the form on the right to add an instant sandbox
                calendar or real OAuth connection.
              </p>
            ) : (
              accounts.map((acc) => (
                <div
                  key={acc.id}
                  className="flex flex-wrap items-center justify-between gap-3 rounded-lg border border-slate-200 p-3 text-xs"
                >
                  <div className="flex items-center gap-3">
                    <span
                      className="h-3 w-3 shrink-0 rounded-full"
                      style={{ backgroundColor: acc.color }}
                    />
                    <div>
                      <div className="flex items-center gap-2">
                        <span className="font-semibold text-slate-900">{acc.label}</span>
                        <span className="rounded bg-slate-100 px-1.5 py-0.5 text-[10px] font-medium uppercase text-slate-600">
                          {acc.provider}
                        </span>
                        <span className="inline-flex items-center gap-1 rounded bg-emerald-50 px-1.5 py-0.5 text-[10px] font-medium text-emerald-700">
                          <CheckCircle2 className="h-3 w-3" />
                          {acc.mode === 'oauth' ? 'OAuth Active' : 'Sandbox Live'}
                        </span>
                      </div>
                      <p className="mt-0.5 text-[11px] text-slate-500">
                        {acc.email} · Watch channel:{' '}
                        <span className="font-mono">{acc.watchChannelId ?? 'pending'}</span>
                      </p>
                    </div>
                  </div>

                  <button
                    type="button"
                    onClick={() => void onDeleteAccount(acc.id)}
                    disabled={busyAction !== null}
                    className="inline-flex items-center gap-1 rounded border border-rose-200 bg-rose-50 px-2.5 py-1 text-[11px] font-medium text-rose-700 hover:bg-rose-100 disabled:opacity-50"
                  >
                    <Trash2 className="h-3 w-3" />
                    <span>Disconnect</span>
                  </button>
                </div>
              ))
            )}
          </div>
        </section>

        {/* Add Account Card */}
        <section className="rounded-xl border border-slate-200 bg-white p-4 shadow-xs lg:col-span-2">
          <h3 className="text-sm font-semibold text-slate-900">
            Add Calendar Account
          </h3>
          <p className="mt-0.5 text-xs text-slate-500">
            Test instantly in Sandbox or authorize live Google / Microsoft Graph OAuth.
          </p>

          <form onSubmit={(e) => void handleAddSandboxAccount(e)} className="mt-3 space-y-3">
            <div className="grid grid-cols-2 gap-2">
              <button
                type="button"
                onClick={() => setProvider('google')}
                className={`rounded-lg border px-3 py-2 text-xs font-semibold transition ${
                  provider === 'google'
                    ? 'border-slate-900 bg-slate-900 text-white'
                    : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-50'
                }`}
              >
                Google Calendar
              </button>
              <button
                type="button"
                onClick={() => setProvider('outlook')}
                className={`rounded-lg border px-3 py-2 text-xs font-semibold transition ${
                  provider === 'outlook'
                    ? 'border-slate-900 bg-slate-900 text-white'
                    : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-50'
                }`}
              >
                Outlook / M365
              </button>
            </div>

            <label className="block text-xs">
              <span className="mb-1 block font-medium text-slate-700">Account Label</span>
              <input
                type="text"
                required
                value={label}
                onChange={(e) => setLabel(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-2.5 py-2 text-xs focus:border-slate-900 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-500/40"
              />
            </label>

            <label className="block text-xs">
              <span className="mb-1 block font-medium text-slate-700">Email Address</span>
              <input
                type="email"
                required
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-2.5 py-2 text-xs focus:border-slate-900 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-500/40"
              />
            </label>

            <div className="text-xs">
              <span className="mb-1.5 block font-medium text-slate-700">Accent Color</span>
              <div className="flex items-center gap-2">
                {COLOR_SWATCHES.map((swatch) => (
                  <button
                    key={swatch}
                    type="button"
                    aria-label={`Color ${swatch}`}
                    onClick={() => setColor(swatch)}
                    className={`h-6 w-6 rounded-full transition ${
                      color === swatch ? 'ring-2 ring-slate-900 ring-offset-2' : ''
                    }`}
                    style={{ backgroundColor: swatch }}
                  />
                ))}
              </div>
            </div>

            <label className="inline-flex items-center gap-2 text-xs text-slate-700">
              <input
                type="checkbox"
                checked={autoMesh}
                onChange={(e) => setAutoMesh(e.target.checked)}
                className="h-4 w-4 rounded border-slate-300"
              />
              <span>Auto-link bidirectional "Busy" rules with primary calendar</span>
            </label>

            <button
              type="submit"
              disabled={busyAction !== null}
              className="flex w-full items-center justify-center gap-1.5 rounded-lg bg-slate-900 px-3 py-2 text-xs font-semibold text-white hover:bg-slate-800 disabled:opacity-50"
            >
              <Plus className="h-3.5 w-3.5 text-emerald-400" />
              <span>Add Sandbox Account Instantly</span>
            </button>
          </form>

          {/* Real OAuth Section */}
          <div className="mt-4 border-t border-slate-100 pt-3">
            <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-800">
              <KeyRound className="h-3.5 w-3.5 text-emerald-600" />
              <span>Production OAuth 2.0 Credentials</span>
            </div>

            <div className="mt-2 grid grid-cols-1 gap-2 sm:grid-cols-2">
              <button
                type="button"
                onClick={() => handleOAuthClick('google')}
                disabled={!oauth.googleConfigured}
                className="inline-flex items-center justify-center gap-1.5 rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-2 text-xs font-medium text-slate-800 hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-60"
              >
                <ExternalLink className="h-3.5 w-3.5" />
                <span>
                  {oauth.googleConfigured ? 'Connect Real Google' : 'Google OAuth (Env needed)'}
                </span>
              </button>

              <button
                type="button"
                onClick={() => handleOAuthClick('outlook')}
                disabled={!oauth.microsoftConfigured}
                className="inline-flex items-center justify-center gap-1.5 rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-2 text-xs font-medium text-slate-800 hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-60"
              >
                <ExternalLink className="h-3.5 w-3.5" />
                <span>
                  {oauth.microsoftConfigured
                    ? 'Connect Real Outlook'
                    : 'Microsoft OAuth (Env needed)'}
                </span>
              </button>
            </div>

            <div className="mt-2.5 space-y-1.5 rounded-lg bg-slate-50 p-2.5 text-[11px] text-slate-600">
              <p className="font-medium text-slate-700">
                Production OAuth Redirect &amp; Webhook URLs:
              </p>
              {[
                { label: 'Google Callback', value: oauth.googleCallbackUrl },
                { label: 'Microsoft Callback', value: oauth.microsoftCallbackUrl },
              ].map((row) => (
                <div key={row.label} className="flex items-center justify-between gap-2">
                  <span className="truncate font-mono text-[10px]">
                    {row.label}: {row.value}
                  </span>
                  <button
                    type="button"
                    onClick={() => void handleCopy(row.value)}
                    className="shrink-0 rounded border border-slate-200 bg-white px-1.5 py-0.5 text-[10px] hover:bg-slate-100"
                  >
                    {copiedUrl === row.value ? 'Copied' : <Copy className="h-3 w-3" />}
                  </button>
                </div>
              ))}
            </div>
          </div>
        </section>
      </div>

      {/* Directed Sync Rule Matrix */}
      <section className="rounded-xl border border-slate-200 bg-white p-4 shadow-xs">
        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 pb-3">
          <div className="flex items-center gap-2">
            <GitBranch className="h-4 w-4 text-slate-700" />
            <div>
              <h2 className="text-sm font-semibold text-slate-900">
                Directional Sync Rules Matrix
              </h2>
              <p className="text-xs text-slate-500">
                Configure exactly how each calendar mirrors clean, locked "Busy" blocks without
                takeover clutter.
              </p>
            </div>
          </div>

          {accounts.length >= 2 && (
            <div className="flex flex-wrap items-center gap-2">
              <select
                aria-label="Source account"
                value={resolvedSourceId}
                onChange={(e) => setSourceAccountId(e.target.value)}
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs"
              >
                {accounts.map((acc) => (
                  <option key={acc.id} value={acc.id}>
                    From: {acc.label}
                  </option>
                ))}
              </select>

              <ArrowRight className="h-3.5 w-3.5 text-slate-400" />

              <select
                aria-label="Target account"
                value={resolvedTargetId}
                onChange={(e) => setTargetAccountId(e.target.value)}
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs"
              >
                {accounts
                  .filter((acc) => acc.id !== resolvedSourceId)
                  .map((acc) => (
                    <option key={acc.id} value={acc.id}>
                      To: {acc.label}
                    </option>
                  ))}
              </select>

              <select
                aria-label="Mirrored title label"
                value={newRuleBusyTitle}
                onChange={(e) => setNewRuleBusyTitle(e.target.value)}
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs"
              >
                {SAFE_BUSY_TITLES.map((lbl) => (
                  <option key={lbl} value={lbl}>
                    Label: {lbl}
                  </option>
                ))}
              </select>

              <button
                type="button"
                disabled={
                  busyAction !== null ||
                  !resolvedSourceId ||
                  !resolvedTargetId ||
                  resolvedSourceId === resolvedTargetId
                }
                onClick={() =>
                  void onCreateRule({
                    sourceAccountId: resolvedSourceId,
                    targetAccountId: resolvedTargetId,
                    busyTitle: newRuleBusyTitle,
                    lockMirrored: true,
                    skipFreeEvents: true,
                    skipDeclinedEvents: true,
                    skipAllDayEvents: true,
                  })
                }
                className="inline-flex items-center gap-1 rounded-lg bg-slate-900 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-800 disabled:opacity-40"
              >
                <Plus className="h-3.5 w-3.5" />
                <span>Add Sync Route</span>
              </button>
            </div>
          )}
        </div>

        <div className="mt-3 overflow-x-auto">
          <table className="w-full min-w-[740px] border-collapse text-left text-xs">
            <thead>
              <tr className="border-b border-slate-200 text-[11px] uppercase tracking-wide text-slate-500">
                <th className="py-2 pr-3 font-medium">Route (Source &rarr; Target)</th>
                <th className="px-3 py-2 font-medium">Status</th>
                <th className="px-3 py-2 font-medium">Masked Title</th>
                <th className="px-3 py-2 font-medium">Lock Block</th>
                <th className="px-3 py-2 font-medium">Skip Free</th>
                <th className="px-3 py-2 font-medium">Skip Declined</th>
                <th className="px-3 py-2 font-medium">Skip All-Day</th>
                <th className="py-2 pl-3 text-right font-medium">Action</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {rules.map((rule) => {
                const src = accountMap.get(rule.sourceAccountId);
                const tgt = accountMap.get(rule.targetAccountId);
                return (
                  <tr key={rule.id} className="hover:bg-slate-50/70">
                    <td className="py-2.5 pr-3 font-medium text-slate-900">
                      <span className="inline-flex items-center gap-1.5">
                        <span>{src?.label ?? rule.sourceAccountId}</span>
                        <ArrowRight className="h-3 w-3 text-slate-400" />
                        <span>{tgt?.label ?? rule.targetAccountId}</span>
                      </span>
                    </td>
                    <td className="px-3 py-2.5">
                      <button
                        type="button"
                        onClick={() =>
                          void onUpdateRule(rule.id, { enabled: !rule.enabled })
                        }
                        disabled={busyAction !== null}
                        className={`rounded-full px-2 py-0.5 text-[10px] font-semibold ring-1 ${
                          rule.enabled
                            ? 'bg-emerald-50 text-emerald-700 ring-emerald-200'
                            : 'bg-slate-100 text-slate-500 ring-slate-200'
                        }`}
                      >
                        {rule.enabled ? 'ACTIVE' : 'PAUSED'}
                      </button>
                    </td>
                    <td className="px-3 py-2.5">
                      <select
                        aria-label="Rule busy title"
                        value={rule.busyTitle}
                        onChange={(e) =>
                          void onUpdateRule(rule.id, { busyTitle: e.target.value })
                        }
                        className="rounded border border-slate-200 bg-white px-2 py-1 text-xs"
                      >
                        {SAFE_BUSY_TITLES.map((lbl) => (
                          <option key={lbl} value={lbl}>
                            {lbl}
                          </option>
                        ))}
                      </select>
                    </td>
                    {(
                      [
                        { key: 'lockMirrored', val: rule.lockMirrored },
                        { key: 'skipFreeEvents', val: rule.skipFreeEvents },
                        { key: 'skipDeclinedEvents', val: rule.skipDeclinedEvents },
                        { key: 'skipAllDayEvents', val: rule.skipAllDayEvents },
                      ] as const
                    ).map((toggle) => (
                      <td key={toggle.key} className="px-3 py-2.5">
                        <input
                          type="checkbox"
                          checked={toggle.val}
                          onChange={(e) =>
                            void onUpdateRule(rule.id, {
                              [toggle.key]: e.target.checked,
                            })
                          }
                          className="h-4 w-4 rounded border-slate-300 text-slate-900"
                        />
                      </td>
                    ))}
                    <td className="py-2.5 pl-3 text-right">
                      <button
                        type="button"
                        onClick={() => void onDeleteRule(rule.id)}
                        disabled={busyAction !== null}
                        className="inline-flex items-center gap-1 rounded border border-slate-200 px-2 py-1 text-[11px] text-slate-600 hover:bg-rose-50 hover:text-rose-700"
                      >
                        <Lock className="h-3 w-3" />
                        <span>Remove</span>
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}
