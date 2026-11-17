import {
  Activity,
  GitBranch,
  Layers,
  Plus,
  RotateCcw,
  Shield,
  Sparkles,
} from 'lucide-react';
import type { Workspace } from '../../shared/types';
import type { WorkspaceStats } from '../lib/api';

export type ActiveTab = 'workspace' | 'matrix' | 'pricing';

interface NavbarProps {
  workspace: Workspace;
  stats: WorkspaceStats;
  activeTab: ActiveTab;
  onSelectTab: (tab: ActiveTab) => void;
  workspaceIds: string[];
  onSwitchWorkspace: (workspaceId: string) => void;
  onCreateWorkspace: () => void;
  onResetDemo: () => void;
  busyAction: string | null;
}

export function Navbar({
  workspace,
  stats,
  activeTab,
  onSelectTab,
  workspaceIds,
  onSwitchWorkspace,
  onCreateWorkspace,
  onResetDemo,
  busyAction,
}: NavbarProps) {
  const planBadge =
    workspace.plan === 'lifetime_39'
      ? '$39 Lifetime · ∞ Calendars'
      : workspace.plan === 'monthly_5'
      ? '$5/mo Flat · ∞ Calendars'
      : '14-Day Trial';

  const navItems: Array<{ id: ActiveTab; label: string; icon: typeof Layers; badge?: number }> = [
    { id: 'workspace', label: 'Live Split Calendar', icon: Layers },
    { id: 'matrix', label: 'Sync Matrix & OAuth', icon: GitBranch, badge: stats.activeRuleCount },
    {
      id: 'pricing',
      label: 'Migrate & Pricing',
      icon: Sparkles,
      badge: stats.legacyClutterCount > 0 ? stats.legacyClutterCount : undefined,
    },
  ];

  return (
    <header className="sticky top-0 z-30 border-b border-slate-200 bg-white/95 backdrop-blur-md">
      <div className="mx-auto flex max-w-7xl flex-wrap items-center justify-between gap-3 px-4 py-3 sm:px-6">
        <div className="flex items-center gap-3">
          <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-slate-900 text-white shadow-sm">
            <Shield className="h-5 w-5 text-emerald-400" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="text-sm font-bold tracking-tight text-slate-900 sm:text-base">
                StealthMirror
              </span>
              <span className="inline-flex items-center gap-1 rounded-full bg-emerald-50 px-2 py-0.5 text-[11px] font-semibold text-emerald-700 ring-1 ring-emerald-200">
                <Activity className="h-3 w-3" />
                {stats.avgLatencyMs}ms avg
              </span>
              <span className="hidden rounded-full bg-slate-100 px-2 py-0.5 text-[11px] font-medium text-slate-700 sm:inline-flex">
                0 leaks
              </span>
            </div>
            <p className="text-xs text-slate-500">
              Private "Busy" blocks across Google &amp; Outlook — zero takeover bloat
            </p>
          </div>
        </div>

        <nav className="flex items-center gap-1 rounded-lg bg-slate-100 p-1">
          {navItems.map((item) => {
            const Icon = item.icon;
            const active = activeTab === item.id;
            return (
              <button
                key={item.id}
                type="button"
                onClick={() => onSelectTab(item.id)}
                className={`inline-flex items-center gap-1.5 rounded-md px-3 py-1.5 text-xs font-medium transition ${
                  active
                    ? 'bg-white text-slate-900 shadow-xs ring-1 ring-slate-200/80'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Icon className="h-3.5 w-3.5" />
                <span>{item.label}</span>
                {item.badge !== undefined && (
                  <span
                    className={`rounded-full px-1.5 py-0.2 text-[10px] font-semibold ${
                      item.id === 'pricing'
                        ? 'bg-amber-100 text-amber-800'
                        : 'bg-slate-200/80 text-slate-700'
                    }`}
                  >
                    {item.badge}
                  </span>
                )}
              </button>
            );
          })}
        </nav>

        <div className="flex items-center gap-2">
          <div className="hidden items-center gap-1.5 rounded-md border border-slate-200 bg-slate-50 px-2 py-1 text-xs md:flex">
            <span className="text-slate-400">WS:</span>
            <select
              aria-label="Workspace selector"
              value={workspace.id}
              onChange={(e) => onSwitchWorkspace(e.target.value)}
              className="bg-transparent font-medium text-slate-800 outline-none"
            >
              {workspaceIds.map((id) => (
                <option key={id} value={id}>
                  {id === 'demo' ? 'demo (3 seeded accounts)' : id}
                </option>
              ))}
            </select>
            <button
              type="button"
              onClick={onCreateWorkspace}
              title="Create blank workspace"
              className="rounded p-0.5 text-slate-500 hover:bg-slate-200 hover:text-slate-900"
            >
              <Plus className="h-3.5 w-3.5" />
            </button>
          </div>

          <button
            type="button"
            onClick={() => onSelectTab('pricing')}
            className="hidden rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 xl:inline-flex"
          >
            {planBadge}
          </button>

          {workspace.id === 'demo' && (
            <button
              type="button"
              onClick={onResetDemo}
              disabled={busyAction !== null}
              className="inline-flex items-center gap-1.5 rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 disabled:opacity-50"
            >
              <RotateCcw className="h-3.5 w-3.5" />
              <span>Reset Demo</span>
            </button>
          )}
        </div>
      </div>
    </header>
  );
}
