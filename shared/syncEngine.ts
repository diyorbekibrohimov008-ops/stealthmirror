import { computeEventFingerprint, isStealthManagedEvent } from './sanitizer';
import type {
  CalendarAccount,
  CalendarEvent,
  MirrorMapping,
  SyncLogEntry,
  SyncRule,
} from './types';

export interface SyncStoragePort {
  getRulesForSource(workspaceId: string, sourceAccountId: string): Promise<SyncRule[]>;
  getAccount(accountId: string): Promise<CalendarAccount | null>;
  getMapping(
    workspaceId: string,
    ruleId: string,
    sourceExternalEventId: string
  ): Promise<MirrorMapping | null>;
  isKnownMirroredTarget(
    workspaceId: string,
    targetAccountId: string,
    externalEventId: string
  ): Promise<boolean>;
  saveMapping(mapping: MirrorMapping): Promise<void>;
  deleteMapping(mappingId: string): Promise<void>;
  upsertDestinationBusyEvent(
    targetAccount: CalendarAccount,
    rule: SyncRule,
    sourceEvent: CalendarEvent,
    existingTargetExternalId?: string
  ): Promise<{ targetExternalEventId: string; payloadPreview: Record<string, unknown> }>;
  deleteDestinationBusyEvent(
    targetAccount: CalendarAccount,
    targetExternalEventId: string
  ): Promise<void>;
  appendLog(entry: SyncLogEntry): Promise<void>;
}

export interface SyncSummary {
  created: number;
  updated: number;
  deleted: number;
  skippedLoop: number;
  skippedPolicy: number;
  errors: number;
}

function makeId(prefix: string): string {
  return `${prefix}_${crypto.randomUUID().slice(0, 12)}`;
}

function getSkipReason(event: CalendarEvent, rule: SyncRule): string | null {
  if (event.deleted) return 'Source event deleted';
  if (rule.skipFreeEvents && event.showAs === 'free') {
    return 'Source event marked as Free (transparent)';
  }
  if (rule.skipDeclinedEvents && event.responseStatus === 'declined') {
    return 'Attendee declined source event';
  }
  if (rule.skipAllDayEvents && event.allDay) {
    return 'All-day event skipped by rule policy';
  }
  return null;
}

export class StealthSyncEngine {
  constructor(private readonly port: SyncStoragePort) {}

  async processEventChange(
    workspaceId: string,
    sourceEvent: CalendarEvent
  ): Promise<SyncSummary> {
    const startedAt = performance.now();
    const summary: SyncSummary = {
      created: 0,
      updated: 0,
      deleted: 0,
      skippedLoop: 0,
      skippedPolicy: 0,
      errors: 0,
    };

    const knownTarget = await this.port.isKnownMirroredTarget(
      workspaceId,
      sourceEvent.accountId,
      sourceEvent.externalEventId
    );

    if (isStealthManagedEvent(sourceEvent) || knownTarget) {
      summary.skippedLoop += 1;
      await this.port.appendLog({
        id: makeId('log'),
        workspaceId,
        sourceAccountId: sourceEvent.accountId,
        sourceEventId: sourceEvent.externalEventId,
        action: 'skipped_loop',
        reason: 'Loop guard triggered: stealthMirrorManaged=true',
        latencyMs: Math.max(1, Math.round(performance.now() - startedAt)),
        createdAt: new Date().toISOString(),
      });
      return summary;
    }

    const rules = await this.port.getRulesForSource(workspaceId, sourceEvent.accountId);

    for (const rule of rules) {
      if (rule.targetAccountId === sourceEvent.accountId) continue;
      const targetAccount = await this.port.getAccount(rule.targetAccountId);
      if (!targetAccount) continue;

      const ruleStarted = performance.now();
      const existingMapping = await this.port.getMapping(
        workspaceId,
        rule.id,
        sourceEvent.externalEventId
      );

      const skipReason = getSkipReason(sourceEvent, rule);

      if (skipReason) {
        if (existingMapping) {
          await this.port.deleteDestinationBusyEvent(
            targetAccount,
            existingMapping.targetExternalEventId
          );
          await this.port.deleteMapping(existingMapping.id);
          summary.deleted += 1;
          await this.port.appendLog({
            id: makeId('log'),
            workspaceId,
            ruleId: rule.id,
            sourceAccountId: sourceEvent.accountId,
            targetAccountId: targetAccount.id,
            sourceEventId: sourceEvent.externalEventId,
            action: 'deleted',
            reason: `Removed mirrored Busy block (${skipReason})`,
            latencyMs: Math.max(1, Math.round(performance.now() - ruleStarted)),
            createdAt: new Date().toISOString(),
          });
        } else {
          summary.skippedPolicy += 1;
          await this.port.appendLog({
            id: makeId('log'),
            workspaceId,
            ruleId: rule.id,
            sourceAccountId: sourceEvent.accountId,
            targetAccountId: targetAccount.id,
            sourceEventId: sourceEvent.externalEventId,
            action: 'skipped_policy',
            reason: skipReason,
            latencyMs: Math.max(1, Math.round(performance.now() - ruleStarted)),
            createdAt: new Date().toISOString(),
          });
        }
        continue;
      }

      const fingerprint = computeEventFingerprint(sourceEvent, rule);
      if (existingMapping && existingMapping.fingerprint === fingerprint) {
        continue;
      }

      try {
        const upsertResult = await this.port.upsertDestinationBusyEvent(
          targetAccount,
          rule,
          sourceEvent,
          existingMapping?.targetExternalEventId
        );

        const nextMapping: MirrorMapping = {
          id: existingMapping?.id ?? makeId('map'),
          workspaceId,
          ruleId: rule.id,
          sourceAccountId: sourceEvent.accountId,
          sourceExternalEventId: sourceEvent.externalEventId,
          targetAccountId: targetAccount.id,
          targetExternalEventId: upsertResult.targetExternalEventId,
          fingerprint,
          updatedAt: new Date().toISOString(),
        };
        await this.port.saveMapping(nextMapping);

        const action = existingMapping ? 'updated' : 'created';
        if (action === 'updated') summary.updated += 1;
        else summary.created += 1;

        await this.port.appendLog({
          id: makeId('log'),
          workspaceId,
          ruleId: rule.id,
          sourceAccountId: sourceEvent.accountId,
          targetAccountId: targetAccount.id,
          sourceEventId: sourceEvent.externalEventId,
          action,
          reason:
            action === 'created'
              ? `Mirrored private "${rule.busyTitle || 'Busy'}" block to ${targetAccount.label}`
              : `Updated mirrored Busy block timing on ${targetAccount.label}`,
          latencyMs: Math.max(1, Math.round(performance.now() - ruleStarted)),
          payloadPreview: upsertResult.payloadPreview,
          createdAt: new Date().toISOString(),
        });
      } catch (err) {
        summary.errors += 1;
        await this.port.appendLog({
          id: makeId('log'),
          workspaceId,
          ruleId: rule.id,
          sourceAccountId: sourceEvent.accountId,
          targetAccountId: targetAccount.id,
          sourceEventId: sourceEvent.externalEventId,
          action: 'error',
          reason: err instanceof Error ? err.message : 'Unknown sync error',
          latencyMs: Math.max(1, Math.round(performance.now() - ruleStarted)),
          createdAt: new Date().toISOString(),
        });
      }
    }

    return summary;
  }
}
