import type { Config } from '@netlify/functions';
import { renewExpiringWatches } from '../lib/router';
import { getStoreClient } from '../lib/store';

export default async (req: Request) => {
  const payload = (await req.json().catch(() => ({}))) as { next_run?: string };
  const result = await renewExpiringWatches(getStoreClient());
  console.log('Watch renewal run', { nextRun: payload.next_run, ...result });
};

export const config: Config = {
  // Runs every 12 hours UTC to keep Google Calendar watch channels & MS Graph subscriptions alive
  schedule: '0 */12 * * *',
};
