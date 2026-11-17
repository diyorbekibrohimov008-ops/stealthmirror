# StealthMirror

Cross-account calendar firewall for Google Calendar and Microsoft Outlook / 365.

Instead of taking over your schedule with rescheduling clutter or charging per-pair fees with 15–30 minute delays, StealthMirror mirrors private, unmovable `"Busy"` blocks across unlimited connected calendars in sub-second webhooks — with zero title, description, location, or attendee leaks.

## Live Deployment

- **Production URL**: [https://stealthmirror-ibroximov.vercel.app](https://stealthmirror-ibroximov.vercel.app)
- **Deployment Platform**: Vercel (Fluid Compute Serverless Functions + Vite SPA + `@vercel/blob` persistent storage)
- **Health Endpoint**: [https://stealthmirror-ibroximov.vercel.app/api/health](https://stealthmirror-ibroximov.vercel.app/api/health)
- **Status**: 24/7 online, free tier, public access (SSO & password protection disabled)

## Core Features

- **Zero-Leak "Stealth Mirror" Sanitization** ([`shared/sanitizer.ts`](shared/sanitizer.ts)):
  - Strips `title`, `description`, `location`, `attendees`, meeting URLs, and collaborator metadata.
  - Emits private Google Calendar (`extendedProperties.private`) and Microsoft Graph (`singleValueExtendedProperties`) Busy payloads locked against accidental edits (`403 LOCKED_MIRROR_BLOCK`).
- **Loop-Safe Sync Engine** ([`shared/syncEngine.ts`](shared/syncEngine.ts)):
  - Prevents infinite echo cascades across N-way mesh calendars via `stealthMirrorManaged` tags, known mirror-mapping lookup, and deterministic timing fingerprints.
  - Honors rule policies (`busyTitle`, `lockMirrored`, `skipFreeEvents`, `skipDeclinedEvents`, `skipAllDayEvents`).
- **Live Split Calendar Sandbox + Real OAuth & Webhooks**:
  - Works immediately with 3 seeded calendars (Google Work, Outlook Client, Personal Gmail) backed by Netlify Blobs (`strong` consistency) and automatic in-memory fallback.
  - Production Google Calendar (`watch` + `syncToken`) and Microsoft Graph (`subscription` + `deltaLink`) handlers in [`netlify/lib/router.ts`](netlify/lib/router.ts) plus 12-hour push channel renewal cron in [`netlify/functions/renew-watches.mts`](netlify/functions/renew-watches.mts).
- **Flat Pricing & 1-Click Legacy Clutter Cleaner**:
  - Purge legacy `[Reclaim.ai]` and `Synced via SyncGene` clutter events cleanly without touching real meetings.
  - Built around `$5/month flat` or `$39 lifetime` for unlimited calendar connections.

## Local Development

```bash
npm install
npm run dev
```

Vite runs with `@netlify/vite-plugin` so both the React frontend and `/api/*` Netlify Functions (plus local Netlify Blobs emulation) run together at `http://localhost:5173`.

## Verification

```bash
npm run typecheck
npm test
npm run build
```

## Production OAuth & Webhook Environment Variables (Optional)

Sandbox mode requires zero environment variables. To enable live Google Calendar & Microsoft Outlook 365 OAuth connections on your deployed site, set:

- `APP_URL` — Public site base URL (e.g. `https://your-site.netlify.app`)
- `GOOGLE_CLIENT_ID` & `GOOGLE_CLIENT_SECRET`
  - Callback URI: `${APP_URL}/api/oauth/google/callback`
  - Webhook URI: `${APP_URL}/api/webhooks/google`
- `MICROSOFT_CLIENT_ID` & `MICROSOFT_CLIENT_SECRET`
  - Callback URI: `${APP_URL}/api/oauth/microsoft/callback`
  - Webhook URI: `${APP_URL}/api/webhooks/microsoft`
