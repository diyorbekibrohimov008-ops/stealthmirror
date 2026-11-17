declare const Netlify:
  | {
      env?: {
        get: (key: string) => string | undefined;
      };
    }
  | undefined;

export function getEnv(key: string): string | undefined {
  try {
    if (typeof Netlify !== 'undefined' && Netlify?.env?.get) {
      const val = Netlify.env.get(key);
      if (val !== undefined) return val;
    }
  } catch {
    // Fallback to process.env outside Netlify runtime
  }
  return typeof process !== 'undefined' ? process.env[key] : undefined;
}

export function getOAuthConfigStatus(origin?: string) {
  const googleClientId = getEnv('GOOGLE_CLIENT_ID');
  const googleClientSecret = getEnv('GOOGLE_CLIENT_SECRET');
  const msClientId = getEnv('MICROSOFT_CLIENT_ID');
  const msClientSecret = getEnv('MICROSOFT_CLIENT_SECRET');
  const appUrl = getEnv('APP_URL') || origin || 'http://localhost:5173';

  return {
    googleConfigured: Boolean(googleClientId && googleClientSecret),
    microsoftConfigured: Boolean(msClientId && msClientSecret),
    appUrl,
    googleCallbackUrl: `${appUrl}/api/oauth/google/callback`,
    microsoftCallbackUrl: `${appUrl}/api/oauth/microsoft/callback`,
    googleWebhookUrl: `${appUrl}/api/webhooks/google`,
    microsoftWebhookUrl: `${appUrl}/api/webhooks/microsoft`,
  };
}
