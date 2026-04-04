function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      'content-type': 'application/json; charset=utf-8',
      'cache-control': 'no-store'
    }
  });
}

function toBase64Url(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

async function signPayload(secret, payload) {
  const encoder = new TextEncoder();
  const key = await crypto.subtle.importKey(
    'raw',
    encoder.encode(secret),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign']
  );

  const signature = await crypto.subtle.sign('HMAC', key, encoder.encode(payload));
  return toBase64Url(signature);
}

function parseBasicAuth(headerValue) {
  if (!headerValue || !headerValue.startsWith('Basic ')) return null;
  try {
    const decoded = atob(headerValue.slice(6));
    const separatorIndex = decoded.indexOf(':');
    if (separatorIndex === -1) return null;
    return {
      email: decoded.slice(0, separatorIndex).trim().toLowerCase(),
      password: decoded.slice(separatorIndex + 1)
    };
  } catch {
    return null;
  }
}

function getAllowedUsers(env) {
  try {
    return JSON.parse(env.CHAT_USERS_JSON || '{}');
  } catch {
    return {};
  }
}

async function forwardToAppsScript(request, env, action, userEmail) {
  const gasBaseUrl = env.GAS_WEBAPP_URL;
  const sharedSecret = env.CHAT_SHARED_SECRET;

  if (!gasBaseUrl) {
    return json({ ok: false, error: 'Missing GAS_WEBAPP_URL' }, 500);
  }

  if (!sharedSecret) {
    return json({ ok: false, error: 'Missing CHAT_SHARED_SECRET' }, 500);
  }

  const ts = String(Math.floor(Date.now() / 1000));
  const sig = await signPayload(sharedSecret, `${userEmail}|${ts}`);

  const gasUrl = new URL(gasBaseUrl);
  gasUrl.searchParams.set('action', action);

  if (request.method === 'GET') {
    const incomingUrl = new URL(request.url);
    incomingUrl.searchParams.forEach((value, key) => {
      gasUrl.searchParams.set(key, value);
    });

    gasUrl.searchParams.set('user', userEmail);
    gasUrl.searchParams.set('ts', ts);
    gasUrl.searchParams.set('sig', sig);

    const gasResponse = await fetch(gasUrl.toString(), {
      method: 'GET',
      redirect: 'follow'
    });

    return new Response(gasResponse.body, {
      status: gasResponse.status,
      headers: {
        'content-type': 'application/json; charset=utf-8',
        'cache-control': 'no-store'
      }
    });
  }

  if (request.method === 'POST') {
    const body = await request.json().catch(() => ({}));
    body.user = userEmail;
    body.ts = ts;
    body.sig = sig;

    const gasResponse = await fetch(gasUrl.toString(), {
      method: 'POST',
      headers: {
        'content-type': 'application/json'
      },
      body: JSON.stringify(body),
      redirect: 'follow'
    });

    return new Response(gasResponse.body, {
      status: gasResponse.status,
      headers: {
        'content-type': 'application/json; charset=utf-8',
        'cache-control': 'no-store'
      }
    });
  }

  return json({ ok: false, error: 'Method not allowed' }, 405);
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname.startsWith('/api/')) {
      const action = url.pathname.replace(/^\/api\//, '');

      const credentials = parseBasicAuth(request.headers.get('Authorization'));
      if (!credentials) {
        return json({ ok: false, error: 'Unauthorized' }, 401);
      }

      const users = getAllowedUsers(env);
      if (!users[credentials.email] || users[credentials.email] !== credentials.password) {
        return json({ ok: false, error: 'Invalid credentials' }, 401);
      }

      return forwardToAppsScript(request, env, action, credentials.email);
    }

    return env.ASSETS.fetch(request);
  }
};
