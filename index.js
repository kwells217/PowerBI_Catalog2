const { app } = require('@azure/functions');

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': process.env.ALLOWED_ORIGIN || '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Content-Type': 'application/json'
};

// Store pending device code sessions in memory
const sessions = {};

async function pbiGet(token, path) {
  const res = await fetch(`https://api.powerbi.com/v1.0/myorg${path}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok) throw new Error(`Power BI API error ${res.status}: ${path}`);
  return res.json();
}

app.http('pbi-proxy', {
  methods: ['GET', 'POST', 'OPTIONS'],
  authLevel: 'function',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') {
      return { status: 204, headers: CORS_HEADERS };
    }

    const url = new URL(request.url);
    const action = url.searchParams.get('action');

    try {
      // Step 1: Start device code flow
      if (action === 'startLogin') {
        const res = await fetch(
          `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/devicecode`,
          {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
              client_id: CLIENT_ID,
              scope: 'https://analysis.windows.net/powerbi/api/.default offline_access'
            })
          }
        );
        const data = await res.json();
        if (data.error) throw new Error(data.error_description || data.error);

        // Store device_code for polling
        const sessionId = Math.random().toString(36).slice(2);
        sessions[sessionId] = {
          device_code: data.device_code,
          interval: data.interval || 5,
          expires_at: Date.now() + (data.expires_in * 1000)
        };

        return {
          status: 200,
          headers: CORS_HEADERS,
          body: JSON.stringify({
            sessionId,
            user_code: data.user_code,
            verification_uri: data.verification_uri,
            message: data.message,
            expires_in: data.expires_in
          })
        };
      }

      // Step 2: Poll for token
      if (action === 'pollToken') {
        const sessionId = url.searchParams.get('sessionId');
        const session = sessions[sessionId];
        if (!session) throw new Error('Session not found or expired');

        const res = await fetch(
          `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
          {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
              grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
              client_id: CLIENT_ID,
              device_code: session.device_code
            })
          }
        );
        const data = await res.json();

        if (data.error === 'authorization_pending') {
          return { status: 200, headers: CORS_HEADERS, body: JSON.stringify({ status: 'pending' }) };
        }
        if (data.error) throw new Error(data.error_description || data.error);

        // Token received — clean up session
        delete sessions[sessionId];
        return {
          status: 200,
          headers: CORS_HEADERS,
          body: JSON.stringify({ status: 'complete', access_token: data.access_token })
        };
      }

      // Step 3: Power BI API calls (token passed from browser)
      const token = request.headers.get('x-pbi-token');
      if (!token) throw new Error('No token provided');

      if (action === 'workspaces') {
        const data = await pbiGet(token, '/groups?$top=1000');
        return { status: 200, headers: CORS_HEADERS, body: JSON.stringify(data) };
      }

      if (action === 'reports') {
        const wsId = url.searchParams.get('wsId');
        const data = await pbiGet(token, `/groups/${wsId}/reports`);
        return { status: 200, headers: CORS_HEADERS, body: JSON.stringify(data) };
      }

      if (action === 'datasets') {
        const wsId = url.searchParams.get('wsId');
        const data = await pbiGet(token, `/groups/${wsId}/datasets`);
        return { status: 200, headers: CORS_HEADERS, body: JSON.stringify(data) };
      }

      if (action === 'datasources') {
        const wsId = url.searchParams.get('wsId');
        const dsId = url.searchParams.get('dsId');
        const data = await pbiGet(token, `/groups/${wsId}/datasets/${dsId}/datasources`);
        return { status: 200, headers: CORS_HEADERS, body: JSON.stringify(data) };
      }

      return {
        status: 400,
        headers: CORS_HEADERS,
        body: JSON.stringify({ error: 'Unknown action' })
      };

    } catch (e) {
      context.error('Proxy error:', e.message);
      return {
        status: 500,
        headers: CORS_HEADERS,
        body: JSON.stringify({ error: e.message })
      };
    }
  }
});
