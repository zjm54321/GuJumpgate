const assert = require('node:assert/strict');
const test = require('node:test');

require('../background/codex2api-api.js');

const { createCodex2ApiApi } = globalThis.MultiPageBackgroundCodex2ApiApi;

test('Codex2API AT payload uses session email and access token only', () => {
  const api = createCodex2ApiApi();

  const result = api.buildCodex2ApiAtAccountPayload({
    session: {
      user: { email: 'plus@example.com' },
      accessToken: 'at-session-token',
    },
    refreshToken: 'rt-ignored',
  });

  assert.deepEqual(result, {
    name: 'plus@example.com',
    payload: {
      name: 'plus@example.com',
      access_token: 'at-session-token',
    },
  });
});

test('Codex2API AT payload requires an access token', () => {
  const api = createCodex2ApiApi();

  assert.throws(
    () => api.buildCodex2ApiAtAccountPayload({ session: { user: { email: 'plus@example.com' } } }),
    /未读取到可导入 Codex2API 的 ChatGPT accessToken/
  );
});

test('Codex2API JSON request posts to AT endpoint with admin key', async () => {
  const requests = [];
  const api = createCodex2ApiApi({
    fetchImpl: async (url, options) => {
      requests.push({ url, options });
      return {
        ok: true,
        status: 200,
        text: async () => '{"ok":true}',
      };
    },
    normalizeCodex2ApiUrl: (value) => String(value).replace(/\/+$/, '/')
  });

  const payload = { name: 'plus@example.com', access_token: 'at-token' };
  const result = await api.fetchCodex2ApiJson('https://proxy.example.com/admin', '/api/admin/accounts/at', {
    method: 'POST',
    adminKey: 'secret-key',
    body: payload,
    timeoutMs: 1000,
  });

  assert.deepEqual(result, { ok: true });
  assert.equal(requests.length, 1);
  assert.equal(requests[0].url, 'https://proxy.example.com/api/admin/accounts/at');
  assert.equal(requests[0].options.method, 'POST');
  assert.equal(requests[0].options.headers['X-Admin-Key'], 'secret-key');
  assert.equal(requests[0].options.body, JSON.stringify(payload));
});

test('Codex2API JSON request surfaces API error messages', async () => {
  const api = createCodex2ApiApi({
    fetchImpl: async () => ({
      ok: false,
      status: 400,
      text: async () => '{"error":"invalid access token"}',
    }),
    normalizeCodex2ApiUrl: (value) => value,
  });

  await assert.rejects(
    () => api.fetchCodex2ApiJson('https://proxy.example.com', '/api/admin/accounts/at', {
      method: 'POST',
      adminKey: 'secret-key',
      body: { name: 'plus@example.com', access_token: 'bad-token' },
      timeoutMs: 1000,
    }),
    /invalid access token/
  );
});
