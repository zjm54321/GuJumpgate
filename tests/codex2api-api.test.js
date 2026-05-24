const assert = require('node:assert/strict');
const test = require('node:test');

require('../background/cpa-api.js');
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

test('Codex2API SESSION JSON artifact preserves Plus plan metadata', () => {
  const api = createCodex2ApiApi();

  const result = api.buildCodex2ApiSessionImportArtifact({
    accessToken: 'at-session-token',
    session: {
      user: { id: 'user-123', email: 'plus@example.com' },
      account: { id: 'acct-123', planType: 'plus' },
      expires: '2026-01-02T03:04:05.000Z',
    },
  }, { now: '2026-01-01T00:00:00.000Z' });

  assert.equal(result.email, 'plus@example.com');
  assert.equal(result.fileName, 'codex-plus@example.com-plus.json');
  assert.equal(result.hasRefreshToken, false);
  assert.equal(result.authJson.access_token, 'at-session-token');
  assert.equal(result.authJson.plan_type, 'plus');
  assert.equal(result.authJson.chatgpt_plan_type, 'plus');
  assert.equal(result.authJson.account_id, 'acct-123');
  assert.equal(result.authJson.chatgpt_account_id, 'acct-123');
  assert.equal(result.authJson.id_token_synthetic, true);
  assert.match(result.authJson.id_token, /^ey/);
  assert.deepEqual(JSON.parse(result.jsonText), result.authJson);
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
    normalizeCodex2ApiUrl: (value) => String(value).replace(/\/+$/, '/'),
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

test('Codex2API session import uploads CPA JSON with format=json', async () => {
  const requests = [];
  const logs = [];
  const api = createCodex2ApiApi({
    addLog: async (message, level) => logs.push({ message, level }),
    buildSessionAuthJson: () => ({
      authJson: {
        email: 'plus@example.com',
        access_token: 'at-token',
        plan_type: 'plus',
        chatgpt_plan_type: 'plus',
      },
      email: 'plus@example.com',
      fileName: 'codex-plus@example.com-plus.json',
      hasRefreshToken: false,
    }),
    fetchImpl: async (url, options) => {
      requests.push({ url, options });
      return {
        ok: true,
        status: 200,
        text: async () => '{"ok":true}',
      };
    },
    normalizeCodex2ApiUrl: (value) => String(value).replace(/\/+$/, '/'),
  });

  const result = await api.importCurrentChatGptSession({
    codex2apiAdminKey: 'secret-key',
    codex2apiUrl: 'https://proxy.example.com/admin',
  }, { timeoutMs: 1000 });

  assert.equal(result.codex2apiImportedEmail, 'plus@example.com');
  assert.equal(result.codex2apiImportedFileName, 'codex-plus@example.com-plus.json');
  assert.equal(requests.length, 1);
  assert.equal(requests[0].url, 'https://proxy.example.com/api/admin/accounts/import');
  assert.equal(requests[0].options.method, 'POST');
  assert.equal(requests[0].options.headers['X-Admin-Key'], 'secret-key');
  assert.equal(requests[0].options.headers['Content-Type'], undefined);

  const formEntries = Array.from(requests[0].options.body.entries());
  assert.equal(formEntries[0][0], 'format');
  assert.equal(formEntries[0][1], 'json');
  assert.equal(formEntries[1][0], 'file');
  assert.equal(formEntries[1][1].name, 'codex-plus@example.com-plus.json');
  assert.deepEqual(JSON.parse(await formEntries[1][1].text()), {
    email: 'plus@example.com',
    access_token: 'at-token',
    plan_type: 'plus',
    chatgpt_plan_type: 'plus',
  });
  assert.ok(logs.some((entry) => entry.level === 'warn' && entry.message.includes('未包含 refresh_token')));
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
