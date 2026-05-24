(function attachBackgroundCodex2ApiApi(root, factory) {
  root.MultiPageBackgroundCodex2ApiApi = factory();
})(typeof self !== 'undefined' ? self : globalThis, function createBackgroundCodex2ApiApiModule() {
  function createCodex2ApiApi(deps = {}) {
    const {
      addLog = async () => {},
      fetchImpl = typeof fetch === 'function' ? fetch.bind(globalThis) : null,
      normalizeCodex2ApiUrl = (value) => String(value || '').trim(),
    } = deps;

    function normalizeString(value = '') {
      return String(value || '').trim();
    }

    function getSessionObject(state = {}) {
      return state?.session && typeof state.session === 'object' && !Array.isArray(state.session)
        ? state.session
        : null;
    }

    function firstNonEmpty(...values) {
      return values.map(normalizeString).find(Boolean) || '';
    }

    function resolveAccountName(state = {}) {
      const session = getSessionObject(state);
      return firstNonEmpty(
        state.codex2apiAccountName,
        state.accountName,
        session?.user?.email,
        session?.email,
        state.accountIdentifier,
        state.email,
        state.step8VerificationTargetEmail,
        state.registrationEmail,
        'chatgpt-session'
      );
    }

    function buildCodex2ApiAtAccountPayload(state = {}) {
      const session = getSessionObject(state);
      const accessToken = firstNonEmpty(
        state.accessToken,
        state.access_token,
        session?.accessToken,
        session?.access_token
      );
      if (!accessToken) {
        throw new Error('未读取到可导入 Codex2API 的 ChatGPT accessToken。');
      }

      const name = resolveAccountName(state);
      return {
        name,
        payload: {
          name,
          access_token: accessToken,
        },
      };
    }

    function getCodex2ApiErrorMessage(payload, responseStatus = 500) {
      const candidates = [
        payload?.error,
        payload?.message,
        payload?.detail,
        payload?.reason,
        payload?.raw,
      ];
      const message = candidates.map(normalizeString).find(Boolean);
      return message || `Codex2API 请求失败（HTTP ${responseStatus}）。`;
    }

    function resolveCodex2ApiEndpoint(rawUrl = '', path = '/') {
      const normalizedUrl = normalizeCodex2ApiUrl(rawUrl);
      if (!normalizedUrl) {
        throw new Error('尚未配置 Codex2API 地址，请先在侧边栏填写。');
      }
      let parsed;
      try {
        parsed = new URL(path, normalizedUrl);
      } catch (error) {
        throw new Error(`Codex2API 地址格式无效，请先在侧边栏检查：${error?.message || error}`);
      }
      return parsed.toString();
    }

    async function parseJsonResponse(response) {
      const text = await response.text();
      const trimmed = normalizeString(text);
      if (!trimmed) {
        return null;
      }
      try {
        return JSON.parse(trimmed);
      } catch (error) {
        return { raw: trimmed };
      }
    }

    async function fetchCodex2ApiJson(rawUrl, path, options = {}) {
      if (typeof fetchImpl !== 'function') {
        throw new Error('当前环境不支持 fetch，无法请求 Codex2API。');
      }
      const adminKey = normalizeString(options.adminKey);
      if (!adminKey) {
        throw new Error('尚未配置 Codex2API 管理密钥，请先在侧边栏填写。');
      }

      const url = resolveCodex2ApiEndpoint(rawUrl, path);
      const timeoutMs = Math.max(1000, Math.floor(Number(options.timeoutMs) || 30000));
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), timeoutMs);

      try {
        const response = await fetchImpl(url, {
          method: options.method || 'GET',
          headers: {
            Accept: 'application/json',
            'Content-Type': 'application/json',
            'X-Admin-Key': adminKey,
          },
          body: options.body === undefined ? undefined : JSON.stringify(options.body),
          signal: controller.signal,
        });
        const payload = await parseJsonResponse(response);
        if (!response.ok) {
          throw new Error(getCodex2ApiErrorMessage(payload, response.status));
        }
        return payload;
      } catch (error) {
        if (error?.name === 'AbortError') {
          throw new Error(`Codex2API 请求超时（>${Math.round(timeoutMs / 1000)} 秒）。`);
        }
        throw error;
      } finally {
        clearTimeout(timer);
      }
    }

    async function logWithOptions(message, level = 'info', options = {}) {
      if (typeof addLog !== 'function') {
        return;
      }
      await addLog(message, level, options.logOptions || {});
    }

    async function importCurrentChatGptSession(state = {}, options = {}) {
      const adminKey = normalizeString(state?.codex2apiAdminKey);
      if (!adminKey) {
        throw new Error('尚未配置 Codex2API 管理密钥，请先在侧边栏填写。');
      }
      const account = buildCodex2ApiAtAccountPayload(state);
      const logLabel = normalizeString(options.logLabel) || 'Codex2API AT 导入';

      await logWithOptions(`${logLabel}：正在通过 Codex2API AT 接口导入当前 ChatGPT 会话...`, 'info', options);
      await fetchCodex2ApiJson(state?.codex2apiUrl, '/api/admin/accounts/at', {
        method: 'POST',
        adminKey,
        body: account.payload,
        timeoutMs: options.importTimeoutMs || options.timeoutMs,
      });

      const verifiedStatus = `Codex2API AT 导入完成：${account.name}`;
      await logWithOptions(verifiedStatus, 'ok', options);
      return {
        verifiedStatus,
        codex2apiImportedAccountName: account.name,
      };
    }

    return {
      buildCodex2ApiAtAccountPayload,
      fetchCodex2ApiJson,
      importCurrentChatGptSession,
      resolveAccountName,
      resolveCodex2ApiEndpoint,
    };
  }

  return {
    createCodex2ApiApi,
  };
});
