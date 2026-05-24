(function attachBackgroundCodex2ApiApi(root, factory) {
  root.MultiPageBackgroundCodex2ApiApi = factory();
})(typeof self !== 'undefined' ? self : globalThis, function createBackgroundCodex2ApiApiModule() {
  const moduleRoot = typeof self !== 'undefined' ? self : globalThis;

  function createCodex2ApiApi(deps = {}) {
    const {
      addLog = async () => {},
      BlobImpl = typeof Blob === 'function' ? Blob : null,
      buildSessionAuthJson = null,
      fetchImpl = typeof fetch === 'function' ? fetch.bind(globalThis) : null,
      FormDataImpl = typeof FormData === 'function' ? FormData : null,
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

    function getSessionAuthJsonBuilder() {
      if (typeof buildSessionAuthJson === 'function') {
        return buildSessionAuthJson;
      }
      const createCpaApi = moduleRoot?.MultiPageBackgroundCpaApi?.createCpaApi;
      if (typeof createCpaApi !== 'function') {
        throw new Error('CPA SESSION JSON 转换模块未加载，无法保留 Codex2API 套餐信息。');
      }
      const cpaApi = createCpaApi({ addLog });
      if (typeof cpaApi?.buildCpaSessionAuthJson !== 'function') {
        throw new Error('CPA SESSION JSON 转换模块不可用，无法保留 Codex2API 套餐信息。');
      }
      return cpaApi.buildCpaSessionAuthJson;
    }

    function buildCodex2ApiSessionImportArtifact(state = {}, options = {}) {
      const sessionAuth = getSessionAuthJsonBuilder()(state, options);
      const authJson = sessionAuth?.authJson;
      if (!authJson || typeof authJson !== 'object' || Array.isArray(authJson)) {
        throw new Error('未生成可导入 Codex2API 的 SESSION JSON。');
      }
      const fileName = normalizeString(sessionAuth.fileName) || `${resolveAccountName(state)}.json`;
      return {
        name: firstNonEmpty(sessionAuth.email, authJson.email, resolveAccountName(state)),
        email: firstNonEmpty(sessionAuth.email, authJson.email),
        fileName,
        authJson,
        jsonText: JSON.stringify(authJson, null, 2),
        hasRefreshToken: Boolean(sessionAuth.hasRefreshToken),
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

    async function fetchCodex2ApiMultipart(rawUrl, path, options = {}) {
      if (typeof fetchImpl !== 'function') {
        throw new Error('当前环境不支持 fetch，无法请求 Codex2API。');
      }
      if (typeof FormDataImpl !== 'function') {
        throw new Error('当前环境不支持 FormData，无法上传 Codex2API 导入文件。');
      }
      if (typeof BlobImpl !== 'function') {
        throw new Error('当前环境不支持 Blob，无法上传 Codex2API 导入文件。');
      }
      const adminKey = normalizeString(options.adminKey);
      if (!adminKey) {
        throw new Error('尚未配置 Codex2API 管理密钥，请先在侧边栏填写。');
      }

      const url = resolveCodex2ApiEndpoint(rawUrl, path);
      const form = new FormDataImpl();
      for (const [key, value] of Object.entries(options.fields || {})) {
        form.append(key, normalizeString(value));
      }
      for (const file of options.files || []) {
        const fieldName = normalizeString(file.fieldName) || 'file';
        const fileName = normalizeString(file.fileName) || 'codex-session.json';
        const content = file.content instanceof BlobImpl
          ? file.content
          : new BlobImpl([normalizeString(file.content)], { type: normalizeString(file.type) || 'application/json' });
        form.append(fieldName, content, fileName);
      }

      const timeoutMs = Math.max(1000, Math.floor(Number(options.timeoutMs) || 30000));
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), timeoutMs);

      try {
        const response = await fetchImpl(url, {
          method: options.method || 'POST',
          headers: {
            Accept: 'application/json',
            'X-Admin-Key': adminKey,
          },
          body: form,
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
      const artifact = buildCodex2ApiSessionImportArtifact(state, options);
      const logLabel = normalizeString(options.logLabel) || 'Codex2API SESSION JSON 导入';

      await logWithOptions(`${logLabel}：正在通过 Codex2API JSON 导入当前 ChatGPT 会话...`, 'info', options);
      if (!artifact.hasRefreshToken) {
        await logWithOptions(`${logLabel}：未包含 refresh_token，access_token 过期后无法自动续期。`, 'warn', options);
      }
      await fetchCodex2ApiMultipart(state?.codex2apiUrl, '/api/admin/accounts/import', {
        method: 'POST',
        adminKey,
        fields: { format: 'json' },
        files: [{
          fieldName: 'file',
          fileName: artifact.fileName,
          content: artifact.jsonText,
          type: 'application/json',
        }],
        timeoutMs: options.importTimeoutMs || options.timeoutMs,
      });

      const verifiedStatus = `Codex2API SESSION JSON 导入完成：${artifact.email || artifact.fileName}`;
      await logWithOptions(verifiedStatus, 'ok', options);
      return {
        verifiedStatus,
        codex2apiImportedAccountName: artifact.name,
        codex2apiImportedEmail: artifact.email || null,
        codex2apiImportedFileName: artifact.fileName,
      };
    }

    return {
      buildCodex2ApiAtAccountPayload,
      buildCodex2ApiSessionImportArtifact,
      fetchCodex2ApiJson,
      fetchCodex2ApiMultipart,
      importCurrentChatGptSession,
      resolveAccountName,
      resolveCodex2ApiEndpoint,
    };
  }

  return {
    createCodex2ApiApi,
  };
});
