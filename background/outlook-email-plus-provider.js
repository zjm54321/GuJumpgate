(function outlookEmailPlusProviderModule(root, factory) {
  root.MultiPageBackgroundOutlookEmailPlusProvider = factory();
})(typeof self !== 'undefined' ? self : globalThis, function createOutlookEmailPlusProviderModule() {
  function createOutlookEmailPlusProvider(deps = {}) {
    const {
      addLog = async () => {},
      buildOutlookEmailPlusAliasAddress,
      buildOutlookEmailPlusHeaders,
      buildOutlookEmailPlusPayPalAliasAddress,
      deriveOutlookEmailPlusBaseAddress,
      generateOutlookEmailPlusTag,
      getOutlookEmailPlusPayPalAliasIndex,
      isOutlookEmailPlusTaggedAlias,
      joinOutlookEmailPlusUrl,
      normalizeOutlookEmailPlusAddress,
      normalizeOutlookEmailPlusBaseUrl,
      normalizeOutlookEmailPlusCallerIdPrefix,
      normalizeOutlookEmailPlusClaim,
      normalizeOutlookEmailPlusProjectKey,
      normalizeOutlookEmailPlusProvider,
      normalizeOutlookEmailPlusVerificationCode,
      OUTLOOK_EMAIL_PLUS_GENERATOR = 'outlook-email-plus',
      OUTLOOK_EMAIL_PLUS_PROVIDER = 'outlook-email-plus',
      fetchImpl = typeof fetch === 'function' ? fetch.bind(globalThis) : null,
      getState = async () => ({}),
      persistRegistrationEmailState = null,
      setEmailState = async () => {},
      setState = async () => {},
      sleepWithStop = async () => {},
      throwIfStopped = () => {},
      unwrapOutlookEmailPlusResponse,
    } = deps;
    const activeClaims = new Map();
    const DEFAULT_ALIAS_MAX_PER_MAILBOX = 5;

    async function persistResolvedEmailState(state = null, email, options = {}) {
      if (typeof persistRegistrationEmailState === 'function') {
        await persistRegistrationEmailState(state, email, options);
        return;
      }
      await setEmailState(email, options);
    }

    function getOutlookEmailPlusConfig(state = {}) {
      return {
        baseUrl: normalizeOutlookEmailPlusBaseUrl(state.outlookEmailPlusBaseUrl),
        apiKey: String(state.outlookEmailPlusApiKey || '').trim(),
        provider: normalizeOutlookEmailPlusProvider(state.outlookEmailPlusProvider) || 'outlook',
        projectKey: normalizeOutlookEmailPlusProjectKey(state.outlookEmailPlusProjectKey) || 'openai',
        callerIdPrefix: normalizeOutlookEmailPlusCallerIdPrefix(state.outlookEmailPlusCallerIdPrefix) || 'gujumpgate',
      };
    }

    function ensureOutlookEmailPlusConfig(state, options = {}) {
      const { requireApiKey = true } = options;
      const config = getOutlookEmailPlusConfig(state);
      if (!config.baseUrl) {
        throw new Error('Outlook Email Plus 服务地址为空或格式无效。');
      }
      if (requireApiKey && !config.apiKey) {
        throw new Error('Outlook Email Plus API Key 为空。');
      }
      if (!config.provider) {
        throw new Error('Outlook Email Plus 邮箱提供商为空。');
      }
      if (!config.projectKey) {
        throw new Error('Outlook Email Plus Project Key 为空。');
      }
      return config;
    }

    async function requestOutlookEmailPlusJson(config, path, options = {}) {
      if (!fetchImpl) {
        throw new Error('Outlook Email Plus 当前运行环境不支持 fetch。');
      }

      const {
        method = 'GET',
        payload,
        searchParams = null,
        timeoutMs = 20000,
      } = options;
      const url = new URL(joinOutlookEmailPlusUrl(config.baseUrl, path));
      if (searchParams && typeof searchParams === 'object') {
        for (const [key, value] of Object.entries(searchParams)) {
          if (value === undefined || value === null || value === '') continue;
          url.searchParams.set(key, String(value));
        }
      }

      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(new Error('timeout')), timeoutMs);
      let response;
      try {
        response = await fetchImpl(url.toString(), {
          method,
          headers: buildOutlookEmailPlusHeaders(config, {
            json: payload !== undefined,
          }),
          body: payload !== undefined ? JSON.stringify(payload) : undefined,
          signal: controller.signal,
        });
      } catch (err) {
        const errorMessage = err?.name === 'AbortError'
          ? `Outlook Email Plus 请求超时（>${Math.round(timeoutMs / 1000)} 秒）`
          : `Outlook Email Plus 请求失败：${err?.message || err}`;
        throw new Error(errorMessage);
      } finally {
        clearTimeout(timeoutId);
      }

      const text = await response.text();
      let parsed;
      try {
        parsed = text ? JSON.parse(text) : {};
      } catch (_error) {
        parsed = text;
      }

      if (!response.ok) {
        let payloadMessage = '';
        try {
          payloadMessage = String(unwrapOutlookEmailPlusResponse(parsed)?.message || '');
        } catch (error) {
          payloadMessage = String(error?.message || '');
        }
        if (!payloadMessage && parsed && typeof parsed === 'object') {
          payloadMessage = String(parsed.message || parsed.error || parsed.msg || '');
        }
        throw new Error(`Outlook Email Plus 请求失败：${payloadMessage || text || `HTTP ${response.status}`}`);
      }

      return unwrapOutlookEmailPlusResponse(parsed);
    }

    function buildRandomIdentifier() {
      if (globalThis.crypto?.randomUUID) {
        return globalThis.crypto.randomUUID();
      }
      return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
    }

    function normalizeIdentifierPart(value = '') {
      return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9._-]+/g, '-')
        .replace(/-{2,}/g, '-')
        .replace(/^[-._]+|[-._]+$/g, '');
    }

    function resolveClaimTaskId(state = {}, options = {}) {
      return normalizeIdentifierPart(
        options.taskId
        || state.currentOutlookEmailPlusClaim?.taskId
        || state.taskId
        || state.activeRunId
        || state.runId
      ) || buildRandomIdentifier();
    }

    function buildCallerId(config, state = {}, options = {}, taskId = '') {
      const prefix = normalizeOutlookEmailPlusCallerIdPrefix(
        options.callerIdPrefix || config.callerIdPrefix || state.outlookEmailPlusCallerIdPrefix || 'gujumpgate'
      ) || 'gujumpgate';
      const explicitCallerId = normalizeIdentifierPart(options.callerId || state.currentOutlookEmailPlusClaim?.callerId);
      if (explicitCallerId) {
        return explicitCallerId;
      }
      const suffix = normalizeIdentifierPart(options.runId || state.runId || taskId) || normalizeIdentifierPart(buildRandomIdentifier());
      return `${prefix}-${suffix}`;
    }

    function getClaimKeys(claim = {}) {
      return [claim.taskId, claim.accountId, claim.address]
        .map((value) => String(value || '').trim().toLowerCase())
        .filter(Boolean);
    }

    function rememberClaim(claim = {}, storedClaim = {}) {
      const secretClaim = {
        ...storedClaim,
        claimToken: claim.claimToken || '',
      };
      for (const key of getClaimKeys(storedClaim)) {
        activeClaims.set(key, secretClaim);
      }
    }

    function forgetClaim(claim = {}) {
      for (const key of getClaimKeys(claim)) {
        activeClaims.delete(key);
      }
    }

    function getRememberedClaim(storedClaim = {}) {
      for (const key of getClaimKeys(storedClaim)) {
        const remembered = activeClaims.get(key);
        if (remembered) return remembered;
      }
      return null;
    }

    function buildStoredOutlookEmailPlusClaim(claim = {}, context = {}) {
      return {
        accountId: claim.accountId || '',
        address: claim.address || '',
        baseAddress: claim.baseAddress || '',
        registrationEmail: claim.registrationEmail || claim.address || '',
        isAliasClaim: Boolean(claim.isAliasClaim),
        domain: claim.domain || '',
        claimedAt: claim.claimedAt || '',
        leaseExpiresAt: claim.leaseExpiresAt || '',
        callerId: context.callerId || '',
        taskId: context.taskId || '',
        projectKey: context.projectKey || '',
        provider: context.provider || '',
        aliasIndex: Math.max(0, Math.floor(Number(claim.aliasIndex) || 0)),
        aliasMax: normalizeAliasMax(claim.aliasMax || context.aliasMax),
        aliasUsed: Boolean(claim.aliasUsed),
      };
    }

    function buildLifecyclePayload(claim = {}, options = {}) {
      const payload = {
        account_id: claim.accountId || undefined,
        email: claim.address || undefined,
        claim_token: claim.claimToken || undefined,
        caller_id: claim.callerId || undefined,
        task_id: claim.taskId || undefined,
        project_key: claim.projectKey || undefined,
        result: options.result || undefined,
        reason: options.reason || undefined,
      };
      for (const [key, value] of Object.entries(payload)) {
        if (value === undefined || value === null || value === '') {
          delete payload[key];
        }
      }
      return payload;
    }

    async function clearStoredClaim(storedClaim = {}) {
      forgetClaim(storedClaim);
      if (typeof setState === 'function') {
        await setState({ currentOutlookEmailPlusClaim: null });
      }
    }

    function resolveLifecycleClaim(state = {}, options = {}) {
      const storedClaim = options.claim && typeof options.claim === 'object'
        ? options.claim
        : (state.currentOutlookEmailPlusClaim || {});
      if (!storedClaim || typeof storedClaim !== 'object') {
        return null;
      }
      const rememberedClaim = getRememberedClaim(storedClaim) || {};
      return {
        ...storedClaim,
        ...rememberedClaim,
      };
    }

    function normalizeAliasMax(value, fallback = DEFAULT_ALIAS_MAX_PER_MAILBOX) {
      const numeric = Math.floor(Number(value));
      if (Number.isFinite(numeric) && numeric >= 1) {
        return Math.min(50, numeric);
      }
      const fallbackNumber = Math.floor(Number(fallback));
      if (Number.isFinite(fallbackNumber) && fallbackNumber >= 1) {
        return Math.min(50, fallbackNumber);
      }
      return DEFAULT_ALIAS_MAX_PER_MAILBOX;
    }

    function getAliasMaxForState(state = {}, fallback = DEFAULT_ALIAS_MAX_PER_MAILBOX) {
      return normalizeAliasMax(
        state.outlookEmailPlusAliasMaxPerMailbox ?? state.outlookAliasMaxPerAccount,
        fallback
      );
    }

    function buildPayPalAlias(baseAddress = '', index = 1) {
      if (typeof buildOutlookEmailPlusPayPalAliasAddress === 'function') {
        return buildOutlookEmailPlusPayPalAliasAddress(baseAddress, index);
      }
      const numericIndex = Math.max(1, Math.floor(Number(index) || 1));
      return buildOutlookEmailPlusAliasAddress(baseAddress, `PayPal${numericIndex}`);
    }

    function getPayPalAliasIndex(aliasAddress = '', baseAddress = '') {
      if (typeof getOutlookEmailPlusPayPalAliasIndex === 'function') {
        return getOutlookEmailPlusPayPalAliasIndex(aliasAddress, baseAddress);
      }
      return null;
    }

    function allocateRegistrationAlias(baseAddress = '', index = 1) {
      const registrationEmail = buildPayPalAlias(baseAddress, index);
      if (!registrationEmail) {
        throw new Error('Outlook Email Plus 无法基于认领邮箱生成注册别名。');
      }
      return registrationEmail;
    }

    function resolveClaimRegistration(claim = {}, context = {}) {
      const claimedAddress = normalizeOutlookEmailPlusAddress(claim.address);
      const baseAddress = deriveOutlookEmailPlusBaseAddress(claimedAddress) || claimedAddress;
      const isAliasClaim = Boolean(isOutlookEmailPlusTaggedAlias(claimedAddress));
      const aliasMax = normalizeAliasMax(context.aliasMax);
      const aliasIndex = Math.max(1, Math.floor(Number(context.aliasIndex) || 1));
      const registrationEmail = allocateRegistrationAlias(baseAddress, aliasIndex);

      return {
        ...claim,
        address: claimedAddress,
        baseAddress,
        registrationEmail,
        isAliasClaim,
        aliasIndex,
        aliasMax,
        aliasUsed: false,
      };
    }

    function getReusableStoredClaim(state = {}) {
      const storedClaim = state.currentOutlookEmailPlusClaim;
      if (!storedClaim || typeof storedClaim !== 'object') {
        return null;
      }
      const rememberedClaim = getRememberedClaim(storedClaim);
      if (!rememberedClaim?.claimToken) {
        return null;
      }
      const address = normalizeOutlookEmailPlusAddress(storedClaim.address);
      const baseAddress = normalizeOutlookEmailPlusAddress(storedClaim.baseAddress)
        || deriveOutlookEmailPlusBaseAddress(address)
        || address;
      if (!address || !baseAddress) {
        return null;
      }
      return {
        ...storedClaim,
        address,
        baseAddress,
        aliasIndex: Math.max(0, Math.floor(Number(storedClaim.aliasIndex) || 0)),
        aliasMax: normalizeAliasMax(storedClaim.aliasMax, getAliasMaxForState(state)),
        aliasUsed: Boolean(storedClaim.aliasUsed),
      };
    }

    function buildReusableRegistrationClaim(storedClaim = {}, context = {}) {
      const aliasMax = normalizeAliasMax(storedClaim.aliasMax, getAliasMaxForState(context.state || {}, storedClaim.aliasMax));
      if (!storedClaim.aliasUsed && storedClaim.registrationEmail) {
        return {
          ...storedClaim,
          aliasMax,
        };
      }
      const nextAliasIndex = Math.max(0, Math.floor(Number(storedClaim.aliasIndex) || 0)) + 1;
      if (nextAliasIndex > aliasMax) {
        return {
          ...storedClaim,
          aliasMax,
          exhausted: true,
        };
      }
      const registrationEmail = allocateRegistrationAlias(storedClaim.baseAddress, nextAliasIndex);
      return {
        ...storedClaim,
        aliasIndex: nextAliasIndex,
        aliasMax,
        aliasUsed: false,
        registrationEmail,
      };
    }

    async function reuseOutlookEmailPlusClaimAddress(state = {}, config = {}, options = {}) {
      const storedClaim = getReusableStoredClaim(state);
      if (!storedClaim) {
        return { reused: false };
      }
      const claim = buildReusableRegistrationClaim(storedClaim, {
        state,
      });
      if (claim.exhausted) {
        return { reused: false, exhausted: true, claim };
      }
      await persistResolvedEmailState(state, claim.registrationEmail, {
        source: `generated:${OUTLOOK_EMAIL_PLUS_GENERATOR}`,
        preserveAccountIdentity: Boolean(options?.preserveAccountIdentity),
      });
      const nextStoredClaim = buildStoredOutlookEmailPlusClaim(claim, {
        callerId: storedClaim.callerId,
        taskId: storedClaim.taskId,
        projectKey: storedClaim.projectKey || config.projectKey,
        provider: storedClaim.provider || config.provider,
      });
      if (typeof setState === 'function') {
        await setState({ currentOutlookEmailPlusClaim: nextStoredClaim });
      }
      await addLog(`Outlook Email Plus：复用 ${claim.address}，本次注册使用 ${claim.registrationEmail}`, 'ok');
      return { reused: true, address: claim.registrationEmail };
    }

    async function claimOutlookEmailPlusAddress(state, options = {}) {
      throwIfStopped();
      let latestState = state || await getState();
      const config = ensureOutlookEmailPlusConfig(latestState);
      const reuseResult = await reuseOutlookEmailPlusClaimAddress(latestState, config, options);
      if (reuseResult.reused) {
        return reuseResult.address;
      }
      if (reuseResult.exhausted) {
        await completeOutlookEmailPlusClaim(latestState, { result: 'success' });
        latestState = await getState();
      }
      const taskId = resolveClaimTaskId(latestState, options);
      const callerId = buildCallerId(config, latestState, options, taskId);
      const aliasMax = getAliasMaxForState(latestState);
      const payload = {
        provider: config.provider,
        project_key: config.projectKey,
        caller_id: callerId,
        task_id: taskId,
      };

      const result = await requestOutlookEmailPlusJson(config, '/api/external/pool/claim-random', {
        method: 'POST',
        payload,
      });
      const claim = resolveClaimRegistration(normalizeOutlookEmailPlusClaim(result), {
        aliasTag: options.aliasTag,
        aliasIndex: 1,
        aliasMax,
        callerId,
        projectKey: config.projectKey,
        taskId,
      });
      if (!claim.address) {
        throw new Error('Outlook Email Plus 未返回可用邮箱地址。');
      }

      await persistResolvedEmailState(latestState, claim.registrationEmail, {
        source: `generated:${OUTLOOK_EMAIL_PLUS_GENERATOR}`,
        preserveAccountIdentity: Boolean(options?.preserveAccountIdentity),
      });
      const storedClaim = buildStoredOutlookEmailPlusClaim(claim, {
        callerId,
        taskId,
        projectKey: config.projectKey,
        provider: config.provider,
      });
      rememberClaim(claim, storedClaim);
      if (typeof setState === 'function') {
        await setState({ currentOutlookEmailPlusClaim: storedClaim });
      }
      await addLog(`Outlook Email Plus：已认领 ${claim.address}，注册使用 ${claim.registrationEmail}`, 'ok');
      return claim.registrationEmail;
    }

    async function markOutlookEmailPlusAliasUsed(state = {}) {
      const latestState = state || await getState();
      const storedClaim = latestState.currentOutlookEmailPlusClaim;
      if (!storedClaim || typeof storedClaim !== 'object') {
        return { handled: false, reason: 'missing_claim' };
      }
      const aliasMax = normalizeAliasMax(storedClaim.aliasMax, getAliasMaxForState(latestState));
      const aliasIndex = Math.max(
        getPayPalAliasIndex(storedClaim.registrationEmail, storedClaim.baseAddress) || 0,
        Math.floor(Number(storedClaim.aliasIndex) || 0)
      );
      const normalizedAliasIndex = Math.max(1, aliasIndex || 1);
      const alreadyUsed = Boolean(storedClaim.aliasUsed);
      const nextStoredClaim = {
        ...storedClaim,
        aliasIndex: normalizedAliasIndex,
        aliasMax,
        aliasUsed: true,
      };
      if (!alreadyUsed && typeof setState === 'function') {
        await setState({ currentOutlookEmailPlusClaim: nextStoredClaim });
      }
      return {
        handled: true,
        alreadyUsed,
        exhausted: normalizedAliasIndex >= aliasMax,
        aliasIndex: normalizedAliasIndex,
        aliasMax,
        registrationEmail: storedClaim.registrationEmail || storedClaim.address || '',
      };
    }

    function resolvePollTargetEmail(state = {}, pollPayload = {}) {
      return normalizeOutlookEmailPlusAddress(
        pollPayload.targetEmail
        || state.registrationEmailState?.current
        || state.email
        || state.currentOutlookEmailPlusClaim?.registrationEmail
        || state.currentOutlookEmailPlusClaim?.address
        || ''
      );
    }

    function resolveSinceMinutes(pollPayload = {}) {
      const configured = Math.floor(Number(pollPayload.sinceMinutes || pollPayload.since_minutes) || 0);
      if (configured > 0) {
        return configured;
      }
      const afterTimestamp = Number(pollPayload.filterAfterTimestamp) || 0;
      if (afterTimestamp <= 0) {
        return 0;
      }
      const ageMs = Math.max(0, Date.now() - afterTimestamp);
      return Math.max(1, Math.ceil(ageMs / 60000));
    }

    async function pollOutlookEmailPlusVerificationCode(step, state, pollPayload = {}) {
      const latestState = state || await getState();
      const config = ensureOutlookEmailPlusConfig(latestState);
      const targetEmail = resolvePollTargetEmail(latestState, pollPayload);
      if (!targetEmail) {
        throw new Error('Outlook Email Plus 轮询前缺少目标邮箱地址，请先获取注册邮箱。');
      }

      const maxAttempts = Math.max(1, Math.floor(Number(pollPayload.maxAttempts) || 5));
      const intervalMs = Math.max(0, Number(pollPayload.intervalMs) || 3000);
      const excludeCodes = new Set(
        (Array.isArray(pollPayload.excludeCodes) ? pollPayload.excludeCodes : [])
          .map((value) => normalizeOutlookEmailPlusVerificationCode(value))
          .filter((value) => typeof value === 'string' && value)
      );
      const sinceMinutes = resolveSinceMinutes(pollPayload);
      const codeLength = Math.max(0, Math.floor(Number(pollPayload.codeLength) || 0));
      const codeRegex = String(pollPayload.codeRegex || '').trim();
      const codeSource = String(pollPayload.codeSource || '').trim();
      let lastError = null;
      let sawNoCode = false;

      for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
        throwIfStopped();
        try {
          const result = await requestOutlookEmailPlusJson(config, '/api/external/verification-code', {
            method: 'GET',
            searchParams: {
              email: targetEmail,
              since_minutes: sinceMinutes > 0 ? sinceMinutes : undefined,
              code_length: codeLength > 0 ? codeLength : undefined,
              code_regex: codeRegex || undefined,
              code_source: codeSource || undefined,
            },
          });
          const verification = normalizeOutlookEmailPlusVerificationCode(result);
          if (verification.code) {
            if (!excludeCodes.has(verification.code)) {
              return {
                ok: true,
                code: verification.code,
                emailTimestamp: verification.emailTimestamp || Date.now(),
                mailId: verification.mailId || '',
              };
            }
            sawNoCode = true;
            lastError = new Error(`步骤 ${step}：Outlook Email Plus 返回了已排除的旧验证码。`);
            await addLog(`步骤 ${step}：Outlook Email Plus 命中过滤掉的旧验证码，继续轮询（${attempt}/${maxAttempts}）。`, 'info');
          } else {
            sawNoCode = true;
            lastError = new Error(`步骤 ${step}：暂未在 Outlook Email Plus 中找到匹配验证码（${attempt}/${maxAttempts}）。`);
            await addLog(lastError.message, attempt === maxAttempts ? 'warn' : 'info');
          }
        } catch (err) {
          lastError = err;
          await addLog(`步骤 ${step}：Outlook Email Plus 轮询失败：${err?.message || err}`, 'warn');
        }

        if (attempt < maxAttempts) {
          await sleepWithStop(intervalMs);
        }
      }

      if (sawNoCode) {
        throw new Error(`步骤 ${step}：未在 Outlook Email Plus 中找到新的匹配验证码。`);
      }
      throw lastError || new Error(`步骤 ${step}：Outlook Email Plus 轮询失败。`);
    }

    async function completeOutlookEmailPlusClaim(state, options = {}) {
      const latestState = state || await getState();
      const config = ensureOutlookEmailPlusConfig(latestState);
      const claim = resolveLifecycleClaim(latestState, options);
      if (!claim?.address && !claim?.accountId) {
        return { completed: false, reason: 'missing_claim' };
      }
      if (!claim.claimToken) {
        return { completed: false, reason: 'missing_claim_token' };
      }
      await requestOutlookEmailPlusJson(config, '/api/external/pool/claim-complete', {
        method: 'POST',
        payload: buildLifecyclePayload(claim, { result: options.result || 'success' }),
      });
      await clearStoredClaim(claim);
      await addLog(`Outlook Email Plus：已完成认领 ${claim.address || claim.accountId}`, 'ok');
      return { completed: true };
    }

    async function releaseOutlookEmailPlusClaim(state, options = {}) {
      const latestState = state || await getState();
      const config = ensureOutlookEmailPlusConfig(latestState);
      const claim = resolveLifecycleClaim(latestState, options);
      if (!claim?.address && !claim?.accountId) {
        return { released: false, reason: 'missing_claim' };
      }
      if (!claim.claimToken) {
        return { released: false, reason: 'missing_claim_token' };
      }
      await requestOutlookEmailPlusJson(config, '/api/external/pool/claim-release', {
        method: 'POST',
        payload: buildLifecyclePayload(claim, { reason: options.reason || 'flow_abandoned' }),
      });
      await clearStoredClaim(claim);
      await addLog(`Outlook Email Plus：已释放认领 ${claim.address || claim.accountId}`, 'warn');
      return { released: true };
    }

    return {
      claimOutlookEmailPlusAddress,
      completeOutlookEmailPlusClaim,
      ensureOutlookEmailPlusConfig,
      getOutlookEmailPlusConfig,
      markOutlookEmailPlusAliasUsed,
      pollOutlookEmailPlusVerificationCode,
      releaseOutlookEmailPlusClaim,
      requestOutlookEmailPlusJson,
    };
  }

  return {
    createOutlookEmailPlusProvider,
  };
});
