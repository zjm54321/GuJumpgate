const assert = require('node:assert/strict');
const test = require('node:test');

require('../shared/flow-capabilities.js');
require('../data/step-definitions.js');

const flowCapabilities = globalThis.MultiPageFlowCapabilities;
const stepDefinitions = globalThis.MultiPageStepDefinitions;

test('Codex2API panel exposes OAuth and SESSION JSON account strategies', () => {
  const registry = flowCapabilities.createFlowCapabilityRegistry();
  const strategy = flowCapabilities.PLUS_ACCOUNT_ACCESS_STRATEGY_CODEX2API_CODEX_SESSION;
  const panelCapabilities = registry.getPanelCapabilities('codex2api');

  assert.deepEqual(panelCapabilities.supportedPlusAccountAccessStrategies, [
    flowCapabilities.PLUS_ACCOUNT_ACCESS_STRATEGY_OAUTH,
    strategy,
  ]);

  const capabilityState = registry.resolveSidepanelCapabilities({
    state: {
      activeFlowId: 'openai',
      panelMode: 'codex2api',
      plusModeEnabled: true,
      plusAccountAccessStrategy: strategy,
      signupMethod: flowCapabilities.SIGNUP_METHOD_EMAIL,
    },
  });

  assert.equal(capabilityState.effectivePlusAccountAccessStrategy, strategy);
  assert.equal(capabilityState.canEditPlusAccountAccessStrategy, true);
  assert.equal(capabilityState.stepDefinitionOptions.plusAccountAccessStrategy, strategy);
});

test('Codex2API SESSION JSON strategy resolves to session import step definitions', () => {
  const strategy = stepDefinitions.PLUS_ACCOUNT_ACCESS_STRATEGY_CODEX2API_CODEX_SESSION;
  const scenarios = [
    { plusPaymentMethod: 'paypal', plusHostedCheckoutIsFinalStep: true },
    { plusPaymentMethod: 'paypal', plusHostedCheckoutIsFinalStep: false },
    { plusPaymentMethod: 'gopay' },
    { plusPaymentMethod: 'gpc-helper' },
  ];

  scenarios.forEach((scenario) => {
    const steps = stepDefinitions.getSteps({
      activeFlowId: 'openai',
      plusModeEnabled: true,
      plusAccountAccessStrategy: strategy,
      signupMethod: stepDefinitions.SIGNUP_METHOD_EMAIL,
      ...scenario,
    });
    const lastStep = steps[steps.length - 1];

    assert.equal(lastStep.key, 'codex2api-session-import', JSON.stringify(scenario));
    assert.equal(lastStep.command, 'codex2api-session-import');
    assert.equal(lastStep.driverId, 'background/codex2api-session-import');
  });
});

test('Codex2API SESSION JSON steps keep OAuth path separate', () => {
  const steps = stepDefinitions.getSteps({
    activeFlowId: 'openai',
    plusModeEnabled: true,
    panelMode: 'codex2api',
    plusPaymentMethod: 'paypal',
    plusAccountAccessStrategy: stepDefinitions.PLUS_ACCOUNT_ACCESS_STRATEGY_OAUTH,
    signupMethod: stepDefinitions.SIGNUP_METHOD_EMAIL,
  });
  const keys = steps.map((step) => step.key);

  assert.equal(keys.includes('codex2api-session-import'), false);
  assert.equal(keys.includes('platform-verify'), true);
});
