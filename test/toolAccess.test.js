const { test, describe, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const {
  normalizeEmail,
  normalizeToolAccess,
  decodeFirestoreFields,
  makeFetchOwnUserDoc,
  resolveToolAccess,
  makeRequireToolAccess,
  DENY_NO_ACCESS,
  DENY_INACTIVE,
} = require('../lib/toolAccess');

// ─── normalizeEmail ──────────────────────────────────────────────
describe('normalizeEmail', () => {
  test('trims and lowercases', () => {
    assert.equal(normalizeEmail('  Huy.Le@H2XStudio.com  '), 'huy.le@h2xstudio.com');
  });
  test('handles missing/empty email', () => {
    assert.equal(normalizeEmail(undefined), '');
    assert.equal(normalizeEmail(''), '');
    assert.equal(normalizeEmail('   '), '');
  });
});

// ─── normalizeToolAccess ─────────────────────────────────────────
describe('normalizeToolAccess', () => {
  test('legacy tools array maps fee-generator -> feeProposalGenerator', () => {
    const access = normalizeToolAccess({ tools: ['fee-generator'] });
    assert.deepEqual(access, {
      feeProposalGenerator: true,
      ffeLooseCalculation: false,
      h2xQsApp: false,
    });
  });
  test('new toolAccess object read directly', () => {
    const access = normalizeToolAccess({ toolAccess: { ffeLooseCalculation: true } });
    assert.equal(access.ffeLooseCalculation, true);
    assert.equal(access.feeProposalGenerator, false);
  });
  test('missing/empty user data -> all false (default-deny)', () => {
    assert.deepEqual(normalizeToolAccess({}), {
      feeProposalGenerator: false,
      ffeLooseCalculation: false,
      h2xQsApp: false,
    });
    assert.deepEqual(normalizeToolAccess(undefined), {
      feeProposalGenerator: false,
      ffeLooseCalculation: false,
      h2xQsApp: false,
    });
  });
  test('legacy + new both present -> union', () => {
    const access = normalizeToolAccess({ tools: ['fee-generator'], toolAccess: { h2xQsApp: true } });
    assert.deepEqual(access, {
      feeProposalGenerator: true,
      ffeLooseCalculation: false,
      h2xQsApp: true,
    });
  });
  test('unknown legacy id / non-boolean toolAccess values are ignored, not thrown', () => {
    const access = normalizeToolAccess({ tools: ['some-other-tool'], toolAccess: { feeProposalGenerator: 'yes' } });
    assert.deepEqual(access, {
      feeProposalGenerator: false,
      ffeLooseCalculation: false,
      h2xQsApp: false,
    });
  });
});

// ─── decodeFirestoreFields (REST document -> plain object) ───────
describe('decodeFirestoreFields', () => {
  test('decodes string, boolean, array and nested map values', () => {
    const fields = {
      name: { stringValue: 'Huy Le' },
      active: { booleanValue: false },
      tools: { arrayValue: { values: [{ stringValue: 'fee-generator' }] } },
      toolAccess: {
        mapValue: {
          fields: {
            feeProposalGenerator: { booleanValue: true },
            h2xQsApp: { booleanValue: false },
          },
        },
      },
    };
    assert.deepEqual(decodeFirestoreFields(fields), {
      name: 'Huy Le',
      active: false,
      tools: ['fee-generator'],
      toolAccess: { feeProposalGenerator: true, h2xQsApp: false },
    });
  });
});

// ─── makeFetchOwnUserDoc (mocks global.fetch — no real network call) ──
describe('makeFetchOwnUserDoc', () => {
  let originalFetch;
  beforeEach(() => { originalFetch = global.fetch; });
  afterEach(() => { global.fetch = originalFetch; });

  test('404 -> { exists: false, data: null }', async () => {
    global.fetch = async () => ({ status: 404, ok: false });
    const fetchDoc = makeFetchOwnUserDoc('h2x-tools-auth');
    const result = await fetchDoc('user@h2xstudio.com', 'tok');
    assert.deepEqual(result, { exists: false, data: null });
  });

  test('200 -> exists true, data decoded', async () => {
    global.fetch = async () => ({
      status: 200,
      ok: true,
      json: async () => ({ fields: { active: { booleanValue: true } } }),
    });
    const fetchDoc = makeFetchOwnUserDoc('h2x-tools-auth');
    const result = await fetchDoc('user@h2xstudio.com', 'tok');
    assert.deepEqual(result, { exists: true, data: { active: true } });
  });

  test('non-ok, non-404 status throws (e.g. malformed/rejected token)', async () => {
    global.fetch = async () => ({ status: 403, ok: false });
    const fetchDoc = makeFetchOwnUserDoc('h2x-tools-auth');
    await assert.rejects(() => fetchDoc('user@h2xstudio.com', 'tok'));
  });

  test('requests the exact users/{email} doc path with the caller bearer token', async () => {
    let capturedUrl, capturedHeaders;
    global.fetch = async (url, opts) => {
      capturedUrl = url;
      capturedHeaders = opts.headers;
      return { status: 404, ok: false };
    };
    const fetchDoc = makeFetchOwnUserDoc('h2x-tools-auth');
    await fetchDoc('user@h2xstudio.com', 'the-token');
    assert.equal(
      capturedUrl,
      'https://firestore.googleapis.com/v1/projects/h2x-tools-auth/databases/(default)/documents/users/user%40h2xstudio.com'
    );
    assert.equal(capturedHeaders.Authorization, 'Bearer the-token');
  });
});

// ─── resolveToolAccess (item 1 + item 4 core rules) ───────────────
describe('resolveToolAccess', () => {
  test('document does not exist -> 403, DENY_NO_ACCESS', async () => {
    const fetchDoc = async () => ({ exists: false, data: null });
    const r = await resolveToolAccess(fetchDoc, 'x@h2xstudio.com', 'tok');
    assert.deepEqual(r, { ok: false, status: 403, error: DENY_NO_ACCESS });
  });

  test('active === false -> 403, DENY_INACTIVE, even if toolAccess would allow', async () => {
    const fetchDoc = async () => ({
      exists: true,
      data: { active: false, toolAccess: { feeProposalGenerator: true } },
    });
    const r = await resolveToolAccess(fetchDoc, 'x@h2xstudio.com', 'tok');
    assert.deepEqual(r, { ok: false, status: 403, error: DENY_INACTIVE });
  });

  test('active missing (legacy doc, field never set) is treated as active', async () => {
    const fetchDoc = async () => ({ exists: true, data: { tools: ['fee-generator'] } });
    const r = await resolveToolAccess(fetchDoc, 'x@h2xstudio.com', 'tok');
    assert.equal(r.ok, true);
    assert.equal(r.access.feeProposalGenerator, true);
  });

  test('active true + toolAccess -> ok, normalized access returned', async () => {
    const fetchDoc = async () => ({
      exists: true,
      data: { active: true, toolAccess: { ffeLooseCalculation: true } },
    });
    const r = await resolveToolAccess(fetchDoc, 'x@h2xstudio.com', 'tok');
    assert.deepEqual(r, {
      ok: true,
      access: { feeProposalGenerator: false, ffeLooseCalculation: true, h2xQsApp: false },
    });
  });
});

// ─── makeRequireToolAccess middleware (item 1) ────────────────────
function fakeRes() {
  const res = { statusCode: null, body: null };
  res.status = (code) => { res.statusCode = code; return res; };
  res.json = (body) => { res.body = body; return res; };
  return res;
}

describe('requireToolAccess middleware', () => {
  test('doc missing -> 403, next NOT called', async () => {
    const fetchDoc = async () => ({ exists: false, data: null });
    const requireToolAccess = makeRequireToolAccess(fetchDoc);
    const req = { userEmail: 'x@h2xstudio.com', rawIdToken: 'tok' };
    const res = fakeRes();
    let nextCalled = false;
    await requireToolAccess('feeProposalGenerator')(req, res, () => { nextCalled = true; });
    assert.equal(res.statusCode, 403);
    assert.equal(nextCalled, false);
  });

  test('active === false -> 403, next NOT called', async () => {
    const fetchDoc = async () => ({ exists: true, data: { active: false, toolAccess: { feeProposalGenerator: true } } });
    const requireToolAccess = makeRequireToolAccess(fetchDoc);
    const req = { userEmail: 'x@h2xstudio.com', rawIdToken: 'tok' };
    const res = fakeRes();
    let nextCalled = false;
    await requireToolAccess('feeProposalGenerator')(req, res, () => { nextCalled = true; });
    assert.equal(res.statusCode, 403);
    assert.equal(nextCalled, false);
  });

  test('toolAccess[toolKey] !== true -> 403, next NOT called', async () => {
    const fetchDoc = async () => ({ exists: true, data: { toolAccess: { ffeLooseCalculation: true } } });
    const requireToolAccess = makeRequireToolAccess(fetchDoc);
    const req = { userEmail: 'x@h2xstudio.com', rawIdToken: 'tok' };
    const res = fakeRes();
    let nextCalled = false;
    await requireToolAccess('feeProposalGenerator')(req, res, () => { nextCalled = true; });
    assert.equal(res.statusCode, 403);
    assert.equal(nextCalled, false);
  });

  test('has access -> next() called, no response written', async () => {
    const fetchDoc = async () => ({ exists: true, data: { toolAccess: { feeProposalGenerator: true } } });
    const requireToolAccess = makeRequireToolAccess(fetchDoc);
    const req = { userEmail: 'x@h2xstudio.com', rawIdToken: 'tok' };
    const res = fakeRes();
    let nextCalled = false;
    await requireToolAccess('feeProposalGenerator')(req, res, () => { nextCalled = true; });
    assert.equal(nextCalled, true);
    assert.equal(res.statusCode, null);
  });

  test('fetchOwnUserDoc throwing (network/Firestore error) -> 502, next NOT called', async () => {
    const fetchDoc = async () => { throw new Error('Firestore lookup failed: 500'); };
    const requireToolAccess = makeRequireToolAccess(fetchDoc);
    const req = { userEmail: 'x@h2xstudio.com', rawIdToken: 'tok' };
    const res = fakeRes();
    let nextCalled = false;
    await requireToolAccess('feeProposalGenerator')(req, res, () => { nextCalled = true; });
    assert.equal(res.statusCode, 502);
    assert.equal(nextCalled, false);
  });
});
