// Shared tool-access logic for server.js. Kept framework-free (no Express, no
// global fetch call at module scope) so it can be unit-tested in isolation.
// Client copies of normalizeToolAccess (public/index.html, public/admin.html)
// must be kept in sync with ALL_TOOL_KEYS / LEGACY_TOOL_MAP below.

const ALL_TOOL_KEYS = ['feeProposalGenerator', 'ffeLooseCalculation', 'h2xQsApp'];
const LEGACY_TOOL_MAP = { 'fee-generator': 'feeProposalGenerator' };

function normalizeEmail(email) {
  return (email || '').trim().toLowerCase();
}

function normalizeToolAccess(userData = {}) {
  const access = Object.fromEntries(ALL_TOOL_KEYS.map(k => [k, false]));
  if (userData.toolAccess) {
    for (const k of ALL_TOOL_KEYS) {
      if (userData.toolAccess[k] === true) access[k] = true;
    }
  }
  if (Array.isArray(userData.tools)) {
    for (const legacyId of userData.tools) {
      const mapped = LEGACY_TOOL_MAP[legacyId];
      if (mapped) access[mapped] = true;
    }
  }
  return access;
}

function decodeFirestoreValue(v) {
  if (!v) return null;
  if (v.stringValue !== undefined) return v.stringValue;
  if (v.booleanValue !== undefined) return v.booleanValue;
  if (v.integerValue !== undefined) return parseInt(v.integerValue, 10);
  if (v.doubleValue !== undefined) return v.doubleValue;
  if (v.arrayValue !== undefined) return (v.arrayValue.values || []).map(decodeFirestoreValue);
  if (v.mapValue !== undefined) return decodeFirestoreFields(v.mapValue.fields || {});
  return null;
}
function decodeFirestoreFields(fields) {
  const out = {};
  for (const [k, v] of Object.entries(fields)) out[k] = decodeFirestoreValue(v);
  return out;
}

// Reads users/{email} via Firestore REST using the caller's own ID token —
// satisfies the existing rule (email == userEmail) without needing firebase-admin.
function makeFetchOwnUserDoc(projectId) {
  return async function fetchOwnUserDoc(email, idToken) {
    const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/users/${encodeURIComponent(email)}`;
    const r = await fetch(url, { headers: { Authorization: `Bearer ${idToken}` } });
    if (r.status === 404) return { exists: false, data: null };
    if (!r.ok) throw new Error(`Firestore lookup failed: ${r.status}`);
    const doc = await r.json();
    return { exists: true, data: doc.fields ? decodeFirestoreFields(doc.fields) : {} };
  };
}

const DENY_NO_ACCESS = 'Chưa được cấp quyền dùng tool này. Liên hệ quản trị viên H2X.';
const DENY_INACTIVE = 'Tài khoản đã bị vô hiệu hoá. Liên hệ quản trị viên H2X.';

// Single choke point used by both requireToolAccess and GET /api/tool-access,
// so "doc missing" / "active === false" / "toolAccess[key] !== true" are each
// decided in exactly one place.
async function resolveToolAccess(fetchOwnUserDoc, email, idToken) {
  const result = await fetchOwnUserDoc(email, idToken);
  if (!result.exists) return { ok: false, status: 403, error: DENY_NO_ACCESS };
  if (result.data.active === false) return { ok: false, status: 403, error: DENY_INACTIVE };
  return { ok: true, access: normalizeToolAccess(result.data) };
}

function makeRequireToolAccess(fetchOwnUserDoc) {
  return function requireToolAccess(toolKey) {
    return async function (req, res, next) {
      try {
        const resolved = await resolveToolAccess(fetchOwnUserDoc, req.userEmail, req.rawIdToken);
        if (!resolved.ok) return res.status(resolved.status).json({ error: resolved.error });
        if (resolved.access[toolKey] !== true) {
          return res.status(403).json({ error: DENY_NO_ACCESS });
        }
        next();
      } catch (e) {
        console.error('requireToolAccess error:', e.message);
        res.status(502).json({ error: 'Không kiểm tra được quyền truy cập, thử lại sau.' });
      }
    };
  };
}

module.exports = {
  ALL_TOOL_KEYS,
  LEGACY_TOOL_MAP,
  normalizeEmail,
  normalizeToolAccess,
  decodeFirestoreFields,
  makeFetchOwnUserDoc,
  resolveToolAccess,
  makeRequireToolAccess,
  DENY_NO_ACCESS,
  DENY_INACTIVE,
};
