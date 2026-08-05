# H2X Shared Access System

One Firebase project (`h2x-tools-auth`) gates every H2X internal web app
(Fee Proposal Generator, FF&E Loose Calculation, Báo cáo Khái toán tự động,
...). This doc describes how that works end to end and how to add a new app
to it. It covers this repo's role as **the app itself** (`server.js`,
`public/index.html`) and **the shared Access Manager** (`public/admin.html`),
which any H2X admin uses to manage access for *all* apps, not just this one.

## 1. Data model — `users/{email}` in Firestore

One doc per user, keyed by lowercased email:

```jsonc
{
  "name": "Huy Le",                 // optional, display only
  "active": true,                   // false = locked out of every app. Missing/undefined = active (legacy docs predate this field)
  "toolAccess": {                   // current format — one boolean per catalog app key
    "feeProposalGenerator": true,
    "ffeLooseCalculation": false,
    "h2xQsApp": false
  },
  "tools": ["fee-generator"],       // LEGACY format, pre-dates toolAccess. Never written by the Access Manager, but still read.
  "updatedAt": "<Firestore Timestamp>" // set by the Access Manager on every save/lock/unlock
}
```

**Legacy compatibility (`tools: ["fee-generator"]`).** Old docs only have the
`tools` array. `normalizeToolAccess()` unions both formats on every read: it
starts from `toolAccess`, then walks `tools` and maps each legacy id through
`LEGACY_TOOL_MAP` (e.g. `"fee-generator"` → `feeProposalGenerator`) onto the
same access object. The Access Manager never deletes the `tools` field — it
only ever writes `toolAccess` (via `set(..., {merge:true})`), so a legacy doc
keeps working indefinitely even if nobody edits it again. If an admin edits
and saves such a user, the resulting checked boxes (computed from the union)
get persisted into `toolAccess`, but the old `tools` array is left in place,
untouched.

**Default-deny.** Any app key not explicitly `true` in `toolAccess` (and not
covered by a legacy `tools` mapping) is `false`. This is what makes item
"new app defaults to false for existing users" automatic: add a new entry to
the catalog (§2) and every existing user doc is simply missing that key, so
`normalizeToolAccess()` returns `false` for it with no migration needed.

## 2. `public/tool-catalog.json` — the single source of truth for "which apps exist"

```jsonc
{
  "apps": [
    { "key": "feeProposalGenerator", "label": "Fee Proposal Generator", "legacyIds": ["fee-generator"] },
    { "key": "ffeLooseCalculation",  "label": "FF&E Loose Calculation",  "legacyIds": [] },
    { "key": "h2xQsApp",             "label": "Báo cáo Khái toán tự động", "legacyIds": [] }
  ]
}
```

- `key` — the field name used inside `toolAccess` and passed to
  `requireToolAccess(key)` on the server that owns that app.
- `label` — what admins see in the Access Manager (dashboard breakdown,
  checkboxes, badges, filter dropdown).
- `legacyIds` — old string ids (from the pre-`toolAccess` era) that should
  map onto this key. Empty array if the app never had a legacy id.

**To add a new app to the ecosystem:** add one object to `apps[]`. Nothing
else needs editing —

- `lib/toolAccess.js` derives `ALL_TOOL_KEYS` / `LEGACY_TOOL_MAP` from this
  file at require-time (`require('../public/tool-catalog.json')`), so every
  server that shares this library picks the new app up automatically.
- `public/admin.html` fetches `/tool-catalog.json` at load time and renders
  the add/edit checkboxes, the app filter dropdown, and the dashboard's
  per-app breakdown from it — no hardcoded app list to update.
- Existing user docs automatically get `false` for the new key (§1).

This file is served as a static asset (`express.static('public')` in
`server.js`), so it's reachable at `/tool-catalog.json` from any app that
serves this `public/` folder, and can be fetched cross-origin by a
*different* app's frontend if that app wants the same catalog (CORS
permitting) instead of vendoring its own copy.

## 3. How access is enforced

**Server-side (authoritative).** `server.js`:
1. `requireAuth` verifies the `Authorization: Bearer <Firebase ID token>`
   header's signature against Google's public keys (no service account
   needed) and sets `req.userEmail` / `req.rawIdToken` from the *verified*
   token — never from the request body.
2. `requireToolAccess('feeProposalGenerator')` (from `lib/toolAccess.js`)
   fetches `users/{email}` via the Firestore REST API using the caller's own
   ID token, applies `active !== false` and `toolAccess[key] === true`, and
   returns 403 if either check fails, 401 upstream if there's no/invalid
   token at all.
3. Every protected route (`/api/research`, `/api/generate`) is wired through
   both middlewares. A signed-out or de-authorized user gets 401/403 on the
   *next* request no matter what the client UI shows.

**Client-side (UX only, not a security boundary).** `public/index.html`
gates the UI behind the same check (`GET` the user's own Firestore doc,
compute `normalizeToolAccess()`, show the auth gate if denied) purely so a
locked-out user sees a clear message instead of a broken app — the real
enforcement is always the server check above.

## 4. H2X Access Manager (`public/admin.html`)

A single admin console for every app's access, not per-app. Restricted to
`ADMIN_EMAILS` (hardcoded in `admin.html`, must match the Firestore rules'
admin list — see §5).

- **Dashboard** — total users, active count, locked count, app count (from
  the catalog), and a per-app user-count breakdown. Computed client-side
  from the full `users` collection snapshot already loaded for the table.
- **Search** — filters the table by email or name substring, client-side.
- **Filters** — by app (only users with that app's `toolAccess[key] === true`)
  and by status (active / locked).
- **User list** — email + name, one badge per granted app, status badge,
  `updatedAt` (formatted, or `—` if the doc predates that field), and
  row actions: Sửa (edit), Khóa/Mở (toggle `active`), Xóa (delete the doc
  entirely).
- **Add/edit form** — renders one checkbox per catalog app (§2), so it never
  needs a code change when an app is added or renamed.
- **Self-protection** — an admin can't lock or delete their *own* row: the
  buttons are disabled in that row, and `toggleLock` / `deleteUser` also
  refuse at the top of the function (defense in depth against a stale
  render). Editing your own name/app access is still allowed — only the
  destructive "lock the whole account" / "delete the doc" actions are
  blocked for self.

## 5. Firestore security rules

`users/{email}` rules must allow:
- any signed-in user to **read** their own doc (`request.auth.token.email == userEmail`) — this is what the client-side gate and every app's own access check rely on.
- **admins only** (matching `ADMIN_EMAILS` in `admin.html`) to **read/write any** doc — this is what the Access Manager needs for the dashboard (reads every doc), search/filter (reads every doc), and add/edit/lock/delete (writes any doc).

⚠️ This repo currently has two rules files that disagree —
`firestore.rules` denies *all* writes (`allow write: if false`), which would
break every Access Manager write (save/lock/delete) if it's the version
actually deployed to the Firebase console. `firestore.rules.txt` has the
admin read/write clause the Access Manager needs. **Before relying on the
Access Manager in production, confirm which version is deployed in the
Firebase console → Firestore → Rules, and deploy the admin-enabled version
if it isn't already live.** This is a manual step outside this repo/CLI
push — Firestore rules are deployed independently of the app itself.

## 6. Account menu / sign-out (already shipped, unrelated to this doc's scope)

`public/index.html`'s header account menu (email/displayName + "Đăng xuất")
and its sign-out flow are a separate, already-completed feature — see the
git history (`Add account menu and sign out`) and `test/account.test.js` /
`public/js/account.js`. Nothing in this doc changes that behavior.
