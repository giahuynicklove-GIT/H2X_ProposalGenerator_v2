// Pure helpers for the header account menu — no DOM, no Firebase, so they
// can run unmodified in the browser (loaded as a plain <script>) and in
// Node test files (via require/module.exports).
(function (root, factory) {
  const exportsObj = factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = exportsObj;
  }
  if (root) {
    Object.assign(root, exportsObj);
  }
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : null), function () {
  function getAccountLabel(user) {
    return (user && (user.displayName || user.email)) || '';
  }

  function getAvatarInitial(label) {
    const trimmed = (label || '').trim();
    return trimmed ? trimmed.charAt(0).toUpperCase() : '?';
  }

  return { getAccountLabel, getAvatarInitial };
});
