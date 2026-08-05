const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const { getAccountLabel, getAvatarInitial } = require('../public/js/account');

// ─── getAccountLabel ─────────────────────────────────────────────
describe('getAccountLabel', () => {
  test('prefers displayName over email', () => {
    assert.equal(getAccountLabel({ displayName: 'Huy Le', email: 'huy@h2xstudio.com' }), 'Huy Le');
  });
  test('falls back to email when displayName missing/empty', () => {
    assert.equal(getAccountLabel({ displayName: '', email: 'huy@h2xstudio.com' }), 'huy@h2xstudio.com');
    assert.equal(getAccountLabel({ email: 'huy@h2xstudio.com' }), 'huy@h2xstudio.com');
  });
  test('missing/null user -> empty string, never throws', () => {
    assert.equal(getAccountLabel(null), '');
    assert.equal(getAccountLabel(undefined), '');
    assert.equal(getAccountLabel({}), '');
  });
});

// ─── getAvatarInitial ────────────────────────────────────────────
describe('getAvatarInitial', () => {
  test('uppercases the first character of the label', () => {
    assert.equal(getAvatarInitial('huy le'), 'H');
    assert.equal(getAvatarInitial('Huy Le'), 'H');
  });
  test('trims leading whitespace before taking the initial', () => {
    assert.equal(getAvatarInitial('  huy@h2xstudio.com'), 'H');
  });
  test('empty/missing label -> "?" placeholder, never throws', () => {
    assert.equal(getAvatarInitial(''), '?');
    assert.equal(getAvatarInitial('   '), '?');
    assert.equal(getAvatarInitial(undefined), '?');
    assert.equal(getAvatarInitial(null), '?');
  });
});
