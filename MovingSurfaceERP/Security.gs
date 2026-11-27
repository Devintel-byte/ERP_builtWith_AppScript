/**
 * Two protections that matter once the app can be reached over a plain web API.
 *
 * 1. A limit on wrong PINs. A PIN is four digits, so unlimited guessing finds it.
 * 2. An owner-only guard on setup and maintenance functions. Any function without a
 *    trailing underscore can be called from the browser console of a page served by
 *    Apps Script, so these check who is really running them.
 */

/*************** WRONG-PIN LIMIT ***************/

const PIN_MAX_FAILURES = 5;
const PIN_LOCK_SECONDS = 300;
const PIN_FAILURE_SLOTS = 100;
const PIN_LOCKED_MESSAGE = 'Too many wrong PIN attempts for this Staff ID. Wait 5 minutes and try again, or use account recovery.';

/** One counter per person, shared by attendance, the staff portal and the admin ERP. */
function pinSubject_(staffId) {
  return String(staffId || '').trim().toUpperCase() || '(NO-STAFF-ID)';
}

/*
 * Each failure takes its own cache slot instead of incrementing one number: two
 * guesses arriving together would both read the old number and one would be lost.
 * The count is the number of occupied slots, and each slot expires on its own.
 */
function pinFailureKeys_(subject) {
  const base = 'PINFAIL_' + sha256Hex_(pinSubject_(subject)).slice(0, 32) + '_';
  const keys = [];
  for (let i = 0; i < PIN_FAILURE_SLOTS; i++) keys.push(base + i);
  return keys;
}

function countPinFailures_(subject) {
  const found = CacheService.getScriptCache().getAll(pinFailureKeys_(subject)) || {};
  return Object.keys(found).length;
}

function isPinLocked_(subject) {
  return countPinFailures_(subject) >= PIN_MAX_FAILURES;
}

function recordPinFailure_(subject) {
  const cache = CacheService.getScriptCache();
  const keys = pinFailureKeys_(subject);
  const used = cache.getAll(keys) || {};
  const free = keys.filter(key => !Object.prototype.hasOwnProperty.call(used, key));
  if (!free.length) return;
  cache.put(free[Math.floor(Math.random() * free.length)], '1', PIN_LOCK_SECONDS);
}

function clearPinFailures_(subject) {
  const cache = CacheService.getScriptCache();
  const used = Object.keys(cache.getAll(pinFailureKeys_(subject)) || {});
  if (used.length) cache.removeAll(used);
}

/*************** OWNER-ONLY FUNCTIONS ***************/

/**
 * Stops anyone but the person actually running the script. From the editor or the
 * spreadsheet menu the active user and the effective user are the same account.
 * An anonymous web visitor has no active user, and a signed-in visitor is not the
 * account the web app executes as.
 */
function requireOwnerContext_() {
  const active = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const effective = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if (!active || active !== effective) {
    throw new Error('Run this from the Apps Script editor or the spreadsheet menu, signed in as an editor of the project.');
  }
}
