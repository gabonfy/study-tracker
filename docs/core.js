/*
 * Study Tracker – pure logic (no Office APIs), so it can be unit-tested in Node.
 *
 * ACTIVE TIME RULE
 *   Every "input" (a highlight added/removed/changed, or text typed) opens a window of
 *   `idleWindowSec`. You count as active from the input until the next input or until the
 *   window runs out, whichever comes first. So each input credits min(gap to next input, window).
 *   Reading for 1 minute between two highlights counts fully; leaving for lunch counts only
 *   one window after your last highlight.
 *
 * PROGRESS RULE
 *   progress = words up to the end of your furthest highlight / total words in the document.
 *
 * ESTIMATE RULE
 *   pace = (progress now − progress when tracking started) / active time tracked
 *   remaining active time = (1 − progress) / pace
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.StudyCore = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var SESSION_GAP_MS = 30 * 60 * 1000; // > 30 min without input starts a new session
  var MAX_SESSIONS = 500;

  function countWords(text) {
    if (!text) return 0;
    var n = 0;
    var parts = String(text).split(/[\s  -​　]+/);
    for (var i = 0; i < parts.length; i++) {
      if (/[0-9A-Za-zÀ-ɏͰ-ϿЀ-ӿ֐-ۿ぀-ヿ一-鿿]/.test(parts[i])) n++;
    }
    return n;
  }

  function hash(str) {
    // FNV-1a 32-bit, as hex
    var h = 0x811c9dc5;
    for (var i = 0; i < str.length; i++) {
      h ^= str.charCodeAt(i);
      h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
    }
    return ("0000000" + h.toString(16)).slice(-8);
  }

  function newState() {
    return {
      v: 1,
      idleWindowSec: 120,
      colorFilter: "any",
      baselineProgress: null, // progress when tracking began on this file
      committedMs: 0,         // active time already credited (excludes the open window)
      lastInputAt: null,
      inputs: 0,
      sessions: [],           // {start, end, activeMs, inputs, pStart, pEnd}
      createdAt: Date.now()
    };
  }

  function normalizeState(s) {
    var d = newState();
    if (!s || typeof s !== "object") return d;
    for (var k in d) if (!(k in s)) s[k] = d[k];
    if (!Array.isArray(s.sessions)) s.sessions = [];
    return s;
  }

  function windowMs(state) {
    return Math.max(5, Number(state.idleWindowSec) || 120) * 1000;
  }

  function lastSession(state) {
    return state.sessions.length ? state.sessions[state.sessions.length - 1] : null;
  }

  /** Register an input at time t (ms). `progress` (0..1 or null) is the current progress. */
  function recordInput(state, t, progress) {
    var W = windowMs(state);
    var prev = state.lastInputAt;
    var cur = lastSession(state);
    if (prev != null && t >= prev) {
      var credit = Math.min(t - prev, W);
      state.committedMs += credit;
      if (cur) {
        cur.activeMs += credit;
        cur.end = Math.max(cur.end, prev + credit);
      }
    }
    var startNew = !cur || prev == null || t - prev > SESSION_GAP_MS || !sameDay(prev, t);
    if (startNew) {
      cur = { start: t, end: t, activeMs: 0, inputs: 0, pStart: progress, pEnd: progress };
      state.sessions.push(cur);
      if (state.sessions.length > MAX_SESSIONS) state.sessions.splice(0, state.sessions.length - MAX_SESSIONS);
    }
    cur.inputs += 1;
    cur.end = Math.max(cur.end, t);
    if (progress != null) {
      if (cur.pStart == null) cur.pStart = progress;
      cur.pEnd = progress;
    }
    state.inputs += 1;
    state.lastInputAt = t;
    return state;
  }

  /** Time still "open" after the last input (0..window). */
  function pendingMs(state, now) {
    if (state.lastInputAt == null) return 0;
    return Math.max(0, Math.min(now - state.lastInputAt, windowMs(state)));
  }

  function isActive(state, now) {
    return state.lastInputAt != null && now - state.lastInputAt < windowMs(state);
  }

  function totalActiveMs(state, now) {
    return state.committedMs + pendingMs(state, now);
  }

  function sessionActiveMs(state, session, now) {
    var extra = session === lastSession(state) ? pendingMs(state, now) : 0;
    return session.activeMs + extra;
  }

  /** Active time today (local time). */
  function todayActiveMs(state, now) {
    var ms = 0;
    for (var i = 0; i < state.sessions.length; i++) {
      var s = state.sessions[i];
      if (sameDay(s.start, now)) ms += sessionActiveMs(state, s, now);
    }
    return ms;
  }

  /** Current session if still running (last input within the session gap). */
  function currentSession(state, now) {
    var s = lastSession(state);
    if (!s || state.lastInputAt == null) return null;
    if (now - state.lastInputAt > SESSION_GAP_MS || !sameDay(state.lastInputAt, now)) return null;
    return s;
  }

  function updateProgress(state, progress) {
    if (progress == null) return;
    if (state.baselineProgress == null) state.baselineProgress = progress;
    var s = lastSession(state);
    if (s && state.lastInputAt != null && s.end >= state.lastInputAt - 1) s.pEnd = progress;
  }

  /**
   * Progress from a per-paragraph word list.
   * paraWords: words per paragraph; lastIdx: index of the paragraph holding the furthest
   * highlight (-1 for none); wordsInLast: words of that paragraph up to the end of the highlight.
   */
  function computeProgress(paraWords, lastIdx, wordsInLast) {
    var total = 0, before = 0;
    for (var i = 0; i < paraWords.length; i++) {
      total += paraWords[i];
      if (i < lastIdx) before += paraWords[i];
    }
    if (total === 0) return { total: 0, reached: 0, progress: 0 };
    var reached = lastIdx < 0 ? 0 : Math.min(total, before + Math.max(0, wordsInLast || 0));
    return { total: total, reached: reached, progress: reached / total };
  }

  var MIN_ACTIVE_FOR_ESTIMATE = 3 * 60 * 1000;
  var MIN_GAIN_FOR_ESTIMATE = 0.005;

  /** Estimate remaining active time. */
  function estimate(state, progress, now) {
    var active = totalActiveMs(state, now);
    if (progress == null) return { ok: false, reason: "no-progress" };
    if (progress >= 0.9995) return { ok: true, done: true, remainingMs: 0, pacePerHour: null };
    var base = state.baselineProgress == null ? progress : state.baselineProgress;
    var gained = progress - base;
    if (active < MIN_ACTIVE_FOR_ESTIMATE || gained < MIN_GAIN_FOR_ESTIMATE) {
      return {
        ok: false, reason: "need-data",
        needMs: Math.max(0, MIN_ACTIVE_FOR_ESTIMATE - active),
        needGain: Math.max(0, MIN_GAIN_FOR_ESTIMATE - gained)
      };
    }
    var pacePerMs = gained / active;
    return {
      ok: true, done: false,
      remainingMs: (1 - progress) / pacePerMs,
      pacePerHour: pacePerMs * 3600000,
      gained: gained, activeMs: active
    };
  }

  function sameDay(a, b) {
    var x = new Date(a), y = new Date(b);
    return x.getFullYear() === y.getFullYear() && x.getMonth() === y.getMonth() && x.getDate() === y.getDate();
  }

  function formatDuration(ms, opts) {
    ms = Math.max(0, Math.round(ms / 1000)) * 1000;
    var s = Math.floor(ms / 1000);
    var h = Math.floor(s / 3600), m = Math.floor((s % 3600) / 60), sec = s % 60;
    if (opts && opts.clock) {
      var p = function (n) { return (n < 10 ? "0" : "") + n; };
      return (h ? h + ":" + p(m) : m) + ":" + p(sec);
    }
    if (h) return h + " h " + (m ? m + " min" : "");
    if (m) return m + " min";
    return sec + " s";
  }

  function toCSV(state, now) {
    var rows = [["date", "start", "end", "active_minutes", "highlights_and_edits", "progress_start_pct", "progress_end_pct"]];
    var pad = function (n) { return (n < 10 ? "0" : "") + n; };
    state.sessions.forEach(function (s) {
      var a = new Date(s.start), b = new Date(s.end);
      rows.push([
        a.getFullYear() + "-" + pad(a.getMonth() + 1) + "-" + pad(a.getDate()),
        pad(a.getHours()) + ":" + pad(a.getMinutes()),
        pad(b.getHours()) + ":" + pad(b.getMinutes()),
        (sessionActiveMs(state, s, now) / 60000).toFixed(1),
        s.inputs,
        s.pStart == null ? "" : (s.pStart * 100).toFixed(1),
        s.pEnd == null ? "" : (s.pEnd * 100).toFixed(1)
      ]);
    });
    return rows.map(function (r) { return r.join(","); }).join("\n");
  }

  // Word's highlight colours, name -> hex (Office.js returns either form depending on platform)
  var HIGHLIGHTS = {
    yellow: "#FFFF00", lime: "#00FF00", brightgreen: "#00FF00", turquoise: "#00FFFF", pink: "#FF00FF",
    blue: "#0000FF", red: "#FF0000", darkblue: "#000080", teal: "#008080", green: "#008000",
    purple: "#800080", violet: "#800080", darkred: "#800000", olive: "#808000", darkyellow: "#808000",
    gray: "#808080", gray50: "#808080", lightgray: "#C0C0C0", gray25: "#C0C0C0", black: "#000000"
  };

  function normColor(c) {
    if (c == null) return null;
    c = String(c).trim();
    if (c === "") return "";
    if (c[0] === "#") return c.toUpperCase();
    var k = c.toLowerCase().replace(/[\s_-]/g, "");
    return HIGHLIGHTS[k] || c.toUpperCase();
  }

  /**
   * Classify a highlight value for a range.
   *  0 = no highlight, 1 = fully highlighted in a counted colour,
   *  2 = mixed (needs a closer look), 3 = highlighted, but in a colour that isn't counted.
   */
  function classifyHighlight(value, filter) {
    var c = normColor(value);
    if (c == null || c === "#FFFFFF" || c === "NONE" || c === "NOHIGHLIGHT") return 0;
    if (c === "") return 2;
    if (!filter || filter === "any") return 1;
    return c === normColor(filter) ? 1 : 3;
  }

  return {
    SESSION_GAP_MS: SESSION_GAP_MS,
    countWords: countWords, hash: hash, newState: newState, normalizeState: normalizeState,
    recordInput: recordInput, pendingMs: pendingMs, isActive: isActive,
    totalActiveMs: totalActiveMs, sessionActiveMs: sessionActiveMs, todayActiveMs: todayActiveMs,
    currentSession: currentSession, updateProgress: updateProgress, computeProgress: computeProgress,
    estimate: estimate, formatDuration: formatDuration, toCSV: toCSV,
    normColor: normColor, classifyHighlight: classifyHighlight, sameDay: sameDay
  };
});
