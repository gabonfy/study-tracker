/*
 * Study Tracker – pure logic (no Office APIs), so it can be unit-tested in Node.
 *
 * ACTIVE TIME RULE
 *   Two things make time count, and overlapping time is never counted twice:
 *   1. Automatic: every input (a highlight added/removed/recoloured, or text typed/deleted) opens a
 *      window of `idleWindowSec`. You count as active until the next input or until the window
 *      runs out. So each input credits min(gap to next input, window).
 *   2. Manual: between pressing "Start clock" and "Stop clock" all time counts, whatever you do.
 *      If the clock is left running with no highlight or typing for MANUAL_MAX_MS it stops itself.
 *
 * PROGRESS RULE
 *   progress = furthest point reached (end of your furthest highlight, or the furthest place you
 *   typed, whichever is further) / total words in the document.
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

  var SESSION_GAP_MS = 30 * 60 * 1000;      // > 30 min with nothing counted starts a new session
  var MANUAL_MAX_MS = 2 * 60 * 60 * 1000;   // a forgotten manual clock stops after 2 h without input
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
      v: 2,
      idleWindowSec: 120,
      colorFilter: "any",
      notifyOnStop: true,
      baselineProgress: null, // progress when tracking began on this file
      committedMs: 0,         // active time credited up to `countedTo`
      countedTo: null,        // timestamp up to which time has been accounted
      lastInputAt: null,      // last highlight or edit
      lastCoveredAt: null,    // last moment that counted as active
      stoppedAt: null,        // last press of "Stop clock" (closes any open window)
      manualSince: null,      // set while the manual clock runs
      heartbeatAt: null,      // last sign of life while the manual clock runs
      editMark: null,         // {words, at}: furthest point reached by typing
      inputs: 0,
      sessions: [],           // {start, end, activeMs, inputs, pStart, pEnd}
      createdAt: Date.now(),
      savedAt: 0
    };
  }

  function normalizeState(s) {
    var d = newState();
    if (!s || typeof s !== "object") return d;
    var wasV1 = !s.v || s.v < 2;
    for (var k in d) if (!(k in s)) s[k] = d[k];
    if (!Array.isArray(s.sessions)) s.sessions = [];
    if (wasV1) {
      // v1 credited everything up to the last input; the window after it was still open.
      s.countedTo = s.lastInputAt;
      s.lastCoveredAt = s.lastInputAt;
      delete s.paused;
      s.v = 2;
    }
    return s;
  }

  function windowMs(state) {
    return Math.max(5, Number(state.idleWindowSec) || 120) * 1000;
  }

  function lastSession(state) {
    return state.sessions.length ? state.sessions[state.sessions.length - 1] : null;
  }

  /** End of the automatic window opened by the last input (or null). */
  function windowEnd(state) {
    if (state.lastInputAt == null) return null;
    var end = state.lastInputAt + windowMs(state);
    if (state.stoppedAt != null && state.stoppedAt >= state.lastInputAt) end = Math.min(end, state.stoppedAt);
    return end;
  }

  /** How much of [from, to] counts under the current mode. */
  function coverage(state, from, to) {
    if (to <= from) return 0;
    if (state.manualSince != null) return Math.max(0, to - Math.max(from, state.manualSince));
    var end = windowEnd(state);
    if (end == null) return 0;
    return Math.max(0, Math.min(to, end) - Math.max(from, state.lastInputAt));
  }

  /** Book all counted time up to t. */
  function advance(state, t) {
    var from = state.countedTo == null ? t : state.countedTo;
    var c = coverage(state, from, t);
    if (c > 0) {
      var end = state.manualSince != null ? t : Math.min(t, windowEnd(state));
      state.committedMs += c;
      var cur = lastSession(state);
      if (cur) { cur.activeMs += c; cur.end = Math.max(cur.end, end); }
      state.lastCoveredAt = Math.max(state.lastCoveredAt || 0, end);
    }
    if (state.countedTo == null || t > state.countedTo) state.countedTo = t;
  }

  function ensureSession(state, t, progress) {
    var cur = lastSession(state);
    var ref = state.lastCoveredAt;
    if (!cur || ref == null || t - ref > SESSION_GAP_MS || !sameDay(ref, t)) {
      cur = { start: t, end: t, activeMs: 0, inputs: 0, pStart: progress, pEnd: progress };
      state.sessions.push(cur);
      if (state.sessions.length > MAX_SESSIONS) state.sessions.splice(0, state.sessions.length - MAX_SESSIONS);
    }
    return cur;
  }

  /** A highlight or an edit happened at time t. `progress` (0..1 or null) is the current progress. */
  function recordInput(state, t, progress) {
    advance(state, t);
    var cur = ensureSession(state, t, progress);
    cur.inputs += 1;
    cur.end = Math.max(cur.end, t);
    if (progress != null) {
      if (cur.pStart == null) cur.pStart = progress;
      cur.pEnd = progress;
    }
    state.inputs += 1;
    state.lastInputAt = t;
    state.lastCoveredAt = Math.max(state.lastCoveredAt || 0, t);
    if (state.manualSince != null) state.heartbeatAt = t;
    return state;
  }

  /**
   * "I'm still studying" answered on the stop notice: the input window had closed at `windowEnd`,
   * the reply came at t. Count the gap (up to capMs) and restart the window.
   */
  function bridge(state, t, capMs, progress) {
    var end = windowEnd(state);
    var extra = end == null ? 0 : Math.max(0, Math.min(t - end, capMs));
    var before = lastSession(state);
    recordInput(state, t, progress);
    if (extra > 0) {
      state.committedMs += extra;
      if (before) before.activeMs += extra;
    }
    return extra;
  }

  function startManual(state, t, progress) {
    if (state.manualSince != null) return state;
    advance(state, t);
    ensureSession(state, t, progress);
    state.manualSince = t;
    state.heartbeatAt = t;
    state.lastCoveredAt = Math.max(state.lastCoveredAt || 0, t);
    return state;
  }

  function stopManual(state, t) {
    advance(state, t);
    state.manualSince = null;
    state.heartbeatAt = null;
    state.stoppedAt = t;
    return state;
  }

  /** Take back time that was counted by mistake (e.g. a clock left running). */
  function uncount(state, ms) {
    ms = Math.max(0, Math.min(ms, state.committedMs));
    state.committedMs -= ms;
    var cur = lastSession(state);
    if (cur) cur.activeMs = Math.max(0, cur.activeMs - ms);
    return ms;
  }

  /** When a running manual clock should stop by itself (no input for MANUAL_MAX_MS). */
  function manualDeadline(state) {
    if (state.manualSince == null) return null;
    return Math.max(state.manualSince, state.lastInputAt || 0) + MANUAL_MAX_MS;
  }

  function pendingMs(state, now) {
    var from = state.countedTo == null ? now : state.countedTo;
    return coverage(state, from, now);
  }

  function isActive(state, now) {
    if (state.manualSince != null) return true;
    var end = windowEnd(state);
    return end != null && now < end;
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

  /** Current session if still running. */
  function currentSession(state, now) {
    var s = lastSession(state);
    if (!s) return null;
    if (state.manualSince != null) return s;
    var ref = Math.max(state.lastCoveredAt || 0, isActive(state, now) ? now : 0);
    if (!ref || now - ref > SESSION_GAP_MS || !sameDay(ref, now)) return null;
    return s;
  }

  function updateProgress(state, progress) {
    if (progress == null) return;
    if (state.baselineProgress == null) state.baselineProgress = progress;
    var s = lastSession(state);
    if (s && (state.manualSince != null || (state.lastCoveredAt != null && s.end >= state.lastCoveredAt - 1))) s.pEnd = progress;
  }

  /** Remember the furthest place you typed (in words from the start of the file). */
  function markEdit(state, words, t) {
    if (state.editMark == null || words > state.editMark.words) state.editMark = { words: words, at: t };
  }

  /**
   * Progress from a per-paragraph word list.
   * paraWords: words per paragraph; lastIdx: index of the paragraph holding the furthest highlight
   * (-1 for none); wordsInLast: words of that paragraph up to the end of the highlight;
   * editWords: furthest typing position in words (or null).
   */
  function computeProgress(paraWords, lastIdx, wordsInLast, editWords) {
    var total = 0, before = 0;
    for (var i = 0; i < paraWords.length; i++) {
      total += paraWords[i];
      if (i < lastIdx) before += paraWords[i];
    }
    if (total === 0) return { total: 0, reached: 0, progress: 0, by: null };
    var hl = lastIdx < 0 ? 0 : Math.min(total, before + Math.max(0, wordsInLast || 0));
    var ed = editWords == null ? 0 : Math.min(total, Math.max(0, editWords));
    var by = ed > hl ? "typing" : (lastIdx < 0 ? null : "highlight");
    var reached = Math.max(hl, ed);
    return { total: total, reached: reached, progress: reached / total, by: by, highlightReached: hl };
  }

  /** Which paragraph holds word number `words` (1-based count), and the word offset inside it. */
  function locateWord(paraWords, words) {
    var acc = 0;
    for (var i = 0; i < paraWords.length; i++) {
      if (acc + paraWords[i] >= words && paraWords[i] > 0) return { idx: i, offset: words - acc };
      acc += paraWords[i];
    }
    return paraWords.length ? { idx: paraWords.length - 1, offset: paraWords[paraWords.length - 1] } : null;
  }

  /**
   * Where did the text change between two lists of paragraph hashes?
   * Returns the index (in `next`) of the last changed paragraph, or null if nothing changed.
   */
  function changedRegion(prev, next) {
    if (!prev || !next) return null;
    if (prev.length === next.length) {
      var same = true;
      for (var k = 0; k < next.length; k++) if (prev[k] !== next[k]) { same = false; break; }
      if (same) return null;
    }
    var a = 0, max = Math.min(prev.length, next.length);
    while (a < max && prev[a] === next[a]) a++;
    var b = 0;
    while (b < max - a && prev[prev.length - 1 - b] === next[next.length - 1 - b]) b++;
    var first = a, last = next.length - 1 - b;
    if (last < first) last = first - 1; // pure deletion: the paragraph just before it
    return { first: first, last: Math.max(0, Math.min(next.length - 1, last)) };
  }

  var MIN_ACTIVE_FOR_ESTIMATE = 3 * 60 * 1000;
  var MIN_GAIN_FOR_ESTIMATE = 0.005;

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
   * 0 = no highlight, 1 = fully highlighted in a counted colour,
   * 2 = mixed (needs a closer look), 3 = highlighted, but in a colour that isn't counted.
   */
  function classifyHighlight(value, filter) {
    var c = normColor(value);
    if (c == null || c === "#FFFFFF" || c === "NONE" || c === "NOHIGHLIGHT") return 0;
    if (c === "") return 2;
    if (!filter || filter === "any") return 1;
    return c === normColor(filter) ? 1 : 3;
  }

  return {
    SESSION_GAP_MS: SESSION_GAP_MS, MANUAL_MAX_MS: MANUAL_MAX_MS,
    countWords: countWords, hash: hash, newState: newState, normalizeState: normalizeState,
    recordInput: recordInput, bridge: bridge, uncount: uncount, startManual: startManual, stopManual: stopManual,
    manualDeadline: manualDeadline, windowEnd: windowEnd, pendingMs: pendingMs, isActive: isActive,
    totalActiveMs: totalActiveMs, sessionActiveMs: sessionActiveMs, todayActiveMs: todayActiveMs,
    currentSession: currentSession, updateProgress: updateProgress, markEdit: markEdit,
    computeProgress: computeProgress, locateWord: locateWord, changedRegion: changedRegion,
    estimate: estimate, formatDuration: formatDuration, toCSV: toCSV,
    normColor: normColor, classifyHighlight: classifyHighlight, sameDay: sameDay
  };
});
