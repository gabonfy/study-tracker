/* Study Tracker – Word task pane. Depends on core.js (window.StudyCore) and Office.js. */
(function () {
  "use strict";
  var C = window.StudyCore;
  var SETTINGS_KEY = "studyTracker";
  var AUTO_OPEN_KEY = "Office.AutoShowTaskpaneWithDocument";

  var state = C.newState();
  var localKey = "studyTracker:untitled";

  // scanner bookkeeping
  var scanning = false, lastScanEnd = 0, lastScanDur = 300, dirty = true, lastSelChange = 0;
  var prevCoarse = null, prevSelIdx = [], fineCache = {};
  var prog = { progress: null, total: 0, reached: 0, lastIdx: -1, wordIdx: null };
  var saveTimer = null, lastError = null;

  var $ = function (id) { return document.getElementById(id); };

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Word) {
      showBanner("This add-in runs in Microsoft Word.");
      return;
    }
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      showBanner("This version of Word is too old for Study Tracker (needs Word 2016 or newer / Microsoft 365).");
      return;
    }
    loadState();
    bindUI();
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function () {
      dirty = true;
      lastSelChange = Date.now();
    });
    render();
    scan();
    setInterval(tick, 1000);
  });

  // ---------- persistence ----------
  function loadState() {
    var url = (Office.context.document && Office.context.document.url) || "";
    localKey = "studyTracker:" + (url || "untitled");
    var fromDoc = null, fromLocal = null;
    try { fromDoc = Office.context.document.settings.get(SETTINGS_KEY); } catch (e) {}
    try { var raw = localStorage.getItem(localKey); if (raw && url) fromLocal = JSON.parse(raw); } catch (e) {}
    var pick = fromDoc;
    if (fromLocal && (!pick || (fromLocal.inputs || 0) > (pick.inputs || 0))) pick = fromLocal;
    state = C.normalizeState(pick ? JSON.parse(JSON.stringify(pick)) : null);
  }

  function persistSoon() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(persistNow, 2500);
  }

  function persistNow() {
    try { localStorage.setItem(localKey, JSON.stringify(state)); } catch (e) {}
    try {
      Office.context.document.settings.set(SETTINGS_KEY, state);
      Office.context.document.settings.saveAsync(function () {});
    } catch (e) {}
  }

  window.addEventListener("pagehide", persistNow);

  // ---------- scanning the document ----------
  function tick() {
    render();
    if (scanning) return;
    var now = Date.now();
    var interval = Math.min(30000, Math.max(2500, lastScanDur * 4));
    var recent = now - lastSelChange < 20000; // keep looking for a bit after you select: the highlight click comes after the selection
    var due = (dirty || recent) ? now - lastScanEnd >= interval : now - lastScanEnd >= 60000;
    if (due) scan();
  }

  function scan() {
    if (scanning) return;
    scanning = true;
    dirty = false;
    var t0 = Date.now();
    var filter = state.colorFilter || "any";

    Word.run(function (ctx) {
      var paras = ctx.document.body.paragraphs;
      paras.load("text,font/highlightColor");
      var selParas = ctx.document.getSelection().paragraphs;
      selParas.load("text");

      return ctx.sync().then(function () {
        var items = paras.items, n = items.length;
        var texts = new Array(n), words = new Array(n), cls = new Array(n), sig = new Array(n);
        for (var i = 0; i < n; i++) {
          var hc = items[i].font.highlightColor;
          texts[i] = items[i].text || "";
          words[i] = C.countWords(texts[i]);
          cls[i] = C.classifyHighlight(hc, filter);
          var nc = C.normColor(hc);
          sig[i] = texts[i].length + "." + (nc == null ? "-" : nc);
        }
        var coarse = n + ":" + C.hash(sig.join("|"));
        var selIdx = locateSelection(texts, selParas.items.slice(0, 6).map(function (p) { return p.text || ""; }));

        // Paragraphs that are only partly highlighted near where you're working: look word by word,
        // so a new highlight inside an already-partly-highlighted paragraph is still noticed.
        var watch = unique(selIdx.concat(prevSelIdx)).filter(function (i) { return i < n && cls[i] === 2; });
        var refined = {};
        watch.forEach(function (i) { refined[i] = queueRefine(items[i]); });

        return ctx.sync().then(function () {
          var fineChanged = false;
          watch.forEach(function (i) {
            var r = readRefine(refined[i], filter);
            refined[i] = r;
            var bits = r.bits.join("");
            var key = String(i);
            if (fineCache[key] && fineCache[key].len === texts[i].length && fineCache[key].bits !== bits) fineChanged = true;
            fineCache[key] = { len: texts[i].length, bits: bits };
          });
          prevSelIdx = selIdx;
          return findFurthest(ctx, items, cls, words, refined, filter).then(function (f) {
            var p = C.computeProgress(words, f.idx, f.wordsIn);
            return { coarse: coarse, fineChanged: fineChanged, p: p, idx: f.idx, wordIdx: f.wordIdx };
          });
        });
      });
    }).then(function (res) {
      var now = Date.now();
      var changed = (prevCoarse !== null && res.coarse !== prevCoarse) || res.fineChanged;
      prevCoarse = res.coarse;
      var oldP = prog.progress;
      prog = { progress: res.p.progress, total: res.p.total, reached: res.p.reached, lastIdx: res.idx, wordIdx: res.wordIdx };
      var hadBaseline = state.baselineProgress != null;
      if (changed && !state.paused) C.recordInput(state, now, prog.progress);
      C.updateProgress(state, prog.progress);
      if (changed || !hadBaseline || oldP !== prog.progress) persistSoon();
      lastError = null;
      if ($("banner").dataset.kind === "error") hideBanner();
    }).catch(function (err) {
      lastError = err;
      console.error(err);
      if (err && err.code !== "GeneralException") showBanner("Couldn't read the document: " + (err.message || err), "error");
    }).then(function () {
      lastScanDur = Date.now() - t0;
      lastScanEnd = Date.now();
      scanning = false;
      render();
    });
  }

  function queueRefine(paragraph) {
    var tr = paragraph.getTextRanges([" "], false);
    tr.load("text,font/highlightColor");
    return tr;
  }

  function readRefine(tr, filter) {
    var texts = [], bits = [];
    tr.items.forEach(function (r) {
      texts.push(r.text || "");
      var c = C.classifyHighlight(r.font.highlightColor, filter);
      bits.push(c === 1 || c === 2 ? 1 : 0);
    });
    return { texts: texts, bits: bits };
  }

  // Walk back from the end of the document to the last highlighted spot.
  function findFurthest(ctx, items, cls, words, refined, filter) {
    var i = items.length - 1, looked = 0;
    function step() {
      for (; i >= 0; i--) {
        if (cls[i] === 1) return Promise.resolve({ idx: i, wordsIn: words[i], wordIdx: null });
        if (cls[i] === 2) break;
      }
      if (i < 0 || looked > 40) return Promise.resolve({ idx: -1, wordsIn: 0, wordIdx: null });
      var idx = i;
      var ready = refined[idx] && refined[idx].bits ? Promise.resolve(refined[idx]) : (function () {
        var tr = queueRefine(items[idx]);
        return ctx.sync().then(function () { return readRefine(tr, filter); });
      })();
      looked++;
      return ready.then(function (r) {
        for (var j = r.bits.length - 1; j >= 0; j--) {
          if (r.bits[j]) {
            return { idx: idx, wordsIn: C.countWords(r.texts.slice(0, j + 1).join(" ")), wordIdx: j };
          }
        }
        i = idx - 1;
        return step();
      });
    }
    return step();
  }

  function locateSelection(texts, selTexts) {
    if (!selTexts.length) return [];
    var first = selTexts[0], best = -1, bestDist = Infinity, anchor = prevSelIdx.length ? prevSelIdx[0] : 0;
    for (var i = 0; i < texts.length; i++) {
      if (texts[i] === first) {
        var d = Math.abs(i - anchor);
        if (d < bestDist) { best = i; bestDist = d; }
      }
    }
    if (best < 0) return [];
    var out = [];
    for (var k = 0; k < selTexts.length && best + k < texts.length; k++) out.push(best + k);
    return out;
  }

  function unique(a) { return a.filter(function (v, i) { return a.indexOf(v) === i; }); }

  // ---------- actions ----------
  function jumpToLastHighlight() {
    if (prog.lastIdx < 0) return;
    var idx = prog.lastIdx, wordIdx = prog.wordIdx;
    Word.run(function (ctx) {
      var paras = ctx.document.body.paragraphs;
      paras.load("text");
      return ctx.sync().then(function () {
        var p = paras.items[idx];
        if (!p) return;
        if (wordIdx == null) { p.getRange("End").select(); return ctx.sync(); }
        var tr = p.getTextRanges([" "], false);
        tr.load("text");
        return ctx.sync().then(function () {
          var target = tr.items[Math.min(wordIdx, tr.items.length - 1)];
          (target || p).getRange("End").select();
          return ctx.sync();
        });
      });
    }).catch(function (e) { console.error(e); });
  }

  function copyCSV() {
    var csv = C.toCSV(state, Date.now());
    var box = $("csvBox");
    var done = function (ok) {
      $("csvBtn").textContent = ok ? "Copied ✓" : "Select the text below and copy it";
      if (!ok) { box.hidden = false; box.value = csv; box.focus(); box.select(); }
      setTimeout(function () { $("csvBtn").textContent = "Copy all sessions as CSV"; }, 2500);
    };
    var fallback = function () {
      box.hidden = false; box.value = csv; box.select();
      var ok = false;
      try { ok = document.execCommand("copy"); } catch (e) {}
      if (ok) box.hidden = true;
      done(ok);
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(csv).then(function () { done(true); }, fallback);
    } else fallback();
  }

  var resetArmed = false;
  function reset() {
    var btn = $("resetBtn");
    if (!resetArmed) {
      resetArmed = true;
      btn.textContent = "Click again to erase all stats for this file";
      btn.classList.add("confirm");
      setTimeout(function () { resetArmed = false; btn.textContent = "Reset tracking for this file"; btn.classList.remove("confirm"); }, 4000);
      return;
    }
    var keep = { idleWindowSec: state.idleWindowSec, colorFilter: state.colorFilter };
    state = C.newState();
    state.idleWindowSec = keep.idleWindowSec;
    state.colorFilter = keep.colorFilter;
    state.baselineProgress = prog.progress;
    resetArmed = false;
    btn.textContent = "Reset tracking for this file";
    btn.classList.remove("confirm");
    persistNow();
    render();
  }

  // ---------- UI ----------
  function bindUI() {
    $("idleSelect").value = String(state.idleWindowSec);
    if (!$("idleSelect").value) $("idleSelect").value = "120";
    $("colorSelect").value = state.colorFilter || "any";
    var auto = false;
    try { auto = !!Office.context.document.settings.get(AUTO_OPEN_KEY); } catch (e) {}
    $("autoOpen").checked = auto;

    $("idleSelect").addEventListener("change", function (e) { state.idleWindowSec = Number(e.target.value); persistNow(); render(); });
    $("colorSelect").addEventListener("change", function (e) {
      state.colorFilter = e.target.value;
      fineCache = {};
      // progress is now measured differently: restart the pace baseline from the new reading
      state.baselineProgress = null;
      persistNow();
      dirty = true; lastScanEnd = 0;
    });
    $("autoOpen").addEventListener("change", function (e) {
      try {
        Office.context.document.settings.set(AUTO_OPEN_KEY, e.target.checked);
        Office.context.document.settings.saveAsync(function () {});
      } catch (err) {}
    });
    $("pauseBtn").addEventListener("click", function () { state.paused = !state.paused; persistNow(); render(); });
    $("jumpBtn").addEventListener("click", jumpToLastHighlight);
    $("csvBtn").addEventListener("click", copyCSV);
    $("resetBtn").addEventListener("click", reset);
  }

  function render() {
    var now = Date.now();
    var active = C.isActive(state, now);

    // status
    var dot = $("statusDot"), label = $("statusLabel"), hint = $("statusHint"), pb = $("pauseBtn");
    dot.className = "dot" + (state.paused ? " paused" : active ? " active" : "");
    if (state.paused) {
      label.textContent = "Paused";
      hint.textContent = "Highlights aren't being timed. Press ▶ to resume.";
    } else if (active) {
      var since = now - state.lastInputAt, left = state.idleWindowSec * 1000 - since;
      label.textContent = "Studying";
      hint.textContent = "Last highlight " + C.formatDuration(since) + " ago · counting for another " + C.formatDuration(left);
    } else if (state.lastInputAt) {
      label.textContent = "Idle";
      hint.textContent = "Not counting. Your next highlight resumes the clock.";
    } else {
      label.textContent = "Ready";
      hint.textContent = "Highlight as you study, the clock starts with your first highlight.";
    }
    pb.innerHTML = state.paused
      ? '<svg viewBox="0 0 20 20" width="16" height="16" aria-hidden="true"><path d="M6 4l10 6-10 6z"/></svg>'
      : '<svg viewBox="0 0 20 20" width="16" height="16" aria-hidden="true"><rect x="5" y="4" width="3.2" height="12" rx="1"/><rect x="11.8" y="4" width="3.2" height="12" rx="1"/></svg>';
    pb.title = state.paused ? "Resume tracking" : "Pause tracking";
    pb.setAttribute("aria-label", pb.title);

    // times
    var cur = C.currentSession(state, now);
    $("sessionTime").textContent = C.formatDuration(cur ? C.sessionActiveMs(state, cur, now) : 0, { clock: true });
    $("todayTime").textContent = C.formatDuration(C.todayActiveMs(state, now));
    $("totalTime").textContent = C.formatDuration(C.totalActiveMs(state, now));
    $("inputCount").textContent = String(state.inputs);

    // progress
    if (prog.progress == null) {
      $("progressPct").textContent = "–";
    } else {
      var pct = prog.progress * 100;
      $("progressPct").textContent = (pct < 10 ? pct.toFixed(1) : Math.round(pct)) + "%";
      $("progressFill").style.width = pct + "%";
      $("progressMark").style.left = "calc(" + pct + "% - 1px)";
      $("progressDetail").textContent = prog.lastIdx < 0
        ? "No highlights yet. " + prog.total.toLocaleString() + " words in this file."
        : "Word " + prog.reached.toLocaleString() + " of " + prog.total.toLocaleString() + " · paragraph " + (prog.lastIdx + 1);
    }
    $("jumpBtn").disabled = prog.lastIdx < 0;

    // estimate
    var e = C.estimate(state, prog.progress, now);
    if (prog.progress == null) {
      $("etaMain").textContent = "–"; $("etaDetail").textContent = "";
    } else if (e.ok && e.done) {
      $("etaMain").textContent = "Done 🎉";
      $("etaDetail").textContent = "You've highlighted to the end of the file.";
    } else if (e.ok) {
      $("etaMain").textContent = "≈ " + C.formatDuration(roundEta(e.remainingMs));
      $("etaDetail").textContent = "Pace " + (e.pacePerHour * 100).toFixed(1) + "% of the file per active hour, from " +
        (e.gained * 100).toFixed(1) + "% covered in " + C.formatDuration(e.activeMs) + ".";
    } else {
      $("etaMain").textContent = "Gathering data…";
      var bits = [];
      if (e.needMs > 0) bits.push(C.formatDuration(e.needMs) + " more active study");
      if (e.needGain > 0) bits.push("a bit more progress (" + (e.needGain * 100).toFixed(1) + "%)");
      $("etaDetail").textContent = "An estimate appears after " + bits.join(" and ") + ".";
    }

    renderSessions(now);
  }

  function roundEta(ms) {
    var m = ms / 60000;
    if (m < 60) return Math.max(1, Math.round(m)) * 60000;
    return Math.round(m / 5) * 5 * 60000;
  }

  var lastSessionsHtml = "";
  function renderSessions(now) {
    var list = state.sessions.slice(-8).reverse();
    var html;
    if (!list.length) {
      html = '<li class="muted small empty">No study sessions yet. Highlight something to start.</li>';
    } else {
      html = list.map(function (s) {
        var gain = s.pStart != null && s.pEnd != null ? (s.pEnd - s.pStart) * 100 : null;
        var g = gain == null ? "" : (gain >= 0 ? "+" : "") + gain.toFixed(1) + "%";
        return '<li><div class="when"><b>' + dayLabel(s.start, now) + '</b> <span>' + hm(s.start) + "–" + hm(s.end) +
          '</span></div><div class="dur">' + C.formatDuration(C.sessionActiveMs(state, s, now)) +
          '</div><div class="gain' + (gain > 0 ? " pos" : "") + '">' + g + "</div></li>";
      }).join("");
    }
    if (html !== lastSessionsHtml) { $("sessionList").innerHTML = html; lastSessionsHtml = html; }
  }

  function hm(t) { var d = new Date(t); return d.getHours() + ":" + (d.getMinutes() < 10 ? "0" : "") + d.getMinutes(); }
  function dayLabel(t, now) {
    if (C.sameDay(t, now)) return "Today";
    if (C.sameDay(t, now - 86400000)) return "Yesterday";
    return new Date(t).toLocaleDateString(undefined, { weekday: "short", day: "numeric", month: "short" });
  }

  function showBanner(msg, kind) { var b = $("banner"); b.textContent = msg; b.hidden = false; b.dataset.kind = kind || ""; }
  function hideBanner() { var b = $("banner"); b.hidden = true; b.dataset.kind = ""; }
})();
