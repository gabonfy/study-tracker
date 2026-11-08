/* Study Tracker – Word task pane (v2). Depends on core.js (window.StudyCore) and Office.js. */
(function () {
  "use strict";
  var C = window.StudyCore;
  var SETTINGS_KEY = "studyTracker";
  var AUTO_OPEN_KEY = "Office.AutoShowTaskpaneWithDocument";
  var BRIDGE_CAP_MS = 5 * 60 * 1000; // after the pop-up, at most this much of the gap is counted back

  var state = C.newState();
  var localKey = "studyTracker:untitled";

  // scanner bookkeeping
  var scanning = false, lastScanEnd = 0, lastScanDur = 300, dirty = true, lastPoke = 0;
  var prevHl = null, prevTextHashes = null, prevSelIdx = [], fineCache = {};
  var prog = { progress: null, total: 0, reached: 0, by: null, lastIdx: -1, wordIdx: null, paraWords: [] };
  var saveTimer = null, lastHeartbeat = 0;
  var wasActive = false, dialog = null;

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
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, poke);
    listenForEdits();
    wasActive = C.isActive(state, Date.now());
    render();
    scan();
    setInterval(tick, 1000);
  });

  function poke() { dirty = true; lastPoke = Date.now(); }

  // Word 2023+ tells us directly when paragraphs change: scan right away instead of waiting.
  function listenForEdits() {
    if (!Office.context.requirements.isSetSupported("WordApi", "1.6")) return;
    var onEdit = function (args) { if (!args || args.source !== "Remote") poke(); };
    Word.run(function (ctx) {
      ctx.document.onParagraphChanged.add(onEdit);
      ctx.document.onParagraphAdded.add(onEdit);
      ctx.document.onParagraphDeleted.add(onEdit);
      return ctx.sync();
    }).catch(function (e) { console.warn("paragraph events unavailable", e); });
  }

  // ---------- persistence ----------
  function loadState() {
    var url = (Office.context.document && Office.context.document.url) || "";
    localKey = "studyTracker:" + (url || "untitled");
    var fromDoc = null, fromLocal = null;
    try { fromDoc = Office.context.document.settings.get(SETTINGS_KEY); } catch (e) {}
    try { var raw = localStorage.getItem(localKey); if (raw && url) fromLocal = JSON.parse(raw); } catch (e) {}
    var pick = fromDoc;
    if (fromLocal && (!pick || (fromLocal.savedAt || 0) > (pick.savedAt || 0) ||
        ((fromLocal.savedAt || 0) === (pick.savedAt || 0) && (fromLocal.inputs || 0) > (pick.inputs || 0)))) pick = fromLocal;
    state = C.normalizeState(pick ? JSON.parse(JSON.stringify(pick)) : null);

    // The clock was running when Word or the panel closed: count it up to the last sign of life.
    if (state.manualSince != null) {
      var until = Math.max(state.heartbeatAt || state.manualSince, state.manualSince);
      C.stopManual(state, Math.min(until, C.manualDeadline(state)));
      showBanner("The clock was still running when the panel closed. It was counted until " + hm(until) + " and is now stopped.", "info");
      persistNow();
    }
  }

  function persistSoon() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(persistNow, 2500);
  }

  function persistLocal() {
    state.savedAt = Date.now();
    try { localStorage.setItem(localKey, JSON.stringify(state)); } catch (e) {}
  }

  function persistNow() {
    clearTimeout(saveTimer);
    persistLocal();
    try {
      Office.context.document.settings.set(SETTINGS_KEY, state);
      Office.context.document.settings.saveAsync(function () {});
    } catch (e) {}
  }

  window.addEventListener("pagehide", persistNow);

  // ---------- the 1-second loop ----------
  function tick() {
    var now = Date.now();

    // a forgotten clock stops itself
    var deadline = C.manualDeadline(state);
    if (deadline != null && now >= deadline) {
      C.stopManual(state, deadline);
      wasActive = false;
      persistNow();
      notifyStop("manual");
    }

    // heartbeat so a closed panel doesn't leave the clock running forever
    if (state.manualSince != null && now - lastHeartbeat > 15000) {
      state.heartbeatAt = now;
      lastHeartbeat = now;
      persistLocal();
    }

    // the automatic window just ran out -> pop-up
    var active = C.isActive(state, now);
    if (wasActive && !active && state.manualSince == null) notifyStop("auto");
    wasActive = active;

    render();

    if (scanning) return;
    var interval = Math.min(30000, Math.max(2500, lastScanDur * 4));
    var recent = now - lastPoke < 20000; // keep looking for a bit: the highlight click comes after the selection
    var due = (dirty || recent) ? now - lastScanEnd >= interval : now - lastScanEnd >= Math.max(10000, interval);
    if (due) scan();
  }

  // ---------- scanning the document ----------
  function scan() {
    if (scanning) return;
    scanning = true;
    dirty = false;
    var t0 = Date.now();
    var filter = state.colorFilter || "any";

    Word.run(function (ctx) {
      var paras = ctx.document.body.paragraphs;
      paras.load("text,font/highlightColor");
      var sel = ctx.document.getSelection();
      var selParas = sel.paragraphs;
      selParas.load("text");

      return ctx.sync().then(function () {
        var items = paras.items, n = items.length;
        var texts = new Array(n), words = new Array(n), cls = new Array(n), hl = new Array(n), th = new Array(n);
        for (var i = 0; i < n; i++) {
          var hc = items[i].font.highlightColor;
          texts[i] = items[i].text || "";
          words[i] = C.countWords(texts[i]);
          cls[i] = C.classifyHighlight(hc, filter);
          var nc = C.normColor(hc);
          hl[i] = nc == null ? "-" : nc;
          th[i] = C.hash(texts[i]);
        }
        var hlSig = n + ":" + C.hash(hl.join("|"));
        var region = C.changedRegion(prevTextHashes, th);      // text typed or deleted?
        var selIdx = locateSelection(texts, selParas.items.slice(0, 6).map(function (p) { return p.text || ""; }));

        // Paragraphs only partly highlighted near where you're working: look word by word, so a new
        // highlight inside an already-partly-highlighted paragraph is still noticed.
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

          // Where exactly did you type? Use the cursor if it sits in the changed paragraphs.
          var cursorRange = null, editIdx = null;
          if (region) {
            editIdx = region.last;
            if (selIdx.length && selIdx[0] >= region.first && selIdx[0] <= region.last && selParas.items[0]) {
              editIdx = selIdx[0];
              cursorRange = selParas.items[0].getRange("Start").expandTo(sel.getRange("Start"));
              cursorRange.load("text");
            }
          }
          var cursorSync = cursorRange ? ctx.sync().then(function () { return cursorRange.text; }, function () { return null; })
                                       : Promise.resolve(null);
          return cursorSync.then(function (cursorText) {
          var editWords = null;
          if (editIdx != null) {
            var before = 0;
            for (var k = 0; k < editIdx; k++) before += words[k];
            var inPara = words[editIdx] || 0;
            if (cursorText != null) inPara = Math.min(inPara, C.countWords(cursorText));
            editWords = before + inPara;
          }

          prevSelIdx = selIdx;
          prevTextHashes = th;
          return findFurthest(ctx, items, cls, words, refined, filter).then(function (f) {
            return {
              hlSig: hlSig, textChanged: !!region, fineChanged: fineChanged, editWords: editWords,
              words: words, idx: f.idx, wordsIn: f.wordsIn, wordIdx: f.wordIdx
            };
          });
          });
        });
      });
    }).then(function (res) {
      var now = Date.now();
      var hlChanged = !res.textChanged && ((prevHl !== null && res.hlSig !== prevHl) || res.fineChanged);
      prevHl = res.hlSig;
      if (res.editWords != null) C.markEdit(state, res.editWords, now);

      var p = C.computeProgress(res.words, res.idx, res.wordsIn, state.editMark ? state.editMark.words : null);
      var oldP = prog.progress;
      prog = { progress: p.progress, total: p.total, reached: p.reached, by: p.by,
               lastIdx: res.idx, wordIdx: res.wordIdx, paraWords: res.words };

      var hadBaseline = state.baselineProgress != null;
      if (res.textChanged || hlChanged) {
        C.recordInput(state, now, prog.progress);
        wasActive = true;
        closeDialog();
      }
      C.updateProgress(state, prog.progress);
      if (res.textChanged || hlChanged || !hadBaseline || oldP !== prog.progress) persistSoon();
      if ($("banner").dataset.kind === "error") hideBanner();
    }).catch(function (err) {
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
          if (r.bits[j]) return { idx: idx, wordsIn: C.countWords(r.texts.slice(0, j + 1).join(" ")), wordIdx: j };
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

  // ---------- the clock button ----------
  function toggleClock() {
    var now = Date.now();
    if (state.manualSince != null) {
      C.stopManual(state, now);
      wasActive = false;
    } else {
      C.startManual(state, now, prog.progress);
      lastHeartbeat = now;
      wasActive = true;
    }
    closeDialog();
    persistNow();
    render();
  }

  // ---------- pop-up when the clock stops ----------
  function notifyStop(kind) {
    if (!state.notifyOnStop || dialog) return;
    if (!Office.context.ui || !Office.context.ui.displayDialogAsync) return;
    var url = location.href.replace(/[?#].*$/, "").replace(/[^\/]*$/, "") + "notify.html?kind=" + kind +
      "&w=" + encodeURIComponent(state.idleWindowSec) + "&v=2";
    var stoppedAt = kind === "auto" ? C.windowEnd(state) : Date.now();
    Office.context.ui.displayDialogAsync(url, { height: 30, width: 26, displayInIframe: true }, function (res) {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return;
      dialog = res.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        var msg = arg && arg.message;
        closeDialog();
        var now = Date.now();
        if (msg === "still") {
          C.bridge(state, now, BRIDGE_CAP_MS, prog.progress);
        } else if (msg === "drop") {
          C.uncount(state, C.MANUAL_MAX_MS);
        } else if (msg === "start") {
          var from = Math.max(stoppedAt || now, now - BRIDGE_CAP_MS, state.countedTo || 0);
          C.startManual(state, Math.min(from, now), prog.progress);
          lastHeartbeat = now;
        }
        wasActive = C.isActive(state, now);
        persistNow();
        render();
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, function () { dialog = null; });
    });
  }

  function closeDialog() {
    if (!dialog) return;
    try { dialog.close(); } catch (e) {}
    dialog = null;
  }

  // ---------- other actions ----------
  function jumpToFurthest() {
    if (prog.by == null) return;
    var target;
    if (prog.by === "highlight") {
      target = { idx: prog.lastIdx, wordIdx: prog.wordIdx, offset: null };
    } else {
      var loc = C.locateWord(prog.paraWords, prog.reached);
      if (!loc) return;
      target = { idx: loc.idx, wordIdx: null, offset: loc.offset };
    }
    Word.run(function (ctx) {
      var paras = ctx.document.body.paragraphs;
      paras.load("text");
      return ctx.sync().then(function () {
        var p = paras.items[target.idx];
        if (!p) return;
        if (target.wordIdx == null && target.offset == null) { p.getRange("End").select(); return ctx.sync(); }
        var tr = p.getTextRanges([" "], false);
        tr.load("text");
        return ctx.sync().then(function () {
          var j = target.wordIdx;
          if (j == null) {
            var acc = "";
            for (j = 0; j < tr.items.length - 1; j++) {
              acc += tr.items[j].text + " ";
              if (C.countWords(acc) >= target.offset) break;
            }
          }
          var r = tr.items[Math.min(j, tr.items.length - 1)];
          (r || p).getRange("End").select();
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
    var keep = { idleWindowSec: state.idleWindowSec, colorFilter: state.colorFilter, notifyOnStop: state.notifyOnStop };
    state = C.newState();
    state.idleWindowSec = keep.idleWindowSec;
    state.colorFilter = keep.colorFilter;
    state.notifyOnStop = keep.notifyOnStop;
    resetArmed = false;
    wasActive = false;
    btn.textContent = "Reset tracking for this file";
    btn.classList.remove("confirm");
    persistNow();
    dirty = true; lastScanEnd = 0; // re-read progress (typing marks are gone) and set a new baseline
    render();
  }

  // ---------- UI ----------
  function bindUI() {
    $("idleSelect").value = String(state.idleWindowSec);
    if (!$("idleSelect").value) $("idleSelect").value = "120";
    $("colorSelect").value = state.colorFilter || "any";
    $("notifyStop").checked = state.notifyOnStop !== false;
    var auto = false;
    try { auto = !!Office.context.document.settings.get(AUTO_OPEN_KEY); } catch (e) {}
    $("autoOpen").checked = auto;

    $("idleSelect").addEventListener("change", function (e) { state.idleWindowSec = Number(e.target.value); persistNow(); render(); });
    $("notifyStop").addEventListener("change", function (e) { state.notifyOnStop = e.target.checked; persistNow(); });
    $("colorSelect").addEventListener("change", function (e) {
      state.colorFilter = e.target.value;
      fineCache = {};
      state.baselineProgress = null; // progress is measured differently now: restart the pace from here
      persistNow();
      dirty = true; lastScanEnd = 0;
    });
    $("autoOpen").addEventListener("change", function (e) {
      try {
        Office.context.document.settings.set(AUTO_OPEN_KEY, e.target.checked);
        Office.context.document.settings.saveAsync(function () {});
      } catch (err) {}
    });
    $("clockBtn").addEventListener("click", toggleClock);
    $("jumpBtn").addEventListener("click", jumpToFurthest);
    $("csvBtn").addEventListener("click", copyCSV);
    $("resetBtn").addEventListener("click", reset);
  }

  var PLAY = '<path d="M6 4l10 6-10 6z"/>';
  var STOP = '<rect x="5" y="5" width="10" height="10" rx="1.5"/>';

  function render() {
    var now = Date.now();
    var manual = state.manualSince != null;
    var active = C.isActive(state, now);

    // status
    var dot = $("statusDot"), label = $("statusLabel"), hint = $("statusHint");
    dot.className = "dot" + (manual ? " active manual" : active ? " active" : "");
    if (manual) {
      label.textContent = "Clock running";
      hint.textContent = "Everything counts until you press Stop.";
    } else if (active) {
      var since = now - state.lastInputAt, left = C.windowEnd(state) - now;
      label.textContent = "Studying";
      hint.textContent = "Last highlight or edit " + C.formatDuration(since) + " ago · counting for another " + C.formatDuration(left);
    } else if (state.lastInputAt || state.committedMs) {
      label.textContent = "Idle";
      hint.textContent = "Not counting. Highlight, type, or start the clock.";
    } else {
      label.textContent = "Ready";
      hint.textContent = "The clock starts with your first highlight or edit, or press Start.";
    }

    // clock button
    $("clockBtn").classList.toggle("running", manual);
    $("clockIcon").innerHTML = manual ? STOP : PLAY;
    $("clockLabel").textContent = manual ? "Stop clock" : "Start clock";
    $("clockHint").textContent = manual
      ? "Started at " + hm(state.manualSince) + ". Stops by itself after 2 h without highlights or typing."
      : "For exercises or reading without highlighting. Runs until you stop it.";

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
      $("progressDetail").textContent = prog.by == null
        ? "No highlights or typing yet. " + prog.total.toLocaleString() + " words in this file."
        : "Word " + prog.reached.toLocaleString() + " of " + prog.total.toLocaleString() +
          (prog.by === "typing" ? " · where you last typed furthest" : " · your furthest highlight");
    }
    $("jumpBtn").disabled = prog.by == null;

    // estimate
    var e = C.estimate(state, prog.progress, now);
    if (prog.progress == null) {
      $("etaMain").textContent = "–"; $("etaDetail").textContent = "";
    } else if (e.ok && e.done) {
      $("etaMain").textContent = "Done 🎉";
      $("etaDetail").textContent = "You've reached the end of the file.";
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
    var newest = state.sessions[state.sessions.length - 1];
    var liveEnd = C.isActive(state, now) ? now : null;
    var html;
    if (!list.length) {
      html = '<li class="muted small empty">No study sessions yet. Highlight, type or start the clock.</li>';
    } else {
      html = list.map(function (s) {
        var gain = s.pStart != null && s.pEnd != null ? (s.pEnd - s.pStart) * 100 : null;
        var g = gain == null ? "" : (gain >= 0 ? "+" : "") + gain.toFixed(1) + "%";
        return '<li><div class="when"><b>' + dayLabel(s.start, now) + '</b> <span>' + hm(s.start) + "–" + hm(s === newest && liveEnd ? liveEnd : s.end) +
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

  function showBanner(msg, kind) {
    var b = $("banner"); b.textContent = msg; b.hidden = false; b.dataset.kind = kind || "";
    b.title = "Click to dismiss"; b.onclick = hideBanner;
  }
  function hideBanner() { var b = $("banner"); b.hidden = true; b.dataset.kind = ""; }
})();
