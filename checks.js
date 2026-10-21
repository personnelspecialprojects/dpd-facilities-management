(window.FM_FILE_VERSIONS = window.FM_FILE_VERSIONS || {})["checks.js"] = "2026.10.21-4";

/*
 * Security checks board (Dashboard, left side).
 * Zones and the watches they belong to come from SharePoint (FM Zones).
 * Clicking a zone's check box logs the signed-in person at the current time.
 * Manual entry records a check at the time it actually happened, for
 * yourself or someone else. "Couldn't check" records a check that wasn't done
 * (who, roughly when it was due, and why). Entries are voided, never deleted.
 */
(function () {
  "use strict";
  var FM = window.FM, C = window.FM_CONFIG, L = C.lists;
  var SELECT = "Id,Created,ZoneId,ZoneName,Watch,BoardDay,CheckedAt,StaffEmail,StaffName,EnteredByEmail,EnteredByName,Method,Reason,Voided,VoidReason,VoidedByEmail";
  var S = { day: null, rows: [], showVoided: false, loadSeq: 0, followToday: true, busy: {} };

  function today() { return FM.dayKey(new Date()); }
  function me() { return FM.state.me; }
  function panel() { return FM.$("#checks-panel"); }

  async function load(day) {
    var seq = ++S.loadSeq;
    var rows;
    try {
      rows = await FM.sp.getAll(L.checks, { $select: SELECT, $filter: "BoardDay eq " + FM.odataStr(day) });
    } catch (e) {
      if (seq === S.loadSeq) FM.panelError(panel(), "Security checks", e);
      return;
    }
    if (seq !== S.loadSeq) return;
    S.day = day;
    S.rows = rows;
    render();
  }

  function zonesFor(watch) {
    var list = FM.data.zones.filter(function (z) { return FM.isActive(z) && FM.zoneWatches(z).indexOf(watch) >= 0; });
    var ids = {};
    list.forEach(function (z) { ids[z.Id] = true; });
    // Zones that were retired or moved off this watch still show on days they were checked.
    S.rows.forEach(function (r) {
      if (r.Watch !== watch || ids[r.ZoneId]) return;
      if (r.Voided && !S.showVoided) return;
      ids[r.ZoneId] = true;
      var z = FM.zoneById(r.ZoneId);
      list.push({ Id: r.ZoneId, Title: z ? z.Title : r.ZoneName, Description: z ? z.Description : "", SortOrder: z ? z.SortOrder : 9999, notCurrent: true });
    });
    return list.sort(FM.bySort);
  }

  function checksFor(zoneId, watch) {
    return S.rows.filter(function (r) {
      return r.ZoneId === zoneId && r.Watch === watch && (S.showVoided || !r.Voided);
    }).sort(function (a, b) { return new Date(a.CheckedAt) - new Date(b.CheckedAt); });
  }

  function chip(r) {
    var mine = FM.sameUser(r.StaffEmail, me().email);
    var missed = r.Method === "Missed";
    var title = (missed ? "Not done (reason recorded)" : r.Method === "Manual" ? "Manual entry" : "Logged") + " · " + FM.fmtDateTime(r.CheckedAt) + (r.Voided ? " · voided" : "");
    return '<button class="chip' + (mine ? " mine" : "") + (missed ? " missed" : "") + (r.Voided ? " voided" : "") + '" data-act="chip" data-id="' + r.Id + '" title="' + FM.esc(title) + '">' +
      FM.esc(r.StaffName || FM.nameForEmail(r.StaffEmail)) + " <time>" + FM.esc(FM.fmtTime(r.CheckedAt)) + "</time>" +
      (r.Method === "Manual" ? "<em>M</em>" : "") + (missed ? "<em>Not done</em>" : "") + "</button>";
  }

  function render() {
    var el = panel(); if (!el || !S.day) return;
    var isToday = S.day === today();
    var anyZones = FM.data.zones.some(FM.isActive);
    var html =
      '<div class="panel-head"><h2>Security checks</h2><div class="toolbar">' +
      '<button class="btn icon" data-act="prev" title="Previous day" aria-label="Previous day">‹</button>' +
      '<input type="date" id="chk-day" value="' + S.day + '" max="' + today() + '" aria-label="Board date">' +
      '<button class="btn icon" data-act="next" title="Next day" aria-label="Next day"' + (isToday ? " disabled" : "") + ">›</button>" +
      (isToday ? "" : '<button class="btn" data-act="today">Today</button>') +
      '<button class="btn" data-act="manual" data-save>Manual entry</button>' +
      '<button class="btn" data-act="missed" data-save>Couldn\'t check</button>' +
      '<label class="small row-gap"><input type="checkbox" data-act="voided"' + (S.showVoided ? " checked" : "") + "> Show voided</label>" +
      "</div></div>";
    html += '<div class="small muted" style="margin:-6px 0 10px">' + FM.esc(FM.fmtDayLong(S.day)) + "</div>";
    if (!isToday) html += '<div class="note">You\'re looking at a past day. Use <b>Manual entry</b> for a check that was done but not logged, or <b>Couldn\'t check</b> to explain one that wasn\'t done.</div>';
    if (FM.state.loadErrors.zones) {
      html += '<div class="load-error">The zones couldn\'t be loaded. See the message at the top of the page.</div>';
    } else if (!anyZones) {
      html += '<div class="empty">No zones are set up yet. ' + (FM.state.isAdmin ? "Add them in Admin → Zones." : "Ask an admin to add them.") + "</div>";
    }
    C.watches.forEach(function (w, i) {
      var zones = zonesFor(w);
      var live = S.rows.filter(function (r) { return r.Watch === w && !r.Voided; });
      var n = live.filter(function (r) { return r.Method !== "Missed"; }).length;
      var nm = live.length - n;
      html += '<section class="watch watch-' + (i + 1) + '" data-watch="' + FM.esc(w) + '"><h3>' + FM.esc(w) + ' watch <span class="count">' + n + " check" + (n === 1 ? "" : "s") +
        (nm ? " · " + nm + " not done" : "") + "</span></h3>";
      if (!zones.length) {
        if (anyZones) html += '<div class="zone-row"><span class="muted small">No zones on this watch.</span></div>';
      }
      zones.forEach(function (z) {
        var list = checksFor(z.Id, w);
        html += '<div class="zone-row' + (z.notCurrent ? " retired" : "") + '" data-zone="' + z.Id + '">' +
          '<div class="zone-name"><b>' + FM.esc(z.Title) + "</b>" + (z.Description ? "<span>" + FM.esc(z.Description) + "</span>" : "") + "</div>" +
          '<div class="chips">' + (list.length ? list.map(chip).join("") : '<span class="none">No checks yet</span>') + "</div>" +
          (isToday && !z.notCurrent
            ? '<button class="check-btn" data-act="log" data-save data-zone="' + z.Id + '" data-watch="' + FM.esc(w) + '" title="Log a check now"><span class="box"></span>Check</button>'
            : "<span></span>") +
          "</div>";
      });
      html += "</section>";
    });
    el.innerHTML = html;
  }

  // ------------------------------------------------------------ actions
  async function autoLog(btn, zoneId, watch) {
    var key = zoneId + "|" + watch;
    if (S.busy[key]) return;
    FM.guardSave();
    var z = FM.zoneById(zoneId);
    if (!z) throw new Error("That zone isn't available anymore. Refresh and try again.");
    var now = new Date();
    var recent = S.rows.find(function (r) {
      return r.ZoneId === zoneId && r.Watch === watch && !r.Voided && r.Method !== "Missed" && FM.sameUser(r.StaffEmail, me().email) &&
        Math.abs(now - new Date(r.CheckedAt)) < 5 * 60000;
    });
    if (recent && !(await FM.confirm("You logged this zone at " + FM.fmtTime(recent.CheckedAt) + ". Log another check now?", "Log again"))) return;

    S.busy[key] = true;
    if (btn) { btn.disabled = true; btn.classList.add("done"); }
    try {
      var row = {
        Title: "Check",
        ZoneId: zoneId,
        ZoneName: z.Title,
        Watch: watch,
        BoardDay: FM.dayKey(now),
        CheckedAt: now.toISOString(),
        StaffEmail: me().email,
        StaffName: me().name,
        EnteredByEmail: me().email,
        EnteredByName: me().name,
        Method: "Auto",
        Voided: false
      };
      var created = await FM.sp.create(L.checks, row);
      created = Object.assign({}, row, created || {});
      if (created.BoardDay === S.day) S.rows.push(created);
      FM.audit("Checks", "Check logged", created.Id, "zone #" + zoneId + " · " + watch + " watch · auto");
      render();
      FM.toast("Check logged at " + FM.fmtTime(now) + ".", {
        action: { label: "Undo", fn: function () { voidCheck(created, "Undone right after logging").catch(function (e) { FM.error(e); }); } }
      });
    } finally {
      S.busy[key] = false;
      if (btn && btn.isConnected) { btn.disabled = false; btn.classList.remove("done"); }
    }
  }

  function canVoid(r) {
    return FM.state.isAdmin || FM.sameUser(r.StaffEmail, me().email) || FM.sameUser(r.EnteredByEmail, me().email);
  }

  async function voidCheck(r, reason) {
    FM.guardSave();
    await FM.sp.update(L.checks, r.Id, { Voided: true, VoidReason: reason, VoidedByEmail: me().email });
    r.Voided = true; r.VoidReason = reason; r.VoidedByEmail = me().email;
    FM.audit("Checks", "Check voided", r.Id, "zone #" + r.ZoneId + " · " + r.Watch + " watch · reason: " + reason);
    render();
    FM.toast("Check voided.");
  }

  function showChip(id) {
    var r = S.rows.find(function (x) { return x.Id === id; });
    if (!r) return;
    var z = FM.zoneById(r.ZoneId);
    var missed = r.Method === "Missed";
    var dl =
      "<dl class=\"kv\">" +
      "<dt>Zone</dt><dd>" + FM.esc(z ? z.Title : r.ZoneName) + "</dd>" +
      "<dt>Watch</dt><dd>" + FM.esc(r.Watch) + "</dd>" +
      "<dt>" + (missed ? "Assigned to" : "Checked by") + "</dt><dd>" + FM.esc(r.StaffName || FM.nameForEmail(r.StaffEmail)) + "</dd>" +
      "<dt>" + (missed ? "Was due about" : "Checked at") + "</dt><dd>" + FM.esc(FM.fmtDateTime(r.CheckedAt)) + "</dd>" +
      "<dt>How</dt><dd>" + (missed ? '<span class="pill bad">Not done</span>' : r.Method === "Manual" ? '<span class="pill warn">Manual entry</span>' : '<span class="pill blue">Logged at the time</span>') + "</dd>" +
      (r.Method === "Manual" || missed
        ? "<dt>Entered by</dt><dd>" + FM.esc(r.EnteredByName || FM.nameForEmail(r.EnteredByEmail)) + " on " + FM.esc(FM.fmtDateTime(r.Created)) + "</dd>" +
          "<dt>Reason</dt><dd>" + FM.esc(r.Reason || "") + "</dd>"
        : "") +
      (r.Voided
        ? "<dt>Voided</dt><dd><span class=\"pill bad\">Voided</span> by " + FM.esc(FM.nameForEmail(r.VoidedByEmail)) + ": " + FM.esc(r.VoidReason || "") + "</dd>"
        : "") +
      "</dl>";
    var actions = [{ label: "Close" }];
    if (!r.Voided && canVoid(r)) {
      actions.unshift({
        label: "Void this check", cls: "danger", id: "void", onClick: async function () {
          var reason = await FM.askReason(missed ? "Void entry" : "Void check", "Why is this " + (missed ? "entry" : "check") + " being voided?", "Void");
          if (!reason) return false;
          await voidCheck(r, reason);
        }
      });
    }
    FM.modal({ title: missed ? "Check not done" : "Security check", body: dl, actions: actions });
  }

  function zoneOptions(watch, selected) {
    var zones = FM.data.zones.filter(function (z) { return FM.isActive(z) && FM.zoneWatches(z).indexOf(watch) >= 0; }).sort(FM.bySort);
    if (!zones.length) return '<option value="">No zones on this watch</option>';
    return zones.map(function (z) {
      return '<option value="' + z.Id + '"' + (Number(selected) === z.Id ? " selected" : "") + ">" + FM.esc(z.Title) + "</option>";
    }).join("");
  }

  // pre.mode: "manual" (a check that was done but not logged) or
  // "missed" (a check that couldn't be done, with the reason).
  function openManual(pre) {
    pre = pre || {};
    var missed = pre.mode === "missed";
    var watch = pre.watch || C.watches[0];
    var day = S.day || today();
    var body =
      (missed
        ? '<p class="small muted">Use this when a check couldn\'t be done. Enter about when it was due and why it wasn\'t done.</p>'
        : '<p class="small muted">Use this when a check was done but wasn\'t logged at the time. Enter when it actually happened.</p>') +
      '<label class="field"><span>' + (missed ? "Who was supposed to do it" : "Checked by") + '</span><select id="m-staff">' + FM.memberOptions(me().email) + "</select></label>" +
      '<div class="field-row">' +
      '<label class="field"><span>Watch</span><select id="m-watch">' +
      C.watches.map(function (w) { return '<option value="' + FM.esc(w) + '"' + (w === watch ? " selected" : "") + ">" + FM.esc(w) + " watch</option>"; }).join("") +
      "</select></label>" +
      '<label class="field"><span>Zone</span><select id="m-zone">' + zoneOptions(watch, pre.zoneId) + "</select></label>" +
      "</div>" +
      '<div class="field-row">' +
      '<label class="field"><span>Date</span><input type="date" id="m-date" value="' + day + '" max="' + FM.ymd(new Date()) + '"></label>' +
      '<label class="field"><span>' + (missed ? "About when it was due" : "Time") + '</span><input type="time" id="m-time"></label>' +
      "</div>" +
      (missed
        ? '<label class="field"><span>Why couldn\'t the check be done?</span><textarea id="m-reason" rows="3" placeholder="e.g. only person working the front desk"></textarea></label>'
        : '<label class="field"><span>Reason for manual entry</span><textarea id="m-reason" rows="2" placeholder="e.g. forgot to log it at the time"></textarea></label>');
    var m = FM.modal({
      title: missed ? "Check couldn't be done" : "Manual check entry",
      body: body,
      actions: [
        { label: "Cancel" },
        { label: missed ? "Save" : "Save check", cls: "primary", id: "save", onClick: function (mm) { return saveManual(mm, missed); } }
      ]
    });
    m.body.querySelector("#m-watch").onchange = function (e) {
      m.body.querySelector("#m-zone").innerHTML = zoneOptions(e.target.value);
    };
  }

  async function saveManual(m, missed) {
    FM.guardSave();
    var b = m.body;
    var staffEmail = b.querySelector("#m-staff").value;
    var watch = b.querySelector("#m-watch").value;
    var zoneId = Number(b.querySelector("#m-zone").value);
    var date = b.querySelector("#m-date").value;
    var time = b.querySelector("#m-time").value;
    var reason = b.querySelector("#m-reason").value.trim();
    var problems = [];
    if (!staffEmail) problems.push(missed ? "who was supposed to do it" : "who did the check");
    if (!zoneId) problems.push("a zone");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) problems.push("the date");
    if (!/^\d{1,2}:\d{2}/.test(time)) problems.push("the time");
    if (!reason) problems.push("a reason");
    if (problems.length) { FM.toast("Please enter " + problems.join(", ") + ".", { type: "error" }); return false; }
    var p = date.split("-").map(Number), t = time.split(":").map(Number);
    var at = new Date(p[0], p[1] - 1, p[2], t[0], t[1], 0);
    if (at.getTime() > Date.now() + 60000) { FM.toast("That time is in the future.", { type: "error" }); return false; }
    var z = FM.zoneById(zoneId);
    if (!z) { FM.toast("Pick a zone.", { type: "error" }); return false; }
    var staffName = FM.nameForEmail(staffEmail);
    var day = FM.dayKey(at);

    var dup = day === S.day && S.rows.find(function (r) {
      return r.ZoneId === zoneId && r.Watch === watch && !r.Voided && FM.sameUser(r.StaffEmail, staffEmail) &&
        (r.Method === "Missed") === !!missed && Math.abs(at - new Date(r.CheckedAt)) < 10 * 60000;
    });
    if (dup && !(await FM.confirm(staffName + " already has " + (missed ? "a not-done entry" : "a check") + " on this zone at " + FM.fmtTime(dup.CheckedAt) + ". Save this one too?", "Save anyway"))) return false;

    var row = {
      Title: "Check",
      ZoneId: zoneId,
      ZoneName: z.Title,
      Watch: watch,
      BoardDay: day,
      CheckedAt: at.toISOString(),
      StaffEmail: staffEmail,
      StaffName: staffName,
      EnteredByEmail: me().email,
      EnteredByName: me().name,
      Method: missed ? "Missed" : "Manual",
      Reason: reason,
      Voided: false
    };
    var created = await FM.sp.create(L.checks, row);
    created = Object.assign({ Created: new Date().toISOString() }, row, created || {});
    FM.audit("Checks", missed ? "Missed check recorded" : "Manual check entered", created.Id,
      "zone #" + zoneId + " · " + watch + " watch · for " + staffName + " · " + (missed ? "due about " : "checked at ") + FM.fmtDateTime(at) + " · reason: " + reason);
    var what = missed ? "Not-done entry" : "Manual check";
    if (day === S.day) {
      S.rows.push(created);
      render();
      FM.toast(what + " saved.");
    } else {
      FM.toast(what + " saved for " + FM.fmtDate(day) + ".", { action: { label: "View day", fn: function () { S.followToday = false; load(day).catch(FM.error); } } });
    }
  }

  // ------------------------------------------------------------ wiring
  function onClick(e) {
    var t = e.target.closest("[data-act]");
    if (!t || !panel().contains(t)) return;
    var act = t.getAttribute("data-act");
    var run = function (p) { Promise.resolve(p).catch(function (err) { FM.error(err); }); };
    if (act === "log") run(autoLog(t, Number(t.getAttribute("data-zone")), t.getAttribute("data-watch")));
    else if (act === "chip") showChip(Number(t.getAttribute("data-id")));
    else if (act === "manual" || act === "missed") { try { FM.guardSave(); openManual({ mode: act }); } catch (err) { FM.error(err); } }
    else if (act === "prev") { S.followToday = false; run(load(FM.addDays(S.day, -1))); }
    else if (act === "next") { var n = FM.addDays(S.day, 1); S.followToday = n === today(); run(load(n)); }
    else if (act === "today") { S.followToday = true; run(load(today())); }
  }
  function onChange(e) {
    if (e.target.id === "chk-day" && e.target.value) {
      var v = e.target.value > today() ? today() : e.target.value;
      S.followToday = v === today();
      load(v).catch(FM.error);
    } else if (e.target.getAttribute("data-act") === "voided") {
      S.showVoided = e.target.checked;
      render();
    }
  }

  FM.modules.checks = {
    init: function () {
      var el = panel();
      el.addEventListener("click", onClick);
      el.addEventListener("change", onChange);
      S.day = today();
      el.innerHTML = '<div class="panel-head"><h2>Security checks</h2></div><div class="empty">Loading…</div>';
    },
    // For tests and other modules.
    state: S,
    load: load,
    openManual: openManual,
    render: render
  };

  FM.onTab("dashboard", function () {
    var d = S.followToday ? today() : S.day;
    return load(d);
  });
  FM.onPoll(function (tab) {
    if (tab !== "dashboard") return;
    if (document.querySelector(".modal-wrap")) return;
    return load(S.followToday ? today() : S.day);
  });
})();
