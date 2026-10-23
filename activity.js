(window.FM_FILE_VERSIONS = window.FM_FILE_VERSIONS || {})["activity.js"] = "2026.10.23-1";

/*
 * Task logging (Dashboard, right side). Task types are set up in Admin.
 * A task type with no checklist steps is logged in one click (optionally
 * with a reference number or a count). A task type with steps starts a
 * tracked work item (work.js).
 */
(function () {
  "use strict";
  var FM = window.FM, C = window.FM_CONFIG, L = C.lists;
  var SELECT = "Id,Created,TaskTypeId,TaskTypeName,Reference,Quantity,StaffEmail,StaffName,LoggedAt,Voided,VoidReason";
  var S = { rows: [], day: null, busy: false };

  function panel() { return FM.$("#activity-panel"); }
  function me() { return FM.state.me; }
  function quickTypes() { return FM.data.taskTypes.filter(function (t) { return FM.isActive(t) && !FM.typeSteps(t).length; }).sort(FM.bySort); }
  function trackedTypes() { return FM.data.taskTypes.filter(function (t) { return FM.isActive(t) && FM.typeSteps(t).length; }).sort(FM.bySort); }

  async function load() {
    var day = FM.ymd(new Date());
    var start = FM.startOfDay(day);
    try {
      S.rows = await FM.sp.getAll(L.activity, { $select: SELECT, $filter: "LoggedAt ge " + FM.odataDate(start) });
    } catch (e) {
      FM.panelError(panel(), "Tasks", e);
      return;
    }
    S.day = day;
    render();
  }

  function render() {
    var el = panel(); if (!el) return;
    var quick = quickTypes(), tracked = trackedTypes();
    var html = '<div class="panel-head"><h2>Tasks</h2></div>';
    if (FM.state.loadErrors.taskTypes) {
      html += '<div class="load-error">The task types couldn\'t be loaded. See the message at the top of the page.</div>';
    } else if (!quick.length && !tracked.length) {
      html += '<div class="empty">No task types are set up yet. ' + (FM.state.isAdmin ? "Add them in Admin → Task types." : "Ask an admin to add them.") + "</div>";
    }
    if (quick.length) {
      html += '<div class="task-list">';
      quick.forEach(function (t) {
        var input = "", lbl = FM.inputLabel(t);
        if (t.InputType === "Reference") input = '<input type="text" name="ref" placeholder="' + FM.esc(lbl) + '" title="' + FM.esc(lbl) + '" aria-label="' + FM.esc(lbl + " for " + t.Title) + '">';
        else if (t.InputType === "Count") input = '<input type="number" name="qty" min="1" step="1" placeholder="' + FM.esc(lbl) + '" title="' + FM.esc(lbl) + '" aria-label="' + FM.esc(lbl + " for " + t.Title) + '">';
        else input = "<span></span>";
        html += '<form class="task-row" data-type="' + t.Id + '"><span class="task-name">' + FM.esc(t.Title) + "</span>" + input +
          '<button class="btn primary" type="submit" data-save>Log</button></form>';
      });
      html += "</div>";
    }
    if (tracked.length) {
      html += "<h4>Start tracked work</h4><div class=\"row-gap\">" +
        tracked.map(function (t) { return '<button class="btn" data-start="' + t.Id + '" data-save>+ ' + FM.esc(t.Title) + "</button>"; }).join("") +
        "</div>";
    }
    var mine = S.rows.filter(function (r) { return FM.sameUser(r.StaffEmail, me().email); })
      .sort(function (a, b) { return new Date(b.LoggedAt) - new Date(a.LoggedAt); }).slice(0, 12);
    html += "<h4>Your entries today</h4>";
    if (!mine.length) html += '<div class="muted small">Nothing logged yet today.</div>';
    else {
      html += '<ul class="recent">' + mine.map(function (r) {
        return '<li class="' + (r.Voided ? "voided" : "") + '"><span>' + FM.esc(r.TaskTypeName) +
          (r.Reference ? " · " + FM.esc(FM.refText(FM.typeById(r.TaskTypeId), r.Reference)) : "") + (r.Quantity > 1 ? " × " + r.Quantity : "") +
          ' <span class="muted">' + FM.esc(FM.fmtTime(r.LoggedAt)) + "</span></span>" +
          (r.Voided ? '<span class="pill">Undone</span>' : '<button class="btn link" data-undo="' + r.Id + '" data-save>Undo</button>') + "</li>";
      }).join("") + "</ul>";
    }
    el.innerHTML = html;
  }

  async function logTask(form) {
    FM.guardSave();
    if (S.busy) return;
    var t = FM.data.taskTypes.find(function (x) { return x.Id === Number(form.getAttribute("data-type")); });
    if (!t) return;
    var ref = "", qty = 1;
    if (t.InputType === "Reference") {
      ref = form.querySelector("[name=ref]").value.trim();
      if (!ref) { FM.toast("Enter the " + FM.inputLabel(t) + " first.", { type: "error" }); return; }
    } else if (t.InputType === "Count") {
      qty = Number(form.querySelector("[name=qty]").value);
      if (!(qty >= 1) || Math.floor(qty) !== qty) { FM.toast(FM.inputLabel(t) + ": enter a whole number of 1 or more.", { type: "error" }); return; }
    }
    var now = new Date();
    if (ref) {
      var dupRef = S.rows.find(function (r) { return r.TaskTypeId === t.Id && !r.Voided && String(r.Reference || "").trim().toLowerCase() === ref.toLowerCase(); });
      if (dupRef && !(await FM.confirm(FM.refText(t, ref) + " was already logged today by " + (dupRef.StaffName || FM.nameForEmail(dupRef.StaffEmail)) + " at " + FM.fmtTime(dupRef.LoggedAt) + ". Log it again?", "Log again"))) return;
    } else {
      var dupClick = S.rows.find(function (r) { return r.TaskTypeId === t.Id && !r.Voided && FM.sameUser(r.StaffEmail, me().email) && now - new Date(r.LoggedAt) < 60000; });
      if (dupClick && !(await FM.confirm("You logged " + t.Title + " less than a minute ago. Log another?", "Log again"))) return;
    }
    S.busy = true;
    var btn = form.querySelector("button"); if (btn) btn.disabled = true;
    try {
      var row = {
        Title: t.Title,
        TaskTypeId: t.Id,
        TaskTypeName: t.Title,
        Reference: ref,
        Quantity: qty,
        StaffEmail: me().email,
        StaffName: me().name,
        LoggedAt: now.toISOString(),
        Voided: false
      };
      var created = Object.assign({}, row, (await FM.sp.create(L.activity, row)) || {});
      S.rows.push(created);
      FM.audit("Tasks", "Task logged", created.Id, t.Title + (ref ? " · " + FM.inputLabel(t) + " " + ref : "") + (qty > 1 ? " · count " + qty : ""));
      render();
      FM.toast(t.Title + " logged.");
      var again = FM.$('form[data-type="' + t.Id + '"] input', panel());
      if (again) again.focus();
    } finally {
      S.busy = false;
      if (btn && btn.isConnected) btn.disabled = false;
    }
  }

  async function undo(id) {
    FM.guardSave();
    var r = S.rows.find(function (x) { return x.Id === id; });
    if (!r || r.Voided) return;
    if (!(await FM.confirm("Undo " + r.TaskTypeName + (r.Reference ? " (" + r.Reference + ")" : "") + " logged at " + FM.fmtTime(r.LoggedAt) + "?", "Undo"))) return;
    await FM.sp.update(L.activity, r.Id, { Voided: true, VoidReason: "Undone by staff" });
    r.Voided = true; r.VoidReason = "Undone by staff";
    FM.audit("Tasks", "Task entry voided", r.Id, r.TaskTypeName + (r.Reference ? " · ref " + r.Reference : "") + " · undone by staff");
    render();
  }

  FM.modules.activity = {
    init: function () {
      var el = panel();
      el.addEventListener("submit", function (e) {
        var f = e.target.closest("form.task-row"); if (!f) return;
        e.preventDefault();
        logTask(f).catch(function (err) { FM.error(err); });
      });
      el.addEventListener("click", function (e) {
        var u = e.target.closest("[data-undo]");
        if (u) { undo(Number(u.getAttribute("data-undo"))).catch(function (err) { FM.error(err); }); return; }
        var s = e.target.closest("[data-start]");
        if (s) {
          try { FM.guardSave(); FM.modules.work.openNew(Number(s.getAttribute("data-start"))); } catch (err) { FM.error(err); }
        }
      });
      el.innerHTML = '<div class="panel-head"><h2>Tasks</h2></div><div class="empty">Loading…</div>';
    },
    state: S,
    load: load,
    render: render
  };

  FM.onTab("dashboard", load);
  FM.onPoll(function (tab) {
    if (tab !== "dashboard") return;
    var el = panel();
    // Don't redraw while someone is typing a reference or count.
    if (el && el.contains(document.activeElement) && document.activeElement.tagName === "INPUT" && document.activeElement.value) return;
    return load();
  });
})();
