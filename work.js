(window.FM_FILE_VERSIONS = window.FM_FILE_VERSIONS || {})["work.js"] = "2026.10.21-4";

/*
 * Tracked work: multi-step items (work orders and the like) with
 * assignments, handoffs, notes, complete, reopen and cancel.
 * Shown on the Dashboard ("Open work") and on the Tracked work tab.
 */
(function () {
  "use strict";
  var FM = window.FM, C = window.FM_CONFIG, L = C.lists;
  var SELECT = "Id,Created,Title,TaskTypeId,TaskTypeName,Description,Reference,Status,AssignedToEmail,AssignedToName,StepsState,DueDate,OpenedAt,OpenedByName,ClosedAt";
  var LOG_SELECT = "Id,Created,Title,WorkItemId,Details,StaffEmail,StaffName";
  var S = {
    open: [], closed: [], closedDays: null,
    filter: { status: "Open", who: "anyone", type: "", q: "", days: 90 }
  };

  function me() { return FM.state.me; }
  function steps(item) {
    try { var v = JSON.parse(item.StepsState || "[]"); return Array.isArray(v) ? v : []; } catch (e) { return []; }
  }
  function progress(item) {
    var st = steps(item), done = st.filter(function (s) { return s.done; }).length;
    return { done: done, total: st.length };
  }
  function isOverdue(item) {
    return item.Status === "Open" && item.DueDate && FM.isoToYmd(item.DueDate) < FM.ymd(new Date());
  }
  function assignedLabel(item) { return item.AssignedToEmail ? (item.AssignedToName || FM.nameForEmail(item.AssignedToEmail)) : "Unassigned"; }
  function statusPill(s) {
    return '<span class="pill ' + (s === "Open" ? "blue" : s === "Completed" ? "ok" : "") + '">' + FM.esc(s) + "</span>";
  }

  async function loadOpen() {
    S.open = await FM.sp.getAll(L.work, { $select: SELECT, $filter: "Status eq 'Open'" });
  }
  async function loadClosed(days) {
    var since = new Date(); since.setDate(since.getDate() - days);
    S.closed = await FM.sp.getAll(L.work, { $select: SELECT, $filter: "Status ne 'Open' and ClosedAt ge " + FM.odataDate(since) });
    S.closedDays = days;
  }
  function upsertLocal(item) {
    S.open = S.open.filter(function (x) { return x.Id !== item.Id; });
    S.closed = S.closed.filter(function (x) { return x.Id !== item.Id; });
    (item.Status === "Open" ? S.open : S.closed).push(item);
  }
  function rerender() { renderMine(); if (FM.state.tab === "work") renderTab(); }

  // ------------------------------------------------- dashboard: open work
  function renderMine() {
    var el = FM.$("#mywork-panel"); if (!el) return;
    var list = S.open.filter(function (i) { return !i.AssignedToEmail || FM.sameUser(i.AssignedToEmail, me().email); })
      .sort(function (a, b) {
        var ad = a.DueDate ? new Date(a.DueDate) : null, bd = b.DueDate ? new Date(b.DueDate) : null;
        if (ad && bd) return ad - bd;
        if (ad) return -1;
        if (bd) return 1;
        return new Date(a.OpenedAt || a.Created) - new Date(b.OpenedAt || b.Created);
      });
    var html = '<div class="panel-head"><h2>Open work</h2><div class="toolbar"><button class="btn" data-goto="work">All tracked work</button></div></div>';
    html += '<div class="small muted" style="margin:-6px 0 8px">Assigned to you or unassigned · ' + S.open.length + " open in total</div>";
    if (!list.length) html += '<div class="empty">Nothing open for you.</div>';
    else {
      html += '<ul class="work-list">' + list.map(function (i) {
        var p = progress(i);
        return '<li data-item="' + i.Id + '"><div class="wl-top"><span>' + FM.esc(i.Title) + "</span>" +
          (p.total ? '<span class="small muted">' + p.done + "/" + p.total + "</span>" : "") + "</div>" +
          '<div class="wl-sub">' + FM.esc(i.TaskTypeName || "") + (i.Reference ? " · #" + FM.esc(i.Reference) : "") +
          " · " + FM.esc(assignedLabel(i)) +
          (i.DueDate ? ' · <span class="' + (isOverdue(i) ? "overdue" : "") + '">due ' + FM.esc(FM.fmtDate(i.DueDate)) + "</span>" : "") +
          "</div></li>";
      }).join("") + "</ul>";
    }
    el.innerHTML = html;
  }

  // ------------------------------------------------------ tracked work tab
  function filtered() {
    var f = S.filter;
    var src = f.status === "Open" ? S.open : f.status === "All" ? S.open.concat(S.closed) : S.closed.filter(function (i) { return i.Status === f.status; });
    var q = f.q.trim().toLowerCase();
    return src.filter(function (i) {
      if (f.type && String(i.TaskTypeId) !== f.type) return false;
      if (f.who === "me" && !FM.sameUser(i.AssignedToEmail, me().email)) return false;
      if (f.who === "none" && i.AssignedToEmail) return false;
      if (f.who.indexOf("@") > 0 && !FM.sameUser(i.AssignedToEmail, f.who)) return false;
      if (q && (i.Title + " " + (i.Reference || "") + " " + (i.Description || "")).toLowerCase().indexOf(q) < 0) return false;
      return true;
    }).sort(function (a, b) { return new Date(b.OpenedAt || b.Created) - new Date(a.OpenedAt || a.Created); });
  }

  function renderTab() {
    var el = FM.$("#work-panel"); if (!el) return;
    var f = S.filter;
    var types = FM.data.taskTypes.filter(function (t) { return FM.typeSteps(t).length; });
    var startable = types.filter(FM.isActive);
    var html = '<div class="panel-head"><h2>Tracked work</h2><div class="toolbar">' +
      (startable.length ? '<select id="w-new" data-save><option value="">+ New item…</option>' +
        startable.map(function (t) { return '<option value="' + t.Id + '">' + FM.esc(t.Title) + "</option>"; }).join("") + "</select>" : "") +
      '<button class="btn" id="w-export">Export to Excel</button></div></div>';
    html += '<div class="toolbar" style="margin-bottom:12px">' +
      '<select id="w-status">' + ["Open", "Completed", "Cancelled", "All"].map(function (s) { return "<option" + (f.status === s ? " selected" : "") + ">" + s + "</option>"; }).join("") + "</select>" +
      '<select id="w-who"><option value="anyone">Anyone</option><option value="me"' + (f.who === "me" ? " selected" : "") + '>Assigned to me</option><option value="none"' + (f.who === "none" ? " selected" : "") + ">Unassigned</option>" +
      FM.activeMembers().map(function (m) { return '<option value="' + FM.esc(m.Title) + '"' + (f.who === m.Title ? " selected" : "") + ">" + FM.esc(FM.memberName(m)) + "</option>"; }).join("") + "</select>" +
      '<select id="w-type"><option value="">All types</option>' + types.map(function (t) { return '<option value="' + t.Id + '"' + (f.type === String(t.Id) ? " selected" : "") + ">" + FM.esc(t.Title) + "</option>"; }).join("") + "</select>" +
      '<input type="search" id="w-q" placeholder="Search title or reference" value="' + FM.esc(f.q) + '">' +
      (f.status !== "Open" ? '<select id="w-days">' + [30, 90, 365, 3650].map(function (d) { return '<option value="' + d + '"' + (f.days === d ? " selected" : "") + ">Closed in last " + (d === 3650 ? "10 years" : d + " days") + "</option>"; }).join("") + "</select>" : "") +
      "</div>";
    var rows = filtered();
    if (!rows.length) html += '<div class="empty">No items match.</div>';
    else {
      html += '<div class="table-wrap"><table class="grid"><thead><tr><th>Title</th><th>Type</th><th>Reference</th><th>Assigned to</th><th>Steps</th><th>Due</th><th>Opened</th><th>Status</th></tr></thead><tbody>' +
        rows.map(function (i) {
          var p = progress(i);
          return '<tr class="clickable" data-item="' + i.Id + '"><td>' + FM.esc(i.Title) + "</td><td>" + FM.esc(i.TaskTypeName || "") + "</td><td>" + FM.esc(i.Reference || "") +
            "</td><td>" + FM.esc(assignedLabel(i)) + "</td><td>" +
            (p.total ? '<span class="progress"><i style="width:' + Math.round(100 * p.done / p.total) + '%"></i></span> ' + p.done + "/" + p.total : "") +
            '</td><td class="' + (isOverdue(i) ? "overdue" : "") + '">' + FM.esc(FM.fmtDate(i.DueDate)) + "</td><td>" + FM.esc(FM.fmtDate(i.OpenedAt || i.Created)) +
            "</td><td>" + statusPill(i.Status) + "</td></tr>";
        }).join("") + "</tbody></table></div>";
    }
    el.innerHTML = html;
  }

  async function refreshTab() {
    await loadOpen();
    if (S.filter.status !== "Open" && S.closedDays !== S.filter.days) await loadClosed(S.filter.days);
    rerender();
  }

  async function exportList() {
    var rows = filtered().map(function (i) {
      var p = progress(i);
      return {
        "ID": i.Id, "Title": i.Title, "Type": i.TaskTypeName || "", "Reference": i.Reference || "", "Status": i.Status,
        "Assigned to": i.AssignedToEmail ? assignedLabel(i) : "", "Steps done": p.done, "Steps total": p.total,
        "Due": FM.fmtDate(i.DueDate), "Opened": FM.fmtDateTime(i.OpenedAt || i.Created), "Opened by": i.OpenedByName || "",
        "Closed": FM.fmtDateTime(i.ClosedAt), "Description": i.Description || ""
      };
    });
    await FM.exportWorkbook("Tracked work " + FM.ymd(new Date()) + ".xlsx", [{ name: "Tracked work", rows: rows }]);
  }

  // ------------------------------------------------------------ new item
  function openNew(typeId) {
    FM.guardSave();
    var t = FM.data.taskTypes.find(function (x) { return x.Id === Number(typeId); });
    if (!t) return;
    var body =
      '<label class="field"><span>Title</span><input type="text" id="n-title" value="' + FM.esc(t.Title) + '"></label>' +
      '<div class="field-row"><label class="field"><span>Reference # (work order, etc.)</span><input type="text" id="n-ref"></label>' +
      '<label class="field"><span>Due date (optional)</span><input type="date" id="n-due"></label></div>' +
      '<label class="field"><span>Assign to</span><select id="n-assign">' + FM.memberOptions(me().email, { blank: "Unassigned" }) + "</select></label>" +
      '<label class="field"><span>Description</span><textarea id="n-desc" rows="3"></textarea></label>' +
      '<div class="small muted">Steps: ' + FM.typeSteps(t).map(FM.esc).join(" → ") + "</div>";
    FM.modal({
      title: "New: " + t.Title,
      body: body,
      actions: [{ label: "Cancel" }, {
        label: "Create", cls: "primary", id: "create", onClick: async function (m) {
          var title = m.body.querySelector("#n-title").value.trim();
          if (!title) { FM.toast("Enter a title.", { type: "error" }); return false; }
          var ref = m.body.querySelector("#n-ref").value.trim();
          if (ref) {
            var dup = S.open.find(function (i) { return String(i.Reference || "").toLowerCase() === ref.toLowerCase(); });
            if (dup && !(await FM.confirm("Open item \"" + dup.Title + "\" already has reference " + ref + ". Create another?", "Create anyway"))) return false;
          }
          var assign = m.body.querySelector("#n-assign").value;
          var due = m.body.querySelector("#n-due").value;
          var row = {
            Title: title,
            TaskTypeId: t.Id,
            TaskTypeName: t.Title,
            Description: m.body.querySelector("#n-desc").value.trim(),
            Reference: ref,
            Status: "Open",
            AssignedToEmail: assign || "",
            AssignedToName: assign ? FM.nameForEmail(assign) : "",
            StepsState: JSON.stringify(FM.typeSteps(t).map(function (s) { return { s: s, done: false }; })),
            DueDate: due ? FM.noonIso(due) : null,
            OpenedAt: new Date().toISOString(),
            OpenedByName: me().name
          };
          var created = Object.assign({}, row, (await FM.sp.create(L.work, row)) || {});
          await addLog(created.Id, "Created", (ref ? "Reference " + ref + ". " : "") + (assign ? "Assigned to " + row.AssignedToName + "." : "Unassigned."));
          FM.audit("Work", "Work item created", created.Id, t.Title + " · " + title + (ref ? " · ref " + ref : ""));
          upsertLocal(created);
          rerender();
          FM.toast("Created.", { action: { label: "Open", fn: function () { openItem(created.Id); } } });
        }
      }]
    });
  }

  // ----------------------------------------------------------- item window
  async function addLog(itemId, action, details) {
    try {
      await FM.sp.create(L.workLog, { Title: action, WorkItemId: itemId, Details: details || "", StaffEmail: me().email, StaffName: me().name });
    } catch (e) {
      FM.error(e, "History note didn't save");
    }
  }

  async function fetchItem(id) {
    var item = await FM.sp.get(L.work, id, SELECT);
    upsertLocal(item);
    return item;
  }

  async function openItem(id) {
    var m = FM.modal({ title: "Loading…", wide: true, body: '<div class="empty">Loading…</div>', actions: [{ label: "Close" }],
      onClose: function () { rerender(); } });
    try {
      var data = await Promise.all([fetchItem(id), FM.sp.getAll(L.workLog, { $select: LOG_SELECT, $filter: "WorkItemId eq " + Number(id) })]);
      drawItem(m, data[0], data[1]);
    } catch (e) { m.close(); FM.error(e); }
  }

  function drawItem(m, item, log) {
    m.el.querySelector(".modal-head h3").textContent = item.Title;
    var st = steps(item);
    var open = item.Status === "Open";
    var html = '<dl class="kv">' +
      "<dt>Status</dt><dd>" + statusPill(item.Status) + (item.ClosedAt ? ' <span class="small muted">' + FM.esc(FM.fmtDateTime(item.ClosedAt)) + "</span>" : "") + "</dd>" +
      "<dt>Type</dt><dd>" + FM.esc(item.TaskTypeName || "") + "</dd>" +
      "<dt>Reference</dt><dd>" + FM.esc(item.Reference || "—") + "</dd>" +
      "<dt>Assigned to</dt><dd>" + FM.esc(assignedLabel(item)) + "</dd>" +
      "<dt>Due</dt><dd" + (isOverdue(item) ? ' class="overdue"' : "") + ">" + FM.esc(FM.fmtDate(item.DueDate) || "—") + "</dd>" +
      "<dt>Opened</dt><dd>" + FM.esc(FM.fmtDateTime(item.OpenedAt || item.Created)) + (item.OpenedByName ? " by " + FM.esc(item.OpenedByName) : "") + "</dd>" +
      (item.Description ? "<dt>Description</dt><dd>" + FM.esc(item.Description).replace(/\n/g, "<br>") + "</dd>" : "") +
      "</dl>";
    html += "<h4>Steps</h4>";
    if (!st.length) html += '<div class="muted small">No steps.</div>';
    else html += '<ul class="steps">' + st.map(function (s, i) {
      return '<li class="' + (s.done ? "done" : "") + '"><label><input type="checkbox" data-step="' + i + '"' + (s.done ? " checked" : "") + (open ? " data-save" : " disabled") + '><span><span class="step-text">' + FM.esc(s.s) + "</span>" +
        (s.done && s.by ? '<br><span class="by">' + FM.esc(s.by) + " · " + FM.esc(FM.fmtDateTime(s.at)) + "</span>" : "") + "</span></label></li>";
    }).join("") + "</ul>";
    if (open) {
      html += "<h4>Hand off or note</h4>" +
        '<div class="field-row"><label class="field"><span>Assign to</span><select id="i-assign">' + FM.memberOptions(item.AssignedToEmail, { blank: "Unassigned" }) + "</select></label>" +
        '<label class="field"><span>Note</span><input type="text" id="i-note" placeholder="Optional for a handoff"></label></div>' +
        '<div class="row-gap"><button class="btn" id="i-handoff" data-save>Save assignment</button><button class="btn" id="i-addnote" data-save>Add note</button></div>';
    }
    html += "<h4>History</h4>";
    var hist = log.slice().sort(function (a, b) { return new Date(b.Created) - new Date(a.Created); });
    html += hist.length ? '<ul class="history">' + hist.map(function (h) {
      return "<li><b>" + FM.esc(h.Title) + "</b>" + (h.Details ? " · " + FM.esc(h.Details) : "") + '<div class="h-meta">' + FM.esc(h.StaffName || FM.nameForEmail(h.StaffEmail)) + " · " + FM.esc(FM.fmtDateTime(h.Created)) + "</div></li>";
    }).join("") + "</ul>" : '<div class="muted small">No history yet.</div>';
    m.body.innerHTML = html;

    var reload = async function () {
      var d = await Promise.all([fetchItem(item.Id), FM.sp.getAll(L.workLog, { $select: LOG_SELECT, $filter: "WorkItemId eq " + Number(item.Id) })]);
      drawItem(m, d[0], d[1]);
    };
    var guard = function (fn) { return function (e) { Promise.resolve().then(function () { FM.guardSave(); return fn(e); }).then(reload).catch(function (err) { FM.error(err); reload().catch(function () {}); }); }; };

    FM.$$("[data-step]", m.body).forEach(function (cb) {
      cb.onchange = guard(async function () {
        var i = Number(cb.getAttribute("data-step"));
        var fresh = await FM.sp.get(L.work, item.Id, SELECT);
        var s2 = steps(fresh);
        if (!s2[i]) return;
        s2[i].done = cb.checked;
        if (cb.checked) { s2[i].by = me().name; s2[i].at = new Date().toISOString(); } else { delete s2[i].by; delete s2[i].at; }
        await FM.sp.update(L.work, item.Id, { StepsState: JSON.stringify(s2) });
        await addLog(item.Id, cb.checked ? "Step done" : "Step unchecked", s2[i].s);
        FM.audit("Work", cb.checked ? "Step done" : "Step unchecked", item.Id, s2[i].s);
      });
    });
    var hb = FM.$("#i-handoff", m.body);
    if (hb) hb.onclick = guard(async function () {
      var to = FM.$("#i-assign", m.body).value;
      var note = FM.$("#i-note", m.body).value.trim();
      if (FM.sameUser(to, item.AssignedToEmail) || (!to && !item.AssignedToEmail)) { FM.toast("That's already the assignment."); return; }
      var toName = to ? FM.nameForEmail(to) : "";
      await FM.sp.update(L.work, item.Id, { AssignedToEmail: to, AssignedToName: toName });
      var d = (to ? "To " + toName : "Unassigned") + (item.AssignedToEmail ? " (was " + assignedLabel(item) + ")" : "") + (note ? ". " + note : "");
      await addLog(item.Id, to ? "Handed off" : "Unassigned", d);
      FM.audit("Work", "Assignment changed", item.Id, d);
    });
    var nb = FM.$("#i-addnote", m.body);
    if (nb) nb.onclick = guard(async function () {
      var note = FM.$("#i-note", m.body).value.trim();
      if (!note) { FM.toast("Type a note first.", { type: "error" }); return; }
      await addLog(item.Id, "Note", note);
      FM.audit("Work", "Note added", item.Id, note);
    });

    var actions = [];
    if (open) {
      actions.push({ label: "Edit details", id: "edit", onClick: function () { editDetails(item, reload); return false; } });
      actions.push({
        label: "Cancel item", cls: "danger", id: "cancel", onClick: async function () {
          FM.guardSave();
          var reason = await FM.askReason("Cancel item", "Why is this being cancelled?", "Cancel item");
          if (!reason) return false;
          await FM.sp.update(L.work, item.Id, { Status: "Cancelled", ClosedAt: new Date().toISOString() });
          await addLog(item.Id, "Cancelled", reason);
          FM.audit("Work", "Work item cancelled", item.Id, reason);
          await reload();
          return false;
        }
      });
      actions.push({
        label: "Complete", cls: "primary", id: "complete", onClick: async function () {
          FM.guardSave();
          var fresh = await FM.sp.get(L.work, item.Id, SELECT);
          var left = steps(fresh).filter(function (s) { return !s.done; }).length;
          if (left && !(await FM.confirm(left + " step" + (left === 1 ? " isn't" : "s aren't") + " checked off. Complete anyway?", "Complete"))) return false;
          await FM.sp.update(L.work, item.Id, { Status: "Completed", ClosedAt: new Date().toISOString() });
          await addLog(item.Id, "Completed", left ? left + " step(s) left unchecked" : "");
          FM.audit("Work", "Work item completed", item.Id, item.Title);
          await reload();
          return false;
        }
      });
    } else {
      actions.push({
        label: "Reopen", id: "reopen", onClick: async function () {
          FM.guardSave();
          await FM.sp.update(L.work, item.Id, { Status: "Open", ClosedAt: null });
          await addLog(item.Id, "Reopened", "");
          FM.audit("Work", "Work item reopened", item.Id, item.Title);
          await reload();
          return false;
        }
      });
    }
    actions.push({ label: "Close" });
    m.setActions(actions);
  }

  function editDetails(item, after) {
    FM.guardSave();
    FM.modal({
      title: "Edit details",
      body:
        '<label class="field"><span>Title</span><input type="text" id="e-title" value="' + FM.esc(item.Title) + '"></label>' +
        '<div class="field-row"><label class="field"><span>Reference #</span><input type="text" id="e-ref" value="' + FM.esc(item.Reference || "") + '"></label>' +
        '<label class="field"><span>Due date</span><input type="date" id="e-due" value="' + FM.esc(FM.isoToYmd(item.DueDate)) + '"></label></div>' +
        '<label class="field"><span>Description</span><textarea id="e-desc" rows="3">' + FM.esc(item.Description || "") + "</textarea></label>",
      actions: [{ label: "Cancel" }, {
        label: "Save", cls: "primary", id: "save", onClick: async function (m) {
          var b = m.body;
          var next = {
            Title: b.querySelector("#e-title").value.trim(),
            Reference: b.querySelector("#e-ref").value.trim(),
            DueDate: b.querySelector("#e-due").value ? FM.noonIso(b.querySelector("#e-due").value) : null,
            Description: b.querySelector("#e-desc").value.trim()
          };
          if (!next.Title) { FM.toast("Enter a title.", { type: "error" }); return false; }
          var changes = [];
          if (next.Title !== item.Title) changes.push("title: " + item.Title + " → " + next.Title);
          if (next.Reference !== (item.Reference || "")) changes.push("reference: " + (item.Reference || "none") + " → " + (next.Reference || "none"));
          if (FM.isoToYmd(next.DueDate) !== FM.isoToYmd(item.DueDate)) changes.push("due: " + (FM.fmtDate(item.DueDate) || "none") + " → " + (FM.fmtDate(next.DueDate) || "none"));
          if (next.Description !== (item.Description || "")) changes.push("description changed");
          if (!changes.length) return;
          await FM.sp.update(L.work, item.Id, next);
          await addLog(item.Id, "Details edited", changes.join("; "));
          FM.audit("Work", "Work item edited", item.Id, changes.join("; "));
          await after();
        }
      }]
    });
  }

  // ------------------------------------------------------------ wiring
  FM.modules.work = {
    init: function () {
      FM.$("#mywork-panel").addEventListener("click", function (e) {
        var li = e.target.closest("[data-item]");
        if (li) { openItem(Number(li.getAttribute("data-item"))); return; }
        if (e.target.closest("[data-goto=work]")) FM.showTab("work");
      });
      var wp = FM.$("#work-panel");
      wp.addEventListener("click", function (e) {
        var tr = e.target.closest("[data-item]");
        if (tr) { openItem(Number(tr.getAttribute("data-item"))); return; }
        if (e.target.id === "w-export") exportList().catch(function (err) { FM.error(err); });
      });
      wp.addEventListener("change", function (e) {
        var id = e.target.id, f = S.filter;
        if (id === "w-new") { var v = e.target.value; e.target.value = ""; if (v) { try { openNew(Number(v)); } catch (err) { FM.error(err); } } return; }
        if (id === "w-status") f.status = e.target.value;
        else if (id === "w-who") f.who = e.target.value;
        else if (id === "w-type") f.type = e.target.value;
        else if (id === "w-days") f.days = Number(e.target.value);
        else return;
        refreshTab().catch(function (err) { FM.error(err); });
      });
      wp.addEventListener("input", function (e) {
        if (e.target.id !== "w-q") return;
        S.filter.q = e.target.value;
        var pos = e.target.selectionStart;
        renderTab();
        var q = FM.$("#w-q"); if (q) { q.focus(); try { q.setSelectionRange(pos, pos); } catch (err) { /* ignore */ } }
      });
      FM.$("#mywork-panel").innerHTML = '<div class="panel-head"><h2>Open work</h2></div><div class="empty">Loading…</div>';
    },
    state: S,
    openNew: openNew,
    openItem: openItem,
    refresh: refreshTab
  };

  FM.onTab("dashboard", function () {
    return loadOpen().then(renderMine, function (e) { FM.panelError(FM.$("#mywork-panel"), "Open work", e); });
  });
  FM.onTab("work", function () {
    return refreshTab().catch(function (e) { FM.panelError(FM.$("#work-panel"), "Tracked work", e); });
  });
  FM.onPoll(function (tab) {
    if (tab !== "dashboard" && tab !== "work") return;
    if (document.querySelector(".modal-wrap")) return;
    if (tab === "work" && document.activeElement && document.activeElement.id === "w-q") return;
    return loadOpen().then(rerender, function (e) {
      FM.panelError(FM.$("#mywork-panel"), "Open work", e);
    });
  });
})();
