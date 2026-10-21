(window.FM_FILE_VERSIONS = window.FM_FILE_VERSIONS || {})["reports.js"] = "2026.10.21-2";

/*
 * Reports tab: security checks by day and watch, by zone and by person
 * (checks that couldn't be done are counted separately, never as checks);
 * tasks by person and type; tracked work opened and closed. Exports run
 * in the browser.
 */
(function () {
  "use strict";
  var FM = window.FM, C = window.FM_CONFIG, L = C.lists;
  var S = { from: null, to: null, data: null, loading: false };

  function panel() { return FM.$("#reports-panel"); }

  async function load() {
    if (S.from > S.to) { var t = S.from; S.from = S.to; S.to = t; }
    S.loading = true; render();
    try {
      var wideStart = FM.startOfDay(FM.addDays(S.from, -1)), wideEnd = FM.startOfDay(FM.addDays(S.to, 2));
      var calStart = FM.startOfDay(S.from), calEnd = FM.startOfDay(FM.addDays(S.to, 1));
      var r = await Promise.all([
        FM.sp.getAll(L.checks, {
          $select: "Id,Created,ZoneId,ZoneName,Watch,BoardDay,CheckedAt,StaffEmail,StaffName,EnteredByEmail,EnteredByName,Method,Reason,Voided,VoidReason,VoidedByEmail",
          $filter: "CheckedAt ge " + FM.odataDate(wideStart) + " and CheckedAt lt " + FM.odataDate(wideEnd)
        }),
        FM.sp.getAll(L.activity, {
          $select: "Id,TaskTypeId,TaskTypeName,Reference,Quantity,StaffEmail,StaffName,LoggedAt,Voided,VoidReason",
          $filter: "LoggedAt ge " + FM.odataDate(calStart) + " and LoggedAt lt " + FM.odataDate(calEnd)
        }),
        FM.sp.getAll(L.work, {
          $select: "Id,Title,TaskTypeName,Reference,Status,AssignedToEmail,AssignedToName,OpenedAt,OpenedByName,ClosedAt",
          $filter: "OpenedAt ge " + FM.odataDate(calStart) + " and OpenedAt lt " + FM.odataDate(calEnd)
        }),
        FM.sp.getAll(L.work, {
          $select: "Id,Title,TaskTypeName,Reference,Status,AssignedToEmail,AssignedToName,OpenedAt,OpenedByName,ClosedAt",
          $filter: "ClosedAt ge " + FM.odataDate(calStart) + " and ClosedAt lt " + FM.odataDate(calEnd)
        })
      ]);
      S.data = {
        checks: r[0].filter(function (c) { return c.BoardDay >= S.from && c.BoardDay <= S.to; }),
        activity: r[1],
        opened: r[2],
        closed: r[3].filter(function (w) { return w.Status !== "Open"; })
      };
    } finally {
      S.loading = false;
      render();
    }
  }

  function zoneName(c) { var z = FM.zoneById(c.ZoneId); return z ? z.Title : c.ZoneName; }
  function inc(map, key, field, n) { map[key] = map[key] || {}; map[key][field] = (map[key][field] || 0) + (n == null ? 1 : n); }

  function summaries() {
    var d = S.data;
    var live = d.checks.filter(function (c) { return !c.Voided; });
    var byDay = {}, byZone = {}, byPerson = {};
    live.forEach(function (c) {
      var zk = zoneName(c) + "\u0000" + c.ZoneId;
      var pk = c.StaffName || FM.nameForEmail(c.StaffEmail);
      if (c.Method === "Missed") {
        inc(byDay, c.BoardDay, "Missed"); inc(byZone, zk, "Missed"); inc(byPerson, pk, "Missed");
        return;
      }
      inc(byDay, c.BoardDay, c.Watch); inc(byDay, c.BoardDay, "Total");
      if (c.Method === "Manual") inc(byDay, c.BoardDay, "Manual");
      inc(byZone, zk, c.Watch); inc(byZone, zk, "Total");
      inc(byPerson, pk, "Total"); if (c.Method === "Manual") inc(byPerson, pk, "Manual");
    });
    var missed = live.filter(function (c) { return c.Method === "Missed"; });
    var tasks = {};
    d.activity.filter(function (a) { return !a.Voided; }).forEach(function (a) {
      inc(tasks, (a.StaffName || FM.nameForEmail(a.StaffEmail)) + "\u0000" + a.TaskTypeName, "Count", Number(a.Quantity) || 1);
    });
    var work = {};
    d.opened.forEach(function (w) { inc(work, w.TaskTypeName || "(none)", "Opened"); });
    d.closed.forEach(function (w) { inc(work, w.TaskTypeName || "(none)", w.Status); });
    return { live: live, missed: missed, byDay: byDay, byZone: byZone, byPerson: byPerson, tasks: tasks, work: work };
  }

  function table(head, rows) {
    if (!rows.length) return '<div class="empty small">Nothing in this period.</div>';
    return '<div class="table-wrap"><table class="grid"><thead><tr>' + head.map(function (h, i) { return "<th" + (i ? ' class="num"' : "") + ">" + FM.esc(h) + "</th>"; }).join("") + "</tr></thead><tbody>" +
      rows.map(function (r) { return "<tr>" + r.map(function (v, i) { return "<td" + (i ? ' class="num"' : "") + ">" + FM.esc(v) + "</td>"; }).join("") + "</tr>"; }).join("") + "</tbody></table></div>";
  }

  function reasonsTable(list) {
    if (!list.length) return '<div class="empty small">None in this period.</div>';
    return '<div class="table-wrap"><table class="grid"><thead><tr><th>Day</th><th>Watch</th><th>Zone</th><th>Who</th><th>Due about</th><th>Reason</th><th>Entered by</th></tr></thead><tbody>' +
      list.slice().sort(function (a, b) { return new Date(b.CheckedAt) - new Date(a.CheckedAt); }).map(function (c) {
        return "<tr><td>" + FM.esc(FM.fmtDate(c.BoardDay)) + "</td><td>" + FM.esc(c.Watch) + "</td><td>" + FM.esc(zoneName(c)) + "</td><td>" + FM.esc(c.StaffName || FM.nameForEmail(c.StaffEmail)) +
          "</td><td>" + FM.esc(FM.fmtTime(c.CheckedAt)) + "</td><td>" + FM.esc(c.Reason || "") + "</td><td>" + FM.esc(c.EnteredByName || FM.nameForEmail(c.EnteredByEmail)) + "</td></tr>";
      }).join("") + "</tbody></table></div>";
  }

  function render() {
    var el = panel(); if (!el) return;
    var html = '<div class="panel-head"><h2>Reports</h2><div class="toolbar">' +
      '<label class="small">From <input type="date" id="r-from" value="' + S.from + '"></label>' +
      '<label class="small">To <input type="date" id="r-to" value="' + S.to + '"></label>' +
      '<button class="btn" data-r="week">Last 7 days</button><button class="btn" data-r="month">Last 30 days</button>' +
      '<button class="btn primary" id="r-export"' + (S.data ? "" : " disabled") + ">Export to Excel</button></div></div>";
    if (S.loading) { el.innerHTML = html + '<div class="empty">Loading…</div>'; return; }
    if (!S.data) { el.innerHTML = html; return; }
    var s = summaries();
    var W = C.watches;
    html += "<h4>Security checks by day</h4>" + table(["Day"].concat(W.map(function (w) { return w + " watch"; }), ["Total", "Manual", "Not done"]),
      Object.keys(s.byDay).sort().reverse().map(function (k) { var r = s.byDay[k]; return [FM.fmtDate(k)].concat(W.map(function (w) { return r[w] || 0; }), [r.Total || 0, r.Manual || 0, r.Missed || 0]); }));
    html += "<h4>Security checks by zone</h4>" + table(["Zone"].concat(W.map(function (w) { return w + " watch"; }), ["Total", "Not done"]),
      Object.keys(s.byZone).sort().map(function (k) { var r = s.byZone[k]; return [k.split("\u0000")[0]].concat(W.map(function (w) { return r[w] || 0; }), [r.Total || 0, r.Missed || 0]); }));
    html += "<h4>Security checks by person</h4>" + table(["Person", "Checks", "Manual", "Not done"],
      Object.keys(s.byPerson).sort().map(function (k) { return [k, s.byPerson[k].Total || 0, s.byPerson[k].Manual || 0, s.byPerson[k].Missed || 0]; }));
    html += "<h4>Checks that couldn't be done</h4>" + reasonsTable(s.missed);
    html += "<h4>Tasks by person</h4>" + table(["Person · task", "Count"],
      Object.keys(s.tasks).sort().map(function (k) { var p = k.split("\u0000"); return [p[0] + " · " + p[1], s.tasks[k].Count]; }));
    html += "<h4>Tracked work</h4>" + table(["Type", "Opened", "Completed", "Cancelled"],
      Object.keys(s.work).sort().map(function (k) { var r = s.work[k]; return [k, r.Opened || 0, r.Completed || 0, r.Cancelled || 0]; }));
    el.innerHTML = html;
  }

  async function exportAll() {
    var s = summaries(), d = S.data, W = C.watches;
    var dayRows = Object.keys(s.byDay).sort().map(function (k) {
      var r = s.byDay[k], o = { "Day": FM.fmtDate(k) };
      W.forEach(function (w) { o[w + " watch"] = r[w] || 0; });
      o.Total = r.Total || 0; o.Manual = r.Manual || 0; o["Not done"] = r.Missed || 0;
      return o;
    });
    var zoneRows = Object.keys(s.byZone).sort().map(function (k) {
      var r = s.byZone[k], o = { "Zone": k.split("\u0000")[0] };
      W.forEach(function (w) { o[w + " watch"] = r[w] || 0; });
      o.Total = r.Total || 0; o["Not done"] = r.Missed || 0;
      return o;
    });
    var logRows = d.checks.slice().sort(function (a, b) { return new Date(a.CheckedAt) - new Date(b.CheckedAt); }).map(function (c) {
      return {
        "Board day": FM.fmtDate(c.BoardDay), "Watch": c.Watch, "Zone": zoneName(c), "Person": c.StaffName || FM.nameForEmail(c.StaffEmail),
        "Time (checked, or due if not done)": FM.fmtDateTime(c.CheckedAt), "Method": c.Method === "Missed" ? "Not done" : c.Method,
        "Entered by": c.Method !== "Auto" ? (c.EnteredByName || FM.nameForEmail(c.EnteredByEmail)) : "",
        "Entered at": c.Method !== "Auto" ? FM.fmtDateTime(c.Created) : "", "Reason": c.Reason || "", "Voided": c.Voided ? "Yes" : "", "Void reason": c.VoidReason || ""
      };
    });
    var taskRows = d.activity.slice().sort(function (a, b) { return new Date(a.LoggedAt) - new Date(b.LoggedAt); }).map(function (a) {
      return { "When": FM.fmtDateTime(a.LoggedAt), "Person": a.StaffName || FM.nameForEmail(a.StaffEmail), "Task": a.TaskTypeName, "Reference": a.Reference || "", "Count": a.Quantity || 1, "Voided": a.Voided ? "Yes" : "" };
    });
    var seen = {};
    var workRows = d.opened.concat(d.closed).filter(function (w) { if (seen[w.Id]) return false; seen[w.Id] = true; return true; }).map(function (w) {
      return { "ID": w.Id, "Title": w.Title, "Type": w.TaskTypeName || "", "Reference": w.Reference || "", "Status": w.Status, "Assigned to": w.AssignedToName || "", "Opened": FM.fmtDateTime(w.OpenedAt), "Opened by": w.OpenedByName || "", "Closed": FM.fmtDateTime(w.ClosedAt) };
    });
    await FM.exportWorkbook("Facilities report " + S.from + " to " + S.to + ".xlsx", [
      { name: "Checks by day", rows: dayRows },
      { name: "Checks by zone", rows: zoneRows },
      { name: "Check log", rows: logRows },
      { name: "Tasks", rows: taskRows },
      { name: "Tracked work", rows: workRows }
    ]);
  }

  FM.modules.reports = {
    init: function () {
      var t = FM.ymd(new Date());
      S.to = t; S.from = FM.addDays(t, -6);
      var el = panel();
      el.addEventListener("change", function (e) {
        if (e.target.id === "r-from" && e.target.value) { S.from = e.target.value; load().catch(function (err) { FM.error(err); }); }
        if (e.target.id === "r-to" && e.target.value) { S.to = e.target.value; load().catch(function (err) { FM.error(err); }); }
      });
      el.addEventListener("click", function (e) {
        var r = e.target.closest("[data-r]");
        if (r) {
          var t2 = FM.ymd(new Date());
          S.to = t2; S.from = FM.addDays(t2, r.getAttribute("data-r") === "week" ? -6 : -29);
          load().catch(function (err) { FM.error(err); });
        }
        if (e.target.id === "r-export") exportAll().catch(function (err) { FM.error(err); });
      });
    },
    state: S,
    load: load
  };
  FM.onTab("reports", load);
})();
