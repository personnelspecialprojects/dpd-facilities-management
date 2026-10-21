(window.FM_FILE_VERSIONS = window.FM_FILE_VERSIONS || {})["audit.js"] = "2026.10.21-1";

/*
 * Audit log tab (admins). Filter by area and period, search, export.
 * Audit text refers to zones as "zone #ID"; names are filled in here
 * from SharePoint at display time.
 */
(function () {
  "use strict";
  var FM = window.FM, C = window.FM_CONFIG, L = C.lists;
  var S = { rows: [], area: "", days: 7, q: "", loadedDays: null };
  var AREAS = ["Checks", "Tasks", "Work", "Admin", "Session"];

  function panel() { return FM.$("#audit-panel"); }

  function withZoneNames(text) {
    return String(text || "").replace(/zone #(\d+)/g, function (m, id) {
      var z = FM.zoneById(id);
      return z ? z.Title + " (#" + id + ")" : m;
    });
  }

  async function load() {
    var since = new Date(); since.setDate(since.getDate() - S.days);
    S.rows = await FM.sp.getAll(L.audit, { $select: "Id,Created,Title,Area,RecordId,Details,StaffEmail", $filter: "Created ge " + FM.odataDate(since) });
    S.rows.sort(function (a, b) { return new Date(b.Created) - new Date(a.Created); });
    S.loadedDays = S.days;
    render();
  }

  function filtered() {
    var q = S.q.trim().toLowerCase();
    return S.rows.filter(function (r) {
      if (S.area && r.Area !== S.area) return false;
      if (q && (r.Title + " " + withZoneNames(r.Details) + " " + r.StaffEmail + " " + r.RecordId + " " + FM.nameForEmail(r.StaffEmail)).toLowerCase().indexOf(q) < 0) return false;
      return true;
    });
  }

  function render() {
    var el = panel(); if (!el) return;
    var rows = filtered();
    var html = '<div class="panel-head"><h2>Audit log</h2><div class="toolbar">' +
      '<select id="a-area"><option value="">All areas</option>' + AREAS.map(function (a) { return "<option" + (S.area === a ? " selected" : "") + ">" + a + "</option>"; }).join("") + "</select>" +
      '<select id="a-days">' + [[1, "Last 24 hours"], [7, "Last 7 days"], [30, "Last 30 days"], [90, "Last 90 days"], [365, "Last year"]].map(function (d) { return '<option value="' + d[0] + '"' + (S.days === d[0] ? " selected" : "") + ">" + d[1] + "</option>"; }).join("") + "</select>" +
      '<input type="search" id="a-q" placeholder="Search" value="' + FM.esc(S.q) + '">' +
      '<button class="btn" id="a-export">Export to Excel</button></div></div>';
    var queued = FM.auditQueueLength();
    if (queued) html += '<div class="note">' + queued + " audit entr" + (queued === 1 ? "y is" : "ies are") + " waiting in this browser to be saved. They'll be retried automatically.</div>";
    html += '<div class="small muted" style="margin-bottom:8px">' + rows.length + " entr" + (rows.length === 1 ? "y" : "ies") + "</div>";
    if (!rows.length) html += '<div class="empty">No audit entries match.</div>';
    else html += '<div class="table-wrap"><table class="grid"><thead><tr><th>When</th><th>Who</th><th>Area</th><th>Action</th><th>Record</th><th>Details</th></tr></thead><tbody>' +
      rows.slice(0, 1000).map(function (r) {
        return "<tr><td>" + FM.esc(FM.fmtDateTime(r.Created)) + "</td><td>" + FM.esc(FM.nameForEmail(r.StaffEmail)) + "</td><td>" + FM.esc(r.Area) + "</td><td>" + FM.esc(r.Title) +
          "</td><td>" + FM.esc(r.RecordId) + "</td><td>" + FM.esc(withZoneNames(r.Details)) + "</td></tr>";
      }).join("") + "</tbody></table></div>" + (rows.length > 1000 ? '<p class="small muted">Showing the newest 1,000. Export to see all.</p>' : "");
    el.innerHTML = html;
  }

  async function exportRows() {
    var rows = filtered().map(function (r) {
      return { "When": FM.fmtDateTime(r.Created), "Who": FM.nameForEmail(r.StaffEmail), "Email": r.StaffEmail, "Area": r.Area, "Action": r.Title, "Record": r.RecordId, "Details": withZoneNames(r.Details) };
    });
    await FM.exportWorkbook("Audit log " + FM.ymd(new Date()) + ".xlsx", [{ name: "Audit log", rows: rows }]);
  }

  FM.modules.audit = {
    init: function () {
      var el = panel();
      el.addEventListener("change", function (e) {
        if (e.target.id === "a-area") { S.area = e.target.value; render(); }
        else if (e.target.id === "a-days") { S.days = Number(e.target.value); load().catch(function (err) { FM.error(err); }); }
      });
      el.addEventListener("input", function (e) {
        if (e.target.id !== "a-q") return;
        S.q = e.target.value;
        render();
        var q = FM.$("#a-q"); if (q) { q.focus(); q.setSelectionRange(q.value.length, q.value.length); }
      });
      el.addEventListener("click", function (e) {
        if (e.target.id === "a-export") exportRows().catch(function (err) { FM.error(err); });
      });
    },
    state: S,
    load: load,
    withZoneNames: withZoneNames
  };
  FM.onTab("audit", load);
})();
