(window.FM_FILE_VERSIONS = window.FM_FILE_VERSIONS || {})["admin.js"] = "2026.10.21-1";

/*
 * Admin tab: zones, task types, team access, settings, setup check.
 * Zone names and descriptions are kept out of the audit text (zones are
 * referenced as "zone #ID"); the audit tab shows the names.
 */
(function () {
  "use strict";
  var FM = window.FM, C = window.FM_CONFIG, L = C.lists;
  var S = { sub: "zones", setupResults: null };

  function panel() { return FM.$("#admin-panel"); }
  function yesNo(v) { return v === false ? '<span class="pill">Retired</span>' : '<span class="pill ok">Active</span>'; }

  async function reloadRef() {
    await FM.loadReference();
    if (FM.modules.checks && FM.modules.checks.state.day) FM.modules.checks.render();
    if (FM.modules.activity) FM.modules.activity.render();
  }

  function render() {
    var el = panel(); if (!el) return;
    var subs = [["zones", "Zones"], ["types", "Task types"], ["team", "Team"], ["settings", "Settings"], ["setup", "SharePoint setup"]];
    var html = '<div class="panel-head"><h2>Admin</h2></div><div class="subtabs">' +
      subs.map(function (s) { return '<button data-sub="' + s[0] + '" class="' + (S.sub === s[0] ? "active" : "") + '">' + s[1] + "</button>"; }).join("") + "</div>";
    html += '<div id="admin-body">' + ({ zones: zonesHtml, types: typesHtml, team: teamHtml, settings: settingsHtml, setup: setupHtml }[S.sub])() + "</div>";
    el.innerHTML = html;
  }

  // -------------------------------------------------------------- zones
  function zonesHtml() {
    var zones = FM.data.zones.slice().sort(FM.bySort);
    var html = '<div class="row-gap" style="margin-bottom:10px"><button class="btn primary" data-a="zone-new" data-save>+ Add zone</button>' +
      '<span class="small muted">Zones are stored in SharePoint only. Retire a zone instead of deleting it; past checks keep its name.</span></div>';
    if (!zones.length) return html + '<div class="empty">No zones yet.</div>';
    return html + '<div class="table-wrap"><table class="grid"><thead><tr><th class="num">Order</th><th>Zone</th><th>Description</th><th>Watches</th><th>Status</th><th></th></tr></thead><tbody>' +
      zones.map(function (z) {
        return '<tr class="' + (FM.isActive(z) ? "" : "dim") + '"><td class="num">' + FM.esc(z.SortOrder == null ? "" : z.SortOrder) + "</td><td><b>" + FM.esc(z.Title) + "</b></td><td>" + FM.esc(z.Description || "") +
          "</td><td>" + FM.esc(FM.zoneWatches(z).join(", ")) + "</td><td>" + yesNo(z.Active) + '</td><td><button class="btn" data-a="zone-edit" data-id="' + z.Id + '" data-save>Edit</button></td></tr>';
      }).join("") + "</tbody></table></div>";
  }

  function editZone(id) {
    FM.guardSave();
    var z = id ? FM.zoneById(id) : null;
    var maxSort = FM.data.zones.reduce(function (m, x) { return Math.max(m, Number(x.SortOrder) || 0); }, 0);
    var watches = z ? FM.zoneWatches(z) : [];
    var body =
      '<label class="field"><span>Zone name</span><input type="text" id="z-title" value="' + FM.esc(z ? z.Title : "") + '"></label>' +
      '<label class="field"><span>Description</span><textarea id="z-desc" rows="2">' + FM.esc(z ? z.Description || "" : "") + "</textarea></label>" +
      '<div class="field"><span>Checked on these watches</span><div class="checks-inline">' +
      C.watches.map(function (w) { return '<label><input type="checkbox" class="z-w" value="' + FM.esc(w) + '"' + (watches.indexOf(w) >= 0 ? " checked" : "") + "> " + FM.esc(w) + "</label>"; }).join("") +
      "</div></div>" +
      '<div class="field-row"><label class="field"><span>Order on the board</span><input type="number" id="z-sort" value="' + (z ? (z.SortOrder == null ? "" : z.SortOrder) : maxSort + 10) + '"></label>' +
      '<label class="field"><span>Status</span><select id="z-active"><option value="1">Active</option><option value="0"' + (z && z.Active === false ? " selected" : "") + ">Retired</option></select></label></div>";
    FM.modal({
      title: z ? "Edit zone" : "Add zone",
      body: body,
      actions: [{ label: "Cancel" }, {
        label: "Save", cls: "primary", id: "save", onClick: async function (m) {
          var b = m.body;
          var next = {
            Title: b.querySelector("#z-title").value.trim(),
            Description: b.querySelector("#z-desc").value.trim(),
            Watches: FM.$$(".z-w", b).filter(function (c) { return c.checked; }).map(function (c) { return c.value; }).join(","),
            SortOrder: b.querySelector("#z-sort").value === "" ? null : Number(b.querySelector("#z-sort").value),
            Active: b.querySelector("#z-active").value === "1"
          };
          if (!next.Title) { FM.toast("Enter a zone name.", { type: "error" }); return false; }
          var clash = FM.data.zones.find(function (x) { return x.Title.trim().toLowerCase() === next.Title.toLowerCase() && (!z || x.Id !== z.Id); });
          if (clash) { FM.toast("Another zone already has that name.", { type: "error" }); return false; }
          if (!next.Watches && next.Active && !(await FM.confirm("This zone isn't on any watch, so it won't appear on the board. Save anyway?", "Save"))) return false;
          if (z) {
            var changed = [];
            if (next.Title !== z.Title) changed.push("name");
            if (next.Description !== (z.Description || "")) changed.push("description");
            if (next.Watches !== FM.zoneWatches(z).join(",")) changed.push("watches " + (FM.zoneWatches(z).join(",") || "none") + " → " + (next.Watches || "none"));
            if (next.SortOrder !== (z.SortOrder == null ? null : z.SortOrder)) changed.push("order " + z.SortOrder + " → " + next.SortOrder);
            if (next.Active !== FM.isActive(z)) changed.push(next.Active ? "restored" : "retired");
            if (!changed.length) return;
            await FM.sp.update(L.zones, z.Id, next);
            FM.audit("Admin", "Zone updated", z.Id, "zone #" + z.Id + " · " + changed.join("; "));
          } else {
            var created = await FM.sp.create(L.zones, next);
            FM.audit("Admin", "Zone added", created.Id, "zone #" + created.Id + " · watches " + (next.Watches || "none"));
          }
          await reloadRef();
          render();
          FM.toast("Zone saved.");
        }
      }]
    });
  }

  // --------------------------------------------------------- task types
  function typesHtml() {
    var types = FM.data.taskTypes.slice().sort(FM.bySort);
    var html = '<div class="row-gap" style="margin-bottom:10px"><button class="btn primary" data-a="type-new" data-save>+ Add task type</button>' +
      '<span class="small muted">No checklist steps = one-click logging on the Dashboard. With steps = tracked work.</span></div>';
    if (!types.length) return html + '<div class="empty">No task types yet.</div>';
    return html + '<div class="table-wrap"><table class="grid"><thead><tr><th class="num">Order</th><th>Task</th><th>Kind</th><th>Input</th><th>Steps</th><th>Status</th><th></th></tr></thead><tbody>' +
      types.map(function (t) {
        var st = FM.typeSteps(t);
        return '<tr class="' + (FM.isActive(t) ? "" : "dim") + '"><td class="num">' + FM.esc(t.SortOrder == null ? "" : t.SortOrder) + "</td><td><b>" + FM.esc(t.Title) + "</b></td><td>" +
          (st.length ? '<span class="pill blue">Tracked work</span>' : '<span class="pill">One-click</span>') + "</td><td>" + (st.length ? "" : FM.esc(t.InputType || "Click Only")) + "</td><td>" +
          FM.esc(st.join(" → ")) + "</td><td>" + yesNo(t.Active) + '</td><td><button class="btn" data-a="type-edit" data-id="' + t.Id + '" data-save>Edit</button></td></tr>';
      }).join("") + "</tbody></table></div>";
  }

  function editType(id) {
    FM.guardSave();
    var t = id ? FM.data.taskTypes.find(function (x) { return x.Id === id; }) : null;
    var maxSort = FM.data.taskTypes.reduce(function (m, x) { return Math.max(m, Number(x.SortOrder) || 0); }, 0);
    var body =
      '<label class="field"><span>Task name</span><input type="text" id="t-title" value="' + FM.esc(t ? t.Title : "") + '"></label>' +
      '<label class="field"><span>Checklist steps (one per line; leave empty for one-click logging)</span><textarea id="t-steps" rows="5">' + FM.esc(t ? FM.typeSteps(t).join("\n") : "") + "</textarea></label>" +
      '<div class="field-row"><label class="field"><span>One-click input</span><select id="t-input">' +
      ["Click Only", "Reference", "Count"].map(function (v) { return "<option" + ((t ? t.InputType || "Click Only" : "Click Only") === v ? " selected" : "") + ">" + v + "</option>"; }).join("") +
      '</select></label><label class="field"><span>Order</span><input type="number" id="t-sort" value="' + (t ? (t.SortOrder == null ? "" : t.SortOrder) : maxSort + 10) + '"></label></div>' +
      '<label class="field"><span>Status</span><select id="t-active"><option value="1">Active</option><option value="0"' + (t && t.Active === false ? " selected" : "") + ">Retired</option></select></label>" +
      '<p class="small muted">Tracked work items always have an optional reference # (work order number).</p>';
    FM.modal({
      title: t ? "Edit task type" : "Add task type",
      body: body,
      actions: [{ label: "Cancel" }, {
        label: "Save", cls: "primary", id: "save", onClick: async function (m) {
          var b = m.body;
          var steps = b.querySelector("#t-steps").value.split(/\r?\n/).map(function (s) { return s.trim(); }).filter(Boolean);
          var next = {
            Title: b.querySelector("#t-title").value.trim(),
            ChecklistSteps: steps.length ? JSON.stringify(steps) : "",
            InputType: b.querySelector("#t-input").value,
            SortOrder: b.querySelector("#t-sort").value === "" ? null : Number(b.querySelector("#t-sort").value),
            Active: b.querySelector("#t-active").value === "1"
          };
          if (!next.Title) { FM.toast("Enter a task name.", { type: "error" }); return false; }
          var clash = FM.data.taskTypes.find(function (x) { return x.Title.trim().toLowerCase() === next.Title.toLowerCase() && (!t || x.Id !== t.Id); });
          if (clash) { FM.toast("Another task type already has that name.", { type: "error" }); return false; }
          if (t) {
            var ch = [];
            if (next.Title !== t.Title) ch.push("name: " + t.Title + " → " + next.Title);
            if (JSON.stringify(steps) !== JSON.stringify(FM.typeSteps(t))) ch.push("steps: " + (FM.typeSteps(t).join(" / ") || "none") + " → " + (steps.join(" / ") || "none"));
            if (next.InputType !== (t.InputType || "Click Only")) ch.push("input: " + (t.InputType || "Click Only") + " → " + next.InputType);
            if (next.SortOrder !== (t.SortOrder == null ? null : t.SortOrder)) ch.push("order " + t.SortOrder + " → " + next.SortOrder);
            if (next.Active !== FM.isActive(t)) ch.push(next.Active ? "restored" : "retired");
            if (!ch.length) return;
            await FM.sp.update(L.taskTypes, t.Id, next);
            FM.audit("Admin", "Task type updated", t.Id, t.Title + " · " + ch.join("; "));
          } else {
            var created = await FM.sp.create(L.taskTypes, next);
            FM.audit("Admin", "Task type added", created.Id, next.Title + (steps.length ? " · steps: " + steps.join(" / ") : " · " + next.InputType));
          }
          await reloadRef();
          render();
          FM.toast("Task type saved.");
        }
      }]
    });
  }

  // --------------------------------------------------------------- team
  function teamHtml() {
    var ms = FM.state.members.slice().sort(function (a, b) { return FM.memberName(a).localeCompare(FM.memberName(b)); });
    var html = '<div class="row-gap" style="margin-bottom:10px"><button class="btn primary" data-a="member-new" data-save>+ Add person</button>' +
      '<span class="small muted">The display name is what shows on checks (for example a last name).</span></div>';
    return html + '<div class="table-wrap"><table class="grid"><thead><tr><th>Display name</th><th>Email</th><th>Role</th><th>Status</th><th></th></tr></thead><tbody>' +
      ms.map(function (m) {
        return '<tr class="' + (FM.isActive(m) ? "" : "dim") + '"><td><b>' + FM.esc(FM.memberName(m)) + "</b></td><td>" + FM.esc(m.Title) + "</td><td>" + FM.esc(m.Role || "Staff") +
          "</td><td>" + (FM.isActive(m) ? '<span class="pill ok">Active</span>' : '<span class="pill">Inactive</span>') + '</td><td><button class="btn" data-a="member-edit" data-id="' + m.Id + '" data-save>Edit</button></td></tr>';
      }).join("") + "</tbody></table></div>";
  }

  function editMember(id) {
    FM.guardSave();
    var p = id ? FM.state.members.find(function (x) { return x.Id === id; }) : null;
    var isMe = p && FM.sameUser(p.Title, FM.state.me.email);
    var body =
      '<label class="field"><span>Work email</span><input type="email" id="p-email" value="' + FM.esc(p ? p.Title : "") + '"' + (p ? " disabled" : "") + "></label>" +
      '<label class="field"><span>Display name</span><input type="text" id="p-name" value="' + FM.esc(p ? p.DisplayName || "" : "") + '"></label>' +
      '<div class="field-row"><label class="field"><span>Role</span><select id="p-role"><option>Staff</option><option' + (p && p.Role === "Admin" ? " selected" : "") + ">Admin</option></select></label>" +
      '<label class="field"><span>Status</span><select id="p-active"><option value="1">Active</option><option value="0"' + (p && p.Active === false ? " selected" : "") + ">Inactive</option></select></label></div>" +
      (isMe ? '<p class="small muted">This is you. You can\'t remove your own admin access or make yourself inactive.</p>' : "");
    FM.modal({
      title: p ? "Edit person" : "Add person",
      body: body,
      actions: [{ label: "Cancel" }, {
        label: "Save", cls: "primary", id: "save", onClick: async function (m) {
          var b = m.body;
          var email = (p ? p.Title : b.querySelector("#p-email").value).trim().toLowerCase();
          var next = {
            DisplayName: b.querySelector("#p-name").value.trim(),
            Role: b.querySelector("#p-role").value,
            Active: b.querySelector("#p-active").value === "1"
          };
          if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) { FM.toast("Enter a valid email.", { type: "error" }); return false; }
          if (isMe && (next.Role !== "Admin" || !next.Active)) { FM.toast("You can't remove your own admin access.", { type: "error" }); return false; }
          if (p) {
            var ch = [];
            if (next.DisplayName !== (p.DisplayName || "")) ch.push("name: " + (p.DisplayName || "none") + " → " + (next.DisplayName || "none"));
            if (next.Role !== (p.Role || "Staff")) ch.push("role: " + (p.Role || "Staff") + " → " + next.Role);
            if (next.Active !== FM.isActive(p)) ch.push(next.Active ? "reactivated" : "made inactive");
            if (!ch.length) return;
            await FM.sp.update(L.members, p.Id, next);
            FM.audit("Admin", "Team member updated", p.Id, email + " · " + ch.join("; "));
          } else {
            if (FM.state.members.some(function (x) { return FM.sameUser(x.Title, email); })) { FM.toast("That person is already on the list.", { type: "error" }); return false; }
            var created = await FM.sp.create(L.members, Object.assign({ Title: email }, next));
            FM.audit("Admin", "Team member added", created.Id, email + " · " + next.Role);
          }
          await FM.loadMembers();
          render();
          FM.toast("Saved.");
        }
      }]
    });
  }

  // ----------------------------------------------------------- settings
  function settingsHtml() {
    var h = FM.dayStartHour();
    var opts = "";
    for (var i = 0; i < 24; i++) {
      var label = i === 0 ? "Midnight (calendar date)" : new Date(2000, 0, 1, i).toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
      opts += '<option value="' + i + '"' + (i === h ? " selected" : "") + ">" + label + "</option>";
    }
    return '<div style="max-width:560px">' +
      '<label class="field"><span>Security check board day starts at</span><select id="set-dayhour">' + opts + "</select></label>" +
      '<p class="small muted">Checks before this time count toward the previous day\'s board. Midnight means each check counts on its calendar date. This is saved in SharePoint, not in the app\'s code.</p>' +
      '<button class="btn primary" data-a="settings-save" data-save>Save settings</button></div>';
  }

  async function saveSettings() {
    FM.guardSave();
    var v = String(Number(FM.$("#set-dayhour").value));
    var cur = FM.data.settings.BoardDayStartHour;
    var old = cur ? cur.value : "0";
    if (String(old) === v) { FM.toast("No changes."); return; }
    if (cur) await FM.sp.update(L.settings, cur.id, { Value: v });
    else await FM.sp.create(L.settings, { Title: "BoardDayStartHour", Value: v });
    FM.audit("Admin", "Setting changed", "BoardDayStartHour", "Board day start hour " + old + " → " + v);
    await reloadRef();
    render();
    FM.toast("Settings saved.");
  }

  // ------------------------------------------------------------- setup
  function setupHtml() {
    var html = '<div class="row-gap" style="margin-bottom:10px"><button class="btn primary" data-a="setup-run">Check SharePoint setup</button>' +
      '<span class="small muted">Checks every FM list, column, type, choice, plain-text setting, index and permission against the app.</span></div>';
    var r = S.setupResults;
    if (!r) return html;
    var bad = r.filter(function (x) { return x.level === "error"; }).length;
    var sec = r.filter(function (x) { return x.level === "security"; }).length;
    var warn = r.filter(function (x) { return x.level === "warn"; }).length;
    html += '<p><b>' + (bad || sec ? "Needs attention: " : "All required items are in place. ") + "</b>" + bad + " problem(s), " + sec + " permission issue(s), " + warn + " suggestion(s).</p>";
    html += '<div class="table-wrap"><table class="grid setup-results"><thead><tr><th>Result</th><th>List</th><th>Item</th><th>Details</th></tr></thead><tbody>' +
      r.map(function (x) {
        var pill = { ok: '<span class="pill ok">OK</span>', error: '<span class="pill bad">Fix</span>', security: '<span class="pill bad">Permissions</span>', warn: '<span class="pill warn">Suggest</span>' }[x.level];
        return "<tr><td>" + pill + "</td><td>" + FM.esc(x.list) + "</td><td>" + FM.esc(x.item || "") + "</td><td>" + FM.esc(x.msg || "") + "</td></tr>";
      }).join("") + "</tbody></table></div>";
    return html;
  }

  // ------------------------------------------------------------ wiring
  function onClick(e) {
    var sub = e.target.closest("[data-sub]");
    if (sub) { S.sub = sub.getAttribute("data-sub"); render(); return; }
    var a = e.target.closest("[data-a]"); if (!a) return;
    var act = a.getAttribute("data-a"), id = Number(a.getAttribute("data-id")) || null;
    var run = function (fn) { try { var p = fn(); if (p && p.catch) p.catch(function (err) { FM.error(err); }); } catch (err) { FM.error(err); } };
    if (act === "zone-new") run(function () { editZone(null); });
    else if (act === "zone-edit") run(function () { editZone(id); });
    else if (act === "type-new") run(function () { editType(null); });
    else if (act === "type-edit") run(function () { editType(id); });
    else if (act === "member-new") run(function () { editMember(null); });
    else if (act === "member-edit") run(function () { editMember(id); });
    else if (act === "settings-save") run(saveSettings);
    else if (act === "setup-run") run(async function () {
      a.disabled = true; a.textContent = "Checking…";
      try { S.setupResults = await FM.runSetupCheck(); } finally { render(); }
    });
  }

  FM.modules.admin = {
    init: function () { panel().addEventListener("click", onClick); },
    state: S,
    render: render
  };
  FM.onTab("admin", function () {
    return Promise.all([FM.loadMembers(), FM.loadReference()]).then(render);
  });
})();
