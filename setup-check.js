(window.FM_FILE_VERSIONS = window.FM_FILE_VERSIONS || {})["setup-check.js"] = "2026.10.21-4";

/*
 * Admin -> SharePoint setup check. Compares every FM list against
 * FM_CONFIG.schema: list exists, columns, types, choices, plain text,
 * indexes, and whether permissions are limited to the team.
 */
(function () {
  "use strict";
  var FM = window.FM, C = window.FM_CONFIG;

  var TYPE_NAMES = {
    Text: "Single line of text",
    Note: "Multiple lines of text (plain text)",
    Number: "Number",
    Boolean: "Yes/No",
    DateTime: "Date and time",
    Choice: "Choice"
  };

  function choicesOf(f) {
    var c = f.Choices;
    if (!c) return [];
    if (Array.isArray(c)) return c;
    if (Array.isArray(c.results)) return c.results;
    return [];
  }

  FM.runSetupCheck = async function () {
    var out = [];
    // Which site the app is talking to, and whether this account can add and edit items there.
    try {
      var web = await FM.sp.raw(C.siteUrl + "/_api/web?$select=Title,Url");
      out.push({ level: "ok", list: "(site)", item: "Site", msg: "Connected to \"" + web.Title + "\" at " + web.Url });
      try {
        var perms = await FM.sp.raw(C.siteUrl + "/_api/web/effectiveBasePermissions");
        var low = Number(perms && perms.Low);
        if (!isNaN(low)) {
          var canAdd = (low & 0x2) !== 0, canEdit = (low & 0x4) !== 0;
          out.push({ level: canAdd && canEdit ? "ok" : "warn", list: "(site)", item: "Your access", msg: "Add items: " + (canAdd ? "yes" : "no") + " · Edit items: " + (canEdit ? "yes" : "no") + " (site level; list permissions can differ)" });
        }
      } catch (e) { /* optional */ }
    } catch (e) {
      out.push({ level: "error", list: "(site)", item: "Site", msg: "Couldn't reach " + C.siteUrl + " (" + e.message + "). Check siteUrl in config.js." });
      return out;
    }
    var keys = Object.keys(C.schema);
    await FM.pool(keys, async function (key) {
      var list = C.lists[key], spec = C.schema[key];
      var base = FM.sp.listUrl(list);
      var info;
      try {
        info = await FM.sp.raw(base + "?$select=Title,ItemCount,HasUniqueRoleAssignments");
      } catch (e) {
        out.push({ level: "error", list: list, item: "List", msg: e.status === 404 ? "List not found. Create it with exactly this name." : "Couldn't read the list (" + e.message + ")." });
        return;
      }
      var fields;
      try {
        var r = await FM.sp.raw(base + "/fields");
        fields = (r && r.value) || [];
      } catch (e) {
        out.push({ level: "error", list: list, item: "Columns", msg: "Couldn't read the columns (" + e.message + ")." });
        return;
      }
      var listOk = true;
      spec.columns.forEach(function (col) {
        var f = fields.find(function (x) { return x.InternalName === col.name; });
        if (!f) {
          var byTitle = fields.find(function (x) { return String(x.Title).replace(/\s/g, "").toLowerCase() === col.name.toLowerCase(); });
          out.push({
            level: "error", list: list, item: col.name,
            msg: byTitle
              ? "A column titled \"" + byTitle.Title + "\" exists but its internal name is " + byTitle.InternalName + ". Delete it and recreate it named exactly " + col.name + " (no spaces)."
              : "Missing column. Add it as " + TYPE_NAMES[col.type] + "."
          });
          listOk = false;
          return;
        }
        var type = f.TypeAsString;
        if (type !== col.type) {
          out.push({ level: "error", list: list, item: col.name, msg: "Wrong type: it's " + type + "; it should be " + TYPE_NAMES[col.type] + "." });
          listOk = false;
          return;
        }
        if (col.type === "Note" && f.RichText) {
          out.push({ level: "error", list: list, item: col.name, msg: "Turn off rich text (column settings → Specify the type of text → Plain text)." });
          listOk = false;
        }
        if (col.choices) {
          var have = choicesOf(f);
          var missing = col.choices.filter(function (c) { return have.indexOf(c) < 0; });
          if (missing.length) {
            out.push({ level: "error", list: list, item: col.name, msg: "Missing choice(s): " + missing.join(", ") + ". Choices should be: " + col.choices.join(", ") + "." });
            listOk = false;
          }
        }
        if (col.indexed && !f.Indexed) {
          out.push({ level: "warn", list: list, item: col.name, msg: "Add an index on this column (List settings → Indexed columns). Needed once the list grows past 5,000 items." });
        }
      });
      if (listOk) out.push({ level: "ok", list: list, item: "Columns", msg: spec.columns.length + " columns OK · " + (info.ItemCount || 0) + " items" });
      if (info.HasUniqueRoleAssignments === false) {
        out.push({
          level: spec.restricted === "required" ? "security" : "warn", list: list, item: "Permissions",
          msg: spec.restricted === "required"
            ? "This list still inherits the site's permissions, so anyone with access to the site can read it. Stop inheriting and grant only the Facilities Management team."
            : "Consider limiting this list to the Facilities Management team (stop inheriting permissions)."
        });
      } else if (info.HasUniqueRoleAssignments === true) {
        out.push({ level: "ok", list: list, item: "Permissions", msg: "Has its own permissions. Make sure only the team is granted access." });
      }
    });
    var order = { error: 0, security: 1, warn: 2, ok: 3 };
    var listOrder = ["(site)"].concat(Object.keys(C.lists).map(function (k) { return C.lists[k]; }));
    out.sort(function (a, b) { return order[a.level] - order[b.level] || listOrder.indexOf(a.list) - listOrder.indexOf(b.list); });
    return out;
  };
})();
