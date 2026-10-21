(window.FM_FILE_VERSIONS = window.FM_FILE_VERSIONS || {})["core.js"] = "2026.10.21-3";

/*
 * Core: sign-in, SharePoint REST (with throttling retry), tabs, polling,
 * modals and toasts, audit writer with retry queue, version guard,
 * stale-tab detection, date and Excel helpers.
 */
(function () {
  "use strict";
  var C = window.FM_CONFIG;
  var FM = (window.FM = window.FM || {});
  FM.modules = FM.modules || {};
  FM.state = { me: null, member: null, isAdmin: false, members: [], blocks: {}, tab: null, started: false };
  FM.data = { zones: [], taskTypes: [], settings: {} };

  var tabHandlers = {};
  var pollHandlers = [];
  var pollCount = 0;

  // ---------------------------------------------------------------- helpers
  FM.$ = function (sel, root) { return (root || document).querySelector(sel); };
  FM.$$ = function (sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); };
  FM.esc = function (s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  };
  function pad(n) { return String(n).padStart(2, "0"); }
  FM.ymd = function (d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); };
  FM.parseYmd = function (s) {
    var p = String(s).split("-").map(Number);
    return new Date(p[0], p[1] - 1, p[2], 12, 0, 0);
  };
  FM.addDays = function (ymdStr, n) { var d = FM.parseYmd(ymdStr); d.setDate(d.getDate() + n); return FM.ymd(d); };
  FM.startOfDay = function (ymdStr) { var p = String(ymdStr).split("-").map(Number); return new Date(p[0], p[1] - 1, p[2], 0, 0, 0); };
  FM.toDate = function (v) { if (!v) return null; var d = v instanceof Date ? v : new Date(v); return isNaN(d) ? null : d; };
  FM.fmtTime = function (v) {
    var d = FM.toDate(v); if (!d) return "";
    return d.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
  };
  FM.fmtDate = function (v) {
    var d = typeof v === "string" && /^\d{4}-\d{2}-\d{2}$/.test(v) ? FM.parseYmd(v) : FM.toDate(v);
    if (!d) return "";
    return (d.getMonth() + 1) + "/" + d.getDate() + "/" + d.getFullYear();
  };
  FM.fmtDateTime = function (v) { var d = FM.toDate(v); return d ? FM.fmtDate(d) + " " + FM.fmtTime(d) : ""; };
  FM.fmtDayLong = function (ymdStr) {
    return FM.parseYmd(ymdStr).toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric" });
  };
  // A date (no time) stored at midday local time so it never shifts a day.
  FM.noonIso = function (ymdStr) { return ymdStr ? FM.parseYmd(ymdStr).toISOString() : null; };
  FM.isoToYmd = function (v) { var d = FM.toDate(v); return d ? FM.ymd(d) : ""; };

  // The board day a moment belongs to. Admin -> Settings can move the start
  // of the day; the default (0) is the calendar date.
  FM.dayStartHour = function () {
    var s = FM.data.settings.BoardDayStartHour;
    var v = s ? parseInt(s.value, 10) : 0;
    return v >= 0 && v <= 23 ? v : 0;
  };
  FM.dayKey = function (date) {
    var d = new Date(date.getTime());
    d.setHours(d.getHours() - FM.dayStartHour());
    return FM.ymd(d);
  };
  FM.zoneWatches = function (z) {
    return String((z && z.Watches) || "").split(/[,;]/).map(function (s) { return s.trim(); }).filter(Boolean);
  };
  FM.bySort = function (a, b) {
    var x = a.SortOrder == null ? 9999 : a.SortOrder, y = b.SortOrder == null ? 9999 : b.SortOrder;
    return x - y || String(a.Title || "").localeCompare(String(b.Title || ""));
  };
  FM.sameUser = function (a, b) {
    var x = String(a || "").trim().toLowerCase().split("@")[0];
    var y = String(b || "").trim().toLowerCase().split("@")[0];
    return !!x && x === y;
  };
  FM.isActive = function (row) { return row && row.Active !== false; };
  FM.activeMembers = function () {
    return FM.state.members.filter(FM.isActive).sort(function (a, b) { return FM.memberName(a).localeCompare(FM.memberName(b)); });
  };
  FM.memberName = function (m) {
    if (!m) return "";
    return (m.DisplayName && m.DisplayName.trim()) || String(m.Title || "").split("@")[0];
  };
  FM.nameForEmail = function (email) {
    var m = FM.state.members.find(function (x) { return FM.sameUser(x.Title, email); });
    return m ? FM.memberName(m) : String(email || "").split("@")[0];
  };
  FM.memberOptions = function (selected, opts) {
    opts = opts || {};
    var html = opts.blank ? '<option value="">' + FM.esc(opts.blank) + "</option>" : "";
    FM.activeMembers().forEach(function (m) {
      var sel = selected && FM.sameUser(m.Title, selected) ? " selected" : "";
      html += '<option value="' + FM.esc(m.Title) + '"' + sel + ">" + FM.esc(FM.memberName(m)) + "</option>";
    });
    return html;
  };

  // ------------------------------------------------------- toasts and modals
  FM.toast = function (msg, opts) {
    opts = opts || {};
    var root = FM.$("#toast-root"); if (!root) return;
    var el = document.createElement("div");
    el.className = "toast " + (opts.type || "ok");
    el.innerHTML = '<span class="toast-msg"></span>';
    el.firstChild.textContent = msg;
    if (opts.action) {
      var b = document.createElement("button");
      b.className = "toast-action";
      b.textContent = opts.action.label;
      b.onclick = function () { el.remove(); opts.action.fn(); };
      el.appendChild(b);
    }
    var x = document.createElement("button");
    x.className = "toast-x"; x.setAttribute("aria-label", "Dismiss"); x.textContent = "×";
    x.onclick = function () { el.remove(); };
    el.appendChild(x);
    root.appendChild(el);
    setTimeout(function () { el.remove(); }, opts.ms || (opts.type === "error" ? 9000 : 6000));
  };
  FM.error = function (e, prefix) {
    var msg = (prefix ? prefix + ": " : "") + (e && e.message ? e.message : String(e));
    if (window.console) console.warn(msg);
    FM.toast(msg, { type: "error" });
  };

  // FM.modal({ title, body (html), wide, actions: [{ label, cls, onClick(m) }] })
  // onClick may return false (or a promise of false) to keep the window open.
  FM.modal = function (o) {
    var root = FM.$("#modal-root");
    var wrap = document.createElement("div");
    wrap.className = "modal-wrap";
    wrap.innerHTML =
      '<div class="modal' + (o.wide ? " wide" : "") + '" role="dialog" aria-modal="true">' +
      '<div class="modal-head"><h3></h3><button class="modal-x" aria-label="Close">×</button></div>' +
      '<div class="modal-body"></div><div class="modal-foot"></div></div>';
    wrap.querySelector("h3").textContent = o.title || "";
    var body = wrap.querySelector(".modal-body");
    body.innerHTML = o.body || "";
    var foot = wrap.querySelector(".modal-foot");
    var m = {
      el: wrap, body: body,
      close: function () { wrap.remove(); document.removeEventListener("keydown", onKey); if (o.onClose) o.onClose(); },
      setActions: function (actions) {
        foot.innerHTML = "";
        (actions || []).forEach(function (a) {
          var b = document.createElement("button");
          b.className = "btn " + (a.cls || "");
          b.textContent = a.label;
          if (a.id) b.setAttribute("data-modal-act", a.id);
          b.onclick = function () {
            if (b.disabled) return;
            if (!a.onClick) { m.close(); return; }
            b.disabled = true;
            Promise.resolve()
              .then(function () { return a.onClick(m); })
              .then(function (r) { if (r !== false) m.close(); else b.disabled = false; })
              .catch(function (e) { b.disabled = false; FM.error(e); });
          };
          foot.appendChild(b);
        });
      }
    };
    function onKey(e) { if (e.key === "Escape" && wrap.parentNode && root.lastChild === wrap) m.close(); }
    wrap.querySelector(".modal-x").onclick = function () { m.close(); };
    document.addEventListener("keydown", onKey);
    m.setActions(o.actions || [{ label: "Close" }]);
    root.appendChild(wrap);
    var first = body.querySelector("input:not([type=hidden]):not([disabled]),select,textarea");
    if (first && first.focus) first.focus();
    return m;
  };

  FM.confirm = function (message, okLabel, title) {
    return new Promise(function (resolve) {
      var done = false;
      FM.modal({
        title: title || "Please confirm",
        body: '<p class="confirm-msg"></p>',
        onClose: function () { if (!done) resolve(false); },
        actions: [
          { label: "Cancel", onClick: function () { done = true; resolve(false); } },
          { label: okLabel || "OK", cls: "primary", id: "ok", onClick: function () { done = true; resolve(true); } }
        ]
      }).body.querySelector(".confirm-msg").textContent = message;
    });
  };

  // Ask for a reason. Resolves to the text, or null if cancelled.
  FM.askReason = function (title, label, okLabel) {
    return new Promise(function (resolve) {
      var done = false;
      FM.modal({
        title: title,
        body: '<label class="field"><span>' + FM.esc(label || "Reason") + '</span><textarea id="fm-reason" rows="3"></textarea></label>',
        onClose: function () { if (!done) resolve(null); },
        actions: [
          { label: "Cancel", onClick: function () { done = true; resolve(null); } },
          {
            label: okLabel || "OK", cls: "primary", id: "ok", onClick: function (m) {
              var v = m.body.querySelector("#fm-reason").value.trim();
              if (!v) { FM.toast("A reason is required.", { type: "error" }); return false; }
              done = true; resolve(v);
            }
          }
        ]
      });
    });
  };

  // ----------------------------------------------- banners and save blocking
  FM.setBlock = function (key, message, opts) {
    FM.state.blocks[key] = { message: message, reload: !!(opts && opts.reload) };
    renderBanners();
  };
  FM.clearBlock = function (key) { delete FM.state.blocks[key]; renderBanners(); };
  FM.saveBlocked = function () { return Object.keys(FM.state.blocks).length > 0; };
  FM.guardSave = function () {
    if (FM.saveBlocked()) {
      var e = new Error("Saving is turned off until the yellow banner at the top is resolved.");
      e.blocked = true;
      throw e;
    }
  };
  function renderBanners() {
    var el = FM.$("#banner-area"); if (!el) return;
    el.innerHTML = Object.keys(FM.state.blocks).map(function (k) {
      var b = FM.state.blocks[k];
      return '<div class="banner warn" data-banner="' + FM.esc(k) + '"><span>' + FM.esc(b.message) + "</span>" +
        (b.reload ? '<button class="btn" onclick="location.reload()">Reload</button>' : "") + "</div>";
    }).join("");
    document.body.classList.toggle("save-blocked", FM.saveBlocked());
  }

  // --------------------------------------------------------- version guard
  function cssVersion() {
    var sheets = document.styleSheets || [];
    for (var i = 0; i < sheets.length; i++) {
      var rules;
      try { rules = sheets[i].cssRules; } catch (e) { continue; }
      if (!rules) continue;
      for (var j = 0; j < rules.length; j++) {
        var sel = rules[j].selectorText || "";
        var m = sel.match(/\.fm-css-v-([0-9_]+-\d+)/);
        if (m) return m[1].replace(/_/g, ".");
      }
    }
    return null;
  }
  FM.checkVersions = function () {
    var exp = window.FM_EXPECTED_VERSIONS || {};
    var got = Object.assign({}, window.FM_FILE_VERSIONS || {});
    got["styles.css"] = cssVersion();
    var bad = Object.keys(exp).filter(function (f) { return got[f] !== exp[f]; }).sort();
    if (bad.length) {
      FM.setBlock("version",
        "These files don't match this release: " +
        bad.map(function (f) { return f + " (found " + (got[f] || "nothing; missing or cut off") + ", expected " + exp[f] + ")"; }).join("; ") +
        ". Saving is turned off. Re-paste the file in GitHub, wait a minute, then reload.");
    } else FM.clearBlock("version");
    return bad;
  };

  // ---------------------------------------------------- stale-tab detection
  var lastStaleCheck = 0;
  FM.checkStale = function () {
    if (!/^https?:/.test(location.protocol)) return Promise.resolve(false);
    lastStaleCheck = Date.now();
    var url = new URL("index.html", location.href).toString() + "?_=" + Date.now();
    return fetch(url, { cache: "no-store" })
      .then(function (r) { return r.ok ? r.text() : null; })
      .then(function (html) {
        if (!html) return false;
        var m = html.match(/FM_PAGE_VERSION\s*=\s*"([^"]+)"/);
        if (m && m[1] !== window.FM_PAGE_VERSION) {
          FM.setBlock("stale", "A new version of the app (" + m[1] + ") is out. Reload this tab to keep working. Saving is turned off until you do.", { reload: true });
          return true;
        }
        return false;
      })
      .catch(function () { return false; });
  };

  // ------------------------------------------------------------- sign-in
  var pca = null, account = null;
  function redirectUri() { return location.origin + location.pathname.replace(/index\.html$/, ""); }

  async function initAuth() {
    if (!window.msal) throw new Error("The sign-in library didn't load. Check your connection and reload.");
    pca = new window.msal.PublicClientApplication({
      auth: { clientId: C.clientId, authority: C.authority, redirectUri: redirectUri() },
      cache: { cacheLocation: "sessionStorage" }
    });
    if (typeof pca.initialize === "function") await pca.initialize();
    var r = await pca.handleRedirectPromise();
    account = (r && r.account) || pca.getAllAccounts()[0] || null;
    if (account && pca.setActiveAccount) pca.setActiveAccount(account);
  }

  FM.getToken = async function (scope) {
    var req = { scopes: [scope], account: account };
    try {
      return (await pca.acquireTokenSilent(req)).accessToken;
    } catch (e) {
      var code = (e && (e.errorCode || e.name)) || "";
      var interactive = (window.msal.InteractionRequiredAuthError && e instanceof window.msal.InteractionRequiredAuthError) ||
        /interaction_required|login_required|consent_required|InteractionRequired/i.test(code);
      if (interactive) {
        await pca.acquireTokenRedirect(req);
        return new Promise(function () {});
      }
      throw e;
    }
  };

  // Shows the reason on screen instead of failing silently.
  FM.signIn = async function () {
    var btn = FM.$("#btn-signin");
    var msg = FM.$("#signin-error");
    if (msg) msg.textContent = "";
    if (btn) { btn.disabled = true; btn.textContent = "Opening sign-in…"; }
    try {
      if (!pca) throw new Error("The sign-in library didn't start. Reload the page.");
      try {
        await pca.loginRedirect({ scopes: [C.graphScope] });
      } catch (e) {
        // A sign-in that was started earlier and never finished blocks a new one. Clear it and try once more.
        if (e && e.errorCode === "interaction_in_progress") {
          try {
            Object.keys(sessionStorage).forEach(function (k) { if (/interaction\.status/i.test(k)) sessionStorage.removeItem(k); });
          } catch (x) { /* storage unavailable */ }
          await pca.loginRedirect({ scopes: [C.graphScope] });
        } else throw e;
      }
    } catch (e) {
      if (btn) { btn.disabled = false; btn.textContent = "Sign in"; }
      var text = "Sign-in didn't start: " + ((e && (e.errorCode ? e.errorCode + " — " : "") + (e.errorMessage || e.message)) || String(e));
      if (msg) msg.textContent = text; else FM.toast(text, { type: "error" });
      if (window.console) console.error(e);
    }
  };
  FM.signOut = function () {
    FM.data = { zones: [], taskTypes: [], settings: {} };
    return pca.logoutRedirect({ account: account, postLogoutRedirectUri: redirectUri() });
  };

  // ---------------------------------------------------------- SharePoint
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  FM.sleep = sleep;

  async function spFetch(url, opts, attempt) {
    opts = opts || {}; attempt = attempt || 0;
    var token = await FM.getToken(C.spScope);
    var headers = Object.assign({ Authorization: "Bearer " + token, Accept: "application/json;odata=nometadata" }, opts.headers || {});
    var res = await fetch(url, { method: opts.method || "GET", headers: headers, body: opts.body });
    if ((res.status === 429 || res.status === 503) && attempt < 6) {
      var ra = parseInt(res.headers.get("Retry-After"), 10);
      await sleep(ra > 0 ? ra * 1000 : Math.min(30000, 1000 * Math.pow(2, attempt)));
      return spFetch(url, opts, attempt + 1);
    }
    if (!res.ok) {
      var text = "";
      try { text = await res.text(); } catch (e) { /* ignore */ }
      var msg = "";
      try { var j = JSON.parse(text); msg = (j["odata.error"] && j["odata.error"].message && j["odata.error"].message.value) || (j.error && j.error.message && (j.error.message.value || j.error.message)) || ""; } catch (e) { /* not JSON */ }
      var err = new Error("SharePoint " + res.status + (msg ? ": " + msg : ""));
      err.status = res.status;
      throw err;
    }
    if (res.status === 204) return null;
    var t = await res.text();
    return t ? JSON.parse(t) : null;
  }

  function qs(o) {
    return Object.keys(o).filter(function (k) { return o[k] != null && o[k] !== ""; })
      .map(function (k) { return k + "=" + encodeURIComponent(o[k]); }).join("&");
  }
  function listUrl(list) { return C.siteUrl + "/_api/web/lists/GetByTitle('" + String(list).replace(/'/g, "''") + "')"; }
  FM.odataStr = function (s) { return "'" + String(s).replace(/'/g, "''") + "'"; };
  FM.odataDate = function (d) { return "datetime'" + (d instanceof Date ? d : new Date(d)).toISOString() + "'"; };

  FM.sp = {
    listUrl: listUrl,
    raw: function (url, opts) { return spFetch(url, opts); },
    getAll: async function (list, q) {
      q = Object.assign({ $top: 5000 }, q || {});
      var url = listUrl(list) + "/items?" + qs(q);
      var out = [];
      while (url) {
        var data = await spFetch(url);
        out = out.concat((data && data.value) || []);
        url = data && (data["odata.nextLink"] || data["@odata.nextLink"]);
      }
      return out;
    },
    get: function (list, id, select) {
      return spFetch(listUrl(list) + "/items(" + Number(id) + ")" + (select ? "?" + qs({ $select: select }) : ""));
    },
    create: function (list, data, opts) {
      if (!(opts && opts.force)) FM.guardSave();
      return spFetch(listUrl(list) + "/items", {
        method: "POST",
        headers: { "Content-Type": "application/json;odata=nometadata" },
        body: JSON.stringify(data)
      });
    },
    update: function (list, id, data) {
      FM.guardSave();
      return spFetch(listUrl(list) + "/items(" + Number(id) + ")", {
        method: "POST",
        headers: { "Content-Type": "application/json;odata=nometadata", "X-HTTP-Method": "MERGE", "IF-MATCH": "*" },
        body: JSON.stringify(data)
      });
    }
  };

  // Runs async work in batches of 4 (keeps SharePoint from throttling).
  FM.pool = async function (items, fn, size) {
    size = size || 4;
    var results = new Array(items.length), i = 0;
    async function worker() { while (i < items.length) { var k = i++; results[k] = await fn(items[k], k); } }
    var ws = []; for (var w = 0; w < Math.min(size, items.length); w++) ws.push(worker());
    await Promise.all(ws);
    return results;
  };

  // --------------------------------------------- audit writer + retry queue
  // Audit details never include zone names or descriptions (zones are
  // referenced as "zone #ID"), so the browser retry queue holds nothing that
  // describes the building.
  var QKEY = "fm-audit-queue-v1";
  function readQueue() { try { return JSON.parse(localStorage.getItem(QKEY) || "[]"); } catch (e) { return []; } }
  function writeQueue(q) { try { localStorage.setItem(QKEY, JSON.stringify(q)); } catch (e) { /* storage unavailable */ } }

  FM.audit = async function (area, action, recordId, details) {
    var me = FM.state.me || {};
    var row = {
      Title: String(action).slice(0, 255),
      Area: area,
      RecordId: recordId == null ? "" : String(recordId),
      Details: details || "",
      StaffEmail: me.email || ""
    };
    try {
      await FM.sp.create(C.lists.audit, row, { force: true });
    } catch (e) {
      row.Details = (row.Details ? row.Details + " " : "") + "[queued " + new Date().toISOString() + " after a failed save]";
      var q = readQueue(); q.push(row); writeQueue(q);
    }
  };
  FM.auditQueueLength = function () { return readQueue().length; };
  FM.flushAuditQueue = async function () {
    var q = readQueue();
    if (!q.length) return 0;
    var left = [];
    for (var i = 0; i < q.length; i++) {
      try { await FM.sp.create(C.lists.audit, q[i], { force: true }); }
      catch (e) { left = left.concat(q.slice(i)); break; }
    }
    writeQueue(left);
    return q.length - left.length;
  };

  // ------------------------------------------------------- reference data
  FM.loadMembers = async function () {
    FM.state.members = await FM.sp.getAll(C.lists.members, { $select: "Id,Title,DisplayName,Role,Active" });
    return FM.state.members;
  };
  FM.loadReference = async function () {
    var r = await Promise.all([
      FM.sp.getAll(C.lists.zones, { $select: "Id,Title,Description,Watches,SortOrder,Active" }),
      FM.sp.getAll(C.lists.taskTypes, { $select: "Id,Title,InputType,ChecklistSteps,SortOrder,Active" }),
      FM.sp.getAll(C.lists.settings, { $select: "Id,Title,Value" })
    ]);
    FM.data.zones = r[0].sort(FM.bySort);
    FM.data.taskTypes = r[1].sort(FM.bySort);
    var s = {};
    r[2].forEach(function (row) { s[row.Title] = { id: row.Id, value: row.Value }; });
    FM.data.settings = s;
  };
  FM.zoneById = function (id) { return FM.data.zones.find(function (z) { return z.Id === Number(id); }) || null; };
  FM.typeSteps = function (t) {
    if (!t || !t.ChecklistSteps) return [];
    try {
      var v = JSON.parse(t.ChecklistSteps);
      return Array.isArray(v) ? v.map(String).filter(Boolean) : [];
    } catch (e) {
      return String(t.ChecklistSteps).split(/\r?\n/).map(function (s) { return s.trim(); }).filter(Boolean);
    }
  };

  // ------------------------------------------------------- tabs and polling
  var TABS = [
    { id: "dashboard", label: "Dashboard" },
    { id: "work", label: "Tracked work" },
    { id: "reports", label: "Reports" },
    { id: "audit", label: "Audit log", admin: true },
    { id: "admin", label: "Admin", admin: true }
  ];
  FM.onTab = function (tab, fn) { (tabHandlers[tab] = tabHandlers[tab] || []).push(fn); };
  FM.onPoll = function (fn) { pollHandlers.push(fn); };
  FM.showTab = function (id) {
    var t = TABS.find(function (x) { return x.id === id; });
    if (!t || (t.admin && !FM.state.isAdmin)) id = "dashboard";
    FM.state.tab = id;
    FM.$$("#tabs button").forEach(function (b) { b.classList.toggle("active", b.getAttribute("data-tab") === id); });
    FM.$$("section.tab").forEach(function (s) { s.hidden = s.id !== "tab-" + id; });
    (tabHandlers[id] || []).forEach(function (fn) {
      Promise.resolve().then(fn).catch(function (e) { FM.error(e); });
    });
  };
  function renderTabs() {
    var nav = FM.$("#tabs");
    nav.innerHTML = TABS.filter(function (t) { return !t.admin || FM.state.isAdmin; })
      .map(function (t) { return '<button data-tab="' + t.id + '">' + FM.esc(t.label) + "</button>"; }).join("");
    nav.hidden = false;
    nav.onclick = function (e) {
      var b = e.target.closest("button[data-tab]");
      if (b) FM.showTab(b.getAttribute("data-tab"));
    };
  }
  FM.poll = async function () {
    pollCount++;
    try { await FM.flushAuditQueue(); } catch (e) { /* retried next time */ }
    if (pollCount % 5 === 0) {
      try { await Promise.all([FM.loadMembers(), FM.loadReference()]); } catch (e) { /* keep old copies */ }
    }
    for (var i = 0; i < pollHandlers.length; i++) {
      try { await pollHandlers[i](FM.state.tab); } catch (e) { if (window.console) console.warn("Refresh failed", e); }
    }
  };

  // ------------------------------------------------------------ screens
  function showScreen(id, message) {
    FM.$$("section.screen").forEach(function (s) { s.hidden = s.id !== "screen-" + id; });
    if (message != null) {
      var m = FM.$("#screen-" + id + " .screen-msg");
      if (m) m.textContent = message;
    }
    if (id !== "app") FM.$$("section.tab").forEach(function (s) { s.hidden = true; });
  }

  // ----------------------------------------------------------- Excel
  var xlsxPromise = null;
  FM.loadXLSX = function () {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (!xlsxPromise) {
      xlsxPromise = new Promise(function (resolve, reject) {
        var s = document.createElement("script");
        s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
        s.onload = function () { resolve(window.XLSX); };
        s.onerror = function () { xlsxPromise = null; reject(new Error("The Excel library didn't load.")); };
        document.head.appendChild(s);
      });
    }
    return xlsxPromise;
  };
  // sheets: [{ name, rows: [ {col: value} ], columns: [names] (optional) }]
  FM.exportWorkbook = async function (filename, sheets) {
    var X = await FM.loadXLSX();
    var wb = X.utils.book_new();
    sheets.forEach(function (sh) {
      var cols = sh.columns || (sh.rows[0] ? Object.keys(sh.rows[0]) : ["(no rows)"]);
      var ws = X.utils.json_to_sheet(sh.rows, { header: cols });
      if (!sh.rows.length) X.utils.sheet_add_aoa(ws, [cols], { origin: "A1" });
      ws["!cols"] = cols.map(function (c) {
        var w = String(c).length;
        sh.rows.forEach(function (r) { var v = r[c]; if (v != null) w = Math.max(w, String(v).length); });
        return { wch: Math.min(60, w + 2) };
      });
      X.utils.book_append_sheet(wb, ws, String(sh.name).slice(0, 31));
    });
    X.writeFile(wb, filename);
  };

  // ----------------------------------------------------------- boot
  async function graphMe() {
    var token = await FM.getToken(C.graphScope);
    var r = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName", {
      headers: { Authorization: "Bearer " + token }
    });
    if (!r.ok) throw new Error("Couldn't read your account details (Graph " + r.status + ").");
    return r.json();
  }

  FM.boot = async function () {
    FM.$("#btn-signin").onclick = function () { FM.signIn(); };
    FM.$("#btn-signout").onclick = function () { FM.signOut(); };
    FM.checkVersions();
    if (/^PASTE_/.test(String(C.clientId)) || /^PASTE_/.test(String(C.authority))) {
      showScreen("error", "config.js still has the placeholder sign-in values. In GitHub, edit config.js and replace clientId and authority with the values from the Personnel Support portal's config.js, then reload.");
      return;
    }
    showScreen("loading", "Signing in…");
    try {
      await initAuth();
    } catch (e) {
      showScreen("error", e.message);
      return;
    }
    if (!account) { showScreen("signin"); return; }

    try {
      var me = await graphMe();
      FM.state.me = {
        email: String(me.mail || me.userPrincipalName || account.username || "").toLowerCase(),
        graphName: me.displayName || ""
      };
      showScreen("loading", "Checking access…");
      try {
        await FM.loadMembers();
      } catch (e) {
        if (e.status === 401 || e.status === 403 || e.status === 404) {
          showScreen("denied", "You don't have access to this app's data. Ask a Facilities Management admin to add you.");
          return;
        }
        throw e;
      }
      var mine = FM.state.members.find(function (m) { return FM.sameUser(m.Title, FM.state.me.email) && FM.isActive(m); });
      if (!mine) {
        showScreen("denied", "You're signed in as " + FM.state.me.email + ", but you aren't on the Facilities Management team list. Ask an admin to add you.");
        return;
      }
      FM.state.member = mine;
      FM.state.isAdmin = mine.Role === "Admin";
      FM.state.me.name = FM.memberName(mine) || FM.state.me.graphName;
      FM.$("#user-name").textContent = FM.state.me.name + (FM.state.isAdmin ? " (Admin)" : "");
      FM.$("#user-box").hidden = false;

      showScreen("loading", "Loading…");
      await FM.loadReference();
      Object.keys(FM.modules).sort().forEach(function (k) {
        var mod = FM.modules[k];
        if (mod.init) { try { mod.init(); } catch (e) { FM.error(e, k + " failed to start"); } }
      });
      showScreen("app");
      renderTabs();
      FM.state.started = true;
      FM.showTab("dashboard");

      FM.flushAuditQueue().catch(function () {});
      FM.audit("Session", "Signed in", "", "");
      setInterval(function () { if (!document.hidden) FM.poll(); }, C.pollSeconds * 1000);
      setInterval(FM.checkStale, C.staleCheckSeconds * 1000);
      var lastFocusPoll = Date.now();
      window.addEventListener("focus", function () {
        if (Date.now() - lastStaleCheck > 30000) FM.checkStale();
        if (FM.state.started && Date.now() - lastFocusPoll > 20000) { lastFocusPoll = Date.now(); FM.poll(); }
      });
    } catch (e) {
      showScreen("error", e.message || String(e));
    }
  };
})();
