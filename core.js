(window.FM_FILE_VERSIONS = window.FM_FILE_VERSIONS || {})["core.js"] = "2026.10.23-1";

/*
 * Core: sign-in (popup, as in the Personnel Support portal), SharePoint REST
 * (request digest, throttling retry, token refresh, paging), tabs, polling,
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
    FM.state.blocks[key] = { message: message, reload: !!(opts && opts.reload), signin: !!(opts && opts.signin) };
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
        (b.reload ? '<button class="btn" onclick="location.reload()">Reload</button>' : "") +
        (b.signin ? '<button class="btn" data-reauth>Sign in again</button>' : "") + "</div>";
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
  // Same pattern as the Personnel Support portal: popup sign-in from the
  // button for the identity token (User.Read), then a separate SharePoint
  // token for <origin>/AllSites.Write. On page load, only try silently.
  var pca = null, account = null;
  function spScope() { return new URL(C.siteUrl).origin + "/AllSites.Write"; }
  FM.spScope = spScope;
  function authority() { return "https://login.microsoftonline.com/" + C.tenantId; }

  function getPca() {
    if (pca) return pca;
    if (typeof window.msal === "undefined") return null;
    pca = new window.msal.PublicClientApplication({
      auth: { clientId: C.clientId, authority: authority(), redirectUri: C.redirectUri },
      cache: { cacheLocation: "sessionStorage" }
    });
    return pca;
  }
  // The second CDN (fallback) loads after the app scripts, so wait briefly for it.
  async function waitForMsal() {
    for (var i = 0; i < 40 && typeof window.msal === "undefined"; i++) await sleep(250);
    return getPca();
  }
  function msalText(e) {
    return ((e && e.errorCode) ? e.errorCode + ": " : "") + ((e && (e.errorMessage || e.message)) || String(e));
  }
  function clearStuckInteraction() {
    try {
      Object.keys(sessionStorage).forEach(function (k) { if (/interaction\.status/i.test(k)) sessionStorage.removeItem(k); });
    } catch (x) { /* storage unavailable */ }
  }
  async function acquireSp(allowPopup, forceRefresh) {
    var m = getPca();
    try {
      return (await m.acquireTokenSilent({ scopes: [spScope()], account: account, forceRefresh: !!forceRefresh })).accessToken;
    } catch (e) {
      if (!allowPopup) throw e;
      return (await m.acquireTokenPopup({ scopes: [spScope()], account: account })).accessToken;
    }
  }

  // SharePoint token for every request (MSAL caches and refreshes it).
  FM.getToken = async function (forceRefresh) {
    try {
      return await acquireSp(false, forceRefresh);
    } catch (e) {
      FM.setBlock("session", "Your sign-in has expired. Click Sign in again to keep working.", { signin: true });
      var err = new Error("Your sign-in has expired. Click Sign in again at the top of the page.");
      err.status = 401;
      throw err;
    }
  };
  // From the "Sign in again" banner button (a user click, so a popup is allowed).
  FM.reauth = async function () {
    try {
      // Straight to the popup: an await before it would let the browser block it.
      await getPca().acquireTokenPopup({ scopes: [spScope()], account: account });
      FM.clearBlock("session");
      FM.toast("Signed in again.");
    } catch (e) {
      FM.toast("Sign-in didn't finish: " + msalText(e), { type: "error" });
    }
  };

  // Sign in button.
  FM.signIn = async function () {
    var btn = FM.$("#btn-signin");
    var msg = FM.$("#signin-error");
    if (msg) msg.textContent = "";
    if (btn) { btn.disabled = true; btn.textContent = "Signing in…"; }
    try {
      var m = getPca();
      if (!m) throw new Error("The sign-in library didn't load. Check your connection and reload the page.");
      var login;
      try {
        login = await m.loginPopup({ scopes: [C.loginScope] });
      } catch (e) {
        if (e && e.errorCode === "interaction_in_progress") {
          clearStuckInteraction();
          login = await m.loginPopup({ scopes: [C.loginScope] });
        } else throw e;
      }
      account = login.account;
      m.setActiveAccount(account);
      try {
        await acquireSp(true);
      } catch (e) {
        throw new Error("Signed in, but couldn't get access to SharePoint (" + msalText(e) + ").");
      }
      await startSession();
    } catch (e) {
      if (btn) { btn.disabled = false; btn.textContent = "Sign in"; }
      var text = /popup_window_error|empty_window_error/.test((e && e.errorCode) || "")
        ? "The sign-in window was blocked. Allow pop-ups for this site, then click Sign in again."
        : (e && e.errorCode === "user_cancelled")
          ? "Sign-in was cancelled."
          : "Sign-in didn't finish: " + msalText(e);
      if (msg) msg.textContent = text; else FM.toast(text, { type: "error" });
      if (window.console) console.error(e);
    }
  };
  FM.signOut = async function () {
    FM.data = { zones: [], taskTypes: [], settings: {} };
    FM.state.members = [];
    var m = getPca();
    try {
      if (m && m.logoutPopup) { await m.logoutPopup({ account: account, mainWindowRedirectUri: C.redirectUri }); return; }
    } catch (e) { /* fall through */ }
    try { sessionStorage.clear(); } catch (e) { /* ignore */ }
    location.reload();
  };

  // ---------------------------------------------------------- SharePoint
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  FM.sleep = sleep;

  function readSpError(text) {
    try {
      var j = JSON.parse(text);
      return (j["odata.error"] && j["odata.error"].message && j["odata.error"].message.value) ||
        (j.error && j.error.message && (j.error.message.value || j.error.message)) || "";
    } catch (e) { return String(text || "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().slice(0, 200); }
  }

  async function spFetch(url, opts) {
    opts = opts || {};
    var token = await FM.getToken();
    var send = function () {
      return fetch(url, {
        method: opts.method || "GET",
        headers: Object.assign({ Authorization: "Bearer " + token, Accept: "application/json;odata=nometadata" }, opts.headers || {}),
        body: opts.body
      });
    };
    var res = await send();
    for (var a = 1; (res.status === 429 || res.status === 503) && a <= 8; a++) {
      var asked = Number(res.headers.get("Retry-After"));
      await sleep(Math.min(120, asked > 0 ? asked : Math.pow(2, a)) * 1000);
      res = await send();
    }
    if (res.status === 401) {
      // Expired token: refresh silently and retry once.
      try { token = await FM.getToken(true); res = await send(); } catch (e) { /* reported below */ }
    }
    if (!res.ok) {
      var text = "";
      try { text = await res.text(); } catch (e) { /* ignore */ }
      var msg = readSpError(text);
      var err = new Error("SharePoint " + res.status + (msg ? ": " + msg : ""));
      err.status = res.status;
      err.spMessage = msg;
      throw err;
    }
    if (res.status === 204) return null;
    var t = await res.text();
    return t ? JSON.parse(t) : null;
  }

  // Writes need a form digest (cached until shortly before it expires).
  var digest = null, digestExpires = 0;
  async function getDigest(force) {
    if (!force && digest && Date.now() < digestExpires) return digest;
    var d = await spFetch(C.siteUrl + "/_api/contextinfo", { method: "POST" });
    digest = d.FormDigestValue;
    digestExpires = Date.now() + ((d.FormDigestTimeoutSeconds || 1800) - 120) * 1000;
    return digest;
  }
  async function spWrite(url, method, body) {
    var attempt = async function (force) {
      var headers = { "X-RequestDigest": await getDigest(force), "Content-Type": "application/json;odata=nometadata" };
      if (method === "MERGE") { headers["IF-MATCH"] = "*"; headers["X-HTTP-Method"] = "MERGE"; }
      return spFetch(url, { method: "POST", headers: headers, body: JSON.stringify(body) });
    };
    try { return await attempt(false); }
    catch (e) { if (e.status === 403) return attempt(true); throw e; }
  }

  // Adds the list name and a plain hint to SharePoint errors.
  function explain(list, verb, e) {
    if (e && !e.explained) {
      var hint = "";
      if (e.status === 400) hint = " A column in this list probably doesn't match what the app expects. Admin → SharePoint setup shows which.";
      else if (e.status === 404) hint = " Check that the list exists on " + C.siteUrl + " with exactly this name.";
      else if (e.status === 403) hint = " You don't have permission on this list.";
      e.message = "Couldn't " + verb + " " + list + " (" + e.message + ")." + hint;
      e.list = list;
      e.explained = true;
    }
    return e;
  }

  function qs(o) {
    return Object.keys(o).filter(function (k) { return o[k] != null && o[k] !== ""; })
      .map(function (k) { return k + "=" + encodeURIComponent(o[k]); }).join("&");
  }
  function listUrl(list) {
    return C.siteUrl + "/_api/web/lists/getbytitle('" + encodeURIComponent(String(list).replace(/'/g, "''")) + "')";
  }
  FM.odataStr = function (s) { return "'" + String(s).replace(/'/g, "''") + "'"; };
  FM.odataDate = function (d) { return "datetime'" + (d instanceof Date ? d : new Date(d)).toISOString() + "'"; };

  FM.sp = {
    listUrl: listUrl,
    raw: function (url, opts) { return spFetch(url, opts); },
    getAll: async function (list, q) {
      q = Object.assign({ $top: 2000 }, q || {});
      var url = listUrl(list) + "/items?" + qs(q);
      var out = [], pages = 0;
      try {
        while (url && pages++ < 200) {
          var data = await spFetch(url);
          out = out.concat((data && data.value) || []);
          url = data && (data["odata.nextLink"] || data["@odata.nextLink"]);
        }
      } catch (e) { throw explain(list, "read", e); }
      return out;
    },
    get: function (list, id, select) {
      return spFetch(listUrl(list) + "/items(" + Number(id) + ")" + (select ? "?" + qs({ $select: select }) : ""))
        .catch(function (e) { throw explain(list, "read", e); });
    },
    create: function (list, data, opts) {
      if (!(opts && opts.force)) FM.guardSave();
      return spWrite(listUrl(list) + "/items", "POST", data).catch(function (e) { throw explain(list, "save to", e); });
    },
    update: function (list, id, data) {
      FM.guardSave();
      return spWrite(listUrl(list) + "/items(" + Number(id) + ")", "MERGE", data).catch(function (e) { throw explain(list, "save to", e); });
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
  // A list that can't be read is reported on screen instead of stopping the
  // whole app, so an admin can still reach Admin -> SharePoint setup.
  FM.state.loadErrors = {};
  FM.reportLoadError = function (key, e) {
    if (e) FM.state.loadErrors[key] = e.message || String(e);
    else delete FM.state.loadErrors[key];
    renderLoadErrors();
  };
  function renderLoadErrors() {
    var el = FM.$("#load-errors"); if (!el) return;
    var keys = Object.keys(FM.state.loadErrors);
    el.hidden = !keys.length;
    el.innerHTML = keys.length
      ? '<div class="load-error"><b>Some SharePoint data couldn\'t be loaded.</b><ul>' +
        keys.map(function (k) { return "<li>" + FM.esc(FM.state.loadErrors[k]) + "</li>"; }).join("") + "</ul>" +
        (FM.state.isAdmin ? '<button class="btn" data-goto-setup>Check SharePoint setup</button>' : "<span class=\"small\">Let an admin know.</span>") + "</div>"
      : "";
  }
  FM.loadMembers = async function () {
    FM.state.members = await FM.sp.getAll(C.lists.members, { $select: "Id,Title,DisplayName,Role,Active" });
    return FM.state.members;
  };
  async function readRef(key, select) {
    try {
      var rows = await FM.sp.getAll(C.lists[key], { $select: select });
      FM.reportLoadError(key, null);
      return rows;
    } catch (e) {
      FM.reportLoadError(key, e);
      return null;
    }
  }
  // Task types: if the InputLabel column hasn't been added yet, read without
  // it (default labels are used) instead of failing the whole list.
  FM.state.hasInputLabel = true;
  async function readTaskTypes() {
    var base = "Id,Title,InputType,ChecklistSteps,SortOrder,Active";
    try {
      var rows = await FM.sp.getAll(C.lists.taskTypes, { $select: base + ",InputLabel" });
      FM.state.hasInputLabel = true;
      FM.reportLoadError("taskTypes", null);
      return rows;
    } catch (e) {
      if (e.status !== 400) { FM.reportLoadError("taskTypes", e); return null; }
      var rows2 = await readRef("taskTypes", base);
      if (rows2) FM.state.hasInputLabel = false;
      return rows2;
    }
  }
  FM.loadReference = async function () {
    var r = await Promise.all([
      readRef("zones", "Id,Title,Description,Watches,SortOrder,Active"),
      readTaskTypes(),
      readRef("settings", "Id,Title,Value")
    ]);
    if (r[0]) FM.data.zones = r[0].sort(FM.bySort);
    if (r[1]) FM.data.taskTypes = r[1].sort(FM.bySort);
    if (r[2]) {
      var s = {};
      r[2].forEach(function (row) { s[row.Title] = { id: row.Id, value: row.Value }; });
      FM.data.settings = s;
    }
  };
  // Shows a load failure inside a dashboard panel instead of "Loading…".
  FM.panelError = function (el, title, e) {
    if (!el) return;
    el.innerHTML = '<div class="panel-head"><h2>' + FM.esc(title) + '</h2></div><div class="load-error">' + FM.esc(e && e.message ? e.message : String(e)) +
      (FM.state.isAdmin ? '<div style="margin-top:8px"><button class="btn" data-goto-setup>Check SharePoint setup</button></div>' : "") + "</div>";
  };
  // The label for a task type's entry box (the gray hint text), e.g. "Employee #".
  FM.inputLabel = function (t) {
    var custom = t && t.InputLabel ? String(t.InputLabel).trim() : "";
    if (custom) return custom;
    if (t && t.InputType === "Count" && !FM.typeSteps(t).length) return "Count";
    return "Reference #";
  };
  FM.typeById = function (id) { return FM.data.taskTypes.find(function (t) { return t.Id === Number(id); }) || null; };
  // "Employee # 12345" for display; the default label shows as "#12345".
  FM.refText = function (t, value) {
    if (!value) return "";
    var custom = t && t.InputLabel ? String(t.InputLabel).trim() : "";
    return custom ? custom + " " + value : "#" + value;
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
  // True inside the sign-in popup window while MSAL reads the response.
  function isAuthPopup() {
    try {
      return !!window.opener && window.opener !== window && /(^|[#&?])(code|error)=/.test(location.hash + location.search);
    } catch (e) { return false; }
  }

  function describeAccessError(e) {
    var list = C.lists.members;
    if (e.status === 404) return "The list \"" + list + "\" wasn't found on " + C.siteUrl + ". Check that it exists there with exactly that name.";
    if (e.status === 401 || e.status === 403) return "SharePoint refused access to \"" + list + "\" (" + e.status + "). Make sure your account has at least Read permission on that list.";
    if (e.status === 400) return "The list \"" + list + "\" doesn't have the columns the app expects (" + (e.spMessage || e.message) + "). It needs DisplayName, Role and Active, created with exactly those names.";
    return e.message || String(e);
  }

  var started = false;
  async function startSession() {
    if (started) return;
    FM.state.me = {
      email: String(account.username || "").toLowerCase(),
      graphName: account.name || ""
    };
    showScreen("loading", "Checking access…");
    try {
      await FM.loadMembers();
    } catch (e) {
      showScreen("denied", describeAccessError(e));
      return;
    }
    var mine = FM.state.members.find(function (m) { return FM.sameUser(m.Title, FM.state.me.email) && FM.isActive(m); });
    if (!mine) {
      showScreen("denied", "You're signed in as " + FM.state.me.email + ", but you aren't on the Facilities Management team list (or you're marked inactive). Ask an admin to add you.");
      return;
    }
    started = true;
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
    renderLoadErrors();
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
  }

  FM.boot = async function () {
    if (isAuthPopup()) {
      showScreen("loading", "Finishing sign-in…");
      return;
    }
    FM.$("#btn-signin").onclick = function () { FM.signIn(); };
    FM.$("#btn-signout").onclick = function () { FM.signOut(); };
    document.addEventListener("click", function (e) {
      if (e.target.closest("[data-goto-setup]") && FM.state.isAdmin) {
        FM.showTab("admin");
        if (FM.modules.admin && FM.modules.admin.openSetup) FM.modules.admin.openSetup();
      }
      if (e.target.closest("[data-reauth]")) FM.reauth();
    });
    FM.checkVersions();
    var GUID = /^[\da-f]{8}(-[\da-f]{4}){3}-[\da-f]{12}$/i;
    if (!GUID.test(String(C.clientId)) || !GUID.test(String(C.tenantId))) {
      showScreen("error", "config.js is missing the clientId or tenantId (each is a 36-character ID). Copy them from the Personnel Support portal's config.js, then reload.");
      return;
    }
    showScreen("loading", "Starting…");
    var m = await waitForMsal();
    if (!m) { showScreen("error", "The sign-in library didn't load. Check your connection and reload the page."); return; }
    try {
      if (typeof m.initialize === "function") await m.initialize();
      var accounts = m.getAllAccounts();
      if (!accounts.length) { showScreen("signin"); return; }
      account = (m.getActiveAccount && m.getActiveAccount()) || accounts[0];
      m.setActiveAccount(account);
      try { await acquireSp(false); } catch (e) { showScreen("signin"); return; }
      await startSession();
    } catch (e) {
      showScreen("error", e.message || String(e));
    }
  };
})();
