/* Pine Hills Therapy Monitoring System — Excel storage for the Microsoft Teams version.
   The app was written against a small database API (collection/doc, add/set/delete, onSnapshot).
   This file provides that same API, backed by three Excel tables in the team's Files:
     Records  — one row per therapy session
     Team     — therapist avatars and roles
     Settings — small app settings (pet name)
   Every teacher signs in with their own school Microsoft account (inside Teams this is automatic). */
(function(){
  "use strict";
  const CFG = window.TIR_CONFIG || {};
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = ["Files.ReadWrite.All", "User.Read"];

  /* ---------- lookups shared with the app (kept in step with TYPES / PROGRESS / MOOD there) ---------- */
  const TYPES = { ot:"Occupational Therapy", pt:"Physiotherapy", speech:"Speech & Language", sensory:"Sensory Regulation", motor:"Fine & Gross Motor", behaviour:"Behavioural Intervention" };
  const PROGRESS = ["Not yet","Emerging","Developing","Nearly there","Achieved"];
  const MOOD = ["Upset","Unsettled","Okay","Engaged","Super keen"];
  const lc = s => String(s == null ? "" : s).trim().toLowerCase();
  const typeId = v => { const s = lc(v); if (!s) return ""; for (const [k, n] of Object.entries(TYPES)) if (lc(n) === s || k === s) return k; return String(v); };
  const scaleIdx = (list, v) => { if (v === "" || v == null) return null; if (typeof v === "number") return v >= 1 && v <= 5 ? Math.round(v) : null; const s = lc(v); const i = list.findIndex(x => lc(x) === s); if (i >= 0) return i + 1; const n = parseInt(s, 10); return n >= 1 && n <= 5 ? n : null; };

  /* ---------- cell conversion ---------- */
  const pad = n => String(n).padStart(2, "0");
  const stamp = ms => { if (!ms) return ""; const d = new Date(ms); return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()) + " " + pad(d.getHours()) + ":" + pad(d.getMinutes()); };
  const serialToLocal = n => { const d = new Date(Math.round((n - 25569) * 86400000)); return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds()); };
  const readDate = v => { if (typeof v === "number" && v > 0) { const d = serialToLocal(v); return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); } const s = String(v || "").trim(); const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); return m ? m[1] + "-" + pad(m[2]) + "-" + pad(m[3]) : s; };
  const readTime = v => { if (typeof v === "number" && v > 0) return serialToLocal(v).getTime(); const t = Date.parse(String(v || "").replace(" ", "T")); return isNaN(t) ? null : t; };
  // Text that Excel would turn into a number, date or formula is stored as plain text.
  const txt = v => { const s = String(v == null ? "" : v); if (!s) return ""; return (/^[=+\-@]/.test(s) || (/^[\d\s.,:\/\-+()%$]+$/.test(s) && /\d/.test(s))) ? "'" + s : s; };
  const str = v => (v == null ? "" : String(v).replace(/^'/, ""));

  const SPECS = {
    records: {
      table: "Records",
      cols: ["ID","Date","Student","Class / Year","Type of support","Therapist / Teacher","Minutes","Goal","Activities","More details","Progress","Engagement and mood","Observations","Next steps and home ideas","Created","Last updated","Entered by","Last edited by"],
      toRow(id, d, prev) {
        return [id, d.date || "", txt(d.student), txt(d.klass), TYPES[d.type] || txt(d.type), txt(d.therapist),
          (typeof d.duration === "number" && isFinite(d.duration)) ? d.duration : "", txt(d.goal), txt((d.tags || []).join("; ")), txt(d.activities),
          d.progress ? PROGRESS[d.progress - 1] || "" : "", d.mood ? MOOD[d.mood - 1] || "" : "", txt(d.notes), txt(d.next),
          stamp(d.createdAt), stamp(d.updatedAt), (prev && prev[16]) || me.email, me.email];
      },
      fromRow(r) {
        const tags = str(r[8]).split(/;\s*/).map(s => s.trim()).filter(Boolean);
        const dur = typeof r[6] === "number" ? r[6] : (parseFloat(r[6]) || null);
        return { student: str(r[2]), klass: str(r[3]), date: readDate(r[1]), therapist: str(r[5]), type: typeId(r[4]), duration: dur,
          goal: str(r[7]), tags, activities: str(r[9]), progress: scaleIdx(PROGRESS, r[10]), mood: scaleIdx(MOOD, r[11]),
          notes: str(r[12]), next: str(r[13]), createdAt: readTime(r[14]), updatedAt: readTime(r[15]), enteredBy: str(r[16]) };
      }
    },
    therapists: {
      table: "Team", cols: ["ID","Name","Role","Data"],
      toRow: (id, d) => [id, txt(d.name), txt(d.role), JSON.stringify(d)],
      fromRow(r) { let d = {}; try { d = JSON.parse(str(r[3]) || "{}"); } catch (e) {} return { ...d, name: str(r[1]) || d.name || "", role: str(r[2]) || d.role || "" }; }
    },
    settings: {
      table: "Settings", cols: ["Key","Value"],
      toRow: (id, d) => [id, JSON.stringify(d)],
      fromRow(r) { try { return JSON.parse(str(r[1]) || "{}"); } catch (e) { return {}; } }
    }
  };

  /* ---------- sign-in ---------- */
  let pca = null, account = null, inTeams = false, loginHint = "";
  const me = { email: "", name: "" };
  async function initAuth() {
    const within = (p, ms) => Promise.race([p, new Promise((_, no) => setTimeout(() => no(new Error("timeout")), ms))]);
    try { if (window.parent === window && !window.opener) throw new Error("not in Teams"); await within(window.microsoftTeams.app.initialize(), 4000); inTeams = true; const ctx = await within(window.microsoftTeams.app.getContext(), 4000); loginHint = (ctx.user && (ctx.user.loginHint || ctx.user.userPrincipalName)) || ""; } catch (e) { inTeams = false; }
    const base = location.origin + location.pathname.replace(/[^/]*$/, "");
    const cfg = { auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId, redirectUri: base }, cache: { cacheLocation: "localStorage" } };
    if (inTeams) { try { pca = await window.msal.createNestablePublicClientApplication(cfg); } catch (e) { pca = null; } }
    if (!pca) { pca = new window.msal.PublicClientApplication(cfg); await pca.initialize(); try { const r = await pca.handleRedirectPromise(); if (r && r.account) account = r.account; } catch (e) {} }
    account = account || pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
    if (account) pca.setActiveAccount(account);
  }
  function remember(r) { if (r && r.account) { account = r.account; try { pca.setActiveAccount(account); } catch (e) {} me.email = account.username || me.email; me.name = account.name || me.name; } return r.accessToken; }
  async function getToken(interactive, force) {
    if (account) { try { return remember(await pca.acquireTokenSilent({ scopes: SCOPES, account, forceRefresh: !!force })); } catch (e) {} }
    if (inTeams) { try { return remember(await pca.acquireTokenSilent({ scopes: SCOPES })); } catch (e) {} }
    if (loginHint) { try { return remember(await pca.ssoSilent({ scopes: SCOPES, loginHint })); } catch (e) {} }
    if (!interactive) throw { code: "needs_signin" };
    if (inTeams) return remember(await pca.acquireTokenPopup({ scopes: SCOPES, loginHint: loginHint || undefined }));
    await pca.loginRedirect({ scopes: SCOPES, loginHint: loginHint || undefined, redirectStartPage: location.href });
    throw { code: "redirecting" };
  }

  /* ---------- Microsoft Graph ---------- */
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  async function graph(method, path, body, opts) {
    opts = opts || {};
    for (let attempt = 0; ; attempt++) {
      const token = await getToken(false, attempt > 0 && opts.lastStatus === 401);
      const headers = { Authorization: "Bearer " + token };
      let payload;
      if (body instanceof Blob || body instanceof ArrayBuffer) { headers["Content-Type"] = "application/octet-stream"; payload = body; }
      else if (body !== undefined) { headers["Content-Type"] = "application/json"; payload = JSON.stringify(body); }
      let res;
      try { res = await fetch(GRAPH + path, { method, headers, body: payload, cache: "no-store" }); }
      catch (e) { if (attempt < 2) { await sleep(800 * (attempt + 1)); continue; } throw { code: "network", message: "Can’t reach Microsoft 365. Check the internet connection." }; }
      if (res.ok) { if (res.status === 204) return null; const ct = res.headers.get("content-type") || ""; return ct.includes("json") ? res.json() : res; }
      let err = {}; try { err = (await res.json()).error || {}; } catch (e) {}
      const retryable = res.status === 429 || res.status === 409 || res.status >= 500;
      if (retryable && attempt < 3) { const ra = parseFloat(res.headers.get("Retry-After")); await sleep((isFinite(ra) ? ra * 1000 : 900 * (attempt + 1)) + Math.random() * 400); continue; }
      if (res.status === 401 && attempt === 0) { opts.lastStatus = 401; continue; }
      throw { status: res.status, code: res.status === 403 ? "invalid_argument" : res.status === 404 ? "not_found" : "graph_error", graphCode: err.code, message: err.message || ("Microsoft 365 said " + res.status) };
    }
  }

  /* ---------- the workbook ---------- */
  let itemId = null, webUrl = "";
  const encPath = p => p.split("/").map(encodeURIComponent).join("/");
  async function ensureFile() {
    const drive = "/drives/" + encodeURIComponent(CFG.driveId);
    try { const it = await graph("GET", drive + "/root:/" + encPath(CFG.filePath) + "?$select=id,webUrl"); itemId = it.id; webUrl = it.webUrl; return; }
    catch (e) { if (e.code !== "not_found") throw e; }
    // First run: put the empty template (with its three tables) into the team's Files.
    const tpl = await (await fetch("template.xlsx", { cache: "no-store" })).blob();
    try { const it = await graph("PUT", drive + "/root:/" + encPath(CFG.filePath) + ":/content?@microsoft.graph.conflictBehavior=fail", tpl); itemId = it.id; webUrl = it.webUrl; }
    catch (e) { if (e.status !== 409) throw e; const it = await graph("GET", drive + "/root:/" + encPath(CFG.filePath) + "?$select=id,webUrl"); itemId = it.id; webUrl = it.webUrl; }
  }
  const wb = () => "/drives/" + encodeURIComponent(CFG.driveId) + "/items/" + encodeURIComponent(itemId) + "/workbook";
  const colLetter = i => { let s = ""; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  const colIndex = L => [...L].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
  function parseAddress(a) {
    const m = String(a).match(/^(?:'?(.+?)'?!)?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/);
    if (!m) throw { code: "graph_error", message: "Unexpected table address " + a };
    return { sheet: m[1], col: colIndex(m[2]), row: parseInt(m[3], 10) };
  }

  const state = {}; // per collection: {rows:[{id, vals, sheetRow}], blank:[sheetRow], sheet, col, sig}
  async function loadTable(name) {
    const spec = SPECS[name];
    const r = await graph("GET", wb() + "/tables/" + spec.table + "/range?$select=address,values");
    const at = parseAddress(r.address), vals = r.values || [];
    const st = { rows: [], blank: [], sheet: at.sheet, col: at.col, sig: "" };
    for (let i = 1; i < vals.length; i++) {
      const row = vals[i], id = str(row[0]).trim(), sheetRow = at.row + i;
      if (!id) { if (row.every(v => v === "" || v == null)) st.blank.push(sheetRow); continue; }
      st.rows.push({ id, vals: row, sheetRow });
    }
    st.sig = JSON.stringify(vals);
    const changed = !state[name] || state[name].sig !== st.sig;
    state[name] = st;
    return changed;
  }
  const rowAddress = (st, sheetRow, n) => colLetter(st.col) + sheetRow + ":" + colLetter(st.col + n - 1) + sheetRow;
  const rangePath = (st, addr) => wb() + "/worksheets/" + encodeURIComponent(st.sheet) + "/range(address='" + addr + "')";

  /* ---------- listeners ---------- */
  const listeners = { records: [], therapists: [], settings: [] }; // settings: [{key, cb}]
  function notify(name) {
    const st = state[name]; if (!st) return;
    if (name === "settings") { listeners.settings.forEach(l => { const r = st.rows.find(x => x.id === l.key); const d = r ? SPECS.settings.fromRow(r.vals) : null; l.cb({ exists: !!r, data: () => d }); }); return; }
    const docs = st.rows.map(r => { const d = SPECS[name].fromRow(r.vals); return { id: r.id, data: () => d }; });
    listeners[name].forEach(l => l.cb({ docs }));
  }
  function fail(name, e) { (name === "settings" ? listeners.settings : listeners[name]).forEach(l => { if (l.err) try { l.err(e); } catch (_) {} }); }

  let refreshing = null, lastRefresh = 0;
  function refresh(force) {
    if (refreshing) return refreshing;
    if (!force && Date.now() - lastRefresh < 4000) return Promise.resolve();
    refreshing = (async () => {
      await Promise.all(Object.keys(SPECS).map(async n => { try { if (await loadTable(n)) notify(n); } catch (e) { fail(n, e); } }));
      lastRefresh = Date.now();
    })().finally(() => { refreshing = null; });
    return refreshing;
  }

  /* ---------- writes (one at a time) ---------- */
  let queue = Promise.resolve();
  const serial = fn => { const p = queue.then(fn, fn); queue = p.catch(() => {}); return p; };
  const newId = p => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  function write(name, id, data) {
    return serial(async () => {
      const spec = SPECS[name];
      await loadTable(name);
      const st = state[name], hit = st.rows.find(r => r.id === id);
      const values = [spec.toRow(id, data, hit && hit.vals)];
      if (hit) await graph("PATCH", rangePath(st, rowAddress(st, hit.sheetRow, spec.cols.length)), { values });
      else if (st.blank.length) await graph("PATCH", rangePath(st, rowAddress(st, st.blank[0], spec.cols.length)), { values });
      else await graph("POST", wb() + "/tables/" + spec.table + "/rows/add", { values });
      if (await loadTable(name)) notify(name);
    });
  }
  function remove(name, id) {
    return serial(async () => {
      const spec = SPECS[name];
      await loadTable(name);
      const st = state[name], hit = st.rows.find(r => r.id === id);
      if (hit) {
        const addr = rowAddress(st, hit.sheetRow, spec.cols.length);
        if (st.rows.length + st.blank.length <= 1) await graph("PATCH", rangePath(st, addr), { values: [spec.cols.map(() => "")] }); // a table always keeps one row
        else await graph("POST", rangePath(st, addr) + "/delete", { shift: "Up" });
      }
      if (await loadTable(name)) notify(name);
    });
  }

  /* ---------- the database API the app uses ---------- */
  const db = {
    collection(name) {
      if (!SPECS[name]) throw new Error("Unknown collection " + name);
      return {
        onSnapshot(cb, err) { listeners[name].push({ cb, err }); if (state[name]) setTimeout(() => notify(name), 0); },
        add(data) { return write(name, newId(name === "records" ? "r" : "t"), data); },
        doc(id) { return { set: data => write(name, id, data), delete: () => remove(name, id) }; }
      };
    },
    doc(path) {
      const key = String(path).replace(/^settings\//, "");
      return {
        onSnapshot(cb, err) { listeners.settings.push({ key, cb, err }); if (state.settings) setTimeout(() => notify("settings"), 0); },
        set: data => write("settings", key, data)
      };
    }
  };
  const downloads = {
    async save({ filename, data }) {
      const blob = data instanceof Blob ? data : new Blob([data], { type: /\.csv$/i.test(filename) ? "text/csv;charset=utf-8" : "application/octet-stream" });
      const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = filename; document.body.append(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 5000);
    }
  };
  const user = { can: async () => true, me };

  /* ---------- start-up, sign-in prompt and live refresh ---------- */
  const $ = s => document.querySelector(s);
  // #wAuth / #wSignIn sit on the welcome page; #wNote and the offline banner are shown by the app when not connected.
  function say(note, showBtn, problem) {
    const wa = $("#wAuth"), b = $("#wSignIn");
    if (wa) { wa.hidden = !note; wa.textContent = note || ""; }
    if (b) b.hidden = !showBtn;
    if (problem) { const wn = $("#wNote"); if (wn) wn.textContent = problem; const ob = $("#offlineBanner p"); if (ob) ob.textContent = problem; }
  }
  let resolveDb;
  const dbReady = new Promise(r => { resolveDb = r; });
  async function connect(interactive) {
    try {
      if (!CFG.clientId || /PASTE/.test(CFG.clientId)) { say("", false, "This app still needs its Microsoft sign-in ID. Ask Mr. Chew or IT to finish the set-up."); resolveDb(null); return; }
      if (!pca) await initAuth();
      await getToken(interactive);
      say("Opening the team’s Excel file…", false);
      await ensureFile();
      await refresh(true);
      const x = $("#xlBtn"); if (x && webUrl) { x.href = webUrl; x.hidden = false; }
      say("", false);
      resolveDb(db);
    } catch (e) {
      if (e && e.code === "redirecting") return;
      if (e && e.errorCode && interactive) { say("Sign-in didn’t finish (" + e.errorCode + "). Tap Sign in to try again.", true); return; }
      if (e && e.code === "needs_signin") { say("Sign in with your Pine Hills Microsoft account to see and save the team’s records.", true); return; }
      const msg = e && e.status === 403 ? "Your account can’t open the team’s records file. Ask Mr. Chew to add you to the Gifted Kids Department team in Microsoft Teams."
        : e && e.code === "not_found" ? "The team’s records file or its tables are missing. Ask Mr. Chew to check the Therapy Records folder in Teams."
        : (e && e.message) || "Couldn’t connect to Microsoft 365.";
      console.error("[Therapy records]", e);
      say("", false, msg);
      resolveDb(null);
    }
  }
  window.TIR_refresh = () => refresh(true);
  window.claude = { use: name => name === "db" ? dbReady : name === "downloads" ? Promise.resolve(downloads) : name === "user" ? dbReady.then(d => d ? user : null) : Promise.resolve(null) };

  document.addEventListener("DOMContentLoaded", () => {
    const b = $("#wSignIn"); if (b) b.addEventListener("click", () => { b.disabled = true; connect(true).finally(() => { b.disabled = false; }); });
    // Opening the Record System or Students tab always re-reads the Excel file.
    document.addEventListener("click", e => { if (e.target.closest && e.target.closest("#t-list,#t-stu,#refreshBtn") && state.records) refresh(true); }, true);
    document.addEventListener("visibilitychange", () => { if (!document.hidden && state.records) refresh(false); });
    setInterval(() => { if (!document.hidden && state.records) refresh(false); }, Math.max(20, CFG.pollSeconds || 60) * 1000);
    connect(false);
  });
})();
