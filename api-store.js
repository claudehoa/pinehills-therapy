/* Pine Hills Therapy Monitoring System — website storage (staff passcode version).
   Staff unlock the site with the shared staff passcode. The Pine Hills data service (a Cloudflare Worker)
   reads and writes the Therapy Records Excel file in the Gifted Kids Department team in Microsoft Teams.
   This file gives the app the same small database API it was written for (collection/doc, add/set/delete, onSnapshot). */
(function(){
  "use strict";
  const CFG = window.TIR_CONFIG || {};
  const API = String(CFG.apiUrl || "").replace(/\/+$/, "");
  const SESSION_KEY = "tir-session-v1", NAME_KEY = "tir-staffname";
  const me = { name: "" };
  try { me.name = localStorage.getItem(NAME_KEY) || ""; } catch (e) {}

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
      toRow(id, d) {
        return [id, d.date || "", txt(d.student), txt(d.klass), TYPES[d.type] || txt(d.type), txt(d.therapist),
          (typeof d.duration === "number" && isFinite(d.duration)) ? d.duration : "", txt(d.goal), txt((d.tags || []).join("; ")), txt(d.activities),
          d.progress ? PROGRESS[d.progress - 1] || "" : "", d.mood ? MOOD[d.mood - 1] || "" : "", txt(d.notes), txt(d.next),
          stamp(d.createdAt), stamp(d.updatedAt), me.name, me.name];
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

  const TABLE_OF = { records: "Records", therapists: "Team", settings: "Settings" };

  /* ---------- talking to the data service ---------- */
  let session = null;
  try { session = JSON.parse(localStorage.getItem(SESSION_KEY) || "null"); if (session && session.exp < Date.now()) session = null; } catch (e) {}
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  async function api(method, path, body) {
    for (let attempt = 0; ; attempt++) {
      let res;
      try {
        res = await fetch(API + path, { method, cache: "no-store", headers: Object.assign({ "Content-Type": "application/json" }, session ? { Authorization: "Bearer " + session.token } : {}), body: body === undefined ? undefined : JSON.stringify(body) });
      } catch (e) { if (attempt < 2 && method === "GET") { await sleep(800 * (attempt + 1)); continue; } throw { code: "network", message: "Can’t reach the school records service. Check the internet connection." }; }
      let j = {}; try { j = await res.json(); } catch (e) {}
      if (res.ok) return j;
      if (res.status === 401 && path !== "/api/session") { lock("Your passcode session has ended. Enter the staff passcode again."); }
      if (res.status >= 500 && attempt < 1 && method === "GET") { await sleep(1000); continue; }
      throw { status: res.status, code: j.error || "server_error", message: j.message || ("The records service said " + res.status) };
    }
  }

  /* ---------- tables ---------- */
  const state = {}; // name -> {rows:[{id, vals}], sig}
  function setTable(name, values) {
    const rows = [];
    (values || []).slice(1).forEach(v => { const id = str(v[0]).trim(); if (id) rows.push({ id, vals: v }); });
    const sig = JSON.stringify(values || []);
    const changed = !state[name] || state[name].sig !== sig;
    state[name] = { rows, sig };
    if (changed) notify(name);
  }
  const listeners = { records: [], therapists: [], settings: [] };
  function notify(name) {
    const st = state[name]; if (!st) return;
    if (name === "settings") { listeners.settings.forEach(l => { const r = st.rows.find(x => x.id === l.key); const d = r ? SPECS.settings.fromRow(r.vals) : null; l.cb({ exists: !!r, data: () => d }); }); return; }
    const docs = st.rows.map(r => { const d = SPECS[name].fromRow(r.vals); return { id: r.id, data: () => d }; });
    listeners[name].forEach(l => l.cb({ docs }));
  }
  function failAll(e) { Object.values(listeners).forEach(list => list.forEach(l => { if (l.err) try { l.err(e); } catch (_) {} })); }
  let webUrl = "", refreshing = null, lastRefresh = 0;
  function refresh(force) {
    if (refreshing) return refreshing;
    if (!session || (!force && Date.now() - lastRefresh < 4000)) return Promise.resolve();
    refreshing = (async () => {
      const d = await api("GET", "/api/data");
      webUrl = d.webUrl || webUrl;
      for (const [name, t] of Object.entries(TABLE_OF)) setTable(name, d.tables && d.tables[t]);
      lastRefresh = Date.now();
    })().catch(e => { if (e.status !== 401) failAll(e); throw e; }).finally(() => { refreshing = null; });
    return refreshing;
  }

  /* ---------- writes (one at a time) ---------- */
  let queue = Promise.resolve();
  const serial = fn => { const p = queue.then(fn, fn); queue = p.catch(() => {}); return p; };
  const newId = p => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  const mapErr = e => (e && e.status === 403) ? Object.assign({}, e, { code: "invalid_argument" }) : e;
  function write(name, id, data) {
    return serial(async () => {
      try {
        const values = SPECS[name].toRow(id, data);
        const r = await api("POST", "/api/upsert", { table: TABLE_OF[name], id, values, keep: name === "records" ? [16] : [] });
        setTable(name, r.values);
      } catch (e) { throw mapErr(e); }
    });
  }
  function remove(name, id) {
    return serial(async () => {
      try { const r = await api("POST", "/api/delete", { table: TABLE_OF[name], id }); setTable(name, r.values); }
      catch (e) { throw mapErr(e); }
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

  /* ---------- passcode screen, start-up and live refresh ---------- */
  const $ = s => document.querySelector(s);
  function say(note, showForm, problem) {
    const wa = $("#wAuth"), f = $("#wPass");
    if (wa) { wa.hidden = !note; wa.textContent = note || ""; }
    if (f) f.hidden = !showForm;
    const en = $("#wEnter"); if (en) en.hidden = !!showForm;
    if (problem) { const wn = $("#wNote"); if (wn) wn.textContent = problem; const ob = $("#offlineBanner p"); if (ob) ob.textContent = problem; }
  }
  function lock(msg) {
    session = null; try { localStorage.removeItem(SESSION_KEY); } catch (e) {}
    const w = $("#welcome"); if (w && w.hidden) w.hidden = false;
    say(msg, true);
  }
  let resolveDb, resolved = false;
  const dbReady = new Promise(r => { resolveDb = v => { if (!resolved) { resolved = true; r(v); } }; });
  async function connect() {
    try {
      if (!API) { say("", false, "This website isn’t linked to the school records service yet. Ask Mr. Chew to finish the set-up."); resolveDb(null); return; }
      if (!session) { say("Enter your name and the staff passcode to see and save the team’s records.", true); return; }
      say("Opening the team’s records…", false);
      await refresh(true);
      const x = $("#xlBtn"); if (x && webUrl) { x.href = webUrl; x.hidden = false; }
      say("", false);
      resolveDb(db);
    } catch (e) {
      if (e && e.status === 401) return;
      console.error("[Therapy records]", e);
      const msg = (e && e.message) || "Couldn’t connect to the school records.";
      if (resolved) { say(msg, false); return; }
      say("", false, msg); resolveDb(null);
    }
  }
  async function unlock(ev) {
    ev.preventDefault();
    const name = ($("#wName").value || "").trim(), code = ($("#wCode").value || "").trim(), btn = $("#wUnlock");
    if (!name) { $("#wName").focus(); say("Please type your name, so records show who entered them.", true); return; }
    if (!code) { $("#wCode").focus(); return; }
    btn.disabled = true;
    try {
      const r = await api("POST", "/api/session", { passcode: code, name });
      session = { token: r.token, exp: r.exp }; me.name = name;
      try { localStorage.setItem(SESSION_KEY, JSON.stringify(session)); localStorage.setItem(NAME_KEY, name); if (!localStorage.getItem("tir-therapist")) localStorage.setItem("tir-therapist", JSON.stringify(name)); } catch (e) {}
      $("#wCode").value = "";
      await connect();
      if (resolved) { try { await refresh(true); } catch (e) {} }
    } catch (e) {
      say((e && e.message) || "Couldn’t check the passcode.", true);
      $("#wCode").select();
    } finally { btn.disabled = false; }
  }
  window.TIR_refresh = () => refresh(true);
  window.TIR_signOut = () => { session = null; try { localStorage.removeItem(SESSION_KEY); } catch (e) {} location.reload(); };
  window.claude = { use: name => name === "db" ? dbReady : name === "downloads" ? Promise.resolve(downloads) : name === "user" ? dbReady.then(d => d ? user : null) : Promise.resolve(null) };

  document.addEventListener("DOMContentLoaded", () => {
    const f = $("#wPass"); if (f) f.addEventListener("submit", unlock);
    if ($("#wName") && me.name) $("#wName").value = me.name;
    // Opening the Record System or Students tab always re-reads the Excel file.
    document.addEventListener("click", e => { if (e.target.closest && e.target.closest("#t-list,#t-stu,#refreshBtn") && state.records) refresh(true).catch(() => {}); }, true);
    document.addEventListener("visibilitychange", () => { if (!document.hidden && state.records) refresh(false).catch(() => {}); });
    setInterval(() => { if (!document.hidden && state.records) refresh(false).catch(() => {}); }, Math.max(20, CFG.pollSeconds || 60) * 1000);
    connect();
  });
})();
