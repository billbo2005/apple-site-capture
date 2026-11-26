/* Project Apple · Site Capture – works-to-go capture app (offline-first PWA, Microsoft 365 / SharePoint back end)
   Project Central Assurance, Nov 2026.  v1.1: progress is logged per activity per room on a slider (Progress %, 0 to 100). */
(() => {
"use strict";
const C = window.APP_CONFIG;
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["Sites.ReadWrite.All", "User.Read"];
const STATUSES = ["Open", "In Progress", "Closed", "On Hold", "At Risk", "Not applicable"];
const FIELDS = ["Title","ItemType","Descr","Building","Floor","Room","Zone","Package","SystemName","WorkOwner","UCBResp","UCBResp2","Cat",
  "DateRaised","TargetDate","ItemStatus","ClosureDate","Notes","SourceRef","LocationID","RowKey","LastBy","LastAt","PhotoIds","Activity","Progress"];
/* The activity list (Parameters tab of the tracker, Seamus's list of 15 Nov 2026): order on screen and discipline. */
const ACT = [["Lighting 1st Fix","Electrical"],["Lighting 2nd Fix","Electrical"],["Power 1st Fix","Electrical"],["Power 2nd Fix","Electrical"],["Fire Alarm 1st Fix","Electrical"],["Fire Alarm 2nd Fix","Electrical"],["Data 1st Fix","Electrical"],["Data 2nd Fix","Electrical"],["Earthing 1st Fix","Electrical"],["Earthing 2nd Fix","Electrical"],["Door Access Control 1st Fix","Electrical"],["Door Access Control 2nd Fix","Electrical"],["Sprinkler 1st Fix","Mechanical"],["Sprinkler 2nd Fix","Mechanical"],["LTHW 1st Fix","Mechanical"],["LTHW 2nd Fix","Mechanical"],["CHW 1st Fix","Mechanical"],["CHW 2nd Fix","Mechanical"],["Lab Gases 1st Fix","Mechanical"],["Lab Gases 2nd Fix","Mechanical"],["Compressed Air 1st Fix","Mechanical"],["Compressed Air 2nd Fix","Mechanical"],["Ventilation 1st Fix","Mechanical"],["Ventilation 2nd Fix","Mechanical"],["Sprinkler Mist 1st Fix","Mechanical"],["Sprinkler Mist 2nd Fix","Mechanical"],["Drainage 1st Fix","Mechanical"],["Drainage 2nd Fix","Mechanical"],["BMS 1st Fix","BMS"],["BMS 2nd Fix","BMS"],["Walls 1st Fix","CSA"],["Walls 2nd Fix","CSA"],["Ceiling 1st Fix","CSA"],["Ceiling 2nd Fix","CSA"],["Flooring / Carpet / Vinyl","CSA"],["Door Install 1st Fix","CSA"],["Door Install 2nd Fix","CSA"],["Painting","CSA"],["Skirting Board","CSA"],["Architrave","CSA"],["BASWA 1st Fix","CSA"],["BASWA 2nd Fix","CSA"],["RAF Install","CSA"],["IPS 1st Fix","Furniture"],["IPS 2nd Fix","Furniture"],["Lab Furniture 1st Fix","Furniture"],["Lab Furniture 2nd Fix","Furniture"],["Office Furniture 1st Fix","Furniture"],["Office Furniture 2nd Fix","Furniture"]];
const ACT_IX = new Map(ACT.map((a, k) => [a[0], { k, d: a[1] }]));
const DISCS = ["Electrical", "Mechanical", "BMS", "CSA", "Furniture", "Other"];
const HOLDS = ["On Hold", "At Risk", "Not applicable"];
const ROOMF = ["Title","Building","Floor","Room","Zone","WorkOwner","UCBResp","UCBResp2"];
const $ = (s, el = document) => el.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const stc = s => "st-" + String(s || "Open").replace(/[^A-Za-z]/g, "");
const MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const fmtD = iso => { if (!iso) return ""; const d = new Date(iso); return isNaN(d) ? "" : `${d.getUTCDate()} ${MON[d.getUTCMonth()]} ${d.getUTCFullYear()}`; };
const todayISO = () => { const n = new Date(); return new Date(Date.UTC(n.getFullYear(), n.getMonth(), n.getDate())).toISOString(); };
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));
const toast = (m, ms = 2600) => { const t = $("#toast"); t.textContent = m; t.classList.add("show"); clearTimeout(t._h); t._h = setTimeout(() => t.classList.remove("show"), ms); };

/* ---------------- IndexedDB ---------------- */
let dbp;
function db() {
  if (!dbp) dbp = new Promise((res, rej) => {
    const r = indexedDB.open("apple-capture", 1);
    r.onupgradeneeded = () => { const d = r.result;
      d.createObjectStore("items", { keyPath: "RowKey" }); d.createObjectStore("rooms", { keyPath: "Title" });
      d.createObjectStore("outbox", { keyPath: "qid", autoIncrement: true }); d.createObjectStore("meta"); };
    r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
  });
  return dbp;
}
async function tx(store, mode, fn) { const d = await db(); return new Promise((res, rej) => {
  const t = d.transaction(store, mode); const s = t.objectStore(store); let out; Promise.resolve(fn(s)).then(v => out = v);
  t.oncomplete = () => res(out); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error); }); }
const req = r => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
const getAll = store => tx(store, "readonly", s => req(s.getAll()));
const putMany = (store, arr) => tx(store, "readwrite", s => { arr.forEach(x => s.put(x)); });
const putOne = (store, v, k) => tx(store, "readwrite", s => req(k === undefined ? s.put(v) : s.put(v, k)));
const delOne = (store, k) => tx(store, "readwrite", s => req(s.delete(k)));
const delMany = (store, keys) => tx(store, "readwrite", s => { keys.forEach(k => s.delete(k)); });
const meta = { get: k => tx("meta", "readonly", s => req(s.get(k))), set: (k, v) => tx("meta", "readwrite", s => req(s.put(v, k))) };

/* ---------------- state ---------------- */
let ITEMV = 0;                                             // bumped whenever an item changes, so the by-room index is rebuilt only when needed
class VMap extends Map { set(k, v) { ITEMV++; return super.set(k, v); } delete(k) { ITEMV++; return super.delete(k); } }
const S = { items: new VMap(), rooms: new Map(), outbox: [], user: null, syncing: false, lastSync: null, filter: { b: null, f: null, q: "" }, roomDisc: {} };
async function loadLocal() {
  (await getAll("items")).forEach(i => S.items.set(i.RowKey, i));
  (await getAll("rooms")).forEach(r => S.rooms.set(r.Title, r));
  S.outbox = await getAll("outbox");
  S.lastSync = await meta.get("lastSync");
  const f = await meta.get("filter"); if (f) S.filter = { ...S.filter, ...f, q: "" };
  const w = await meta.get("who"); if (w) { S.who = w.name; S.whoDay = w.day; }
  S.whoNames = (await meta.get("whoNames")) || [];
}
/* ---- who is using a shared tablet today ---- */
const localDay = () => { const n = new Date(); return `${n.getFullYear()}-${String(n.getMonth() + 1).padStart(2, "0")}-${String(n.getDate()).padStart(2, "0")}`; };
const isShared = () => !!S.user && (C.sharedSignIns || []).some(x => x.toLowerCase() === String(S.user.email || "").toLowerCase());
const whoNow = () => (isShared() && S.who && S.whoDay === localDay()) ? S.who : "";
const needWho = () => isShared() && !whoNow();
const byLine = () => !S.user ? "" : (whoNow() ? `${whoNow()} (${S.user.email})` : S.user.email);   // goes into 'Last changed by' on every entry
async function setWho(name) {
  S.who = name; S.whoDay = localDay();
  S.whoNames = [name, ...(S.whoNames || []).filter(n => n.toLowerCase() !== name.toLowerCase())].slice(0, 6);
  await meta.set("who", { name, day: S.whoDay }); await meta.set("whoNames", S.whoNames); paintWho();
}
function paintWho() { const el = $(".brand .t2"); if (el) el.textContent = "UCB · works to go" + (whoNow() ? " · " + whoNow() : ""); }

/* ---------------- sign-in ---------------- */
let pca, account;
async function initAuth() {
  pca = new msal.PublicClientApplication({
    auth: { clientId: C.clientId, authority: `https://login.microsoftonline.com/${C.tenantId}`,
            redirectUri: location.origin + location.pathname.replace(/index\.html$/, ""), navigateToLoginRequestUrl: false },
    cache: { cacheLocation: "localStorage" } });
  await pca.initialize();
  try { const r = await pca.handleRedirectPromise(); if (r && r.account) account = r.account; } catch (e) { console.warn(e); }
  account = account || pca.getAllAccounts()[0] || null;
  if (account) { pca.setActiveAccount(account); S.user = { name: account.name || account.username, email: account.username }; }
}
const signIn = () => pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
const signOut = () => pca.logoutRedirect({ account });
async function token() {
  try { return (await pca.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
  catch (e) { if (e instanceof msal.InteractionRequiredAuthError && navigator.onLine) { await pca.acquireTokenRedirect({ scopes: SCOPES, account }); } throw e; }
}
async function g(path, opt = {}) {
  const t = await token();
  const ac = new AbortController(); const timer = setTimeout(() => ac.abort(), opt.method === "PUT" ? 120000 : 90000);   // a stalled request gives up instead of hanging for ever
  let r;
  try { r = await fetch(path.startsWith("http") ? path : GRAPH + path, { ...opt, signal: ac.signal, headers: { Authorization: "Bearer " + t, ...(opt.headers || {}) } }); }
  catch (e) { throw new Error(e.name === "AbortError" ? "No answer from SharePoint (timed out)" : e.message); }
  finally { clearTimeout(timer); }
  if (r.status === 429 || r.status === 503) { await new Promise(z => setTimeout(z, 3000)); return g(path, opt); }
  if (!r.ok) { const txt = await r.text(); throw new Error(`Graph ${r.status}: ${txt.slice(0, 200)}`); }
  return r.status === 204 ? null : r.json();
}

/* ---------------- SharePoint via Graph ---------------- */
let site;
async function siteIds() {
  if (site) return site;
  site = await meta.get("site");
  if (site) return site;
  const s = await g(`/sites/${C.siteHost}:${C.sitePath}?$select=id`);
  const lists = await g(`/sites/${s.id}/lists?$select=id,name,displayName`);
  const by = n => (lists.value.find(l => l.displayName === n || l.name === n) || {}).id;
  const drv = await g(`/sites/${s.id}/lists/${by(C.photosList)}/drive?$select=id`);
  site = { id: s.id, items: by(C.itemsList), rooms: by(C.roomsList), drive: drv.id };
  await meta.set("site", site); return site;
}
function fromSP(it, keys) { const f = it.fields || {}; const o = { spId: it.id }; keys.forEach(k => { if (f[k] !== undefined && f[k] !== null) o[k] = f[k]; }); return o; }
async function pullAll(listId, keys, onRow) {
  let url = `/sites/${site.id}/lists/${listId}/items?$expand=fields($select=${keys.join(",")})&$top=999`;
  while (url) { const p = await g(url); p.value.forEach(it => onRow(fromSP(it, keys))); url = p["@odata.nextLink"]; }
}
async function pullDelta(listId, keys, metaKey, onRow, onDel) {
  let url = await meta.get(metaKey);
  const full = !url;
  if (!url) url = `/sites/${site.id}/lists/${listId}/items/delta?$expand=fields($select=${keys.join(",")})&$top=999`;
  try {
    while (url) { const p = await g(url);
      p.value.forEach(it => it.deleted ? onDel(it.id) : (it.fields && onRow(fromSP(it, keys))));
      if (p["@odata.deltaLink"]) { await meta.set(metaKey, p["@odata.deltaLink"]); url = null; } else url = p["@odata.nextLink"]; }
    return full;
  } catch (e) { await meta.set(metaKey, null); throw e; }
}
async function refresh() {
  await siteIds();
  { const rooms = []; await pullAll(site.rooms, ROOMF, r => rooms.push(r)); await putMany("rooms", rooms); rooms.forEach(r => S.rooms.set(r.Title, r)); }   // every time: new rooms reach phones that already have a room list
  const changed = [];
  const pend = new Set(S.outbox.map(o => o.rowKey));
  const byId = new Map([...S.items.values()].map(i => [i.spId, i.RowKey]));
  const seen = new Set(), gone = [];
  const onRow = r => { seen.add(r.spId); if (!r.RowKey) return; if (pend.has(r.RowKey)) return; S.items.set(r.RowKey, r); changed.push(r); if (changed.length % 200 === 0) { S.phase = `Updating list… ${changed.length.toLocaleString("en-GB")}`; paintChip(); } };
  const onDel = id => { const k = byId.get(id); if (k) gone.push(k); };
  let full;
  try { full = await pullDelta(site.items, FIELDS, "deltaItems", onRow, onDel); }
  catch (e) { console.warn("delta failed, full pull", e); await pullAll(site.items, FIELDS, onRow); full = true; }
  await putMany("items", changed);
  if (full) {                                               // after a whole-list read: rows this phone holds that are no longer on the list
    const local = [...S.items.values()].filter(i => i.spId), miss = local.filter(i => !seen.has(i.spId));
    if (seen.size && miss.length <= local.length / 2) { miss.forEach(i => gone.push(i.RowKey)); await meta.set("recon", "1"); }   // more than half missing = the read is not to be trusted: remove nothing
    else console.warn("list read looks wrong – nothing removed", seen.size, local.length);
  }
  if (gone.length) { gone.forEach(k => S.items.delete(k)); await delMany("items", gone); }
  await prune();
}
// a change waiting for a row that has been taken off the list can never be sent: it is set aside, on record, instead of waiting for ever
async function prune() {
  const creating = new Set(S.outbox.filter(o => o.type === "create").map(o => o.rowKey));
  const dead = S.outbox.filter(o => o.type !== "create" && !creating.has(o.rowKey) && !(S.items.get(o.rowKey) || {}).spId);
  if (!dead.length) return;
  const log = (await meta.get("unsent")) || [];
  for (const o of dead) { log.unshift({ at: o.at, what: o.label, detail: o.detail || "", room: o.room || "", photos: o.photos || [], why: "This row is no longer on the list" }); await delOne("outbox", o.qid); }
  S.outbox = S.outbox.filter(o => !dead.includes(o));
  await meta.set("unsent", log.slice(0, 300));
}
const why = e => { const m = String((e && e.message) || e);
  return /Graph 403/.test(m) ? "This sign-in is not allowed to change the list – tell the office"
    : /Graph 404|not on server/.test(m) ? "Row not found on the list"
    : /Graph 401/.test(m) ? "The sign-in has run out – sign in again"
    : /timed out|Failed to fetch|Load failed|NetworkError/i.test(m) ? "No answer from SharePoint – will try again"
    : "Could not send (" + m.slice(0, 60) + ")"; };

/* ---------------- outbox ---------------- */
async function queue(op) { op.at = new Date().toISOString(); op.by = byLine(); const qid = await putOne("outbox", op); op.qid = qid; S.outbox.push(op); paintChip(); }
function nextId() { let m = (C.newIdsFrom || 1) - 1;       // items raised in the app are numbered clear of the tracker's own numbers
  for (const i of S.items.values()) { const n = /^WTG-(\d+)$/.exec(i.Title || ""); if (n) m = Math.max(m, +n[1]); } return "WTG-" + String(m + 1).padStart(4, "0"); }
async function uploadPhotos(rowKey, photos) {
  const out = [];
  for (let n = 0; n < photos.length; n++) {
    const p = photos[n]; const stamp = p.at.replace(/[-:T]/g, "").slice(0, 14);
    const name = `${encodeURIComponent(rowKey)}/${stamp}-${n + 1}.jpg`;
    const r = await g(`/drives/${site.drive}/root:/${name}:/content`, { method: "PUT", headers: { "Content-Type": "image/jpeg" }, body: p.blob });
    out.push(r.webUrl);
  }
  return out;
}
async function sync(manual) {
  if (S.syncing) { S.again = true; return; }
  if (!navigator.onLine || !account) { if (manual && !navigator.onLine) toast("No signal – changes stay on this phone and send later."); return; }
  S.again = false;
  S.syncing = true; paintChip();
  let sent = 0, failed = 0, lost = false;
  try {
    await siteIds();
    const send = async ops => { let k = 0; for (const op of ops) {
      S.phase = `Sending… ${++k} of ${ops.length}`; paintChip();
      try {
        const item = S.items.get(op.rowKey) || {};
        const urls = op.photos && op.photos.length ? await uploadPhotos(item.Title && !item.Title.startsWith("NEW") ? item.Title : op.rowKey, op.photos) : [];
        const fields = { ...op.fields, LastBy: op.by, LastAt: op.at };
        if (urls.length) fields.PhotoIds = [item.PhotoIds, ...urls].filter(Boolean).join("\n");
        if (op.type === "create") {
          fields.Title = nextId(); fields.RowKey = op.rowKey;
          const r = await g(`/sites/${site.id}/lists/${site.items}/items`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ fields }) });
          const local = { ...item, ...fields, spId: r.id }; S.items.set(op.rowKey, local); await putOne("items", local);
        } else {
          if (!item.spId) throw new Error("item not on server yet");
          await g(`/sites/${site.id}/lists/${site.items}/items/${item.spId}/fields`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify(fields) });
          const local = { ...item, ...fields }; S.items.set(op.rowKey, local); await putOne("items", local);
        }
        await delOne("outbox", op.qid); S.outbox = S.outbox.filter(o => o.qid !== op.qid);
        const sentLog = (await meta.get("sent")) || []; sentLog.unshift({ at: new Date().toISOString(), what: op.label, room: op.detail || op.room });
        await meta.set("sent", sentLog.slice(0, 30)); sent++;
      } catch (e) { console.error(e); failed++; op.err = why(e); if (op.type !== "create" && /Graph 404/.test(String(e.message))) lost = true; }
    } };
    S.phase = "Sending…"; paintChip();
    await send(S.outbox.filter(o => o.type !== "create"));  // changes to existing items go first: they do not need the list refreshed
    if (lost && !S.reconTried) { S.reconTried = true; await meta.set("deltaItems", null); }   // a row was not found: read the whole list once and drop what is no longer on it
    S.phase = "Updating list…"; paintChip();
    const before = S.outbox.length;
    await refresh();                                        // fresh numbers before creating new items
    const setAside = before - S.outbox.length; failed = Math.max(0, failed - setAside);
    const rest = [...S.outbox].filter(o => o.type === "create");
    if (rest.length) { S.phase = "Sending…"; paintChip(); await send(rest); }
    S.lastSync = new Date().toISOString(); await meta.set("lastSync", S.lastSync);
    const barred = S.outbox.some(o => /not allowed/.test(o.err || ""));
    if (barred) toast("This sign-in is not allowed to change the list – tell the office", 7000);
    else if (setAside) toast(`${setAside} change${setAside > 1 ? "s" : ""} set aside: the row${setAside > 1 ? "s are" : " is"} no longer on the list – see Outbox`, 7000);
    else if (manual || sent) toast(failed ? `${sent} sent, ${failed} will retry` : sent ? `${sent} change${sent > 1 ? "s" : ""} sent` : "Up to date");
  } catch (e) { console.error(e); if (manual) toast("Couldn't reach SharePoint – will try again"); }
  // the screen is redrawn after a sync, except one somebody is filling in (an activity or a new item)
  finally { S.syncing = false; paintChip(); if (!/^#\/(item|add)\//.test(location.hash) && (!document.activeElement || !/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName))) route(); if (S.again || (sent && S.outbox.length)) setTimeout(() => sync(), 300); }
}

/* ---------------- photos ---------------- */
function shrink(file) { return new Promise((res, rej) => {
  const img = new Image(); const u = URL.createObjectURL(file);
  img.onload = () => { const m = 1600, s = Math.min(1, m / Math.max(img.width, img.height));
    const c = document.createElement("canvas"); c.width = Math.round(img.width * s); c.height = Math.round(img.height * s);
    c.getContext("2d").drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(u);
    c.toBlob(b => b ? res(b) : rej(new Error("photo")), "image/jpeg", 0.72); };
  img.onerror = rej; img.src = u; }); }

/* ---------------- helpers ---------------- */
const BAND = p => p >= 100 ? "var(--b100)" : p >= 75 ? "var(--b75)" : p >= 50 ? "var(--b50)" : p >= 25 ? "var(--b25)" : "var(--b0)";
const SHORT = b => String(b || "").replace("Building ", "").replace(" Annex", " Annexe");
const FLOOR_ORDER = ["LGF", "GF", "UGF", "Mezz", "Plant / Switch Room", "Interstitial", "Roof", "All"];
const counted = i => i.ItemStatus !== "Not applicable";
const discOf = i => (ACT_IX.get(i.Activity) || {}).d || (DISCS.includes(i.Package) ? i.Package : "Other");   // items with no listed activity sit under their package
const ordOf = i => (ACT_IX.get(i.Activity) || { k: 999 }).k;
const nameOf = i => i.Activity || i.Descr || "(no description)";
const pctOf = i => typeof i.Progress === "number" ? Math.max(0, Math.min(100, i.Progress)) : (i.ItemStatus === "Closed" ? 100 : 0);   // Closed counts as 100%, as on the tracker
const pctTxt = p => (Math.round(p * 10) / 10) + "%";
let LOCV = -1, LOC = new Map();
function byLoc(loc) {
  if (LOCV !== ITEMV) { LOC = new Map(); for (const i of S.items.values()) { const a = LOC.get(i.LocationID); if (a) a.push(i); else LOC.set(i.LocationID, [i]); } LOCV = ITEMV; }
  return LOC.get(loc) || [];
}
function roomStats(loc) { let t = 0, c = 0, o = 0, sum = 0; for (const i of byLoc(loc)) if (counted(i)) { t++; sum += pctOf(i); if (i.ItemStatus === "Closed") c++; else o++; } return { t, c, o, p: t ? Math.round(sum / t) : 0, sum }; }
function pending(rowKey) { return S.outbox.some(o => o.rowKey === rowKey); }
const chev = '<svg class="chev" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg>';
const backI = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" aria-hidden="true"><path d="M15 6l-6 6 6 6"/></svg>';
const camI = '<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round" aria-hidden="true"><path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="4"/></svg>';

function paintChip() {
  const ch = $("#netchip"); const n = S.outbox.length;
  if (S.syncing) { ch.className = "chip busy"; ch.textContent = S.phase || "Sending…"; }
  else if (!navigator.onLine) { ch.className = "chip off"; ch.textContent = `No signal · ${n} waiting`; }
  else { ch.className = n ? "chip off" : "chip on"; ch.textContent = n ? `${n} waiting` : "All sent"; }
  $("#navout").textContent = n ? `Outbox · ${n}` : "Outbox";
}
function setTab(t) { document.querySelectorAll("#nav a").forEach(a => a.classList.toggle("act", a.dataset.tab === t)); }

/* ---------------- views ---------------- */
function vSignIn() {
  setTab("");
  $("#view").innerHTML = `<div class="signin"><h1>Sign in</h1>
    <p>Use your Project Central account, or the email address you were invited with.<br>You only need to do this once on this phone.</p>
    <button class="big navy" id="si" type="button">Sign in with Microsoft</button>
    <p class="sub" style="margin-top:24px">Tip: on iPhone, tap Share → <b>Add to Home Screen</b> so the app works with no signal.</p></div>`;
  $("#si").onclick = () => navigator.onLine ? signIn() : toast("Sign-in needs signal the first time.");
}
function vWho() {
  if ($("#whoform")) return;                                // already on screen: do not wipe a name being typed
  setTab("");
  const names = S.whoNames || [];
  $("#view").innerHTML = `<div id="whoform"><h1>Who is using this tablet today?</h1>
    <p class="sub">Your name goes on each entry you log today, so the office knows who to ask about it.</p>
    ${names.length ? `<div class="lab">Tap your name</div><div class="grid g2" id="wnm">${names.map(n => `<button class="big line" type="button" data-n="${esc(n)}">${esc(n)}</button>`).join("")}</div><div class="lab">Or someone new</div>` : ""}
    <label class="field"><span class="lab" style="margin:0">First name and surname</span><input id="wn" autocomplete="off" autocapitalize="words" maxlength="40" placeholder="e.g. Sam Jones"></label>
    <div style="margin-top:14px"><button class="big navy" id="wgo" type="button">Start</button></div></div>`;
  const go = async name => { name = String(name || "").replace(/\s+/g, " ").trim();
    if (name.length < 2) { toast("Type your name"); $("#wn").focus(); return; }
    await setWho(name); $("#whoform").remove(); toast(`Logging as ${name}`); route(); };
  document.querySelectorAll("#wnm [data-n]").forEach(b => b.onclick = () => go(b.dataset.n));
  $("#wgo").onclick = () => go($("#wn").value);
  $("#wn").onkeydown = e => { if (e.key === "Enter") go($("#wn").value); };
}
function vLoading(msg) { $("#view").innerHTML = `<div class="empty">${esc(msg)}</div>`; }

function vRooms() {
  setTab("rooms");
  const rooms = [...S.rooms.values()];
  const blds = [...new Set(rooms.map(r => r.Building))].sort();
  if (!S.filter.b || !blds.includes(S.filter.b)) S.filter.b = blds.includes("Building B") ? "Building B" : blds[0];
  const floors = [...new Set(rooms.filter(r => r.Building === S.filter.b).map(r => r.Floor))].sort((a, b) => FLOOR_ORDER.indexOf(a) - FLOOR_ORDER.indexOf(b));
  if (!S.filter.f || !floors.includes(S.filter.f)) S.filter.f = floors[0];
  const q = S.filter.q.trim().toUpperCase().replace(/\./g, "");
  let list = q ? rooms.filter(r => String(r.Room || "").toUpperCase().replace(/\./g, "").includes(q) || String(r.Title).toUpperCase().includes(q))
               : rooms.filter(r => r.Building === S.filter.b && r.Floor === S.filter.f);
  const zones = new Map(); list.forEach(r => { const z = q ? `${r.Building} · ${r.Floor}` : (r.Zone ? "Zone " + r.Zone : "No zone"); if (!zones.has(z)) zones.set(z, []); zones.get(z).push(r); });
  let html = `<div class="lab">Building</div><div class="grid g4">${blds.map(b => `<button type="button" class="pick ${b === S.filter.b && !q ? "sel" : ""}" data-b="${esc(b)}">${esc(SHORT(b))}</button>`).join("")}</div>
    <div class="lab">Floor</div><div class="grid g4">${floors.map(f => `<button type="button" class="pick ${f === S.filter.f && !q ? "sel" : ""}" data-f="${esc(f)}">${esc(f)}</button>`).join("")}</div>
    <label class="search"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/></svg>
    <input id="q" type="search" inputmode="search" placeholder="Type a room number, e.g. 053" aria-label="Search rooms" value="${esc(S.filter.q)}"></label>`;
  if (!list.length) html += `<div class="empty">No rooms match.</div>`;
  for (const [z, rs] of [...zones.entries()].sort()) {
    rs.sort((a, b) => String(a.Room).localeCompare(String(b.Room)));
    let zt = 0, zs = 0; rs.forEach(r => { const s = roomStats(r.Title); zt += s.t; zs += s.sum; });
    html += `<div class="zonehd"><b>${esc(z)}</b><span class="sub">${rs.length} rooms${zt ? ` · ${Math.round(zs / zt)}% done` : ""}</span></div><div class="list">`;
    for (const r of rs) { const s = roomStats(r.Title); const p = s.p;
      html += `<a class="row" href="#/room/${encodeURIComponent(r.Title)}"><span class="grow"><span class="nm">${esc(r.Room)}</span>
        <span class="bar" aria-hidden="true"><span style="width:${Math.max(p, 3)}%;background:${BAND(p)}"></span></span></span>
        <span class="cnt"><b class="pct">${p}%</b><br><span class="s">${s.c} of ${s.t} closed</span></span>${chev}</a>`; }
    html += `</div>`;
  }
  html += `<div class="legend"><span>Bar = average progress:</span><span><i style="background:var(--b0)"></i>0–24%</span><span><i style="background:var(--b25)"></i>25–49%</span><span><i style="background:var(--b50)"></i>50–74%</span><span><i style="background:var(--b75)"></i>75–99%</span><span><i style="background:var(--b100)"></i>100%</span></div>`;
  $("#view").innerHTML = html;
  $("#view").querySelectorAll("[data-b]").forEach(b => b.onclick = () => { S.filter = { b: b.dataset.b, f: null, q: "" }; meta.set("filter", S.filter); vRooms(); });
  $("#view").querySelectorAll("[data-f]").forEach(b => b.onclick = () => { S.filter.f = b.dataset.f; S.filter.q = ""; meta.set("filter", S.filter); vRooms(); });
  const qi = $("#q"); qi.oninput = () => { S.filter.q = qi.value; clearTimeout(qi._t); qi._t = setTimeout(() => { vRooms(); const n = $("#q"); n.focus(); n.setSelectionRange(n.value.length, n.value.length); }, 250); };
}

function vRoom(loc) {
  setTab("rooms");
  const r = S.rooms.get(loc); if (!r) return vRooms();
  const items = byLoc(loc).slice().sort((a, b) => ordOf(a) - ordOf(b) || String(a.Title).localeCompare(String(b.Title), "en", { numeric: true }));
  const s = roomStats(loc);
  const discs = DISCS.filter(d => items.some(i => discOf(i) === d));
  let disc = S.roomDisc[loc]; if (!discs.includes(disc)) disc = discs[0];
  const shown = items.filter(i => discOf(i) === disc);
  $("#view").innerHTML = `<a class="back" href="#/">${backI} ${esc(SHORT(r.Building))} · ${esc(r.Floor)} rooms</a>
    <div style="display:flex;justify-content:space-between;align-items:flex-end;gap:12px;flex-wrap:wrap">
      <div><h1>${esc(r.Room)}</h1><div class="sub">${r.Zone ? "Zone " + esc(r.Zone) + " · " : ""}work owner ${esc(r.WorkOwner || "–")}</div></div>
      <div style="text-align:right"><div class="pct" style="font-size:30px">${s.p}%</div><div class="sub">room progress</div></div></div>
    <div class="bar" style="margin-top:10px" aria-hidden="true"><span style="width:${s.p}%;background:var(--navy)"></span></div>
    <div class="sub" style="font-size:13px;margin-top:6px">${s.c} of ${s.t} activities closed</div>
    <div class="pills">${discs.map(d => `<button type="button" class="pill ${d === disc ? "sel" : ""}" data-p="${esc(d)}">${esc(d)} · ${items.filter(i => discOf(i) === d).length}</button>`).join("")}</div>
    <div class="sub" style="font-size:13px;margin-bottom:8px">Tap an activity to log its progress.</div>
    <div class="list">${shown.map(i => { const live = counted(i), p = pctOf(i); return `<a class="row" href="#/item/${encodeURIComponent(i.RowKey)}">
      <span class="grow"><span class="nm">${esc(nameOf(i))}</span>
        <span class="bar thin" aria-hidden="true"><span style="width:${live ? p : 0}%;background:var(--navy)"></span></span>
        ${pending(i.RowKey) ? '<span class="pend">waiting to send</span>' : ""}</span>
      <span class="rowr"><b class="pct">${live ? pctTxt(p) : "–"}</b><span class="tag sm ${stc(i.ItemStatus)}">${esc(i.ItemStatus || "Open")}</span></span>${chev}</a>`; }).join("") || `<div class="empty">No activities.</div>`}</div>
    <div style="margin-top:14px"><a class="big line" href="#/add/${encodeURIComponent(loc)}">＋ Add another works to go item</a></div>`;
  $("#view").querySelectorAll("[data-p]").forEach(b => b.onclick = () => { S.roomDisc[loc] = b.dataset.p; vRoom(loc); });
}

function photoPicker(holder, photos, existing) {
  const paint = () => {
    holder.innerHTML = existing.map(u => `<a class="ph" href="${esc(u)}" target="_blank" rel="noopener" style="background-color:#3A4756;align-items:center;justify-content:center;font-size:13px">Saved photo</a>`).join("") +
      photos.map((p, n) => `<div class="ph" style="background-image:url(${p.url})"><button type="button" data-x="${n}" aria-label="Remove photo">✕</button></div>`).join("") +
      `<label class="addph">${camI}<span>Take photo</span><input type="file" accept="image/*" capture="environment" hidden></label>`;
    holder.querySelectorAll("[data-x]").forEach(b => b.onclick = () => { photos.splice(+b.dataset.x, 1); paint(); });
    holder.querySelector("input").onchange = async e => { const f = e.target.files[0]; if (!f) return;
      try { const blob = await shrink(f); photos.push({ blob, url: URL.createObjectURL(blob), at: new Date().toISOString() }); paint(); } catch { toast("Couldn't read that photo"); } };
  };
  paint();
}

function vItem(rowKey) {
  setTab("rooms");
  const i = S.items.get(rowKey); if (!i) return vRooms();
  const was = pctOf(i), st0 = i.ItemStatus || "Open";
  let pct = was, flag = HOLDS.includes(st0) ? st0 : "";
  const photos = []; const existing = String(i.PhotoIds || "").split("\n").filter(Boolean);
  const oldName = i.Activity && i.Descr && i.Descr !== i.Activity ? ` · on the old list as "${esc(i.Descr)}"` : "";
  const statusNow = () => flag ? flag : pct >= 100 ? "Closed" : pct > 0 ? "In Progress" : (st0 === "In Progress" && pct === was ? "In Progress" : "Open");
  $("#view").innerHTML = `<a class="back" href="#/room/${encodeURIComponent(i.LocationID)}">${backI} ${esc(i.Room)}</a>
    <div class="sub" style="font-family:Consolas,monospace">${esc(i.Title)}${pending(rowKey) ? ' · <span class="pend">change waiting to send</span>' : ""}</div>
    <h1 style="font-size:26px">${esc(nameOf(i))}</h1>
    <div class="sub">${esc(i.Activity ? discOf(i) : [i.Package, i.SystemName].filter(Boolean).join(" · "))}${oldName}</div>
    <div class="facts three">
      <div><span>Where</span><b>${esc(SHORT(i.Building))} · ${esc(i.Floor)}${i.Zone ? " · Zone " + esc(i.Zone) : ""}</b></div>
      <div><span>Work owner</span><b>${esc(i.WorkOwner || "–")}</b></div>
      <div><span>Target</span><b>${fmtD(i.TargetDate) || "–"}</b></div></div>
    <div class="lab">Progress</div>
    <div class="prog">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:10px"><span class="pctbig" id="pv"></span><span class="tag" id="stg"></span></div>
      <div id="sl">
        <input type="range" id="rg" class="rng" min="0" max="100" step="5" value="${Math.round(was / 5) * 5}" aria-label="${esc(nameOf(i))}, % complete">
        <div class="grid g5" id="qk">${[0, 25, 50, 75, 100].map(v => `<button type="button" class="qk" data-v="${v}">${v}</button>`).join("")}</div>
      </div>
      <div class="sub" id="pn" style="font-size:14px;line-height:1.35"></div>
    </div>
    <div class="lab">If it cannot move</div>
    <div class="grid g3" id="hl">${HOLDS.map(h => `<button type="button" class="stbtn ${h.replace(/[^A-Za-z]/g, "")}" data-h="${esc(h)}">${esc(h)}</button>`).join("")}</div>
    <div class="lab" id="phl">Photos</div><div class="photos" id="ph"></div>
    <label class="field"><span class="lab" style="margin:8px 0 0">Comment</span><textarea id="cm" placeholder="What was done, or what's stopping it">${esc(i.Notes || "")}</textarea></label>
    <div style="margin-top:18px"><button class="big navy" id="sv" type="button">Save</button></div>
    <p class="sub" id="why" style="text-align:center"></p>`;
  photoPicker($("#ph"), photos, existing);
  const paint = () => {
    const status = statusNow(), na = flag === "Not applicable", closing = status === "Closed" && st0 !== "Closed";
    $("#pv").textContent = na ? "–" : pctTxt(pct);
    const tg = $("#stg"); tg.className = "tag " + stc(status); tg.textContent = status;
    $("#sl").style.display = na ? "none" : "";
    $("#qk").querySelectorAll("[data-v]").forEach(b => b.classList.toggle("sel", +b.dataset.v === pct));
    $("#hl").querySelectorAll("[data-h]").forEach(b => b.classList.toggle("sel", b.dataset.h === flag));
    $("#pn").textContent = na ? "Not applicable: this activity is left out of every count and of the room progress."
      : flag ? `${flag}: progress stays at ${pctTxt(pct)} until it is cleared. Say why in the comment.`
      : closing ? `Closes today, ${fmtD(todayISO())}. The date goes to the tracker with the save.`
      : status === "Closed" ? `Closed${i.ClosureDate ? " " + fmtD(i.ClosureDate) : ""}. Slide back to reopen it.`
      : pct !== was ? `Total done so far. Was ${pctTxt(was)} at the last save.` : "Slide, or tap a figure, to the total % done so far.";
    $("#phl").textContent = C.photoToClose && closing && !existing.length ? "Photo (needed to close)" : "Photos";
    $("#sv").textContent = closing ? `Save · close ${i.Title}` : "Save";
  };
  const rg = $("#rg"); rg.oninput = () => { pct = +rg.value; paint(); };
  $("#qk").querySelectorAll("[data-v]").forEach(b => b.onclick = () => { pct = +b.dataset.v; rg.value = pct; paint(); });
  $("#hl").querySelectorAll("[data-h]").forEach(b => b.onclick = () => { flag = flag === b.dataset.h ? "" : b.dataset.h; paint(); });
  paint();
  $("#sv").onclick = async () => {
    const status = statusNow(), closingNow = status === "Closed" && st0 !== "Closed";
    if (C.photoToClose && closingNow && !photos.length && !existing.length) { $("#why").textContent = "Take a photo of the finished work to close this item."; toast("Add a photo to close"); return; }
    const fields = {}; const note = $("#cm").value.trim();
    const firstFigure = typeof i.Progress !== "number" && status !== st0;      // a row with no figure whose status changes (e.g. no longer Not applicable): send the figure too, or the tracker cannot take the status
    if ((pct !== was || firstFigure) && flag !== "Not applicable") fields.Progress = pct;
    if (status !== st0) fields.ItemStatus = status;
    if (closingNow) fields.ClosureDate = todayISO();
    if (status !== "Closed" && i.ClosureDate) fields.ClosureDate = null;
    if (note !== (i.Notes || "")) fields.Notes = note;
    if (!Object.keys(fields).length && !photos.length) { history.back(); return; }
    const local = { ...i, ...fields }; S.items.set(rowKey, local); await putOne("items", local);
    const detail = [fields.Progress !== undefined && pct !== was ? `${pctTxt(was)} to ${pctTxt(pct)}` : "", fields.ItemStatus ? (closingNow ? `closed ${fmtD(todayISO())}` : fields.ItemStatus) : "",
      fields.Notes !== undefined ? "comment" : ""].filter(Boolean).join(" · ") || "photo added";
    await queue({ type: "update", rowKey, fields, photos: photos.map(p => ({ blob: p.blob, at: p.at })), label: `${i.Room} · ${nameOf(i)}`, detail, room: i.Room });
    toast(navigator.onLine ? "Saved – sending…" : "Saved on this phone – sends when there's signal");
    location.hash = "#/room/" + encodeURIComponent(i.LocationID); sync();
  };
}

function vAdd(loc) {
  setTab("rooms");
  const r = S.rooms.get(loc); if (!r) return vRooms();
  const pk = new Map(); for (const i of S.items.values()) { if (!i.Package) continue; if (!pk.has(i.Package)) pk.set(i.Package, new Set()); if (i.SystemName) pk.get(i.Package).add(i.SystemName); }
  const pkgs = [...pk.keys()].sort(); const photos = [];
  const t14 = new Date(Date.now() + 14 * 864e5).toISOString().slice(0, 10);
  $("#view").innerHTML = `<a class="back" href="#/room/${encodeURIComponent(loc)}">${backI} Cancel</a>
    <h1>New works to go item</h1>
    <div class="where"><span>Where (filled in from the room)</span><b>${esc(r.Room)}${r.Zone ? " · Zone " + esc(r.Zone) : ""} · ${esc(r.Building)} · ${esc(r.Floor)}</b></div>
    <label class="field"><span class="lab" style="margin:0">What needs doing</span><textarea id="ds" placeholder="e.g. Door signage to install"></textarea></label>
    <div class="grid g2"><label class="field"><span class="lab" style="margin:0">Package</span><select id="pk"><option value="">Choose…</option>${pkgs.map(p => `<option>${esc(p)}</option>`).join("")}</select></label>
      <label class="field"><span class="lab" style="margin:0">System</span><select id="sy"></select></label></div>
    <div class="grid g2"><label class="field"><span class="lab" style="margin:0">Work owner</span><input id="wo" value="${esc(r.WorkOwner || "")}"></label>
      <label class="field"><span class="lab" style="margin:0">Target date</span><input id="td" type="date" value="${t14}"></label></div>
    <div class="lab">Photos</div><div class="photos" id="ph"></div>
    <div style="margin-top:18px"><button class="big navy" id="sv" type="button">Save works to go item</button></div>
    <p class="sub" style="text-align:center">Raised today by ${esc(whoNow() || (S.user ? S.user.name : "you"))}. The WTG number is given when it reaches SharePoint.</p>`;
  const fillSys = () => { const set = [...(pk.get($("#pk").value) || [])].sort(); $("#sy").innerHTML = `<option value="">–</option>` + set.map(s => `<option>${esc(s)}</option>`).join(""); };
  $("#pk").onchange = fillSys; fillSys();
  photoPicker($("#ph"), photos, []);
  $("#sv").onclick = async () => {
    const ds = $("#ds").value.trim(); if (!ds) { toast("Say what needs doing"); $("#ds").focus(); return; }
    if (!$("#pk").value) { toast("Choose the package"); $("#pk").focus(); return; }
    const rowKey = uid();
    const fields = { ItemType: "Works to Go", Descr: ds, Building: r.Building, Floor: r.Floor, Room: r.Room, Zone: r.Zone || "", Package: $("#pk").value,
      SystemName: $("#sy").value, WorkOwner: $("#wo").value.trim(), UCBResp: r.UCBResp || "", UCBResp2: r.UCBResp2 || "", DateRaised: todayISO(),
      TargetDate: $("#td").value ? $("#td").value + "T00:00:00Z" : null, ItemStatus: "Open", LocationID: loc, SourceRef: "Site Capture app" };
    const local = { ...fields, RowKey: rowKey, Title: "NEW (not sent yet)" }; S.items.set(rowKey, local); await putOne("items", local);
    await queue({ type: "create", rowKey, fields, photos: photos.map(p => ({ blob: p.blob, at: p.at })), label: `New · ${ds}`, room: r.Room });
    toast(navigator.onLine ? "Saved – sending…" : "Saved on this phone – sends when there's signal");
    location.hash = "#/room/" + encodeURIComponent(loc); sync();
  };
}

async function vOutbox() {
  setTab("outbox");
  const sent = (await meta.get("sent")) || [], unsent = (await meta.get("unsent")) || [];
  const hm = iso => { const d = new Date(iso); return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`; };
  const n = S.outbox.length;
  $("#view").innerHTML = `<h1>Outbox</h1>
    ${n ? `<div class="banner warn"><div><b>${n} change${n > 1 ? "s" : ""} saved on this phone</b><br>${navigator.onLine ? "Sending now – keep the app open for a moment." : "There's no signal here. Nothing is lost: they send by themselves when you're back in signal. Carry on working."}</div></div>`
         : `<div class="banner ok"><div><b>Everything is sent.</b><br>Last checked ${S.lastSync ? fmtD(S.lastSync) + " " + hm(S.lastSync) : "–"}.</div></div>`}
    <div class="lab">Waiting to send</div><div class="list">${S.outbox.slice().reverse().map(o => `<div class="row"><span class="cnt" style="width:52px;text-align:left">${hm(o.at)}</span><span class="grow"><span class="nm">${esc(o.label)}</span><span class="s">${esc(o.detail || o.room || "")}${o.photos && o.photos.length ? ` · ${o.photos.length} photo${o.photos.length > 1 ? "s" : ""}` : ""}${o.err ? `<br><b style="color:var(--warn,#B4531A)">${esc(o.err)}</b>` : ""}</span></span><span class="pend">Waiting</span></div>`).join("") || `<div class="empty">Nothing waiting.</div>`}</div>
    ${unsent.length ? `<div class="lab">Not sent – row no longer on the list (${unsent.length})</div><div class="list">${unsent.slice(0, 60).map(o => `<div class="row"><span class="cnt" style="width:52px;text-align:left">${hm(o.at)}</span><span class="grow"><span class="nm">${esc(o.what)}</span><span class="s">${esc(o.detail || o.room || "")}${o.photos && o.photos.length ? ` · ${o.photos.length} photo${o.photos.length > 1 ? "s" : ""} kept on this phone` : ""} · ${fmtD(o.at)}</span></span><span class="pend" style="background:none">Set aside</span></div>`).join("")}</div>
      <p class="sub">These changes were made to rows that have since been taken off the list in the office, so there is nowhere to send them. Tell the office if any of them matter.</p>` : ""}
    <div class="lab">Sent recently</div><div class="list">${sent.slice(0, 10).map(o => `<div class="row"><span class="cnt" style="width:52px;text-align:left">${hm(o.at)}</span><span class="grow"><span class="nm">${esc(o.what)}</span><span class="s">${esc(o.room || "")}</span></span><span style="color:var(--ok);font-weight:700">Sent</span></div>`).join("") || `<div class="empty">Nothing sent from this phone yet.</div>`}</div>
    ${isShared() ? `<div class="where" style="display:flex;justify-content:space-between;align-items:center;gap:10px"><div><span>Using this tablet today</span><b id="whon">${esc(whoNow())}</b></div><button class="big line" id="wc" type="button" style="width:auto;padding:10px 16px">Change</button></div>` : ""}
    <div style="margin-top:14px" class="grid g2"><button class="big line" id="sy" type="button">Sync now</button><button class="big line" id="so" type="button">Sign out</button></div>
    <div class="office"><div style="display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap"><b style="color:var(--navy);font-size:18px">Office · Friday close-out</b><span class="sub">${esc(S.user ? S.user.email : "")}</span></div>
      <p>Exports every activity with its Progress %, status and closure date, under the tracker's column names, for the Friday tracker update.</p>
      <button class="big amber" id="ex" type="button">Export full list to Excel</button></div>
    <p class="sub" style="text-align:center;margin-top:18px">${S.items.size.toLocaleString("en-GB")} items · ${S.rooms.size} rooms on this phone · v${esc(C.version)}</p>`;
  $("#sy").onclick = () => sync(true);
  $("#so").onclick = () => { if (S.outbox.length && !confirm("There are changes not sent yet. Sign out anyway? They stay on this phone.")) return; signOut(); };
  $("#ex").onclick = exportXlsx;
  if ($("#wc")) $("#wc").onclick = () => { S.whoDay = null; paintWho(); route(); };
}

/* ---------------- export ---------------- */
let xlsxP;
function loadXlsx() { if (!xlsxP) xlsxP = new Promise((res, rej) => { const s = document.createElement("script"); s.src = "vendor/xlsx.full.min.js"; s.onload = res; s.onerror = rej; document.head.appendChild(s); }); return xlsxP; }
async function exportXlsx() {
  if (S.outbox.length && !confirm(`${S.outbox.length} change(s) on this phone haven't reached SharePoint yet. Export anyway (they are included)?`)) return;
  await loadXlsx();
  const H = ["ID","Type","Description","Building","Floor","Room","Zone / Area","Package","System","Work Owner (Contractor / Supplier)","UCB Responsible","UCB Responsible 2","Category","Date Raised","Target Date","Status","Progress %","Closure Date","Comment","Source (migration)","Activity","Last changed by","Last changed"];
  const d = iso => iso ? new Date(iso) : null;
  const rows = [...S.items.values()].sort((a, b) => String(a.Title).localeCompare(String(b.Title), "en", { numeric: true }))
    .map(i => [i.Title, i.ItemType || "Works to Go", i.Descr || "", i.Building || "", i.Floor || "", i.Room || "", i.Zone || "", i.Package || "", i.SystemName || "",
      i.WorkOwner || "", i.UCBResp || "", i.UCBResp2 || "", i.Cat || "", d(i.DateRaised), d(i.TargetDate), i.ItemStatus || "", typeof i.Progress === "number" ? i.Progress : null,
      d(i.ClosureDate), i.Notes || "", i.SourceRef || "", i.Activity || "", i.LastBy || "", d(i.LastAt)]);
  const ws = XLSX.utils.aoa_to_sheet([H, ...rows], { cellDates: true, dateNF: "d mmm yyyy" });
  ws["!cols"] = H.map((h, k) => ({ wch: [10, 12, 34, 16, 8, 20, 14, 14, 18, 14, 12, 12, 8, 12, 12, 12, 11, 12, 30, 30, 28, 26, 12][k] }));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Works to Go Register");
  wb.Props = { Title: "Project Apple – Works to Go export", Author: "Project Central Assurance", Company: "Project Central Assurance" };
  const n = new Date(); const stamp = `${n.getDate()} ${MON[n.getMonth()]} ${n.getFullYear()}`;
  XLSX.writeFile(wb, `Project Apple - Works to Go export - ${stamp}.xlsx`, { compression: true });
  toast(`Exported ${rows.length.toLocaleString("en-GB")} items`);
}

/* ---------------- router ---------------- */
function route() {
  if (!account) return vSignIn();
  if (needWho()) return vWho();
  if (!S.rooms.size) return vLoading(navigator.onLine ? "Loading rooms and items for the first time…" : "Connect to signal once to load the rooms.");
  const h = decodeURIComponent(location.hash.slice(1) || "/");
  const [, a, b] = h.split("/");
  window.scrollTo(0, 0);
  if (a === "room") return vRoom(b); if (a === "item") return vItem(b); if (a === "add") return vAdd(b); if (a === "outbox") return vOutbox();
  return vRooms();
}

/* ---------------- start ---------------- */
async function start() {
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(e => console.warn(e));
  await loadLocal(); paintChip();
  { const fk = FIELDS.join(","), had = await meta.get("fkey");                                    // a full reload only when the fields change, not at every app update
    if ((had && had !== fk) || (!had && /^1\.0\./.test((await meta.get("ver")) || ""))) await meta.set("deltaItems", null);   // v1.0 phones lack Activity and Progress
    if (had && (await meta.get("recon")) !== "1") await meta.set("deltaItems", null);              // once, on phones loaded before v1.2.2: a whole-list read, so rows since taken off the list are dropped
    await meta.set("fkey", fk); await meta.set("ver", C.version); }
  try { await initAuth(); } catch (e) { console.error(e); }
  paintWho(); route();
  window.addEventListener("hashchange", route);
  window.addEventListener("online", () => { paintChip(); sync(); });
  window.addEventListener("offline", paintChip);
  $("#netchip").onclick = () => location.hash = "#/outbox";
  if (account && navigator.onLine) {
    if (!S.rooms.size) { try { await siteIds(); await refresh(); S.lastSync = new Date().toISOString(); await meta.set("lastSync", S.lastSync); } catch (e) { console.error(e); vLoading("Couldn't load from SharePoint: " + e.message); return; } route(); }
    sync();
  }
  setInterval(() => { if (navigator.onLine && S.outbox.length) sync(); }, 60000);
  document.addEventListener("visibilitychange", () => { if (document.visibilityState !== "visible") return; if (needWho()) { paintWho(); route(); } if (navigator.onLine) sync(); });   // a tablet left open overnight asks again in the morning
}
start();
})();
