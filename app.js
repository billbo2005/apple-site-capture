/* Project Apple · Site Capture – works-to-go capture app (offline-first PWA, Microsoft 365 / SharePoint back end)
   Project Central Assurance, Nov 2026 */
(() => {
"use strict";
const C = window.APP_CONFIG;
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["Sites.ReadWrite.All", "User.Read"];
const STATUSES = ["Open", "In Progress", "Closed", "On Hold", "At Risk", "Not applicable"];
const FIELDS = ["Title","ItemType","Descr","Building","Floor","Room","Zone","Package","SystemName","WorkOwner","UCBResp","UCBResp2","Cat",
  "DateRaised","TargetDate","ItemStatus","ClosureDate","Notes","SourceRef","LocationID","RowKey","LastBy","LastAt","PhotoIds"];
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
const meta = { get: k => tx("meta", "readonly", s => req(s.get(k))), set: (k, v) => tx("meta", "readwrite", s => req(s.put(v, k))) };

/* ---------------- state ---------------- */
const S = { items: new Map(), rooms: new Map(), outbox: [], user: null, syncing: false, lastSync: null, filter: { b: null, f: null, q: "" }, roomPill: "All" };
async function loadLocal() {
  (await getAll("items")).forEach(i => S.items.set(i.RowKey, i));
  (await getAll("rooms")).forEach(r => S.rooms.set(r.Title, r));
  S.outbox = await getAll("outbox");
  S.lastSync = await meta.get("lastSync");
  const f = await meta.get("filter"); if (f) S.filter = { ...S.filter, ...f, q: "" };
}

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
  const r = await fetch(path.startsWith("http") ? path : GRAPH + path, { ...opt, headers: { Authorization: "Bearer " + t, ...(opt.headers || {}) } });
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
  if (!S.rooms.size) { const rooms = []; await pullAll(site.rooms, ROOMF, r => rooms.push(r)); await putMany("rooms", rooms); rooms.forEach(r => S.rooms.set(r.Title, r)); }
  const changed = [];
  const pend = new Set(S.outbox.map(o => o.rowKey));
  const byId = new Map([...S.items.values()].map(i => [i.spId, i.RowKey]));
  const onRow = r => { if (!r.RowKey) return; if (pend.has(r.RowKey)) return; S.items.set(r.RowKey, r); changed.push(r); };
  const onDel = id => { const k = byId.get(id); if (k && !pend.has(k)) { S.items.delete(k); delOne("items", k); } };
  try { await pullDelta(site.items, FIELDS, "deltaItems", onRow, onDel); }
  catch (e) { console.warn("delta failed, full pull", e); await pullAll(site.items, FIELDS, onRow); }
  await putMany("items", changed);
}

/* ---------------- outbox ---------------- */
async function queue(op) { op.at = new Date().toISOString(); op.by = S.user ? S.user.email : ""; const qid = await putOne("outbox", op); op.qid = qid; S.outbox.push(op); paintChip(); }
function nextId() { let m = 0; for (const i of S.items.values()) { const n = /^WTG-(\d+)$/.exec(i.Title || ""); if (n) m = Math.max(m, +n[1]); } return "WTG-" + String(m + 1).padStart(4, "0"); }
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
  let sent = 0, failed = 0;
  try {
    await siteIds();
    await refresh();                                        // fresh numbers before creating new items
    for (const op of [...S.outbox]) {
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
        const sentLog = (await meta.get("sent")) || []; sentLog.unshift({ at: new Date().toISOString(), what: op.label, room: op.room });
        await meta.set("sent", sentLog.slice(0, 30)); sent++;
      } catch (e) { console.error(e); failed++; }
    }
    S.lastSync = new Date().toISOString(); await meta.set("lastSync", S.lastSync);
    if (manual || sent) toast(failed ? `${sent} sent, ${failed} will retry` : sent ? `${sent} change${sent > 1 ? "s" : ""} sent` : "Up to date");
  } catch (e) { console.error(e); if (manual) toast("Couldn't reach SharePoint – will try again"); }
  finally { S.syncing = false; paintChip(); if (!document.activeElement || !/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName)) route(); if (S.again || (sent && S.outbox.length)) setTimeout(() => sync(), 300); }
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
function roomStats(loc) { let t = 0, c = 0, o = 0; for (const i of S.items.values()) if (i.LocationID === loc && counted(i)) { t++; if (i.ItemStatus === "Closed") c++; else o++; } return { t, c, o }; }
function pending(rowKey) { return S.outbox.some(o => o.rowKey === rowKey); }
const chev = '<svg class="chev" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg>';
const backI = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" aria-hidden="true"><path d="M15 6l-6 6 6 6"/></svg>';
const camI = '<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round" aria-hidden="true"><path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="4"/></svg>';

function paintChip() {
  const ch = $("#netchip"); const n = S.outbox.length;
  if (S.syncing) { ch.className = "chip busy"; ch.textContent = "Sending…"; }
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
    let zt = 0, zc = 0; rs.forEach(r => { const s = roomStats(r.Title); zt += s.t; zc += s.c; });
    html += `<div class="zonehd"><b>${esc(z)}</b><span class="sub">${rs.length} rooms · ${zc} of ${zt} closed${zt ? ` (${Math.round(zc / zt * 100)}%)` : ""}</span></div><div class="list">`;
    for (const r of rs) { const s = roomStats(r.Title); const p = s.t ? Math.round(s.c / s.t * 100) : 0;
      html += `<a class="row" href="#/room/${encodeURIComponent(r.Title)}"><span class="grow"><span class="nm">${esc(r.Room)}</span>
        <span class="bar" aria-hidden="true"><span style="width:${Math.max(p, 3)}%;background:${BAND(p)}"></span></span></span>
        <span class="cnt">${s.c} of ${s.t} closed<br><span class="s">${s.o} open</span></span>${chev}</a>`; }
    html += `</div>`;
  }
  html += `<div class="legend"><span>Bar = share closed:</span><span><i style="background:var(--b0)"></i>0–24%</span><span><i style="background:var(--b25)"></i>25–49%</span><span><i style="background:var(--b50)"></i>50–74%</span><span><i style="background:var(--b75)"></i>75–99%</span><span><i style="background:var(--b100)"></i>100%</span></div>`;
  $("#view").innerHTML = html;
  $("#view").querySelectorAll("[data-b]").forEach(b => b.onclick = () => { S.filter = { b: b.dataset.b, f: null, q: "" }; meta.set("filter", S.filter); vRooms(); });
  $("#view").querySelectorAll("[data-f]").forEach(b => b.onclick = () => { S.filter.f = b.dataset.f; S.filter.q = ""; meta.set("filter", S.filter); vRooms(); });
  const qi = $("#q"); qi.oninput = () => { S.filter.q = qi.value; clearTimeout(qi._t); qi._t = setTimeout(() => { vRooms(); const n = $("#q"); n.focus(); n.setSelectionRange(n.value.length, n.value.length); }, 250); };
}

function vRoom(loc) {
  setTab("rooms");
  const r = S.rooms.get(loc); if (!r) return vRooms();
  const items = [...S.items.values()].filter(i => i.LocationID === loc).sort((a, b) => String(a.Title).localeCompare(String(b.Title)));
  const s = roomStats(loc);
  const cnt = st => items.filter(i => (i.ItemStatus || "Open") === st).length;
  const pills = ["All", "Open", "In Progress", "Closed"].concat(["On Hold", "At Risk", "Not applicable"].filter(cnt));
  const pill = S.roomPill;
  const shown = pill === "All" ? items : items.filter(i => (i.ItemStatus || "Open") === pill);
  $("#view").innerHTML = `<a class="back" href="#/">${backI} ${esc(SHORT(r.Building))} · ${esc(r.Floor)} rooms</a>
    <div style="display:flex;justify-content:space-between;align-items:flex-end;gap:12px;flex-wrap:wrap">
      <div><h1>${esc(r.Room)}</h1><div class="sub">${r.Zone ? "Zone " + esc(r.Zone) + " · " : ""}work owner ${esc(r.WorkOwner || "–")} · UCB ${esc(r.UCBResp || "–")}</div></div>
      <div style="text-align:right"><div style="font-size:28px;font-weight:700">${s.c} of ${s.t}</div><div class="sub">activities closed</div></div></div>
    <div style="margin-top:14px"><a class="big amber" href="#/add/${encodeURIComponent(loc)}">＋ Add works to go item</a></div>
    <div class="pills">${pills.map(p => `<button type="button" class="pill ${p === pill ? "sel" : ""}" data-p="${esc(p)}">${esc(p)} · ${p === "All" ? items.length : cnt(p)}</button>`).join("")}</div>
    <div class="list">${shown.map(i => `<a class="row" href="#/item/${encodeURIComponent(i.RowKey)}">
      <span class="grow"><span class="nm">${esc(i.Descr || "(no description)")}</span><span class="s">${esc(i.Title)} · ${esc([i.Package, i.SystemName].filter(Boolean).join(" · "))}${pending(i.RowKey) ? ' · <span class="pend">waiting to send</span>' : ""}</span></span>
      <span class="tag ${stc(i.ItemStatus)}">${esc(i.ItemStatus || "Open")}</span>${chev}</a>`).join("") || `<div class="empty">No items.</div>`}</div>`;
  $("#view").querySelectorAll("[data-p]").forEach(b => b.onclick = () => { S.roomPill = b.dataset.p; vRoom(loc); });
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
  let status = i.ItemStatus || "Open"; const photos = []; const existing = String(i.PhotoIds || "").split("\n").filter(Boolean);
  const days = i.TargetDate ? Math.round((new Date(i.TargetDate) - new Date(todayISO())) / 864e5) : null;
  $("#view").innerHTML = `<a class="back" href="#/room/${encodeURIComponent(i.LocationID)}">${backI} ${esc(i.Room)}</a>
    <div class="sub" style="font-family:Consolas,monospace">${esc(i.Title)}${pending(rowKey) ? ' · <span class="pend">change waiting to send</span>' : ""}</div>
    <h1>${esc(i.Descr || "(no description)")}</h1>
    <div class="facts">
      <div><span>Room · zone</span><b>${esc(i.Room)}${i.Zone ? " · " + esc(i.Zone) : ""}</b></div>
      <div><span>Package · system</span><b>${esc(i.Package || "–")} · ${esc(i.SystemName || "–")}</b></div>
      <div><span>Work owner · UCB</span><b>${esc(i.WorkOwner || "–")} · ${esc(i.UCBResp || "–")}</b></div>
      <div><span>Raised</span><b>${fmtD(i.DateRaised) || "–"}</b></div>
      <div><span>Target</span><b>${fmtD(i.TargetDate) || "–"}</b></div>
      <div><span>Days to target</span><b>${days === null ? "–" : days + " cd"}</b></div></div>
    <div class="lab">Status</div><div class="grid g3" id="sts">${STATUSES.map(s => `<button type="button" class="stbtn ${s.replace(/[^A-Za-z]/g, "")}" data-s="${esc(s)}">${esc(s)}</button>`).join("")}</div>
    <div id="cd" class="where" style="display:none"><span>Closure date</span><b id="cdv"></b></div>
    <div class="lab" id="phl">Photos</div><div class="photos" id="ph"></div>
    <label class="field"><span class="lab" style="margin:8px 0 0">Comment</span><textarea id="cm" placeholder="What was done, or what's stopping it">${esc(i.Notes || "")}</textarea></label>
    <div style="margin-top:18px"><button class="big navy" id="sv" type="button">Save</button></div>
    <p class="sub" id="why" style="text-align:center"></p>`;
  photoPicker($("#ph"), photos, existing);
  const paint = () => {
    $("#sts").querySelectorAll("[data-s]").forEach(b => b.classList.toggle("sel", b.dataset.s === status));
    const closing = status === "Closed";
    $("#cd").style.display = closing ? "" : "none";
    $("#cdv").textContent = closing ? (i.ItemStatus === "Closed" && i.ClosureDate ? fmtD(i.ClosureDate) : fmtD(todayISO()) + " (today)") : "";
    $("#phl").textContent = closing && !existing.length ? "Photo (needed to close)" : "Photos";
    $("#sv").textContent = closing && i.ItemStatus !== "Closed" ? `Save · close ${i.Title}` : "Save";
  };
  $("#sts").querySelectorAll("[data-s]").forEach(b => b.onclick = () => { status = b.dataset.s; paint(); });
  paint();
  $("#sv").onclick = async () => {
    const closingNow = status === "Closed" && i.ItemStatus !== "Closed";
    if (closingNow && !photos.length && !existing.length) { $("#why").textContent = "Take a photo of the finished work to close this item."; toast("Add a photo to close"); return; }
    const fields = {}; const note = $("#cm").value.trim();
    if (status !== (i.ItemStatus || "Open")) fields.ItemStatus = status;
    if (closingNow) fields.ClosureDate = todayISO();
    if (status !== "Closed" && i.ClosureDate) fields.ClosureDate = null;
    if (note !== (i.Notes || "")) fields.Notes = note;
    if (!Object.keys(fields).length && !photos.length) { history.back(); return; }
    const local = { ...i, ...fields }; S.items.set(rowKey, local); await putOne("items", local);
    const label = `${i.Title} ${i.Descr || ""}` + (fields.ItemStatus ? ` → ${fields.ItemStatus}` : photos.length ? " · photo added" : " · comment");
    await queue({ type: "update", rowKey, fields, photos: photos.map(p => ({ blob: p.blob, at: p.at })), label, room: i.Room });
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
    <p class="sub" style="text-align:center">Raised today by ${esc(S.user ? S.user.name : "you")}. The WTG number is given when it reaches SharePoint.</p>`;
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
  const sent = (await meta.get("sent")) || [];
  const hm = iso => { const d = new Date(iso); return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`; };
  const n = S.outbox.length;
  $("#view").innerHTML = `<h1>Outbox</h1>
    ${n ? `<div class="banner warn"><div><b>${n} change${n > 1 ? "s" : ""} saved on this phone</b><br>${navigator.onLine ? "Sending now – keep the app open for a moment." : "There's no signal here. Nothing is lost: they send by themselves when you're back in signal. Carry on working."}</div></div>`
         : `<div class="banner ok"><div><b>Everything is sent.</b><br>Last checked ${S.lastSync ? fmtD(S.lastSync) + " " + hm(S.lastSync) : "–"}.</div></div>`}
    <div class="lab">Waiting to send</div><div class="list">${S.outbox.slice().reverse().map(o => `<div class="row"><span class="cnt" style="width:52px;text-align:left">${hm(o.at)}</span><span class="grow"><span class="nm">${esc(o.label)}</span><span class="s">${esc(o.room || "")}${o.photos && o.photos.length ? ` · ${o.photos.length} photo${o.photos.length > 1 ? "s" : ""}` : ""}</span></span><span class="pend">Waiting</span></div>`).join("") || `<div class="empty">Nothing waiting.</div>`}</div>
    <div class="lab">Sent recently</div><div class="list">${sent.slice(0, 10).map(o => `<div class="row"><span class="cnt" style="width:52px;text-align:left">${hm(o.at)}</span><span class="grow"><span class="nm">${esc(o.what)}</span><span class="s">${esc(o.room || "")}</span></span><span style="color:var(--ok);font-weight:700">Sent</span></div>`).join("") || `<div class="empty">Nothing sent from this phone yet.</div>`}</div>
    <div style="margin-top:14px" class="grid g2"><button class="big line" id="sy" type="button">Sync now</button><button class="big line" id="so" type="button">Sign out</button></div>
    <div class="office"><div style="display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap"><b style="color:var(--navy);font-size:18px">Office · Friday close-out</b><span class="sub">${esc(S.user ? S.user.email : "")}</span></div>
      <p>Exports every item in the same column order as the tracker's Works to Go Register (A to S), ready to paste over it with Paste Values.</p>
      <button class="big amber" id="ex" type="button">Export full list to Excel</button></div>
    <p class="sub" style="text-align:center;margin-top:18px">${S.items.size.toLocaleString("en-GB")} items · ${S.rooms.size} rooms on this phone · v${esc(C.version)}</p>`;
  $("#sy").onclick = () => sync(true);
  $("#so").onclick = () => { if (S.outbox.length && !confirm("There are changes not sent yet. Sign out anyway? They stay on this phone.")) return; signOut(); };
  $("#ex").onclick = exportXlsx;
}

/* ---------------- export ---------------- */
let xlsxP;
function loadXlsx() { if (!xlsxP) xlsxP = new Promise((res, rej) => { const s = document.createElement("script"); s.src = "vendor/xlsx.full.min.js"; s.onload = res; s.onerror = rej; document.head.appendChild(s); }); return xlsxP; }
async function exportXlsx() {
  if (S.outbox.length && !confirm(`${S.outbox.length} change(s) on this phone haven't reached SharePoint yet. Export anyway (they are included)?`)) return;
  await loadXlsx();
  const H = ["ID","Type","Description","Building","Floor","Room","Zone / Area","Package","System","Work Owner (Contractor / Supplier)","UCB Responsible","UCB Responsible 2","Category","Date Raised","Target Date","Status","Closure Date","Comment","Source (migration)"];
  const d = iso => iso ? new Date(iso) : null;
  const rows = [...S.items.values()].sort((a, b) => String(a.Title).localeCompare(String(b.Title), "en", { numeric: true }))
    .map(i => [i.Title, i.ItemType || "Works to Go", i.Descr || "", i.Building || "", i.Floor || "", i.Room || "", i.Zone || "", i.Package || "", i.SystemName || "",
      i.WorkOwner || "", i.UCBResp || "", i.UCBResp2 || "", i.Cat || "", d(i.DateRaised), d(i.TargetDate), i.ItemStatus || "", d(i.ClosureDate), i.Notes || "", i.SourceRef || ""]);
  const ws = XLSX.utils.aoa_to_sheet([H, ...rows], { cellDates: true, dateNF: "d mmm yyyy" });
  ws["!cols"] = H.map((h, k) => ({ wch: [10, 12, 34, 16, 8, 20, 14, 14, 18, 14, 12, 12, 8, 12, 12, 12, 12, 30, 30][k] }));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Works to Go Register");
  wb.Props = { Title: "Project Apple – Works to Go export", Author: "Project Central Assurance", Company: "Project Central Assurance" };
  const n = new Date(); const stamp = `${n.getDate()} ${MON[n.getMonth()]} ${n.getFullYear()}`;
  XLSX.writeFile(wb, `Project Apple - Works to Go export - ${stamp}.xlsx`, { compression: true });
  toast(`Exported ${rows.length.toLocaleString("en-GB")} items`);
}

/* ---------------- router ---------------- */
function route() {
  if (!account) return vSignIn();
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
  try { await initAuth(); } catch (e) { console.error(e); }
  route();
  window.addEventListener("hashchange", route);
  window.addEventListener("online", () => { paintChip(); sync(); });
  window.addEventListener("offline", paintChip);
  $("#netchip").onclick = () => location.hash = "#/outbox";
  if (account && navigator.onLine) {
    if (!S.rooms.size) { try { await siteIds(); await refresh(); S.lastSync = new Date().toISOString(); await meta.set("lastSync", S.lastSync); } catch (e) { console.error(e); vLoading("Couldn't load from SharePoint: " + e.message); return; } route(); }
    sync();
  }
  setInterval(() => { if (navigator.onLine && S.outbox.length) sync(); }, 60000);
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible" && navigator.onLine) sync(); });
}
start();
})();
