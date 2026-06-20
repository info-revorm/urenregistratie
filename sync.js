/* ============================================================
   revorm urenregistratie — cloud-synchronisatie
   MSAL (Microsoft-login) + Microsoft Graph (OneDrive app-map)
   + offline-wachtrij + veilige 3-weg merge.

   Hooks die de app (index.html) levert (globaal):
     getState() / setState(s) / defaultState() / normalizeState(s)
     saveLocal() / rerenderAll() / seedIfEmpty() / stateIsEmpty()
     setSyncStatus(text,kind) / updateAuthUI(account) / STORE_KEY
   Deze module levert terug (globaal):
     scheduleCloudFlush() / msSignIn() / msSignOut() / msSyncNow() / initSync()
   ============================================================ */

let msalInstance = null;
let flushTimer = null;
let busy = false;

/* ---------- localStorage helpers voor sync-metadata ---------- */
function lsGet(k){ try{ return localStorage.getItem(k); }catch(e){ return null; } }
function lsSet(k,v){ try{ localStorage.setItem(k,v); }catch(e){} }
function lsDel(k){ try{ localStorage.removeItem(k); }catch(e){} }

function baseKey(){ return STORE_KEY + "__base"; }
function etagKey(){ return STORE_KEY + "__etag"; }
function dirtyKey(){ return STORE_KEY + "__dirty"; }

function markDirty(){ lsSet(dirtyKey(), "1"); }
function clearDirty(){ lsDel(dirtyKey()); }
function isDirty(){ return lsGet(dirtyKey()) === "1"; }
function getETag(){ return lsGet(etagKey()); }
function setETag(e){ if(e) lsSet(etagKey(), e); }
function loadBase(){ try{ const r=lsGet(baseKey()); return r?JSON.parse(r):null; }catch(e){ return null; } }
function saveBase(s){ try{ lsSet(baseKey(), JSON.stringify(s)); }catch(e){} }

function isSignedIn(){ return !!(msalInstance && msalInstance.getActiveAccount()); }

/* ---------- 3-weg merge (verliest geen edits bij normaal gebruik) ---------- */
function eq(a,b){ return JSON.stringify(a) === JSON.stringify(b); }

function merge3Map(base, server, local){
  const out = {};
  const keys = new Set([
    ...Object.keys(server||{}), ...Object.keys(local||{}), ...Object.keys(base||{})
  ]);
  keys.forEach(k=>{
    const b = base ? base[k] : undefined;
    const s = server ? server[k] : undefined;
    const l = local ? local[k] : undefined;
    const lc = !eq(l,b), sc = !eq(s,b);
    let pick;
    if(lc && !sc) pick = l;          // alleen lokaal gewijzigd
    else if(sc && !lc) pick = s;     // alleen server gewijzigd
    else if(!lc && !sc) pick = s;    // niets gewijzigd
    else pick = l;                   // beide gewijzigd -> lokaal wint
    if(pick !== undefined) out[k] = pick;
  });
  return out;
}

function byId(arr){ const m={}; (arr||[]).forEach(x=>{ if(x && x.id) m[x.id]=x; }); return m; }

function merge3List(base, server, local){
  const bm=byId(base), sm=byId(server), lm=byId(local);
  const ids=[], seen=new Set();
  (local||[]).forEach(x=>{ if(x&&!seen.has(x.id)){ ids.push(x.id); seen.add(x.id);} });
  (server||[]).forEach(x=>{ if(x&&!seen.has(x.id)){ ids.push(x.id); seen.add(x.id);} });
  (base||[]).forEach(x=>{ if(x&&!seen.has(x.id)){ ids.push(x.id); seen.add(x.id);} });
  const out=[];
  ids.forEach(id=>{
    const b=bm[id], s=sm[id], l=lm[id];
    const lc=!eq(l,b), sc=!eq(s,b);
    let pick;
    if(lc && !sc) pick=l;
    else if(sc && !lc) pick=s;
    else if(!lc && !sc) pick=s;
    else pick=l;                     // beide gewijzigd -> lokaal wint
    if(pick !== undefined) out.push(pick);
  });
  return out;
}

function mergeStates(base, server, local){
  return {
    settings: merge3Map(base.settings||{}, server.settings||{}, local.settings||{}),
    clients:  merge3List(base.clients||[],  server.clients||[],  local.clients||[]),
    projects: merge3List(base.projects||[], server.projects||[], local.projects||[]),
    entries:  merge3Map(base.entries||{},   server.entries||{},   local.entries||{})
  };
}

/* ---------- Microsoft Graph ---------- */
function httpErr(res){ const e=new Error("HTTP "+res.status); e.status=res.status; return e; }
const GRAPH_BASE = "https://graph.microsoft.com/v1.0/me/drive/special/approot:/";

async function getToken(){
  const account = msalInstance.getActiveAccount();
  if(!account){ const e=new Error("not-signed-in"); throw e; }
  try{
    const r = await msalInstance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
    return r.accessToken;
  }catch(e){
    if(typeof msal !== "undefined" && e instanceof msal.InteractionRequiredAuthError){
      await msalInstance.acquireTokenRedirect({ scopes: GRAPH_SCOPES });
    }
    throw e;
  }
}

async function getServer(){
  const token = await getToken();
  const fileUrl = GRAPH_BASE + encodeURIComponent(GRAPH_FILE);
  const metaRes = await fetch(fileUrl, { headers:{ Authorization:"Bearer "+token } });
  if(metaRes.status === 404) return { status:404 };
  if(!metaRes.ok) throw httpErr(metaRes);
  const meta = await metaRes.json();
  const contentRes = await fetch(fileUrl + ":/content", { headers:{ Authorization:"Bearer "+token } });
  if(contentRes.status === 404) return { status:404 };
  if(!contentRes.ok) throw httpErr(contentRes);
  let parsed = null;
  try{ parsed = JSON.parse(await contentRes.text()); }catch(e){ parsed = null; }
  return { status:200, eTag: meta.eTag || meta.cTag, state: parsed };
}

async function putServer(stateObj, eTag){
  const token = await getToken();
  const url = GRAPH_BASE + encodeURIComponent(GRAPH_FILE) + ":/content";
  const headers = { Authorization:"Bearer "+token, "Content-Type":"application/json" };
  if(eTag) headers["If-Match"] = eTag;
  const res = await fetch(url, { method:"PUT", headers, body: JSON.stringify(stateObj) });
  if(res.status === 412){ const e=new Error("conflict"); e.status=412; throw e; }
  if(!res.ok) throw httpErr(res);
  const item = await res.json();
  return { eTag: item.eTag || item.cTag };
}

/* ---------- Sync-orkestratie ---------- */
function scheduleCloudFlush(){
  markDirty();
  if(!isSignedIn()) return;                 // lokaal blijven tot aangemeld
  setSyncStatus("Wijziging opslaan…", "sync");
  clearTimeout(flushTimer);
  flushTimer = setTimeout(function(){ flushToCloud(); }, 1500);
}

async function flushToCloud(){
  if(!isSignedIn()) return;
  if(!navigator.onLine){ setSyncStatus("Offline — wijzigingen worden later gesynct", "off"); return; }
  if(busy) return;
  if(!getETag()){ await syncFromCloud(); return; }   // etag onbekend -> eerst veilig synchroniseren
  busy = true;
  try{
    setSyncStatus("Synchroniseren…", "sync");
    let put;
    try{
      put = await putServer(getState(), getETag());
    }catch(err){
      if(err.status === 412){ busy = false; await syncFromCloud(); return; }
      throw err;
    }
    saveBase(getState()); setETag(put.eTag); clearDirty();
    setSyncStatus("Gesynchroniseerd ✓", "ok");
  }catch(e){
    setSyncStatus("Opslaan mislukt — probeert later opnieuw", "off");
  }finally{ busy = false; }
}

async function syncFromCloud(){
  if(!isSignedIn()) return;
  if(!navigator.onLine){ setSyncStatus("Offline — wijzigingen worden later gesynct", "off"); return; }
  if(busy) return;
  busy = true;
  try{
    setSyncStatus("Synchroniseren…", "sync");
    const srv = await getServer();

    if(srv.status === 404){
      // Nog geen cloud-bestand: maak het aan vanuit de lokale data.
      if(stateIsEmpty()) seedIfEmpty();
      const put = await putServer(getState(), null);
      saveBase(getState()); setETag(put.eTag); clearDirty();
      rerenderAll();
      setSyncStatus("Gesynchroniseerd ✓", "ok");
      return;
    }

    const server = normalizeState(srv.state || {});
    if(isDirty()){
      // Lokale, nog niet gesyncte wijzigingen -> 3-weg merge en terugschrijven.
      const base = loadBase() || server;
      const merged = normalizeState(mergeStates(base, server, getState()));
      setState(merged); saveLocal(); rerenderAll();
      const put = await putServer(merged, srv.eTag);
      saveBase(merged); setETag(put.eTag); clearDirty();
      setSyncStatus("Gesynchroniseerd ✓", "ok");
    }else{
      // Niets openstaand -> server overnemen.
      setState(server); saveLocal(); saveBase(server); setETag(srv.eTag); rerenderAll();
      setSyncStatus("Gesynchroniseerd ✓", "ok");
    }
  }catch(e){
    if(e && e.message === "not-signed-in") setSyncStatus("Meld je aan bij Microsoft om te synchroniseren", "off");
    else setSyncStatus("Synchronisatie mislukt — probeer 'Nu synchroniseren'", "off");
  }finally{ busy = false; }
}

/* ---------- Publieke acties (knoppen in Beheer) ---------- */
function msSignIn(){
  if(!MSAL_CONFIGURED){ alert("De Microsoft-koppeling is nog niet ingesteld.\nVul je client-ID en tenant-ID in msal-config.js in (zie SETUP.md)."); return; }
  if(!msalInstance){ alert("Microsoft-bibliotheek niet geladen."); return; }
  msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
}
function msSignOut(){
  if(!msalInstance) return;
  msalInstance.logoutRedirect();
}
function msSyncNow(){
  if(!isSignedIn()){ msSignIn(); return; }
  if(isDirty()) flushToCloud(); else syncFromCloud();
}

/* ---------- Initialisatie ---------- */
async function initSync(){
  // Geen MSAL beschikbaar/ingesteld -> app werkt puur lokaal (zoals het offline bestand).
  if(typeof msal === "undefined"){
    updateAuthUI(null);
    setSyncStatus("Microsoft-bibliotheek niet geladen — app werkt lokaal", "off");
    if(stateIsEmpty()) { seedIfEmpty(); rerenderAll(); }
    return;
  }
  if(!MSAL_CONFIGURED){
    updateAuthUI(null);
    setSyncStatus("Nog niet gekoppeld — werkt lokaal. Zie SETUP.md.", "off");
    if(stateIsEmpty()) { seedIfEmpty(); rerenderAll(); }
    return;
  }

  msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
  await msalInstance.initialize();
  try{
    const resp = await msalInstance.handleRedirectPromise();
    if(resp && resp.account) msalInstance.setActiveAccount(resp.account);
  }catch(e){ /* redirect-afhandeling mislukt; ga door als niet-aangemeld */ }

  let acct = msalInstance.getActiveAccount();
  if(!acct){
    const all = msalInstance.getAllAccounts();
    if(all.length){ acct = all[0]; msalInstance.setActiveAccount(acct); }
  }

  if(acct){
    updateAuthUI(acct);
    await syncFromCloud();
  }else{
    updateAuthUI(null);
    setSyncStatus("Niet aangemeld — meld je aan om te synchroniseren", "off");
    // Bewust NIET seeden: voorkomt dat voorbeelddata over je cloud-data wordt gemerged.
  }

  window.addEventListener("online", function(){
    setSyncStatus("Weer online — synchroniseren…", "sync");
    if(isDirty()) flushToCloud(); else syncFromCloud();
  });
  window.addEventListener("offline", function(){
    setSyncStatus("Offline — wijzigingen worden later gesynct", "off");
  });
}

/* Start zodra de app-globals geladen zijn (dit script staat ná het app-script). */
initSync();
