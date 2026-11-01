'use strict';
/* Zeiterfassung ProTax – Oberfläche, Anmeldung (Microsoft 365) und Speicher (SharePoint-Listen über Microsoft Graph) */
const CFG = window.ZE_CONFIG || {}, Z = window.ZC;
const APP_VERSION = 'ProTax Personal v.1.4';
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const WD = ['', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'];
const MONATE = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];
const ARTEN = ['Urlaub', 'Krank', 'Fortbildung', 'Sonderurlaub', 'Überstundenabbau'];
const today = () => Z.iso(new Date());
const nowMin = () => { const d = new Date(); return d.getHours() * 60 + d.getMinutes(); };
const nowHHMM = () => Z.hhmm(nowMin());
const de = iso => iso ? iso.slice(8, 10) + '.' + iso.slice(5, 7) + '.' + iso.slice(0, 4) : '—';
const deKurz = iso => iso ? iso.slice(8, 10) + '.' + iso.slice(5, 7) + '.' : '';
const lc = s => String(s || '').trim().toLowerCase();
function toast(msg, kind = '') { const el = document.createElement('div'); el.className = 't ' + kind; el.textContent = msg; $('#toast').appendChild(el); setTimeout(() => el.remove(), kind === 'err' ? 7000 : 3000); }
function download(name, text, mime) { const u = URL.createObjectURL(new Blob([text], { type: mime })); const a = document.createElement('a'); a.href = u; a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(u), 2000); }
const signed = m => `<span class="${m < 0 ? 'neg' : m > 0 ? 'pos' : ''}">${Z.dur(m, true)}</span>`;

/* ======================= Listen-Definition ======================= */
// key: [interner Spaltenname, Typ, indiziert, Anzeigename in SharePoint]
const LISTS = {
  buchungen: { name: 'ZE_Buchungen', title: b => `${b.name} ${b.datum}`, fields: {
    upn: ['ZeUpn', 'text', true, 'Mitarbeiter (E-Mail)'], name: ['ZeName', 'text', false, 'Name'], datum: ['ZeDatum', 'text', true, 'Datum'],
    beginn: ['ZeBeginn', 'text', false, 'Beginn'], ende: ['ZeEnde', 'text', false, 'Ende'], pausen: ['ZePausen', 'json', false, 'Pausen'],
    pauseMin: ['ZePauseMin', 'number', false, 'Pause pauschal (min)'], ort: ['ZeOrt', 'text', false, 'Ort'], bemerkung: ['ZeBemerkung', 'note', false, 'Bemerkung'],
    nachtrag: ['ZeNachtrag', 'bool', false, 'Nachtrag'], grund: ['ZeGrund', 'note', false, 'Begründung'], status: ['ZeStatus', 'text', false, 'Status'],
    korrVon: ['ZeKorrVon', 'text', false, 'Korrigiert von'], korrAm: ['ZeKorrAm', 'text', false, 'Korrigiert am'] } },
  abw: { name: 'ZE_Abwesenheiten', title: a => `${a.name} ${a.art} ${a.von}`, fields: {
    upn: ['ZeUpn', 'text', true, 'Mitarbeiter (E-Mail)'], name: ['ZeName', 'text', false, 'Name'], art: ['ZeArt', 'text', false, 'Art'],
    von: ['ZeVon', 'text', true, 'Von'], bis: ['ZeBis', 'text', false, 'Bis'], halbtag: ['ZeHalbtag', 'bool', false, 'Halber Tag'],
    bemerkung: ['ZeBemerkung', 'note', false, 'Bemerkung'], status: ['ZeStatus', 'text', false, 'Status'],
    entschVon: ['ZeEntschVon', 'text', false, 'Entschieden von'], entschAm: ['ZeEntschAm', 'text', false, 'Entschieden am'] } },
  ma: { name: 'ZE_Mitarbeiter', title: m => m.name, fields: {
    upn: ['ZeUpn', 'text', true, 'E-Mail (Anmeldename)'], name: ['ZeName', 'text', false, 'Name'], wochenstunden: ['ZeWochenstunden', 'number', false, 'Wochenstunden'],
    arbeitstage: ['ZeArbeitstage', 'text', false, 'Arbeitstage (1=Mo … 7=So)'], urlaubstage: ['ZeUrlaubstage', 'number', false, 'Urlaubstage pro Jahr'],
    restVorjahr: ['ZeRestVorjahr', 'number', false, 'Resturlaub Vorjahr'], saldoStart: ['ZeSaldoStart', 'text', false, 'Saldo ab Datum'],
    saldoStunden: ['ZeSaldoStunden', 'number', false, 'Saldo-Startwert (Stunden)'], eintritt: ['ZeEintritt', 'text', false, 'Eintritt'], austritt: ['ZeAustritt', 'text', false, 'Austritt'], leitung: ['ZeLeitung', 'bool', false, 'Leitung'],
    akteDrive: ['ZeAkteDrive', 'text', false, 'Personalakte Bibliothek'], akteId: ['ZeAkteId', 'text', false, 'Personalakte Ordner'], eingangId: ['ZeEingangId', 'text', false, 'Personalakte Eingang'] } }
};
const AKTE = { name: 'ZE_Personalakte', kategorien: ['Vertrag & Nachträge', 'Zeugnisse & Qualifikationen', 'Bescheinigungen', 'Unterweisungen', 'Sonstiges'],
  columns: [{ name: 'ZeKategorie', displayName: 'Kategorie', text: {} }, { name: 'ZeDokDatum', displayName: 'Dokumentdatum', text: {} }, { name: 'ZeBemerkung', displayName: 'Bemerkung', text: { allowMultipleLines: true } }] };
const MAX_UPLOAD_MB = 25;
function akteItem(it, ort) {
  const f = it.listItem?.fields || {};
  return { id: it.id, name: it.name, size: it.size || 0, folder: !!it.folder, ort, webUrl: it.webUrl || '', download: it['@microsoft.graph.downloadUrl'] || '',
    kategorie: f.ZeKategorie || '', dokDatum: f.ZeDokDatum || '', bemerkung: f.ZeBemerkung || '',
    by: it.createdBy?.user?.displayName || '', at: it.createdDateTime || '' };
}
function fromItem(key, it) {
  const f = it.fields || {}, o = { id: String(it.id) };
  for (const [k, [col, typ]] of Object.entries(LISTS[key].fields)) {
    const v = f[col];
    if (typ === 'json') { try { o[k] = v ? JSON.parse(v) : []; } catch (e) { o[k] = []; } }
    else if (typ === 'number') o[k] = v == null || v === '' ? null : Number(v);
    else if (typ === 'bool') o[k] = v === true || v === 'true' || v === 1;
    else o[k] = v == null ? '' : String(v);
  }
  o.upn = lc(o.upn);
  o._created = it.createdDateTime || ''; o._createdBy = it.createdBy?.user?.displayName || '';
  return o;
}
function toFields(key, obj, withTitle) {
  const def = LISTS[key], f = {};
  for (const [k, [col, typ]] of Object.entries(def.fields)) {
    if (!(k in obj)) continue;
    const v = obj[k];
    f[col] = typ === 'json' ? JSON.stringify(v || []) : typ === 'number' ? (v === '' || v == null ? null : Number(v)) : typ === 'bool' ? !!v : (v ?? '');
  }
  if (withTitle) f.Title = def.title(obj).slice(0, 255);
  return f;
}
LISTS.pd = { name: 'ZE_Personendaten', optional: true, title: p => p.name || p.upn, fields: {
  upn: ['ZeUpn', 'text', true, 'E-Mail (Anmeldename)'], name: ['ZeName', 'text', false, 'Name'], anrede: ['ZeAnrede', 'text', false, 'Anrede'], titel: ['ZeTitel', 'text', false, 'Titel'],
  vorname: ['ZeVorname', 'text', false, 'Vorname'], nachname: ['ZeNachname', 'text', false, 'Nachname'], geburtsdatum: ['ZeGeburtsdatum', 'text', false, 'Geburtsdatum'],
  strasse: ['ZeStrasse', 'text', false, 'Straße'], hausnr: ['ZeHausnr', 'text', false, 'Hausnummer'], plz: ['ZePlz', 'text', false, 'PLZ'], bundesland: ['ZeBundesland', 'text', false, 'Bundesland'], adresszusatz: ['ZeAdresszusatz', 'text', false, 'Adresszusatz'], staatsang: ['ZeStaatsang', 'text', false, 'Staatsangehörigkeit'], ort: ['ZeOrt', 'text', false, 'Ort'], land: ['ZeLand', 'text', false, 'Land'],
  privMail: ['ZePrivMail', 'text', false, 'Private E-Mail'], mobil: ['ZeMobil', 'text', false, 'Mobil privat'], festnetz: ['ZeFestnetz', 'text', false, 'Festnetz privat'],
  notfallName: ['ZeNotfallName', 'text', false, 'Notfallkontakt Name'], notfallBez: ['ZeNotfallBez', 'text', false, 'Notfallkontakt Beziehung'], notfallTel: ['ZeNotfallTel', 'text', false, 'Notfallkontakt Telefon'],
  geaendertAm: ['ZeGeaendertAm', 'text', false, 'Zuletzt geändert'] } };
function columnDefs(key, withIndex) {
  return Object.values(LISTS[key].fields).map(([name, typ, idx, label]) => {
    const c = { name, displayName: label };
    if (typ === 'text') c.text = {};
    else if (typ === 'note' || typ === 'json') c.text = { allowMultipleLines: true, linesForEditing: 3 };
    else if (typ === 'number') c.number = {};
    else if (typ === 'bool') c.boolean = {};
    if (withIndex && idx) c.indexed = true;
    return c;
  });
}

/* ======================= Speicher: Microsoft 365 ======================= */
class GraphStore {
  constructor() {
    this.scopes = ['User.Read', 'Sites.ReadWrite.All'];
    this.site = null; this.lists = {};
  }
  async login() {
    if (!window.msal) throw new Error('Anmeldebibliothek (lib/msal-browser.min.js) wurde nicht geladen.');
    if (/HIER-/.test(CFG.clientId + CFG.tenantId)) throw new Error('config.js ist noch nicht ausgefüllt (tenantId / clientId).');
    this.pca = new msal.PublicClientApplication({
      auth: { clientId: CFG.clientId, authority: 'https://login.microsoftonline.com/' + CFG.tenantId, redirectUri: CFG.redirectUri || (location.origin + location.pathname.replace(/index\.html$/, '')) },
      cache: { cacheLocation: 'localStorage' }
    });
    await this.pca.initialize();
    const r = await this.pca.handleRedirectPromise();
    this.account = r?.account || this.pca.getAllAccounts()[0] || null;
    if (!this.account) return null;
    this.pca.setActiveAccount(this.account);
    return { upn: lc(this.account.username), name: this.account.name || this.account.username };
  }
  loginRedirect() { return this.pca.loginRedirect({ scopes: this.scopes, prompt: 'select_account' }); }
  logout() { return this.pca.logoutRedirect({ account: this.account }); }
  async token(scopes) {
    try { return (await this.pca.acquireTokenSilent({ scopes, account: this.account })).accessToken; }
    catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) { await this.pca.acquireTokenRedirect({ scopes, account: this.account }); return new Promise(() => {}); }
      throw e;
    }
  }
  async g(method, path, body, scopes) {
    const tok = await this.token(scopes || this.scopes);
    const url = path.startsWith('https://') ? path : 'https://graph.microsoft.com/v1.0' + path;
    const r = await fetch(url, { method, headers: { Authorization: 'Bearer ' + tok, 'Content-Type': 'application/json', Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' }, body: body ? JSON.stringify(body) : undefined });
    if (!r.ok) {
      let msg = 'HTTP ' + r.status; try { const j = await r.json(); msg = j.error?.message || msg; } catch (e) {}
      if (r.status === 403) msg = 'Keine Berechtigung (' + msg + ')';
      const err = new Error(msg); err.status = r.status; throw err;
    }
    return r.status === 204 ? null : r.json();
  }
  async init() {
    const sp = CFG.sharepoint || {};
    this.site = await this.g('GET', `/sites/${sp.hostname}:${sp.sitePath}?$select=id,webUrl,displayName`);
    const ls = await this.g('GET', `/sites/${this.site.id}/lists?$select=id,displayName,name&$top=200`);
    this.lists = {};
    for (const [k, d] of Object.entries(LISTS)) { const f = ls.value.find(x => x.displayName === d.name || x.name === d.name); if (f) this.lists[k] = f.id; }
    const ak = ls.value.find(x => x.displayName === AKTE.name || x.name === AKTE.name); this.akteList = ak ? ak.id : null;
    return { site: this.site.webUrl, missing: Object.keys(LISTS).filter(k => !LISTS[k].optional && !this.lists[k]).map(k => LISTS[k].name) };
  }
  async setup() {
    const scopes = ['Sites.Manage.All'];
    for (const [k, d] of Object.entries(LISTS)) {
      if (this.lists[k] || d.optional) continue;
      const body = c => ({ displayName: d.name, columns: columnDefs(k, c), list: { template: 'genericList' } });
      let res;
      try { res = await this.g('POST', `/sites/${this.site.id}/lists`, body(true), scopes); }
      catch (e) { res = await this.g('POST', `/sites/${this.site.id}/lists`, body(false), scopes); }
      this.lists[k] = res.id;
    }
  }
  async load(key, filter) {
    let url = `/sites/${this.site.id}/lists/${this.lists[key]}/items?$expand=fields&$top=999` + (filter ? '&$filter=' + encodeURIComponent(filter) : '');
    const out = [];
    while (url) { const j = await this.g('GET', url); out.push(...j.value.map(it => fromItem(key, it))); url = j['@odata.nextLink'] || null; }
    return out;
  }
  loadBuchungen(from) { return this.load('buchungen', `fields/ZeDatum ge '${from}'`); }
  async create(key, obj) {
    const it = await this.g('POST', `/sites/${this.site.id}/lists/${this.lists[key]}/items`, { fields: toFields(key, obj, true) });
    return fromItem(key, it);
  }
  async update(key, id, partial, full) {
    const f = toFields(key, partial, !!full); if (full) f.Title = LISTS[key].title(full).slice(0, 255);
    await this.g('PATCH', `/sites/${this.site.id}/lists/${this.lists[key]}/items/${id}/fields`, f);
  }
  // Leitung legt Eintrag für eine/n Mitarbeiter/in an: "Erstellt von" auf diese Person setzen,
  // damit sie den Eintrag trotz Elementberechtigung ("nur eigene Elemente") sieht und bearbeiten kann.
  async assignAuthor(key, id, upn) {
    const sp = CFG.sharepoint || {};
    const tok = await this.token([`https://${sp.hostname}/AllSites.Write`]);
    const url = `https://${sp.hostname}${sp.sitePath}/_api/web/lists(guid'${this.lists[key]}')/items(${id})/ValidateUpdateListItem`;
    const r = await fetch(url, { method: 'POST', headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
      body: JSON.stringify({ formValues: [{ FieldName: 'Author', FieldValue: JSON.stringify([{ Key: 'i:0#.f|membership|' + upn }]) }], bNewDocumentUpdate: true }) });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const j = await r.json(); const bad = (j.value || []).find(x => x.HasException);
    if (bad) throw new Error(bad.ErrorMessage || 'Feld „Erstellt von“ nicht gesetzt');
  }
  async versions(key, id) {
    const base = `/sites/${this.site.id}/lists/${this.lists[key]}/items/${id}/versions`;
    let j; try { j = await this.g('GET', base + '?$expand=fields'); } catch (e) { j = await this.g('GET', base); }
    return (j.value || []).map(v => ({ at: v.lastModifiedDateTime, by: v.lastModifiedBy?.user?.displayName || '', data: v.fields ? fromItem(key, { id, fields: v.fields }) : null }));
  }
}

/* ======================= Speicher: Demo (nur dieser Browser) ======================= */
/* ---- Persönliche Daten (Liste mit "nur eigene Elemente") ---- */
GraphStore.prototype.pdStatus = async function () {
  if (!this.lists.pd) return { status: 'missing' };
  try {
    const sp = CFG.sharepoint || {}, tok = await this.token([`https://${sp.hostname}/AllSites.Write`]);
    const r = await fetch(`https://${sp.hostname}${sp.sitePath}/_api/web/lists(guid'${this.lists.pd}')?$select=ReadSecurity,WriteSecurity`, { headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=nometadata' } });
    if (!r.ok) return { status: 'ok', unchecked: true };
    const j = await r.json();
    return { status: j.ReadSecurity === 2 && j.WriteSecurity === 2 ? 'ok' : 'unsafe' };
  } catch (e) { return { status: 'ok', unchecked: true }; }
};
GraphStore.prototype.pdSetup = async function () {
  const res = await this.g('POST', `/sites/${this.site.id}/lists`, { displayName: LISTS.pd.name, columns: columnDefs('pd', true), list: { template: 'genericList' } }, ['Sites.Manage.All'])
    .catch(() => this.g('POST', `/sites/${this.site.id}/lists`, { displayName: LISTS.pd.name, columns: columnDefs('pd', false), list: { template: 'genericList' } }, ['Sites.Manage.All']));
  this.lists.pd = res.id;
};
GraphStore.prototype.pdFind = async function () {
  const ls = await this.g('GET', `/sites/${this.site.id}/lists?$select=id,displayName,name&$top=200`);
  const f = ls.value.find(x => x.displayName === LISTS.pd.name || x.name === LISTS.pd.name); if (f) this.lists.pd = f.id;
};

/* ---- Personalakte (SharePoint-Dokumentbibliothek) ---- */
GraphStore.prototype.raw = async function (method, path, body, contentType) {
  const tok = await this.token(this.scopes);
  const r = await fetch('https://graph.microsoft.com/v1.0' + path, { method, headers: { Authorization: 'Bearer ' + tok, 'Content-Type': contentType || 'application/octet-stream' }, body });
  if (!r.ok) { let msg = 'HTTP ' + r.status; try { const j = await r.json(); msg = j.error?.message || msg; } catch (e) {} const err = new Error(msg); err.status = r.status; throw err; }
  return r.status === 204 ? null : r.json();
};
// Fehlende Spalten in bestehenden Listen ergänzen (z. B. nach einem Update der App)
GraphStore.prototype.ensureColumns = async function (key) {
  const have = new Set(((await this.g('GET', `/sites/${this.site.id}/lists/${this.lists[key]}/columns?$select=name&$top=200`)).value || []).map(c => c.name));
  for (const c of columnDefs(key, false)) if (!have.has(c.name)) await this.g('POST', `/sites/${this.site.id}/lists/${this.lists[key]}/columns`, c, ['Sites.Manage.All']);
};
GraphStore.prototype.akteStatus = async function () {
  if (!this.akteList) return { status: 'missing' };
  if (!this.akteDrive) this.akteDrive = (await this.g('GET', `/sites/${this.site.id}/lists/${this.akteList}/drive?$select=id`)).id;
  // Sicherheitsprüfung: Die Bibliothek darf nicht für die Gruppen Mitglieder/Besucher freigegeben sein
  let unsafe = false;
  try {
    const p = await this.g('GET', `/drives/${this.akteDrive}/root/permissions`);
    unsafe = (p.value || []).some(x => /mitglieder|besucher|members|visitors/i.test(x.grantedToV2?.siteGroup?.displayName || x.grantedTo?.user?.displayName || ''));
  } catch (e) { /* Prüfung nicht möglich – kein Abbruch */ }
  return { status: unsafe ? 'unsafe' : 'ok', drive: this.akteDrive };
};
GraphStore.prototype.akteSetup = async function () {
  const res = await this.g('POST', `/sites/${this.site.id}/lists`, { displayName: AKTE.name, columns: AKTE.columns, list: { template: 'documentLibrary' } }, ['Sites.Manage.All']);
  this.akteList = res.id; this.akteDrive = null;
};
GraphStore.prototype.akteEnsureFolder = async function (ma) {
  await this.ensureColumns('ma');
  const d = this.akteDrive, mk = async (parent, name) => {
    try { return await this.g('POST', `/drives/${d}/items/${parent}/children`, { name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }); }
    catch (e) { if (e.status === 409) return this.g('GET', `/drives/${d}/items/${parent}:/${encodeURIComponent(name)}`); throw e; }
  };
  const root = await this.g('GET', `/drives/${d}/root?$select=id`);
  const f = await mk(root.id, ma.upn), ein = await mk(f.id, 'Eingang');
  const invite = (id, role) => this.g('POST', `/drives/${d}/items/${id}/invite`, { requireSignIn: true, sendInvitation: false, roles: [role], recipients: [{ email: ma.upn }] });
  await invite(f.id, 'read'); await invite(ein.id, 'write');
  const patch = { akteDrive: d, akteId: f.id, eingangId: ein.id };
  await this.update('ma', ma.id, patch);
  return patch;
};
GraphStore.prototype.akteList_ = async function (ma) {
  const q = '?$expand=listItem($expand=fields)&$top=500';
  const [a, e] = await Promise.all([this.g('GET', `/drives/${ma.akteDrive}/items/${ma.akteId}/children${q}`), this.g('GET', `/drives/${ma.akteDrive}/items/${ma.eingangId}/children${q}`)]);
  return [...a.value.filter(x => !x.folder).map(x => akteItem(x, 'akte')), ...e.value.filter(x => !x.folder).map(x => akteItem(x, 'eingang'))];
};
GraphStore.prototype.akteUpload = async function (ma, ort, file, meta) {
  const parent = ort === 'eingang' ? ma.eingangId : ma.akteId;
  const it = await this.raw('PUT', `/drives/${ma.akteDrive}/items/${parent}:/${encodeURIComponent(file.name)}:/content?@microsoft.graph.conflictBehavior=rename`, file, file.type || 'application/octet-stream');
  await this.g('PATCH', `/drives/${ma.akteDrive}/items/${it.id}/listItem/fields`, { ZeKategorie: meta.kategorie, ZeDokDatum: meta.dokDatum, ZeBemerkung: meta.bemerkung });
};
GraphStore.prototype.akteMove = async function (ma, item, kategorie) {
  await this.g('PATCH', `/drives/${ma.akteDrive}/items/${item.id}`, { parentReference: { id: ma.akteId } });
  await this.g('PATCH', `/drives/${ma.akteDrive}/items/${item.id}/listItem/fields`, { ZeKategorie: kategorie });
};
GraphStore.prototype.akteSetKategorie = async function (ma, item, kategorie) {
  await this.g('PATCH', `/drives/${ma.akteDrive}/items/${item.id}/listItem/fields`, { ZeKategorie: kategorie });
};
GraphStore.prototype.akteDelete = async function (ma, item) { await this.g('DELETE', `/drives/${ma.akteDrive}/items/${item.id}`); };
GraphStore.prototype.akteOpen = async function (ma, item) {
  const it = await this.g('GET', `/drives/${ma.akteDrive}/items/${item.id}?$select=id,webUrl,@microsoft.graph.downloadUrl`);
  return { view: it.webUrl, download: it['@microsoft.graph.downloadUrl'] };
};

class DemoStore {
  constructor() {
    this.k = 'ze_demo_v1';
    try { this.d = JSON.parse(localStorage.getItem(this.k)); } catch (e) { this.d = null; }
    if (!this.d) { this.d = DemoStore.seed(); this.save(); }
  }
  save() { try { localStorage.setItem(this.k, JSON.stringify(this.d)); } catch (e) {} }
  static seed() {
    let s = 7; const rnd = () => (s = (s * 9301 + 49297) % 233280) / 233280;
    const ri = (a, b) => a + Math.floor(rnd() * (b - a + 1));
    const t = today(), start = Z.addDays(t, -45);
    const ma = [
      { id: 'm1', leitung: true, upn: 'anna.leitung@demo.de', name: 'Anna Leitung', wochenstunden: 40, arbeitstage: '1,2,3,4,5', urlaubstage: 30, restVorjahr: 2, saldoStart: start, saldoStunden: 0, eintritt: '2020-01-01', austritt: '' },
      { id: 'm2', upn: 'ben.muster@demo.de', name: 'Ben Muster', wochenstunden: 40, arbeitstage: '1,2,3,4,5', urlaubstage: 30, restVorjahr: 0, saldoStart: start, saldoStunden: 4.5, eintritt: '2022-04-01', austritt: '' },
      { id: 'm3', upn: 'clara.teilzeit@demo.de', name: 'Clara Teilzeit', wochenstunden: 25, arbeitstage: '1,2,3,4', urlaubstage: 24, restVorjahr: 1, saldoStart: start, saldoStunden: -1, eintritt: '2023-09-01', austritt: '' }
    ];
    const b = [], abw = []; let id = 1;
    const add = (m, datum, beginn, ende, pausen, extra) => b.push(Object.assign({ id: 'b' + id++, upn: m.upn, name: m.name, datum, beginn, ende, pausen: pausen || [], pauseMin: 0, ort: 'Büro', bemerkung: '', nachtrag: false, grund: '', status: 'aktiv', korrVon: '', korrAm: '' }, extra || {}));
    abw.push({ id: 'a1', upn: ma[1].upn, name: ma[1].name, art: 'Urlaub', von: Z.addDays(t, -12), bis: Z.addDays(t, -11), halbtag: false, bemerkung: '', status: 'genehmigt', entschVon: 'Anna Leitung', entschAm: '' });
    abw.push({ id: 'a2', upn: ma[0].upn, name: ma[0].name, art: 'Krank', von: Z.addDays(t, -20), bis: Z.addDays(t, -20), halbtag: false, bemerkung: '', status: 'gemeldet', entschVon: '', entschAm: '' });
    abw.push({ id: 'a3', upn: ma[2].upn, name: ma[2].name, art: 'Urlaub', von: Z.addDays(t, 20), bis: Z.addDays(t, 24), halbtag: false, bemerkung: 'Familienfeier', status: 'beantragt', entschVon: '', entschAm: '' });
    for (let d = start; d < t; d = Z.addDays(d, 1)) {
      ma.forEach((m, i) => {
        if (Z.sollTag(m, d, CFG) === 0) return;
        if (Z.abwAm(abw, m.upn, d).length) return;
        const soll = Z.sollTag(m, d, CFG), bg = 7 * 60 + 30 + ri(0, 50), p0 = 12 * 60 + ri(-15, 20);
        const ende = bg + soll + 30 + ri(-25, 35);
        add(m, d, Z.hhmm(bg), Z.hhmm(ende), soll > 360 ? [{ v: Z.hhmm(p0), b: Z.hhmm(p0 + 30) }] : [], { ort: rnd() < 0.3 ? 'Homeoffice' : 'Büro' });
      });
    }
    // Auffälligkeiten für die Vorführung
    const d1 = Z.addDays(t, -6), d2 = Z.addDays(t, -4), d3 = Z.addDays(t, -3);
    const fix = (upn, d, patch) => { const x = b.find(y => y.upn === upn && y.datum === d); if (x) Object.assign(x, patch); };
    fix(ma[1].upn, d1, { beginn: '07:30', ende: '18:30', pausen: [{ v: '12:30', b: '13:00' }] });
    fix(ma[2].upn, d2, { beginn: '08:00', ende: '15:00', pausen: [] });
    fix(ma[1].upn, d3, { nachtrag: true, grund: 'Stempeln vergessen' });
    const m2 = ma[1]; if (Z.sollTag(m2, t, CFG) > 0) add(m2, t, Z.hhmm(Math.max(0, Math.min(nowMin() - 30, 7 * 60 + 58))), '', [], {});
    return { ma, buchungen: b, abw, hist: {}, n: 1000 };
  }
  users() { return this.d.ma.map(m => ({ upn: m.upn, name: m.name })); }
  async login() {
    let u = null; try { u = localStorage.getItem('ze_demo_user'); } catch (e) {}
    const m = this.d.ma.find(x => x.upn === u) || this.d.ma[0];
    return { upn: m.upn, name: m.name };
  }
  isDemoLeitung(upn) { return upn === 'anna.leitung@demo.de'; }
  async init() { return { site: 'Demo (nur dieser Browser)', missing: [] }; }
  async load(key) { return JSON.parse(JSON.stringify(this.d[key] || [])); }
  async loadBuchungen(from) { return JSON.parse(JSON.stringify(this.d.buchungen.filter(b => b.datum >= from))); }
  hist(key, obj) { (this.d.hist[key + obj.id] = this.d.hist[key + obj.id] || []).push({ at: new Date().toISOString(), by: S.me.name, data: JSON.parse(JSON.stringify(obj)) }); }
  async create(key, obj) { this.d[key] = this.d[key] || []; const o = Object.assign({}, obj, { id: 'x' + (this.d.n++), _created: new Date().toISOString() }); this.d[key].push(o); this.hist(key, o); this.save(); return JSON.parse(JSON.stringify(o)); }
  async update(key, id, partial) { const o = this.d[key].find(x => x.id === id); if (!o) throw new Error('Eintrag nicht gefunden'); Object.assign(o, partial); this.hist(key, o); this.save(); }
  async versions(key, id) { return (this.d.hist[key + id] || []).slice().reverse(); }
  async setup() {}
  reset() { localStorage.removeItem(this.k); }
  async pdStatus() { this.d.pd = this.d.pd || []; return { status: this.d.pdLib ? 'ok' : 'missing' }; }
  async pdSetup() { this.d.pdLib = true; this.d.pd = this.d.pd || []; this.save(); }
  async pdFind() {}
  _ak() { return (this.d.akte = this.d.akte || { lib: false, files: [], n: 1 }); }
  async ensureColumns() {}
  async akteStatus() { return { status: this._ak().lib ? 'ok' : 'missing', drive: 'demo' }; }
  async akteSetup() { this._ak().lib = true; this.save(); }
  async akteEnsureFolder(ma) { const p = { akteDrive: 'demo', akteId: 'f-' + ma.id, eingangId: 'e-' + ma.id }; await this.update('ma', ma.id, p); return p; }
  async akteList_(ma) { return this._ak().files.filter(f => f.upn === ma.upn).map(f => ({ ...f, webUrl: f.data, download: f.data })); }
  async akteUpload(ma, ort, file, meta) {
    if (file.size > 1.5 * 1024 * 1024) throw new Error('Im Demo-Modus nur Dateien bis 1,5 MB (Browser-Speicher). Im Echtbetrieb bis ' + MAX_UPLOAD_MB + ' MB.');
    const data = await new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(file); });
    const ak = this._ak(); ak.files.push({ id: 'd' + (ak.n++), upn: ma.upn, ort, name: file.name, size: file.size, data, kategorie: meta.kategorie, dokDatum: meta.dokDatum, bemerkung: meta.bemerkung, by: S.me.name, at: new Date().toISOString() });
    this.save();
  }
  async akteMove(ma, item, kategorie) { const f = this._ak().files.find(x => x.id === item.id); f.ort = 'akte'; f.kategorie = kategorie; this.save(); }
  async akteSetKategorie(ma, item, kategorie) { this._ak().files.find(x => x.id === item.id).kategorie = kategorie; this.save(); }
  async akteDelete(ma, item) { const ak = this._ak(); ak.files = ak.files.filter(x => x.id !== item.id); this.save(); }
  async akteOpen(ma, item) { const f = this._ak().files.find(x => x.id === item.id); return { view: f.data, download: f.data }; }
}

/* ======================= Zustand ======================= */
const S = { store: null, me: null, leitung: false, ma: [], buchungen: [], abw: [], loadedFrom: null, view: 'heute', month: today().slice(0, 7), viewUpn: null, info: null, missing: [], ort: (CFG.orte || ['Büro'])[0], busy: false, akte: null, pd: null };
const myMa = () => S.ma.find(m => m.upn === S.me.upn);
const maBy = upn => S.ma.find(m => m.upn === upn);
const dataFor = upn => ({ buchungen: S.buchungen.filter(b => b.upn === upn), abw: S.abw });
// Leitung: in den Stammdaten markiert, in config.js eingetragen, oder Ersteinrichtung (noch keine Listen / keine Stammdaten).
// Die echte Zugriffsgrenze setzt SharePoint: Mitarbeitende können Stammdaten nur lesen und Listen nicht anlegen.
const isLeitung = upn => (CFG.leitung || []).map(lc).includes(lc(upn)) || !!(S.ma.find(m => m.upn === lc(upn)) || {}).leitung
  || (S.store && S.store.isDemoLeitung && S.store.isDemoLeitung(upn)) || (!CFG.demo && (S.missing.length > 0 || S.ma.length === 0));
function canEdit(upn, datum) {
  if (S.leitung) return true;
  if (upn !== S.me.upn) return false;
  return Z.daysBetween(datum, today()) <= (+CFG.korrekturFristTage || 0) && datum <= today();
}

async function loadAll() {
  S.ma = (await S.store.load('ma')).sort((a, b) => a.name.localeCompare(b.name));
  const y = +today().slice(0, 4);
  let from = (y - 1) + '-01-01';
  S.ma.forEach(m => { const s = m.saldoStart || m.eintritt; if (s && s < from) from = s; });
  S.loadedFrom = from;
  const [b, a] = await Promise.all([S.store.loadBuchungen(from), S.store.load('abw')]);
  S.buchungen = b; S.abw = a;
  if (!S.leitung) { S.buchungen = S.buchungen.filter(x => x.upn === S.me.upn); S.abw = S.abw.filter(x => x.upn === S.me.upn); }
}
async function ensureLoaded(from) {
  if (from >= S.loadedFrom) return;
  const more = await S.store.loadBuchungen(from);
  const known = new Set(S.buchungen.map(b => b.id));
  S.buchungen.push(...more.filter(b => !known.has(b.id) && (S.leitung || b.upn === S.me.upn)));
  S.loadedFrom = from;
}
async function refreshRecent() {
  const rec = await S.store.loadBuchungen(Z.addDays(today(), -3));
  const ids = new Set(rec.map(b => b.id));
  S.buchungen = S.buchungen.filter(b => !ids.has(b.id)).concat(rec.filter(b => S.leitung || b.upn === S.me.upn));
}
async function act(fn, okMsg) {
  if (S.busy) return; S.busy = true;
  try { await fn(); if (okMsg) toast(okMsg, 'ok'); }
  catch (e) { console.error(e); toast('Fehler: ' + (e.message || e), 'err'); }
  finally { S.busy = false; render(); }
}
async function fuerAndere(key, o) {
  if (o.upn === S.me.upn || !S.store.assignAuthor) return;
  try { await S.store.assignAuthor(key, o.id, o.upn); }
  catch (e) { console.warn(e); toast('Gespeichert, aber „Erstellt von“ konnte nicht auf ' + o.name + ' gesetzt werden – die Person sieht den Eintrag evtl. nicht (' + e.message + '). Siehe Anleitung, Fehlerbehebung.', 'err'); }
}
function upsertLocal(arr, obj) { const i = arr.findIndex(x => x.id === obj.id); if (i >= 0) arr[i] = obj; else arr.push(obj); }

/* ======================= Stempeln ======================= */
function offeneBloecke(upn) {
  const t = today();
  return S.buchungen.filter(b => b.upn === upn && b.status !== 'storniert' && !b.ende && b.datum <= t).sort((a, c) => (a.datum + a.beginn).localeCompare(c.datum + c.beginn));
}
// aktuell laufender Block: von heute, oder von gestern wenn es noch vor 6 Uhr ist (Arbeit über Mitternacht)
function laufenderBlock(upn) {
  const t = today(), g = Z.addDays(t, -1);
  return offeneBloecke(upn).filter(b => b.datum === t || (b.datum === g && nowMin() < 360)).pop() || null;
}
function stempelStatus(upn) {
  const b = laufenderBlock(upn);
  if (!b) return { s: 'aus' };
  const p = (b.pausen || [])[b.pausen.length - 1];
  return { s: p && !p.b ? 'pause' : 'an', b, seit: p && !p.b ? p.v : b.beginn };
}
const kommen = () => act(async () => {
  await refreshRecent();
  if (laufenderBlock(S.me.upn)) throw new Error('Du bist bereits eingestempelt.');
  const m = myMa();
  const o = await S.store.create('buchungen', { upn: S.me.upn, name: m ? m.name : S.me.name, datum: today(), beginn: nowHHMM(), ende: '', pausen: [], pauseMin: 0, ort: S.ort, bemerkung: '', nachtrag: false, grund: '', status: 'aktiv', korrVon: '', korrAm: '' });
  upsertLocal(S.buchungen, o);
}, 'Eingestempelt ' + nowHHMM());
const pause = () => act(async () => {
  await refreshRecent();
  const b = laufenderBlock(S.me.upn); if (!b) throw new Error('Nicht eingestempelt.');
  const ps = (b.pausen || []).slice(); if (ps.length && !ps[ps.length - 1].b) throw new Error('Pause läuft bereits.');
  ps.push({ v: nowHHMM() }); await S.store.update('buchungen', b.id, { pausen: ps }); b.pausen = ps;
}, 'Pause gestartet');
const pauseEnde = () => act(async () => {
  await refreshRecent();
  const b = laufenderBlock(S.me.upn); if (!b) throw new Error('Nicht eingestempelt.');
  const ps = (b.pausen || []).slice(); const l = ps[ps.length - 1]; if (!l || l.b) throw new Error('Keine laufende Pause.');
  l.b = nowHHMM(); await S.store.update('buchungen', b.id, { pausen: ps }); b.pausen = ps;
}, 'Pause beendet');
const gehen = () => act(async () => {
  await refreshRecent();
  const b = laufenderBlock(S.me.upn); if (!b) throw new Error('Nicht eingestempelt.');
  const ps = (b.pausen || []).slice(); const l = ps[ps.length - 1]; if (l && !l.b) l.b = nowHHMM();
  const ende = nowHHMM();
  await S.store.update('buchungen', b.id, { pausen: ps, ende }); b.pausen = ps; b.ende = ende;
}, 'Ausgestempelt ' + nowHHMM());

/* ======================= Darstellung ======================= */
const ICON = {
  heute: '<path d="M12 7v5l3 2"/><circle cx="12" cy="12" r="8.5"/>',
  monat: '<rect x="4" y="5" width="16" height="15" rx="2"/><path d="M4 10h16M9 3v4M15 3v4M8 14h3M8 17h6"/>',
  abw: '<path d="M4 19h16M7 19c0-5 2.5-9 5-9s5 4 5 9"/><path d="M12 10V5M9.5 6.5 12 5l2.5 1.5"/>',
  team: '<circle cx="9" cy="9" r="3"/><path d="M3.5 19c.6-3 2.8-5 5.5-5s4.9 2 5.5 5"/><circle cx="17" cy="10" r="2.3"/><path d="M15.8 14.3c2.4.2 4 1.9 4.6 4.7"/>',
  daten: '<circle cx="12" cy="8.5" r="3.5"/><path d="M5 19.5c.9-3.4 3.7-5.5 7-5.5s6.1 2.1 7 5.5"/>',
  akte: '<path d="M3.5 7.5a2 2 0 0 1 2-2h4l2 2h7a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2h-13a2 2 0 0 1-2-2z"/><path d="M8 13h8"/>',
  upload: '<path d="M12 16V5M7.5 9.5 12 5l4.5 4.5"/><path d="M5 19h14"/>',
  stamm: '<rect x="3.5" y="5" width="17" height="14" rx="2"/><circle cx="9" cy="11" r="2.2"/><path d="M5.8 16.5c.5-1.8 1.7-2.7 3.2-2.7s2.7.9 3.2 2.7M14.5 10h3.5M14.5 13.5h3.5"/>',
  refresh: '<path d="M19 12a7 7 0 1 1-2.05-4.95"/><path d="M19 5v4h-4"/>',
  logout: '<path d="M14 5h4a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2h-4"/><path d="M10 16l-4-4 4-4M6 12h9"/>',
  play: '<path d="M8 5.5v13l10.5-6.5z" fill="currentColor" stroke="none"/>',
  stop: '<rect x="7" y="7" width="10" height="10" rx="1.5" fill="currentColor" stroke="none"/>',
  pause: '<path d="M9 6v12M15 6v12"/>'
};
const svg = k => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON[k]}</svg>`;
const initials = n => String(n || '?').split(/\s+/).filter(Boolean).map(x => x[0]).slice(0, 2).join('').toUpperCase();
const VIEW_LABEL = { heute: 'Heute', monat: 'Monatsnachweis', abw: 'Abwesenheiten', daten: 'Persönliche Daten', akte: 'Personalakte', team: 'Team', stamm: 'Stammdaten' };
function renderSide() {
  const offen = S.abw.filter(a => a.status === 'beantragt').length;
  const items = [['heute'], ['monat'], ['abw'], ['daten'], ['akte']];
  if (S.leitung) items.push(['team', offen], ['stamm']);
  const m = myMa();
  $('#side').innerHTML = `<div class="user"><div class="avatar">${esc(initials(m ? m.name : S.me.name))}</div><div><div class="n">${esc(m ? m.name : S.me.name)}</div><div class="r">${S.leitung ? 'Leitung' : 'Mitarbeiter/in'}</div></div></div>`
    + items.map(([v, n]) => `<div class="nav ${S.view === v ? 'on' : ''}" role="link" tabindex="0" onclick="go('${v}')" onkeydown="if(event.key==='Enter')go('${v}')">${svg(v)}<span>${VIEW_LABEL[v]}</span>${n ? `<span class="cnt">${n}</span>` : ''}</div>`).join('')
    + `<div class="ver">${esc(APP_VERSION)}</div>`;
  const demo = CFG.demo ? `<select class="inp" style="width:auto" onchange="demoUser(this.value)" title="Demo: Benutzer wechseln">${S.store.users().map(u => `<option value="${esc(u.upn)}" ${u.upn === S.me.upn ? 'selected' : ''}>${esc(u.name)}</option>`).join('')}</select><span class="badge b-warn">Demo</span>
    <button class="btn sm" onclick="demoReset()">Demo zurücksetzen</button>` : '';
  $('#who').innerHTML = `${demo}
    <button class="iconbtn no-print" onclick="reload()" title="Daten neu laden" aria-label="Daten neu laden">${svg('refresh')}</button>
    ${CFG.demo ? '' : `<button class="iconbtn no-print" onclick="S.store.logout()" title="Abmelden" aria-label="Abmelden">${svg('logout')}</button>`}`;
  $('#appname').textContent = CFG.appName || 'Personal';
  document.title = CFG.appName || 'Personal';
}
// Schwebender Timer unten rechts: laufende Sitzung + Tagessumme, Start/Stopp mit einem Klick
function renderTimer() {
  const el = $('#timer'); if (!el) return;
  const m = S.me && myMa();
  if (!m || S.missing.length) { el.hidden = true; return; }
  const st = stempelStatus(S.me.upn), t = today();
  const row = Z.tagesZeile(m, t, dataFor(S.me.upn), CFG, t, nowMin());
  let lauf = '00:00:00';
  if (st.s !== 'aus') {
    const [h, mi] = st.seit.split(':').map(Number), now = new Date();
    let sec = (now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds()) - (h * 3600 + mi * 60);
    if (sec < 0) sec += 86400;
    lauf = [Math.floor(sec / 3600), Math.floor(sec / 60) % 60, sec % 60].map(Z.pad).join(':');
  }
  const btn = st.s === 'aus' ? `<button class="tb" onclick="kommen()" aria-label="Kommen" title="Kommen">${svg('play')}</button>`
    : st.s === 'pause' ? `<button class="tb" onclick="pauseEnde()" aria-label="Pause beenden" title="Pause beenden">${svg('play')}</button>`
    : `<button class="tb run" onclick="gehen()" aria-label="Gehen" title="Gehen">${svg('stop')}</button>`;
  const label = st.s === 'aus' ? 'Nicht eingestempelt' : st.s === 'pause' ? 'Pause' : 'Anwesend';
  el.innerHTML = `${btn}<div onclick="go('heute')" role="link" tabindex="0"><div class="tt">${lauf}</div><div class="ts">${label} · Heute ${Z.dur(row.ist)} h</div></div>`;
  el.hidden = false;
}
function render() {
  if (!S.me) return;
  renderSide();
  if (S.missing.length && S.view !== 'stamm') { $('#main').innerHTML = viewMissing(); renderTimer(); return; }
  const v = { heute: viewHeute, monat: viewMonat, abw: viewAbw, daten: viewDaten, akte: viewAkte, team: viewTeam, stamm: viewStamm }[S.view] || viewHeute;
  $('#main').innerHTML = `<div class="crumb no-print">${esc(CFG.firma || 'ProTax')}<i></i>${VIEW_LABEL[S.view] || ''}</div>` + v() + `<div class="ver-m no-print">${esc(APP_VERSION)}</div>`;
  renderTimer();
}
const go = v => { S.view = v; if ((v === 'monat' || v === 'akte' || v === 'daten') && !S.viewUpn) S.viewUpn = S.me.upn; if (v === 'akte') S.akte = null; if (v === 'daten') S.pd = null; render(); window.scrollTo(0, 0); };
const reload = () => act(async () => { await loadAll(); }, 'Aktualisiert');

function viewMissing() {
  return `<div class="card"><div class="card-h"><div class="card-t">Einrichtung nicht abgeschlossen</div></div><div class="card-b">
    <p>In der SharePoint-Website fehlen die Listen: <b>${S.missing.map(esc).join(', ')}</b>.</p>
    <p class="muted sm" style="margin-top:8px">${S.leitung ? 'Unter „Stammdaten“ kannst du sie automatisch anlegen.' : 'Bitte die Kanzleileitung informieren.'}</p>
    ${S.leitung ? `<div class="actions"><button class="btn primary" onclick="go('stamm')">Zur Einrichtung</button></div>` : ''}</div></div>`;
}
function noMa() {
  return `<div class="banner warn">Für <b>${esc(S.me.upn)}</b> ist noch kein Mitarbeiter-Stammsatz angelegt. ${S.leitung ? 'Lege ihn unter „Stammdaten“ an.' : 'Bitte die Kanzleileitung ansprechen.'}</div>`;
}
function checksHtml(checks, withInfo = true) {
  const list = checks.filter(c => withInfo || c.l !== 'info');
  return list.length ? `<div class="checks">${list.map(c => `<span class="ck ${c.l}">${esc(c.t)}</span>`).join('')}</div>` : '';
}

/* ---------- Heute ---------- */
function viewHeute() {
  const m = myMa(); if (!m) return noMa();
  const t = today(), st = stempelStatus(S.me.upn), data = dataFor(S.me.upn), nm = nowMin();
  const row = Z.tagesZeile(m, t, data, CFG, t, nm);
  // Woche (Mo bis heute)
  const mo = Z.addDays(t, -(Z.isoWd(t) - 1)); let wIst = 0, wSoll = 0;
  for (let d = mo; d <= t; d = Z.addDays(d, 1)) { const r = Z.tagesZeile(m, d, data, CFG, t, nm); wIst += r.ist + r.gut; wSoll += r.soll; }
  const saldo = Z.saldo(m, data, CFG, Z.addDays(t, -1));
  const ul = Z.urlaub(m, +t.slice(0, 4), S.abw, CFG);
  const alteOffen = offeneBloecke(S.me.upn).filter(b => b !== st.b);
  const abwHeute = Z.abwAm(S.abw, S.me.upn, t);
  const stateTxt = st.s === 'aus' ? '<span class="badge b-grey">Nicht eingestempelt</span>'
    : st.s === 'pause' ? `<span class="badge b-warn">Pause seit ${esc(st.seit)}</span>`
    : `<span class="badge b-ok">Anwesend seit ${esc(st.b.beginn)} · ${esc(st.b.ort)}</span>`;
  const btns = st.s === 'aus'
    ? `<button class="btn big ok" onclick="kommen()" ${S.busy ? 'disabled' : ''}>Kommen</button>`
    : st.s === 'an'
      ? `<button class="btn big" onclick="pause()" ${S.busy ? 'disabled' : ''}>Pause</button><button class="btn big danger" onclick="gehen()" ${S.busy ? 'disabled' : ''}>Gehen</button>`
      : `<button class="btn big ok" onclick="pauseEnde()" ${S.busy ? 'disabled' : ''}>Pause beenden</button><button class="btn big danger" onclick="gehen()" ${S.busy ? 'disabled' : ''}>Gehen</button>`;
  const orte = st.s === 'aus' ? `<div class="seg">${(CFG.orte || []).map(o => `<label><input type="radio" name="ort" value="${esc(o)}" ${o === S.ort ? 'checked' : ''} onchange="S.ort=this.value">${esc(o)}</label>`).join('')}</div>` : '';
  const heuteBl = S.buchungen.filter(b => b.upn === S.me.upn && b.datum === t && b.status !== 'storniert').sort((a, c) => a.beginn.localeCompare(c.beginn));
  return `
  ${alteOffen.map(b => `<div class="banner err">Am ${de(b.datum)} fehlt das Ausstempeln (Beginn ${esc(b.beginn)}). <button class="link" onclick="editBlock('${b.id}')">Jetzt nachtragen</button></div>`).join('')}
  ${abwHeute.length ? `<div class="banner info">Heute eingetragen: ${abwHeute.map(a => esc(a.art) + (a.halbtag ? ' (halber Tag)' : '') + ' – ' + esc(a.status)).join(', ')}</div>` : ''}
  <div class="card"><div class="card-b">
    <div class="clock"><div class="time" id="clock">${nowHHMM()}</div><div class="muted sm">${WD[Z.isoWd(t)]}, ${de(t)}${row.feiertag ? ' · ' + esc(row.feiertag) : ''}</div><div class="state" style="margin-top:6px">${stateTxt}</div></div>
    ${orte ? `<div class="row" style="justify-content:center;margin-bottom:14px">${orte}</div>` : ''}
    <div class="stamp">${btns}</div>
    <div class="row" style="justify-content:center;margin-top:14px"><button class="link" onclick="newBlock('${t}','${S.me.upn}')">Zeit nachtragen</button></div>
  </div></div>
  <div class="card"><div class="card-b"><div class="kv">
    <div class="stat"><div class="l">Heute</div><div class="v">${Z.dur(row.ist)}</div><div class="muted sm">Soll ${Z.dur(row.soll)}</div></div>
    <div class="stat"><div class="l">Diese Woche</div><div class="v">${Z.dur(wIst)}</div><div class="muted sm">Soll ${Z.dur(wSoll)}</div></div>
    <div class="stat"><div class="l">Gleitzeitsaldo (bis gestern)</div><div class="v ${saldo < 0 ? 'neg' : saldo > 0 ? 'pos' : ''}">${Z.dur(saldo, true)}</div><div class="muted sm">Stunden</div></div>
    <div class="stat"><div class="l">Resturlaub ${t.slice(0, 4)}</div><div class="v">${fmtTage(ul.rest)}</div><div class="muted sm">${ul.beantragt ? fmtTage(ul.beantragt) + ' beantragt' : 'Tage'}</div></div>
  </div></div></div>
  <div class="card"><div class="card-h"><div class="card-t">Buchungen heute</div><span class="dot ${row.ampel}"></span></div>
    ${heuteBl.length ? `<div class="tblwrap"><table class="tbl"><thead><tr><th>Zeit</th><th>Pausen</th><th>Ort</th><th></th></tr></thead><tbody>
    ${heuteBl.map(b => `<tr><td>${esc(b.beginn)} – ${esc(b.ende || 'läuft')}${b.nachtrag ? ' <span class="badge b-grey">Nachtrag</span>' : ''}</td><td class="sm">${esc(Z.pausenText(b.pausen)) || '—'}${b.pauseMin ? ` · pauschal ${b.pauseMin} min` : ''}</td><td class="sm">${esc(b.ort)}</td>
      <td style="text-align:right"><button class="link" onclick="editBlock('${b.id}')">Korrigieren</button></td></tr>`).join('')}</tbody></table></div>`
    : '<div class="empty">Heute noch nichts erfasst.</div>'}
    ${row.checks.filter(c => c.l !== 'info').length ? `<div class="card-b" style="padding-top:0">${checksHtml(row.checks, false)}</div>` : ''}
  </div>`;
}
const fmtTage = n => (Math.round(n * 2) / 2).toLocaleString('de-DE');

/* ---------- Monatsnachweis ---------- */
function monthNav() {
  const [y, m] = S.month.split('-').map(Number);
  const opts = S.leitung ? `<select class="inp" style="width:auto" onchange="S.viewUpn=this.value;render()">${S.ma.map(x => `<option value="${esc(x.upn)}" ${x.upn === S.viewUpn ? 'selected' : ''}>${esc(x.name)}</option>`).join('')}</select>` : '';
  return `<div class="row no-print">${opts}<button class="btn sm" onclick="shiftMonth(-1)">‹</button><b style="min-width:120px;text-align:center">${MONATE[m - 1]} ${y}</b><button class="btn sm" onclick="shiftMonth(1)">›</button></div>`;
}
async function shiftMonth(n) {
  let [y, m] = S.month.split('-').map(Number); m += n; if (m < 1) { m = 12; y--; } if (m > 12) { m = 1; y++; }
  S.month = y + '-' + Z.pad(m);
  const from = S.month + '-01';
  if (from < S.loadedFrom) { await act(() => ensureLoaded(Z.addDays(from, -2))); return; }
  render();
}
function viewMonat() {
  const upn = S.leitung ? (S.viewUpn || S.me.upn) : S.me.upn, m = maBy(upn);
  if (!m) return S.leitung && S.ma.length ? (S.viewUpn = S.ma[0].upn, viewMonat()) : noMa();
  const [y, mo] = S.month.split('-').map(Number), t = today();
  const data = dataFor(upn), mon = Z.monat(m, y, mo, data, CFG, t, nowMin());
  const ende = S.month + '-' + Z.pad(new Date(y, mo, 0).getDate());
  const bis = ende < t ? ende : Z.addDays(t, -1);
  const saldo = Z.saldo(m, data, CFG, bis);
  const rows = mon.rows.map(r => {
    const bl = r.blocks.slice().sort((a, c) => a.beginn.localeCompare(c.beginn));
    const zeiten = bl.map(b => `<div class="blk">${esc(b.beginn)}–${esc(b.ende || (r.date === t ? 'läuft' : '?'))}
        <span class="muted sm">${(b.pausen || []).length || b.pauseMin ? 'P ' + esc(Z.pausenText(b.pausen)) + (b.pauseMin ? ' +' + b.pauseMin + ' min' : '') : ''} ${esc(b.ort)}</span>
        <span class="no-print">${canEdit(upn, r.date) ? `<button class="link" onclick="editBlock('${b.id}')">ändern</button>` : ''} ${b.nachtrag || b.korrAm ? `<button class="link" onclick="verlauf('${b.id}')">Verlauf</button>` : ''}</span></div>`).join('');
    const abw = [r.feiertag, ...r.abw.map(a => a.art + (a.halbtag ? ' ½' : '') + (a.status === 'beantragt' ? ' (beantragt)' : ''))].filter(Boolean).map(esc).join(', ');
    const cls = (r.soll === 0 && !r.blocks.length ? 'we ' : '') + (r.date === t ? 'today' : '');
    return `<tr class="${cls}"><td style="white-space:nowrap">${WD[r.wd]} ${deKurz(r.date)}</td>
      <td>${zeiten || (canEdit(upn, r.date) && !r.future ? `<button class="link no-print" onclick="newBlock('${r.date}','${upn}')">+ erfassen</button>` : '')}</td>
      <td class="sm">${abw}</td><td class="num">${r.soll ? Z.dur(r.soll) : ''}</td><td class="num">${r.ist || r.gut ? Z.dur(r.ist + r.gut) : ''}</td>
      <td class="num">${r.future || (!r.soll && !r.ist) ? '' : signed(r.diff)}</td>
      <td>${r.future || (!r.blocks.length && !r.checks.length) ? '' : `<span class="dot ${r.ampel}"></span>`}${checksHtml(r.checks)}</td></tr>`;
  }).join('');
  return `<div class="card">
    <div class="card-h"><div><div class="card-t">Monatsnachweis · ${esc(m.name)}</div><div class="muted sm">${m.wochenstunden} h/Woche · Arbeitstage ${esc(m.arbeitstage.split(',').map(x => WD[+x]).join(', '))}</div></div>${monthNav()}</div>
    <div class="print-only" style="padding:10px 18px"><b>${esc(CFG.firma || '')}</b> – Arbeitszeitnachweis ${MONATE[mo - 1]} ${y} – ${esc(m.name)}</div>
    <div class="card-b"><div class="kv">
      <div class="stat"><div class="l">Soll</div><div class="v">${Z.dur(mon.soll)}</div></div>
      <div class="stat"><div class="l">Ist (inkl. Gutschriften)</div><div class="v">${Z.dur(mon.ist + mon.gut)}</div><div class="muted sm">davon Gutschrift ${Z.dur(mon.gut)}</div></div>
      <div class="stat"><div class="l">Differenz Monat</div><div class="v ${mon.diff < 0 ? 'neg' : mon.diff > 0 ? 'pos' : ''}">${Z.dur(mon.diff, true)}</div></div>
      <div class="stat"><div class="l">Saldo per ${de(bis)}</div><div class="v ${saldo < 0 ? 'neg' : saldo > 0 ? 'pos' : ''}">${Z.dur(saldo, true)}</div><div class="muted sm">${mon.rot ? mon.rot + ' Tag(e) mit Verstoß' : 'keine Verstöße'}</div></div>
    </div>
    <div class="row no-print" style="margin-top:14px;justify-content:flex-end"><button class="btn sm" onclick="csvMonat('${upn}')">Excel (CSV)</button><button class="btn sm" onclick="window.print()">Drucken / PDF</button></div></div>
    <div class="tblwrap"><table class="tbl"><thead><tr><th>Tag</th><th>Zeiten</th><th>Abwesenheit</th><th class="num">Soll</th><th class="num">Ist</th><th class="num">Diff.</th><th>Prüfung</th></tr></thead>
    <tbody>${rows}</tbody><tfoot><tr><td colspan="3">Summe</td><td class="num">${Z.dur(mon.soll)}</td><td class="num">${Z.dur(mon.ist + mon.gut)}</td><td class="num">${signed(mon.diff)}</td><td></td></tr></tfoot></table></div>
    <div class="print-only" style="padding:0 18px 18px"><div class="sign"><div>Datum, Unterschrift Mitarbeiter/in</div><div>Datum, Unterschrift Arbeitgeber</div></div></div>
  </div>`;
}
function csvMonat(upn) {
  const m = maBy(upn), [y, mo] = S.month.split('-').map(Number), mon = Z.monat(m, y, mo, dataFor(upn), CFG, today(), nowMin());
  const q = v => '"' + String(v ?? '').replace(/"/g, '""') + '"', h = n => (Math.round(n / 60 * 100) / 100).toFixed(2).replace('.', ',');
  const head = ['Datum', 'Tag', 'Zeiten', 'Pausen (min)', 'Ort', 'Abwesenheit/Feiertag', 'Soll (h)', 'Ist (h)', 'Gutschrift (h)', 'Differenz (h)', 'Prüfung'];
  const lines = mon.rows.map(r => [de(r.date), WD[r.wd], r.blocks.map(b => b.beginn + '-' + (b.ende || '')).join(' / '), r.t.pauseGesamt || '', [...new Set(r.blocks.map(b => b.ort))].join(' / '),
    [r.feiertag, ...r.abw.map(a => a.art + (a.halbtag ? ' (halb)' : '') + ' ' + a.status)].filter(Boolean).join(', '), h(r.soll), h(r.ist), h(r.gut), r.future ? '' : h(r.diff),
    r.checks.filter(c => c.l !== 'info').map(c => c.t).join(' | ')].map(q).join(';'));
  download(`Arbeitszeit_${m.name.replace(/\W+/g, '_')}_${S.month}.csv`, '﻿' + [head.map(q).join(';'), ...lines].join('\r\n'), 'text/csv;charset=utf-8');
}

/* ---------- Nachtrag / Korrektur ---------- */
function modal(html) { $('#modal').innerHTML = `<div class="back" onclick="if(event.target===this)closeModal()"><div class="dlg">${html}</div></div>`; }
const closeModal = () => { $('#modal').innerHTML = ''; };
function blockForm(b, isNew) {
  const ok = canEdit(b.upn, b.datum);
  const frist = !ok && !S.leitung ? `<div class="banner warn" style="margin-bottom:12px">Korrekturen sind nur ${CFG.korrekturFristTage} Tage rückwirkend möglich. Bitte die Kanzleileitung ansprechen.</div>` : '';
  modal(`<div class="card-h"><div class="card-t">${isNew ? 'Zeit nachtragen' : 'Buchung korrigieren'} · ${esc(b.name)}</div></div><div class="card-b">${frist}
    <div class="fg2">
      <div class="f"><label>Datum</label><input id="ff-datum" type="date" value="${esc(b.datum)}" max="${today()}" ${isNew ? '' : 'disabled'}></div>
      <div class="f"><label>Ort</label><select id="ff-ort">${(CFG.orte || []).map(o => `<option ${o === b.ort ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select></div>
      <div class="f"><label>Beginn</label><input id="ff-beginn" type="time" value="${esc(b.beginn)}"></div>
      <div class="f"><label>Ende</label><input id="ff-ende" type="time" value="${esc(b.ende)}"></div>
      <div class="f span2"><label>Pausen mit Uhrzeit, z. B. 12:00-12:30, 15:00-15:15</label><input id="ff-pausen" value="${esc(Z.pausenText(b.pausen).replace(/–/g, '-').replace(/…/g, ''))}"></div>
      <div class="f"><label>oder Pause pauschal (Minuten)</label><input id="ff-pmin" type="number" min="0" step="5" value="${b.pauseMin || ''}"></div>
      <div class="f"><label>Bemerkung</label><input id="ff-bem" value="${esc(b.bemerkung)}"></div>
      <div class="f span2"><label>Begründung (Pflicht) – wird mit Name und Zeitpunkt gespeichert</label><input id="ff-grund" placeholder="z. B. Ausstempeln vergessen, Termin beim Mandanten"></div>
    </div>
    ${b.grund ? `<div class="muted sm" style="margin-top:8px">Bisher: ${esc(b.grund)}</div>` : ''}
    <div class="actions">
      ${!isNew && ok ? `<button class="btn danger" onclick="stornoBlock('${b.id}')">Stornieren</button>` : ''}
      <button class="btn" onclick="closeModal()">Abbrechen</button>
      <button class="btn primary" onclick="saveBlock(${isNew ? 'null' : `'${b.id}'`},'${esc(b.upn)}')" ${ok ? '' : 'disabled'}>Speichern</button></div></div>`);
}
function newBlock(datum, upn) {
  const m = maBy(upn) || { name: S.me.name };
  blockForm({ upn, name: m.name, datum, beginn: '', ende: '', pausen: [], pauseMin: 0, ort: (CFG.orte || [''])[0], bemerkung: '', grund: '' }, true);
}
function editBlock(id) { const b = S.buchungen.find(x => x.id === id); if (b) blockForm(b, false); }
const stamp = txt => `${de(today())} ${nowHHMM()} ${S.me.name}: ${txt}`;
function saveBlock(id, upn) {
  const v = n => $('#ff-' + n).value.trim();
  const datum = v('datum'), beginn = v('beginn'), ende = v('ende'), grund = v('grund'), pmin = +v('pmin') || 0;
  const pp = Z.parsePausen(v('pausen'));
  if (!datum || datum > today()) return toast('Datum fehlt oder liegt in der Zukunft', 'err');
  if (!canEdit(upn, datum)) return toast('Keine Berechtigung für dieses Datum', 'err');
  if (Z.tm(beginn) == null) return toast('Beginn fehlt', 'err');
  if (!ende && datum !== today()) return toast('Ende fehlt', 'err');
  if (ende && ende === beginn) return toast('Beginn und Ende sind gleich', 'err');
  if (pp.fehler.length) return toast('Pausen nicht lesbar: ' + pp.fehler.join(', '), 'err');
  if (!grund) return toast('Bitte eine Begründung angeben', 'err');
  const fields = { beginn, ende, pausen: pp.pausen, pauseMin: pmin, ort: v('ort'), bemerkung: v('bem') };
  closeModal();
  act(async () => {
    if (!id) {
      const m = maBy(upn);
      const o = await S.store.create('buchungen', Object.assign({ upn, name: m ? m.name : S.me.name, datum, nachtrag: true, grund: stamp(grund), status: 'aktiv', korrVon: '', korrAm: '' }, fields));
      upsertLocal(S.buchungen, o); await fuerAndere('buchungen', o);
    } else {
      const b = S.buchungen.find(x => x.id === id);
      Object.assign(fields, { grund: (b.grund ? b.grund + ' | ' : '') + stamp(grund), korrVon: S.me.name, korrAm: new Date().toISOString() });
      await S.store.update('buchungen', id, fields); Object.assign(b, fields);
    }
  }, 'Gespeichert');
}
function stornoBlock(id) {
  const grund = $('#ff-grund').value.trim(); if (!grund) return toast('Bitte eine Begründung für die Stornierung angeben', 'err');
  const b = S.buchungen.find(x => x.id === id); closeModal();
  act(async () => {
    const f = { status: 'storniert', grund: (b.grund ? b.grund + ' | ' : '') + stamp('Storno – ' + grund), korrVon: S.me.name, korrAm: new Date().toISOString() };
    await S.store.update('buchungen', id, f); Object.assign(b, f);
  }, 'Buchung storniert');
}
async function verlauf(id) {
  modal(`<div class="card-h"><div class="card-t">Änderungsverlauf</div></div><div class="empty">Lade …</div>`);
  try {
    const vs = await S.store.versions('buchungen', id);
    modal(`<div class="card-h"><div class="card-t">Änderungsverlauf</div><button class="btn sm" onclick="closeModal()">Schließen</button></div>
      <div class="tblwrap"><table class="tbl"><thead><tr><th>Zeitpunkt</th><th>Von</th><th>Stand</th></tr></thead><tbody>
      ${vs.map(v => `<tr><td class="sm">${new Date(v.at).toLocaleString('de-DE')}</td><td class="sm">${esc(v.by)}</td><td class="sm">${v.data ? `${esc(v.data.beginn)}–${esc(v.data.ende)} · P ${esc(Z.pausenText(v.data.pausen))}${v.data.pauseMin ? ' +' + v.data.pauseMin + ' min' : ''} · ${esc(v.data.status)}` : '—'}</td></tr>`).join('') || '<tr><td colspan="3" class="muted">Keine Versionen gefunden.</td></tr>'}
      </tbody></table></div><div class="card-b muted sm">Der vollständige Verlauf steht zusätzlich in SharePoint (Liste ZE_Buchungen → Element → Versionsverlauf).</div>`);
  } catch (e) { modal(`<div class="card-b"><div class="banner err">${esc(e.message)}</div><div class="actions"><button class="btn" onclick="closeModal()">Schließen</button></div></div>`); }
}

/* ---------- Abwesenheiten ---------- */
function viewAbw() {
  const upn = S.leitung ? (S.viewUpn || S.me.upn) : S.me.upn, m = maBy(upn), y = +today().slice(0, 4);
  const ul = m ? Z.urlaub(m, y, S.abw, CFG) : null;
  const list = S.abw.filter(a => S.leitung || a.upn === S.me.upn).filter(a => (a.bis || a.von) >= (y - 1) + '-01-01').sort((a, b) => b.von.localeCompare(a.von));
  const badge = s => `<span class="badge ${{ genehmigt: 'b-ok', gemeldet: 'b-ok', beantragt: 'b-warn', abgelehnt: 'b-err', storniert: 'b-grey' }[s] || 'b-grey'}">${esc(s)}</span>`;
  return `${m ? `<div class="card"><div class="card-h"><div class="card-t">Urlaub ${y} · ${esc(m.name)}</div>
    ${S.leitung ? `<select class="inp" style="width:auto" onchange="S.viewUpn=this.value;render()">${S.ma.map(x => `<option value="${esc(x.upn)}" ${x.upn === upn ? 'selected' : ''}>${esc(x.name)}</option>`).join('')}</select>` : ''}</div>
    <div class="card-b"><div class="kv">
    <div class="stat"><div class="l">Anspruch inkl. Vorjahr</div><div class="v">${fmtTage(ul.anspruch)}</div></div>
    <div class="stat"><div class="l">Genehmigt</div><div class="v">${fmtTage(ul.genommen)}</div></div>
    <div class="stat"><div class="l">Beantragt</div><div class="v">${fmtTage(ul.beantragt)}</div></div>
    <div class="stat"><div class="l">Rest</div><div class="v ${ul.rest < 0 ? 'neg' : ''}">${fmtTage(ul.rest)}</div></div></div></div></div>` : noMa()}
  <div class="card"><div class="card-h"><div class="card-t">${S.leitung ? 'Abwesenheiten aller Mitarbeitenden' : 'Meine Abwesenheiten'}</div><button class="btn primary sm" onclick="abwForm()">+ Eintragen</button></div>
  ${list.length ? `<div class="tblwrap"><table class="tbl"><thead><tr>${S.leitung ? '<th>Mitarbeiter</th>' : ''}<th>Art</th><th>Zeitraum</th><th class="num">Arbeitstage</th><th>Status</th><th></th></tr></thead><tbody>
  ${list.map(a => { const mm = maBy(a.upn); const tage = mm ? Z.arbeitstageZaehlen(mm, a.von, a.bis || a.von, CFG, a.halbtag) : ''; return `<tr>${S.leitung ? `<td>${esc(a.name)}</td>` : ''}<td>${esc(a.art)}${a.bemerkung ? `<div class="muted sm">${esc(a.bemerkung)}</div>` : ''}</td>
    <td>${de(a.von)}${a.bis && a.bis !== a.von ? ' – ' + de(a.bis) : ''}${a.halbtag ? ' (½)' : ''}</td><td class="num">${fmtTage(tage)}</td><td>${badge(a.status)}${a.entschVon ? `<div class="muted sm">${esc(a.entschVon)}</div>` : ''}</td>
    <td style="text-align:right;white-space:nowrap">${abwActions(a)}</td></tr>`; }).join('')}</tbody></table></div>` : '<div class="empty">Keine Abwesenheiten eingetragen.</div>'}</div>`;
}
function abwActions(a) {
  const out = [];
  if (S.leitung && a.status === 'beantragt') out.push(`<button class="btn sm ok" onclick="abwStatus('${a.id}','genehmigt')">Genehmigen</button>`, `<button class="btn sm danger" onclick="abwStatus('${a.id}','abgelehnt')">Ablehnen</button>`);
  const eigen = a.upn === S.me.upn && (a.status === 'beantragt' || ((a.status === 'gemeldet' || a.status === 'genehmigt') && a.von > today()));
  if ((S.leitung && !['storniert', 'abgelehnt'].includes(a.status)) || eigen) out.push(`<button class="btn sm" onclick="abwStatus('${a.id}','storniert')">Stornieren</button>`);
  return out.join(' ');
}
function abwForm() {
  const sel = S.leitung ? `<div class="f span2"><label>Mitarbeiter</label><select id="af-upn">${S.ma.map(x => `<option value="${esc(x.upn)}" ${x.upn === (S.viewUpn || S.me.upn) ? 'selected' : ''}>${esc(x.name)}</option>`).join('')}</select></div>` : '';
  modal(`<div class="card-h"><div class="card-t">Abwesenheit eintragen</div></div><div class="card-b"><div class="fg2">${sel}
    <div class="f span2"><label>Art</label><select id="af-art">${ARTEN.map(a => `<option>${a}</option>`).join('')}</select></div>
    <div class="f"><label>Von</label><input id="af-von" type="date" value="${today()}" oninput="abwTage()"></div>
    <div class="f"><label>Bis</label><input id="af-bis" type="date" value="${today()}" oninput="abwTage()"></div>
    <label class="chk span2"><input type="checkbox" id="af-halb" onchange="abwTage()"> Halber Tag (nur bei einem Tag)</label>
    <div class="f span2"><label>Bemerkung</label><input id="af-bem"></div></div>
    <div class="muted sm" id="af-tage" style="margin-top:8px"></div>
    <div class="muted sm" style="margin-top:4px">${S.leitung ? 'Von der Leitung eingetragen = direkt genehmigt.' : 'Urlaub, Sonderurlaub und Überstundenabbau werden beantragt; Krankheit und Fortbildung gelten als gemeldet. Eine AU-Bescheinigung bitte wie gewohnt einreichen – sie gehört nicht in dieses System.'}</div>
    <div class="actions"><button class="btn" onclick="closeModal()">Abbrechen</button><button class="btn primary" onclick="saveAbw()">Speichern</button></div></div>`);
  abwTage();
}
function abwTage() {
  const upn = $('#af-upn')?.value || S.me.upn, m = maBy(upn), von = $('#af-von').value, bis = $('#af-bis').value || von;
  $('#af-tage').textContent = m && von && bis >= von ? `= ${fmtTage(Z.arbeitstageZaehlen(m, von, bis, CFG, $('#af-halb').checked && von === bis))} Arbeitstag(e)` : '';
}
function saveAbw() {
  const upn = $('#af-upn')?.value || S.me.upn, m = maBy(upn), art = $('#af-art').value, von = $('#af-von').value, bis = $('#af-bis').value || von;
  const halbtag = $('#af-halb').checked;
  if (!m) return toast('Kein Stammsatz für diesen Mitarbeiter', 'err');
  if (!von || bis < von) return toast('Zeitraum prüfen', 'err');
  if (halbtag && von !== bis) return toast('Halber Tag nur bei einem einzelnen Tag', 'err');
  const ueber = S.abw.find(a => a.upn === upn && !['storniert', 'abgelehnt'].includes(a.status) && a.von <= bis && von <= (a.bis || a.von));
  if (ueber && !confirm(`Überschneidung mit „${ueber.art}“ ${de(ueber.von)}–${de(ueber.bis)}. Trotzdem speichern?`)) return;
  const status = S.leitung ? (['Krank', 'Fortbildung'].includes(art) ? 'gemeldet' : 'genehmigt') : (['Krank', 'Fortbildung'].includes(art) ? 'gemeldet' : 'beantragt');
  closeModal();
  act(async () => {
    const o = await S.store.create('abw', { upn, name: m.name, art, von, bis, halbtag, bemerkung: $('#af-bem')?.value || '', status, entschVon: S.leitung ? S.me.name : '', entschAm: S.leitung ? new Date().toISOString() : '' });
    upsertLocal(S.abw, o); await fuerAndere('abw', o);
  }, status === 'beantragt' ? 'Antrag gestellt' : 'Eingetragen');
}
function abwStatus(id, status) {
  const a = S.abw.find(x => x.id === id); if (!a) return;
  if (status !== 'genehmigt' && !confirm(`${a.art} ${de(a.von)} wirklich ${status === 'storniert' ? 'stornieren' : 'ablehnen'}?`)) return;
  act(async () => {
    const f = { status, entschVon: S.me.name, entschAm: new Date().toISOString() };
    await S.store.update('abw', id, f); Object.assign(a, f);
  }, 'Status: ' + status);
}

/* ---------- Persönliche Daten ---------- */
const BUNDESLAENDER = ['', 'Baden-Württemberg', 'Bayern', 'Berlin', 'Brandenburg', 'Bremen', 'Hamburg', 'Hessen', 'Mecklenburg-Vorpommern', 'Niedersachsen', 'Nordrhein-Westfalen', 'Rheinland-Pfalz', 'Saarland', 'Sachsen', 'Sachsen-Anhalt', 'Schleswig-Holstein', 'Thüringen'];
const PD_SECTIONS = [
  ['Allgemein', [['titel', 'Titel', 'select', ['', 'Dr.', 'Prof.', 'Prof. Dr.']], ['anrede', 'Anrede', 'select', ['', 'Frau', 'Herr', 'Divers']], ['vorname', 'Vorname', 'text'], ['nachname', 'Nachname', 'text'], ['geburtsdatum', 'Geburtsdatum', 'date'], ['_dienst', 'E-Mail (dienstlich)', 'fix']]],
  ['Anschrift', [['land', 'Staat', 'text'], ['plz', 'Postleitzahl', 'text'], ['bundesland', 'Bundesland', 'select', BUNDESLAENDER], ['ort', 'Stadt', 'text'], ['_row', [['strasse', 'Straße', 'text'], ['hausnr', 'Hausnummer', 'text']]], ['adresszusatz', 'Zusätzliche Adressinformationen', 'text']]],
  ['Zusätzliche Information', [['staatsang', 'Staatsangehörigkeit', 'text'], ['privMail', 'Private E-Mail', 'email'], ['mobil', 'Private Telefonnummer', 'tel'], ['festnetz', 'Private Festnetznummer', 'tel'], ['_h', 'Notfallkontakt'], ['notfallName', 'Name', 'text'], ['notfallBez', 'Beziehung (z. B. Partner/in)', 'text'], ['notfallTel', 'Telefon', 'tel']]]];
const pdKeys = () => PD_SECTIONS.flatMap(([, fs]) => fs.flatMap(f => f[0] === '_row' ? f[1].map(x => x[0]) : f[0][0] === '_' ? [] : [f[0]]));
async function loadPd() {
  S.pd = { loading: true }; render();
  try {
    if (S.store.pdFind && !(S.store.lists || {}).pd) await S.store.pdFind();
    const st = await S.store.pdStatus();
    if (st.status !== 'ok') { S.pd = { status: st.status }; return render(); }
    const items = await S.store.load('pd');
    S.pd = { status: 'ok', unchecked: st.unchecked, items: S.leitung ? items : items.filter(x => x.upn === S.me.upn) };
  } catch (e) { console.error(e); S.pd = { status: 'error', error: e.message || String(e) }; }
  render();
}
function viewDaten() {
  const upn = S.leitung ? (S.viewUpn || S.me.upn) : S.me.upn, m = maBy(upn), P = S.pd;
  const sel = S.leitung && S.ma.length ? `<select class="inp" style="width:auto" onchange="S.viewUpn=this.value;render()">${S.ma.map(x => `<option value="${esc(x.upn)}" ${x.upn === upn ? 'selected' : ''}>${esc(x.name)}</option>`).join('')}</select>` : '';
  const name = m ? m.name : S.me.name;
  const head = `<div class="card-h"><div><div class="card-t">Persönliche Daten · ${esc(name)}</div><div class="muted sm">${upn === S.me.upn ? 'Bitte aktuell halten – sichtbar nur für dich und die Kanzleileitung' : 'Von der Mitarbeiterin bzw. dem Mitarbeiter gepflegt – sichtbar nur für sie/ihn und die Leitung'}</div></div>${sel}</div>`;
  if (!P) { setTimeout(loadPd); return `<div class="card">${head}<div class="empty">Lade …</div></div>`; }
  if (P.loading) return `<div class="card">${head}<div class="empty">Lade …</div></div>`;
  if (P.status === 'error') return `<div class="card">${head}<div class="card-b"><div class="banner err">${esc(P.error)}</div><div class="actions"><button class="btn" onclick="loadPd()">Erneut versuchen</button></div></div></div>`;
  if (P.status === 'missing') return `<div class="card">${head}<div class="card-b">${S.leitung
    ? `<p>Für die persönlichen Daten wird in SharePoint eine eigene Liste <b>${LISTS.pd.name}</b> angelegt. Jede Person sieht dort nur ihren eigenen Datensatz.</p><div class="actions"><button class="btn primary" onclick="pdSetupAction()">Bereich einrichten</button></div>`
    : '<div class="empty">Dieser Bereich wird gerade von der Kanzleileitung eingerichtet.</div>'}</div></div>`;
  if (P.status === 'unsafe') return `<div class="card">${head}<div class="card-b">${S.leitung ? `
    <div class="banner warn"><b>Liste noch nicht abgesichert.</b> In ${LISTS.pd.name} können sonst alle Mitglieder die Daten der anderen lesen. Bitte in SharePoint einstellen:</div>
    <ol class="steps"><li>SharePoint-Website → <b>Websiteinhalte</b> → bei <b>${LISTS.pd.name}</b> auf <b>⋯ → Einstellungen</b></li>
    <li><b>Erweiterte Einstellungen</b> → Lesezugriff: <b>Elemente lesen, die vom Benutzer erstellt wurden</b></li>
    <li>Erstellungs- und Bearbeitungszugriff: <b>Elemente erstellen und bearbeiten, die vom Benutzer erstellt wurden</b></li>
    <li>Anlagen <b>Deaktiviert</b>, Kommentare <b>Nein</b> → <b>OK</b>; unter <b>Versionsverwaltungseinstellungen</b> den Versionsverlauf aktivieren</li></ol>
    <div class="actions"><button class="btn primary" onclick="loadPd()">Erneut prüfen</button></div>` : '<div class="empty">Dieser Bereich wird gerade von der Kanzleileitung eingerichtet.</div>'}</div></div>`;
  const p = (P.items || []).find(x => x.upn === upn) || {};
  if (!p.vorname && !p.nachname && m) { const parts = m.name.split(' '); p.vorname = p.vorname || parts.slice(0, -1).join(' '); p.nachname = p.nachname || parts.slice(-1).join(''); }
  if (!p.land) p.land = 'Deutschland'; if (!p.staatsang && !p.id) p.staatsang = '';
  const field = ([k, l, t, opts]) => k === '_row' ? `<div class="pd-row">${l.map(field).join('')}</div>`
    : k === '_h' ? `<div class="pd-sub">${l}</div>`
    : t === 'fix' ? `<div class="f"><label>${l}</label><input value="${esc(upn)}" disabled></div>`
    : t === 'select'
    ? `<div class="f"><label>${l}</label><select id="pd-${k}">${opts.map(o => `<option value="${esc(o)}" ${o === (p[k] || '') ? 'selected' : ''}>${esc(o || '–')}</option>`).join('')}</select></div>`
    : `<div class="f"><label>${l}</label><input id="pd-${k}" type="${t}" value="${esc(p[k] || '')}" ${t === 'tel' ? 'inputmode="tel"' : ''} autocomplete="off"></div>`;
  const kann = S.leitung || upn === S.me.upn;
  return `<div class="card">${head}<div class="card-b">
    ${P.unchecked && S.leitung ? '<div class="banner info" style="margin-bottom:16px">Die Absicherung der Liste konnte nicht automatisch geprüft werden. Bitte sicherstellen, dass in SharePoint „nur eigene Elemente lesen/bearbeiten“ eingestellt ist.</div>' : ''}
    <div class="pd-grid">${PD_SECTIONS.map(([t, fs]) => `<section><h3 class="pd-h">${t}</h3><div class="pd-fields">${fs.map(field).join('')}</div></section>`).join('')}</div>
    <div class="actions"><span class="muted sm" style="margin-right:auto;align-self:center">${p.geaendertAm ? 'Zuletzt geändert: ' + new Date(p.geaendertAm).toLocaleString('de-DE') : ''}</span>
      <button class="btn primary" onclick="pdSaveAction('${esc(upn)}')" ${kann ? '' : 'disabled'}>Speichern</button></div></div></div>`;
}
const pdSetupAction = () => act(async () => { await S.store.pdSetup(); S.pd = null; }, 'Bereich angelegt');
function pdSaveAction(upn) {
  const v = k => ($('#pd-' + k)?.value || '').trim(), o = { upn };
  pdKeys().forEach(k => { o[k] = v(k); });
  if (o.privMail && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(o.privMail)) return toast('Private E-Mail ist ungültig', 'err');
  if (o.plz && /deutschland/i.test(o.land) && !/^\d{5}$/.test(o.plz)) return toast('Postleitzahl muss 5 Ziffern haben', 'err');
  for (const k of ['mobil', 'festnetz', 'notfallTel']) if (o[k] && !/^[+\d][\d\s\/()-]{4,}$/.test(o[k])) return toast('Telefonnummer prüfen: ' + o[k], 'err');
  const m = maBy(upn); o.name = m ? m.name : [o.vorname, o.nachname].filter(Boolean).join(' '); o.geaendertAm = new Date().toISOString();
  act(async () => {
    const ex = (S.pd.items || []).find(x => x.upn === upn);
    if (ex) { await S.store.update('pd', ex.id, o, Object.assign({}, ex, o)); Object.assign(ex, o); }
    else { const n = await S.store.create('pd', o); S.pd.items.push(n); if (upn !== S.me.upn) await fuerAndere('pd', n); }
  }, 'Gespeichert');
}

/* ---------- Personalakte ---------- */
const fmtSize = n => n >= 1048576 ? (n / 1048576).toFixed(1).replace('.', ',') + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';
async function loadAkte(upn) {
  const ma = maBy(upn);
  S.akte = { upn, loading: true }; render();
  try {
    if (S.leitung) {
      const st = await S.store.akteStatus();
      if (st.status !== 'ok') { S.akte = { upn, status: st.status }; return render(); }
    }
    if (!ma || !ma.akteId) { S.akte = { upn, status: 'nofolder' }; return render(); }
    S.akte = { upn, status: 'ok', items: await S.store.akteList_(ma) };
  } catch (e) { console.error(e); S.akte = { upn, status: 'error', error: e.message || String(e) }; }
  render();
}
function viewAkte() {
  const upn = S.leitung ? (S.viewUpn || S.me.upn) : S.me.upn, m = maBy(upn), A = S.akte;
  if (!m && !S.leitung) return noMa();
  const sel = S.leitung && S.ma.length ? `<select class="inp" style="width:auto" onchange="S.viewUpn=this.value;loadAkte(this.value)">${S.ma.map(x => `<option value="${esc(x.upn)}" ${x.upn === upn ? 'selected' : ''}>${esc(x.name)}</option>`).join('')}</select>` : '';
  const head = `<div class="card-h"><div><div class="card-t">Personalakte${m ? ' · ' + esc(m.name) : ''}</div><div class="muted sm">${S.leitung ? 'Dokumente ablegen und eingereichte Unterlagen übernehmen' : 'Deine Unterlagen – nur für dich und die Kanzleileitung sichtbar'}</div></div>${sel}</div>`;
  if (!A || A.upn !== upn) { setTimeout(() => loadAkte(upn)); return `<div class="card">${head}<div class="empty">Lade Personalakte …</div></div>`; }
  if (A.loading) return `<div class="card">${head}<div class="empty">Lade Personalakte …</div></div>`;
  if (A.status === 'error') return `<div class="card">${head}<div class="card-b"><div class="banner err">${esc(A.error)}</div><div class="actions"><button class="btn" onclick="loadAkte('${esc(upn)}')">Erneut versuchen</button></div></div></div>`;
  if (A.status === 'missing') return `<div class="card">${head}<div class="card-b">
    <p>Für die Personalakten wird in SharePoint eine eigene Dokumentbibliothek <b>${AKTE.name}</b> angelegt. Pro Mitarbeiter/in entsteht darin ein Ordner mit eigenen Rechten.</p>
    <div class="actions"><button class="btn primary" onclick="akteSetupAction()">Personalakte einrichten</button></div></div></div>`;
  if (A.status === 'unsafe') return `<div class="card">${head}<div class="card-b">
    <div class="banner warn"><b>Bibliothek noch nicht abgesichert.</b> Die Bibliothek ${AKTE.name} ist noch für alle Mitglieder der Website freigegeben. Bevor Akten angelegt werden, bitte in SharePoint einschränken:</div>
    <ol class="steps"><li>SharePoint-Website öffnen → <b>Websiteinhalte</b> → bei <b>${AKTE.name}</b> auf <b>⋯ → Einstellungen</b></li>
    <li><b>Berechtigungen für diese Dokumentbibliothek</b> → <b>Berechtigungsvererbung beenden</b></li>
    <li>Haken bei <b>Mitglieder von …</b> und <b>Besucher von …</b> setzen → <b>Benutzerberechtigungen entfernen</b></li>
    <li>Übrig bleibt nur <b>Besitzer von …</b> (Vollzugriff)</li></ol>
    <div class="actions"><button class="btn primary" onclick="loadAkte('${esc(upn)}')">Erneut prüfen</button></div></div></div>`;
  if (A.status === 'nofolder') return `<div class="card">${head}<div class="card-b">${S.leitung
    ? `<p>Für ${esc(m.name)} ist noch keine Personalakte angelegt. Beim Anlegen entsteht ein Ordner, den nur die Leitung und ${esc(m.name)} sehen, plus ein Eingang zum Einreichen von Unterlagen.</p><div class="actions"><button class="btn primary" onclick="akteFolderAction('${esc(upn)}')">Akte anlegen</button></div>`
    : '<div class="empty">Deine Personalakte ist noch nicht eingerichtet. Die Kanzleileitung legt sie an.</div>'}</div></div>`;
  const items = A.items || [], eingang = items.filter(i => i.ort === 'eingang'), akte = items.filter(i => i.ort === 'akte');
  const kats = AKTE.kategorien, katOf = i => kats.includes(i.kategorie) ? i.kategorie : 'Sonstiges';
  const link = i => `<a class="doclink" href="${esc(i.webUrl)}" target="_blank" rel="noopener" ${String(i.webUrl).startsWith('data:') ? `download="${esc(i.name)}"` : ''}>${svg('akte')}<span>${esc(i.name)}</span></a>`;
  const meta = i => `${i.dokDatum ? de(i.dokDatum) : '—'}`;
  const by = i => `<div class="sm">${esc(i.by)}</div><div class="muted sm">${i.at ? new Date(i.at).toLocaleDateString('de-DE') : ''} · ${fmtSize(i.size)}</div>`;
  const upload = `<div class="card">${head}<div class="card-b"><div class="akte-h" style="padding:0 0 12px">${S.leitung ? 'Dokument ablegen' : 'Dokument einreichen'}</div>
    <label class="drop" id="ak-drop">${svg('upload')}<span id="ak-fname">Datei hier ablegen oder klicken (PDF, Bild, Word · max. ${MAX_UPLOAD_MB} MB)</span>
      <input type="file" id="ak-file" multiple onchange="akteFileChosen(this)"></label>
    <div class="fg2" style="margin-top:18px">
      <div class="f"><label>Kategorie</label><select id="ak-kat">${kats.map(k => `<option>${esc(k)}</option>`).join('')}</select></div>
      <div class="f"><label>Dokumentdatum</label><input type="date" id="ak-datum"></div>
      <div class="f span2"><label>Bemerkung (optional)</label><input id="ak-bem" placeholder="z. B. Nachtrag Arbeitszeit ab 01.01."></div></div>
    ${S.leitung ? '' : '<div class="muted sm" style="margin-top:10px">Eingereichte Dokumente prüft die Kanzleileitung und übernimmt sie in deine Akte.</div>'}
    <div class="actions"><button class="btn primary" onclick="akteUploadAction('${esc(upn)}')">${S.leitung ? 'Hochladen' : 'Einreichen'}</button></div></div></div>`;
  const eingangCard = eingang.length || !S.leitung ? `<div class="card"><div class="card-h"><div class="card-t">${S.leitung ? 'Eingereicht – zu prüfen' : 'Eingereicht'} <span class="muted">(${eingang.length})</span></div></div>
    ${eingang.length ? `<div class="tblwrap"><table class="tbl"><thead><tr><th>Dokument</th><th>Kategorie</th><th>Datum</th><th>Eingereicht</th><th></th></tr></thead><tbody>
    ${eingang.map(i => `<tr><td>${link(i)}${i.bemerkung ? `<div class="muted sm">${esc(i.bemerkung)}</div>` : ''}</td>
      <td>${S.leitung ? `<select class="inp" id="mv-${esc(i.id)}" style="width:auto">${kats.map(k => `<option ${k === katOf(i) ? 'selected' : ''}>${esc(k)}</option>`).join('')}</select>` : esc(katOf(i))}</td>
      <td>${meta(i)}</td><td>${by(i)}</td>
      <td style="text-align:right;white-space:nowrap">${S.leitung ? `<button class="btn sm primary" onclick="akteMoveAction('${esc(upn)}','${esc(i.id)}')">Übernehmen</button> <button class="btn sm danger" onclick="akteDeleteAction('${esc(upn)}','${esc(i.id)}')">Löschen</button>` : '<span class="badge b-warn">wird geprüft</span>'}</td></tr>`).join('')}
    </tbody></table></div>` : '<div class="empty">Nichts eingereicht.</div>'}</div>` : '';
  const sections = kats.map(k => { const fs = akte.filter(i => katOf(i) === k).sort((x, y) => (y.dokDatum || y.at).localeCompare(x.dokDatum || x.at));
    return `<div class="akte-sec"><div class="akte-h">${esc(k)} <span class="muted">(${fs.length})</span></div>${fs.length ? `<div class="tblwrap"><table class="tbl"><tbody>
      ${fs.map(i => `<tr><td>${link(i)}${i.bemerkung ? `<div class="muted sm">${esc(i.bemerkung)}</div>` : ''}</td><td class="hide-m" style="width:110px">${meta(i)}</td><td class="hide-m" style="width:190px">${by(i)}</td>
        <td style="text-align:right;white-space:nowrap;width:1%"><a class="btn sm" href="${esc(i.download || i.webUrl)}" ${String(i.download || i.webUrl).startsWith('data:') ? `download="${esc(i.name)}"` : 'target="_blank" rel="noopener"'}>Download</a>${S.leitung ? ` <button class="btn sm danger" onclick="akteDeleteAction('${esc(upn)}','${esc(i.id)}')">Löschen</button>` : ''}</td></tr>`).join('')}
      </tbody></table></div>` : '<div class="muted sm akte-empty">Keine Dokumente</div>'}</div>`; }).join('');
  return `${upload}${eingangCard}<div class="card"><div class="card-h"><div class="card-t">Dokumente <span class="muted">(${akte.length})</span></div></div><div class="akte-body">${sections}</div></div>`;
}
function akteFileChosen(inp) { const n = [...inp.files].map(f => f.name).join(', '); $('#ak-fname').textContent = n || 'Datei hier ablegen oder klicken'; }
const akteSetupAction = () => act(async () => { await S.store.akteSetup(); S.akte = null; }, 'Bibliothek angelegt');
const akteFolderAction = upn => act(async () => { const m = maBy(upn); Object.assign(m, await S.store.akteEnsureFolder(m)); S.akte = null; }, 'Personalakte angelegt');
function akteUploadAction(upn) {
  const m = maBy(upn), files = [...($('#ak-file').files || [])];
  if (!files.length) return toast('Bitte zuerst eine Datei auswählen', 'err');
  const big = files.find(f => f.size > MAX_UPLOAD_MB * 1048576); if (big) return toast(`${big.name} ist größer als ${MAX_UPLOAD_MB} MB`, 'err');
  const meta = { kategorie: $('#ak-kat').value, dokDatum: $('#ak-datum').value, bemerkung: $('#ak-bem').value.trim() };
  act(async () => { for (const f of files) await S.store.akteUpload(m, S.leitung ? 'akte' : 'eingang', f, meta); S.akte = null; },
    files.length + (files.length === 1 ? ' Dokument ' : ' Dokumente ') + (S.leitung ? 'abgelegt' : 'eingereicht'));
}
function akteMoveAction(upn, id) {
  const m = maBy(upn), it = S.akte.items.find(x => x.id === id), k = $('#mv-' + CSS.escape(id)).value;
  act(async () => { await S.store.akteMove(m, it, k); S.akte = null; }, 'In die Akte übernommen');
}
function akteDeleteAction(upn, id) {
  const m = maBy(upn), it = S.akte.items.find(x => x.id === id);
  if (!confirm(`„${it.name}“ löschen? Die Datei landet im SharePoint-Papierkorb.`)) return;
  act(async () => { await S.store.akteDelete(m, it); S.akte = null; }, 'Gelöscht');
}

/* ---------- Team (Leitung) ---------- */
function viewTeam() {
  const [y, mo] = S.month.split('-').map(Number), t = today();
  const ende = S.month + '-' + Z.pad(new Date(y, mo, 0).getDate()), bis = ende < t ? ende : Z.addDays(t, -1);
  const rows = S.ma.filter(m => !m.austritt || m.austritt >= S.month + '-01').map(m => {
    const data = dataFor(m.upn), mon = Z.monat(m, y, mo, data, CFG, t, nowMin()), st = stempelStatus(m.upn), ul = Z.urlaub(m, y, S.abw, CFG);
    const abwH = Z.abwAm(S.abw, m.upn, t).filter(Z.zaehlt);
    const heute = st.s === 'an' ? `<span class="badge b-ok">da seit ${esc(st.b.beginn)} · ${esc(st.b.ort)}</span>` : st.s === 'pause' ? `<span class="badge b-warn">Pause seit ${esc(st.seit)}</span>`
      : abwH.length ? `<span class="badge b-info">${esc(abwH[0].art)}</span>` : '<span class="badge b-grey">—</span>';
    return { m, mon, heute, saldo: Z.saldo(m, data, CFG, bis), ul, offen: offeneBloecke(m.upn).filter(b => b.datum < t).length };
  });
  const antraege = S.abw.filter(a => a.status === 'beantragt').sort((a, b) => a.von.localeCompare(b.von));
  return `${antraege.length ? `<div class="card"><div class="card-h"><div class="card-t">Offene Anträge (${antraege.length})</div></div><div class="tblwrap"><table class="tbl"><tbody>
    ${antraege.map(a => { const mm = maBy(a.upn); return `<tr><td>${esc(a.name)}</td><td>${esc(a.art)}${a.bemerkung ? `<div class="muted sm">${esc(a.bemerkung)}</div>` : ''}</td><td>${de(a.von)}${a.bis !== a.von ? ' – ' + de(a.bis) : ''}</td><td class="num">${mm ? fmtTage(Z.arbeitstageZaehlen(mm, a.von, a.bis || a.von, CFG, a.halbtag)) : ''} Tg.</td><td style="text-align:right">${abwActions(a)}</td></tr>`; }).join('')}</tbody></table></div></div>` : ''}
  <div class="card"><div class="card-h"><div class="card-t">Team · ${MONATE[mo - 1]} ${y}</div><div class="row"><button class="btn sm" onclick="shiftMonth(-1)">‹</button><button class="btn sm" onclick="shiftMonth(1)">›</button><button class="btn sm" onclick="csvTeam()">Excel (CSV)</button></div></div>
  ${rows.length ? `<div class="tblwrap"><table class="tbl"><thead><tr><th>Mitarbeiter</th><th>Heute</th><th class="num">Soll</th><th class="num">Ist</th><th class="num">Diff. Monat</th><th class="num">Saldo</th><th class="num">Resturlaub</th><th>Prüfung</th></tr></thead><tbody>
  ${rows.map(r => `<tr><td><button class="link" style="font-size:13px" onclick="S.viewUpn='${esc(r.m.upn)}';go('monat')">${esc(r.m.name)}</button></td><td>${r.heute}</td>
    <td class="num">${Z.dur(r.mon.soll)}</td><td class="num">${Z.dur(r.mon.ist + r.mon.gut)}</td><td class="num">${signed(r.mon.diff)}</td><td class="num">${signed(r.saldo)}</td><td class="num">${fmtTage(r.ul.rest)}</td>
    <td>${r.mon.rot ? `<span class="badge b-err">${r.mon.rot} rot</span> ` : ''}${r.mon.gelb ? `<span class="badge b-warn">${r.mon.gelb} gelb</span> ` : ''}${r.offen ? `<span class="badge b-err">${r.offen}× Gehen fehlt</span>` : ''}${!r.mon.rot && !r.mon.gelb && !r.offen ? '<span class="dot green"></span>' : ''}</td></tr>`).join('')}
  </tbody></table></div>` : '<div class="empty">Noch keine Mitarbeiter angelegt.</div>'}</div>`;
}
function csvTeam() {
  const [y, mo] = S.month.split('-').map(Number), t = today(), ende = S.month + '-' + Z.pad(new Date(y, mo, 0).getDate()), bis = ende < t ? ende : Z.addDays(t, -1);
  const q = v => '"' + String(v ?? '').replace(/"/g, '""') + '"', h = n => (Math.round(n / 60 * 100) / 100).toFixed(2).replace('.', ',');
  const lines = S.ma.map(m => { const d = dataFor(m.upn), mon = Z.monat(m, y, mo, d, CFG, t, null), ul = Z.urlaub(m, y, S.abw, CFG);
    return [m.name, h(mon.soll), h(mon.ist), h(mon.gut), h(mon.diff), h(Z.saldo(m, d, CFG, bis)), ul.genommen, ul.rest, mon.rot, mon.gelb].map(q).join(';'); });
  download(`Arbeitszeit_Team_${S.month}.csv`, '﻿' + [['Mitarbeiter', 'Soll (h)', 'Ist (h)', 'Gutschrift (h)', 'Differenz (h)', 'Saldo per ' + de(bis) + ' (h)', 'Urlaub genommen', 'Resturlaub', 'Tage rot', 'Tage gelb'].map(q).join(';'), ...lines].join('\r\n'), 'text/csv;charset=utf-8');
}

/* ---------- Stammdaten & Einrichtung (Leitung) ---------- */
function viewStamm() {
  const setup = `<div class="card"><div class="card-h"><div class="card-t">Einrichtung</div></div><div class="card-b">
    <div class="sm">Modus: <b>${CFG.demo ? 'Demo – Daten nur in diesem Browser' : 'Microsoft 365'}</b>${S.info?.site ? ` · Website: ${esc(S.info.site)}` : ''}</div>
    ${S.missing.length ? `<div class="banner warn" style="margin-top:10px">Fehlende Listen: ${S.missing.map(esc).join(', ')}. Zum Anlegen brauchst du Besitzerrechte auf der SharePoint-Website.</div>
      <div class="actions"><button class="btn primary" onclick="setupLists()">Listen jetzt anlegen</button></div>`
      : `<div class="banner ok" style="margin-top:10px">Einrichtung abgeschlossen – alle drei SharePoint-Listen sind verbunden.</div>
         <div class="muted sm" style="margin-top:8px">Für Admins: In SharePoint müssen bei ZE_Buchungen und ZE_Abwesenheiten „Nur eigene Elemente lesen“ und der Versionsverlauf aktiv bleiben.</div>`}
  </div></div>`;
  if (S.missing.length) return setup;
  return `<div class="card"><div class="card-h"><div class="card-t">Mitarbeiter</div><button class="btn primary sm" onclick="maForm()">+ Neu</button></div>
  ${S.ma.length ? `<div class="tblwrap"><table class="tbl"><thead><tr><th>Name</th><th>Anmeldename</th><th class="num">h/Woche</th><th>Arbeitstage</th><th class="num">Urlaub</th><th>Saldo ab</th><th>Eintritt / Austritt</th><th></th></tr></thead><tbody>
  ${S.ma.map(m => `<tr><td>${esc(m.name)}${m.leitung || (CFG.leitung || []).map(lc).includes(m.upn) || (S.store.isDemoLeitung && S.store.isDemoLeitung(m.upn)) ? ' <span class="badge b-info">Leitung</span>' : ''}</td><td class="sm">${esc(m.upn)}</td><td class="num">${m.wochenstunden ?? ''}</td>
    <td class="sm">${esc(String(m.arbeitstage || '').split(',').map(x => WD[+x] || '').join(' '))}</td><td class="num">${m.urlaubstage ?? ''}${m.restVorjahr ? ' + ' + m.restVorjahr : ''}</td>
    <td class="sm">${de(m.saldoStart)}${m.saldoStunden ? ` (${m.saldoStunden > 0 ? '+' : ''}${String(m.saldoStunden).replace('.', ',')} h)` : ''}</td><td class="sm">${de(m.eintritt)}${m.austritt ? ' / ' + de(m.austritt) : ''}</td>
    <td style="text-align:right"><button class="btn sm" onclick="maForm('${m.id}')">Bearbeiten</button></td></tr>`).join('')}</tbody></table></div>`
  : '<div class="empty">Noch keine Mitarbeiter. Lege zuerst dich selbst an – mit genau der E-Mail-Adresse, mit der du dich bei Microsoft anmeldest.</div>'}
  <div class="card-b muted sm" style="border-top:.5px solid var(--border)">Jahreswechsel: Resturlaub Vorjahr eintragen und „Saldo ab“ auf den 01.01. mit dem Saldo per 31.12. als Startwert setzen. Das hält die Berechnung schnell und schreibt den Stand fest.</div></div>
  ${setup}`;
}
function maForm(id) {
  const m = id ? S.ma.find(x => x.id === id) : { name: '', upn: '', wochenstunden: 40, arbeitstage: '1,2,3,4,5', urlaubstage: 30, restVorjahr: 0, saldoStart: today().slice(0, 8) + '01', saldoStunden: 0, eintritt: '', austritt: '' };
  const at = String(m.arbeitstage || '').split(',').map(Number);
  modal(`<div class="card-h"><div class="card-t">${id ? 'Mitarbeiter bearbeiten' : 'Neuer Mitarbeiter'}</div></div><div class="card-b"><div class="fg2">
    <div class="f"><label>Name</label><input id="mf-name" value="${esc(m.name)}"></div>
    <div class="f"><label>Anmeldename (E-Mail bei Microsoft 365)</label><input id="mf-upn" value="${esc(m.upn)}" ${id ? 'disabled' : ''}></div>
    <div class="f"><label>Wochenstunden</label><input id="mf-wh" type="number" step="0.25" min="0" value="${m.wochenstunden ?? ''}"></div>
    <div class="f"><label>Arbeitstage</label><div class="row" style="padding-top:4px">${[1, 2, 3, 4, 5, 6, 7].map(d => `<label class="chk"><input type="checkbox" class="mf-at" value="${d}" ${at.includes(d) ? 'checked' : ''}>${WD[d]}</label>`).join('')}</div></div>
    <div class="f"><label>Urlaubstage pro Jahr</label><input id="mf-ut" type="number" step="0.5" min="0" value="${m.urlaubstage ?? ''}"></div>
    <div class="f"><label>Resturlaub aus Vorjahr</label><input id="mf-rv" type="number" step="0.5" value="${m.restVorjahr ?? 0}"></div>
    <div class="f"><label>Saldo ab Datum</label><input id="mf-ss" type="date" value="${esc(m.saldoStart)}"></div>
    <div class="f"><label>Saldo-Startwert in Stunden (z. B. 3,5 oder -2)</label><input id="mf-sw" value="${String(m.saldoStunden ?? 0).replace('.', ',')}"></div>
    <div class="f"><label>Eintritt</label><input id="mf-ein" type="date" value="${esc(m.eintritt)}"></div>
    <div class="f"><label>Austritt</label><input id="mf-aus" type="date" value="${esc(m.austritt)}"></div>
    <label class="chk span2"><input type="checkbox" id="mf-leitung" ${m.leitung ? 'checked' : ''}> Leitung (Team-Übersicht, Genehmigungen, Stammdaten) – zusätzlich in SharePoint in die Gruppe „Besitzer“ aufnehmen</label>
  </div><div class="actions"><button class="btn" onclick="closeModal()">Abbrechen</button><button class="btn primary" onclick="saveMa(${id ? `'${id}'` : 'null'})">Speichern</button></div></div>`);
}
function saveMa(id) {
  const v = n => $('#mf-' + n).value.trim();
  const o = { name: v('name'), wochenstunden: +String(v('wh')).replace(',', '.') || 0, arbeitstage: [...document.querySelectorAll('.mf-at:checked')].map(x => x.value).join(','),
    urlaubstage: +String(v('ut')).replace(',', '.') || 0, restVorjahr: +String(v('rv')).replace(',', '.') || 0, saldoStart: v('ss'), saldoStunden: +String(v('sw')).replace(',', '.') || 0, eintritt: v('ein'), austritt: v('aus'), leitung: $('#mf-leitung').checked };
  if (!id) o.upn = lc(v('upn'));
  if (!o.name) return toast('Name fehlt', 'err');
  if (!id && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(o.upn)) return toast('Anmeldename muss eine E-Mail-Adresse sein', 'err');
  if (!id && S.ma.some(m => m.upn === o.upn)) return toast('Für diese Adresse gibt es schon einen Stammsatz', 'err');
  if (!o.arbeitstage) return toast('Mindestens einen Arbeitstag wählen', 'err');
  if (!o.saldoStart) return toast('„Saldo ab Datum“ fehlt', 'err');
  closeModal();
  act(async () => {
    if (id) { const m = S.ma.find(x => x.id === id); await S.store.update('ma', id, o, Object.assign({}, m, o)); Object.assign(m, o); }
    else { const n = await S.store.create('ma', o); S.ma.push(n); S.ma.sort((a, b) => a.name.localeCompare(b.name)); }
    if (o.saldoStart < S.loadedFrom) await ensureLoaded(o.saldoStart);
    S.leitung = !!isLeitung(S.me.upn);
  }, 'Gespeichert');
}
const setupLists = () => act(async () => {
  await S.store.setup(); const st = await S.store.init(); S.info = st; S.missing = st.missing;
  if (!S.missing.length) await loadAll();
  S.leitung = !!isLeitung(S.me.upn);
}, 'Listen angelegt');

/* ---------- Demo ---------- */
function demoUser(upn) { try { localStorage.setItem('ze_demo_user', upn); } catch (e) {} location.reload(); }
function demoReset() { if (!confirm('Demo-Daten zurücksetzen?')) return; S.store.reset(); location.reload(); }

/* ======================= Start ======================= */
function renderLogin(msg) {
  document.body.classList.add('login-mode');
  $('#side').style.display = 'none';
  const ico = (d, k) => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;
  const ms = '<svg viewBox="0 0 21 21" aria-hidden="true"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>';
  const feat = [[ICON.heute, 'Kommen, Pause und Gehen mit einem Klick'], [ICON.monat, 'Monatsnachweis mit Saldo und Prüfung'], [ICON.abw, 'Urlaub beantragen, Abwesenheiten melden']];
  $('#main').innerHTML = `<div class="login">
    <div class="login-ico">${ico(ICON.heute)}</div>
    <h1>Willkommen bei ${esc(CFG.appName || 'Personal')}</h1>
    <p class="login-sub">Arbeitszeit und Abwesenheiten für ${esc(CFG.firma || 'die Kanzlei')}</p>
    <ul class="login-feat">${feat.map(([d, t]) => `<li>${ico(d)}<span>${t}</span></li>`).join('')}</ul>
    ${msg ? `<div class="banner err">${esc(msg)}</div>` : ''}
    <button class="btn primary login-btn" onclick="S.store.loginRedirect()">${ms}<span>Mit Microsoft anmelden</span></button>
    <p class="login-note">${ico('<rect x="5" y="11" width="14" height="9" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/>')}Anmeldung mit dem Kanzlei-Konto. Ihre Daten bleiben im Microsoft 365 der Kanzlei.</p>
    <p class="login-ver">${esc(APP_VERSION)}</p>
  </div>`;
}
function renderFatal(e) {
  console.error(e);
  $('#main').innerHTML = `<div class="card"><div class="card-h"><div class="card-t">Start nicht möglich</div></div><div class="card-b"><div class="banner err">${esc(e.message || e)}</div>
    <p class="muted sm" style="margin-top:10px">Häufige Ursachen: config.js nicht ausgefüllt, SharePoint-Adresse falsch, Administratorzustimmung für die App-Berechtigungen fehlt (Anleitung Schritt 2), keine Berechtigung auf der SharePoint-Website (Schritt 4).</p>
    <div class="actions"><button class="btn" onclick="location.reload()">Erneut versuchen</button>${S.store && S.store.logout && S.me ? '<button class="btn" onclick="S.store.logout()">Abmelden</button>' : ''}</div></div></div>`;
}
function tick() {
  renderTimer();
  const c = $('#clock'); if (c) c.textContent = nowHHMM();
  if (S.view === 'heute' && !$('#modal').innerHTML && !S.busy && new Date().getSeconds() === 0) render();
}
(async () => {
  try {
    S.store = CFG.demo ? new DemoStore() : new GraphStore();
    const me = await S.store.login();
    if (!me) return renderLogin();
    S.me = me; S.leitung = !!isLeitung(me.upn);
    S.info = await S.store.init(); S.missing = S.info.missing;
    if (!S.missing.length) await loadAll();
    S.leitung = !!isLeitung(me.upn);
    if (S.missing.length) S.view = 'stamm';
    S.viewUpn = me.upn;
    render();
    setInterval(tick, 1000);
    if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).catch(() => {});
  } catch (e) { renderFatal(e); }
})();
