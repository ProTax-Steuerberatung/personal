/* Zeiterfassung – Rechenkern (ohne UI, ohne Speicher).
   Alle Zeiten intern in Minuten. Datumswerte als 'JJJJ-MM-TT'. */
(function (root) {
  'use strict';
  const pad = n => String(n).padStart(2, '0');
  const iso = d => d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  const parseIso = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
  const addDays = (s, n) => { const d = parseIso(s); d.setDate(d.getDate() + n); return iso(d); };
  const isoWd = s => { const x = parseIso(s).getDay(); return x === 0 ? 7 : x; }; // 1 = Mo … 7 = So
  const daysBetween = (a, b) => Math.round((parseIso(b) - parseIso(a)) / 864e5);
  function tm(s) { const m = /^(\d{1,2}):(\d{2})$/.exec(String(s || '').trim()); if (!m) return null; const h = +m[1], mi = +m[2]; return h > 23 || mi > 59 ? null : h * 60 + mi; }
  const hhmm = min => pad(Math.floor(min / 60) % 24) + ':' + pad(Math.round(min) % 60);
  function dur(min, signed) {
    if (min == null || isNaN(min)) return '—';
    const neg = min < 0, a = Math.round(Math.abs(min));
    const s = Math.floor(a / 60) + ':' + pad(a % 60);
    return (neg ? '−' : signed && a ? '+' : '') + s;
  }

  /* ---------- Feiertage ---------- */
  function easter(y) {
    const a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25),
      g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4,
      l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451), mon = Math.floor((h + l - 7 * m + 114) / 31),
      day = ((h + l - 7 * m + 114) % 31) + 1;
    return iso(new Date(y, mon - 1, day));
  }
  const _ft = {};
  function feiertage(y, cfg) {
    cfg = cfg || {};
    const key = y + '|' + (cfg.mariaeHimmelfahrt !== false);
    if (_ft[key]) return _ft[key];
    const e = easter(y), f = {};
    f[y + '-01-01'] = 'Neujahr';
    f[y + '-01-06'] = 'Heilige Drei Könige';
    f[addDays(e, -2)] = 'Karfreitag';
    f[addDays(e, 1)] = 'Ostermontag';
    f[y + '-05-01'] = 'Tag der Arbeit';
    f[addDays(e, 39)] = 'Christi Himmelfahrt';
    f[addDays(e, 50)] = 'Pfingstmontag';
    f[addDays(e, 60)] = 'Fronleichnam';
    if (cfg.mariaeHimmelfahrt !== false) f[y + '-08-15'] = 'Mariä Himmelfahrt';
    f[y + '-10-03'] = 'Tag der Deutschen Einheit';
    f[y + '-11-01'] = 'Allerheiligen';
    f[y + '-12-25'] = '1. Weihnachtstag';
    f[y + '-12-26'] = '2. Weihnachtstag';
    return (_ft[key] = f);
  }
  const feiertag = (date, cfg) => feiertage(+date.slice(0, 4), cfg)[date] || '';

  /* ---------- Pausen-Text ---------- */
  // "12:00-12:30, 15:00-15:15"  <->  [{v:'12:00',b:'12:30'}, …]
  function parsePausen(txt) {
    const out = [], err = [];
    String(txt || '').split(/[,;\n]/).map(s => s.trim()).filter(Boolean).forEach(p => {
      const m = /^(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})$/.exec(p);
      if (!m || tm(m[1]) == null || tm(m[2]) == null) err.push(p); else out.push({ v: m[1].padStart(5, '0'), b: m[2].padStart(5, '0') });
    });
    return { pausen: out, fehler: err };
  }
  const pausenText = ps => (ps || []).map(p => p.v + '–' + (p.b || '…')).join(', ');

  /* ---------- Block & Tag ---------- */
  // b: {beginn, ende, pausen:[{v,b}], pauseMin}; nowMin: Minuten seit 0 Uhr, wenn der Block "heute" ist (laufende Erfassung)
  function block(b, nowMin) {
    const beg = tm(b.beginn);
    if (beg == null) return { ok: false, fehler: 'Beginn fehlt' };
    let end = tm(b.ende), laufend = false;
    if (end == null) {
      if (nowMin == null) return { ok: true, beg, end: null, offen: true, work: [], netto: 0, pause: 0, pauseMin: +b.pauseMin || 0 };
      end = nowMin < beg ? nowMin + 1440 : nowMin; laufend = true;
    } else if (end < beg) end += 1440; // über Mitternacht
    const ps = (b.pausen || []).map(p => {
      let v = tm(p.v), e = tm(p.b);
      if (v == null) return null;
      if (v < beg) v += 1440;
      if (e == null) e = laufend ? end : v; else if (e < v) e += 1440;
      v = Math.max(beg, Math.min(end, v)); e = Math.max(beg, Math.min(end, e));
      return e > v ? [v, e] : null;
    }).filter(Boolean).sort((a, c) => a[0] - c[0]);
    // Pausen zusammenführen, falls überlappend
    const merged = [];
    ps.forEach(p => { const l = merged[merged.length - 1]; if (l && p[0] <= l[1]) l[1] = Math.max(l[1], p[1]); else merged.push(p.slice()); });
    const work = []; let cur = beg;
    merged.forEach(p => { if (p[0] > cur) work.push([cur, p[0]]); cur = Math.max(cur, p[1]); });
    if (end > cur) work.push([cur, end]);
    const pauseInt = merged.reduce((s, p) => s + p[1] - p[0], 0);
    const pauseMin = Math.max(0, +b.pauseMin || 0);
    const brutto = end - beg;
    return { ok: true, beg, end, laufend, offen: false, work, pauseInt, pauseMin, pause: pauseInt + pauseMin, netto: Math.max(0, brutto - pauseInt - pauseMin) };
  }

  // blocks: Buchungen eines Mitarbeiters an einem Tag (nicht storniert)
  function tag(blocks, nowMin) {
    const calc = blocks.map(b => ({ b, c: block(b, nowMin) })).filter(x => x.c.ok);
    const offen = calc.filter(x => x.c.offen);
    const fertig = calc.filter(x => !x.c.offen).sort((a, c) => a.c.beg - c.c.beg);
    const ist = fertig.reduce((s, x) => s + x.c.netto, 0);
    const pauschal = fertig.reduce((s, x) => s + x.c.pauseMin, 0);
    // Arbeitsintervalle über alle Blöcke; Unterbrechungen < 15 min zählen nicht als Ruhepause (§ 4 ArbZG)
    const work = fertig.flatMap(x => x.c.work).sort((a, c) => a[0] - c[0]);
    const seg = []; let pauseAnr = 0;
    work.forEach(w => {
      const l = seg[seg.length - 1];
      if (l && w[0] - l[1] < 15) { l[1] = Math.max(l[1], w[1]); }
      else { if (l) pauseAnr += w[0] - l[1]; seg.push(w.slice()); }
    });
    if (pauschal >= 15) pauseAnr += pauschal;
    const pauseGesamt = fertig.reduce((s, x) => s + x.c.pause, 0);
    // längster Abschnitt ohne Ruhepause – nur sicher bestimmbar ohne Pauschalpause
    const maxAbschnitt = pauschal > 0 ? null : seg.reduce((m, s) => Math.max(m, s[1] - s[0]), 0);
    const first = fertig.length ? fertig[0].c.beg : null;
    const last = fertig.length ? Math.max(...fertig.map(x => x.c.end)) : null;
    return { ist, pauseGesamt, pauseAnr, maxAbschnitt, first, last, offen: offen.map(x => x.b), laufend: fertig.some(x => x.c.laufend), bloecke: fertig };
  }

  /* ---------- Soll & Abwesenheit ---------- */
  const GUTSCHRIFT = { 'Urlaub': 1, 'Krank': 1, 'Fortbildung': 1, 'Sonderurlaub': 1, 'Überstundenabbau': 0 };
  const zaehlt = a => a.status === 'genehmigt' || a.status === 'gemeldet';
  function arbeitstage(ma) { return String(ma.arbeitstage || '1,2,3,4,5').split(/[,\s]+/).map(Number).filter(x => x >= 1 && x <= 7); }
  function aktiv(ma, date) { if (ma.eintritt && date < ma.eintritt) return false; if (ma.austritt && date > ma.austritt) return false; return true; }
  function sollTag(ma, date, cfg) {
    if (!aktiv(ma, date)) return 0;
    const at = arbeitstage(ma); if (!at.length || !at.includes(isoWd(date))) return 0;
    if (feiertag(date, cfg)) return 0;
    let s = (+ma.wochenstunden || 0) * 60 / at.length;
    if ((cfg && cfg.halbeTage || []).includes(date.slice(5))) s /= 2;
    return Math.round(s);
  }
  function abwAm(abw, upn, date) { return abw.filter(a => a.upn === upn && a.status !== 'storniert' && a.status !== 'abgelehnt' && a.von <= date && date <= (a.bis || a.von)); }

  /* ---------- Prüfungen nach ArbZG ---------- */
  // ctx: {date, t (tag()), soll, abw[], heute(bool), vortagEnde (abs. Minuten bezogen auf 0 Uhr dieses Tages, z. B. -60 = 23:00 Vortag), cfg, blocks}
  function pruefe(ctx) {
    const c = [], t = ctx.t, E = x => c.push({ l: 'err', t: x }), W = x => c.push({ l: 'warn', t: x }), I = x => c.push({ l: 'info', t: x });
    if (t.offen.length && !ctx.heute) E('Gehen fehlt – Buchung ist nicht abgeschlossen');
    if (!ctx.heute || !t.laufend) {
      if (t.ist > 600) E(`Arbeitszeit ${dur(t.ist)} h über 10 h (§ 3 ArbZG)`);
      else if (t.ist > 480) W(`Arbeitszeit ${dur(t.ist)} h über 8 h – Ausgleich im 6-Monats-Schnitt nötig`);
      if (t.ist > 540 && t.pauseAnr < 45) E(`Pause ${dur(t.pauseAnr)} h zu kurz – bei über 9 h mind. 45 min (§ 4 ArbZG)`);
      else if (t.ist > 360 && t.pauseAnr < 30) E(`Pause ${dur(t.pauseAnr)} h zu kurz – bei über 6 h mind. 30 min (§ 4 ArbZG)`);
      if (t.maxAbschnitt != null && t.maxAbschnitt > 360) E(`${dur(t.maxAbschnitt)} h am Stück ohne Pause – max. 6 h (§ 4 ArbZG)`);
    }
    if (t.first != null && ctx.vortagEnde != null) {
      const ruhe = t.first - ctx.vortagEnde;
      if (ruhe < 660) E(`Ruhezeit nur ${dur(ruhe)} h seit Arbeitsende am Vortag – mind. 11 h (§ 5 ArbZG)`);
    }
    if (t.bloecke.length || t.offen.length) {
      if (isoWd(ctx.date) === 7) W('Arbeit an einem Sonntag (§ 9 ArbZG) – Ausnahme dokumentieren');
      const ft = feiertag(ctx.date, ctx.cfg); if (ft) W(`Arbeit am Feiertag (${ft}) – Ausnahme dokumentieren`);
      const urlaub = (ctx.abw || []).filter(a => zaehlt(a) && GUTSCHRIFT[a.art] && !a.halbtag);
      if (urlaub.length) W(`Arbeitszeit erfasst trotz ganztägiger Abwesenheit (${urlaub.map(a => a.art).join(', ')})`);
    }
    // letzter Eintrag aus dem Begründungs-Verlauf ("TT.MM.JJJJ HH:MM Name: Text | …") ohne Stempel
    const letzterGrund = g => { const s = String(g || '').split(' | ').pop() || ''; const i = s.indexOf(': '); return (i > 0 && /^\d{2}\.\d{2}\.\d{4} /.test(s) ? s.slice(i + 2) : s).trim(); };
    (ctx.blocks || []).filter(b => b.nachtrag).forEach(b => I(`Nachtrag ${b.beginn}${b.ende ? '–' + b.ende : ''} – Grund: ${letzterGrund(b.grund) || 'nicht angegeben'}`));
    (ctx.blocks || []).filter(b => b.korrAm).forEach(b => I(`Korrigiert von ${b.korrVon || '?'} am ${new Date(b.korrAm).toLocaleString('de-DE')} – Grund: ${letzterGrund(b.grund) || 'nicht angegeben'}`));
    return c;
  }
  const ampel = checks => checks.some(x => x.l === 'err') ? 'red' : checks.some(x => x.l === 'warn') ? 'yellow' : 'green';

  /* ---------- Auswertungen ---------- */
  // data: {buchungen[], abw[]} – nur Einträge dieses Mitarbeiters oder alle (wird gefiltert)
  function tagesZeile(ma, date, data, cfg, heuteIso, nowMin) {
    const blocks = data.buchungen.filter(b => b.upn === ma.upn && b.datum === date && b.status !== 'storniert');
    const heute = date === heuteIso;
    const t = tag(blocks, heute ? nowMin : null);
    const soll = sollTag(ma, date, cfg);
    const abw = abwAm(data.abw, ma.upn, date);
    let gut = 0;
    abw.filter(zaehlt).forEach(a => { gut = Math.max(gut, Math.round(soll * (GUTSCHRIFT[a.art] || 0) * (a.halbtag ? 0.5 : 1))); });
    // Ruhezeit: Arbeitsende des letzten Arbeitstags innerhalb der letzten 2 Tage
    let vortagEnde = null;
    for (let i = 1; i <= 2 && vortagEnde == null; i++) {
      const d = addDays(date, -i);
      const tv = tag(data.buchungen.filter(b => b.upn === ma.upn && b.datum === d && b.status !== 'storniert'), null);
      if (tv.last != null) vortagEnde = tv.last - 1440 * i;
    }
    const checks = pruefe({ date, t, soll, abw, heute, vortagEnde, cfg, blocks });
    const future = date > heuteIso;
    const diff = future ? 0 : t.ist + gut - soll;
    return { date, wd: isoWd(date), feiertag: feiertag(date, cfg), soll, ist: t.ist, gut, diff, future, abw, blocks, t, checks, ampel: ampel(checks) };
  }
  function monat(ma, y, m, data, cfg, heuteIso, nowMin) {
    const rows = [], first = `${y}-${pad(m)}-01`;
    for (let d = first; d.slice(0, 7) === first.slice(0, 7); d = addDays(d, 1)) rows.push(tagesZeile(ma, d, data, cfg, heuteIso, nowMin));
    const sum = k => rows.reduce((s, r) => s + (r.future ? 0 : r[k]), 0);
    return { rows, soll: sum('soll'), ist: sum('ist'), gut: sum('gut'), diff: sum('diff'), rot: rows.filter(r => r.ampel === 'red').length, gelb: rows.filter(r => r.ampel === 'yellow').length };
  }
  // Saldo aller Tage von Saldo-Startdatum bis einschließlich 'bis'
  function saldo(ma, data, cfg, bis) {
    const start = ma.saldoStart || ma.eintritt; let s = Math.round((+ma.saldoStunden || 0) * 60);
    if (!start || start > bis) return s;
    for (let d = start; d <= bis; d = addDays(d, 1)) {
      const r = tagesZeile(ma, d, data, cfg, '9999-12-31', null);
      s += r.diff;
    }
    return s;
  }
  // Urlaubstage eines Jahres
  function urlaub(ma, y, abw, cfg) {
    const res = { anspruch: (+ma.urlaubstage || 0) + (+ma.restVorjahr || 0), genommen: 0, beantragt: 0 };
    abw.filter(a => a.upn === ma.upn && a.art === 'Urlaub' && (a.status === 'genehmigt' || a.status === 'beantragt')).forEach(a => {
      let tage = 0;
      for (let d = a.von; d <= (a.bis || a.von); d = addDays(d, 1)) if (d.slice(0, 4) === String(y) && sollTag(ma, d, cfg) > 0) tage += a.halbtag ? 0.5 : 1;
      res[a.status === 'genehmigt' ? 'genommen' : 'beantragt'] += tage;
    });
    res.rest = res.anspruch - res.genommen - res.beantragt;
    return res;
  }
  function arbeitstageZaehlen(ma, von, bis, cfg, halbtag) {
    let n = 0; for (let d = von; d <= bis; d = addDays(d, 1)) if (sollTag(ma, d, cfg) > 0) n += halbtag ? 0.5 : 1; return n;
  }

  const api = { pad, iso, parseIso, addDays, isoWd, daysBetween, tm, hhmm, dur, easter, feiertage, feiertag, parsePausen, pausenText, block, tag, sollTag, abwAm, pruefe, ampel, tagesZeile, monat, saldo, urlaub, arbeitstageZaehlen, GUTSCHRIFT, zaehlt };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.ZC = api;
})(this);
