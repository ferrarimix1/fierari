
'use strict';
/* Facturi Fierari: splits the house's utility bills between the four apartments.
   The page renders from the JSON in #ff-data. People who can edit the page save a month by
   publishing the page again with new JSON (the artifact capability); every open view then
   reloads to the new version, so everyone sees the same figures. */

var PAGE_HEAD = '<title>Facturi Fierari</title>\n' +
  '<meta name="claude:viewport-fit-sides" content="cover">\n' +
  '<link rel="preconnect" href="https://fonts.googleapis.com">\n' +
  '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>\n' +
  '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Archivo:wdth,wght@75..125,500..800&amp;family=IBM+Plex+Mono:wght@500;600&amp;family=Inter:wght@400;500;600;700&amp;display=swap">\n';
var PAGE_BODY = '<div id="app"></div>\n';
var SCRIPT_CLOSE = '</scr' + 'ipt>';
var FALLBACK_RESET = ':root{color-scheme:light;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}body{margin:0;font:14px system-ui,-apple-system,sans-serif}img{max-width:100%}[hidden]{display:none!important}';
var XLSX_URLS = ['https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'];

/* gunoi is a fixed amount per apartment; the others are whole bills */
var BILLS = ['gaz', 'digi', 'gunoi', 'apa', 'curent'];
var APTS = [0, 1, 2, 3];
var METERS = [0, 1, 2, 3, 4];
var LIMIT_BILL = 100000;
var LIMIT_KWH = 10000;
var MAX_EL = 6;

/* ------------------------------------------------------------------ page source */

function pageContent(css, dataJson, js) {
  return PAGE_HEAD +
    '<style id="ff-style">' + css + '</style>\n' +
    PAGE_BODY +
    '<script type="application/json" id="ff-data">' + dataJson + SCRIPT_CLOSE + '\n' +
    '<script id="ff-app">' + js + SCRIPT_CLOSE + '\n';
}

function fullDocument(reset, content) {
  return '<!doctype html><html><head><meta charset=utf8><meta name=viewport content="width=device-width,initial-scale=1,viewport-fit=cover"><style>' +
    reset + '</style></head><body>' + content + '</body></html>';
}

function num(v) {
  return typeof v === 'number' && isFinite(v);
}

/* Price per kWh of a month's electricity bills: what they cost in total over the kWh they bill. */
function elPrice(el) {
  if (!el || !el.length) return null;
  var L = sum(el.map(function (b) { return b.lei; })), K = sum(el.map(function (b) { return b.kwh; }));
  if (!(K > 0) || !(L / K > 0)) return null;
  return L / K;
}

function normalizeMonth(m) {
  var o = {};
  /* The record can begin with the starting meter readings only: no bills, nothing to pay. */
  if (m.start) {
    o.start = true;
    o.r = (m.r || []).slice(0, 5).map(Number);
    o.p = (m.p || []).slice(0, 4).map(Number);
    if (m.t) o.t = String(m.t);
    return o;
  }
  BILLS.forEach(function (b) {
    if (b === 'gunoi' && m.gunoi == null) return;
    o[b] = Number(m[b]);
    if (m[b + 'Text']) o[b + 'Text'] = String(m[b + 'Text']);
  });
  /* el = the electricity bills ({lei, kwh}), pk = the price per kWh the apartments pay (from those
     bills, or carried over from an earlier month when there were none). */
  if (Array.isArray(m.el)) {
    var el = m.el.filter(function (b) { return b && b.lei != null && b.kwh != null && num(Number(b.lei)) && num(Number(b.kwh)); })
      .slice(0, MAX_EL).map(function (b) { return { lei: Number(b.lei), kwh: Number(b.kwh) }; });
    if (el.length) {
      o.el = el;
      o.curent = round2(sum(el.map(function (b) { return b.lei; })));
      delete o.curentText;
    }
  }
  var pk = elPrice(o.el);
  if (pk == null && m.pk != null && num(Number(m.pk)) && Number(m.pk) > 0) pk = Number(m.pk);
  if (pk != null) o.pk = pk;
  o.r = (m.r || []).slice(0, 5).map(Number);
  if (m.kwh && typeof m.kwh === 'object') {
    var keys = Object.keys(m.kwh).filter(function (k) { return /^[0-4]$/.test(k) && num(Number(m.kwh[k])); }).sort();
    if (keys.length) {
      o.kwh = {};
      keys.forEach(function (k) { o.kwh[k] = Number(m.kwh[k]); });
    }
  }
  o.p = (m.p || []).slice(0, 4).map(Number);
  if (Array.isArray(m.g) && m.g.some(function (v) { return v != null && v !== ''; })) {
    o.g = m.g.slice(0, 3).map(function (v) { return v == null || v === '' ? null : Number(v); });
  }
  if (Array.isArray(m.manual)) o.manual = m.manual.slice(0, 4).map(Number);
  if (m.note) o.note = String(m.note).slice(0, 200);
  if (m.t) o.t = String(m.t);
  return o;
}

function parseYm(s) {
  var m = /^(\d{4})-(\d{2})$/.exec(String(s || ''));
  if (!m || +m[2] < 1 || +m[2] > 12) return null;
  return { y: +m[1], m: +m[2] - 1 };
}

/* Calendar months left out of the month names (the Excel has no March 2023). */
function normalizeSkip(a) {
  var seen = {};
  return (Array.isArray(a) ? a : []).map(String).filter(function (s) {
    if (!parseYm(s) || seen[s]) return false;
    seen[s] = true;
    return true;
  }).sort().slice(0, 24);
}

function normalizeData(d) {
  d = d || {};
  var s = d.settings || {};
  var surfaces = Array.isArray(s.surfaces) && s.surfaces.length === 4 ? s.surfaces.map(Number) : [25, 25, 25, 25];
  return {
    v: 1,
    settings: { start: /^\d{4}-\d{2}$/.test(s.start || '') ? s.start : '2022-09', surfaces: surfaces, skip: normalizeSkip(s.skip) },
    updated: d.updated || null,
    months: (Array.isArray(d.months) ? d.months : []).map(normalizeMonth)
  };
}

function escJson(s) {
  return s.replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
}

function serializeData(d) {
  var n = normalizeData(d);
  var head = JSON.stringify({ v: n.v, settings: n.settings, updated: n.updated });
  var body = n.months.map(function (m) { return JSON.stringify(m); }).join(',\n');
  return escJson(head.slice(0, -1) + ',"months":[\n' + body + '\n]}');
}

/* ------------------------------------------------------------------ months & money */

function parseStart(s) {
  var m = /^(\d{4})-(\d{2})$/.exec(String(s || ''));
  if (!m) return { y: 2022, m: 8 };
  return { y: +m[1], m: Math.min(11, Math.max(0, +m[2] - 1)) };
}

function monthAt(settings, i) {
  var s = parseStart(settings.start);
  var t0 = s.y * 12 + s.m, t = t0 + i;
  var skips = (settings.skip || []).map(parseYm).filter(Boolean)
    .map(function (d) { return d.y * 12 + d.m; }).sort(function (a, b) { return a - b; });
  for (var k = 0; k < skips.length; k++) if (skips[k] > t0 && skips[k] <= t) t++;
  return { y: Math.floor(t / 12), m: ((t % 12) + 12) % 12 };
}

function sum(a) {
  var s = 0;
  for (var k = 0; k < a.length; k++) s += a[k];
  return s;
}

function round2(v) {
  return Math.round(v * 100) / 100;
}

function shares(surfaces) {
  var t = sum(surfaces);
  return surfaces.map(function (v) { return t > 0 ? v / t : 0; });
}

/* Rounds the four amounts to the ban so that they add up exactly to the bills' total. */
function splitCents(pay, totalCents) {
  var raw = pay.map(function (v) { return v * 100; });
  var fl = raw.map(function (v) { return Math.floor(v + 1e-7); });
  var k = Math.round(totalCents - sum(fl));
  if (!(k >= 0 && k <= pay.length)) return raw.map(function (v) { return Math.round(v); });
  var order = raw.map(function (v, i) { return { i: i, r: v - fl[i] }; })
    .sort(function (a, b) { return (b.r - a.r) || (a.i - b.i); });
  var out = fl.slice();
  for (var n = 0; n < k; n++) out[order[n].i] += 1;
  return out;
}

function computeMonth(data, i) {
  var m = data.months[i];
  var sh = shares(data.settings.surfaces);
  if (m.start) return { i: i, start: true, manual: false, shares: sh, cents: [NaN, NaN, NaN, NaN], total: NaN };
  var gunoi = num(m.gunoi) ? m.gunoi : 0;
  var billTotal = round2(m.gaz + m.digi + gunoi * APTS.length + m.apa + m.curent);
  if (i === 0 || m.manual) {
    var man = (m.manual || [0, 0, 0, 0]).map(Number);
    var mc = man.map(function (v) { return Math.round(v * 100); });
    return { i: i, manual: true, billTotal: billTotal, shares: sh, pay: man, cents: mc, total: sum(mc) / 100 };
  }
  var prev = data.months[i - 1];
  var kwh = METERS.map(function (j) {
    return m.kwh && m.kwh[j] != null ? m.kwh[j] : m.r[j] - prev.r[j];
  });
  var K = sum(kwh);
  var P = sum(m.p);
  var kwhPay = APTS.map(function (a) { return kwh[a] + sh[a] * kwh[4]; });
  /* With a price per kWh (November 2026 on), each apartment pays its kWh at that price. Before,
     the whole electricity bill was shared out in proportion to the kWh. */
  var byPrice = num(m.pk);
  var parts = {
    gaz: sh.map(function (s) { return m.gaz * s; }),
    digi: m.p.map(function (p) { return P > 0 ? m.digi * p / P : NaN; }),
    gunoi: APTS.map(function () { return gunoi; }),
    apa: m.p.map(function (p) { return P > 0 ? m.apa * p / P : NaN; }),
    curent: kwhPay.map(function (k) { return byPrice ? m.pk * k : K === 0 ? 0 : m.curent * k / K; })
  };
  var elDue = byPrice ? m.pk * K : m.curent;
  var due = m.gaz + m.digi + gunoi * APTS.length + m.apa + elDue;
  var pay = APTS.map(function (a) { return parts.gaz[a] + parts.digi[a] + parts.gunoi[a] + parts.apa[a] + parts.curent[a]; });
  var ok = pay.every(num);
  var cents = ok ? splitCents(pay, Math.round(due * 100)) : [NaN, NaN, NaN, NaN];
  return {
    i: i, manual: false, byPrice: byPrice, pk: byPrice ? m.pk : null, elDue: elDue, due: round2(due), billTotal: billTotal,
    shares: sh, kwh: kwh, K: K, P: P, kwhPay: kwhPay, parts: parts, pay: pay, cents: cents, total: sum(cents) / 100
  };
}

function lastPrice(data, i) {
  for (var k = Math.min(i, data.months.length) - 1; k >= 0; k--) {
    if (num(data.months[k].pk)) return { pk: data.months[k].pk, i: k };
  }
  return null;
}

function computeAll(data) {
  return data.months.map(function (_, i) { return computeMonth(data, i); });
}

/* ------------------------------------------------------------------ input parsing */

/* Reads "269,4", "269.4", "150+120" or "88+4*22". Decimals may use a comma or a point.
   kind: "money" (sums allowed) or "plain" (a single number). */
function parseExpr(text, kind) {
  var src = String(text == null ? '' : text).trim();
  if (!src) return { empty: true };
  var s = src.replace(/\s+/g, '').replace(/[×xX]/g, '*').replace(/[÷:]/g, '/').replace(/[−–]/g, '-');
  var pos = 0, binOps = 0, err = null;
  function peek() { return s.charAt(pos); }
  function number() {
    var m = /^(\d+)(?:[.,](\d+))?/.exec(s.slice(pos));
    if (!m) { err = err || 'number'; return NaN; }
    if (m[2] && m[2].length === 3) err = err || 'thousands';
    if (/^[.,]\d/.test(s.slice(pos + m[0].length))) err = err || 'thousands';
    pos += m[0].length;
    return parseFloat(m[1] + (m[2] ? '.' + m[2] : ''));
  }
  function factor() {
    var c = peek();
    if (c === '-') { pos++; return -factor(); }
    if (c === '+') { pos++; return factor(); }
    if (c === '(') {
      pos++;
      var v = expr();
      if (peek() !== ')') { err = err || 'number'; return NaN; }
      pos++;
      binOps++;
      return v;
    }
    return number();
  }
  function term() {
    var v = factor();
    while (!err && (peek() === '*' || peek() === '/')) {
      var op = s.charAt(pos++);
      binOps++;
      var w = factor();
      if (op === '*') v = v * w;
      else if (w === 0) { err = err || 'number'; v = NaN; }
      else v = v / w;
    }
    return v;
  }
  function expr() {
    var v = term();
    while (!err && (peek() === '+' || peek() === '-')) {
      var op = s.charAt(pos++);
      binOps++;
      var w = term();
      v = op === '+' ? v + w : v - w;
    }
    return v;
  }
  var v = expr();
  if (!err && pos !== s.length) err = 'number';
  if (!err && !num(v)) err = 'number';
  if (err) return { error: err };
  if (kind === 'plain' && binOps > 0) return { error: 'plain' };
  return { value: Math.round(v * 1e9) / 1e9, isExpr: binOps > 0, text: s };
}

function recentMax(data, i, get) {
  var vals = [];
  for (var k = Math.max(0, i - 12); k < i; k++) {
    var v = get(k);
    if (num(v)) vals.push(Math.abs(v));
  }
  return vals.length >= 3 ? Math.max.apply(null, vals) : null;
}

/* Checks a month typed into the form. Returns field errors and warnings, form-wide errors,
   the parsed month when everything is valid, and its computed split. */
function validateDraft(data, i, d) {
  var errors = {}, warnings = {}, global = [];
  var prev = i > 0 ? data.months[i - 1] : null;
  var next = i + 1 < data.months.length ? data.months[i + 1] : null;
  var out = {};
  var usage = [null, null, null, null, null];
  var comps = i > 0 ? computeAll(data) : [];
  var elMode = !!d.elMode;
  var elInfo = null;

  /* months recorded before rubbish had its own line (it was inside DIGI) may leave it empty */
  var oldStyle = i < data.months.length && !data.months[i].start && !num(data.months[i].gunoi);
  BILLS.forEach(function (b) {
    if (b === 'curent' && elMode) return;
    var r = parseExpr(d[b], 'money');
    if (r.empty && b === 'gunoi' && oldStyle) return;
    if (r.empty) errors[b] = { code: 'required' };
    else if (r.error) errors[b] = { code: r.error === 'plain' ? 'number' : r.error };
    else if (Math.abs(r.value) > LIMIT_BILL) errors[b] = { code: 'tooBig' };
    else {
      out[b] = r.value;
      if (r.isExpr) out[b + 'Text'] = r.text;
      var mx = recentMax(data, i, function (k) { return data.months[k][b]; });
      if (mx != null && Math.abs(r.value) > 2.5 * mx && Math.abs(r.value) > 50) warnings[b] = { code: 'billHigh', v: mx };
    }
  });

  /* Electricity bills: each row is one bill's total (lei) and the kWh it bills. Empty rows are
     skipped; with no bill at all the last month's price carries over. */
  if (elMode) {
    var rows = [], rowsOk = true;
    (d.el || []).forEach(function (row, n) {
      var l = parseExpr(row && row.l, 'money'), k = parseExpr(row && row.k, 'money');
      if (l.empty && k.empty) return;
      var good = true;
      [['l', l, LIMIT_BILL], ['k', k, LIMIT_BILL]].forEach(function (x) {
        var key = 'el' + n + x[0], r = x[1];
        if (r.empty) { errors[key] = { code: 'required' }; good = false; }
        else if (r.error) { errors[key] = { code: r.error === 'plain' ? 'number' : r.error }; good = false; }
        else if (Math.abs(r.value) > x[2]) { errors[key] = { code: 'tooBig' }; good = false; }
      });
      if (good) rows.push({ lei: l.value, kwh: k.value });
      else rowsOk = false;
    });
    var lp = lastPrice(data, i);
    elInfo = { rows: rows, pk: null, from: null, L: null, K: null };
    if (rowsOk && rows.length) {
      var L = sum(rows.map(function (b) { return b.lei; })), Kb = sum(rows.map(function (b) { return b.kwh; }));
      elInfo.L = L;
      elInfo.K = Kb;
      if (!(Kb > 0)) global.push({ code: 'elKwh' });
      else if (!(L / Kb > 0)) global.push({ code: 'elPrice' });
      else {
        elInfo.pk = L / Kb;
        out.curent = round2(L);
        out.el = rows;
        out.pk = elInfo.pk;
        if (lp && (elInfo.pk > lp.pk * 1.5 || elInfo.pk < lp.pk / 1.5)) warnings.el = { code: 'priceOff', v: lp.pk };
      }
    } else if (rowsOk) {
      if (lp) {
        elInfo.pk = lp.pk;
        elInfo.from = lp.i;
        out.curent = 0;
        out.pk = lp.pk;
      } else global.push({ code: 'elNeeded' });
    }
  }

  out.r = [];
  var kwhOverride = {};
  var readingsOk = true;
  METERS.forEach(function (j) {
    var r = parseExpr(d.r[j], 'plain');
    if (r.empty) { errors['r' + j] = { code: 'required' }; readingsOk = false; return; }
    if (r.error) { errors['r' + j] = { code: r.error === 'plain' ? 'plain' : r.error }; readingsOk = false; return; }
    if (r.value < 0) { errors['r' + j] = { code: 'negative' }; readingsOk = false; return; }
    out.r[j] = r.value;
    if (d.x[j]) {
      var o = parseExpr(d.o[j], 'plain');
      if (o.empty || o.error || o.value < 0) { errors['o' + j] = { code: 'override' }; readingsOk = false; }
      else if (o.value > LIMIT_KWH) { errors['o' + j] = { code: 'kwhBig' }; readingsOk = false; }
      else { kwhOverride[j] = o.value; usage[j] = o.value; }
    } else if (prev) {
      var c = Math.round((r.value - prev.r[j]) * 1e6) / 1e6;
      if (c < 0) { errors['r' + j] = { code: 'lower', v: prev.r[j] }; readingsOk = false; }
      else if (c > LIMIT_KWH) { errors['r' + j] = { code: 'kwhBig' }; readingsOk = false; }
      else {
        usage[j] = c;
        var mx = recentMax(data, i, function (k) { return k > 0 && comps[k] && comps[k].kwh ? comps[k].kwh[j] : NaN; });
        if (mx != null && c > 3 * mx && c > 30) warnings['r' + j] = { code: 'kwhHigh', v: mx };
      }
    }
  });
  if (Object.keys(kwhOverride).length) out.kwh = kwhOverride;

  out.p = [];
  var peopleOk = true;
  APTS.forEach(function (a) {
    var t = String(d.p[a] == null ? '' : d.p[a]).trim();
    if (!/^\d{1,2}$/.test(t) || +t > 20) { errors['p' + a] = { code: 'persons' }; peopleOk = false; return; }
    out.p[a] = +t;
  });
  if (peopleOk && sum(out.p) === 0) { global.push({ code: 'personsZero' }); peopleOk = false; }

  var g = [null, null, null], anyG = false;
  [0, 1, 2].forEach(function (j) {
    var r = parseExpr(d.g[j], 'plain');
    if (r.empty) return;
    if (r.error || r.value < 0) { errors['g' + j] = { code: r.error === 'thousands' ? 'thousands' : 'plain' }; return; }
    g[j] = r.value;
    anyG = true;
  });
  if (anyG) out.g = g;

  var note = String(d.note || '').trim();
  if (note.length > 200) errors.note = { code: 'note' };
  else if (note) out.note = note;

  if (!elMode && prev && readingsOk && !errors.curent) {
    var K = sum(METERS.map(function (j) { return usage[j]; }));
    if (K === 0 && out.curent !== 0) global.push({ code: 'kwhZero' });
  }
  if (next && readingsOk) {
    METERS.forEach(function (j) {
      if (next.kwh && next.kwh[j] != null) return;
      if (next.r[j] - out.r[j] < 0) global.push({ code: 'nextLower', i: i + 1, j: j });
    });
  }

  var valid = !Object.keys(errors).length && !global.length;
  var month = null, comp = null;
  if (valid) {
    month = normalizeMonth(out);
    var tmp = { settings: data.settings, months: data.months.slice(0, i).concat([month]) };
    comp = i > 0 ? computeMonth(tmp, i) : null;
    if (comp && !comp.cents.every(num)) { valid = false; month = null; comp = null; global.push({ code: 'personsZero' }); }
  }
  return { errors: errors, warnings: warnings, global: global, usage: usage, el: elInfo, month: month, comp: comp };
}

/* Checks the starting readings (and people) typed into their form. */
function validateStart(data, d) {
  var errors = {}, global = [], out = { start: true, r: [], p: [] };
  var readingsOk = true, peopleOk = true;
  METERS.forEach(function (j) {
    var r = parseExpr(d.r[j], 'plain');
    if (r.empty) errors['r' + j] = { code: 'required' };
    else if (r.error) errors['r' + j] = { code: r.error === 'plain' ? 'plain' : r.error };
    else if (r.value < 0) errors['r' + j] = { code: 'negative' };
    else { out.r[j] = r.value; return; }
    readingsOk = false;
  });
  APTS.forEach(function (a) {
    var t = String(d.p[a] == null ? '' : d.p[a]).trim();
    if (!/^\d{1,2}$/.test(t) || +t > 20) { errors['p' + a] = { code: 'persons' }; peopleOk = false; return; }
    out.p[a] = +t;
  });
  if (peopleOk && sum(out.p) === 0) global.push({ code: 'personsZero' });
  var next = data.months[1];
  if (next && readingsOk) {
    METERS.forEach(function (j) {
      if (next.kwh && next.kwh[j] != null) return;
      if (next.r[j] - out.r[j] < 0) global.push({ code: 'nextLower', i: 1, j: j });
    });
  }
  var valid = !Object.keys(errors).length && !global.length;
  return { errors: errors, warnings: {}, global: global, usage: [null, null, null, null, null], el: null, month: valid ? normalizeMonth(out) : null, comp: null };
}

function validateSettings(d) {
  var errors = {}, surfaces = [];
  APTS.forEach(function (a) {
    var r = parseExpr(d.s[a], 'plain');
    if (r.empty || r.error || !(r.value > 0)) errors['s' + a] = { code: 'positive' };
    else surfaces[a] = r.value;
  });
  var y = +d.sy, mo = +d.sm;
  if (!(y >= 2000 && y <= 2100) || !(mo >= 0 && mo <= 11)) errors.sy = { code: 'number' };
  var ok = !Object.keys(errors).length;
  return {
    errors: errors,
    settings: ok ? { start: y + '-' + String(mo + 1).padStart(2, '0'), surfaces: surfaces } : null
  };
}

/* ------------------------------------------------------------------ text */

var I18N = {
  ro: {
    locale: 'ro-RO',
    tagline: 'Utilitățile casei, împărțite pe apartamente',
    lang: 'Limba',
    myApt: 'Apartamentul meu',
    all: 'Toate',
    apt: 'Ap. {n}',
    you: 'tu',
    prevMonth: 'Luna anterioară',
    nextMonth: 'Luna următoare',
    stOk: 'Calculat',
    stStart: 'Lună de start',
    stBase: 'Citiri de pornire',
    baseNote: 'Evidența pornește de la aceste citiri ale contoarelor. Prima lună de completat: {m}.',
    editBase: 'Editează citirile',
    fBase: 'Citirile de pornire, {m}',
    baseSave: 'Salvează citirile',
    baseT: 'Citirile de pornire au fost salvate.',
    histEmpty: 'Încă nu e nicio lună calculată. După ce adaugi prima lună, sumele apar aici.',
    add: 'Adaugă {m}',
    editorsOnly: 'Doar cei care pot modifica pagina văd acest buton.',
    edit: 'Editează',
    totalBills: 'Total facturi',
    totalPaid: 'Total plătit',
    totalDue: 'Total de plată',
    sumOk: 'Sumele pe apartamente adună exact totalul facturilor.',
    sumOkDue: 'Sumele pe apartamente adună exact totalul de plată.',
    manualNote: 'Prima lună din evidență: sumele au fost stabilite manual, nu calculate. Facturile lunii au totalizat {v}.',
    billsTitle: 'Facturile lunii',
    dueTitle: 'Ce se împarte luna aceasta',
    gaz: 'Gaz',
    digi: 'DIGI',
    gunoi: 'Gunoi',
    gunoiField: 'Gunoi (lei pe apartament)',
    gunoiExpr: '{n} × {v}',
    byFixed: 'aceeași sumă pe apartament',
    lineFixed: 'sumă fixă pe apartament',
    apa: 'Apă',
    curent: 'Curent (Sus + Jos)',
    curentP: 'Curent',
    byArea: 'după suprafață',
    byPeople: 'după persoane',
    byKwh: 'proporțional cu kWh',
    byPrice: 'kWh × preț',
    kwhTimes: '{k} kWh × {p}',
    priceTitle: 'Prețul curentului',
    billN: 'Factura {n}',
    priceCalc: '{L} ÷ {K} kWh = {p}.',
    priceUse: 'La acest preț, fiecare apartament plătește consumul propriului contor plus partea sa din contorul comun.',
    priceCarried: 'Nicio factură de curent luna aceasta: se folosește prețul din {m}.',
    howTitle: 'Cum s-a calculat',
    howFor: 'Cum s-a ajuns la {v} pentru {a}',
    pickApt: 'Alege apartamentul tău mai sus ca să vezi calculul pas cu pas.',
    allApts: 'Toate apartamentele',
    lineArea: '{b} × {pct} cota de suprafață',
    linePeople: '{b} × {p} din {P} {w}',
    lineKwh: '{b} × {k} din {K} kWh',
    kwhDetail: '{k} kWh de plată = {own} kWh consum propriu + {com} kWh din consumul comun ({pct} din {c} kWh)',
    personOne: 'persoană',
    personMany: 'persoane',
    bill: 'Factura',
    amount: 'Suma',
    total: 'Total',
    basis: 'Baza de calcul',
    persons: 'Persoane',
    share: 'Cotă de suprafață',
    ownKwh: 'Consum propriu (kWh)',
    comKwh: 'Din consumul comun (kWh)',
    payKwh: 'kWh de plată',
    meters: 'Contoare',
    startReadings: 'Citirile de pornire',
    meter: 'Contor',
    lastMonth: 'Luna trecută',
    thisMonth: 'Luna aceasta',
    reading: 'Citire',
    usage: 'Consum',
    common: 'Comun',
    replacedNote: 'contor înlocuit, consum introdus manual',
    commonNote: 'Fiecare apartament plătește consumul propriului contor; doar consumul contorului comun se împarte după cota de suprafață.',
    note: 'Observații',
    general: 'Contoare generale',
    genGaz: 'Gaz',
    genSus: 'Electric Sus',
    genJos: 'Electric Jos',
    history: 'Istoric',
    range: 'Interval',
    r12: '12 luni',
    r24: '24 de luni',
    rAll: 'Tot',
    chartTitle: 'Plata lunară pe apartament (lei)',
    chartAria: 'Grafic cu plata lunară pe apartament. Folosește săgețile stânga și dreapta ca să parcurgi lunile. Aceleași valori sunt în tabelul de dedesubt.',
    month: 'Luna',
    showAll: 'Arată toate lunile ({n})',
    showLess: 'Arată doar ultimele 12 luni',
    rulesTitle: 'Cum se împart facturile',
    ruleGaz: 'Gazul se împarte după cota de suprafață a fiecărui apartament.',
    rulePeople: 'DIGI și apa se împart după numărul de persoane din luna respectivă.',
    ruleGunoi: 'Gunoiul nu se împarte: fiecare apartament plătește aceeași sumă fixă, {v} pe lună.',
    ruleKwh: 'Fiecare apartament plătește consumul de pe propriul contor. Doar consumul contorului comun se împarte, după cota de suprafață. Toți kWh-ii se plătesc la prețul din facturile de curent ale lunii (total lei ÷ total kWh facturați).',
    ruleKwhOld: 'Până în {m} inclusiv, factura de curent s-a împărțit proporțional cu acești kWh.',
    ruleKwhMixed: 'Unele luni au împărțit factura de curent proporțional cu acești kWh; fiecare lună arată regula folosită.',
    ruleKwhProp: 'Curentul se împarte după kWh: consumul contorului propriu plus partea din consumul comun (după cota de suprafață), raportat la consumul celor 5 contoare.',
    shares: 'Cote de suprafață',
    peopleIn: 'Persoane în {m}',
    updated: 'Ultima modificare: {d}.',
    fromExcel: 'Lunile până în {m} au fost preluate din fișierul Excel.',
    live: 'Pagina se actualizează singură când se salvează o lună nouă.',
    exportXlsx: 'Descarcă Excel',
    exporting: 'Se pregătește…',
    settings: 'Setări',
    fAdd: 'Adaugă {m}',
    fEdit: 'Editează {m}',
    close: 'Închide',
    secBills: 'Facturi (lei)',
    exprHint: 'Poți scrie și un calcul, de ex. 150+120. Zecimalele merg cu virgulă sau cu punct.',
    secEl: 'Facturile de curent ale lunii',
    modePrice: 'După preț pe kWh',
    modeProp: 'Proporțional',
    elAmount: 'Factura de curent (lei)',
    propHint: 'Factura se împarte proporțional cu kWh-ul fiecărui apartament (contorul propriu plus partea din contorul comun), ca în Excel.',
    elLei: 'Total de plată (lei)',
    elKwh: 'Consum facturat (kWh)',
    elAdd: 'Adaugă o factură',
    elDel: 'Șterge factura {n}',
    elHint: 'Pentru fiecare factură de curent scrie totalul de plată și kWh-ul facturat. Prețul pe kWh = total lei ÷ total kWh. Dacă luna aceasta nu ai nicio factură, lasă gol: se folosește prețul de luna trecută.',
    elPriceLive: 'Preț: {p} ({L} ÷ {K} kWh)',
    elPriceCarry: 'Fără factură: se folosește prețul din {m}, {p}.',
    secReadings: 'Citiri contoare (kWh)',
    lastRead: 'luna trecută: {v}',
    usageNow: 'consum: {v} kWh',
    replaced: 'Contor înlocuit',
    overrideLbl: 'Consum real în luna aceasta (kWh)',
    secPeople: 'Persoane',
    less: 'Mai puțin',
    more: 'Mai mult',
    secNote: 'Observații (opțional)',
    notePh: 'ex. Ap. 2 s-a mutat pe 15',
    secGeneral: 'Contoare generale (opțional, nu intră în calcul)',
    preview: 'Previzualizare',
    fillIn: 'Completează facturile și citirile.',
    ready: 'Totul e în regulă. Poți salva.',
    problems1: 'Mai e o problemă de rezolvat.',
    problemsN: 'Mai sunt {n} probleme de rezolvat.',
    save: 'Salvează luna',
    saving: 'Se salvează…',
    del: 'Șterge luna',
    delQ: 'Ștergi {m}? Celelalte luni rămân neschimbate.',
    delYes: 'Da, șterge',
    delNo: 'Nu',
    eRequired: 'Completează câmpul.',
    eNumber: 'Scrie un număr, de ex. 269,4 sau 150+120.',
    ePlain: 'Scrie doar un număr, de ex. 4023.',
    eThousands: 'Scrie fără separator de mii, de ex. 1234,56.',
    eTooBig: 'Valoare prea mare. Verifică suma.',
    eNegative: 'Citirea nu poate fi negativă.',
    eLower: 'E mai mică decât luna trecută ({v}). Dacă s-a schimbat contorul, bifează „Contor înlocuit”.',
    eKwhBig: 'Peste 10.000 kWh într-o lună. Verifică citirea.',
    eOverride: 'Scrie consumul real, în kWh.',
    ePersons: 'Un număr întreg între 0 și 20.',
    ePersonsZero: 'Trebuie să fie cel puțin o persoană în casă.',
    eKwhZero: 'Consumul total este 0 kWh, dar factura de curent nu e 0.',
    eElNeeded: 'Adaugă cel puțin o factură de curent: nu există un preț de luna trecută.',
    eElKwh: 'Facturile de curent trebuie să aibă împreună peste 0 kWh.',
    eElPrice: 'Prețul pe kWh trebuie să fie peste 0. Verifică sumele.',
    eNextLower: 'Cu această citire, {m} ar avea consum negativ la {x}.',
    eNote: 'Cel mult 200 de caractere.',
    ePositive: 'Un număr mai mare decât 0.',
    wBillHigh: 'Mult peste ultimele luni (maximum {v}). Verifică.',
    wKwhHigh: 'Consum mult peste ultimele luni (maximum {v} kWh). Verifică.',
    wPriceOff: 'Prețul e mult diferit de cel anterior ({v}). Verifică sumele și kWh-ul de pe facturi.',
    savedT: '{m} a fost salvată.',
    deletedT: '{m} a fost ștearsă.',
    settingsT: 'Setările au fost salvate.',
    conflict: 'Între timp s-a salvat o altă versiune. Pagina se reîncarcă, iar ce ai scris rămâne în formular.',
    restored: 'Am păstrat ce scriseseși. Verifică și salvează din nou.',
    notWriter: 'Poți vedea pagina, dar nu o poți modifica.',
    unavailable: 'Salvarea nu e disponibilă aici. Deschide pagina în Claude ca să adaugi luni.',
    rateLimited: 'Prea multe salvări într-un timp scurt. Așteaptă un minut și încearcă din nou.',
    saveFailed: 'Salvarea nu a reușit. Încearcă din nou.',
    exportedT: 'Fișierul Excel a fost descărcat.',
    exportFail: 'Descărcarea nu a reușit.',
    setTitle: 'Setări',
    setStart: 'Prima lună din evidență',
    setStartHint: 'Toate lunile primesc numele pornind de aici. Schimb-o doar dacă etichetele lunilor sunt decalate.',
    setSkip: 'Lipsește din numerotare, ca în Excel: {m}.',
    xElLei: 'Facturi curent (lei)',
    xElKwh: 'Facturi curent (kWh)',
    xPrice: 'Preț curent (lei/kWh)',
    xElDue: 'Curent de plată (lei)',
    setLastWouldBe: 'Ultima lună din evidență ar deveni: {m}.',
    setMonth: 'Luna',
    setYear: 'Anul',
    setSurf: 'Suprafață (% din casă)',
    setSurfHint: 'Dacă totalul nu e 100, cotele se recalculează proporțional. Noile cote se aplică tuturor lunilor, inclusiv celor trecute.',
    setSave: 'Salvează setările',
    usedShares: 'Cote folosite: {v}',
    fatal: 'Datele paginii nu pot fi citite.'
  },
  en: {
    locale: 'en-GB',
    tagline: 'House utilities, split between the apartments',
    lang: 'Language',
    myApt: 'My apartment',
    all: 'All',
    apt: 'Ap. {n}',
    you: 'you',
    prevMonth: 'Previous month',
    nextMonth: 'Next month',
    stOk: 'Calculated',
    stStart: 'Start month',
    stBase: 'Starting readings',
    baseNote: 'The record starts from these meter readings. The first month to fill in: {m}.',
    editBase: 'Edit readings',
    fBase: 'Starting readings, {m}',
    baseSave: 'Save readings',
    baseT: 'Starting readings saved.',
    histEmpty: 'No month calculated yet. Once the first month is added, the amounts show here.',
    add: 'Add {m}',
    editorsOnly: 'Only people who can edit the page see this button.',
    edit: 'Edit',
    totalBills: 'Total bills',
    totalPaid: 'Total paid',
    totalDue: 'Total to pay',
    sumOk: 'The apartment amounts add up exactly to the bills.',
    sumOkDue: 'The apartment amounts add up exactly to the total to pay.',
    manualNote: "First month on record: the amounts were set by hand, not calculated. That month's bills came to {v}.",
    billsTitle: "This month's bills",
    dueTitle: "What is split this month",
    gaz: 'Gas',
    digi: 'DIGI',
    gunoi: 'Rubbish',
    gunoiField: 'Rubbish (lei per apartment)',
    gunoiExpr: '{n} × {v}',
    byFixed: 'same amount per apartment',
    lineFixed: 'fixed amount per apartment',
    apa: 'Water',
    curent: 'Electricity (upstairs + downstairs)',
    curentP: 'Electricity',
    byArea: 'by floor area',
    byPeople: 'by people',
    byKwh: 'in proportion to kWh',
    byPrice: 'kWh × price',
    kwhTimes: '{k} kWh × {p}',
    priceTitle: 'Electricity price',
    billN: 'Bill {n}',
    priceCalc: '{L} ÷ {K} kWh = {p}.',
    priceUse: 'At this price, each apartment pays for its own meter plus its share of the common meter.',
    priceCarried: 'No electricity bill this month: the price from {m} is used.',
    howTitle: 'How it was worked out',
    howFor: 'How {a} comes to {v}',
    pickApt: 'Pick your apartment above to see the calculation step by step.',
    allApts: 'All apartments',
    lineArea: '{b} × {pct} floor-area share',
    linePeople: '{b} × {p} of {P} {w}',
    lineKwh: '{b} × {k} of {K} kWh',
    kwhDetail: '{k} kWh to pay = {own} kWh own use + {com} kWh of the common use ({pct} of {c} kWh)',
    personOne: 'person',
    personMany: 'people',
    bill: 'Bill',
    amount: 'Amount',
    total: 'Total',
    basis: 'Basis',
    persons: 'People',
    share: 'Floor-area share',
    ownKwh: 'Own use (kWh)',
    comKwh: 'Of the common use (kWh)',
    payKwh: 'kWh to pay',
    meters: 'Meters',
    startReadings: 'Starting readings',
    meter: 'Meter',
    lastMonth: 'Last month',
    thisMonth: 'This month',
    reading: 'Reading',
    usage: 'Use',
    common: 'Common',
    replacedNote: 'meter replaced, use typed in by hand',
    commonNote: 'Each apartment pays for its own meter; only the common meter is split by floor-area share.',
    note: 'Notes',
    general: 'Main meters',
    genGaz: 'Gas',
    genSus: 'Electric upstairs',
    genJos: 'Electric downstairs',
    history: 'History',
    range: 'Range',
    r12: '12 months',
    r24: '24 months',
    rAll: 'All',
    chartTitle: 'Monthly amount per apartment (lei)',
    chartAria: 'Chart of the monthly amount per apartment. Use the left and right arrow keys to move through the months. The same values are in the table below.',
    month: 'Month',
    showAll: 'Show all months ({n})',
    showLess: 'Show only the last 12 months',
    rulesTitle: 'How the bills are split',
    ruleGaz: "Gas is split by each apartment's share of the floor area.",
    rulePeople: 'DIGI and water are split by the number of people living there that month.',
    ruleGunoi: 'Rubbish is not split: every apartment pays the same fixed amount, {v} a month.',
    ruleKwh: "Each apartment pays for the use on its own meter. Only the common meter's use is split, by floor-area share. All kWh are paid at the price from that month's electricity bills (total lei ÷ total kWh billed).",
    ruleKwhOld: 'Up to and including {m}, the electricity bill was shared out in proportion to these kWh.',
    ruleKwhMixed: 'Some months shared out the electricity bill in proportion to these kWh instead; each month shows the rule it used.',
    ruleKwhProp: "Electricity is split by kWh: each apartment's own meter plus its floor-area share of the common meter, out of the total on all 5 meters.",
    shares: 'Floor-area shares',
    peopleIn: 'People in {m}',
    updated: 'Last change: {d}.',
    fromExcel: 'Months up to {m} were taken over from the Excel file.',
    live: 'This page updates by itself when a new month is saved.',
    exportXlsx: 'Download Excel',
    exporting: 'Preparing…',
    settings: 'Settings',
    fAdd: 'Add {m}',
    fEdit: 'Edit {m}',
    close: 'Close',
    secBills: 'Bills (lei)',
    exprHint: 'You can also type a sum, e.g. 150+120. Decimals work with a comma or a point.',
    secEl: "This month's electricity bills",
    modePrice: 'By price per kWh',
    modeProp: 'Proportional',
    elAmount: 'Electricity bill (lei)',
    propHint: "The bill is shared out in proportion to each apartment's kWh (its own meter plus its share of the common meter), as in the Excel.",
    elLei: 'Total to pay (lei)',
    elKwh: 'Billed use (kWh)',
    elAdd: 'Add a bill',
    elDel: 'Remove bill {n}',
    elHint: "For each electricity bill, type its total to pay and the kWh it bills. Price per kWh = total lei ÷ total kWh. If there is no bill this month, leave this empty: last month's price is used.",
    elPriceLive: 'Price: {p} ({L} ÷ {K} kWh)',
    elPriceCarry: 'No bill: the price from {m} is used, {p}.',
    secReadings: 'Meter readings (kWh)',
    lastRead: 'last month: {v}',
    usageNow: 'use: {v} kWh',
    replaced: 'Meter replaced',
    overrideLbl: 'Actual use this month (kWh)',
    secPeople: 'People',
    less: 'Fewer',
    more: 'More',
    secNote: 'Notes (optional)',
    notePh: 'e.g. Ap. 2 moved out on the 15th',
    secGeneral: 'Main meters (optional, not used in the split)',
    preview: 'Preview',
    fillIn: 'Fill in the bills and the readings.',
    ready: 'Everything checks out. You can save.',
    problems1: 'One thing still needs fixing.',
    problemsN: '{n} things still need fixing.',
    save: 'Save month',
    saving: 'Saving…',
    del: 'Delete month',
    delQ: 'Delete {m}? The other months stay as they are.',
    delYes: 'Yes, delete',
    delNo: 'No',
    eRequired: 'Fill this in.',
    eNumber: 'Type a number, e.g. 269.4 or 150+120.',
    ePlain: 'Type just a number, e.g. 4023.',
    eThousands: 'Leave out thousands separators, e.g. 1234.56.',
    eTooBig: 'That is too large. Check the amount.',
    eNegative: 'A reading cannot be negative.',
    eLower: 'Lower than last month ({v}). If the meter was replaced, tick "Meter replaced".',
    eKwhBig: 'Over 10,000 kWh in one month. Check the reading.',
    eOverride: 'Type the actual use, in kWh.',
    ePersons: 'A whole number from 0 to 20.',
    ePersonsZero: 'At least one person must live in the house.',
    eKwhZero: 'Total use is 0 kWh, but the electricity bill is not 0.',
    eElNeeded: 'Add at least one electricity bill: there is no earlier price to use.',
    eElKwh: 'The electricity bills must add up to more than 0 kWh.',
    eElPrice: 'The price per kWh must be above 0. Check the amounts.',
    eNextLower: 'With this reading, {m} would have negative use on {x}.',
    eNote: '200 characters at most.',
    ePositive: 'A number above 0.',
    wBillHigh: 'Well above recent months (highest {v}). Check it.',
    wKwhHigh: 'Use well above recent months (highest {v} kWh). Check it.',
    wPriceOff: 'The price is far from the previous one ({v}). Check the amounts and kWh on the bills.',
    savedT: '{m} saved.',
    deletedT: '{m} deleted.',
    settingsT: 'Settings saved.',
    conflict: 'Someone saved another version meanwhile. The page is reloading, and what you typed stays in the form.',
    restored: 'Your entries were kept. Check them and save again.',
    notWriter: 'You can view this page but not change it.',
    unavailable: 'Saving is not available here. Open the page in Claude to add months.',
    rateLimited: 'Too many saves in a short time. Wait a minute and try again.',
    saveFailed: 'Saving did not work. Try again.',
    exportedT: 'Excel file downloaded.',
    exportFail: 'The download did not work.',
    setTitle: 'Settings',
    setStart: 'First month on record',
    setStartHint: 'All months are named counting from this one. Change it only if the month labels are off.',
    setSkip: 'Left out of the month names, as in the Excel: {m}.',
    xElLei: 'Electricity bills (lei)',
    xElKwh: 'Electricity bills (kWh)',
    xPrice: 'Electricity price (lei/kWh)',
    xElDue: 'Electricity to pay (lei)',
    setLastWouldBe: 'The latest month would become: {m}.',
    setMonth: 'Month',
    setYear: 'Year',
    setSurf: 'Floor area (% of the house)',
    setSurfHint: 'If the total is not 100, the shares are scaled to add up. New shares apply to every month, past ones included.',
    setSave: 'Save settings',
    usedShares: 'Shares used: {v}',
    fatal: 'The page data cannot be read.'
  }
};

/* ------------------------------------------------------------------ app state */

var state = {
  data: null,
  comp: [],
  lang: 'ro',
  apt: null,
  sel: 0,
  range: 24,
  histAll: false,
  userCanEdit: false,
  artifact: null,
  downloads: null,
  canEdit: false,
  sheet: null,
  pendingStash: null,
  exporting: false
};
var root = null;

function $(id) {
  return document.getElementById(id);
}

function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function tr(key, params) {
  var dict = I18N[state.lang] || I18N.ro;
  var s = dict[key] != null ? dict[key] : I18N.ro[key];
  if (s == null) return key;
  if (params) s = s.replace(/\{(\w+)\}/g, function (m, k) { return params[k] != null ? params[k] : m; });
  return s;
}

function lsGet(k) { try { return window.localStorage.getItem(k); } catch (e) { return null; } }
function lsSet(k, v) { try { window.localStorage.setItem(k, v); } catch (e) { /* storage unavailable */ } }
function lsDel(k) { try { window.localStorage.removeItem(k); } catch (e) { /* storage unavailable */ } }
function ssGet(k) { try { return window.sessionStorage.getItem(k); } catch (e) { return null; } }
function ssSet(k, v) { try { window.sessionStorage.setItem(k, v); } catch (e) { /* storage unavailable */ } }
function ssDel(k) { try { window.sessionStorage.removeItem(k); } catch (e) { /* storage unavailable */ } }

var NF = {};
function fmt(n, min, max, group) {
  if (!num(n)) return '—';
  var key = state.lang + '|' + min + '|' + max + '|' + group;
  var f = NF[key];
  if (!f) f = NF[key] = new Intl.NumberFormat(I18N[state.lang].locale, { minimumFractionDigits: min, maximumFractionDigits: max, useGrouping: group });
  return f.format(Math.abs(n) < 0.005 && min === 2 ? 0 : n);
}
function money(n) { return fmt(n, 2, 2, true); }
function lei(n) { return money(n) + '\u00a0lei'; }
function kwhStr(n) { return fmt(num(n) ? Math.round(n * 100) / 100 : n, 0, 2, false); }
function kwhTotal(n) { return fmt(num(n) ? Math.round(n * 100) / 100 : n, 0, 2, true); }
function priceStr(p) { return fmt(p, 4, 4, false) + ' lei/kWh'; }
function pct(n) { return fmt(n * 100, 2, 2, false) + (state.lang === 'ro' ? '\u00a0%' : '%'); }
function drum(v) {
  if (!num(v)) return '—';
  var neg = v < 0, a = Math.abs(v);
  var ip = Math.floor(a), fp = Math.round((a - ip) * 100);
  var s = String(ip).padStart(5, '0');
  if (fp) s += (state.lang === 'ro' ? ',' : '.') + String(fp).padStart(2, '0').replace(/0$/, '');
  return (neg ? '-' : '') + s;
}

function ym(i) { return monthAt(state.data.settings, i); }
function mName(i, style) {
  var d = ym(i);
  var o = style === 'month' ? { month: 'long' } : style === 'short' ? { month: 'short' } :
    style === 'shortYear' ? { month: 'short', year: 'numeric' } : { month: 'long', year: 'numeric' };
  o.timeZone = 'UTC';
  return new Intl.DateTimeFormat(I18N[state.lang].locale, o).format(new Date(Date.UTC(d.y, d.m, 1)));
}
function cap(s) { return s ? s.charAt(0).toLocaleUpperCase() + s.slice(1) : s; }
function aptName(a) { return tr('apt', { n: a + 1 }); }
function meterName(j) { return j < 4 ? aptName(j) : tr('common'); }
function peopleWord(n) {
  if (state.lang === 'ro') return n === 1 ? tr('personOne') : (n % 100 >= 20 || (n >= 100 && n % 100 === 0) ? 'de ' : '') + tr('personMany');
  return n === 1 ? tr('personOne') : tr('personMany');
}
function numToInput(v) {
  if (v == null || !num(v)) return '';
  return state.lang === 'ro' ? String(v).replace('.', ',') : String(v);
}
function exprDisplay(t) {
  var s = String(t).replace(/\*/g, '×').replace(/\//g, '÷').replace(/([+\-×÷])/g, ' $1 ');
  if (state.lang === 'ro') s = s.replace(/(\d)\.(\d)/g, '$1,$2');
  return s.replace(/\s+/g, ' ').trim();
}
function todayISO() { return new Date().toISOString().slice(0, 10); }
function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
function clone(o) { return JSON.parse(JSON.stringify(o)); }

var ICONS = {
  prev: '<path d="M15 18l-6-6 6-6"/>',
  next: '<path d="M9 18l6-6-6-6"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  edit: '<path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/>',
  check: '<path d="M20 6 9 17l-5-5"/>',
  x: '<path d="M18 6 6 18M6 6l12 12"/>',
  down: '<path d="M12 3v12"/><path d="m7 10 5 5 5-5"/><path d="M5 21h14"/>',
  sliders: '<path d="M4 6h9M17 6h3M4 12h3M11 12h9M4 18h11M19 18h1"/><circle cx="15" cy="6" r="2"/><circle cx="9" cy="12" r="2"/><circle cx="17" cy="18" r="2"/>'
};
function eyebrow(s) {
  return esc(s).replace('(kWh)', '<span class="nt">(kWh)</span>');
}
function icon(n) {
  return '<svg class="i" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + ICONS[n] + '</svg>';
}

/* ------------------------------------------------------------------ boot */

function pickLang() {
  var s = lsGet('ff-lang');
  if (s === 'ro' || s === 'en') return s;
  var langs = (navigator.languages && navigator.languages.length) ? navigator.languages : [navigator.language || ''];
  for (var k = 0; k < langs.length; k++) if (/^ro\b/i.test(langs[k])) return 'ro';
  return 'en';
}

function aptFromHash() {
  var m = /^#ap-([1-4])$/.exec(location.hash || '');
  if (m) return +m[1] - 1;
  if (location.hash === '#all') return null;
  return undefined;
}

function pickApt() {
  var h = aptFromHash();
  if (h !== undefined) return h;
  var s = lsGet('ff-apt');
  return /^[0-3]$/.test(s || '') ? +s : null;
}

function boot() {
  root = $('app');
  try {
    state.data = normalizeData(JSON.parse($('ff-data').textContent));
    if (!state.data.months.length) throw new Error('empty');
  } catch (e) {
    root.innerHTML = '<p class="fatal">' + esc(I18N.ro.fatal) + ' / ' + esc(I18N.en.fatal) + '</p>';
    return;
  }
  state.lang = pickLang();
  state.apt = pickApt();
  var r = lsGet('ff-range');
  state.range = r === '12' ? 12 : r === '0' ? 0 : 24;
  state.sel = state.data.months.length - 1;
  state.comp = computeAll(state.data);

  var after = null;
  try { after = JSON.parse(ssGet('ff-after') || 'null'); } catch (e) { after = null; }
  ssDel('ff-after');
  if (after && typeof after.index === 'number' && after.index >= 0 && after.index < state.data.months.length) state.sel = after.index;
  var stash = null;
  try { stash = JSON.parse(ssGet('ff-stash') || 'null'); } catch (e) { stash = null; }
  ssDel('ff-stash');
  if (stash && stash.index === state.data.months.length && stash.draft) state.pendingStash = stash;
  try {
    var ld = JSON.parse(lsGet('ff-draft') || 'null');
    if (ld && ld.index !== state.data.months.length) lsDel('ff-draft');
  } catch (e) { lsDel('ff-draft'); }

  document.documentElement.lang = state.lang;
  var toast = document.createElement('div');
  toast.className = 'toast';
  toast.id = 'toast';
  toast.setAttribute('role', 'status');
  toast.setAttribute('aria-live', 'polite');
  document.body.appendChild(toast);
  renderAll();
  bindEvents();
  if (after) showToast(afterMessage(after));
  connect();
}

function afterMessage(a) {
  if (a.type === 'settings') return tr('settingsT');
  if (a.type === 'base') return tr('baseT');
  if (a.type === 'deleted') return tr('deletedT', { m: cap(mName(a.index)) });
  return tr('savedT', { m: cap(mName(a.index)) });
}

function connect() {
  var c = window.claude;
  if (!c || typeof c.use !== 'function') return;
  c.use('user').then(function (u) {
    return u && typeof u.canEdit === 'function' ? u.canEdit() : false;
  }).then(function (v) { state.userCanEdit = v === true; updateEditability(); }, function () {});
  c.use('artifact').then(function (a) { state.artifact = a || null; updateEditability(); }, function () {});
  c.use('downloads').then(function (d) { state.downloads = d || null; renderFooter(); }, function () {});
}

function updateEditability() {
  var can = !!(state.userCanEdit && state.artifact);
  if (can === state.canEdit) return;
  state.canEdit = can;
  if (!can && state.sheet) state.sheet = null;
  renderAll();
  if (can && state.pendingStash) {
    var st = state.pendingStash;
    state.pendingStash = null;
    openMonthSheet('add', null, st.draft);
    setSheetMsg(tr('restored'), 'warn');
  }
}

/* ------------------------------------------------------------------ rendering */

function renderAll() {
  if (!root) return;
  var h = '';
  h += '<header class="band"><div class="band-inner"><div class="brand"><h1>Facturi Fierari</h1><p>' + esc(tr('tagline')) + '</p></div>';
  h += '<div class="lang" role="group" aria-label="' + esc(tr('lang')) + '">' +
    ['ro', 'en'].map(function (l) {
      return '<button type="button" data-action="lang" data-v="' + l + '" aria-pressed="' + (state.lang === l) + '" lang="' + l + '">' + l.toUpperCase() + '</button>';
    }).join('') + '</div></div></header>';
  h += '<div class="wrap">';
  h += '<div class="myapt" role="group" aria-labelledby="mine-lbl"><p class="eyebrow" id="mine-lbl">' + esc(tr('myApt')) + '</p><div class="chips">';
  h += '<button type="button" class="chip" data-action="apt" data-v="" aria-pressed="' + (state.apt == null) + '">' + esc(tr('all')) + '</button>';
  APTS.forEach(function (a) {
    h += '<button type="button" class="chip a' + a + '" data-action="apt" data-v="' + a + '" aria-pressed="' + (state.apt === a) + '"><span class="dot" aria-hidden="true"></span>' + esc(aptName(a)) + '</button>';
  });
  h += '</div></div>';
  if (state.canEdit) {
    h += '<div class="editor"><button type="button" class="btn primary" data-action="add">' + icon('plus') +
      esc(tr('add', { m: mName(state.data.months.length) })) + '</button><span class="muted">' + esc(tr('editorsOnly')) + '</span></div>';
  }
  h += '<div class="cols"><section class="panel" id="month-panel" aria-labelledby="month-title"></section>' +
    '<section class="panel" id="history-panel" aria-labelledby="hist-title"></section></div>';
  h += '<section class="panel" id="rules-panel" aria-labelledby="rules-title"></section>';
  h += '<footer class="foot" id="foot"></footer>';
  h += '</div><div id="sheet-host"></div>';
  root.innerHTML = h;
  renderMonth();
  renderHistory();
  renderRules();
  renderFooter();
  if (state.sheet) renderSheet();
}

function pill(kind) {
  if (kind === 'ok') return '<span class="pill">' + esc(tr('stOk')) + '</span>';
  return '<span class="pill neutral">' + esc(tr(kind === 'base' ? 'stBase' : 'stStart')) + '</span>';
}

function renderMonth() {
  var el = $('month-panel');
  if (!el) return;
  var i = state.sel, n = state.data.months.length;
  var c = state.comp[i], m = state.data.months[i];
  var h = '<div class="mhead"><div class="mnav">';
  h += '<button type="button" class="nav-btn" data-action="prev" aria-label="' + esc(tr('prevMonth')) + '"' + (i > 0 ? '' : ' disabled') + '>' + icon('prev') + '</button>';
  h += '<h2 class="mtitle" id="month-title">' + esc(cap(mName(i))) + '</h2>';
  h += '<button type="button" class="nav-btn" data-action="next" aria-label="' + esc(tr('nextMonth')) + '"' + (i < n - 1 ? '' : ' disabled') + '>' + icon('next') + '</button>';
  h += '</div><div class="mmeta">' + pill(c.start ? 'base' : c.manual ? 'start' : 'ok');
  if (state.canEdit && c.start) h += '<button type="button" class="btn sm" data-action="edit-base">' + icon('edit') + esc(tr('editBase')) + '</button>';
  else if (state.canEdit && i > 0) h += '<button type="button" class="btn sm" data-action="edit">' + icon('edit') + esc(tr('edit')) + '</button>';
  h += '</div></div>';

  /* the starting readings: no bills, nothing to pay */
  if (c.start) {
    h += '<p class="basenote">' + esc(tr('baseNote', { m: mName(i + 1) })) + '</p>' + metersHTML(i);
    el.innerHTML = h;
    return;
  }

  h += '<div class="cwrap"><div class="coupons">';
  APTS.forEach(function (a) {
    var mine = state.apt === a;
    var cents = c.cents[a];
    h += '<div class="coupon a' + a + (mine ? ' mine' : '') + '">';
    h += '<div class="who"><span class="key" aria-hidden="true"></span><span>' + esc(aptName(a)) + '</span>' + (mine ? '<span class="you">' + esc(tr('you')) + '</span>' : '') + '</div>';
    h += '<div class="amt">' + esc(money(cents / 100)) + '<span class="cur">lei</span></div></div>';
  });
  h += '</div></div>';

  if (c.manual) {
    h += '<div class="totalrow"><span class="eyebrow">' + esc(tr('totalPaid')) + '</span><span class="tval">' + esc(lei(c.total)) + '</span></div>';
    h += '<p class="noteline">' + esc(tr('manualNote', { v: lei(c.billTotal) })) + '</p>';
  } else if (c.byPrice) {
    h += '<div class="totalrow"><span class="eyebrow">' + esc(tr('totalDue')) + '</span><span class="tval">' + esc(lei(c.due)) + '</span></div>';
    if (Math.round(c.due * 100) === sum(c.cents)) h += '<p class="okline">' + icon('check') + esc(tr('sumOkDue')) + '</p>';
  } else {
    h += '<div class="totalrow"><span class="eyebrow">' + esc(tr('totalBills')) + '</span><span class="tval">' + esc(lei(c.billTotal)) + '</span></div>';
    if (Math.round(c.billTotal * 100) === sum(c.cents)) h += '<p class="okline">' + icon('check') + esc(tr('sumOk')) + '</p>';
  }

  h += '<h3 class="h3">' + esc(tr(c.byPrice ? 'dueTitle' : 'billsTitle')) + '</h3><div class="bills">';
  BILLS.forEach(function (b) {
    if (b === 'gunoi' && !num(m.gunoi)) return;
    var elRow = b === 'curent' && c.byPrice;
    var val = elRow ? c.elDue : b === 'gunoi' ? m.gunoi * APTS.length : m[b];
    var sub = elRow ? tr('kwhTimes', { k: kwhStr(c.K), p: priceStr(c.pk) }) : b === 'gunoi' ? tr('gunoiExpr', { n: APTS.length, v: lei(m.gunoi) }) :
      m[b + 'Text'] ? exprDisplay(m[b + 'Text']) : '';
    h += '<div class="bill"><span class="bname">' + esc(tr(elRow ? 'curentP' : b)) + '</span>' +
      (c.manual ? '' : '<span class="brule">' + esc(tr(ruleOf(c, b))) + '</span>') +
      '<span class="bval">' + esc(lei(val)) + (sub ? '<span class="expr">' + esc(sub) + '</span>' : '') + '</span></div>';
  });
  h += '</div>';
  if (c.byPrice) h += priceBoxHTML(i);
  if (m.note) h += '<p class="note"><b>' + esc(tr('note')) + ':</b> ' + esc(m.note) + '</p>';

  if (!c.manual) {
    h += '<h3 class="h3">' + esc(tr('howTitle')) + '</h3>';
    if (state.apt != null) {
      h += narrativeHTML(i, state.apt);
      h += '<details class="more"><summary>' + esc(tr('allApts')) + '</summary>' + splitTableHTML(i) + basisTableHTML(i) + '</details>';
    } else {
      h += '<p class="hint">' + esc(tr('pickApt')) + '</p>' + splitTableHTML(i) + basisTableHTML(i);
    }
  }
  h += metersHTML(i);
  if (m.g) {
    var gl = ['genGaz', 'genSus', 'genJos'].map(function (k, j) { return m.g[j] == null ? null : tr(k) + ' ' + kwhStr(m.g[j]); }).filter(Boolean);
    if (gl.length) h += '<p class="gen-line">' + esc(tr('general')) + ': ' + esc(gl.join(' · ')) + '</p>';
  }
  el.innerHTML = h;
}

function ruleOf(c, b) {
  return b === 'gaz' ? 'byArea' : b === 'gunoi' ? 'byFixed' : b === 'curent' ? (c.byPrice ? 'byPrice' : 'byKwh') : 'byPeople';
}

/* Where the month's price per kWh comes from: its electricity bills, or an earlier month's price. */
function priceBoxHTML(i) {
  var c = state.comp[i], m = state.data.months[i];
  var h = '<div class="pricebox"><div class="pb-head"><span class="eyebrow">' + esc(tr('priceTitle')) + '</span><b class="pb-val">' + esc(priceStr(c.pk)) + '</b></div>';
  if (m.el && m.el.length) {
    var L = sum(m.el.map(function (b) { return b.lei; })), K = sum(m.el.map(function (b) { return b.kwh; }));
    h += '<table class="pb-tbl"><tbody>';
    m.el.forEach(function (b, n) {
      h += '<tr><th scope="row">' + esc(tr('billN', { n: n + 1 })) + '</th><td class="num">' + esc(lei(b.lei)) + '</td><td class="num">' + esc(kwhTotal(b.kwh)) + ' kWh</td></tr>';
    });
    if (m.el.length > 1) {
      h += '<tr class="pb-tot"><th scope="row">' + esc(tr('total')) + '</th><td class="num">' + esc(lei(L)) + '</td><td class="num">' + esc(kwhTotal(K)) + ' kWh</td></tr>';
    }
    h += '</tbody></table><p class="pb-note">' + esc(tr('priceCalc', { L: lei(L), K: kwhTotal(K), p: priceStr(c.pk) })) + ' ' + esc(tr('priceUse')) + '</p>';
  } else {
    var lp = lastPrice(state.data, i);
    h += '<p class="pb-note">' + esc(tr('priceCarried', { m: lp ? mName(lp.i) : '—' })) + ' ' + esc(tr('priceUse')) + '</p>';
  }
  return h + '</div>';
}

function narrativeHTML(i, a) {
  var c = state.comp[i], m = state.data.months[i];
  var elLine = c.byPrice ? tr('kwhTimes', { k: kwhStr(c.kwhPay[a]), p: priceStr(c.pk) }) :
    tr('lineKwh', { b: money(m.curent), k: kwhStr(c.kwhPay[a]), K: kwhStr(c.K) });
  var rows = [
    ['gaz', tr('lineArea', { b: money(m.gaz), pct: pct(c.shares[a]) }), c.parts.gaz[a], ''],
    ['digi', tr('linePeople', { b: money(m.digi), p: m.p[a], P: c.P, w: peopleWord(c.P) }), c.parts.digi[a], ''],
    ['gunoi', tr('lineFixed'), c.parts.gunoi[a], ''],
    ['apa', tr('linePeople', { b: money(m.apa), p: m.p[a], P: c.P, w: peopleWord(c.P) }), c.parts.apa[a], ''],
    ['curent', elLine, c.parts.curent[a],
      tr('kwhDetail', { k: kwhStr(c.kwhPay[a]), own: kwhStr(c.kwh[a]), com: kwhStr(c.shares[a] * c.kwh[4]), pct: pct(c.shares[a]), c: kwhStr(c.kwh[4]) })]
  ];
  var h = '<div class="calc a' + a + '"><p class="calc-title"><span class="key" aria-hidden="true"></span>' +
    esc(tr('howFor', { a: aptName(a), v: lei(c.cents[a] / 100) })) + '</p>';
  rows.forEach(function (r) {
    if (r[0] === 'gunoi' && !num(m.gunoi)) return;
    h += '<div class="crow"><span class="cname">' + esc(tr(r[0] === 'curent' && c.byPrice ? 'curentP' : r[0])) + '</span><span class="chow">' + esc(r[1]) +
      (r[3] ? '<small>' + esc(r[3]) + '</small>' : '') + '</span><span class="cval">' + esc(money(r[2])) + '</span></div>';
  });
  h += '<div class="crow tot"><span class="cname">' + esc(tr('total')) + '</span><span class="chow"></span><span class="cval">' + esc(lei(c.cents[a] / 100)) + '</span></div></div>';
  return h;
}

function aptHeadCells() {
  return APTS.map(function (a) {
    return '<th scope="col" class="num a' + a + (state.apt === a ? ' mine' : '') + '"><span class="key" aria-hidden="true"></span>' + esc(aptName(a)) + '</th>';
  }).join('');
}
function aptCells(vals, f) {
  return APTS.map(function (a) {
    return '<td class="num a' + a + (state.apt === a ? ' mine' : '') + '">' + esc(f(vals[a])) + '</td>';
  }).join('');
}

function splitTableHTML(i) {
  var c = state.comp[i], m = state.data.months[i];
  var h = '<div class="scroll"><table class="tbl split"><thead><tr><th scope="col">' + esc(tr('bill')) + '</th><th scope="col" class="num">' + esc(tr('amount')) + '</th>' + aptHeadCells() + '</tr></thead><tbody>';
  BILLS.forEach(function (b) {
    if (b === 'gunoi' && !num(m.gunoi)) return;
    var elRow = b === 'curent' && c.byPrice;
    var amount = elRow ? c.elDue : b === 'gunoi' ? m.gunoi * APTS.length : m[b];
    h += '<tr><th scope="row">' + esc(tr(elRow ? 'curentP' : b)) + '<small>' + esc(tr(ruleOf(c, b))) + '</small></th><td class="num">' + esc(money(amount)) + '</td>' + aptCells(c.parts[b], money) + '</tr>';
  });
  h += '</tbody><tfoot><tr><th scope="row">' + esc(tr('total')) + '</th><td class="num">' + esc(money(c.byPrice ? c.due : c.billTotal)) + '</td>' +
    aptCells(c.cents.map(function (v) { return v / 100; }), money) + '</tr></tfoot></table></div>';
  return h;
}

function basisTableHTML(i) {
  var c = state.comp[i], m = state.data.months[i];
  var com = APTS.map(function (a) { return c.shares[a] * c.kwh[4]; });
  var rows = [
    [tr('persons'), m.p, c.P, function (v) { return fmt(v, 0, 0, false); }],
    [tr('share'), c.shares, 1, pct],
    [tr('ownKwh'), c.kwh.slice(0, 4), sum(c.kwh.slice(0, 4)), kwhStr],
    [tr('comKwh'), com, c.kwh[4], kwhStr],
    [tr('payKwh'), c.kwhPay, c.K, kwhStr]
  ];
  var h = '<div class="scroll"><table class="tbl basis"><thead><tr><th scope="col">' + esc(tr('basis')) + '</th><th scope="col" class="num">' + esc(tr('total')) + '</th>' + aptHeadCells() + '</tr></thead><tbody>';
  rows.forEach(function (r) {
    h += '<tr><th scope="row">' + esc(r[0]) + '</th><td class="num">' + esc(r[3](r[2])) + '</td>' + aptCells(r[1], r[3]) + '</tr>';
  });
  return h + '</tbody></table></div>';
}

function metersHTML(i) {
  var m = state.data.months[i], c = state.comp[i];
  var prev = i > 0 ? state.data.months[i - 1] : null;
  var first = c.manual || c.start;
  var h = '<h3 class="h3">' + esc(tr(first ? 'startReadings' : 'meters')) + '</h3><div class="scroll"><table class="tbl meters"><thead><tr><th scope="col">' + esc(tr('meter')) + '</th>';
  if (first) h += '<th scope="col">' + esc(tr('reading')) + '</th>';
  else h += '<th scope="col">' + esc(tr('lastMonth')) + '</th><th scope="col">' + esc(tr('thisMonth')) + '</th><th scope="col" class="num">' + esc(tr('usage')) + '</th>';
  h += '</tr></thead><tbody>';
  METERS.forEach(function (j) {
    var key = j < 4 ? '<span class="key a' + j + '" aria-hidden="true"></span> ' : '';
    h += '<tr' + (j < 4 && state.apt === j ? ' class="a' + j + '"' : '') + '><th scope="row">' + key + esc(meterName(j)) + '</th>';
    if (first) h += '<td><span class="drum">' + esc(drum(m.r[j])) + '</span></td>';
    else {
      var replaced = m.kwh && m.kwh[j] != null;
      h += '<td><span class="drum">' + esc(drum(prev.r[j])) + '</span></td><td><span class="drum">' + esc(drum(m.r[j])) + '</span></td>' +
        '<td class="num">' + esc(kwhStr(c.kwh[j])) + '\u00a0kWh' + (replaced ? '<span class="tag">' + esc(tr('replacedNote')) + '</span>' : '') + '</td>';
    }
    h += '</tr>';
  });
  h += '</tbody></table></div>';
  if (!first) h += '<p class="hint">' + esc(tr('commonNote')) + '</p>';
  return h;
}

/* months with amounts to show: everything but the starting readings */
function paidMonths() {
  var out = [];
  state.comp.forEach(function (c, i) { if (!c.start) out.push(i); });
  return out;
}

function renderHistory() {
  var el = $('history-panel');
  if (!el) return;
  var paid = paidMonths(), n = paid.length;
  var h = '<div class="hhead"><h2 class="h2" id="hist-title">' + esc(tr('history')) + '</h2>';
  if (!n) {
    el.innerHTML = h + '</div><p class="hint hist-empty">' + esc(tr('histEmpty')) + '</p>';
    return;
  }
  h += '<div class="seg range" role="group" aria-label="' + esc(tr('range')) + '">' +
    [[12, 'r12'], [24, 'r24'], [0, 'rAll']].map(function (r) {
      return '<button type="button" data-action="range" data-v="' + r[0] + '" aria-pressed="' + (state.range === r[0]) + '">' + esc(tr(r[1])) + '</button>';
    }).join('') + '</div></div>';
  h += '<p class="chart-title" id="chart-title">' + esc(tr('chartTitle')) + '</p>';
  h += '<ul class="legend">' + APTS.map(function (a) {
    var cls = 'a' + a + (state.apt === a ? ' mine' : state.apt != null ? ' dim' : '');
    return '<li class="' + cls + '"><span class="key" aria-hidden="true"></span>' + esc(aptName(a)) + '</li>';
  }).join('') + '</ul>';
  h += '<div class="chart-wrap" id="chart" tabindex="0" role="group" aria-roledescription="chart" aria-labelledby="chart-title" aria-describedby="chart-desc"></div>';
  h += '<p class="hint" id="chart-desc" hidden>' + esc(tr('chartAria')) + '</p>';

  var rows = paid.slice().reverse();
  var shown = state.histAll ? rows : rows.slice(0, 12);
  h += '<div class="scroll"><table class="tbl hist"><thead><tr><th scope="col">' + esc(tr('month')) + '</th>' + aptHeadCells() + '<th scope="col" class="num">' + esc(tr('total')) + '</th></tr></thead><tbody>';
  shown.forEach(function (i) {
    var c = state.comp[i];
    h += '<tr' + (i === state.sel ? ' class="sel"' : '') + '><th scope="row"><button type="button" class="rowbtn" data-action="pick" data-i="' + i + '"' +
      (i === state.sel ? ' aria-current="true"' : '') + '>' + esc(cap(mName(i, 'shortYear'))) + '</button></th>' +
      aptCells(c.cents.map(function (v) { return v / 100; }), money) + '<td class="num">' + esc(money(c.total)) + '</td></tr>';
  });
  h += '</tbody></table></div>';
  if (n > 12) h += '<button type="button" class="linkbtn" data-action="hist-all">' + esc(state.histAll ? tr('showLess') : tr('showAll', { n: n })) + '</button>';
  el.innerHTML = h;
  renderChart();
  var ch = $('chart');
  ch.addEventListener('pointermove', chartPointer);
  ch.addEventListener('pointerdown', chartPointer);
  ch.addEventListener('pointerleave', function () { if (document.activeElement !== ch) hideCrosshair(); });
  ch.addEventListener('focus', function () { var g = ch._geo; if (g) showCrosshair(g.idx.length - 1); });
  ch.addEventListener('blur', hideCrosshair);
  ch.addEventListener('keydown', chartKey);
}

function niceScale(lo, hi, count) {
  if (hi <= lo) hi = lo + 1;
  var raw = (hi - lo) / count;
  var p = Math.pow(10, Math.floor(Math.log(raw) / Math.LN10));
  var f = raw / p;
  var step = (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p;
  var a = Math.floor(lo / step) * step, b = Math.ceil(hi / step) * step;
  var ticks = [];
  for (var v = a; v <= b + step / 2; v += step) ticks.push(Math.round(v * 100) / 100);
  return { lo: a, hi: b, ticks: ticks };
}

function renderChart() {
  var el = $('chart');
  if (!el) return;
  var W = Math.max(280, Math.round(el.clientWidth || 600));
  var R = state.range, paid = paidMonths();
  var idx = R ? paid.slice(-R) : paid;
  if (!idx.length) return;
  var k;
  var series = APTS.map(function (a) { return idx.map(function (i) { return state.comp[i].cents[a] / 100; }); });
  var all = [].concat.apply([], series).filter(num);
  var sc = niceScale(Math.min(0, Math.min.apply(null, all)), Math.max.apply(null, all), 4);
  var H = 240, padT = 12, padB = 30, padL = 46;
  var plotH = H - padT - padB;
  function y(v) { return padT + plotH * (1 - (v - sc.lo) / (sc.hi - sc.lo)); }
  var ends = APTS.map(function (a) { return y(series[a][idx.length - 1]); });
  var sorted = ends.slice().sort(function (p, q) { return p - q; });
  var fits = true;
  for (k = 1; k < sorted.length; k++) if (sorted[k] - sorted[k - 1] < 14) fits = false;
  var padR = fits ? 54 : 16;
  var plotW = W - padL - padR;
  function x(j) { return padL + (idx.length === 1 ? plotW / 2 : j * plotW / (idx.length - 1)); }

  var s = '<svg class="chart" viewBox="0 0 ' + W + ' ' + H + '" width="' + W + '" height="' + H + '" aria-hidden="true" focusable="false">';
  sc.ticks.forEach(function (t) {
    var yy = Math.round(y(t)) + 0.5;
    s += '<line class="grid' + (t === 0 ? ' base' : '') + '" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + yy + '" y2="' + yy + '"/>';
    s += '<text class="ax" x="' + (padL - 8) + '" y="' + (yy + 4) + '" text-anchor="end">' + esc(fmt(t, 0, 0, true)) + '</text>';
  });
  var last = -Infinity;
  idx.forEach(function (i, j) {
    var d = ym(i), label = null;
    if (d.m === 0) label = String(d.y);
    else if (R && R <= 24 && d.m % 3 === 0) label = mName(i, 'short');
    if (label == null) return;
    var xx = x(j);
    if (xx - last < 34) return;
    last = xx;
    s += '<line class="tickm" x1="' + (Math.round(xx) + 0.5) + '" x2="' + (Math.round(xx) + 0.5) + '" y1="' + (padT + plotH) + '" y2="' + (padT + plotH + 4) + '"/>';
    s += '<text class="ax" x="' + xx + '" y="' + (H - 9) + '" text-anchor="middle">' + esc(label) + '</text>';
  });
  var order = state.apt == null ? APTS.slice() : APTS.filter(function (a) { return a !== state.apt; }).concat([state.apt]);
  order.forEach(function (a) {
    var dim = state.apt != null && state.apt !== a;
    var d = series[a].map(function (v, j) { return (j ? 'L' : 'M') + x(j).toFixed(1) + ' ' + y(v).toFixed(1); }).join('');
    s += '<path class="ln s' + a + (dim ? ' dim' : '') + '" d="' + d + '"/>';
  });
  order.forEach(function (a) {
    var dim = state.apt != null && state.apt !== a;
    s += '<circle class="dot f' + a + (dim ? ' dim' : '') + '" cx="' + x(idx.length - 1).toFixed(1) + '" cy="' + ends[a].toFixed(1) + '" r="4"/>';
  });
  if (fits) {
    APTS.forEach(function (a) {
      s += '<text class="elab' + (state.apt === a ? ' mine' : '') + '" x="' + (x(idx.length - 1) + 9).toFixed(1) + '" y="' + (ends[a] + 4).toFixed(1) + '">' + esc(aptName(a)) + '</text>';
    });
  }
  s += '<g id="xh" visibility="hidden"><line class="xh-l" id="xh-l" y1="' + padT + '" y2="' + (padT + plotH) + '"/>' +
    APTS.map(function (a) { return '<circle class="xh-d f' + a + '" id="xh-d' + a + '" r="4.5"/>'; }).join('') + '</g>';
  s += '<rect class="hit" x="' + (padL - 8) + '" y="0" width="' + (plotW + 16) + '" height="' + H + '"/>';
  s += '</svg><div class="tip" id="chart-tip" hidden></div>';
  el.innerHTML = s;
  el._geo = { idx: idx, x: x, y: y, padL: padL, plotW: plotW, W: W, series: series };
}

function chartPointer(ev) {
  var el = $('chart'), g = el && el._geo;
  if (!g) return;
  var svg = el.querySelector('svg');
  var rect = svg.getBoundingClientRect();
  var px = (ev.clientX - rect.left) * (g.W / rect.width);
  var step = g.idx.length > 1 ? g.plotW / (g.idx.length - 1) : 1;
  var j = Math.round((px - g.padL) / step);
  showCrosshair(Math.max(0, Math.min(g.idx.length - 1, j)));
}

function chartKey(ev) {
  var el = $('chart'), g = el && el._geo;
  if (!g) return;
  var j = el._j == null ? g.idx.length - 1 : el._j;
  if (ev.key === 'ArrowLeft') j--;
  else if (ev.key === 'ArrowRight') j++;
  else if (ev.key === 'Home') j = 0;
  else if (ev.key === 'End') j = g.idx.length - 1;
  else return;
  ev.preventDefault();
  showCrosshair(Math.max(0, Math.min(g.idx.length - 1, j)));
}

function showCrosshair(j) {
  var el = $('chart'), g = el && el._geo;
  if (!g) return;
  el._j = j;
  var i = g.idx[j], xx = g.x(j);
  var xh = $('xh');
  xh.setAttribute('visibility', 'visible');
  var l = $('xh-l');
  l.setAttribute('x1', xx);
  l.setAttribute('x2', xx);
  APTS.forEach(function (a) {
    var d = $('xh-d' + a);
    d.setAttribute('cx', xx);
    d.setAttribute('cy', g.y(g.series[a][j]));
  });
  var tip = $('chart-tip');
  var c = state.comp[i];
  var h = '<p class="tip-m">' + esc(cap(mName(i))) + '</p>';
  APTS.forEach(function (a) {
    h += '<div class="tip-r a' + a + '"><span class="lk" aria-hidden="true"></span><b>' + esc(money(c.cents[a] / 100)) + '</b><span>' + esc(aptName(a)) + '</span></div>';
  });
  h += '<div class="tip-r tot"><span></span><b>' + esc(money(c.total)) + '</b><span>' + esc(tr('total')) + '</span></div>';
  tip.innerHTML = h;
  tip.hidden = false;
  var svg = el.querySelector('svg');
  var scale = svg.getBoundingClientRect().width / g.W;
  var px = xx * scale, tw = tip.offsetWidth, cw = el.clientWidth;
  var left = px + 14;
  if (left + tw > cw) left = px - tw - 14;
  tip.style.left = Math.max(0, left) + 'px';
}

function hideCrosshair() {
  var el = $('chart');
  if (!el) return;
  el._j = null;
  var xh = $('xh');
  if (xh) xh.setAttribute('visibility', 'hidden');
  var tip = $('chart-tip');
  if (tip) tip.hidden = true;
}

function renderRules() {
  var el = $('rules-panel');
  if (!el) return;
  var sh = shares(state.data.settings.surfaces);
  var last = state.data.months.length - 1, lm = state.data.months[last];
  /* which calculated months use the price per kWh and which the proportional split */
  var byPrice = [], byProp = [];
  var fee = null;
  state.data.months.forEach(function (mm, k) {
    if (num(mm.gunoi)) fee = mm.gunoi;
    if (k === 0 || mm.manual || mm.start) return;
    (num(mm.pk) ? byPrice : byProp).push(k);
  });
  /* new months use the price per kWh, so that is the rule until a month says otherwise */
  var elRule = tr('ruleKwhProp');
  if (byPrice.length || !byProp.length) {
    elRule = tr('ruleKwh');
    if (byProp.length) {
      var lastProp = byProp[byProp.length - 1];
      elRule += ' ' + (lastProp < byPrice[0] ? tr('ruleKwhOld', { m: mName(lastProp) }) : tr('ruleKwhMixed'));
    }
  }
  var h = '<h2 class="h2" id="rules-title">' + esc(tr('rulesTitle')) + '</h2><div class="rules-grid"><ul class="rules-list">';
  [['gaz', tr('ruleGaz')], ['digi', tr('rulePeople')], ['gunoi', tr('ruleGunoi', { v: lei(fee == null ? 22 : fee) })], ['curent', elRule]].forEach(function (r) {
    var tag = r[0] === 'digi' ? tr('digi') + ' · ' + tr('apa') : r[0] === 'curent' ? tr('curentP') : tr(r[0]);
    h += '<li><span class="rtag">' + esc(tag) + '</span><span>' + esc(r[1]) + '</span></li>';
  });
  h += '</ul><div><p class="eyebrow">' + esc(tr('shares')) + '</p><div class="segbar" role="img" aria-label="' +
    esc(APTS.map(function (a) { return aptName(a) + ' ' + pct(sh[a]); }).join(', ')) + '">';
  APTS.forEach(function (a) {
    h += '<span class="a' + a + '" data-w="' + (sh[a] * 100).toFixed(4) + '">' + (sh[a] >= 0.12 ? esc(pct(sh[a])) : '') + '</span>';
  });
  h += '</div><ul class="seglegend">' + APTS.map(function (a) {
    return '<li class="a' + a + '"><span class="sw" aria-hidden="true"></span>' + esc(aptName(a)) + (sh[a] < 0.12 ? ' ' + esc(pct(sh[a])) : '') + '</li>';
  }).join('') + '</ul>';
  h += '<p class="eyebrow">' + esc(tr('peopleIn', { m: mName(last) })) + '</p><div class="people">' +
    APTS.map(function (a) { return '<span class="a' + a + '"><span class="key" aria-hidden="true"></span>' + esc(aptName(a)) + ' <b>' + esc(String(lm.p[a])) + '</b></span>'; }).join('') +
    '</div></div></div>';
  el.innerHTML = h;
  Array.prototype.forEach.call(el.querySelectorAll('.segbar [data-w]'), function (sp) {
    sp.style.flex = sp.getAttribute('data-w') + ' 1 0';
  });
}

function renderFooter() {
  var el = $('foot');
  if (!el) return;
  var d = state.data;
  var info = d.updated ? tr('updated', { d: new Intl.DateTimeFormat(I18N[state.lang].locale, { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' }).format(new Date(d.updated + 'T00:00:00Z')) }) :
    tr('fromExcel', { m: mName(d.months.length - 1) });
  var h = '<div><p>' + esc(info) + '</p><p>' + esc(tr('live')) + '</p></div><div class="foot-actions">';
  if (state.downloads) h += '<button type="button" class="btn sm" data-action="export"' + (state.exporting ? ' disabled' : '') + '>' + icon('down') + esc(state.exporting ? tr('exporting') : tr('exportXlsx')) + '</button>';
  if (state.canEdit) h += '<button type="button" class="btn sm" data-action="settings">' + icon('sliders') + esc(tr('settings')) + '</button>';
  el.innerHTML = h + '</div>';
}

/* ------------------------------------------------------------------ forms */

function emptyElRows(n) {
  var rows = [];
  for (var k = 0; k < n; k++) rows.push({ l: '', k: '' });
  return rows;
}

/* New months start on the price per kWh; DIGI and rubbish carry over from the month before. */
var DEFAULT_DIGI = 98, DEFAULT_GUNOI = 22;
function draftForAdd() {
  var ms = state.data.months, last = ms[ms.length - 1];
  return {
    gaz: '', digi: num(last.digi) ? last.digiText || numToInput(last.digi) : numToInput(DEFAULT_DIGI),
    gunoi: numToInput(num(last.gunoi) ? last.gunoi : DEFAULT_GUNOI), apa: '', curent: '',
    elMode: true, el: emptyElRows(Math.min(MAX_EL, Math.max(2, last.el ? last.el.length : 0))),
    r: ['', '', '', '', ''], x: [false, false, false, false, false], o: ['', '', '', '', ''],
    p: last.p.map(String), g: ['', '', ''], note: ''
  };
}

/* A draft kept from an earlier visit may predate a field; fill in whatever is missing. */
function fixDraft(d, base) {
  var o = {};
  Object.keys(base).forEach(function (k) { o[k] = d && d[k] != null ? d[k] : base[k]; });
  ['r', 'x', 'o', 'p', 'g'].forEach(function (k) {
    if (!Array.isArray(o[k]) || o[k].length !== base[k].length) o[k] = base[k];
  });
  if (!Array.isArray(o.el) || !o.el.length) o.el = base.el;
  o.el = o.el.slice(0, MAX_EL).map(function (row) { return { l: row && row.l != null ? String(row.l) : '', k: row && row.k != null ? String(row.k) : '' }; });
  o.elMode = d && typeof d.elMode === 'boolean' ? d.elMode : base.elMode;
  return o;
}

/* A month's electricity as bill rows, for switching it to the price per kWh: the amounts it was
   entered with ("208+134" gives two bills), with the kWh still to fill in. */
function rowsFromAmount(m) {
  var parts = [];
  if (m.curentText && /^\d+(\.\d+)?(\+\d+(\.\d+)?)+$/.test(m.curentText)) parts = m.curentText.split('+').map(Number);
  else if (num(m.curent) && m.curent > 0) parts = [m.curent];
  var rows = parts.slice(0, MAX_EL).map(function (v) { return { l: numToInput(v), k: '' }; });
  return rows.length ? rows : emptyElRows(2);
}

/* Both ways of entering the electricity are filled in, so the form can switch between them. */
function draftFromMonth(m) {
  var d = { r: [], x: [], o: [], p: m.p.map(String), g: ['', '', ''], note: m.note || '' };
  BILLS.forEach(function (b) { d[b] = m[b + 'Text'] || numToInput(m[b]); });
  d.elMode = num(m.pk);
  d.el = m.el ? m.el.map(function (b) { return { l: numToInput(b.lei), k: numToInput(b.kwh) }; }) :
    d.elMode ? emptyElRows(2) : rowsFromAmount(m);
  METERS.forEach(function (j) {
    d.r[j] = numToInput(m.r[j]);
    var has = !!(m.kwh && m.kwh[j] != null);
    d.x[j] = has;
    d.o[j] = has ? numToInput(m.kwh[j]) : '';
  });
  if (m.g) d.g = m.g.map(function (v) { return v == null ? '' : numToInput(v); });
  return d;
}

function openMonthSheet(mode, index, draft) {
  var n = state.data.months.length;
  if (mode === 'add') {
    index = n;
    if (!draft) {
      try {
        var ld = JSON.parse(lsGet('ff-draft') || 'null');
        if (ld && ld.index === n && ld.draft) draft = ld.draft;
      } catch (e) { draft = null; }
    }
    draft = draft ? fixDraft(draft, draftForAdd()) : draftForAdd();
  } else {
    draft = draftFromMonth(state.data.months[index]);
  }
  state.sheet = { kind: 'month', mode: mode, index: index, draft: draft, touched: {}, submitted: false, saving: false, opener: document.activeElement };
  renderSheet();
  var first = $(mode === 'add' ? 'f-gaz' : 'f-gaz');
  if (first) first.focus();
}

function openSettingsSheet() {
  var s = state.data.settings, st = parseStart(s.start);
  state.sheet = { kind: 'settings', draft: { sm: String(st.m), sy: String(st.y), s: s.surfaces.map(numToInput) }, touched: {}, submitted: false, saving: false, opener: document.activeElement };
  renderSheet();
  var f = $('f-sm');
  if (f) f.focus();
}

function closeSheet() {
  var op = state.sheet && state.sheet.opener;
  state.sheet = null;
  renderSheet();
  if (op && document.body.contains(op)) op.focus();
}

function inputField(f, label, value, mode, extra) {
  var id = 'f-' + f;
  return '<div class="field' + (extra && extra.wide ? ' wide' : '') + '" data-wrap="' + f + '"><label for="' + id + '">' + (extra && extra.key != null ? '<span class="key a' + extra.key + '" aria-hidden="true"></span>' : '') + esc(label) + '</label>' +
    '<input type="text" id="' + id + '" data-f="' + f + '" inputmode="' + mode + '" autocomplete="off" spellcheck="false" enterkeyhint="next" value="' + esc(value == null ? '' : value) + '"' +
    (extra && extra.ph ? ' placeholder="' + esc(extra.ph) + '"' : '') + (extra && extra.max ? ' maxlength="' + extra.max + '"' : '') +
    ' aria-describedby="' + id + '-i ' + id + '-m">' +
    (extra && extra.info != null ? '<small class="info" id="' + id + '-i">' + esc(extra.info) + '</small>' : '') +
    '<small class="msg" id="' + id + '-m"></small></div>';
}

/* The electricity part of the form: a switch for how the month's electricity is split, then
   either its bills (price per kWh) or the one amount shared out in proportion (the Excel's rule). */
function elBoxHTML(d) {
  var h = '<div class="elbox"><div class="elmode"><p class="elhead" id="el-title">' + esc(tr('curentP')) + '</p>' +
    '<div class="elseg" role="radiogroup" aria-labelledby="el-title">';
  [['price', 'modePrice', true], ['prop', 'modeProp', false]].forEach(function (o) {
    h += '<label class="segopt"><input type="radio" name="elmode" id="f-elmode-' + o[0] + '" value="' + o[0] + '" data-f="elmode"' +
      (!!d.elMode === o[2] ? ' checked' : '') + '><span>' + esc(tr(o[1])) + '</span></label>';
  });
  h += '</div></div>';
  if (d.elMode) h += elBillsHTML(d) + '<p class="hint">' + esc(tr('elHint')) + '</p>';
  else h += '<div class="fgrid three">' + inputField('curent', tr('elAmount'), d.curent, 'text') + '</div><p class="hint">' + esc(tr('propHint')) + '</p>';
  return h + '</div>';
}

/* One row per electricity bill: its total (lei) and the kWh it bills. */
function elBillsHTML(d) {
  var h = '<div class="elgrid" role="group" aria-label="' + esc(tr('secEl')) + '"><span aria-hidden="true"></span><span class="elcol" aria-hidden="true">' + esc(tr('elLei')) + '</span>' +
    '<span class="elcol" aria-hidden="true">' + esc(tr('elKwh')) + '</span><span aria-hidden="true"></span>';
  d.el.forEach(function (row, n) {
    var bn = tr('billN', { n: n + 1 });
    h += '<span class="eln" aria-hidden="true">' + (n + 1) + '</span>';
    [['l', tr('elLei'), 'text'], ['k', tr('elKwh'), 'decimal']].forEach(function (c) {
      var f = 'el' + n + c[0], id = 'f-' + f;
      h += '<div class="field" data-wrap="' + f + '"><input type="text" id="' + id + '" data-f="' + f + '" inputmode="' + c[2] + '" autocomplete="off" spellcheck="false" enterkeyhint="next" value="' +
        esc(row[c[0]]) + '" aria-label="' + esc(bn + ', ' + c[1]) + '" aria-describedby="' + id + '-m"><small class="msg" id="' + id + '-m"></small></div>';
    });
    h += '<button type="button" class="icon-btn el-del" data-action="el-del" data-n="' + n + '" aria-label="' + esc(tr('elDel', { n: n + 1 })) + '">' + icon('x') + '</button>';
  });
  h += '</div>';
  if (d.el.length < MAX_EL) h += '<button type="button" class="btn sm el-add" data-action="el-add">' + icon('plus') + esc(tr('elAdd')) + '</button>';
  return h + '<p class="elprice" id="elprice"></p>';
}

function peopleFieldsetHTML(d) {
  var h = '<fieldset class="fs"><legend class="eyebrow">' + esc(tr('secPeople')) + '</legend><div class="fgrid four">';
  APTS.forEach(function (a) {
    var id = 'f-p' + a;
    h += '<div class="field" data-wrap="p' + a + '"><label for="' + id + '"><span class="key a' + a + '" aria-hidden="true"></span>' + esc(aptName(a)) + '</label><div class="step">' +
      '<button type="button" class="sbtn" data-action="step" data-f="p' + a + '" data-d="-1" aria-label="' + esc(tr('less') + ', ' + aptName(a)) + '">−</button>' +
      '<input type="text" id="' + id + '" data-f="p' + a + '" inputmode="numeric" autocomplete="off" value="' + esc(d.p[a]) + '" aria-describedby="' + id + '-m">' +
      '<button type="button" class="sbtn" data-action="step" data-f="p' + a + '" data-d="1" aria-label="' + esc(tr('more') + ', ' + aptName(a)) + '">+</button>' +
      '</div><small class="msg" id="' + id + '-m"></small></div>';
  });
  return h + '</div></fieldset>';
}

/* The starting readings (and the people living there then), for when one was copied wrong. */
function openBaseSheet() {
  var b = state.data.months[0];
  state.sheet = { kind: 'base', index: 0, draft: { r: b.r.map(numToInput), p: b.p.map(String) }, touched: {}, submitted: false, saving: false, opener: document.activeElement };
  renderSheet();
  var f = $('f-r0');
  if (f) f.focus();
}

var BASE_KEYS = ['r0', 'r1', 'r2', 'r3', 'r4', 'p0', 'p1', 'p2', 'p3'];

function baseSheetHTML(sh) {
  var d = sh.draft;
  var h = '<div class="sheet"><div class="sheet-scrim" data-action="close-sheet"></div>' +
    '<div class="sheet-card" role="dialog" aria-modal="true" aria-labelledby="sheet-title">' +
    '<div class="sheet-head"><h2 class="h2" id="sheet-title">' + esc(tr('fBase', { m: mName(0) })) + '</h2><button type="button" class="icon-btn" data-action="close-sheet" aria-label="' + esc(tr('close')) + '">' + icon('x') + '</button></div>' +
    '<form class="sheet-body" id="bform" novalidate>';
  h += '<fieldset class="fs"><legend class="eyebrow">' + eyebrow(tr('secReadings')) + '</legend><div class="fgrid five">';
  METERS.forEach(function (j) { h += inputField('r' + j, meterName(j), d.r[j], 'decimal', { key: j < 4 ? j : null }); });
  h += '</div></fieldset>' + peopleFieldsetHTML(d);
  h += '<ul class="pv-global" id="base-global"></ul><p class="sheet-msg" id="sheet-msg" role="alert"></p>';
  h += '<div class="sheet-actions"><button type="submit" class="btn primary" id="save-btn">' + icon('check') + esc(tr('baseSave')) + '</button>' +
    '<button type="button" class="btn" data-action="close-sheet">' + esc(tr('close')) + '</button></div>';
  return h + '</form></div></div>';
}

function refreshBase() {
  var sh = state.sheet, v = validateStart(state.data, sh.draft);
  sh.v = v;
  BASE_KEYS.forEach(function (f) { setFieldMsg(f, v.errors[f], null, sh.submitted || sh.touched[f]); });
  var g = $('base-global');
  if (g) g.innerHTML = v.global.map(function (x) { return '<li>' + esc(errText(x)) + '</li>'; }).join('');
}

function submitBase() {
  var sh = state.sheet;
  if (!sh || sh.saving) return;
  sh.submitted = true;
  refreshSheet();
  if (!sh.v.month) { focusFirstInvalid(); return; }
  var nd = clone(state.data);
  var m = clone(sh.v.month);
  m.t = todayISO();
  nd.months[0] = m;
  saveData(nd, { type: 'base', index: 0 });
}

function monthSheetHTML(sh) {
  var i = sh.index, d = sh.draft, n = state.data.months.length;
  var prev = i > 0 ? state.data.months[i - 1] : null;
  var title = cap(tr(sh.mode === 'add' ? 'fAdd' : 'fEdit', { m: mName(i) }));
  var h = '<div class="sheet"><div class="sheet-scrim" data-action="close-sheet"></div>' +
    '<div class="sheet-card" role="dialog" aria-modal="true" aria-labelledby="sheet-title">' +
    '<div class="sheet-head"><h2 class="h2" id="sheet-title">' + esc(title) + '</h2><button type="button" class="icon-btn" data-action="close-sheet" aria-label="' + esc(tr('close')) + '">' + icon('x') + '</button></div>' +
    '<form class="sheet-body" id="mform" novalidate>';
  h += '<fieldset class="fs"><legend class="eyebrow">' + esc(tr('secBills')) + '</legend><div class="fgrid four">';
  ['gaz', 'digi', 'gunoi', 'apa'].forEach(function (b) { h += inputField(b, tr(b === 'gunoi' ? 'gunoiField' : b), d[b], 'text'); });
  h += '</div>' + elBoxHTML(d) + '<p class="hint">' + esc(tr('exprHint')) + '</p></fieldset>';
  h += '<fieldset class="fs"><legend class="eyebrow">' + eyebrow(tr('secReadings')) + '</legend><div class="fgrid five">';
  METERS.forEach(function (j) {
    var id = 'f-r' + j;
    h += '<div class="field" data-wrap="r' + j + '"><label for="' + id + '">' + (j < 4 ? '<span class="key a' + j + '" aria-hidden="true"></span>' : '') + esc(meterName(j)) + '</label>' +
      '<input type="text" id="' + id + '" data-f="r' + j + '" inputmode="decimal" autocomplete="off" spellcheck="false" enterkeyhint="next" value="' + esc(d.r[j]) + '" aria-describedby="' + id + '-i ' + id + '-m">' +
      '<small class="info" id="' + id + '-i">' + esc(prev ? tr('lastRead', { v: kwhStr(prev.r[j]) }) : '') + '</small>' +
      '<small class="msg" id="' + id + '-m"></small>' +
      '<label class="check" id="xw' + j + '"' + (d.x[j] ? '' : ' hidden') + '><input type="checkbox" data-f="x' + j + '"' + (d.x[j] ? ' checked' : '') + '>' + esc(tr('replaced')) + '</label>' +
      '<div class="override" id="ow' + j + '"' + (d.x[j] ? '' : ' hidden') + '><label for="f-o' + j + '">' + esc(tr('overrideLbl')) + '</label>' +
      '<input type="text" id="f-o' + j + '" data-f="o' + j + '" inputmode="decimal" autocomplete="off" value="' + esc(d.o[j]) + '" aria-describedby="f-o' + j + '-m"><small class="msg" id="f-o' + j + '-m"></small></div></div>';
  });
  h += '</div></fieldset>';
  h += peopleFieldsetHTML(d);
  h += '<div class="fs"><div class="fgrid">' + inputField('note', tr('secNote'), d.note, 'text', { wide: true, ph: tr('notePh'), max: 200 }) + '</div></div>';
  h += '<details class="gen"' + (d.g.some(function (v) { return v; }) ? ' open' : '') + '><summary>' + esc(tr('secGeneral')) + '</summary><div class="fgrid three">' +
    ['genGaz', 'genSus', 'genJos'].map(function (k, j) { return inputField('g' + j, tr(k), d.g[j], 'decimal'); }).join('') + '</div></details>';
  h += '<div class="preview" id="preview" aria-live="polite"></div>';
  h += '<p class="sheet-msg" id="sheet-msg" role="alert"></p>';
  h += '<div class="sheet-actions"><button type="submit" class="btn primary" id="save-btn">' + icon('check') + esc(tr('save')) + '</button>' +
    '<button type="button" class="btn" data-action="close-sheet">' + esc(tr('close')) + '</button>' +
    (sh.mode === 'edit' && i === n - 1 && n > 1 ? '<button type="button" class="btn danger" data-action="ask-delete">' + esc(tr('del')) + '</button>' : '') + '</div>';
  h += '<div class="confirm" id="confirm" hidden><p>' + esc(tr('delQ', { m: mName(i) })) + '</p>' +
    '<button type="button" class="btn danger solid" data-action="del-yes">' + esc(tr('delYes')) + '</button><button type="button" class="btn" data-action="del-no">' + esc(tr('delNo')) + '</button></div>';
  return h + '</form></div></div>';
}

function settingsSheetHTML(sh) {
  var d = sh.draft;
  var h = '<div class="sheet"><div class="sheet-scrim" data-action="close-sheet"></div>' +
    '<div class="sheet-card" role="dialog" aria-modal="true" aria-labelledby="sheet-title">' +
    '<div class="sheet-head"><h2 class="h2" id="sheet-title">' + esc(tr('setTitle')) + '</h2><button type="button" class="icon-btn" data-action="close-sheet" aria-label="' + esc(tr('close')) + '">' + icon('x') + '</button></div>' +
    '<form class="sheet-body" id="sform" novalidate>';
  var mf = new Intl.DateTimeFormat(I18N[state.lang].locale, { month: 'long', timeZone: 'UTC' });
  h += '<fieldset class="fs"><legend class="eyebrow">' + esc(tr('setStart')) + '</legend><div class="fgrid">';
  h += '<div class="field" data-wrap="sm"><label for="f-sm">' + esc(tr('setMonth')) + '</label><select id="f-sm" data-f="sm">';
  for (var mo = 0; mo < 12; mo++) h += '<option value="' + mo + '"' + (String(mo) === d.sm ? ' selected' : '') + '>' + esc(cap(mf.format(new Date(Date.UTC(2020, mo, 1))))) + '</option>';
  h += '</select></div><div class="field" data-wrap="sy"><label for="f-sy">' + esc(tr('setYear')) + '</label><select id="f-sy" data-f="sy">';
  var cy = parseStart(state.data.settings.start).y;
  for (var yr = Math.min(2015, cy); yr <= Math.max(2030, cy); yr++) h += '<option value="' + yr + '"' + (String(yr) === d.sy ? ' selected' : '') + '>' + yr + '</option>';
  h += '</select><small class="msg" id="f-sy-m"></small></div></div><p class="hint">' + esc(tr('setStartHint')) + '</p>';
  var skip = state.data.settings.skip || [];
  if (skip.length) {
    var lf = new Intl.DateTimeFormat(I18N[state.lang].locale, { month: 'long', year: 'numeric', timeZone: 'UTC' });
    h += '<p class="hint">' + esc(tr('setSkip', { m: skip.map(function (s) { var p = parseYm(s); return lf.format(new Date(Date.UTC(p.y, p.m, 1))); }).join(', ') })) + '</p>';
  }
  h += '<p class="hint" id="last-would-be"></p></fieldset>';
  h += '<fieldset class="fs"><legend class="eyebrow">' + esc(tr('setSurf')) + '</legend><div class="fgrid four">';
  APTS.forEach(function (a) { h += inputField('s' + a, aptName(a), d.s[a], 'decimal', { key: a }); });
  h += '</div><p class="shares-pv" id="shares-pv"></p><p class="hint">' + esc(tr('setSurfHint')) + '</p></fieldset>';
  h += '<p class="sheet-msg" id="sheet-msg" role="alert"></p>';
  h += '<div class="sheet-actions"><button type="submit" class="btn primary" id="save-btn">' + icon('check') + esc(tr('setSave')) + '</button>' +
    '<button type="button" class="btn" data-action="close-sheet">' + esc(tr('close')) + '</button></div>';
  return h + '</form></div></div>';
}

function renderSheet() {
  var host = $('sheet-host');
  if (!host) return;
  var sh = state.sheet;
  if (!sh) {
    host.innerHTML = '';
    document.body.classList.remove('locked');
    return;
  }
  document.body.classList.add('locked');
  host.innerHTML = sh.kind === 'month' ? monthSheetHTML(sh) : sh.kind === 'base' ? baseSheetHTML(sh) : settingsSheetHTML(sh);
  refreshSheet();
  setSaving(sh.saving);
}

/* Redraws the open form (after a bill row is added or removed) without losing the scroll position. */
function rerenderSheet(focusId) {
  var card = root.querySelector('.sheet-card');
  var top = card ? card.scrollTop : 0;
  renderSheet();
  var card2 = root.querySelector('.sheet-card');
  if (card2) card2.scrollTop = top;
  var f = focusId && $(focusId);
  if (f) f.focus({ preventScroll: true });
}

function errText(e) {
  var c = e.code;
  if (c === 'lower') return tr('eLower', { v: kwhStr(e.v) });
  if (c === 'nextLower') return tr('eNextLower', { m: mName(e.i), x: meterName(e.j) });
  var map = {
    required: 'eRequired', number: 'eNumber', plain: 'ePlain', thousands: 'eThousands', tooBig: 'eTooBig', negative: 'eNegative',
    kwhBig: 'eKwhBig', override: 'eOverride', persons: 'ePersons', personsZero: 'ePersonsZero', kwhZero: 'eKwhZero', note: 'eNote', positive: 'ePositive',
    elNeeded: 'eElNeeded', elKwh: 'eElKwh', elPrice: 'eElPrice'
  };
  return tr(map[c] || 'eNumber');
}
function warnText(w) {
  if (w.code === 'priceOff') return tr('wPriceOff', { v: priceStr(w.v) });
  return w.code === 'billHigh' ? tr('wBillHigh', { v: lei(w.v) }) : tr('wKwhHigh', { v: kwhStr(w.v) });
}

function setFieldMsg(f, e, w, show) {
  var input = $('f-' + f), msg = $('f-' + f + '-m');
  if (!msg) return;
  if (e && show) {
    msg.textContent = errText(e);
    msg.className = 'msg bad';
    if (input) input.setAttribute('aria-invalid', 'true');
  } else {
    if (w && show) {
      msg.textContent = warnText(w);
      msg.className = 'msg warn';
    } else {
      msg.textContent = '';
      msg.className = 'msg';
    }
    if (input) input.removeAttribute('aria-invalid');
  }
}

/* The form's fields in screen order, for messages and for focusing the first problem. */
function formKeys(d) {
  var keys = d.elMode ? ['gaz', 'digi', 'gunoi', 'apa'] : BILLS.slice();
  if (d.elMode) d.el.forEach(function (_, n) { keys.push('el' + n + 'l', 'el' + n + 'k'); });
  return keys.concat(['r0', 'r1', 'r2', 'r3', 'r4', 'o0', 'o1', 'o2', 'o3', 'o4', 'p0', 'p1', 'p2', 'p3', 'note', 'g0', 'g1', 'g2']);
}

function refreshElPrice(sh, v) {
  var el = $('elprice');
  if (!el) return;
  var e = v.el, t = '', cls = 'elprice';
  if (e && e.pk != null && e.from == null) t = tr('elPriceLive', { p: priceStr(e.pk), L: lei(e.L), K: kwhTotal(e.K) });
  else if (e && e.pk != null) t = tr('elPriceCarry', { m: mName(e.from), p: priceStr(e.pk) });
  if (v.warnings.el && t) { t += ' ' + warnText(v.warnings.el); cls += ' warn'; }
  el.textContent = t;
  el.className = cls;
}

function refreshSheet() {
  var sh = state.sheet;
  if (!sh) return;
  if (sh.kind === 'settings') return refreshSettings();
  if (sh.kind === 'base') return refreshBase();
  var v = validateDraft(state.data, sh.index, sh.draft);
  sh.v = v;
  formKeys(sh.draft).forEach(function (f) {
    setFieldMsg(f, v.errors[f], v.warnings[f], sh.submitted || sh.touched[f]);
  });
  if (sh.draft.elMode) refreshElPrice(sh, v);
  var prev = sh.index > 0 ? state.data.months[sh.index - 1] : null;
  METERS.forEach(function (j) {
    var er = v.errors['r' + j];
    var showX = sh.draft.x[j] || (er && er.code === 'lower');
    var xw = $('xw' + j), ow = $('ow' + j), info = $('f-r' + j + '-i');
    if (xw) xw.hidden = !showX;
    if (ow) ow.hidden = !sh.draft.x[j];
    if (info) {
      var t = prev ? tr('lastRead', { v: kwhStr(prev.r[j]) }) : '';
      if (v.usage[j] != null) t += ' · ' + tr('usageNow', { v: kwhStr(v.usage[j]) });
      info.textContent = t;
    }
  });
  var el = $('preview');
  if (!el) return;
  var errs = Object.keys(v.errors).length + v.global.length;
  var anyTouched = sh.submitted || Object.keys(sh.touched).length > 0;
  var h = '<p class="eyebrow">' + esc(tr('preview')) + ' · ' + esc(cap(mName(sh.index))) + '</p>';
  var c = v.comp;
  h += '<div class="pv-grid">' + APTS.map(function (a) {
    return '<div class="pv a' + a + '"><span class="pv-n"><span class="key" aria-hidden="true"></span>' + esc(aptName(a)) + '</span><b>' + esc(c ? money(c.cents[a] / 100) : '—') + '</b></div>';
  }).join('') + '<div class="pv"><span class="pv-n">' + esc(tr('total')) + '</span><b>' + esc(c ? lei(c.total) : '—') + '</b></div></div>';
  if (!errs) h += '<p class="pv-status ok">' + icon('check') + esc(tr('ready')) + '</p>';
  else if (!anyTouched && (sh.mode === 'add' || !v.global.length)) h += '<p class="pv-status muted">' + esc(tr('fillIn')) + '</p>';
  else {
    h += '<p class="pv-status bad">' + esc(errs === 1 ? tr('problems1') : tr('problemsN', { n: errs })) + '</p>';
    if (v.global.length) h += '<ul class="pv-global">' + v.global.map(function (g) { return '<li>' + esc(errText(g)) + '</li>'; }).join('') + '</ul>';
  }
  el.innerHTML = h;
}

function refreshSettings() {
  var sh = state.sheet, v = validateSettings(sh.draft);
  sh.v = v;
  APTS.forEach(function (a) { setFieldMsg('s' + a, v.errors['s' + a], null, sh.submitted || sh.touched['s' + a]); });
  setFieldMsg('sy', v.errors.sy, null, true);
  var pv = $('shares-pv');
  if (pv) {
    var vals = APTS.map(function (a) { var r = parseExpr(sh.draft.s[a], 'plain'); return r.value > 0 ? r.value : 0; });
    var s2 = shares(vals);
    pv.textContent = sum(vals) > 0 ? tr('usedShares', { v: APTS.map(function (a) { return aptName(a) + ' ' + pct(s2[a]); }).join(' · ') }) : '';
  }
  var lw = $('last-would-be');
  if (lw && v.settings) {
    var d = monthAt({ start: v.settings.start, skip: state.data.settings.skip }, state.data.months.length - 1);
    lw.textContent = tr('setLastWouldBe', { m: new Intl.DateTimeFormat(I18N[state.lang].locale, { month: 'long', year: 'numeric', timeZone: 'UTC' }).format(new Date(Date.UTC(d.y, d.m, 1))) });
  }
}

var draftTimer = null;
function saveLocalDraft() {
  var sh = state.sheet;
  if (!sh || sh.kind !== 'month' || sh.mode !== 'add') return;
  clearTimeout(draftTimer);
  draftTimer = setTimeout(function () {
    lsSet('ff-draft', JSON.stringify({ index: sh.index, draft: sh.draft }));
  }, 300);
}

function setSheetMsg(text, kind) {
  var el = $('sheet-msg');
  if (!el) return;
  el.textContent = text || '';
  el.className = 'sheet-msg' + (kind ? ' ' + kind : '');
}

function setSaving(on) {
  var sh = state.sheet;
  if (sh) sh.saving = !!on;
  var b = $('save-btn');
  if (!b) return;
  b.disabled = !!on;
  var label = tr(sh && sh.kind === 'settings' ? 'setSave' : sh && sh.kind === 'base' ? 'baseSave' : 'save');
  b.innerHTML = on ? esc(tr('saving')) : icon('check') + esc(label);
}

function focusFirstInvalid() {
  var sh = state.sheet;
  if (!sh || !sh.v) return;
  var keys = sh.kind === 'settings' ? ['s0', 's1', 's2', 's3', 'sy'] : sh.kind === 'base' ? BASE_KEYS : formKeys(sh.draft);
  for (var k = 0; k < keys.length; k++) {
    if (sh.v.errors[keys[k]]) {
      var el = $('f-' + keys[k]);
      if (el) { el.focus(); return; }
    }
  }
  var pv = $('preview');
  if (pv) pv.scrollIntoView({ block: 'nearest' });
}

function submitMonth() {
  var sh = state.sheet;
  if (!sh || sh.saving) return;
  sh.submitted = true;
  refreshSheet();
  var v = sh.v;
  if (!v.month) { focusFirstInvalid(); return; }
  var nd = clone(state.data);
  var m = clone(v.month);
  m.t = todayISO();
  if (sh.mode === 'add') nd.months.push(m);
  else nd.months[sh.index] = m;
  saveData(nd, { type: 'saved', index: sh.index });
}

function submitSettings() {
  var sh = state.sheet;
  if (!sh || sh.saving) return;
  sh.submitted = true;
  refreshSheet();
  if (!sh.v.settings) { focusFirstInvalid(); return; }
  var nd = clone(state.data);
  nd.settings = { start: sh.v.settings.start, surfaces: sh.v.settings.surfaces, skip: (state.data.settings.skip || []).slice() };
  saveData(nd, { type: 'settings' });
}

function skeletonReset() {
  var styles = document.head ? document.head.querySelectorAll('style') : [];
  for (var k = 0; k < styles.length; k++) {
    var t = styles[k].textContent || '';
    if (t.indexOf('[hidden]') !== -1 && t.indexOf('safe-area-inset') !== -1) return t;
  }
  return FALLBACK_RESET;
}

function buildDocument(nd) {
  var css = $('ff-style'), js = $('ff-app');
  if (!css || !js) throw new Error('source missing');
  var json = serializeData(nd);
  var back = normalizeData(JSON.parse(json));
  if (back.months.length !== nd.months.length) throw new Error('round trip');
  var html = fullDocument(skeletonReset(), pageContent(css.textContent, json, js.textContent));
  if (html.indexOf('id="ff-app"') === -1 || html.length > 3000000) throw new Error('bad document');
  return html;
}

function saveData(nd, after) {
  var sh = state.sheet;
  if (!state.artifact) { setSheetMsg(tr('unavailable'), 'bad'); return Promise.resolve(); }
  nd.updated = todayISO();
  var html;
  try { html = buildDocument(nd); } catch (e) { setSheetMsg(tr('saveFailed'), 'bad'); return Promise.resolve(); }
  ssSet('ff-after', JSON.stringify(after));
  if (sh && sh.kind === 'month' && sh.mode === 'add') ssSet('ff-stash', JSON.stringify({ index: sh.index, draft: sh.draft }));
  setSaving(true);
  setSheetMsg('', '');
  var attempt = 0;
  function go() {
    return state.artifact.publish(html).then(function () {
      lsDel('ff-draft');
      ssDel('ff-stash');
      /* every open view, this one included, reloads to the new version; if this one has not
         reloaded after a few seconds, show the saved data in place */
      setTimeout(function () { applyInPlace(nd, after); }, 6000);
    }, function (e) {
      var code = (e && e.code) || 'upstream_error';
      if (code === 'upstream_error' && attempt === 0) {
        attempt++;
        return sleep(800 + Math.random() * 1200).then(go);
      }
      ssDel('ff-after');
      setSaving(false);
      if (code === 'conflict') { setSheetMsg(tr('conflict'), 'warn'); return; }
      ssDel('ff-stash');
      if (code === 'not_writer') {
        setSheetMsg(tr('notWriter'), 'bad');
        state.userCanEdit = false;
      } else if (code === 'not_granted' || code === 'not_declared' || code === 'consent_required' || code === 'capability_disabled' || code === 'capability_removed') {
        setSheetMsg(tr('unavailable'), 'bad');
      } else if (code === 'rate_limited') {
        setSheetMsg(tr('rateLimited'), 'warn');
      } else {
        setSheetMsg(tr('saveFailed'), 'bad');
      }
    });
  }
  return go();
}

function applyInPlace(nd, after) {
  if (!state.data || state.data === nd) return;
  ssDel('ff-after');
  state.data = normalizeData(nd);
  state.comp = computeAll(state.data);
  state.sheet = null;
  if (after && typeof after.index === 'number' && after.index < state.data.months.length) state.sel = after.index;
  else state.sel = state.data.months.length - 1;
  renderAll();
  document.body.classList.remove('locked');
  showToast(afterMessage(after));
}

/* ------------------------------------------------------------------ export */

var xlsxPromise = null;
function loadXlsx() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  if (xlsxPromise) return xlsxPromise;
  xlsxPromise = new Promise(function (resolve, reject) {
    var urls = XLSX_URLS.slice();
    (function next() {
      var u = urls.shift();
      if (!u) { xlsxPromise = null; reject(new Error('load')); return; }
      var s = document.createElement('script');
      s.src = u;
      s.async = true;
      s.onload = function () { if (window.XLSX) resolve(window.XLSX); else next(); };
      s.onerror = function () { s.remove(); next(); };
      document.head.appendChild(s);
    })();
  });
  return xlsxPromise;
}

function exportXlsx() {
  if (!state.downloads || state.exporting) return;
  state.exporting = true;
  renderFooter();
  loadXlsx().then(function (X) {
    var M2 = '#,##0.00';
    /* one column per entry: header, number format, value for month i */
    /* bills and amounts stay empty on the starting readings */
    function paid(f) { return function (m, c, i) { if (c.start) return ''; var v = f(m, c, i); return v == null ? '' : v; }; }
    var cols = [
      [tr('month'), 'mmm yyyy', function (m, c, i) { var d = ym(i); return (Date.UTC(d.y, d.m, 1) - Date.UTC(1899, 11, 30)) / 86400000; }],
      [tr('gaz'), M2, paid(function (m) { return m.gaz; })],
      [tr('digi'), M2, paid(function (m) { return m.digi; })],
      [tr('gunoiField'), M2, paid(function (m) { return num(m.gunoi) ? m.gunoi : ''; })],
      [tr('apa'), M2, paid(function (m) { return m.apa; })],
      [tr('xElLei'), M2, paid(function (m) { return m.curent; })],
      [tr('xElKwh'), '#,##0', paid(function (m) { return m.el ? sum(m.el.map(function (b) { return b.kwh; })) : ''; })],
      [tr('xPrice'), '0.0000', paid(function (m) { return num(m.pk) ? m.pk : ''; })],
      [tr('xElDue'), M2, paid(function (m, c) { return c.manual ? '' : round2(c.elDue); })]
    ];
    METERS.forEach(function (j) { cols.push([tr('reading') + ' ' + meterName(j), null, function (m) { return m.r[j]; }]); });
    METERS.forEach(function (j) { cols.push([tr('usage') + ' ' + meterName(j) + ' (kWh)', null, paid(function (m, c) { return c.manual ? '' : c.kwh[j]; })]); });
    APTS.forEach(function (a) { cols.push([tr('persons') + ' ' + aptName(a), null, function (m) { return m.p[a]; }]); });
    APTS.forEach(function (a) { cols.push([aptName(a) + ' (lei)', M2, paid(function (m, c) { return c.cents[a] / 100; })]); });
    cols.push([tr('total') + ' (lei)', M2, paid(function (m, c) { return c.total; })]);
    cols.push([tr('note'), null, function (m) { return m.note || (m.start ? tr('stBase') : ''); }]);
    var head = cols.map(function (k) { return k[0]; });
    var rows = [head];
    state.data.months.forEach(function (m, i) {
      var c = state.comp[i];
      rows.push(cols.map(function (k) { return k[2](m, c, i); }));
    });
    var ws = X.utils.aoa_to_sheet(rows);
    for (var r = 1; r < rows.length; r++) {
      for (var col = 0; col < cols.length; col++) {
        var ref = X.utils.encode_cell({ r: r, c: col });
        if (cols[col][1] && ws[ref] && ws[ref].t === 'n') ws[ref].z = cols[col][1];
      }
    }
    ws['!cols'] = head.map(function (t, k) { return { wch: k === 0 ? 10 : Math.min(28, Math.max(9, String(t).length + 2)) }; });
    var wb = X.utils.book_new();
    X.utils.book_append_sheet(wb, ws, 'Lunar');
    var sh = shares(state.data.settings.surfaces);
    var ws2 = X.utils.aoa_to_sheet([
      [''].concat(APTS.map(aptName)),
      [tr('setSurf')].concat(state.data.settings.surfaces),
      [tr('share')].concat(sh)
    ]);
    ['B3', 'C3', 'D3', 'E3'].forEach(function (k) { if (ws2[k]) ws2[k].z = '0.00%'; });
    ws2['!cols'] = [{ wch: 26 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }];
    X.utils.book_append_sheet(wb, ws2, 'Setari');
    var buf = X.write(wb, { bookType: 'xlsx', type: 'array' });
    return state.downloads.save({ filename: 'facturi-fierari.xlsx', data: buf });
  }).then(function () {
    showToast(tr('exportedT'));
  }, function (e) {
    if (!(e && e.code === 'declined')) showToast(tr('exportFail'));
  }).then(function () {
    state.exporting = false;
    renderFooter();
  });
}

/* ------------------------------------------------------------------ events */

var toastTimer = null;
function showToast(msg) {
  var el = $('toast');
  if (!el || !msg) return;
  el.textContent = msg;
  el.classList.add('on');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () { el.classList.remove('on'); }, 4500);
}

function setApt(a) {
  state.apt = a;
  lsSet('ff-apt', a == null ? '' : String(a));
  try { location.replace(a == null ? '#all' : '#ap-' + (a + 1)); } catch (e) { /* hash not writable here */ }
  var y = window.scrollY;
  renderAll();
  window.scrollTo(0, y);
}

function selectMonth(i, scroll) {
  state.sel = Math.max(0, Math.min(state.data.months.length - 1, i));
  renderMonth();
  renderHistory();
  if (scroll && window.innerWidth < 1000) {
    var p = $('month-panel');
    if (p) p.scrollIntoView({ block: 'start', behavior: window.matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth' });
  }
}

function onClick(ev) {
  var t = ev.target.closest('[data-action]');
  if (!t || !root.contains(t)) return;
  var act = t.getAttribute('data-action');
  if (act === 'lang') {
    state.lang = t.getAttribute('data-v') === 'en' ? 'en' : 'ro';
    lsSet('ff-lang', state.lang);
    document.documentElement.lang = state.lang;
    renderAll();
  } else if (act === 'apt') {
    var v = t.getAttribute('data-v');
    setApt(v === '' ? null : +v);
  } else if (act === 'prev') selectMonth(state.sel - 1);
  else if (act === 'next') selectMonth(state.sel + 1);
  else if (act === 'pick') selectMonth(+t.getAttribute('data-i'), true);
  else if (act === 'range') {
    state.range = +t.getAttribute('data-v');
    lsSet('ff-range', String(state.range));
    renderHistory();
  } else if (act === 'hist-all') {
    state.histAll = !state.histAll;
    renderHistory();
  } else if (act === 'add') { if (state.canEdit) openMonthSheet('add'); }
  else if (act === 'edit') { if (state.canEdit && state.sel > 0 && !state.data.months[state.sel].start) openMonthSheet('edit', state.sel); }
  else if (act === 'edit-base') { if (state.canEdit && state.data.months[0].start) openBaseSheet(); }
  else if (act === 'settings') { if (state.canEdit) openSettingsSheet(); }
  else if (act === 'close-sheet') { if (!(state.sheet && state.sheet.saving)) closeSheet(); }
  else if (act === 'step') {
    var f = t.getAttribute('data-f'), a2 = +f.slice(1), sh = state.sheet;
    if (!sh || (sh.kind !== 'month' && sh.kind !== 'base')) return;
    var cur = parseInt(sh.draft.p[a2], 10);
    if (!(cur >= 0)) cur = 0;
    var nv = Math.max(0, Math.min(20, cur + (+t.getAttribute('data-d'))));
    sh.draft.p[a2] = String(nv);
    sh.touched[f] = true;
    var inp = $('f-' + f);
    if (inp) inp.value = String(nv);
    refreshSheet();
    saveLocalDraft();
  } else if (act === 'el-add' || act === 'el-del') {
    var s3 = state.sheet;
    if (!s3 || s3.kind !== 'month' || !s3.draft.elMode || s3.saving) return;
    var rows = s3.draft.el, focusId;
    if (act === 'el-add') {
      if (rows.length >= MAX_EL) return;
      rows.push({ l: '', k: '' });
      focusId = 'f-el' + (rows.length - 1) + 'l';
    } else {
      var at = +t.getAttribute('data-n');
      if (rows.length > 1) rows.splice(at, 1);
      else rows[0] = { l: '', k: '' };
      /* messages follow their row */
      Object.keys(s3.touched).forEach(function (k) {
        var mt = /^el(\d+)([lk])$/.exec(k);
        if (!mt) return;
        delete s3.touched[k];
        var r = +mt[1];
        if (r > at) s3.touched['el' + (r - 1) + mt[2]] = true;
        else if (r < at) s3.touched[k] = true;
      });
      focusId = rows.length ? 'f-el' + Math.min(at, rows.length - 1) + 'l' : null;
    }
    rerenderSheet(focusId);
    saveLocalDraft();
  } else if (act === 'ask-delete') {
    var cf = $('confirm');
    if (cf) { cf.hidden = false; var y2 = cf.querySelector('[data-action="del-yes"]'); if (y2) y2.focus(); }
  } else if (act === 'del-no') {
    var cf2 = $('confirm');
    if (cf2) cf2.hidden = true;
  } else if (act === 'del-yes') {
    var s2 = state.sheet;
    if (!s2 || s2.saving || s2.kind !== 'month') return;
    var nd = clone(state.data);
    nd.months.pop();
    saveData(nd, { type: 'deleted', index: nd.months.length });
  } else if (act === 'export') exportXlsx();
}

function onInput(ev) {
  var t = ev.target, f = t.getAttribute && t.getAttribute('data-f');
  var sh = state.sheet;
  if (!f || !sh) return;
  var d = sh.draft;
  if (sh.kind === 'base') {
    if (/^r[0-4]$/.test(f)) d.r[+f.slice(1)] = t.value;
    else if (/^p[0-3]$/.test(f)) d.p[+f.slice(1)] = t.value;
    else return;
    if (ev.type === 'change' || String(t.value).trim() !== '') sh.touched[f] = true;
    refreshSheet();
    return;
  }
  if (sh.kind === 'settings') {
    if (f === 'sm' || f === 'sy') d[f] = t.value;
    else if (/^s[0-3]$/.test(f)) d.s[+f.slice(1)] = t.value;
    sh.touched[f] = true;
    refreshSheet();
    return;
  }
  if (f === 'elmode') {
    var toPrice = t.value === 'price';
    if (t.checked && toPrice !== !!d.elMode) {
      d.elMode = toPrice;
      rerenderSheet('f-elmode-' + t.value);
      saveLocalDraft();
    }
    return;
  }
  var mEl = /^el(\d+)([lk])$/.exec(f);
  if (BILLS.indexOf(f) !== -1 || f === 'note') d[f] = t.value;
  else if (mEl) { if (d.el[+mEl[1]]) d.el[+mEl[1]][mEl[2]] = t.value; else return; }
  else if (/^r[0-4]$/.test(f)) d.r[+f.slice(1)] = t.value;
  else if (/^o[0-4]$/.test(f)) d.o[+f.slice(1)] = t.value;
  else if (/^x[0-4]$/.test(f)) d.x[+f.slice(1)] = !!t.checked;
  else if (/^p[0-3]$/.test(f)) d.p[+f.slice(1)] = t.value;
  else if (/^g[0-2]$/.test(f)) d.g[+f.slice(1)] = t.value;
  else return;
  if (ev.type === 'change' || t.type === 'checkbox' || String(t.value).trim() !== '') sh.touched[f] = true;
  refreshSheet();
  saveLocalDraft();
}

/* Messages that appear when a field is left can push the button being clicked out from under the
   pointer, so while a pointer is pressed they wait until the click is done. */
var pointerDown = false, refreshAfterClick = false;
function onPointerDown() { pointerDown = true; }
function onPointerUp() {
  pointerDown = false;
  if (refreshAfterClick) {
    refreshAfterClick = false;
    setTimeout(refreshSheet, 0);
  }
}

function onFocusOut(ev) {
  var t = ev.target, f = t.getAttribute && t.getAttribute('data-f');
  var sh = state.sheet;
  if (!f || !sh) return;
  var keys = [f];
  /* leaving a bill row with one half filled in flags the other half too */
  var me = /^el(\d+)([lk])$/.exec(f);
  if (me && String(t.value).trim() !== '') {
    var other = 'el' + me[1] + (me[2] === 'l' ? 'k' : 'l');
    var rel = ev.relatedTarget;
    if (!(rel && rel.getAttribute && rel.getAttribute('data-f') === other)) keys.push(other);
  }
  var changed = false;
  keys.forEach(function (k) { if (!sh.touched[k]) { sh.touched[k] = true; changed = true; } });
  if (!changed) return;
  if (pointerDown) { refreshAfterClick = true; return; }
  refreshSheet();
}

function onSubmit(ev) {
  if (ev.target.id === 'mform') { ev.preventDefault(); submitMonth(); }
  else if (ev.target.id === 'sform') { ev.preventDefault(); submitSettings(); }
  else if (ev.target.id === 'bform') { ev.preventDefault(); submitBase(); }
}

function onKey(ev) {
  if (!state.sheet) return;
  if (ev.key === 'Escape') {
    var cf = $('confirm');
    if (cf && !cf.hidden) cf.hidden = true;
    else if (!state.sheet.saving) closeSheet();
    ev.preventDefault();
    return;
  }
  if (ev.key === 'Tab') {
    var card = root.querySelector('.sheet-card');
    if (!card) return;
    var items = Array.prototype.filter.call(card.querySelectorAll('button, input, select, summary, [tabindex]'), function (e) {
      return !e.disabled && e.offsetParent !== null;
    });
    if (!items.length) return;
    var first = items[0], last = items[items.length - 1];
    if (ev.shiftKey && document.activeElement === first) { last.focus(); ev.preventDefault(); }
    else if (!ev.shiftKey && document.activeElement === last) { first.focus(); ev.preventDefault(); }
  }
}

function bindEvents() {
  root.addEventListener('click', onClick);
  root.addEventListener('input', onInput);
  root.addEventListener('change', onInput);
  root.addEventListener('focusout', onFocusOut);
  root.addEventListener('pointerdown', onPointerDown, true);
  document.addEventListener('pointerup', onPointerUp, true);
  document.addEventListener('pointercancel', onPointerUp, true);
  root.addEventListener('submit', onSubmit);
  document.addEventListener('keydown', onKey);
  window.addEventListener('hashchange', function () {
    var h = aptFromHash();
    if (h !== undefined && h !== state.apt) {
      state.apt = h;
      lsSet('ff-apt', h == null ? '' : String(h));
      renderAll();
    }
  });
  var lastW = root.clientWidth, rt = null;
  var onResize = function () {
    if (root.clientWidth === lastW) return;
    lastW = root.clientWidth;
    clearTimeout(rt);
    rt = setTimeout(function () { hideCrosshair(); renderChart(); }, 80);
  };
  if (typeof ResizeObserver === 'function') new ResizeObserver(onResize).observe(root);
  else window.addEventListener('resize', onResize);
}

if (typeof document !== 'undefined' && document.getElementById('ff-data')) {
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
}


/* ====================================================================== website layer
   The same app, served by GitHub Pages at fierari.avrusu.eu. This file comes after the app's
   own script and replaces its Claude-specific parts (boot, saving, footer).
   - The months live in data.json, in the open: anyone with the address sees the bills, with no
     password (the owner's choice, 23 Nov 2026).
   - Saving needs a GitHub key with write access to the repository; it is kept on the owner's own
     devices only (localStorage) and sent nowhere but api.github.com.
   - The readings panel (#citiri): whoever has its password can send the five meter readings
     (citiri.json). The owner switches it on under Administrare, which stores a copy of the
     owner's key in panel.json, encrypted with the panel password (AES-GCM, PBKDF2-SHA256). */

var WEB = { owner: 'ferrarimix1', repo: 'fierari', branch: 'main', path: 'data.json', pending: 'citiri.json', panel: 'panel.json', iter: 310000, minPanelPass: 10, api: 'https://api.github.com' };
XLSX_URLS = ['vendor/xlsx.full.min.js'];

Object.assign(I18N.ro, {
  editorsOnly: 'Doar administratorul vede acest buton.',
  live: 'O lună nouă apare aici la 1–2 minute după ce e salvată.',
  secNote: 'Observații (opțional; le vede oricine deschide pagina)',
  savedT: '{m} a fost salvată. Pe site apare pentru toți în 1–2 minute.',
  deletedT: '{m} a fost ștearsă. Pe site se vede în 1–2 minute.',
  settingsT: 'Setările au fost salvate.',
  baseT: 'Citirile de pornire au fost salvate.',
  unavailable: 'Ca să salvezi, conectează acest dispozitiv la Administrare.',
  loading: 'Se încarcă…',
  loadFail: 'Datele nu au putut fi încărcate. Verifică internetul și reîncarcă pagina.',
  webAdmin: 'Administrare',
  webConflict: 'Între timp s-a salvat altceva. Am reîncărcat datele: verifică și salvează din nou.',
  webBadToken: 'Cheia GitHub nu mai e valabilă (a expirat sau a fost ștearsă). Pune una nouă la Administrare.',
  webNoAccess: 'Cheia GitHub nu are voie să scrie în depozitul „fierari”.',
  webOffline: 'Nu e conexiune la internet. Încearcă din nou.',
  adminTitle: 'Administrare',
  adminIntro: 'Ca să adaugi sau să modifici luni de pe acest site, conectează acest dispozitiv cu cheia GitHub a site-ului. Cheia rămâne doar pe acest dispozitiv.',
  adminToken: 'Cheia GitHub',
  adminConnect: 'Conectează',
  adminChecking: 'Se verifică…',
  adminBadToken: 'Cheia nu e valabilă. Verifică dacă ai copiat-o întreagă.',
  adminNoRepo: 'Cheia nu are acces la depozitul „fierari” sau la fișierul cu date.',
  adminWrongAccount: 'Cheia e din contul GitHub „{u}”, dar site-ul e în contul „{o}”. Fă cheia fiind conectat ca {o}.',
  webStale: 'Pagina s-a actualizat între timp. Reîncarcă pagina, apoi încearcă din nou.',
  webNewVersion: 'Există o versiune nouă a paginii. Reîncarc-o când termini.',
  adminConnected: 'Acest dispozitiv e conectat ca {u}. Poți adăuga și modifica luni.',
  adminDisconnect: 'Deconectează acest dispozitiv',
  webConnectedT: 'Conectat. Acum poți adăuga și modifica luni.',
  webDisconnectedT: 'Dispozitivul a fost deconectat.',
  webKey: 'Cheia GitHub',
  fEyebrow: 'De plată',
  fTotalsTitle: 'Totaluri pe categorii',
  fGazHow: '{b} × cota de suprafață a fiecărui apartament',
  fPeopleHow: '{b} ÷ {P} {w} = {v} de persoană',
  fPeopleHowApprox: '{b} ÷ {P} {w} ≈ {v} de persoană',
  fGunoiHow: '{v} de fiecare apartament × {n} = {t}',
  fElHow: '{K} kWh × {p} = {v}',
  fElPropHow: '{b} se împarte proporțional cu kWh-ii de plată ({K} kWh în total)',
  fColApt: 'Apartament',
  fColShare: 'Cota',
  fColPeople: 'Persoane',
  fColLei: 'Lei',
  fColOwn: 'Contor propriu (kWh)',
  fColCommon: 'Din comun (kWh)',
  fColPay: 'De plată (kWh)',
  fReadings: 'Citirile contoarelor',
  fRound: 'Sumele sunt rotunjite la ban, astfel încât se adună exact: pe fiecare factură și pe fiecare apartament.',
  fEmptyTitle: 'Încă nu sunt facturi de plată',
  fEmptyText: 'Prima lună de plată, {m}, apare aici după ce se adaugă facturile ei.',
  fBack: 'Facturile de plată',
  fAdminNote: 'Administrare: aici se adaugă și se modifică lunile.',
  citLink: 'Trimite citirile',
  citTitle: 'Citirile contoarelor',
  citIntro: 'Scrie indexul de pe fiecare contor. Citirile apar pe pagina principală în 1–2 minute și intră în calculul pentru {m}.',
  citOff: 'Panoul de citiri nu e pornit încă. Îl pornește proprietarul, din Administrare, alegând o parolă.',
  citPass: 'Parola',
  citOpen: 'Deschide',
  citChecking: 'Se verifică…',
  citWrong: 'Parola nu e corectă.',
  citRemember: 'Ține minte pe acest dispozitiv',
  citSave: 'Trimite citirile',
  citSaving: 'Se trimite…',
  citDone: 'Citirile au fost trimise. Apar pe pagina principală în 1–2 minute.',
  citLower: 'E mai mică decât ultima citire ({v}). Verifică; dacă s-a schimbat contorul, anunță proprietarul.',
  citExpired: 'Panoul nu mai poate salva: cheia lui nu mai e valabilă. Anunță proprietarul să activeze panoul din nou.',
  citFix: 'Corectează citirile marcate.',
  citLock: 'Închide panoul pe acest dispozitiv',
  citSent: 'Ultimele citiri au fost trimise pe {d}. Le poți corecta trimițându-le din nou.',
  pendTitle: 'Citiri noi pentru {m}',
  pendNote: 'Trimise pe {d}. Facturile lunii nu sunt adăugate încă; consumul e față de ultima citire.',
  pendLast: 'Ultima citire',
  pendNew: 'Citirea nouă',
  panTitle: 'Parola',
  panOffText: 'Cu o parolă, Administrarea și pagina „Trimite citirile” se deschid pe orice dispozitiv doar cu ea, fără cheie GitHub.',
  panOnText: 'Parola deschide Administrarea și pagina „Trimite citirile”. Aici o poți schimba sau opri.',
  panRisk: 'Parola deblochează o copie a cheii tale GitHub, păstrată pe site. Cine o are poate schimba orice la care ajunge cheia, așa că dă parola doar cui ai încredere. Dacă îți faci o cheie nouă, folosește „Am o cheie GitHub nouă”.',
  panPass: 'Parola',
  panPass2: 'Repetă parola',
  panOn: 'Salvează parola',
  panChange: 'Schimbă parola',
  panOffBtn: 'Oprește parola',
  panShort: 'Cel puțin {n} caractere.',
  panMismatch: 'Parolele nu coincid.',
  panOnT: 'Parola a fost salvată. Pe celelalte dispozitive merge în 1–2 minute.',
  panOffT: 'Parola a fost oprită: Administrarea merge doar cu cheia GitHub, iar panoul de citiri e închis.',
  panChecking: 'Se verifică…',
  webUnreachable: 'Nu se poate ajunge la GitHub (api.github.com): {d}. Poate o extensie a browserului (de ex. un blocker de reclame) sau rețeaua îl blochează. Încearcă din alt browser.',
  adminPassIntro: 'Scrie parola ca să adaugi și să modifici luni.',
  adminPassLbl: 'Parola',
  adminOpen: 'Deschide',
  adminWrongPass: 'Parola nu e corectă.',
  adminNewKey: 'Am o cheie GitHub nouă',
  adminUsePass: 'Am deja parola',
  adminSetupIntro: 'Prima dată: lipește cheia GitHub și alege o parolă. De acum înainte, pe orice dispozitiv, parola e tot ce trebuie: pentru Administrare și pentru trimiterea citirilor.',
  adminSetupBtn: 'Conectează și salvează parola',
  adminPassSavedT: 'Gata. De acum, Administrarea și citirile se deschid doar cu parola.',
  adminKeyExpired: 'Cheia păstrată cu parola nu mai e valabilă (a expirat sau a fost ștearsă). Apasă „Am o cheie GitHub nouă”.'
});
Object.assign(I18N.en, {
  editorsOnly: 'Only the admin sees this button.',
  live: 'A new month appears here 1–2 minutes after it is saved.',
  secNote: 'Notes (optional; anyone who opens the page sees them)',
  savedT: '{m} saved. Everyone sees it on the site within 1–2 minutes.',
  deletedT: '{m} deleted. The site shows it within 1–2 minutes.',
  settingsT: 'Settings saved.',
  baseT: 'Starting readings saved.',
  unavailable: 'To save, connect this device under Admin.',
  loading: 'Loading…',
  loadFail: 'The data could not be loaded. Check the internet connection and reload the page.',
  webAdmin: 'Admin',
  webConflict: 'Something else was saved meanwhile. The data has been reloaded: check it and save again.',
  webBadToken: 'The GitHub key no longer works (it expired or was deleted). Add a new one under Admin.',
  webNoAccess: 'The GitHub key may not write to the "fierari" repository.',
  webOffline: 'No internet connection. Try again.',
  adminTitle: 'Admin',
  adminIntro: "To add or change months from this site, connect this device with the site's GitHub key. The key stays on this device only.",
  adminToken: 'GitHub key',
  adminConnect: 'Connect',
  adminChecking: 'Checking…',
  adminBadToken: 'That key does not work. Check that it was copied whole.',
  adminNoRepo: 'The key has no access to the "fierari" repository or its data file.',
  adminWrongAccount: 'The key belongs to the GitHub account "{u}", but the site is in the account "{o}". Make the key while signed in as {o}.',
  webStale: 'The page was updated in the meantime. Reload it, then try again.',
  webNewVersion: 'A new version of the page is available. Reload it when you are done.',
  adminConnected: 'This device is connected as {u}. You can add and change months.',
  adminDisconnect: 'Disconnect this device',
  webConnectedT: 'Connected. You can now add and change months.',
  webDisconnectedT: 'This device was disconnected.',
  webKey: 'GitHub key',
  fEyebrow: 'To pay',
  fTotalsTitle: 'Totals by category',
  fGazHow: "{b} × each apartment's floor-area share",
  fPeopleHow: '{b} ÷ {P} {w} = {v} per person',
  fPeopleHowApprox: '{b} ÷ {P} {w} ≈ {v} per person',
  fGunoiHow: '{v} per apartment × {n} = {t}',
  fElHow: '{K} kWh × {p} = {v}',
  fElPropHow: '{b} is shared out in proportion to the kWh to pay ({K} kWh in all)',
  fColApt: 'Apartment',
  fColShare: 'Share',
  fColPeople: 'People',
  fColLei: 'Lei',
  fColOwn: 'Own meter (kWh)',
  fColCommon: 'Of common (kWh)',
  fColPay: 'To pay (kWh)',
  fReadings: 'Meter readings',
  fRound: 'Amounts are rounded to the ban so that everything adds up exactly: for each bill and for each apartment.',
  fEmptyTitle: 'No bills to pay yet',
  fEmptyText: 'The first month to pay, {m}, shows here once its bills are added.',
  fBack: 'Bills to pay',
  fAdminNote: 'Admin: months are added and changed here.',
  citLink: 'Send readings',
  citTitle: 'Meter readings',
  citIntro: 'Type the reading on each meter. The readings show on the front page within 1–2 minutes and go into the calculation for {m}.',
  citOff: 'The readings panel is not on yet. The owner turns it on under Admin by choosing a password.',
  citPass: 'Password',
  citOpen: 'Open',
  citChecking: 'Checking…',
  citWrong: 'That password is not right.',
  citRemember: 'Remember on this device',
  citSave: 'Send readings',
  citSaving: 'Sending…',
  citDone: 'Readings sent. They show on the front page within 1–2 minutes.',
  citLower: 'Lower than the last reading ({v}). Check it; if the meter was replaced, tell the owner.',
  citExpired: 'The panel can no longer save: its key no longer works. Ask the owner to switch the panel on again.',
  citFix: 'Correct the marked readings.',
  citLock: 'Close the panel on this device',
  citSent: 'The last readings were sent on {d}. You can correct them by sending them again.',
  pendTitle: 'New readings for {m}',
  pendNote: "Sent on {d}. The month's bills are not added yet; the use is since the last reading.",
  pendLast: 'Last reading',
  pendNew: 'New reading',
  panTitle: 'Password',
  panOffText: 'With a password, Admin and the "Send readings" page open on any device with just that, without a GitHub key.',
  panOnText: 'The password opens Admin and the "Send readings" page. Here you can change it or switch it off.',
  panRisk: 'The password unlocks a copy of your GitHub key kept on the site. Whoever has it can change anything that key reaches, so give the password only to people you trust. If you make a new key, use "I have a new GitHub key".',
  panPass: 'Password',
  panPass2: 'Password again',
  panOn: 'Save the password',
  panChange: 'Change password',
  panOffBtn: 'Switch the password off',
  panShort: 'At least {n} characters.',
  panMismatch: 'The passwords do not match.',
  panOnT: 'Password saved. Other devices can use it within 1–2 minutes.',
  panOffT: 'The password is off: Admin works only with the GitHub key, and the readings panel is closed.',
  panChecking: 'Checking…',
  webUnreachable: 'GitHub (api.github.com) cannot be reached: {d}. A browser extension (an ad blocker, say) or the network may be blocking it. Try another browser.',
  adminPassIntro: 'Type the password to add and change months.',
  adminPassLbl: 'Password',
  adminOpen: 'Open',
  adminWrongPass: 'That password is not right.',
  adminNewKey: 'I have a new GitHub key',
  adminUsePass: 'I already have the password',
  adminSetupIntro: 'First time: paste the GitHub key and choose a password. From then on, on any device, the password is all you need: for Admin and for sending the readings.',
  adminSetupBtn: 'Connect and save the password',
  adminPassSavedT: 'Done. From now on, Admin and the readings open with just the password.',
  adminKeyExpired: 'The key kept with the password no longer works (it expired or was deleted). Press "I have a new GitHub key".'
});
ICONS.key = '<circle cx="7.5" cy="15.5" r="3.5"/><path d="m10 13 8-8"/><path d="m15 8 3 3"/><path d="m17 6 2 2"/>';

/* ------------------------------------------------------------------ text ⇄ base64 (the GitHub API sends files as base64) */

function b64enc(bytes) {
  var s = '';
  for (var k = 0; k < bytes.length; k++) s += String.fromCharCode(bytes[k]);
  return btoa(s);
}
function b64dec(str) {
  var s = atob(String(str).replace(/\s+/g, '')), out = new Uint8Array(s.length);
  for (var k = 0; k < s.length; k++) out[k] = s.charCodeAt(k);
  return out;
}
function utf8ToB64(s) { return b64enc(new TextEncoder().encode(s)); }
function b64ToUtf8(s) { return new TextDecoder().decode(b64dec(s)); }

/* ------------------------------------------------------------------ encryption (the readings panel's copy of the key) */

function deriveKey(pass, salt, iter) {
  return crypto.subtle.importKey('raw', new TextEncoder().encode(pass), 'PBKDF2', false, ['deriveKey']).then(function (base) {
    return crypto.subtle.deriveKey({ name: 'PBKDF2', salt: salt, iterations: iter, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
  });
}
/* {v, kdf, iter, salt, iv, ct}: ct is the object as JSON, encrypted with a key from the password. */
function encryptJSON(obj, pass) {
  var salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12));
  return deriveKey(pass, salt, WEB.iter).then(function (key) {
    return crypto.subtle.encrypt({ name: 'AES-GCM', iv: iv }, key, new TextEncoder().encode(JSON.stringify(obj)));
  }).then(function (ct) {
    return { v: 1, kdf: 'PBKDF2-SHA256', iter: WEB.iter, salt: b64enc(salt), iv: b64enc(iv), ct: b64enc(new Uint8Array(ct)) };
  });
}
/* Rejects when the password is wrong (AES-GCM checks the whole message). */
function decryptJSON(env, pass) {
  if (!env || env.v !== 1 || !env.salt || !env.iv || !env.ct) return Promise.reject(new Error('format'));
  return deriveKey(pass, b64dec(env.salt), env.iter || WEB.iter).then(function (key) {
    return crypto.subtle.decrypt({ name: 'AES-GCM', iv: b64dec(env.iv) }, key, b64dec(env.ct));
  }).then(function (buf) { return JSON.parse(new TextDecoder().decode(buf)); });
}

/* ------------------------------------------------------------------ GitHub */

/* Plain requests only: some browsers turn fetch's cache option into extra headers that GitHub's
   CORS rules refuse, so freshness comes from a query string instead. */
function gh(method, url, token, body) {
  var headers = { 'Accept': 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28', 'Authorization': 'Bearer ' + token };
  if (body) headers['Content-Type'] = 'application/json';
  if (method === 'GET') url += (url.indexOf('?') === -1 ? '?' : '&') + 't=' + Date.now();
  var req;
  try {
    req = fetch(WEB.api + url, { method: method, headers: headers, body: body ? JSON.stringify(body) : undefined });
  } catch (err) {
    return Promise.reject({ code: 'unreachable', detail: String(err && err.message || err) });
  }
  return req.then(function (r) {
    if (r.ok) return r.json();
    var code = r.status === 401 ? 'badToken' : r.status === 403 || r.status === 404 ? 'noAccess' : r.status === 409 || r.status === 422 ? 'conflict' : 'failed';
    throw { code: code, status: r.status };
  }, function (err) {
    throw { code: typeof navigator !== 'undefined' && navigator.onLine === false ? 'offline' : 'unreachable', detail: String(err && err.message || err) };
  });
}

/* A key as pasted: whatever surrounds it (spaces, "Bearer", invisible characters) is dropped. */
function cleanToken(s) {
  var m = /(github_pat_[A-Za-z0-9_]+|gh[pousr]_[A-Za-z0-9]+)/.exec(String(s || ''));
  return m ? m[1] : String(s || '').replace(/[^A-Za-z0-9_]/g, '');
}
function ghFileUrl(owner) { return '/repos/' + owner + '/' + WEB.repo + '/contents/' + WEB.path; }

/* The owner's copy comes from the GitHub API: always the latest, with the version (sha) a save needs. */
function loadFromGitHub(token, owner) {
  return gh('GET', ghFileUrl(owner) + '?ref=' + WEB.branch, token).then(function (f) {
    return { sha: f.sha, data: normalizeData(JSON.parse(b64ToUtf8(f.content))) };
  });
}

/* One commit of data.json, written one month per line so GitHub shows each change plainly. */
function putData(nd, message) {
  var g = state.gh;
  var body = { message: message, content: utf8ToB64(serializeData(nd) + '\n'), branch: WEB.branch };
  if (g.sha) body.sha = g.sha;
  return gh('PUT', ghFileUrl(g.owner), g.token, body).then(function (r) {
    g.sha = r && r.content && r.content.sha;
  });
}

/* Other files in the repository (the readings, the panel's key), with key = {token, owner}. */
function ghPath(key, path) { return '/repos/' + key.owner + '/' + WEB.repo + '/contents/' + path; }
function ghGetFile(key, path) {
  return gh('GET', ghPath(key, path) + '?ref=' + WEB.branch, key.token).catch(function (e) {
    if (e && e.status === 404) return null;
    throw e;
  });
}
/* creates or replaces a file; tries once more if someone else wrote it in between */
function ghWriteFile(key, path, text, message) {
  function attempt(n) {
    return ghGetFile(key, path).then(function (f) {
      var body = { message: message, content: utf8ToB64(text), branch: WEB.branch };
      if (f && f.sha) body.sha = f.sha;
      return gh('PUT', ghPath(key, path), key.token, body);
    }).catch(function (e) {
      if (e && e.code === 'conflict' && n < 1) return attempt(n + 1);
      throw e;
    });
  }
  return attempt(0);
}
function ghFileJSON(f) {
  try { return f && f.content ? JSON.parse(b64ToUtf8(f.content)) : null; } catch (e) { return null; }
}
function ghErrorText(e) {
  var code = e && e.code;
  if (code === 'unreachable') return tr('webUnreachable', { d: e.detail || '?' });
  return tr(code === 'stale' ? 'webStale' : code === 'conflict' ? 'webConflict' : code === 'badToken' ? 'webBadToken' : code === 'noAccess' ? 'webNoAccess' : code === 'offline' ? 'webOffline' : 'saveFailed') +
    (code === 'failed' && e.status ? ' (GitHub: ' + e.status + ')' : '');
}

/* ------------------------------------------------------------------ boot */

function boot() {
  root = $('app');
  state.lang = pickLang();
  state.apt = pickApt();
  var r = lsGet('ff-range');
  state.range = r === '12' ? 12 : r === '0' ? 0 : 24;
  state.downloads = { save: saveFile };
  document.documentElement.lang = state.lang;
  var toast = document.createElement('div');
  toast.className = 'toast';
  toast.id = 'toast';
  toast.setAttribute('role', 'status');
  toast.setAttribute('aria-live', 'polite');
  document.body.appendChild(toast);
  /* the page used to ask for a password; a device that remembered one no longer needs it */
  lsDel('ff-pass');
  ssDel('ff-pass');
  bindEvents();
  bindWebEvents();
  loadSite();
}

function bandHTML() {
  return '<header class="band"><div class="band-inner"><div class="brand"><h1>Facturi Fierari</h1><p>' + esc(tr('tagline')) + '</p></div>' +
    '<div class="lang" role="group" aria-label="' + esc(tr('lang')) + '">' + ['ro', 'en'].map(function (l) {
      return '<button type="button" data-action="web-lang" data-v="' + l + '" aria-pressed="' + (state.lang === l) + '" lang="' + l + '">' + l.toUpperCase() + '</button>';
    }).join('') + '</div></div></header>';
}

/* The published data, fresh: the query string gets past any cached copy. */
function fetchData() {
  return fetch(WEB.path + '?t=' + Date.now(), { cache: 'no-store' }).then(function (r) {
    if (!r.ok) throw new Error('http ' + r.status);
    return r.json();
  }).then(function (j) {
    var d = normalizeData(j);
    if (!d.months.length) throw new Error('empty');
    return d;
  });
}

/* a file that may not exist (the readings, the panel's key): null when it is missing */
function fetchOptional(path) {
  return fetch(path + '?t=' + Date.now(), { cache: 'no-store' }).then(function (r) {
    return r.ok ? r.json() : null;
  }).catch(function () { return null; });
}

function loadSite() {
  root.innerHTML = bandHTML() + '<div class="wrap"><p class="web-status">' + esc(tr('loading')) + '</p></div>';
  Promise.all([fetchData(), fetchOptional(WEB.pending)]).then(function (res) {
    state.pendingRaw = res[1];
    start(res[0]);
  }, function () {
    root.innerHTML = bandHTML() + '<div class="wrap"><p class="web-status bad">' + esc(tr('loadFail')) + '</p></div>';
  });
}

function start(d) {
  state.data = d;
  state.comp = computeAll(d);
  state.sel = d.months.length - 1;
  state.canEdit = false;
  state.view = viewFromHash();
  state.fsel = null;
  renderAll();
  connectOwner(lsGet('ff-token'), false);
  if (state.view === 'admin' && !lsGet('ff-token')) openAdmin();
  if (state.view === 'panel') loadPanelInfo();
}

/* ------------------------------------------------------------------ owner */

/* Checks a GitHub key: that it belongs to the site's account and can read the data file in the
   "fierari" repository. On success the owner's copy of the data replaces the one from the site. */
function connectOwner(token, fromDialog) {
  if (!token) return Promise.resolve(false);
  return gh('GET', '/user', token).then(function (u) {
    var login = u && u.login ? String(u.login) : '';
    if (login.toLowerCase() !== WEB.owner.toLowerCase()) throw { code: 'wrongAccount', login: login };
    return loadFromGitHub(token, WEB.owner).then(function (f) {
      var d = f.data;
      state.gh = { token: token, owner: WEB.owner, login: login, sha: f.sha };
      lsSet('ff-token', token);
      if (JSON.stringify(d) !== JSON.stringify(state.data)) {
        state.data = d;
        state.comp = computeAll(d);
        state.sel = d.months.length - 1;
      }
      state.webCanEdit = true;
      state.canEdit = false;
      updateEditability();
      return true;
    });
  }).catch(function (e) {
    state.webCanEdit = false;
    state.gh = null;
    if (!fromDialog) {
      if (e && e.code === 'badToken') { lsDel('ff-token'); showToast(tr('webBadToken')); }
      return false;
    }
    throw e;
  });
}

function updateEditability() {
  var can = !!state.webCanEdit;
  if (can === state.canEdit) { renderFooter(); return; }
  state.canEdit = can;
  if (!can && state.sheet) state.sheet = null;
  renderAll();
}

function commitMessage(after) {
  var lang = state.lang;
  state.lang = 'ro';
  var m = after ? afterMessage(after) : tr('settingsT');
  state.lang = lang;
  return m.split('.')[0];
}

function saveData(nd, after) {
  if (!state.webCanEdit || !state.gh) { setSheetMsg(tr('unavailable'), 'bad'); return Promise.resolve(); }
  nd.updated = todayISO();
  setSaving(true);
  setSheetMsg('', '');
  /* never write with outdated code: the typed form is kept in this browser (ff-draft) for after the reload */
  return isCurrent().then(function (cur) {
    if (!cur) throw { code: 'stale' };
    return putData(nd, commitMessage(after));
  }).then(function () {
    lsDel('ff-draft');
    applyInPlace(nd, after);
  }, function (e) {
    setSaving(false);
    var code = e && e.code;
    if (code === 'conflict') {
      setSheetMsg(tr('webConflict'), 'warn');
      refreshFromGitHub();
    } else if (code === 'stale') setSheetMsg(tr('webStale'), 'warn');
    else setSheetMsg(ghErrorText(e), 'bad');
  });
}

/* After a conflict: take the latest from GitHub, keep the open form as typed. */
function refreshFromGitHub() {
  var g = state.gh;
  if (!g) return Promise.resolve();
  return loadFromGitHub(g.token, g.owner).then(function (f) {
    var d = f.data;
    g.sha = f.sha;
    state.data = d;
    state.comp = computeAll(d);
    if (state.sel > d.months.length - 1) state.sel = d.months.length - 1;
    if (state.sheet && state.sheet.mode === 'add') state.sheet.index = d.months.length;
    renderAll();
    setSheetMsg(tr('webConflict'), 'warn');
  }).catch(function () { /* the message about the conflict stays */ });
}

/* Excel and other files: an ordinary download. */
function saveFile(o) {
  var blob = new Blob([o.data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  var url = URL.createObjectURL(blob), a = document.createElement('a');
  a.href = url;
  a.download = o.filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 1500);
  return Promise.resolve({ status: 'saved' });
}

function renderFooter() {
  var el = $('foot');
  if (!el || !state.data) return;
  var d = state.data;
  var info = d.updated ? tr('updated', { d: new Intl.DateTimeFormat(I18N[state.lang].locale, { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' }).format(new Date(d.updated + 'T00:00:00Z')) }) : '';
  var h = '<div><p>' + esc(info) + '</p><p>' + esc(tr('live')) + '</p></div><div class="foot-actions">';
  if (state.view === 'admin') {
    h += '<button type="button" class="btn sm" data-action="export"' + (state.exporting ? ' disabled' : '') + '>' + icon('down') + esc(state.exporting ? tr('exporting') : tr('exportXlsx')) + '</button>';
    if (state.canEdit) h += '<button type="button" class="btn sm" data-action="settings">' + icon('sliders') + esc(tr('settings')) + '</button>';
    h += '<button type="button" class="btn sm" data-action="web-admin">' + icon('key') + esc(tr('webKey')) + '</button>';
  } else {
    /* the front page only shows; readings go through the panel, everything else under Administrare */
    if (state.view !== 'panel') h += '<a class="btn sm ghost" href="#citiri" data-route="panel">' + icon('edit') + esc(tr('citLink')) + '</a>';
    h += '<a class="btn sm ghost" href="#admin" data-route="admin">' + icon('key') + esc(tr('webAdmin')) + '</a>';
  }
  el.innerHTML = h + '</div>';
}

/* ------------------------------------------------------------------ the front page

   What everyone sees first: the latest month's bills, read-only. Each apartment's amount with its
   part of every bill, the total of each bill, and how each one was split, with the month's own
   numbers. The full app, where the owner adds and changes months, is at #admin. */

var VIEW_HASH = { admin: '#admin', panel: '#citiri' };
function viewFromHash() { return location.hash === '#admin' ? 'admin' : location.hash === '#citiri' ? 'panel' : 'front'; }

function setView(v, push) {
  if (v === state.view) return;
  state.view = v;
  if (push) {
    try { history.pushState(null, '', VIEW_HASH[v] || location.pathname + location.search); } catch (e) { location.hash = VIEW_HASH[v] || ''; }
  }
  if (v === 'panel') loadPanelInfo();
  if (v !== 'admin') {
    state.sheet = null;
    state.fsel = null;
    document.body.classList.remove('locked');
    var dlg = $('admin-dlg');
    if (dlg && dlg.open) dlg.close();
  }
  if (!state.data) return;
  renderAll();
  window.scrollTo(0, 0);
  if (v === 'admin' && !state.gh && !lsGet('ff-token')) openAdmin();
}

/* The apartment chips of the full app keep the #admin address (the app's own version writes #ap-N). */
function setApt(a) {
  state.apt = a;
  lsSet('ff-apt', a == null ? '' : String(a));
  var y = window.scrollY;
  renderAll();
  window.scrollTo(0, y);
}

/* the month shown on the front page: the latest with bills, unless another one was picked */
function frontIndex(paid) {
  if (!paid.length) return -1;
  return state.fsel != null && paid.indexOf(state.fsel) !== -1 ? state.fsel : paid[paid.length - 1];
}

function billLabel(b) { return tr(b === 'curent' ? 'curentP' : b); }

/* A month's amounts per bill and apartment, in bani. Each bill's four parts add up to the bill,
   and each apartment's parts add up to what it pays (the same totals as everywhere else): every
   part is its exact amount rounded down or up to the ban. */
function centsTable(i) {
  var c = state.comp[i], m = state.data.months[i];
  var out = { rows: [], cell: {}, rowTot: {}, total: sum(c.cents), ok: true };
  if (c.start || c.manual || !c.cents.every(num)) return out;
  out.rows = BILLS.filter(function (b) { return b !== 'gunoi' || num(m.gunoi); });
  var R = out.rows.map(function (b) {
    return b === 'curent' ? 0 : b === 'gunoi' ? Math.round(m.gunoi * 100) * APTS.length : Math.round(m[b] * 100);
  });
  /* electricity takes the rest, so the bills add up to the total to pay to the ban */
  R[out.rows.indexOf('curent')] = out.total - sum(R);
  var real = out.rows.map(function (b) { return c.parts[b].map(function (v) { return v * 100; }); });
  var X = roundTable(real, R, c.cents);
  if (!X) { X = roundLoose(real, c.cents); out.ok = false; }
  out.rows.forEach(function (b, r) { out.cell[b] = X[r]; out.rowTot[b] = R[r]; });
  return out;
}

/* Rounds a table of amounts to whole numbers with the given row and column totals: each cell goes
   down or up by less than 1, only cells that are not already whole move, and the result stays as
   close as it can to the exact amounts (a cell at .67 goes up before one at .33). It is a tiny
   minimum-cost flow: source → rows (the "ups" each row needs) → cells → columns → sink, where
   rounding a cell up instead of down costs 1 − 2 × its fraction. Null when it cannot be done. */
function roundTable(real, R, C) {
  var nr = real.length, nc = C.length, EPS = 1e-6;
  var X = real.map(function (row) { return row.map(function (v) { return Math.floor(v + EPS); }); });
  var fr = real.map(function (row, r) { return row.map(function (v, a) { return v - X[r][a]; }); });
  var needR = R.map(function (t, r) { return t - sum(X[r]); });
  var needC = C.map(function (t, a) { return t - sum(X.map(function (row) { return row[a]; })); });
  if (needR.some(function (v) { return v < 0; }) || needC.some(function (v) { return v < 0; }) || sum(needR) !== sum(needC)) return null;
  var S = nr + nc, T = S + 1, g = [], cells = [];
  for (var n = 0; n <= T; n++) g.push([]);
  function edge(u, v, cap, cost) {
    g[u].push({ v: v, cap: cap, cost: cost, rev: g[v].length });
    g[v].push({ v: u, cap: 0, cost: -cost, rev: g[u].length - 1 });
    return g[u][g[u].length - 1];
  }
  needR.forEach(function (k, r) { if (k > 0) edge(S, r, k, 0); });
  needC.forEach(function (k, a) { if (k > 0) edge(nr + a, T, k, 0); });
  fr.forEach(function (row, r) {
    row.forEach(function (f, a) { if (f > EPS) cells.push({ r: r, a: a, e: edge(r, nr + a, 1, 1 - 2 * f) }); });
  });
  for (var flow = 0, need = sum(needR); flow < need; flow++) {
    /* the cheapest way to hand out one more "up" (Bellman–Ford: the costs can be negative) */
    var dist = g.map(function () { return Infinity; }), prev = [];
    dist[S] = 0;
    for (var it = 0; it < g.length; it++) {
      var changed = false;
      for (var u = 0; u < g.length; u++) {
        if (dist[u] === Infinity) continue;
        for (var j = 0; j < g[u].length; j++) {
          var e = g[u][j];
          if (e.cap > 0 && dist[u] + e.cost < dist[e.v] - 1e-9) { dist[e.v] = dist[u] + e.cost; prev[e.v] = [u, j]; changed = true; }
        }
      }
      if (!changed) break;
    }
    if (dist[T] === Infinity) return null;
    for (var v = T; v !== S; v = prev[v][0]) {
      var pe = g[prev[v][0]][prev[v][1]];
      pe.cap -= 1;
      g[v][pe.rev].cap += 1;
    }
  }
  cells.forEach(function (c) { if (c.e.cap === 0) X[c.r][c.a] += 1; });
  return X;
}

/* fallback: round each cell, then make each apartment's column add up to what it pays */
function roundLoose(real, C) {
  var X = real.map(function (row) { return row.map(function (v) { return Math.round(v); }); });
  C.forEach(function (t, a) {
    var d = t - sum(X.map(function (row) { return row[a]; })), s = d > 0 ? 1 : -1;
    var rows = X.map(function (_, r) { return r; }).sort(function (p, q) { return s * ((real[q][a] - X[q][a]) - (real[p][a] - X[p][a])); });
    for (var n = 0; d !== 0 && n < rows.length * 4; n++) { X[rows[n % rows.length]][a] += s; d -= s; }
  });
  return X;
}

function renderFront() {
  var paid = paidMonths(), i = frontIndex(paid);
  var h = bandHTML() + '<main class="wrap front">' + (i < 0 ? frontEmptyHTML() : frontMonthHTML(i, paid)) +
    '<footer class="foot" id="foot"></footer></main>';
  root.innerHTML = h;
  renderFooter();
  if (i < 0) renderRules();
}

/* nothing to pay yet: when the first month will appear, the starting readings, and the rules */
function frontEmptyHTML() {
  var n = state.data.months.length, last = state.comp[n - 1];
  var h = '<section class="panel stmt" aria-labelledby="f-title"><p class="eyebrow">' + esc(tr('fEyebrow')) + '</p>' +
    '<h2 class="mtitle f-empty-title" id="f-title">' + esc(tr('fEmptyTitle')) + '</h2>' +
    '<p class="stmt-lead">' + esc(tr('fEmptyText', { m: mName(n) })) + '</p>';
  if (last && last.start) h += metersHTML(n - 1);
  return h + '</section>' + pendingHTML() + '<section class="panel" id="rules-panel" aria-labelledby="rules-title"></section>';
}

/* ------------------------------------------------------------------ readings sent through the panel

   citiri.json = {t: "YYYY-MM-DD", r: [five readings]}. They count as new readings for the next month
   until a month uses them: none may be below the latest month's readings and one must be above. */
function pendingReadings() {
  var p = state.pendingRaw, ms = state.data && state.data.months;
  if (!p || !Array.isArray(p.r) || p.r.length !== METERS.length || !ms || !ms.length) return null;
  var last = ms[ms.length - 1].r, r = p.r.map(Number);
  if (!r.every(num) || r.some(function (v, j) { return v < last[j]; }) || r.every(function (v, j) { return v === last[j]; })) return null;
  return { r: r, t: typeof p.t === 'string' ? p.t : '', last: last };
}

function fmtDay(iso) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso || '')) return '—';
  return new Intl.DateTimeFormat(I18N[state.lang].locale, { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' }).format(new Date(iso + 'T00:00:00Z'));
}

/* on the front page: the new readings and the use since the last ones, before the bills come */
function pendingHTML() {
  var p = pendingReadings();
  if (!p) return '';
  var h = '<section class="panel" id="f-pending" aria-labelledby="f-pend-title"><h2 class="h2" id="f-pend-title">' + esc(tr('pendTitle', { m: mName(state.data.months.length) })) + '</h2>' +
    '<p class="hint">' + esc(tr('pendNote', { d: fmtDay(p.t) })) + '</p><div class="scroll"><table class="tbl meters"><thead><tr><th scope="col">' + esc(tr('meter')) + '</th>' +
    '<th scope="col">' + esc(tr('pendLast')) + '</th><th scope="col">' + esc(tr('pendNew')) + '</th><th scope="col" class="num">' + esc(tr('usage')) + '</th></tr></thead><tbody>';
  METERS.forEach(function (j) {
    var key = j < 4 ? '<span class="key a' + j + '" aria-hidden="true"></span> ' : '';
    h += '<tr><th scope="row">' + key + esc(meterName(j)) + '</th><td><span class="drum">' + esc(drum(p.last[j])) + '</span></td><td><span class="drum">' + esc(drum(p.r[j])) + '</span></td>' +
      '<td class="num">' + esc(kwhStr(p.r[j] - p.last[j])) + ' kWh</td></tr>';
  });
  return h + '</tbody></table></div></section>';
}

/* the owner's form for the next month starts from the readings sent through the panel */
var coreDraftForAdd = draftForAdd;
draftForAdd = function () {
  var d = coreDraftForAdd(), p = pendingReadings();
  if (p) d.r = p.r.map(numToInput);
  return d;
};

function frontNavBtn(act, enabled, label, ic) {
  return '<button type="button" class="nav-btn" data-action="' + act + '" aria-label="' + esc(tr(label)) + '"' + (enabled ? '' : ' disabled') + '>' + icon(ic) + '</button>';
}

function frontMonthHTML(i, paid) {
  var c = state.comp[i], m = state.data.months[i], k = paid.indexOf(i), t = centsTable(i);
  var h = '<section class="panel stmt" aria-labelledby="f-title"><div class="stmt-head"><div class="stmt-titles">' +
    '<p class="eyebrow">' + esc(tr('fEyebrow')) + '</p><div class="mnav">';
  if (paid.length > 1) h += frontNavBtn('f-prev', k > 0, 'prevMonth', 'prev');
  h += '<h2 class="mtitle" id="f-title">' + esc(cap(mName(i))) + '</h2>';
  if (paid.length > 1) h += frontNavBtn('f-next', k < paid.length - 1, 'nextMonth', 'next');
  h += '</div></div><div class="stmt-total"><span class="eyebrow">' + esc(tr(c.manual ? 'totalPaid' : 'totalDue')) + '</span><b id="f-total">' + esc(lei(t.total / 100)) + '</b></div></div>';
  h += '<div class="cwrap"><div class="aptgrid">';
  APTS.forEach(function (a) {
    h += '<article class="aptcard a' + a + '" aria-labelledby="f-ap' + a + '"><h3 class="ac-name" id="f-ap' + a + '"><span class="key" aria-hidden="true"></span>' + esc(aptName(a)) + '</h3>' +
      '<p class="ac-amt">' + esc(money(c.cents[a] / 100)) + '<span class="cur">lei</span></p>';
    if (t.rows.length) {
      h += '<dl class="ac-lines">' + t.rows.map(function (b) {
        return '<div><dt>' + esc(billLabel(b)) + '</dt><dd>' + esc(money(t.cell[b][a] / 100)) + '</dd></div>';
      }).join('') + '</dl>';
    }
    h += '</article>';
  });
  h += '</div></div>';
  if (c.manual) h += '<p class="noteline">' + esc(tr('manualNote', { v: lei(c.billTotal) })) + '</p>';
  if (m.note) h += '<p class="note"><b>' + esc(tr('note')) + ':</b> ' + esc(m.note) + '</p>';
  h += '</section>' + pendingHTML();
  if (t.rows.length) h += frontTotalsHTML(i, t) + frontHowHTML(i, t);
  return h;
}

/* each bill's total and how it is split */
function frontTotalsHTML(i, t) {
  var c = state.comp[i], m = state.data.months[i];
  var h = '<section class="panel" aria-labelledby="f-tot-title"><h2 class="h2" id="f-tot-title">' + esc(tr('fTotalsTitle')) + '</h2><div class="bills f-bills">';
  t.rows.forEach(function (b) {
    var sub = b === 'curent' && c.byPrice ? tr('kwhTimes', { k: kwhTotal(c.K), p: priceStr(c.pk) }) :
      b === 'gunoi' ? tr('gunoiExpr', { n: APTS.length, v: lei(m.gunoi) }) : m[b + 'Text'] ? exprDisplay(m[b + 'Text']) : '';
    h += '<div class="bill"><span class="bname">' + esc(billLabel(b)) + '</span><span class="brule">' + esc(tr(ruleOf(c, b))) + '</span>' +
      '<span class="bval">' + esc(lei(t.rowTot[b] / 100)) + (sub ? '<span class="expr">' + esc(sub) + '</span>' : '') + '</span></div>';
  });
  h += '</div><div class="totalrow"><span class="eyebrow">' + esc(tr('totalDue')) + '</span><span class="tval">' + esc(lei(t.total / 100)) + '</span></div>';
  if (t.ok) h += '<p class="okline">' + icon('check') + esc(tr('sumOkDue')) + '</p>';
  return h + '</section>';
}

function frontHowHTML(i, t) {
  var h = '<section class="panel f-how" aria-labelledby="f-how-title"><h2 class="h2" id="f-how-title">' + esc(tr('howTitle')) + '</h2>';
  t.rows.forEach(function (b) { h += howBlockHTML(i, t, b); });
  return h + '<p class="hint f-round">' + esc(tr('fRound')) + '</p></section>';
}

function howBlockHTML(i, t, b) {
  var c = state.comp[i], m = state.data.months[i];
  var h = '<div class="howblock" id="f-how-' + b + '"><div class="hb-head"><h3 class="h3">' + esc(billLabel(b)) + '</h3><b>' + esc(lei(t.rowTot[b] / 100)) + '</b></div>';
  var formula = function (s) { return '<p class="hb-formula">' + esc(s) + '</p>'; };
  if (b === 'gaz') {
    h += formula(tr('fGazHow', { b: lei(m.gaz) }));
    h += aptTableHTML([tr('fColShare')], function (a) { return [pct(c.shares[a])]; }, [pct(1)], t.cell.gaz, t.rowTot.gaz);
  } else if (b === 'digi' || b === 'apa') {
    var per = m[b] / c.P, exact = Math.abs(per * 100 - Math.round(per * 100)) < 1e-6;
    h += formula(tr(exact ? 'fPeopleHow' : 'fPeopleHowApprox', { b: lei(m[b]), P: c.P, w: peopleWord(c.P), v: lei(per) }));
    h += aptTableHTML([tr('fColPeople')], function (a) { return [String(m.p[a])]; }, [String(c.P)], t.cell[b], t.rowTot[b]);
  } else if (b === 'gunoi') {
    h += formula(tr('fGunoiHow', { v: lei(m.gunoi), n: APTS.length, t: lei(t.rowTot.gunoi / 100) }));
  } else {
    h += formula(c.byPrice ? tr('fElHow', { K: kwhTotal(c.K), p: priceStr(c.pk), v: lei(t.rowTot.curent / 100) }) :
      tr('fElPropHow', { b: lei(m.curent), K: kwhTotal(c.K) }));
    if (c.byPrice) h += priceBoxHTML(i);
    h += readingsHTML(i);
    var com = APTS.map(function (a) { return c.shares[a] * c.kwh[4]; });
    h += aptTableHTML([tr('fColOwn'), tr('fColCommon'), tr('fColPay')], function (a) { return [kwhStr(c.kwh[a]), kwhStr(com[a]), kwhStr(c.kwhPay[a])]; },
      [kwhStr(sum(c.kwh.slice(0, 4))), kwhStr(c.kwh[4]), kwhStr(c.K)], t.cell.curent, t.rowTot.curent);
    h += '<p class="hint">' + esc(tr('commonNote')) + '</p>';
  }
  return h + '</div>';
}

/* one row per apartment: the basis of the split, then its amount in lei */
function aptTableHTML(cols, rowFn, tot, cents, totCents) {
  var h = '<div class="scroll"><table class="tbl hb-tbl"><thead><tr><th scope="col">' + esc(tr('fColApt')) + '</th>' +
    cols.map(function (s) { return '<th scope="col" class="num">' + esc(s) + '</th>'; }).join('') +
    '<th scope="col" class="num">' + esc(tr('fColLei')) + '</th></tr></thead><tbody>';
  APTS.forEach(function (a) {
    h += '<tr><th scope="row"><span class="key a' + a + '" aria-hidden="true"></span> ' + esc(aptName(a)) + '</th>' +
      rowFn(a).map(function (v) { return '<td class="num">' + esc(v) + '</td>'; }).join('') +
      '<td class="num">' + esc(money(cents[a] / 100)) + '</td></tr>';
  });
  return h + '</tbody><tfoot><tr><th scope="row">' + esc(tr('total')) + '</th>' + tot.map(function (v) { return '<td class="num">' + esc(v) + '</td>'; }).join('') +
    '<td class="num">' + esc(money(totCents / 100)) + '</td></tr></tfoot></table></div>';
}

/* the five meters: last month, this month, the use */
function readingsHTML(i) {
  var m = state.data.months[i], c = state.comp[i], prev = state.data.months[i - 1];
  var h = '<p class="hb-sub">' + esc(tr('fReadings')) + '</p><div class="scroll"><table class="tbl meters"><thead><tr><th scope="col">' + esc(tr('meter')) + '</th>' +
    '<th scope="col">' + esc(tr('lastMonth')) + '</th><th scope="col">' + esc(tr('thisMonth')) + '</th><th scope="col" class="num">' + esc(tr('usage')) + '</th></tr></thead><tbody>';
  METERS.forEach(function (j) {
    var key = j < 4 ? '<span class="key a' + j + '" aria-hidden="true"></span> ' : '';
    var replaced = m.kwh && m.kwh[j] != null;
    h += '<tr><th scope="row">' + key + esc(meterName(j)) + '</th><td><span class="drum">' + esc(drum(prev.r[j])) + '</span></td><td><span class="drum">' + esc(drum(m.r[j])) + '</span></td>' +
      '<td class="num">' + esc(kwhStr(c.kwh[j])) + ' kWh' + (replaced ? '<span class="tag">' + esc(tr('replacedNote')) + '</span>' : '') + '</td></tr>';
  });
  return h + '</tbody></table></div>';
}

/* the full app gets a way back to the front page */
function adminBarHTML() {
  return '<nav class="adminbar" aria-label="' + esc(tr('webAdmin')) + '"><a class="btn sm" href="' + esc(location.pathname + location.search) + '" data-route="front">' +
    icon('prev') + esc(tr('fBack')) + '</a><span class="muted">' + esc(tr('fAdminNote')) + '</span></nav>';
}

/* ------------------------------------------------------------------ the readings panel (#citiri)

   With the panel password, anyone can send the five meter readings for the next month. The
   password decrypts panel.json, the copy of the owner's key the owner stored when switching the
   panel on; with it the page commits citiri.json and nothing else. */

state.cit = { info: undefined, key: null, draft: ['', '', '', '', ''], touched: {}, passMsg: '', saving: false };

/* whether the panel is switched on (panel.json exists); a remembered password opens it at once */
function loadPanelInfo() {
  var c = state.cit;
  return fetchOptional(WEB.panel).then(function (env) {
    c.info = env && env.v === 1 && env.ct ? env : null;
    var saved = lsGet('ff-cit');
    if (c.info && saved && !c.key) return unlockPanel(saved, true, true);
  }).then(function () { if (state.view === 'panel') renderAll(); });
}

function unlockPanel(pass, remember, silent) {
  var c = state.cit;
  return decryptJSON(c.info, pass).then(function (k) {
    if (!k || !k.token || !k.owner) throw new Error('format');
    c.key = { token: k.token, owner: k.owner };
    if (remember) lsSet('ff-cit', pass);
    c.passMsg = '';
    var p = pendingReadings();
    c.draft = METERS.map(function (j) { return p ? numToInput(p.r[j]) : ''; });
    c.touched = {};
  }, function () {
    if (silent) { lsDel('ff-cit'); return; }
    c.passMsg = tr('citWrong');
  });
}

function renderPanel() {
  var c = state.cit, n = state.data.months.length;
  var h = bandHTML() + '<main class="wrap front"><nav class="adminbar"><a class="btn sm" href="' + esc(location.pathname + location.search) + '" data-route="front">' +
    icon('prev') + esc(tr('fBack')) + '</a></nav><section class="panel cit" aria-labelledby="c-title"><h2 class="h2" id="c-title">' + esc(tr('citTitle')) + '</h2>';
  if (c.info === undefined) h += '<p class="hint">' + esc(tr('panChecking')) + '</p>';
  else if (!c.info) h += '<p class="basenote cit-off">' + esc(tr('citOff')) + '</p>';
  else if (!c.key) {
    h += '<form id="cit-unlock" class="cit-unlock" novalidate><div class="field"><label for="c-pass">' + esc(tr('citPass')) + '</label>' +
      '<input type="password" id="c-pass" autocomplete="current-password" autocapitalize="none" spellcheck="false" enterkeyhint="go" aria-describedby="c-pass-m"' + (c.passMsg ? ' aria-invalid="true"' : '') + '></div>' +
      '<label class="check"><input type="checkbox" id="c-remember">' + esc(tr('citRemember')) + '</label>' +
      '<p class="msg bad" id="c-pass-m" role="alert">' + esc(c.passMsg) + '</p>' +
      '<div class="sheet-actions"><button type="submit" class="btn primary" id="c-open">' + esc(tr('citOpen')) + '</button></div></form>';
  } else {
    var p = pendingReadings();
    h += '<p class="stmt-lead">' + esc(tr('citIntro', { m: mName(n) })) + '</p>' + (p ? '<p class="hint">' + esc(tr('citSent', { d: fmtDay(p.t) })) + '</p>' : '') +
      '<form id="cit-form" novalidate><div class="fgrid five">';
    METERS.forEach(function (j) {
      var id = 'c-r' + j;
      h += '<div class="field"><label for="' + id + '">' + (j < 4 ? '<span class="key a' + j + '" aria-hidden="true"></span>' : '') + esc(meterName(j)) + '</label>' +
        '<input type="text" id="' + id + '" data-cit="' + j + '" inputmode="decimal" autocomplete="off" spellcheck="false" enterkeyhint="next" value="' + esc(c.draft[j]) + '" aria-describedby="' + id + '-i ' + id + '-m">' +
        '<small class="info" id="' + id + '-i"></small><small class="msg" id="' + id + '-m"></small></div>';
    });
    h += '</div><p class="sheet-msg" id="c-msg" role="status"></p><div class="sheet-actions"><button type="submit" class="btn primary" id="c-save">' + icon('check') + esc(tr('citSave')) + '</button>' +
      '<button type="button" class="btn ghost" data-action="cit-lock">' + esc(tr('citLock')) + '</button></div></form>';
  }
  root.innerHTML = h + '</section><footer class="foot" id="foot"></footer></main>';
  renderFooter();
  if (c.key) refreshCit(false);
  var f = $('c-pass');
  if (f && c.passMsg) f.focus();
}

/* the readings as typed: each must be a number, not below the last reading, and not absurdly far above */
function citCheck(draft) {
  var last = state.data.months[state.data.months.length - 1].r, out = { r: [], errors: {}, ok: true };
  METERS.forEach(function (j) {
    var p = parseExpr(draft[j], 'plain'), e = null;
    if (p.empty) e = 'eRequired';
    else if (p.error) e = p.error === 'plain' ? 'ePlain' : p.error === 'thousands' ? 'eThousands' : 'eNumber';
    else if (p.value < 0) e = 'eNegative';
    else if (p.value < last[j]) e = 'lower';
    else if (p.value - last[j] > LIMIT_KWH) e = 'eKwhBig';
    out.r[j] = p.value;
    if (e) { out.errors[j] = e; out.ok = false; }
  });
  return out;
}

function refreshCit(all) {
  var c = state.cit, last = state.data.months[state.data.months.length - 1].r, chk = citCheck(c.draft);
  METERS.forEach(function (j) {
    var inp = $('c-r' + j), m = $('c-r' + j + '-m'), info = $('c-r' + j + '-i');
    if (!inp) return;
    var e = chk.errors[j], show = !!e && (all || (c.touched[j] && e !== 'eRequired'));
    m.className = 'msg' + (show ? ' bad' : '');
    m.textContent = show ? (e === 'lower' ? tr('citLower', { v: kwhStr(last[j]) }) : tr(e)) : '';
    inp.setAttribute('aria-invalid', show ? 'true' : 'false');
    info.textContent = tr('lastRead', { v: kwhStr(last[j]) }) + (e ? '' : ' · ' + tr('usageNow', { v: kwhStr(chk.r[j] - last[j]) }));
  });
  return chk;
}

function savePanelReadings() {
  var c = state.cit, msg = $('c-msg'), btn = $('c-save');
  if (!c.key || c.saving || !msg) return;
  var chk = refreshCit(true);
  if (!chk.ok) {
    msg.className = 'sheet-msg bad';
    msg.textContent = tr('citFix');
    var bad = METERS.filter(function (j) { return chk.errors[j]; })[0];
    if (bad != null) $('c-r' + bad).focus();
    return;
  }
  var body = { t: todayISO(), r: chk.r };
  c.saving = true;
  btn.disabled = true;
  btn.textContent = tr('citSaving');
  msg.className = 'sheet-msg';
  msg.textContent = '';
  isCurrent().then(function (cur) {
    if (!cur) throw { code: 'stale' };
    return ghWriteFile(c.key, WEB.pending, JSON.stringify(body) + '\n', 'Citiri contoare ' + body.t);
  }).then(function () {
    c.saving = false;
    state.pendingRaw = body;
    c.touched = {};
    renderAll();
    var m2 = $('c-msg');
    if (m2) { m2.className = 'sheet-msg ok'; m2.textContent = tr('citDone'); }
    showToast(tr('citDone'));
  }, function (e) {
    c.saving = false;
    btn.disabled = false;
    btn.innerHTML = icon('check') + esc(tr('citSave'));
    msg.className = 'sheet-msg bad';
    var code = e && e.code;
    msg.textContent = code === 'badToken' || code === 'noAccess' ? tr('citExpired') : ghErrorText(e);
  });
}

/* ------------------------------------------------------------------ keeping an open page current

   A page left open across an update keeps running the old code. When it comes back into view it
   compares its app.js fingerprint with the published index.html and reloads (unless something is
   being typed); after a long time away it reloads anyway, for fresh figures. Saves refuse to run
   with outdated code. */
function runningVersion() {
  var s = document.querySelector('script[src*="app.js"]');
  var m = s && /[?&]v=([0-9a-f]+)/.exec(s.getAttribute('src') || '');
  return m ? m[1] : null;
}
/* true when this page runs the published code, or when that cannot be told */
function isCurrent() {
  var mine = runningVersion();
  if (!mine) return Promise.resolve(true);
  return fetch('index.html?t=' + Date.now(), { cache: 'no-store' }).then(function (r) { return r.ok ? r.text() : ''; }).then(function (t) {
    var m = /app\.js\?v=([0-9a-f]+)/.exec(t);
    return !m || m[1] === mine;
  }).catch(function () { return true; });
}
function busy() {
  var dlg = $('admin-dlg');
  return !!state.sheet || !!(dlg && dlg.open) || (state.view === 'panel' && !!state.cit.key);
}
var hiddenAt = 0;
function onVisibility() {
  if (document.visibilityState === 'hidden') { hiddenAt = Date.now(); return; }
  var away = hiddenAt ? Date.now() - hiddenAt : 0;
  if (!state.data || away < 60000) return;
  isCurrent().then(function (cur) {
    if (!cur) { if (busy()) showToast(tr('webNewVersion')); else location.reload(); }
    else if (away > 15 * 60000 && !busy()) location.reload();
  });
}

function submitUnlock() {
  var c = state.cit, input = $('c-pass'), btn = $('c-open');
  if (!input || !c.info || btn.disabled) return;
  if (!input.value) { input.focus(); return; }
  btn.disabled = true;
  btn.textContent = tr('citChecking');
  unlockPanel(input.value, !!($('c-remember') && $('c-remember').checked), false).then(function () {
    renderAll();
    var first = $('c-r0');
    if (first) first.focus();
  });
}

/* ------------------------------------------------------------------ admin dialog */

function adminDialog() {
  var dlg = $('admin-dlg');
  if (!dlg) {
    dlg = document.createElement('dialog');
    dlg.id = 'admin-dlg';
    dlg.className = 'web-dlg';
    dlg.setAttribute('aria-labelledby', 'adm-title');
    document.body.appendChild(dlg);
    dlg.addEventListener('click', onAdminClick);
    dlg.addEventListener('submit', onAdminSubmit);
    dlg.addEventListener('close', function () { document.body.classList.remove('locked'); });
  }
  return dlg;
}

function renderAdmin() {
  var dlg = adminDialog(), g = state.gh, info = state.cit.info;
  var h = '<div class="sheet-head"><h2 class="h2" id="adm-title">' + esc(tr('adminTitle')) + '</h2>' +
    '<button type="button" class="icon-btn" data-adm="close" aria-label="' + esc(tr('close')) + '">' + icon('x') + '</button></div><div class="sheet-body">';
  if (g) {
    h += '<p class="okline">' + icon('check') + esc(tr('adminConnected', { u: g.owner })) + '</p>';
    h += panelAdminHTML();
    h += '<div class="sheet-actions adm-foot"><button type="button" class="btn danger" data-adm="disconnect">' + esc(tr('adminDisconnect')) + '</button></div>';
  } else if (info === undefined) {
    h += '<p class="hint">' + esc(tr('panChecking')) + '</p>';
  } else if (info && state.admMode !== 'setup') {
    /* the usual way in: the password opens the key kept on the site */
    h += '<form id="adm-pass-form" novalidate><p class="hint">' + esc(tr('adminPassIntro')) + '</p>' +
      '<div class="field"><label for="adm-pass">' + esc(tr('adminPassLbl')) + '</label>' +
      '<input type="password" id="adm-pass" autocomplete="current-password" autocapitalize="none" spellcheck="false" enterkeyhint="go" aria-describedby="adm-pass-m">' +
      '<small class="msg" id="adm-pass-m" role="alert"></small></div>' +
      '<div class="sheet-actions"><button type="submit" class="btn primary" id="adm-pass-btn">' + icon('key') + esc(tr('adminOpen')) + '</button>' +
      '<button type="button" class="btn ghost" data-adm="newkey">' + esc(tr('adminNewKey')) + '</button></div></form>';
  } else {
    /* the first time, or a new key: the key and the password that will lock it */
    h += '<form id="adm-setup" novalidate><p class="hint">' + esc(tr('adminSetupIntro')) + '</p>' +
      '<div class="field"><label for="adm-token">' + esc(tr('adminToken')) + '</label>' +
      '<input type="password" id="adm-token" autocomplete="off" autocapitalize="none" spellcheck="false" aria-describedby="adm-token-m">' +
      '<small class="msg" id="adm-token-m" role="alert"></small></div><div class="fgrid adm-pw">' +
      '<div class="field"><label for="adm-p1">' + esc(tr('panPass')) + '</label><input type="password" id="adm-p1" autocomplete="new-password" autocapitalize="none" spellcheck="false"></div>' +
      '<div class="field"><label for="adm-p2">' + esc(tr('panPass2')) + '</label><input type="password" id="adm-p2" autocomplete="new-password" autocapitalize="none" spellcheck="false"></div>' +
      '</div><small class="msg" id="adm-setup-m" role="alert"></small><p class="hint pan-risk">' + esc(tr('panRisk')) + '</p>' +
      '<div class="sheet-actions"><button type="submit" class="btn primary" id="adm-connect-btn">' + icon('key') + esc(tr('adminSetupBtn')) + '</button>' +
      (info ? '<button type="button" class="btn ghost" data-adm="usepass">' + esc(tr('adminUsePass')) + '</button>' : '') + '</div></form>';
  }
  dlg.innerHTML = h + '</div>';
}

/* the password, once connected: change it, or switch it off (panel.json is then just {}) */
state.panAdmin = { status: undefined };
state.admMode = null;
function panelAdminHTML() {
  var s = state.panAdmin.status, on = s === 'on';
  var h = '<section class="adm-sec" aria-labelledby="pan-title"><h3 class="h3" id="pan-title">' + esc(tr('panTitle')) + '</h3>';
  if (s === undefined) return h + '<p class="hint">' + esc(tr('panChecking')) + '</p></section>';
  h += '<p class="hint">' + esc(tr(on ? 'panOnText' : 'panOffText')) + '</p><p class="hint pan-risk">' + esc(tr('panRisk')) + '</p>' +
    '<form id="pan-form" novalidate><div class="fgrid">' +
    '<div class="field"><label for="pan-p1">' + esc(tr('panPass')) + '</label><input type="password" id="pan-p1" autocomplete="new-password" autocapitalize="none" spellcheck="false"></div>' +
    '<div class="field"><label for="pan-p2">' + esc(tr('panPass2')) + '</label><input type="password" id="pan-p2" autocomplete="new-password" autocapitalize="none" spellcheck="false"></div>' +
    '</div><p class="msg" id="pan-m" role="alert"></p><div class="sheet-actions"><button type="submit" class="btn primary" id="pan-btn">' + esc(tr(on ? 'panChange' : 'panOn')) + '</button>' +
    (on ? '<button type="button" class="btn danger" data-adm="pan-off">' + esc(tr('panOffBtn')) + '</button>' : '') + '</div></form>';
  return h + '</section>';
}

function checkPanelStatus() {
  var g = state.gh;
  if (!g) return;
  ghGetFile(g, WEB.panel).then(function (f) {
    var env = ghFileJSON(f);
    state.panAdmin.status = env && env.ct ? 'on' : 'off';
  }, function () { state.panAdmin.status = 'off'; }).then(function () {
    var dlg = $('admin-dlg');
    if (dlg && dlg.open && state.gh) renderAdmin();
  });
}

/* locks the connected key with a password and stores it on the site (panel.json) */
function savePassword(p1) {
  var g = state.gh;
  return encryptJSON({ token: g.token, owner: g.owner }, p1).then(function (env) {
    return ghWriteFile(g, WEB.panel, JSON.stringify(env) + '\n', 'Parola de acces salvată').then(function () {
      state.cit.info = env;
      state.cit.key = { token: g.token, owner: g.owner };
      state.panAdmin.status = 'on';
      lsSet('ff-cit', p1);
    });
  });
}

function passwordProblem(p1, p2) {
  if (p1.length < WEB.minPanelPass) return { text: tr('panShort', { n: WEB.minPanelPass }), field: 1 };
  if (p1 !== p2) return { text: tr('panMismatch'), field: 2 };
  return null;
}

function connectErrorText(e) {
  var code = e && e.code;
  if (!code && typeof console !== 'undefined') console.error(e);
  if (code === 'unreachable' || code === 'offline') return ghErrorText(e);
  return code === 'badToken' ? tr('adminBadToken') : code === 'stale' ? tr('webStale') :
    code === 'wrongAccount' ? tr('adminWrongAccount', { u: e.login || '?', o: WEB.owner }) :
    tr('adminNoRepo') + (e && e.status ? ' (GitHub: ' + e.status + ')' : '');
}

/* the first time: check the key, then lock it with the chosen password */
function setupWithKey() {
  var input = $('adm-token'), btn = $('adm-connect-btn'), msg = $('adm-token-m'), pm = $('adm-setup-m');
  var token = cleanToken(input.value), p1 = $('adm-p1').value, p2 = $('adm-p2').value;
  msg.className = 'msg bad';
  pm.className = 'msg bad';
  msg.textContent = '';
  pm.textContent = '';
  if (!token) { input.focus(); return; }
  var bad = passwordProblem(p1, p2);
  if (bad) { pm.textContent = bad.text; $('adm-p' + bad.field).focus(); return; }
  if (btn.disabled) return;
  btn.disabled = true;
  btn.textContent = tr('adminChecking');
  isCurrent().then(function (cur) {
    if (!cur) throw { code: 'stale' };
    return connectOwner(token, true);
  }).then(function () {
    return savePassword(p1).then(function () {
      state.admMode = null;
      $('admin-dlg').close();
      showToast(tr('adminPassSavedT'));
    }, function (e) {
      /* connected, but the password could not be stored: say so where it can be tried again */
      state.panAdmin.status = 'off';
      renderAdmin();
      var m = $('pan-m');
      if (m) { m.className = 'msg bad'; m.textContent = ghErrorText(e); }
    });
  }, function (e) {
    btn.disabled = false;
    btn.innerHTML = icon('key') + esc(tr('adminSetupBtn'));
    msg.textContent = connectErrorText(e);
  });
}

/* the usual way: the password opens the key kept on the site */
function openWithPassword() {
  var input = $('adm-pass'), btn = $('adm-pass-btn'), msg = $('adm-pass-m'), pass = input.value;
  if (!pass) { input.focus(); return; }
  if (btn.disabled) return;
  btn.disabled = true;
  btn.textContent = tr('adminChecking');
  msg.className = 'msg bad';
  msg.textContent = '';
  decryptJSON(state.cit.info, pass).then(function (k) {
    if (!k || !k.token) throw { code: 'wrongPass' };
    return k;
  }, function () { throw { code: 'wrongPass' }; }).then(function (k) {
    return isCurrent().then(function (cur) {
      if (!cur) throw { code: 'stale' };
      return connectOwner(k.token, true);
    }).then(function () {
      state.cit.key = { token: k.token, owner: WEB.owner };
      lsSet('ff-cit', pass);
      state.panAdmin.status = 'on';
      $('admin-dlg').close();
      showToast(tr('webConnectedT'));
    });
  }).catch(function (e) {
    btn.disabled = false;
    btn.innerHTML = icon('key') + esc(tr('adminOpen'));
    var code = e && e.code;
    msg.textContent = code === 'wrongPass' ? tr('adminWrongPass') : code === 'badToken' ? tr('adminKeyExpired') : connectErrorText(e);
    if (code === 'wrongPass') input.select();
  });
}

/* connected: a new password for the key kept on the site */
function activatePanel() {
  var p1 = $('pan-p1').value, p2 = $('pan-p2').value, m = $('pan-m'), btn = $('pan-btn');
  m.className = 'msg bad';
  var bad = passwordProblem(p1, p2);
  if (bad) { m.textContent = bad.text; $('pan-p' + bad.field).focus(); return; }
  if (!state.gh || btn.disabled) return;
  btn.disabled = true;
  m.textContent = '';
  isCurrent().then(function (cur) {
    if (!cur) throw { code: 'stale' };
    return savePassword(p1);
  }).then(function () {
    $('admin-dlg').close();
    showToast(tr('panOnT'));
  }, function (e) {
    btn.disabled = false;
    m.textContent = ghErrorText(e);
  });
}

function deactivatePanel() {
  var g = state.gh;
  if (!g) return;
  ghWriteFile(g, WEB.panel, '{}\n', 'Parola de acces oprită').then(function () {
    state.panAdmin.status = 'off';
    state.cit.info = null;
    state.cit.key = null;
    lsDel('ff-cit');
    $('admin-dlg').close();
    showToast(tr('panOffT'));
  }, function (e) {
    var m = $('pan-m');
    if (m) { m.className = 'msg bad'; m.textContent = ghErrorText(e); }
  });
}

function focusAdmin() {
  var f = $('adm-pass') || $('adm-token');
  if (f) f.focus();
}

function openAdmin() {
  var dlg = adminDialog();
  state.admMode = null;
  if (state.gh) { state.panAdmin.status = undefined; checkPanelStatus(); }
  else if (state.cit.info === undefined) loadPanelInfo().then(function () { if (dlg.open && !state.gh) { renderAdmin(); focusAdmin(); } });
  renderAdmin();
  document.body.classList.add('locked');
  if (!dlg.open) dlg.showModal();
  focusAdmin();
}

function onAdminClick(ev) {
  var t = ev.target.closest('[data-adm]');
  if (!t) {
    if (ev.target === $('admin-dlg')) $('admin-dlg').close();
    return;
  }
  var a = t.getAttribute('data-adm');
  if (a === 'close') $('admin-dlg').close();
  else if (a === 'newkey' || a === 'usepass') { state.admMode = a === 'newkey' ? 'setup' : null; renderAdmin(); focusAdmin(); }
  else if (a === 'pan-off') deactivatePanel();
  else if (a === 'disconnect') {
    lsDel('ff-token');
    lsDel('ff-cit');
    state.gh = null;
    state.cit.key = null;
    state.webCanEdit = false;
    $('admin-dlg').close();
    updateEditability();
    showToast(tr('webDisconnectedT'));
  }
}

function onAdminSubmit(ev) {
  ev.preventDefault();
  var id = ev.target.id;
  if (id === 'adm-setup') setupWithKey();
  else if (id === 'adm-pass-form') openWithPassword();
  else if (id === 'pan-form') activatePanel();
}

/* ------------------------------------------------------------------ events */

function bindWebEvents() {
  root.addEventListener('click', function (ev) {
    var link = ev.target.closest('a[data-route]');
    if (link && !(ev.metaKey || ev.ctrlKey || ev.shiftKey || ev.altKey)) {
      ev.preventDefault();
      setView(link.getAttribute('data-route'), true);
      return;
    }
    var t = ev.target.closest('[data-action]');
    if (!t) return;
    var act = t.getAttribute('data-action');
    if (act === 'f-prev' || act === 'f-next') {
      var paid = paidMonths(), k = paid.indexOf(frontIndex(paid)) + (act === 'f-prev' ? -1 : 1);
      if (k < 0 || k >= paid.length) return;
      state.fsel = paid[k];
      var y = window.scrollY;
      renderAll();
      window.scrollTo(0, y);
      var again = root.querySelector('[data-action="' + act + '"]:not([disabled])') || root.querySelector('[data-action^="f-"]:not([disabled])');
      if (again) again.focus();
      return;
    }
    if (act === 'web-lang') {
      state.lang = t.getAttribute('data-v') === 'en' ? 'en' : 'ro';
      lsSet('ff-lang', state.lang);
      document.documentElement.lang = state.lang;
      if (state.data) renderAll();
      else loadSite();
    } else if (act === 'web-admin') openAdmin();
    else if (act === 'cit-lock') {
      lsDel('ff-cit');
      state.cit.key = null;
      state.cit.passMsg = '';
      renderAll();
    }
  });
  root.addEventListener('input', function (ev) {
    var t = ev.target, j = t.getAttribute && t.getAttribute('data-cit');
    if (j == null || !state.cit.key) return;
    state.cit.draft[+j] = t.value;
    if (String(t.value).trim() !== '') state.cit.touched[+j] = true;
    refreshCit(false);
  });
  root.addEventListener('focusout', function (ev) {
    var t = ev.target, j = t.getAttribute && t.getAttribute('data-cit');
    if (j == null || !state.cit.key || state.cit.touched[+j]) return;
    state.cit.touched[+j] = true;
    refreshCit(false);
  });
  root.addEventListener('submit', function (ev) {
    if (ev.target.id === 'cit-unlock') { ev.preventDefault(); submitUnlock(); }
    else if (ev.target.id === 'cit-form') { ev.preventDefault(); savePanelReadings(); }
  });
  document.addEventListener('visibilitychange', onVisibility);
  /* the browser's back and forward buttons move between the front page, #admin and #citiri */
  window.addEventListener('popstate', function () { setView(viewFromHash(), false); });
  window.addEventListener('hashchange', function () { setView(viewFromHash(), false); });
}

/* The front page by default, the full app under #admin. Until the data has loaded there is
   nothing to draw (a link change would otherwise try). */
var coreRenderAll = renderAll;
renderAll = function () {
  if (!state.data) return;
  if (state.view === 'panel') { renderPanel(); return; }
  if (state.view !== 'admin') { renderFront(); return; }
  coreRenderAll();
  var w = root.querySelector('.wrap');
  if (w) w.insertAdjacentHTML('afterbegin', adminBarHTML());
};

if (typeof document !== 'undefined' && document.getElementById('app') && !document.getElementById('ff-data')) {
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
}
