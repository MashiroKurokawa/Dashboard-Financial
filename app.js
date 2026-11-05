/* Dashboard Keuangan Warung Mie Ayam
   Alur: Excel (database) -> parse() -> calc(skenario) -> render() */
const DATA_URL = 'data/kelompok_3_warung_makan_mie_sederhana.xlsx';
const $ = s => document.querySelector(s);
const rp = n => 'Rp\u00A0' + Math.round(n).toLocaleString('id-ID');
const jt = n => Math.abs(n) >= 1e6 ? 'Rp\u00A0' + (n / 1e6).toLocaleString('id-ID', {maximumFractionDigits: 1}) + ' jt' : rp(n);
const pct = n => (n * 100).toLocaleString('id-ID', {maximumFractionDigits: 1}) + '%';
let D, base, S, charts = {};

/* ---------- 1. Excel -> objek data ---------- */
function toSheets(wb) {
  const o = {};
  wb.SheetNames.forEach(n => o[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], {header: 1, defval: null}));
  return o;
}
function parse(sh) {
  const num = (s, label) => {
    const r = (sh[s] || []).find(r => String(r[0] || '').trim().startsWith(label));
    const v = r && r.slice(1).find(x => typeof x === 'number');
    if (v === undefined) throw new Error(`Baris "${label}" tidak ditemukan di sheet "${s}"`);
    return v;
  };
  const between = (s, from, to) => {
    const rows = sh[s] || [], i = rows.findIndex(r => r[0] === from);
    const j = rows.findIndex((r, k) => k > i && String(r[0] || '').startsWith(to));
    return rows.slice(i + 1, j < 0 ? undefined : j).filter(r => r[0]);
  };
  const H = 'HPP';
  return {
    recipe: between(H, 'Komponen', 'Total Biaya Bahan Baku').map(r => ({name: r[0], qty: r[1], unit: r[2], price: r[3], total: r[4]})),
    bowls: num(H, 'Mangkok per batch'), pack: num(H, 'Kemasan, sumpit'),
    daily: num(H, 'Target penjualan harian'), days: num(H, 'Hari kerja per bulan'),
    tubes: num(H, 'Gas LPG 3 kg: jumlah tabung'), tubePrice: num(H, 'Gas LPG 3 kg: harga'), sanit: num(H, 'Sanitasi dapur per bulan'),
    price: num('Harga Jual', 'Rekomendasi harga jual'), lo: num('Harga Jual', 'Kisaran harga pasar Jakarta Pusat (bawah)'), hi: num('Harga Jual', 'Kisaran harga pasar Jakarta Pusat (atas)'),
    fixed: [['Listrik & air PAM', 'Listrik'], ['Penyusutan peralatan', 'Penyusutan'], ['Iuran RT & kebersihan', 'Iuran RT']].map(([n, k]) => ({name: n, v: num('BEP', k)})),
    cls: between('Klasifikasi', 'Komponen Biaya', 'Asumsi').map(r => ({name: r[0], type: r[1], why: r[2]})),
    assump: Object.fromEntries(['Biaya Operasional', 'Klasifikasi', 'HPP', 'Harga Jual', 'BEP', 'Proyeksi_LR'].map(s => {
      const rows = sh[s] || [], i = rows.findIndex(r => r[0] === 'Asumsi');
      return [s, i < 0 ? [] : rows.slice(i + 1).map(r => r[0]).filter(Boolean)];
    }))
  };
}

/* ---------- 2. Model keuangan ---------- */
function calc(d, s) {
  const k = 1 + s.infl / 100, vol = s.daily * s.days;
  const items = d.recipe.map(x => ({n: x.name, v: x.total * k / d.bowls}));
  items.push({n: 'Kemasan & alat sekali pakai', v: d.pack}, {n: 'Gas LPG 3 kg', v: d.tubes * d.tubePrice / vol}, {n: 'Sanitasi dapur', v: d.sanit / vol});
  const hpp = items.reduce((a, x) => a + x.v, 0), cm = s.price - hpp, fc = d.fixed.reduce((a, x) => a + x.v, 0);
  const rev = s.price * vol, vc = hpp * vol, gp = rev - vc, net = gp - fc, bep = cm > 0 ? Math.ceil(fc / cm) : Infinity;
  return {items, hpp, cm, fc, vol, rev, vc, gp, net, bep, margin: net / rev, mos: (vol - bep) / vol};
}

/* ---------- 3. Tampilan ---------- */
const COL = ['#1F6B47', '#E0A32B', '#C23B2A', '#5B3A24', '#7FA88F', '#F0C96B', '#D9826F', '#9C7A5F', '#3E8E66', '#B9C4BB'];
function mk(id, cfg) { charts[id] = new Chart($('#' + id), Object.assign({options: {}}, cfg)); }
Chart.defaults.font.family = '"IBM Plex Sans", sans-serif'; Chart.defaults.maintainAspectRatio = false;

function render() {
  const m = calc(D, S), pos = m.net >= 0;
  $('#lead').textContent = `Dengan ${S.daily} mangkok per hari selama ${S.days} hari dan harga ${rp(S.price)}, warung ini ${pos ? 'menghasilkan laba bersih' : 'mengalami rugi'} ${jt(Math.abs(m.net))} per bulan.`;
  $('#kpis').innerHTML = [
    ['Laba bersih / bulan', jt(m.net), `margin ${pct(m.margin)}`], ['Harga pokok / mangkok', rp(m.hpp), `margin kontribusi ${rp(m.cm)}`],
    ['Titik impas', isFinite(m.bep) ? m.bep + ' mangkok' : 'tidak tercapai', isFinite(m.bep) ? `≈ ${Math.ceil(m.bep / S.days)} per hari · omzet ${jt(m.bep * S.price)}` : 'harga di bawah HPP'],
    ['Batas aman penjualan', isFinite(m.bep) ? pct(m.mos) : '–', 'penjualan boleh turun sebelum rugi']
  ].map(([a, b, c]) => `<div class="kpi"><small>${a}</small><span class="big">${b}</span><small>${c}</small></div>`).join('');
  $('#rcSub').textContent = `${m.vol.toLocaleString('id-ID')} mangkok · ${S.days} hari kerja`;
  $('#rc').innerHTML = [['Omzet', m.rev, ''], ['Biaya variabel (HPP)', -m.vc, 'neg'], ['Laba kotor', m.gp, ''], ['Biaya tetap', -m.fc, 'neg'], ['Laba bersih', m.net, 'tot']]
    .map(([a, v, c]) => `<div class="${c}"><dt>${a}</dt><dd>${v < 0 ? '−' : ''}${rp(Math.abs(v))}</dd></div>`).join('');
  const warn = S.price < D.lo || S.price > D.hi ? ` Harga ${rp(S.price)} di luar kisaran pasar ${rp(D.lo)}–${rp(D.hi)}.` : '';
  $('#bepNote').textContent = `Biaya tetap ${rp(m.fc)} per bulan. Setiap mangkok menyumbang ${rp(m.cm)} setelah HPP.` + warn;

  $('#tHpp').innerHTML = '<tr><th>Komponen</th><th class="r">Per mangkok</th><th class="r">Porsi</th></tr>' +
    m.items.map(x => `<tr><td>${x.n}</td><td class="r">${rp(x.v)}</td><td class="r">${pct(x.v / m.hpp)}</td></tr>`).join('') +
    `<tr><th>Total HPP</th><th class="r">${rp(m.hpp)}</th><th></th></tr>`;
  $('#tCls').innerHTML = '<tr><th>Komponen</th><th>Klasifikasi</th><th>Dasar penentuan</th></tr>' +
    D.cls.map(x => { const f = /tetap/i.test(x.type); return `<tr><td>${x.name}</td><td><span class="tag ${f ? 'f' : 'v'}">${f ? 'Tetap' : 'Variabel'}</span></td><td>${x.why}</td></tr>`; }).join('');

  const xmax = Math.max(m.vol * 1.4, (isFinite(m.bep) ? m.bep : 0) * 2);
  const pts = {rev: [{x: 0, y: 0}, {x: xmax, y: S.price * xmax}], cost: [{x: 0, y: m.fc}, {x: xmax, y: m.fc + m.hpp * xmax}]};
  const cv = charts.cCvp.data.datasets;
  cv[0].data = pts.rev; cv[1].data = pts.cost; cv[2].data = isFinite(m.bep) ? [{x: m.bep, y: m.bep * S.price}] : []; cv[3].data = [{x: m.vol, y: m.rev}];
  charts.cHpp.data.labels = m.items.map(x => x.n); charts.cHpp.data.datasets[0].data = m.items.map(x => Math.round(x.v));
  charts.cLr.data.datasets[0].data = [m.rev, m.vc, m.gp, m.fc, m.net];
  Object.values(charts).forEach(c => c.update('none'));
}

function build() {
  S = {price: D.price, daily: D.daily, days: D.days, infl: 0}; base = {...S};
  const defs = [['price', 'Harga jual per mangkok', 10000, 25000, 500, rp], ['daily', 'Penjualan per hari (mangkok)', 10, 100, 1, v => v], ['days', 'Hari kerja per bulan', 15, 30, 1, v => v], ['infl', 'Perubahan harga bahan baku', -20, 40, 1, v => (v > 0 ? '+' : '') + v + '%']];
  $('#sliders').innerHTML = defs.map(([k, t, a, b, st]) => `<div><label for="s_${k}">${t}<output id="o_${k}"></output></label><input id="s_${k}" type="range" min="${a}" max="${b}" step="${st}"></div>`).join('');
  const sync = () => defs.forEach(([k, , , , , f]) => { $('#s_' + k).value = S[k]; $('#o_' + k).textContent = f(S[k]); });
  defs.forEach(([k]) => $('#s_' + k).oninput = e => { S[k] = +e.target.value; sync(); render(); });
  $('#reset').onclick = () => { S = {...base}; sync(); render(); };
  $('#assump').innerHTML = Object.entries(D.assump).filter(([, a]) => a.length).map(([s, a]) => `<details><summary>Asumsi · ${s.replace('_', ' ')}</summary><ol>${a.map(t => `<li>${t.replace(/^\d+\.\s*/, '')}</li>`).join('')}</ol></details>`).join('');

  const tip = {callbacks: {label: c => ' ' + (c.dataset.label ? c.dataset.label + ': ' : '') + rp(c.parsed.y ?? c.parsed.x ?? c.parsed)}};
  Object.values(charts).forEach(c => c.destroy()); charts = {};
  mk('cHpp', {type: 'doughnut', data: {labels: [], datasets: [{data: [], backgroundColor: COL, borderWidth: 2, borderColor: '#FBFCF9'}]},
    options: {maintainAspectRatio: false, plugins: {legend: {position: 'right', labels: {boxWidth: 12}}, tooltip: {callbacks: {label: c => ` ${c.label}: ${rp(c.parsed)}`}}}}});
  const ln = (label, color, extra) => Object.assign({label, borderColor: color, backgroundColor: color, pointRadius: 0, borderWidth: 2.5, data: []}, extra);
  mk('cCvp', {type: 'line', data: {datasets: [ln('Omzet', '#1F6B47'), ln('Total biaya', '#C23B2A'), ln('Titik impas', '#E0A32B', {showLine: false, pointRadius: 8}), ln('Target bulan ini', '#1D2B24', {showLine: false, pointRadius: 6, pointStyle: 'rectRot'})]},
    options: {maintainAspectRatio: false, parsing: false, scales: {x: {type: 'linear', title: {display: true, text: 'Mangkok terjual per bulan'}}, y: {ticks: {callback: v => jt(v)}}}, plugins: {title: {display: true, text: 'Omzet vs total biaya'}, tooltip: {callbacks: {label: c => ` ${c.dataset.label}: ${Math.round(c.parsed.x)} mangkok · ${rp(c.parsed.y)}`}}}}});
  mk('cLr', {type: 'bar', data: {labels: ['Omzet', 'HPP', 'Laba kotor', 'Biaya tetap', 'Laba bersih'], datasets: [{data: [], backgroundColor: ['#1F6B47', '#C23B2A', '#7FA88F', '#5B3A24', '#E0A32B'], borderRadius: 3}]},
    options: {maintainAspectRatio: false, plugins: {legend: {display: false}, title: {display: true, text: 'Laba rugi bulanan'}, tooltip: {callbacks: {label: c => ' ' + rp(c.parsed.y)}}}, scales: {y: {ticks: {callback: v => jt(v)}}}}});
  sync(); render();
}

/* ---------- 4. Memuat database Excel ---------- */
function status(txt, cls) { const p = $('#src'); p.textContent = txt; p.className = 'pill ' + cls; }
function load(sheets, label, cls) {
  try { D = parse(sheets); build(); status(label, cls); }
  catch (e) { status('Format Excel tidak sesuai', 'warn'); alert(e.message); }
}
async function init() {
  try {
    const r = await fetch(DATA_URL + '?t=' + Date.now()); if (!r.ok) throw 0;
    load(toSheets(XLSX.read(await r.arrayBuffer())), 'Excel terbaca langsung', 'ok');
  } catch { load(window.SNAPSHOT, 'Snapshot (buka via Live Server untuk baca Excel)', 'warn'); }
}
$('#file').onchange = async e => { const f = e.target.files[0]; if (f) load(toSheets(XLSX.read(await f.arrayBuffer())), 'Excel: ' + f.name, 'ok'); };
init();