import React, { useState, useMemo, useCallback } from "react";
import {
  ResponsiveContainer, ComposedChart, BarChart, Bar, Line,
  XAxis, YAxis, CartesianGrid, Tooltip, Cell, ReferenceLine,
} from "recharts";
import * as XLSX from "xlsx";

/* ============================================================================
   COCKPIT DE VALORISATION · 5 methodes (DCF · Comparables · Patrimoniale ·
   Dividendes · EVA/MVA). Moteur live, sensibilites, football field,
   ponderation ajustable, export Excel / PDF.
   Donnees par defaut : HPS (Hightech Payment Systems), Bourse de Casablanca.
============================================================================ */

const M = {
  dcf:   { label: "DCF",          n: "01", color: "#38BDF8", sub: "Flux de tresorerie actualises" },
  comps: { label: "Comparables",  n: "02", color: "#A78BFA", sub: "Multiples boursiers" },
  anr:   { label: "Patrimoniale", n: "03", color: "#FBBF24", sub: "Actif net reevalue (ANR)" },
  ddm:   { label: "Dividendes",   n: "04", color: "#34D399", sub: "Gordon-Shapiro (DDM)" },
  eva:   { label: "EVA / MVA",    n: "05", color: "#FB7185", sub: "Creation de valeur" },
};
const KEYS = ["dcf", "comps", "anr", "ddm", "eva"];
const POS = "#34D399", NEG = "#F87171", MKT = "#EDEAE0";

const nf0 = new Intl.NumberFormat("fr-FR", { maximumFractionDigits: 0 });
const nf2 = new Intl.NumberFormat("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const f0 = (x) => (Number.isFinite(x) ? nf0.format(x) : "n.d.");
const f2 = (x) => (Number.isFinite(x) ? nf2.format(x) : "n.d.");
const fAct = (x) => (Number.isFinite(x) ? nf2.format(x) + " MAD" : "n.d.");
const fPct = (x) => (Number.isFinite(x) ? (x >= 0 ? "+" : "") + nf2.format(x * 100) + " %" : "n.d.");
const fC = (x) => {
  if (!Number.isFinite(x)) return "n.d.";
  const a = Math.abs(x);
  if (a >= 1e9) return nf2.format(x / 1e9) + " Md";
  if (a >= 1e6) return nf2.format(x / 1e6) + " M";
  return nf0.format(x);
};
const num = (v) => (v === "" || v === null || v === undefined ? 0 : parseFloat(v) || 0);

const mean = (a) => (a.length ? a.reduce((x, y) => x + y, 0) / a.length : NaN);
const median = (a) => {
  if (!a.length) return NaN;
  const b = [...a].sort((x, y) => x - y), m = Math.floor(b.length / 2);
  return b.length % 2 ? b[m] : (b[m - 1] + b[m]) / 2;
};
const steps = (base, count, step) => {
  const h = Math.floor(count / 2);
  return Array.from({ length: count }, (_, i) => +(base + (i - h) * step).toFixed(4));
};
const niceNum = (x) => {
  if (!Number.isFinite(x) || x <= 0) return 1;
  const exp = Math.floor(Math.log10(x)), f = x / Math.pow(10, exp);
  const nf = f < 1.5 ? 1 : f < 3 ? 2 : f < 7 ? 5 : 10;
  return nf * Math.pow(10, exp);
};

/* ------------------------- Moteurs de valorisation ----------------------- */
function calcDCF(d, shares) {
  const wacc = d.wacc / 100, g = d.g / 100, n = d.fcf.length;
  const disc = d.fcf.map((f, i) => num(f) / Math.pow(1 + wacc, i + 1));
  const sumDisc = disc.reduce((a, b) => a + b, 0);
  const valid = wacc > g;
  const last = num(d.fcf[n - 1]);
  const tv = valid ? (last * (1 + g)) / (wacc - g) : NaN;
  const tvDisc = valid ? tv / Math.pow(1 + wacc, n) : NaN;
  const ev = sumDisc + (Number.isFinite(tvDisc) ? tvDisc : 0);
  const equity = ev - num(d.netDebt);
  return { disc, sumDisc, tv, tvDisc, ev, equity, perShare: equity / shares, valid, n };
}
const dcfAt = (d, shares, wp, gp) => {
  const wacc = wp / 100, g = gp / 100, n = d.fcf.length;
  if (wacc <= g) return NaN;
  const s = d.fcf.reduce((a, f, i) => a + num(f) / Math.pow(1 + wacc, i + 1), 0);
  const tv = (num(d.fcf[n - 1]) * (1 + g)) / (wacc - g);
  return (s + tv / Math.pow(1 + wacc, n) - num(d.netDebt)) / shares;
};

function calcComps(c, shares) {
  const agg = (sel) => { const a = c.peers.map((p) => num(p[sel])); return { mean: mean(a), median: median(a) }; };
  const defs = [
    { key: "per",      label: "PER",       mult: agg("per"),      base: num(c.target.netIncome), ev: false },
    { key: "evEbitda", label: "EV/EBITDA", mult: agg("evEbitda"), base: num(c.target.ebitda),    ev: true },
    { key: "evEbit",   label: "EV/EBIT",   mult: agg("evEbit"),   base: num(c.target.ebit),      ev: true },
    { key: "evSales",  label: "EV/CA",     mult: agg("evSales"),  base: num(c.target.sales),     ev: true },
  ];
  const eq = (m, base, isEv) => { const v = m * base; return ((isEv ? v - num(c.netDebt) : v)) / shares; };
  const lines = defs.map((l) => ({ ...l, perMean: eq(l.mult.mean, l.base, l.ev), perMedian: eq(l.mult.median, l.base, l.ev) }));
  const mv = lines.map((l) => l.perMean);
  return { lines, centralMean: mean(mv), centralMedian: mean(lines.map((l) => l.perMedian)), low: Math.min(...mv), high: Math.max(...mv) };
}

function calcANR(p, shares) {
  const adj = p.adjustments.reduce((a, x) => a + num(x.amount), 0);
  const anr = num(p.bookEquity) + adj;
  const global = anr + num(p.goodwill);
  return { anr, global, adj, perShareANR: anr / shares, perShareGlobal: global / shares };
}

function calcDDM(m) {
  const ke = m.ke / 100, g = m.g / 100, valid = ke > g;
  const d1 = num(m.d0) * (1 + g);
  return { d1, value: valid ? d1 / (ke - g) : NaN, valid };
}
const ddmAt = (m, kp, gp) => { const ke = kp / 100, g = gp / 100; return ke > g ? (num(m.d0) * (1 + g)) / (ke - g) : NaN; };

function calcEVA(e, shares) {
  const wacc = e.wacc / 100, g = e.g / 100, n = e.series.length;
  const evas = e.series.map((s) => num(s.nopat) - num(s.ci) * wacc);
  const ci0 = n ? num(e.series[n - 1].ci) : 0;
  const pvEvas = evas.map((v, i) => v / Math.pow(1 + wacc, i + 1));
  const pvEva = pvEvas.reduce((a, b) => a + b, 0);
  const valid = wacc > g;
  const tv = valid ? (evas[n - 1] * (1 + g)) / (wacc - g) : NaN;
  const pvTV = valid ? tv / Math.pow(1 + wacc, n) : NaN;
  const ev = ci0 + pvEva + (Number.isFinite(pvTV) ? pvTV : 0);
  const equity = ev - num(e.netDebt);
  const marketEV = num(e.marketCap) + num(e.mvaNetDebt);
  return { evas, pvEvas, pvEva, tv, pvTV, ev, equity, perShare: equity / shares, ci0,
    valid, mvaMarket: marketEV - num(e.capitalInvested), mvaIntrinsic: pvEva + (Number.isFinite(pvTV) ? pvTV : 0), marketEV };
}
const evaAt = (e, shares, wp, gp) => {
  const wacc = wp / 100, g = gp / 100, n = e.series.length;
  if (wacc <= g) return NaN;
  const evas = e.series.map((s) => num(s.nopat) - num(s.ci) * wacc);
  const ci0 = num(e.series[n - 1].ci);
  const pv = evas.reduce((a, v, i) => a + v / Math.pow(1 + wacc, i + 1), 0);
  const tv = (evas[n - 1] * (1 + g)) / (wacc - g);
  return (ci0 + pv + tv / Math.pow(1 + wacc, n) - num(e.netDebt)) / shares;
};

function waterfall(items) {
  let run = 0;
  return items.map((s) => {
    if (s.kind === "total") {
      const v = s.amount != null ? s.amount : run;
      const row = { label: s.label, base: 0, bar: v, raw: v, kind: "total" };
      run = v; return row;
    }
    const start = run, end = run + s.amount; run = end;
    return { label: s.label, base: Math.min(start, end), bar: Math.abs(s.amount), raw: s.amount, kind: s.amount >= 0 ? "up" : "down" };
  });
}

/* ============================== UI primitives ============================ */
function Field({ label, value, onChange, suffix, step = "any" }) {
  return (
    <label className="vck-fld">
      <span>{label}</span>
      <div className="vck-inwrap">
        <input type="number" step={step} value={Number.isFinite(value) ? value : ""}
          onChange={(e) => onChange(e.target.value === "" ? 0 : parseFloat(e.target.value))} />
        {suffix && <em>{suffix}</em>}
      </div>
    </label>
  );
}
function Slider({ value, min, max, step, onChange, color }) {
  return (
    <input className="vck-range" type="range" min={min} max={max} step={step}
      value={value} style={{ accentColor: color }}
      onChange={(e) => onChange(parseFloat(e.target.value))} />
  );
}
function Kpi({ label, value, accent, big }) {
  return (
    <div className={"vck-kpi" + (big ? " big" : "")}>
      <span className="vck-kpi-l">{label}</span>
      <span className="vck-kpi-v mono" style={accent ? { color: accent } : null}>{value}</span>
    </div>
  );
}
function Tag({ ps, price, color }) {
  const up = Number.isFinite(ps) && Number.isFinite(price) ? (ps - price) / price : NaN;
  return (
    <div className="vck-tag">
      <b className="mono" style={{ color }}>{fAct(ps)}</b>
      <span className="mono" style={{ color: up >= 0 ? POS : NEG }}>{fPct(up)} vs cours</span>
    </div>
  );
}
const ChartTip = ({ active, payload, label }) => {
  if (!active || !payload || !payload.length) return null;
  return (
    <div className="vck-cttip">
      <div className="vck-cttip-l">{label}</div>
      {payload.filter((p) => p.dataKey !== "base").map((p, i) => (
        <div key={i} className="mono"><i style={{ background: p.color || p.fill }} />{p.name} : {fC(p.value)}</div>
      ))}
    </div>
  );
};
const WfTip = ({ active, payload }) => {
  if (!active || !payload || !payload.length) return null;
  const r = payload.find((p) => p.dataKey === "bar"); if (!r) return null;
  return <div className="vck-cttip"><div className="vck-cttip-l">{r.payload.label}</div><div className="mono">{fC(r.payload.raw)} MAD</div></div>;
};

function heatColor(v, mkt) {
  if (!Number.isFinite(v)) return "transparent";
  const t = Math.max(-1, Math.min(1, (v - mkt) / mkt / 0.6));
  const a = 0.1 + 0.55 * Math.abs(t);
  return t >= 0 ? "rgba(52,211,153," + a + ")" : "rgba(248,113,113," + a + ")";
}
function Heatmap({ cols, rows, fn, baseCol, baseRow, price, color, xlab, ylab }) {
  return (
    <div className="vck-heat">
      <table>
        <thead>
          <tr><th className="vck-heat-corner">{ylab}<i>/</i>{xlab}</th>
            {cols.map((c, i) => <th key={i} className={"mono" + (Math.abs(c - baseCol) < 1e-6 ? " base" : "")}>{f2(c)}</th>)}</tr>
        </thead>
        <tbody>
          {rows.map((r, ri) => (
            <tr key={ri}>
              <th className={"mono row" + (Math.abs(r - baseRow) < 1e-6 ? " base" : "")}>{f2(r)}</th>
              {cols.map((c, ci) => {
                const v = fn(c, r);
                const isBase = Math.abs(c - baseCol) < 1e-6 && Math.abs(r - baseRow) < 1e-6;
                return (
                  <td key={ci} className="mono" style={{ background: heatColor(v, price), boxShadow: isBase ? "inset 0 0 0 2px " + color : "none" }}
                    title={f2(v) + " MAD (WACC " + f2(c) + " / g " + f2(r) + ")"}>{Number.isFinite(v) ? f0(v) : ""}</td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function domainOf(bars, extra) {
  const lows = bars.map((b) => b.low).filter(Number.isFinite);
  const highs = bars.map((b) => b.high).filter(Number.isFinite);
  const ex = extra.filter(Number.isFinite);
  let lo = Math.min(...lows, ...ex);
  let hi = Math.max(...highs, ...ex);
  if (!Number.isFinite(lo)) lo = 0;
  if (!Number.isFinite(hi)) hi = 1;
  lo = lo >= 0 ? 0 : lo * 1.05;
  hi = hi * 1.06;
  return [lo, hi];
}
function FootballField({ bars, price, weighted }) {
  const [lo, hi] = domainOf(bars, [price, weighted]);
  const span = hi - lo || 1;
  const xp = (v) => Math.max(0, Math.min(100, ((v - lo) / span) * 100));
  const step = niceNum(span / 5);
  const ticks = [];
  for (let t = Math.ceil(lo / step) * step; t <= hi; t += step) ticks.push(+t.toFixed(2));
  return (
    <div className="vck-ff">
      <div className="vck-ff-top">
        <div className="vck-ff-mark" style={{ left: xp(price) + "%" }}><span className="vck-ff-pin mkt">Cours {f0(price)}</span></div>
        {Number.isFinite(weighted) && (
          <div className="vck-ff-mark" style={{ left: xp(weighted) + "%" }}><span className="vck-ff-pin wtd">Ponderee {f0(weighted)}</span></div>
        )}
      </div>
      <div className="vck-ff-plot">
        <div className="vck-ff-layer">
          <div className="vck-ff-grid">{ticks.map((t, i) => <i key={i} style={{ left: xp(t) + "%" }} />)}</div>
          {Number.isFinite(price) && <div className="vck-ff-vline mkt" style={{ left: xp(price) + "%" }} />}
          {Number.isFinite(weighted) && <div className="vck-ff-vline wtd" style={{ left: xp(weighted) + "%" }} />}
        </div>
        {bars.map((b) => {
          const L = xp(b.low), W = Math.max(xp(b.high) - L, 0.8), C = xp(b.central);
          return (
            <div className="vck-ff-row" key={b.key}>
              <div className="vck-ff-lab"><i style={{ background: b.color }} /><span className="disp">{b.label}</span><b className="mono">{f0(b.central)}</b></div>
              <div className="vck-ff-track">
                <div className="vck-ff-bar" style={{ left: L + "%", width: W + "%", background: "linear-gradient(90deg, " + b.color + "22, " + b.color + "66)", borderColor: b.color + "88" }} />
                <div className="vck-ff-tick" style={{ left: C + "%", background: b.color }} />
              </div>
            </div>
          );
        })}
      </div>
      <div className="vck-ff-axis">{ticks.map((t, i) => <span key={i} className="mono" style={{ left: xp(t) + "%" }}>{f0(t)}</span>)}</div>
    </div>
  );
}

function Section({ id, m, ps, price, children }) {
  return (
    <section id={id} className="vck-sec">
      <header className="vck-sec-h">
        <span className="vck-sec-n disp" style={{ color: m.color }}>{m.n}</span>
        <div className="vck-sec-t"><h2 className="disp">{m.label}</h2><p>{m.sub}</p></div>
        <div className="vck-sec-r"><Tag ps={ps} price={price} color={m.color} /></div>
      </header>
      {children}
    </section>
  );
}
const ChartBox = ({ title, children, h = 248 }) => (
  <div className="vck-chart">
    <div className="vck-chart-t disp">{title}</div>
    <div style={{ width: "100%", height: h }}><ResponsiveContainer>{children}</ResponsiveContainer></div>
  </div>
);

/* ================================== APP ================================== */
export default function App() {
  const [company, setCompany] = useState("HPS · Hightech Payment Systems");
  const [shares, setShares] = useState(7406190);
  const [price, setPrice] = useState(630);

  const [dcf, setDcf] = useState({
    fcf: [65464467, 259631981, 253208304, 192119682, 134189245],
    years: [2021, 2022, 2023, 2024, 2025],
    wacc: 8.24, g: 4.03, netDebt: 422306692,
  });
  const [comps, setComps] = useState({
    peers: [{ name: "Disway", per: 16.86, evEbitda: 12.25, evSales: 0.796, evEbit: 15.37 }],
    target: { netIncome: 106000000, ebitda: 286000000, ebit: 219702344, sales: 1493574508 },
    netDebt: 196000000,
  });
  const [anr, setAnr] = useState({
    bookEquity: 816636777,
    adjustments: [
      { label: "(-) Dividendes proposes", amount: -59249520 },
      { label: "(+) Plus-value immobilier", amount: 0 },
      { label: "(+) Plus-value stocks / en-cours", amount: 0 },
      { label: "(-) Provisions complementaires", amount: 0 },
    ],
    goodwill: 326818703,
  });
  const [ddm, setDdm] = useState({
    d0: 8, g: 4.03, ke: 8.725, projYears: 5,
    history: [{ year: 2021, dps: 5.5 }, { year: 2022, dps: 6 }, { year: 2023, dps: 6.8 }, { year: 2024, dps: 7 }, { year: 2025, dps: 8 }],
  });
  const [eva, setEva] = useState({
    series: [
      { year: 2021, nopat: 144400000, ci: 431562261 },
      { year: 2022, nopat: 163458407, ci: 563813433 },
      { year: 2023, nopat: 196353611, ci: 662262864 },
      { year: 2024, nopat: 153565401, ci: 1155617871 },
      { year: 2025, nopat: 219702344, ci: 1052675155 },
    ],
    wacc: 8.24, g: 4.03, netDebt: 422306692,
    marketCap: 4665899700, mvaNetDebt: 492171904, capitalInvested: 1052675155,
  });
  const [weights, setWeights] = useState({ dcf: 30, comps: 25, anr: 10, ddm: 15, eva: 20 });

  const rDcf = useMemo(() => calcDCF(dcf, shares), [dcf, shares]);
  const rComps = useMemo(() => calcComps(comps, shares), [comps, shares]);
  const rAnr = useMemo(() => calcANR(anr, shares), [anr, shares]);
  const rDdm = useMemo(() => calcDDM(ddm), [ddm]);
  const rEva = useMemo(() => calcEVA(eva, shares), [eva, shares]);

  const dcfGrid = useMemo(() => {
    const cols = steps(dcf.wacc, 7, 0.5), rows = steps(dcf.g, 5, 0.5).reverse();
    let mn = Infinity, mx = -Infinity;
    rows.forEach((r) => cols.forEach((c) => { const v = dcfAt(dcf, shares, c, r); if (Number.isFinite(v)) { mn = Math.min(mn, v); mx = Math.max(mx, v); } }));
    return { cols, rows, min: mn, max: mx };
  }, [dcf, shares]);
  const ddmGrid = useMemo(() => {
    const cols = steps(ddm.ke, 7, 0.5), rows = steps(ddm.g, 5, 0.5).reverse();
    let mn = Infinity, mx = -Infinity;
    rows.forEach((r) => cols.forEach((c) => { const v = ddmAt(ddm, c, r); if (Number.isFinite(v)) { mn = Math.min(mn, v); mx = Math.max(mx, v); } }));
    return { cols, rows, min: mn, max: mx };
  }, [ddm]);
  const evaGrid = useMemo(() => {
    const cols = steps(eva.wacc, 7, 0.5), rows = steps(eva.g, 5, 0.5).reverse();
    let mn = Infinity, mx = -Infinity;
    rows.forEach((r) => cols.forEach((c) => { const v = evaAt(eva, shares, c, r); if (Number.isFinite(v)) { mn = Math.min(mn, v); mx = Math.max(mx, v); } }));
    return { cols, rows, min: mn, max: mx };
  }, [eva, shares]);

  const central = { dcf: rDcf.perShare, comps: rComps.centralMean, anr: rAnr.perShareGlobal, ddm: rDdm.value, eva: rEva.perShare };
  const ranges = {
    dcf:   { low: dcfGrid.min, high: dcfGrid.max },
    comps: { low: rComps.low, high: rComps.high },
    anr:   { low: Math.min(rAnr.perShareANR, rAnr.perShareGlobal), high: Math.max(rAnr.perShareANR, rAnr.perShareGlobal) },
    ddm:   { low: ddmGrid.min, high: ddmGrid.max },
    eva:   { low: evaGrid.min, high: evaGrid.max },
  };
  const ffBars = KEYS.map((k) => ({ key: k, label: M[k].label, color: M[k].color, central: central[k], low: ranges[k].low, high: ranges[k].high }));

  const weighted = useMemo(() => {
    let ws = 0, acc = 0;
    KEYS.forEach((k) => { if (Number.isFinite(central[k])) { ws += weights[k]; acc += weights[k] * central[k]; } });
    return ws > 0 ? acc / ws : NaN;
  }, [weights, central.dcf, central.comps, central.anr, central.ddm, central.eva]);
  const wSum = KEYS.reduce((a, k) => a + weights[k], 0);
  const upside = Number.isFinite(weighted) ? (weighted - price) / price : NaN;
  const verdict = !Number.isFinite(upside) ? "" : upside >= 0 ? "Sous-evaluee · potentiel " + fPct(upside) : "Surevaluee · ecart " + fPct(upside);

  const upDcf = (patch) => setDcf((d) => ({ ...d, ...patch }));
  const setFcf = (i, v) => setDcf((d) => ({ ...d, fcf: d.fcf.map((x, j) => (j === i ? v : x)) }));
  const addFcf = () => setDcf((d) => ({ ...d, fcf: [...d.fcf, d.fcf[d.fcf.length - 1] || 0], years: [...d.years, (d.years[d.years.length - 1] || 2025) + 1] }));
  const delFcf = () => setDcf((d) => (d.fcf.length > 1 ? { ...d, fcf: d.fcf.slice(0, -1), years: d.years.slice(0, -1) } : d));
  const setPeer = (i, key, v) => setComps((c) => ({ ...c, peers: c.peers.map((p, j) => (j === i ? { ...p, [key]: v } : p)) }));
  const addPeer = () => setComps((c) => ({ ...c, peers: [...c.peers, { name: "Pair " + (c.peers.length + 1), per: 15, evEbitda: 10, evSales: 1, evEbit: 12 }] }));
  const delPeer = (i) => setComps((c) => (c.peers.length > 1 ? { ...c, peers: c.peers.filter((_, j) => j !== i) } : c));
  const setTgt = (k, v) => setComps((c) => ({ ...c, target: { ...c.target, [k]: v } }));
  const setAdj = (i, v) => setAnr((p) => ({ ...p, adjustments: p.adjustments.map((a, j) => (j === i ? { ...a, amount: v } : a)) }));
  const setEvaRow = (i, k, v) => setEva((e) => ({ ...e, series: e.series.map((s, j) => (j === i ? { ...s, [k]: v } : s)) }));

  const exportExcel = useCallback(() => {
    const wb = XLSX.utils.book_new();
    const r2 = (x) => (Number.isFinite(x) ? +x.toFixed(2) : null);
    const synth = [
      ["COCKPIT DE VALORISATION"], ["Societe", company], ["Cours de bourse (MAD)", price], ["Nombre d'actions", shares], [],
      ["Methode", "Valeur / action (MAD)", "Fourchette basse", "Fourchette haute", "Ecart vs cours (%)", "Ponderation (%)"],
      ...KEYS.map((k) => [M[k].label, r2(central[k]), r2(ranges[k].low), r2(ranges[k].high),
        r2(Number.isFinite(central[k]) ? ((central[k] - price) / price) * 100 : NaN), weights[k]]),
      [], ["Valeur ponderee (MAD)", r2(weighted)], ["Potentiel vs cours (%)", r2(upside * 100)],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(synth), "Synthese");
    const dcfS = [["DCF"], [], ["Annee", ...dcf.years], ["FCF (MAD)", ...dcf.fcf], ["FCF actualise (MAD)", ...rDcf.disc.map(r2)], [],
      ["WACC (%)", dcf.wacc], ["g terminal (%)", dcf.g], ["Dette nette (MAD)", dcf.netDebt], [],
      ["Somme FCF actualises", r2(rDcf.sumDisc)], ["Valeur terminale", r2(rDcf.tv)], ["VT actualisee", r2(rDcf.tvDisc)],
      ["Valeur d'entreprise", r2(rDcf.ev)], ["Capitaux propres", r2(rDcf.equity)], ["Valeur / action", r2(rDcf.perShare)]];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(dcfS), "DCF");
    const cS = [["COMPARABLES"], [], ["Pair", "PER", "EV/EBITDA", "EV/EBIT", "EV/CA"],
      ...comps.peers.map((p) => [p.name, p.per, p.evEbitda, p.evEbit, p.evSales]), [],
      ["Cible RN", comps.target.netIncome], ["EBITDA", comps.target.ebitda], ["EBIT", comps.target.ebit], ["CA", comps.target.sales], ["Dette nette", comps.netDebt], [],
      ["Methode", "Valeur/action (moy.)", "Valeur/action (med.)"],
      ...rComps.lines.map((l) => [l.label, r2(l.perMean), r2(l.perMedian)]),
      ["Moyenne globale", r2(rComps.centralMean), r2(rComps.centralMedian)]];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cS), "Comparables");
    const aS = [["PATRIMONIALE (ANR)"], [], ["Capitaux propres comptables", anr.bookEquity],
      ...anr.adjustments.map((a) => [a.label, a.amount]), ["ANR (ANCC)", r2(rAnr.anr)],
      ["+ Goodwill", anr.goodwill], ["Valeur globale", r2(rAnr.global)], [],
      ["Valeur/action ANR", r2(rAnr.perShareANR)], ["Valeur/action globale", r2(rAnr.perShareGlobal)]];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aS), "Patrimoniale");
    const dS = [["DIVIDENDES (DDM)"], [], ["Annee", ...ddm.history.map((h) => h.year)], ["DPA (MAD)", ...ddm.history.map((h) => h.dps)], [],
      ["D0 (MAD)", ddm.d0], ["g (%)", ddm.g], ["Ke (%)", ddm.ke], ["D1 (MAD)", r2(rDdm.d1)], ["Valeur / action", r2(rDdm.value)]];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(dS), "Dividendes");
    const eS = [["EVA / MVA"], [], ["Annee", ...eva.series.map((s) => s.year)], ["NOPAT", ...eva.series.map((s) => s.nopat)],
      ["Capital investi", ...eva.series.map((s) => s.ci)], ["EVA", ...rEva.evas.map(r2)], [],
      ["WACC (%)", eva.wacc], ["g (%)", eva.g], ["Somme PV(EVA)", r2(rEva.pvEva)], ["PV(VT)", r2(rEva.pvTV)],
      ["VE intrinseque", r2(rEva.ev)], ["Valeur / action", r2(rEva.perShare)], [],
      ["MVA intrinseque", r2(rEva.mvaIntrinsic)], ["MVA de marche", r2(rEva.mvaMarket)]];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(eS), "EVA-MVA");
    XLSX.writeFile(wb, "Valorisation_" + company.split(" ")[0] + ".xlsx");
  }, [company, price, shares, dcf, comps, anr, ddm, eva, weights, rDcf, rComps, rAnr, rDdm, rEva]);

  const sci = (id) => { const el = document.getElementById(id); if (el) el.scrollIntoView({ behavior: "smooth", block: "start" }); };

  const dcfWf = waterfall([
    ...rDcf.disc.map((v, i) => ({ label: "FCF act. " + dcf.years[i], amount: v, kind: "delta" })),
    { label: "VT actualisee", amount: rDcf.tvDisc || 0, kind: "delta" },
    { label: "VE", kind: "total" },
    { label: "- Dette nette", amount: -num(dcf.netDebt), kind: "delta" },
    { label: "Capitaux propres", kind: "total" },
  ]);
  const anrWf = waterfall([
    { label: "Cap. propres compt.", amount: num(anr.bookEquity), kind: "total" },
    ...anr.adjustments.filter((a) => num(a.amount) !== 0).map((a) => ({ label: a.label, amount: num(a.amount), kind: "delta" })),
    { label: "ANR (ANCC)", kind: "total" },
    { label: "+ Goodwill", amount: num(anr.goodwill), kind: "delta" },
    { label: "Valeur globale", kind: "total" },
  ]);
  const mvaWf = waterfall([
    { label: "Capital investi", amount: rEva.ci0, kind: "total" },
    { label: "+ PV(EVA)", amount: rEva.pvEva, kind: "delta" },
    { label: "+ PV(VT)", amount: rEva.pvTV || 0, kind: "delta" },
    { label: "VE intrinseque", kind: "total" },
  ]);
  const wfColor = (k, c) => (k === "total" ? c : k === "up" ? c + "99" : NEG);
  const multiplesData = ["PER", "EV/EBITDA", "EV/EBIT", "EV/CA"].map((lab, i) => {
    const key = ["per", "evEbitda", "evEbit", "evSales"][i];
    const row = { multiple: lab };
    comps.peers.forEach((p) => { row[p.name] = num(p[key]); });
    return row;
  });
  const dividendData = (() => {
    const out = ddm.history.map((h) => ({ year: String(h.year), hist: num(h.dps), proj: null }));
    const last = ddm.history[ddm.history.length - 1];
    let cur = num(last.dps);
    out[out.length - 1].proj = cur;
    for (let i = 1; i <= ddm.projYears; i++) { cur = cur * (1 + ddm.g / 100); out.push({ year: String(num(last.year) + i), hist: null, proj: cur }); }
    return out;
  })();
  const evaData = eva.series.map((s, i) => ({ year: String(s.year), eva: rEva.evas[i] }));

  return (
    <div className="vck">
      <style>{CSS}</style>

      <header className="vck-top no-print">
        <div className="vck-top-l">
          <div className="vck-logo disp">VALO<span>·</span>5</div>
          <input className="vck-co" value={company} onChange={(e) => setCompany(e.target.value)} spellCheck={false} />
        </div>
        <div className="vck-top-r">
          <label className="vck-mini"><span>Cours</span><input type="number" value={price} onChange={(e) => setPrice(num(e.target.value))} /><em>MAD</em></label>
          <label className="vck-mini"><span>Actions</span><input type="number" value={shares} onChange={(e) => setShares(num(e.target.value) || 1)} /></label>
          <button className="vck-btn" onClick={exportExcel}>{ICON.xls}Excel</button>
          <button className="vck-btn" onClick={() => window.print()}>{ICON.pdf}PDF</button>
        </div>
      </header>

      <div className="vck-wrap">
        <nav className="vck-nav no-print">
          {KEYS.map((k) => (
            <button key={k} onClick={() => sci(k)} style={{ "--c": M[k].color }}><span className="disp">{M[k].n}</span>{M[k].label}</button>
          ))}
          <button onClick={() => sci("synthese")} style={{ "--c": MKT }}><span className="disp">S</span>Synthese</button>
        </nav>

        <main className="vck-main">
          <div className="vck-hero">
            <div className="vck-hero-l">
              <span className="vck-eyebrow">Triangulation de la valeur intrinseque</span>
              <div className="vck-hero-big mono">{f2(weighted)}<i>MAD / action (ponderee)</i></div>
              <div className="vck-hero-cmp">
                <div><span>Cours de bourse</span><b className="mono">{f0(price)} MAD</b></div>
                <div className="vck-hero-arrow" style={{ color: upside >= 0 ? POS : NEG }}>{upside >= 0 ? "\u25B2" : "\u25BC"}</div>
                <div><span>Potentiel</span><b className="mono" style={{ color: upside >= 0 ? POS : NEG }}>{fPct(upside)}</b></div>
              </div>
              <div className="vck-verdict" style={{ borderColor: (upside >= 0 ? POS : NEG) + "55", color: upside >= 0 ? POS : NEG }}>{verdict}</div>
            </div>
            <div className="vck-hero-r"><FootballField bars={ffBars} price={price} weighted={weighted} /></div>
          </div>

          <div className="vck-weights no-print">
            <div className="vck-weights-h">
              <span className="vck-eyebrow">Ponderation des methodes</span>
              <div className="vck-weights-act">
                <button onClick={() => setWeights({ dcf: 20, comps: 20, anr: 20, ddm: 20, eva: 20 })}>Egales</button>
                <button onClick={() => setWeights({ dcf: 30, comps: 25, anr: 10, ddm: 15, eva: 20 })}>Defaut</button>
              </div>
            </div>
            <div className="vck-weights-grid">
              {KEYS.map((k) => (
                <div className="vck-wcell" key={k}>
                  <div className="vck-wcell-h"><i style={{ background: M[k].color }} /><span className="disp">{M[k].label}</span><b className="mono">{wSum ? Math.round((weights[k] / wSum) * 100) : 0}%</b></div>
                  <Slider value={weights[k]} min={0} max={100} step={1} color={M[k].color} onChange={(v) => setWeights((w) => ({ ...w, [k]: v }))} />
                  <span className="vck-wcell-v mono">{fAct(central[k])}</span>
                </div>
              ))}
            </div>
          </div>

          <Section id="dcf" m={M.dcf} ps={rDcf.perShare} price={price}>
            <div className="vck-grid">
              <div className="vck-panel">
                <h3 className="disp">Hypotheses</h3>
                <div className="vck-table">
                  <div className="vck-tr vck-th"><span>Annee</span><span>FCF previsionnel (MAD)</span></div>
                  {dcf.fcf.map((v, i) => (
                    <div className="vck-tr" key={i}><span className="mono">{dcf.years[i]}</span><input type="number" value={v} onChange={(e) => setFcf(i, num(e.target.value))} /></div>
                  ))}
                </div>
                <div className="vck-rowbtns"><button onClick={addFcf}>+ annee</button><button onClick={delFcf}>- annee</button></div>
                <div className="vck-sl">
                  <div className="vck-sl-h"><span>WACC</span><b className="mono">{f2(dcf.wacc)} %</b></div>
                  <Slider value={dcf.wacc} min={3} max={18} step={0.05} color={M.dcf.color} onChange={(v) => upDcf({ wacc: v })} />
                  <div className="vck-sl-h"><span>Croissance terminale g</span><b className="mono">{f2(dcf.g)} %</b></div>
                  <Slider value={dcf.g} min={0} max={Math.max(0.1, dcf.wacc - 0.1)} step={0.05} color={M.dcf.color} onChange={(v) => upDcf({ g: v })} />
                </div>
                <Field label="Dette nette" value={dcf.netDebt} onChange={(v) => upDcf({ netDebt: v })} suffix="MAD" />
                {!rDcf.valid && <p className="vck-warn">WACC doit etre superieur a g pour la valeur terminale.</p>}
              </div>
              <div className="vck-out">
                <div className="vck-kpis">
                  <Kpi label="Somme FCF actualises" value={fC(rDcf.sumDisc)} />
                  <Kpi label="Valeur terminale" value={fC(rDcf.tv)} />
                  <Kpi label="VT actualisee" value={fC(rDcf.tvDisc)} />
                  <Kpi label="Valeur d'entreprise" value={fC(rDcf.ev)} accent={M.dcf.color} />
                  <Kpi label="Capitaux propres" value={fC(rDcf.equity)} />
                  <Kpi label="Valeur / action" value={fAct(rDcf.perShare)} accent={M.dcf.color} big />
                </div>
                <ChartBox title="Waterfall des flux actualises vers capitaux propres">
                  <BarChart data={dcfWf} margin={{ top: 8, right: 8, left: 0, bottom: 4 }}>
                    <CartesianGrid stroke="#22273340" vertical={false} />
                    <XAxis dataKey="label" tick={{ fontSize: 9, fill: "#8A93A4" }} interval={0} angle={-22} textAnchor="end" height={64} />
                    <YAxis tick={{ fontSize: 10, fill: "#8A93A4" }} tickFormatter={fC} width={56} />
                    <Tooltip content={<WfTip />} cursor={{ fill: "#ffffff08" }} />
                    <Bar dataKey="base" stackId="a" fill="transparent" />
                    <Bar dataKey="bar" stackId="a" radius={[3, 3, 0, 0]}>{dcfWf.map((r, i) => <Cell key={i} fill={wfColor(r.kind, M.dcf.color)} />)}</Bar>
                  </BarChart>
                </ChartBox>
                <div className="vck-chart">
                  <div className="vck-chart-t disp">Matrice de sensibilite · WACC x g <em>(couleur = ecart vs cours)</em></div>
                  <Heatmap cols={dcfGrid.cols} rows={dcfGrid.rows} fn={(c, r) => dcfAt(dcf, shares, c, r)} baseCol={dcf.wacc} baseRow={dcf.g} price={price} color={M.dcf.color} xlab="WACC" ylab="g" />
                </div>
              </div>
            </div>
          </Section>

          <Section id="comps" m={M.comps} ps={rComps.centralMean} price={price}>
            <div className="vck-grid">
              <div className="vck-panel">
                <h3 className="disp">Panier de pairs</h3>
                <div className="vck-table peers">
                  <div className="vck-tr vck-th"><span>Pair</span><span>PER</span><span>EV/EBITDA</span><span>EV/EBIT</span><span>EV/CA</span><span /></div>
                  {comps.peers.map((p, i) => (
                    <div className="vck-tr" key={i}>
                      <input value={p.name} onChange={(e) => setPeer(i, "name", e.target.value)} />
                      <input type="number" value={p.per} onChange={(e) => setPeer(i, "per", num(e.target.value))} />
                      <input type="number" value={p.evEbitda} onChange={(e) => setPeer(i, "evEbitda", num(e.target.value))} />
                      <input type="number" value={p.evEbit} onChange={(e) => setPeer(i, "evEbit", num(e.target.value))} />
                      <input type="number" value={p.evSales} onChange={(e) => setPeer(i, "evSales", num(e.target.value))} />
                      <button className="vck-x" onClick={() => delPeer(i)}>x</button>
                    </div>
                  ))}
                </div>
                <div className="vck-rowbtns"><button onClick={addPeer}>+ pair</button></div>
                <h3 className="disp" style={{ marginTop: 18 }}>Agregats cibles</h3>
                <Field label="Resultat net" value={comps.target.netIncome} onChange={(v) => setTgt("netIncome", v)} suffix="MAD" />
                <Field label="EBITDA" value={comps.target.ebitda} onChange={(v) => setTgt("ebitda", v)} suffix="MAD" />
                <Field label="EBIT" value={comps.target.ebit} onChange={(v) => setTgt("ebit", v)} suffix="MAD" />
                <Field label="Chiffre d'affaires" value={comps.target.sales} onChange={(v) => setTgt("sales", v)} suffix="MAD" />
                <Field label="Dette nette (multiples EV)" value={comps.netDebt} onChange={(v) => setComps((c) => ({ ...c, netDebt: v }))} suffix="MAD" />
                <p className="vck-note">Multiples EV nets de la dette nette pour le passage a la valeur des fonds propres.</p>
              </div>
              <div className="vck-out">
                <div className="vck-kpis">
                  <Kpi label="Valeur moyenne" value={fAct(rComps.centralMean)} accent={M.comps.color} big />
                  <Kpi label="Valeur mediane" value={fAct(rComps.centralMedian)} />
                  <Kpi label="Fourchette" value={f0(rComps.low) + " · " + f0(rComps.high)} />
                </div>
                <div className="vck-chart">
                  <div className="vck-chart-t disp">Football field · valeur/action par multiple</div>
                  <div className="vck-pf">
                    {(() => {
                      const vals = rComps.lines.map((l) => l.perMean).concat([price]);
                      const hi = Math.max(...vals) * 1.05, sp = hi || 1;
                      const xp = (v) => Math.max(0, Math.min(100, (v / sp) * 100));
                      return (<>
                        <div className="vck-pf-vline mkt" style={{ left: xp(price) + "%" }}><span>Cours {f0(price)}</span></div>
                        <div className="vck-pf-vline avg" style={{ left: xp(rComps.centralMean) + "%", borderColor: M.comps.color }}><span style={{ color: M.comps.color }}>Moy. {f0(rComps.centralMean)}</span></div>
                        {rComps.lines.map((l) => (
                          <div className="vck-pf-row" key={l.key}>
                            <div className="vck-pf-lab disp">{l.label}<b className="mono">{f0(l.perMean)}</b></div>
                            <div className="vck-pf-track"><div className="vck-pf-bar" style={{ width: xp(l.perMean) + "%", background: "linear-gradient(90deg," + M.comps.color + "33," + M.comps.color + "aa)" }} /></div>
                          </div>
                        ))}
                      </>);
                    })()}
                  </div>
                </div>
                <ChartBox title="Comparaison des multiples par pair">
                  <BarChart data={multiplesData} margin={{ top: 8, right: 8, left: 0, bottom: 4 }}>
                    <CartesianGrid stroke="#22273340" vertical={false} />
                    <XAxis dataKey="multiple" tick={{ fontSize: 10, fill: "#8A93A4" }} />
                    <YAxis tick={{ fontSize: 10, fill: "#8A93A4" }} width={36} />
                    <Tooltip content={<ChartTip />} cursor={{ fill: "#ffffff08" }} />
                    {comps.peers.map((p, i) => (<Bar key={i} dataKey={p.name} fill={["#A78BFA", "#7C5CFC", "#C4B5FD", "#6D28D9"][i % 4]} radius={[3, 3, 0, 0]} />))}
                  </BarChart>
                </ChartBox>
              </div>
            </div>
          </Section>

          <Section id="anr" m={M.anr} ps={rAnr.perShareGlobal} price={price}>
            <div className="vck-grid">
              <div className="vck-panel">
                <h3 className="disp">Bilan & retraitements</h3>
                <Field label="Capitaux propres comptables" value={anr.bookEquity} onChange={(v) => setAnr((p) => ({ ...p, bookEquity: v }))} suffix="MAD" />
                <div className="vck-table" style={{ marginTop: 10 }}>
                  <div className="vck-tr vck-th"><span>Retraitement (juste valeur)</span><span>Montant (MAD)</span></div>
                  {anr.adjustments.map((a, i) => (
                    <div className="vck-tr" key={i}><span className="vck-adj">{a.label}</span><input type="number" value={a.amount} onChange={(e) => setAdj(i, num(e.target.value))} /></div>
                  ))}
                </div>
                <Field label="Goodwill (rente de superprofit)" value={anr.goodwill} onChange={(v) => setAnr((p) => ({ ...p, goodwill: v }))} suffix="MAD" />
                <p className="vck-note">ANR = capitaux propres comptables + retraitements en juste valeur. Valeur globale = ANR + Goodwill.</p>
              </div>
              <div className="vck-out">
                <div className="vck-kpis">
                  <Kpi label="ANR (ANCC)" value={fC(rAnr.anr)} />
                  <Kpi label="Valeur globale" value={fC(rAnr.global)} accent={M.anr.color} />
                  <Kpi label="Valeur/action ANR" value={fAct(rAnr.perShareANR)} />
                  <Kpi label="Valeur/action globale" value={fAct(rAnr.perShareGlobal)} accent={M.anr.color} big />
                </div>
                <ChartBox title="Waterfall · valeur comptable vers ANR vers valeur globale" h={300}>
                  <BarChart data={anrWf} margin={{ top: 8, right: 8, left: 0, bottom: 4 }}>
                    <CartesianGrid stroke="#22273340" vertical={false} />
                    <XAxis dataKey="label" tick={{ fontSize: 9, fill: "#8A93A4" }} interval={0} angle={-22} textAnchor="end" height={70} />
                    <YAxis tick={{ fontSize: 10, fill: "#8A93A4" }} tickFormatter={fC} width={56} />
                    <Tooltip content={<WfTip />} cursor={{ fill: "#ffffff08" }} />
                    <Bar dataKey="base" stackId="a" fill="transparent" />
                    <Bar dataKey="bar" stackId="a" radius={[3, 3, 0, 0]}>{anrWf.map((r, i) => <Cell key={i} fill={wfColor(r.kind, M.anr.color)} />)}</Bar>
                  </BarChart>
                </ChartBox>
              </div>
            </div>
          </Section>

          <Section id="ddm" m={M.ddm} ps={rDdm.value} price={price}>
            <div className="vck-grid">
              <div className="vck-panel">
                <h3 className="disp">Historique des dividendes</h3>
                <div className="vck-table">
                  <div className="vck-tr vck-th"><span>Annee</span><span>DPA ajuste (MAD)</span></div>
                  {ddm.history.map((h, i) => (
                    <div className="vck-tr" key={i}><span className="mono">{h.year}</span><input type="number" value={h.dps} onChange={(e) => setDdm((d) => ({ ...d, history: d.history.map((x, j) => (j === i ? { ...x, dps: num(e.target.value) } : x)) }))} /></div>
                  ))}
                </div>
                <Field label="Dernier dividende D0" value={ddm.d0} onChange={(v) => setDdm((d) => ({ ...d, d0: v }))} suffix="MAD" />
                <div className="vck-sl">
                  <div className="vck-sl-h"><span>Cout des fonds propres Ke</span><b className="mono">{f2(ddm.ke)} %</b></div>
                  <Slider value={ddm.ke} min={3} max={20} step={0.05} color={M.ddm.color} onChange={(v) => setDdm((d) => ({ ...d, ke: v }))} />
                  <div className="vck-sl-h"><span>Croissance perpetuelle g</span><b className="mono">{f2(ddm.g)} %</b></div>
                  <Slider value={ddm.g} min={0} max={Math.max(0.1, ddm.ke - 0.1)} step={0.05} color={M.ddm.color} onChange={(v) => setDdm((d) => ({ ...d, g: v }))} />
                </div>
                {!rDdm.valid && <p className="vck-warn">Ke doit etre superieur a g (modele de Gordon).</p>}
              </div>
              <div className="vck-out">
                <div className="vck-kpis">
                  <Kpi label="D1 = D0 x (1+g)" value={fAct(rDdm.d1)} />
                  <Kpi label="Ecart Ke - g" value={f2(ddm.ke - ddm.g) + " pp"} />
                  <Kpi label="Valeur / action" value={fAct(rDdm.value)} accent={M.ddm.color} big />
                </div>
                <ChartBox title="Historique & projection des dividendes">
                  <ComposedChart data={dividendData} margin={{ top: 8, right: 8, left: 0, bottom: 4 }}>
                    <CartesianGrid stroke="#22273340" vertical={false} />
                    <XAxis dataKey="year" tick={{ fontSize: 10, fill: "#8A93A4" }} />
                    <YAxis tick={{ fontSize: 10, fill: "#8A93A4" }} width={36} />
                    <Tooltip content={<ChartTip />} cursor={{ fill: "#ffffff08" }} />
                    <Bar dataKey="hist" name="DPA historique" fill={M.ddm.color} radius={[3, 3, 0, 0]} barSize={26} />
                    <Line dataKey="proj" name="DPA projete" stroke={M.ddm.color} strokeWidth={2} strokeDasharray="5 4" dot={{ r: 3 }} connectNulls />
                  </ComposedChart>
                </ChartBox>
                <div className="vck-chart">
                  <div className="vck-chart-t disp">Matrice de sensibilite · Ke x g <em>(couleur = ecart vs cours)</em></div>
                  <Heatmap cols={ddmGrid.cols} rows={ddmGrid.rows} fn={(c, r) => ddmAt(ddm, c, r)} baseCol={ddm.ke} baseRow={ddm.g} price={price} color={M.ddm.color} xlab="Ke" ylab="g" />
                </div>
              </div>
            </div>
          </Section>

          <Section id="eva" m={M.eva} ps={rEva.perShare} price={price}>
            <div className="vck-grid">
              <div className="vck-panel">
                <h3 className="disp">Serie de creation de valeur</h3>
                <div className="vck-table eva">
                  <div className="vck-tr vck-th"><span>Annee</span><span>NOPAT (MAD)</span><span>Capital investi</span></div>
                  {eva.series.map((s, i) => (
                    <div className="vck-tr" key={i}><span className="mono">{s.year}</span>
                      <input type="number" value={s.nopat} onChange={(e) => setEvaRow(i, "nopat", num(e.target.value))} />
                      <input type="number" value={s.ci} onChange={(e) => setEvaRow(i, "ci", num(e.target.value))} /></div>
                  ))}
                </div>
                <div className="vck-sl">
                  <div className="vck-sl-h"><span>WACC</span><b className="mono">{f2(eva.wacc)} %</b></div>
                  <Slider value={eva.wacc} min={3} max={18} step={0.05} color={M.eva.color} onChange={(v) => setEva((e) => ({ ...e, wacc: v }))} />
                  <div className="vck-sl-h"><span>Croissance g</span><b className="mono">{f2(eva.g)} %</b></div>
                  <Slider value={eva.g} min={0} max={Math.max(0.1, eva.wacc - 0.1)} step={0.05} color={M.eva.color} onChange={(v) => setEva((e) => ({ ...e, g: v }))} />
                </div>
                <Field label="Dette nette (VE vers fonds propres)" value={eva.netDebt} onChange={(v) => setEva((e) => ({ ...e, netDebt: v }))} suffix="MAD" />
                <h3 className="disp" style={{ marginTop: 16 }}>MVA de marche</h3>
                <Field label="Capitalisation boursiere" value={eva.marketCap} onChange={(v) => setEva((e) => ({ ...e, marketCap: v }))} suffix="MAD" />
                <Field label="Dette nette de marche" value={eva.mvaNetDebt} onChange={(v) => setEva((e) => ({ ...e, mvaNetDebt: v }))} suffix="MAD" />
                <Field label="Capital investi" value={eva.capitalInvested} onChange={(v) => setEva((e) => ({ ...e, capitalInvested: v }))} suffix="MAD" />
                <p className="vck-note">EVA = NOPAT - (capital investi x WACC). MVA = valeur de marche - capital investi.</p>
              </div>
              <div className="vck-out">
                <div className="vck-kpis">
                  <Kpi label="Somme PV(EVA)" value={fC(rEva.pvEva)} />
                  <Kpi label="VE intrinseque" value={fC(rEva.ev)} accent={M.eva.color} />
                  <Kpi label="MVA intrinseque" value={fC(rEva.mvaIntrinsic)} />
                  <Kpi label="MVA de marche" value={fC(rEva.mvaMarket)} />
                  <Kpi label="Valeur / action" value={fAct(rEva.perShare)} accent={M.eva.color} big />
                </div>
                <ChartBox title="Evolution de l'EVA">
                  <BarChart data={evaData} margin={{ top: 8, right: 8, left: 0, bottom: 4 }}>
                    <CartesianGrid stroke="#22273340" vertical={false} />
                    <XAxis dataKey="year" tick={{ fontSize: 10, fill: "#8A93A4" }} />
                    <YAxis tick={{ fontSize: 10, fill: "#8A93A4" }} tickFormatter={fC} width={56} />
                    <Tooltip content={<ChartTip />} cursor={{ fill: "#ffffff08" }} />
                    <ReferenceLine y={0} stroke="#3A4150" />
                    <Bar dataKey="eva" name="EVA" radius={[3, 3, 0, 0]}>{evaData.map((d, i) => <Cell key={i} fill={d.eva >= 0 ? M.eva.color : NEG} />)}</Bar>
                  </BarChart>
                </ChartBox>
                <ChartBox title="Bridge MVA · capital investi vers valeur intrinseque">
                  <BarChart data={mvaWf} margin={{ top: 8, right: 8, left: 0, bottom: 4 }}>
                    <CartesianGrid stroke="#22273340" vertical={false} />
                    <XAxis dataKey="label" tick={{ fontSize: 9, fill: "#8A93A4" }} interval={0} angle={-18} textAnchor="end" height={56} />
                    <YAxis tick={{ fontSize: 10, fill: "#8A93A4" }} tickFormatter={fC} width={56} />
                    <Tooltip content={<WfTip />} cursor={{ fill: "#ffffff08" }} />
                    <ReferenceLine y={rEva.marketEV} stroke={MKT} strokeDasharray="4 4" />
                    <Bar dataKey="base" stackId="a" fill="transparent" />
                    <Bar dataKey="bar" stackId="a" radius={[3, 3, 0, 0]}>{mvaWf.map((r, i) => <Cell key={i} fill={wfColor(r.kind, M.eva.color)} />)}</Bar>
                  </BarChart>
                </ChartBox>
              </div>
            </div>
          </Section>

          <section id="synthese" className="vck-sec">
            <header className="vck-sec-h">
              <span className="vck-sec-n disp" style={{ color: MKT }}>S</span>
              <div className="vck-sec-t"><h2 className="disp">Synthese</h2><p>Comparaison des 5 valorisations vs cours de bourse</p></div>
            </header>
            <div className="vck-synth">
              <table className="vck-syntable">
                <thead><tr><th>Methode</th><th>Valeur / action</th><th>Fourchette</th><th>Ecart vs cours</th><th>Ponderation</th><th>Contribution</th></tr></thead>
                <tbody>
                  {KEYS.map((k) => {
                    const v = central[k], up = Number.isFinite(v) ? (v - price) / price : NaN;
                    const wn = wSum ? weights[k] / wSum : 0;
                    return (
                      <tr key={k}>
                        <td><i className="vck-dot" style={{ background: M[k].color }} />{M[k].label}</td>
                        <td className="mono" style={{ color: M[k].color }}>{fAct(v)}</td>
                        <td className="mono dim">{f0(ranges[k].low)} · {f0(ranges[k].high)}</td>
                        <td className="mono" style={{ color: up >= 0 ? POS : NEG }}>{fPct(up)}</td>
                        <td className="mono">{Math.round(wn * 100)} %</td>
                        <td className="mono">{fAct(Number.isFinite(v) ? v * wn : NaN)}</td>
                      </tr>
                    );
                  })}
                  <tr className="vck-syn-tot"><td className="disp">Valeur ponderee</td><td className="mono" style={{ color: MKT }}>{fAct(weighted)}</td><td /><td className="mono" style={{ color: upside >= 0 ? POS : NEG }}>{fPct(upside)}</td><td className="mono">100 %</td><td /></tr>
                  <tr className="vck-syn-mkt"><td className="disp">Cours de bourse</td><td className="mono">{fAct(price)}</td><td colSpan={4} /></tr>
                </tbody>
              </table>
              <div className="vck-synth-ff"><div className="vck-chart-t disp">Football field global</div><FootballField bars={ffBars} price={price} weighted={weighted} /></div>
            </div>
            <p className="vck-foot">Donnees par defaut : HPS (Bourse de Casablanca), exercice 2025. Conventions : multiples EV nets de la dette nette ; actualisation standard des EVA au taux WACC. Recalcul integral en direct sur toute modification d'hypothese.</p>
          </section>
        </main>
      </div>
    </div>
  );
}

const ICON = {
  xls: (<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="3" width="18" height="18" rx="2" /><path d="M8 8l8 8M16 8l-8 8" /></svg>),
  pdf: (<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z" /><path d="M14 3v6h6" /></svg>),
};

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

.vck{
  --bg:#0D0F14; --panel:#14171F; --panel2:#191D27; --line:#252A36; --line2:#2E3442;
  --ink:#E9ECF3; --dim:#99A2B2; --faint:#5C6473;
  font-family:'Inter',system-ui,sans-serif; color:var(--ink); background:var(--bg);
  min-height:100vh; -webkit-font-smoothing:antialiased; letter-spacing:-0.01em;
}
.vck *{box-sizing:border-box;}
.vck .disp{font-family:'Space Grotesk','Inter',sans-serif;}
.vck .mono{font-family:'IBM Plex Mono',ui-monospace,monospace; font-variant-numeric:tabular-nums;}
.vck .dim{color:var(--dim);}
.vck h2,.vck h3{margin:0;}
.vck p{margin:0;}
.vck input{font-family:'IBM Plex Mono',monospace; font-variant-numeric:tabular-nums;}
.vck input[type=number]{-moz-appearance:textfield;}
.vck input[type=number]::-webkit-inner-spin-button,.vck input[type=number]::-webkit-outer-spin-button{-webkit-appearance:none;margin:0;}
.vck button{font-family:'Inter',sans-serif; cursor:pointer;}

/* header */
.vck-top{position:sticky;top:0;z-index:40;display:flex;align-items:center;justify-content:space-between;gap:16px;
  padding:12px 22px;background:rgba(13,15,20,.86);backdrop-filter:blur(12px);border-bottom:1px solid var(--line);}
.vck-top-l{display:flex;align-items:center;gap:16px;min-width:0;}
.vck-logo{font-weight:700;font-size:18px;letter-spacing:.06em;color:var(--ink);white-space:nowrap;}
.vck-logo span{color:#38BDF8;margin:0 1px;}
.vck-co{background:transparent;border:1px solid transparent;color:var(--ink);font-family:'Space Grotesk',sans-serif;font-size:15px;font-weight:500;
  padding:6px 10px;border-radius:8px;min-width:240px;width:34vw;max-width:420px;}
.vck-co:hover,.vck-co:focus{background:var(--panel);border-color:var(--line);outline:none;}
.vck-top-r{display:flex;align-items:center;gap:10px;flex-shrink:0;}
.vck-mini{display:flex;align-items:center;gap:6px;background:var(--panel);border:1px solid var(--line);border-radius:9px;padding:5px 9px;}
.vck-mini span{font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:var(--faint);}
.vck-mini input{background:transparent;border:none;color:var(--ink);width:80px;font-size:13px;outline:none;text-align:right;}
.vck-mini em{font-style:normal;font-size:10px;color:var(--faint);}
.vck-btn{display:inline-flex;align-items:center;gap:6px;background:var(--panel2);color:var(--ink);border:1px solid var(--line2);
  border-radius:9px;padding:7px 13px;font-size:13px;font-weight:500;transition:.15s;}
.vck-btn:hover{background:#222838;border-color:#3a4252;}
.vck-btn svg{opacity:.8;}

/* layout */
.vck-wrap{display:grid;grid-template-columns:188px 1fr;max-width:1320px;margin:0 auto;}
.vck-nav{position:sticky;top:61px;align-self:start;height:calc(100vh - 61px);padding:24px 14px;display:flex;flex-direction:column;gap:3px;border-right:1px solid var(--line);}
.vck-nav button{display:flex;align-items:center;gap:11px;background:transparent;border:none;color:var(--dim);text-align:left;
  padding:10px 12px;border-radius:9px;font-size:13.5px;font-weight:500;transition:.15s;}
.vck-nav button span{font-size:11px;font-weight:600;color:var(--faint);width:16px;text-align:center;}
.vck-nav button:hover{background:var(--panel);color:var(--ink);box-shadow:inset 2px 0 0 var(--c);}
.vck-nav button:hover span{color:var(--c);}
.vck-main{padding:26px 30px 80px;min-width:0;}

/* hero */
.vck-eyebrow{font-size:10.5px;text-transform:uppercase;letter-spacing:.16em;color:var(--faint);font-weight:600;}
.vck-hero{display:grid;grid-template-columns:300px 1fr;gap:30px;background:linear-gradient(135deg,#14171F,#101319);
  border:1px solid var(--line);border-radius:18px;padding:26px 28px;margin-bottom:14px;box-shadow:0 1px 0 rgba(255,255,255,.04) inset, 0 24px 60px -40px #000;}
.vck-hero-l{display:flex;flex-direction:column;gap:14px;}
.vck-hero-big{font-size:46px;font-weight:600;line-height:1;letter-spacing:-.02em;color:var(--ink);}
.vck-hero-big i{display:block;font-style:normal;font-size:11px;color:var(--faint);letter-spacing:.05em;margin-top:9px;}
.vck-hero-cmp{display:flex;align-items:center;gap:16px;padding:14px 0;border-top:1px solid var(--line);border-bottom:1px solid var(--line);}
.vck-hero-cmp>div{display:flex;flex-direction:column;gap:4px;}
.vck-hero-cmp span{font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:var(--faint);}
.vck-hero-cmp b{font-size:18px;}
.vck-hero-arrow{font-size:15px;align-self:center;}
.vck-verdict{align-self:flex-start;font-size:12.5px;font-weight:600;padding:7px 14px;border:1px solid;border-radius:999px;}
.vck-hero-r{min-width:0;align-self:center;}

/* football field */
.vck-ff{position:relative;--lab:172px;}
.vck-ff-top{position:relative;height:26px;margin-left:var(--lab);}
.vck-ff-mark{position:absolute;top:0;transform:translateX(-50%);white-space:nowrap;}
.vck-ff-pin{font-size:10px;font-weight:600;padding:2px 7px;border-radius:5px;font-family:'IBM Plex Mono',monospace;}
.vck-ff-pin.mkt{background:rgba(237,234,224,.14);color:#EDEAE0;border:1px solid rgba(237,234,224,.35);}
.vck-ff-pin.wtd{background:rgba(56,189,248,.12);color:#7FD7FB;border:1px solid rgba(127,215,251,.4);}
.vck-ff-plot{position:relative;}
.vck-ff-layer{position:absolute;left:var(--lab);right:0;top:0;bottom:0;pointer-events:none;z-index:3;}
.vck-ff-grid{position:absolute;inset:0;}
.vck-ff-grid i{position:absolute;top:0;bottom:0;width:1px;background:rgba(255,255,255,.04);}
.vck-ff-vline{position:absolute;top:-2px;bottom:-2px;width:0;}
.vck-ff-vline.mkt{border-left:1.5px dashed rgba(237,234,224,.6);}
.vck-ff-vline.wtd{border-left:1.5px solid rgba(127,215,251,.85);}
.vck-ff-row{display:grid;grid-template-columns:var(--lab) 1fr;align-items:center;height:42px;}
.vck-ff-lab{display:flex;align-items:center;gap:8px;padding-right:14px;font-size:12.5px;}
.vck-ff-lab i{width:8px;height:8px;border-radius:2px;flex-shrink:0;}
.vck-ff-lab span{flex:1;color:var(--ink);font-weight:500;}
.vck-ff-lab b{font-size:12px;color:var(--dim);font-weight:500;}
.vck-ff-track{position:relative;height:100%;}
.vck-ff-bar{position:absolute;top:50%;transform:translateY(-50%);height:14px;border-radius:7px;border:1px solid;}
.vck-ff-tick{position:absolute;top:50%;transform:translate(-50%,-50%);width:2px;height:22px;border-radius:2px;z-index:2;}
.vck-ff-axis{position:relative;height:18px;margin-left:var(--lab);margin-top:4px;}
.vck-ff-axis span{position:absolute;transform:translateX(-50%);font-size:9.5px;color:var(--faint);}

/* weights */
.vck-weights{background:var(--panel);border:1px solid var(--line);border-radius:16px;padding:18px 22px;margin-bottom:26px;}
.vck-weights-h{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;}
.vck-weights-act{display:flex;gap:7px;}
.vck-weights-act button{background:var(--panel2);border:1px solid var(--line2);color:var(--dim);font-size:11.5px;padding:5px 11px;border-radius:7px;transition:.15s;}
.vck-weights-act button:hover{color:var(--ink);border-color:#3a4252;}
.vck-weights-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:18px;}
.vck-wcell-h{display:flex;align-items:center;gap:7px;margin-bottom:9px;}
.vck-wcell-h i{width:9px;height:9px;border-radius:2px;}
.vck-wcell-h span{flex:1;font-size:12.5px;font-weight:500;}
.vck-wcell-h b{font-size:13px;}
.vck-wcell-v{display:block;font-size:11px;color:var(--faint);margin-top:7px;}

/* range */
.vck-range{-webkit-appearance:none;appearance:none;width:100%;height:4px;border-radius:4px;background:var(--line2);outline:none;cursor:pointer;}
.vck-range::-webkit-slider-thumb{-webkit-appearance:none;width:15px;height:15px;border-radius:50%;background:currentColor;border:2px solid var(--bg);box-shadow:0 0 0 1px rgba(255,255,255,.12);}
.vck-range::-moz-range-thumb{width:13px;height:13px;border-radius:50%;background:currentColor;border:2px solid var(--bg);}

/* sections */
.vck-sec{padding-top:30px;margin-top:18px;border-top:1px solid var(--line);scroll-margin-top:74px;}
.vck-sec-h{display:flex;align-items:center;gap:16px;margin-bottom:20px;}
.vck-sec-n{font-size:30px;font-weight:600;line-height:1;width:40px;opacity:.55;}
.vck-sec-t{flex:1;}
.vck-sec-t h2{font-size:21px;font-weight:600;letter-spacing:-.01em;}
.vck-sec-t p{font-size:12.5px;color:var(--dim);margin-top:3px;}
.vck-tag{text-align:right;display:flex;flex-direction:column;gap:3px;}
.vck-tag b{font-size:18px;}
.vck-tag span{font-size:11.5px;}

/* grid */
.vck-grid{display:grid;grid-template-columns:330px 1fr;gap:22px;align-items:start;}
.vck-panel{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px;}
.vck-panel h3{font-size:13px;font-weight:600;color:var(--ink);margin-bottom:12px;letter-spacing:.01em;}
.vck-out{display:flex;flex-direction:column;gap:16px;min-width:0;}

/* tables */
.vck-table{display:flex;flex-direction:column;border:1px solid var(--line);border-radius:10px;overflow:hidden;}
.vck-tr{display:grid;grid-template-columns:1fr 1.4fr;align-items:center;border-bottom:1px solid var(--line);}
.vck-tr:last-child{border-bottom:none;}
.vck-th{background:var(--panel2);}
.vck-th span{font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--faint);padding:8px 10px;font-weight:600;}
.vck-tr>span{padding:0 10px;font-size:12.5px;color:var(--dim);}
.vck-tr input{background:transparent;border:none;border-left:1px solid var(--line);color:var(--ink);padding:9px 10px;font-size:12.5px;width:100%;outline:none;text-align:right;}
.vck-tr input:focus{background:rgba(56,189,248,.06);}
.vck-table.peers .vck-tr{grid-template-columns:1.3fr .8fr 1fr .9fr .8fr 24px;}
.vck-table.peers .vck-th span{padding:8px 6px;font-size:9px;}
.vck-table.peers input{padding:8px 6px;font-size:11.5px;}
.vck-table.peers input:first-child{text-align:left;border-left:none;}
.vck-table.eva .vck-tr{grid-template-columns:.7fr 1.3fr 1.3fr;}
.vck-adj{padding:9px 10px;font-size:11.5px;color:var(--dim);}
.vck-x{background:transparent;border:none;color:var(--faint);font-size:15px;line-height:1;padding:0;}
.vck-x:hover{color:#F87171;}
.vck-rowbtns{display:flex;gap:8px;margin-top:10px;}
.vck-rowbtns button{background:var(--panel2);border:1px solid var(--line2);color:var(--dim);font-size:11.5px;padding:6px 11px;border-radius:7px;transition:.15s;}
.vck-rowbtns button:hover{color:var(--ink);border-color:#3a4252;}

/* sliders block */
.vck-sl{margin:16px 0 6px;}
.vck-sl-h{display:flex;align-items:baseline;justify-content:space-between;margin:14px 0 8px;}
.vck-sl-h:first-child{margin-top:0;}
.vck-sl-h span{font-size:12px;color:var(--dim);}
.vck-sl-h b{font-size:13px;color:var(--ink);}

/* field */
.vck-fld{display:block;margin-top:12px;}
.vck-fld>span{display:block;font-size:11.5px;color:var(--dim);margin-bottom:6px;}
.vck-inwrap{display:flex;align-items:center;background:var(--panel2);border:1px solid var(--line);border-radius:9px;padding:0 11px;}
.vck-inwrap input{flex:1;background:transparent;border:none;color:var(--ink);padding:9px 0;font-size:13px;outline:none;}
.vck-inwrap em{font-style:normal;font-size:10.5px;color:var(--faint);}
.vck-inwrap:focus-within{border-color:#3a4252;}
.vck-warn{font-size:11.5px;color:#FBBF24;margin-top:12px;line-height:1.5;}
.vck-note{font-size:10.5px;color:var(--faint);margin-top:12px;line-height:1.55;}

/* kpis */
.vck-kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:1px;background:var(--line);border:1px solid var(--line);border-radius:12px;overflow:hidden;}
.vck-kpi{background:var(--panel);padding:13px 15px;display:flex;flex-direction:column;gap:6px;}
.vck-kpi.big{grid-column:span 1;background:var(--panel2);}
.vck-kpi-l{font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--faint);font-weight:600;}
.vck-kpi-v{font-size:16px;font-weight:500;color:var(--ink);}
.vck-kpi.big .vck-kpi-v{font-size:19px;}

/* charts */
.vck-chart{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:14px 16px 8px;}
.vck-chart-t{font-size:12px;font-weight:600;color:var(--ink);margin-bottom:10px;}
.vck-chart-t em{font-style:normal;font-weight:400;color:var(--faint);font-size:10.5px;}

/* heatmap */
.vck-heat{overflow-x:auto;}
.vck-heat table{border-collapse:collapse;width:100%;}
.vck-heat th,.vck-heat td{text-align:center;font-size:11px;padding:7px 4px;border:1px solid var(--bg);}
.vck-heat thead th{color:var(--faint);font-weight:500;font-size:10.5px;background:transparent;}
.vck-heat th.row{color:var(--faint);font-weight:500;background:transparent;}
.vck-heat th.base{color:#38BDF8;font-weight:600;}
.vck-heat td{color:var(--ink);font-weight:500;min-width:42px;}
.vck-heat-corner{font-size:9px !important;color:var(--faint);}
.vck-heat-corner i{font-style:normal;opacity:.5;margin:0 2px;}

/* comparables football */
.vck-pf{position:relative;padding-top:22px;}
.vck-pf-vline{position:absolute;top:0;bottom:6px;width:0;z-index:3;}
.vck-pf-vline.mkt{border-left:1.5px dashed rgba(237,234,224,.55);}
.vck-pf-vline.avg{border-left:1.5px solid;}
.vck-pf-vline span{position:absolute;top:-2px;left:5px;font-size:9.5px;font-family:'IBM Plex Mono',monospace;color:#EDEAE0;white-space:nowrap;}
.vck-pf-vline.avg span{left:5px;}
.vck-pf-row{display:grid;grid-template-columns:118px 1fr;align-items:center;height:34px;}
.vck-pf-lab{display:flex;align-items:center;gap:8px;font-size:12px;padding-right:12px;}
.vck-pf-lab b{font-size:11px;color:var(--dim);font-weight:500;margin-left:auto;}
.vck-pf-track{position:relative;height:14px;background:rgba(255,255,255,.03);border-radius:7px;}
.vck-pf-bar{position:absolute;left:0;top:0;height:100%;border-radius:7px;min-width:2px;}

/* chart tooltip */
.vck-cttip{background:#0B0D12;border:1px solid var(--line2);border-radius:9px;padding:9px 12px;box-shadow:0 10px 30px -10px #000;}
.vck-cttip-l{font-size:11px;color:var(--dim);margin-bottom:5px;font-weight:500;}
.vck-cttip div{font-size:12px;color:var(--ink);display:flex;align-items:center;gap:7px;}
.vck-cttip i{width:8px;height:8px;border-radius:2px;display:inline-block;}

/* synthese */
.vck-synth{display:grid;grid-template-columns:1fr;gap:22px;}
.vck-syntable{width:100%;border-collapse:collapse;background:var(--panel);border:1px solid var(--line);border-radius:14px;overflow:hidden;}
.vck-syntable th{text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--faint);font-weight:600;padding:12px 16px;border-bottom:1px solid var(--line);background:var(--panel2);}
.vck-syntable th:not(:first-child){text-align:right;}
.vck-syntable td{padding:13px 16px;font-size:13px;border-bottom:1px solid var(--line);}
.vck-syntable td:not(:first-child){text-align:right;}
.vck-syntable td:first-child{display:flex;align-items:center;gap:9px;color:var(--ink);font-weight:500;}
.vck-dot{width:9px;height:9px;border-radius:2px;}
.vck-syn-tot td{border-top:2px solid var(--line2);border-bottom:none;font-weight:600;background:rgba(255,255,255,.02);padding-top:15px;padding-bottom:15px;}
.vck-syn-mkt td{border-bottom:none;color:var(--dim);}
.vck-syn-mkt td:first-child{color:var(--dim);}
.vck-synth-ff{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px 20px;}
.vck-foot{font-size:10.5px;color:var(--faint);line-height:1.6;margin-top:22px;max-width:760px;}

/* responsive */
@media (max-width:1080px){
  .vck-wrap{grid-template-columns:1fr;}
  .vck-nav{display:none;}
  .vck-main{padding:22px 18px 60px;}
  .vck-hero{grid-template-columns:1fr;gap:22px;}
  .vck-weights-grid{grid-template-columns:repeat(2,1fr);}
  .vck-grid{grid-template-columns:1fr;}
}
@media (max-width:680px){
  .vck-top{flex-wrap:wrap;gap:10px;padding:10px 14px;}
  .vck-co{width:100%;min-width:0;}
  .vck-top-r{width:100%;flex-wrap:wrap;}
  .vck-hero-big{font-size:38px;}
  .vck-kpis{grid-template-columns:repeat(2,1fr);}
  .vck-weights-grid{grid-template-columns:1fr;}
  .vck-ff{--lab:120px;}
}

/* print */
@media print{
  .vck{background:#fff;color:#111;}
  .no-print{display:none !important;}
  .vck-wrap{grid-template-columns:1fr;max-width:100%;}
  .vck-main{padding:0;}
  .vck-hero,.vck-panel,.vck-chart,.vck-weights,.vck-syntable,.vck-synth-ff{background:#fff;border-color:#ccc;box-shadow:none;}
  .vck-sec{break-inside:avoid;border-top-color:#ddd;}
  .vck-hero-big,.vck-sec-t h2,.vck-kpi-v,.vck-syntable td,.vck-ff-lab span{color:#111;}
  .vck-kpi.big,.vck-th,.vck-syntable th{background:#f4f4f4;}
  *{-webkit-print-color-adjust:exact;print-color-adjust:exact;}
}
`;
