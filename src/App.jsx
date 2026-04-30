import { useState, useCallback, useMemo, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { DEFAULT_Y, DEFAULT_FEATURES, META_SKIP, compShort } from './constants.jsx'
import { toNum, isNumCol, standardize, removeOutliers, ols, ridge, buildFixedEffects, computeRelativeMultiple } from './math.js'
import ResultsDashboard from './components/ResultsDashboard.jsx'

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0c0f14;--bg1:#111520;--bg2:#161b26;--bg3:#1c2333;
  --border:rgba(255,255,255,0.07);--border-h:rgba(255,255,255,0.13);
  --blue:#3b82f6;--blue-d:rgba(59,130,246,0.15);
  --green:#10b981;--green-d:rgba(16,185,129,0.12);
  --red:#f43f5e;--red-d:rgba(244,63,94,0.12);
  --amber:#f59e0b;--amber-d:rgba(245,158,11,0.12);
  --text:#f1f5f9;--text2:#94a3b8;--text3:#475569;
  --font:'Sora',sans-serif;--mono:'JetBrains Mono',monospace;
}
body{font-family:var(--font);background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden}
::-webkit-scrollbar{width:4px;height:4px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:var(--bg3);border-radius:2px}
@keyframes spin{to{transform:rotate(360deg)}}
@keyframes fadeUp{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:translateY(0)}}
.fade-up{animation:fadeUp 0.4s ease forwards}
.d1{animation-delay:0.05s;opacity:0}.d2{animation-delay:0.1s;opacity:0}.d3{animation-delay:0.15s;opacity:0}.d4{animation-delay:0.2s;opacity:0}.d5{animation-delay:0.25s;opacity:0}
.chip{display:inline-flex;align-items:center;gap:6px;padding:7px 13px;border-radius:20px;cursor:pointer;font-size:12px;font-weight:500;border:1.5px solid var(--border);background:var(--bg2);color:var(--text3);transition:all 0.15s ease;user-select:none;font-family:var(--font)}
.chip:hover{border-color:var(--border-h);color:var(--text2);transform:translateY(-1px)}
.chip.on{background:var(--blue-d);border-color:var(--blue);color:var(--blue)}
.chip.on-amber{background:var(--amber-d);border-color:var(--amber);color:var(--amber)}
.chip.on-green{background:var(--green-d);border-color:var(--green);color:var(--green)}
.mcard{padding:16px;border-radius:12px;cursor:pointer;border:1.5px solid var(--border);background:var(--bg2);transition:all 0.2s ease;flex:1;min-width:150px}
.mcard:hover{border-color:var(--border-h);transform:translateY(-2px)}
.mcard.sel{border-color:var(--blue);background:var(--blue-d)}
.mcard .mname{font-size:13px;font-weight:600;margin:8px 0 4px;color:var(--text2)}
.mcard.sel .mname{color:var(--blue)}
.mcard .mdesc{font-size:11px;color:var(--text3);line-height:1.5}
.radio{width:15px;height:15px;border-radius:50%;border:2px solid var(--text3);display:flex;align-items:center;justify-content:center;transition:all 0.15s}
.mcard.sel .radio{border-color:var(--blue)}
.mcard.sel .radio::after{content:'';width:6px;height:6px;border-radius:50%;background:var(--blue);display:block}
.run-btn{display:flex;align-items:center;justify-content:center;gap:10px;padding:15px 36px;border-radius:12px;border:none;cursor:pointer;font-family:var(--font);font-size:14px;font-weight:600;background:var(--blue);color:white;transition:all 0.2s ease}
.run-btn:hover{background:#2563eb;transform:translateY(-2px);box-shadow:0 8px 24px rgba(59,130,246,0.3)}
.run-btn:active{transform:translateY(0)}
.run-btn:disabled{opacity:0.4;cursor:not-allowed;transform:none;box-shadow:none}
.spinner{width:16px;height:16px;border:2px solid rgba(255,255,255,0.3);border-top-color:white;border-radius:50%;animation:spin 0.7s linear infinite}
.glass{background:var(--bg1);border:1px solid var(--border);border-radius:16px;padding:22px}
.sel{width:100%;padding:11px 14px;border-radius:10px;background:var(--bg2);border:1.5px solid var(--border);color:var(--text);font-family:var(--font);font-size:13px;outline:none;transition:border-color 0.15s;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%2394a3b8' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 12px center;padding-right:34px;cursor:pointer}
.sel:focus{border-color:var(--blue)}
.slabel{font-size:10px;font-weight:600;letter-spacing:0.12em;text-transform:uppercase;color:var(--text3);margin-bottom:8px}
.warn{padding:10px 14px;border-radius:10px;background:var(--amber-d);border:1px solid rgba(245,158,11,0.25);font-size:12px;color:var(--amber);margin-top:12px}
.tab-bar{display:flex;gap:3px;padding:3px;background:var(--bg2);border-radius:11px;overflow-x:auto}
.tab{padding:8px 15px;border-radius:8px;border:none;cursor:pointer;font-family:var(--font);font-size:12px;font-weight:500;background:transparent;color:var(--text3);transition:all 0.15s;white-space:nowrap}
.tab:hover{color:var(--text2)}
.tab.active{background:var(--bg3);color:var(--text);box-shadow:0 1px 4px rgba(0,0,0,0.35)}
.kpi{flex:1;min-width:100px;padding:18px;background:var(--bg1);border-radius:14px;border:1px solid var(--border)}
.kpi-val{font-size:22px;font-weight:700;letter-spacing:-0.02em}
.kpi-label{font-size:10px;color:var(--text3);letter-spacing:0.1em;text-transform:uppercase;margin-top:4px}
.kpi-sub{font-size:11px;margin-top:3px}
.fs-overlay{position:fixed;inset:0;z-index:9999;background:rgba(8,10,15,0.97);backdrop-filter:blur(16px);display:flex;flex-direction:column;animation:fadeUp 0.2s ease}
.fs-btn{width:34px;height:34px;border-radius:8px;background:var(--bg2);border:1px solid var(--border);color:var(--text2);font-size:16px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all 0.15s;font-family:inherit}
.fs-btn:hover{background:var(--bg3);color:var(--text)}
.expand-btn{padding:6px 10px;border-radius:7px;border:1px solid var(--border);background:transparent;color:var(--text3);font-size:11px;cursor:pointer;font-family:var(--font);transition:all 0.15s;display:flex;align-items:center;gap:5px}
.expand-btn:hover{border-color:var(--border-h);color:var(--text2)}
.num-input{width:100%;padding:8px 10px;border-radius:8px;background:var(--bg2);border:1.5px solid var(--border);color:var(--text);font-family:var(--mono);font-size:12px;outline:none;transition:border-color 0.15s}
.num-input:focus{border-color:var(--blue)}
.num-input.mod{border-color:var(--amber)}
.reset-btn{padding:5px 9px;border-radius:6px;border:1px solid rgba(245,158,11,0.3);background:transparent;color:var(--amber);font-size:10px;cursor:pointer;font-family:var(--font);transition:all 0.15s}
.reset-btn:hover{background:var(--amber-d)}
.insight{padding:14px 16px;border-radius:12px;font-size:12px;color:var(--text2);line-height:1.8;margin-top:12px}
.badge{display:inline-flex;align-items:center;gap:6px;padding:5px 11px;border-radius:8px;font-size:10px;font-weight:600;letter-spacing:0.06em}
`

function StyleInject() {
  useEffect(() => {
    const el = document.createElement('style')
    el.textContent = CSS
    document.head.appendChild(el)
    return () => document.head.removeChild(el)
  }, [])
  return null
}

// ─── Upload ───────────────────────────────────────────────────────────────────
function UploadZone({ onData }) {
  const [drag, setDrag] = useState(false)
  const [loading, setLoading] = useState(false)
  const handle = f => {
    if (!f) return
    setLoading(true)
    const r = new FileReader()
    r.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: false })
      const name = wb.SheetNames.includes('Annualized_Panel') ? 'Annualized_Panel' : wb.SheetNames[0]
      const data = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: null, raw: true })
      setTimeout(() => { onData(data, name); setLoading(false) }, 200)
    }
    r.readAsArrayBuffer(f)
  }
  return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 32 }}>
      <div style={{ maxWidth: 480, width: '100%', textAlign: 'center' }}>
        <div className="fade-up" style={{ marginBottom: 48 }}>
          <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.2em', color: 'var(--blue)', textTransform: 'uppercase', marginBottom: 14 }}>ValuationEngine</div>
          <div style={{ fontSize: 34, fontWeight: 700, lineHeight: 1.2, marginBottom: 12 }}>Regression-Based<br />Valuation Support</div>
          <div style={{ fontSize: 15, color: 'var(--text2)', fontWeight: 300 }}>M&A Advisory · Comparable Analysis</div>
        </div>
        <div className="fade-up d2"
          onDragOver={e => { e.preventDefault(); setDrag(true) }}
          onDragLeave={() => setDrag(false)}
          onDrop={e => { e.preventDefault(); setDrag(false); handle(e.dataTransfer.files[0]) }}
          onClick={() => !loading && document.getElementById('xlsxIn').click()}
          style={{ border: `2px dashed ${drag ? 'var(--blue)' : 'var(--border-h)'}`, borderRadius: 20, padding: '44px 32px', cursor: loading ? 'wait' : 'pointer', background: drag ? 'var(--blue-d)' : 'var(--bg1)', transition: 'all 0.2s ease', transform: drag ? 'scale(1.02)' : 'scale(1)' }}>
          {loading
            ? <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12 }}><div className="spinner" style={{ width: 28, height: 28, border: '3px solid var(--bg3)', borderTopColor: 'var(--blue)' }} /><div style={{ color: 'var(--text2)', fontSize: 14 }}>Reading your data…</div></div>
            : <><div style={{ fontSize: 34, marginBottom: 14 }}>📊</div><div style={{ fontSize: 16, fontWeight: 600, marginBottom: 6 }}>Drop your Excel here</div><div style={{ fontSize: 13, color: 'var(--text2)' }}>or click to browse</div><div style={{ marginTop: 18, fontSize: 11, color: 'var(--text3)' }}>Sheet "Annualized_Panel" · company metrics by year</div></>}
        </div>
        <input id="xlsxIn" type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={e => handle(e.target.files[0])} />
      </div>
    </div>
  )
}

// ─── Setup ────────────────────────────────────────────────────────────────────
function SetupScreen({ rawData, sheetName, onRun, onReset }) {
  const { rows, numCols } = rawData
  const allCompanies = useMemo(() => [...new Set(rows.map(r => r._company).filter(Boolean))], [rows])
  const allYears = [2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025]
  const targetCols = useMemo(() => numCols.filter(c => !META_SKIP.has(c)), [numCols])
  const [target, setTarget] = useState(DEFAULT_Y)
  const featureCols = useMemo(() => numCols.filter(c => !META_SKIP.has(c) && c !== target), [numCols, target])
  const [client, setClient] = useState(allCompanies[0] || '')
  const [selCos, setSelCos] = useState(allCompanies)
  const [yearRange, setYearRange] = useState([2018, 2025])
  const [features, setFeatures] = useState(DEFAULT_FEATURES)
  const [method, setMethod] = useState('ols')
  const [filterOut, setFilterOut] = useState(true)
  const [useRelative, setUseRelative] = useState(false)
  const [running, setRunning] = useState(false)
  const [warning, setWarning] = useState('')

  const validF = features.filter(f => f !== target)
  const filteredRows = useMemo(() => rows.filter(r => selCos.includes(r._company) && r._year >= yearRange[0] && r._year <= yearRange[1]), [rows, selCos, yearRange])
  const obsCount = filteredRows.filter(r => isFinite(r[target]) && validF.every(f => isFinite(r[f]))).length

  const toggleYear = y => setYearRange(([lo, hi]) => {
    const on = y >= lo && y <= hi
    if (on && y === lo && lo < hi) return [lo + 1, hi]
    if (on && y === hi && hi > lo) return [lo, hi - 1]
    if (!on) return [Math.min(lo, y), Math.max(hi, y)]
    return [lo, hi]
  })

  const handleRun = () => {
    setWarning('')
    // Optionally compute relative multiple as target
    const workRows = useRelative ? computeRelativeMultiple(filteredRows, target) : filteredRows
    const effectiveTarget = useRelative ? '_relativeMultiple' : target
    let clean = workRows.filter(r => isFinite(r[effectiveTarget]) && validF.every(f => isFinite(r[f])))
    if (filterOut && clean.length > 4) clean = removeOutliers(clean, effectiveTarget)
    if (clean.length < validF.length + 3) { setWarning(`Only ${clean.length} valid obs. Need ≥ ${validF.length + 3}.`); return }
    setRunning(true)
    setTimeout(() => {
      let res = null, labels = []
      if (method === 'fe' || method === 'fe_ridge') {
        const fe = buildFixedEffects(clean, validF, effectiveTarget)
        const wr = clean.map((r, i) => { const o = {}; validF.forEach((f, j) => { o[f] = fe.XWithin[i][j] }); return o })
        const { Xs, ms, ss } = standardize(wr, validF)
        res = method === 'fe_ridge' ? ridge(Xs, fe.yWithin, 5) : ols(Xs, fe.yWithin)
        labels = ['Intercept (within)', ...validF]
        if (res) { res.companyMeansY = fe.companyMeansY; res.companyMeansX = fe.companyMeansX; res.isFE = true }
      } else {
        const { Xs } = standardize(clean, validF)
        const y = clean.map(r => r[effectiveTarget])
        res = method === 'ridge' ? ridge(Xs, y, 5) : ols(Xs, y)
        labels = ['Intercept', ...validF]
      }
      if (!res) { setWarning('Matrix error — try fewer/less correlated variables.'); setRunning(false); return }
      const { Xs: Xstd } = standardize(clean, validF)
      const oRes = ols(Xstd, clean.map(r => r[effectiveTarget]))
      const chartData = oRes ? clean.map((r, i) => ({
        company: r._company.split(' ')[0], year: r._year,
        actual: clean[i][effectiveTarget],
        actualAbsolute: clean[i][target],
        sectorMedian: clean[i]['_sectorMedian'] ?? null,
        predicted: oRes.yHat[i], residual: oRes.residuals[i],
        index: i, isClient: r._company === client
      })) : []
      const featureImp = validF.map((f, i) => ({ name: f, short: f.length > 30 ? f.slice(0, 30) + '…' : f, tStat: Math.abs(res.tS?.[i + 1] ?? 0), sig: (res.pV?.[i + 1] ?? 1) < 0.05, beta: res.beta[i + 1] ?? 0, pVal: res.pV?.[i + 1] ?? 1 })).sort((a, b) => b.tStat - a.tStat)
      onRun({ result: { ...res, chartData, labels, validF, featureImp }, config: { target, effectiveTarget, useRelative, clientCompany: client, selectedCos: selCos, yearRange, features, method, filterOutliers: filterOut, allRows: rows } })
      setRunning(false)
    }, 80)
  }

  const methods = [
    { val: 'ols', label: 'Pooled OLS', desc: 'Standard across all companies' },
    { val: 'ridge', label: 'Ridge', desc: 'Better with correlated vars' },
    { val: 'fe', label: 'Fixed Effects', desc: 'Recommended for panel data' },
    { val: 'fe_ridge', label: 'FE + Ridge', desc: 'Most robust option' },
  ]

  return (
    <div style={{ minHeight: '100vh', display: 'flex', flexDirection: 'column' }}>
      <div style={{ borderBottom: '1px solid var(--border)', padding: '14px 28px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: 'var(--bg1)' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
          <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.15em', color: 'var(--blue)', textTransform: 'uppercase' }}>ValuationEngine</span>
          <span style={{ width: 1, height: 14, background: 'var(--border)', display: 'inline-block' }} />
          <span style={{ fontSize: 12, color: 'var(--text2)' }}>{sheetName} · {allCompanies.length} companies</span>
        </div>
        <button onClick={onReset} style={{ fontSize: 12, color: 'var(--text3)', background: 'none', border: 'none', cursor: 'pointer', fontFamily: 'var(--font)' }}>← Change file</button>
      </div>
      <div style={{ flex: 1, padding: '28px 28px', maxWidth: 1080, margin: '0 auto', width: '100%' }}>
        <div className="fade-up" style={{ marginBottom: 28 }}>
          <div style={{ fontSize: 22, fontWeight: 700, marginBottom: 4 }}>Configure your model</div>
          <div style={{ fontSize: 13, color: 'var(--text2)', fontWeight: 300 }}>Choose your settings — then run the regression to see results</div>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14, marginBottom: 14 }}>
          <div className="glass fade-up d1">
            <div className="slabel">Predict (Y)</div>
            <select className="sel" value={target} onChange={e => { setTarget(e.target.value); setFeatures(f => f.filter(x => x !== e.target.value)) }}>{targetCols.map(c => <option key={c} value={c}>{c}</option>)}</select>
          </div>
          <div className="glass fade-up d1">
            <div className="slabel">★ Your client</div>
            <select className="sel" value={client} onChange={e => setClient(e.target.value)}>{allCompanies.map(c => <option key={c} value={c}>{c}</option>)}</select>
          </div>
        </div>

        <div className="glass fade-up d2" style={{ marginBottom: 14 }}>
          <div className="slabel">Regression method</div>
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
            {methods.map(m => (
              <div key={m.val} className={`mcard ${method === m.val ? 'sel' : ''}`} onClick={() => setMethod(m.val)}>
                <div className="radio" />
                <div className="mname">{m.label}</div>
                <div className="mdesc">{m.desc}</div>
              </div>
            ))}
          </div>
          <label style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: 14, cursor: 'pointer', fontSize: 12, color: filterOut ? 'var(--amber)' : 'var(--text3)' }}>
            <input type="checkbox" checked={filterOut} onChange={e => setFilterOut(e.target.checked)} style={{ accentColor: 'var(--amber)', width: 14, height: 14 }} />
            Filter extreme outliers (IQR × 3) — removes negative multiples like QXO
          </label>
        </div>

        {/* Target mode toggle */}
        <div className="glass fade-up d2" style={{ marginBottom: 14 }}>
          <div className="slabel">Prediction mode</div>
          <div style={{ display: 'flex', gap: 10 }}>
            <div className={`mcard ${!useRelative ? 'sel' : ''}`} onClick={() => setUseRelative(false)} style={{ maxWidth: 300 }}>
              <div className="radio" />
              <div className="mname">Absolute multiple</div>
              <div className="mdesc">Predicts the raw EV/EBITDA multiple directly. Wider prediction range but simpler to interpret.</div>
            </div>
            <div className={`mcard ${useRelative ? 'sel' : ''}`} onClick={() => setUseRelative(true)} style={{ maxWidth: 300 }}>
              <div className="radio" />
              <div className="mname" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>Relative to sector <span style={{ fontSize: 10, padding: '2px 7px', borderRadius: 5, background: 'var(--green-d, rgba(16,185,129,0.12))', color: 'var(--green)', border: '1px solid rgba(16,185,129,0.25)', fontWeight: 500 }}>Recommended</span></div>
              <div className="mdesc">Predicts how much above or below the sector median a company should trade. Removes macro noise (rate cycles, sentiment) and focuses on company-specific premium. Cleaner signal.</div>
            </div>
          </div>
        </div>

        <div className="glass fade-up d3" style={{ marginBottom: 14 }}>
          <div className="slabel">Year range</div>
          <div style={{ display: 'flex', gap: 7, flexWrap: 'wrap' }}>
            {allYears.map(y => { const on = y >= yearRange[0] && y <= yearRange[1]; return <div key={y} className={`chip ${on ? 'on-green' : ''}`} onClick={() => toggleYear(y)}><span style={{ width: 6, height: 6, borderRadius: '50%', background: 'currentColor', opacity: on ? 1 : 0.3 }} />{y}</div> })}
          </div>
        </div>

        <div className="glass fade-up d3" style={{ marginBottom: 14 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
            <div className="slabel" style={{ margin: 0 }}>Comparable companies ({selCos.length}/{allCompanies.length})</div>
            <div style={{ display: 'flex', gap: 10 }}>
              <button onClick={() => setSelCos(allCompanies)} style={{ fontSize: 11, color: 'var(--blue)', background: 'none', border: 'none', cursor: 'pointer', fontFamily: 'var(--font)' }}>All</button>
              <button onClick={() => setSelCos(p => p.includes(client) ? p : [client])} style={{ fontSize: 11, color: 'var(--text3)', background: 'none', border: 'none', cursor: 'pointer', fontFamily: 'var(--font)' }}>Client only</button>
            </div>
          </div>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 7 }}>
            {allCompanies.map(co => { const on = selCos.includes(co); const isC = co === client; return <div key={co} className={`chip ${on ? (isC ? 'on-amber' : 'on') : ''}`} onClick={() => { if (isC) return; setSelCos(p => p.includes(co) ? p.filter(x => x !== co) : [...p, co]) }} style={{ opacity: on ? 1 : 0.45 }}>{isC && <span style={{ fontSize: 9 }}>★</span>}{compShort(co)}</div> })}
          </div>
        </div>

        <div className="glass fade-up d4" style={{ marginBottom: 24 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
            <div className="slabel" style={{ margin: 0 }}>Independent variables X ({validF.length} selected)</div>
            <button onClick={() => setFeatures(DEFAULT_FEATURES)} style={{ fontSize: 11, color: 'var(--blue)', background: 'none', border: 'none', cursor: 'pointer', fontFamily: 'var(--font)' }}>Reset to recommended</button>
          </div>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 7 }}>
            {featureCols.map(f => { const on = features.includes(f); const rec = DEFAULT_FEATURES.includes(f); return <div key={f} className={`chip ${on ? 'on' : ''}`} onClick={() => setFeatures(p => p.includes(f) ? p.filter(x => x !== f) : [...p, f])} style={{ opacity: on ? 1 : 0.4 }}>{rec && <span style={{ fontSize: 9, opacity: 0.6 }}>★</span>}{f}</div> })}
          </div>
          <div style={{ marginTop: 10, fontSize: 11, color: 'var(--text3)', lineHeight: 1.6 }}>★ = recommended · Keep 4–7 variables for best results. The recommended variables (growth, margin, ROIC, scale, cyclicality, theme score) are fundamental drivers that work across different target multiples — they explain <em>why</em> companies trade at a premium or discount regardless of which multiple you predict.</div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: 16, flexWrap: 'wrap' }}>
          <button className="run-btn" onClick={handleRun} disabled={running || obsCount < 5}>
            {running ? <><div className="spinner" />Running…</> : '▶ Run Regression'}
          </button>
          <div style={{ fontSize: 13, color: 'var(--text3)' }}>{validF.length} variables · <span style={{ color: 'var(--text2)' }}>{obsCount} observations</span></div>
        </div>
        {warning && <div className="warn">⚠ {warning}</div>}
      </div>
    </div>
  )
}

export default function App() {
  const [rawData, setRawData] = useState(null)
  const [sheetName, setSheetName] = useState('')
  const [runResult, setRunResult] = useState(null)

  const handleData = useCallback((data, sheet) => {
    setSheetName(sheet)
    const withCalc = data.map(row => ({ ...row, 'EV/Sales': (toNum(row['EV ($mm)']) / toNum(row['Sales ($mm)'])) || null }))
    const numCols = Object.keys(withCalc[0] || {}).filter(k => !META_SKIP.has(k) && k !== 'Company Name' && k !== 'Bucket' && k !== 'Theme Bucket' && isNumCol(withCalc, k))
    const cleaned = withCalc.map(row => { const r = { _company: row['Company Name'] || '', _year: toNum(row['As-of Year']), _bucket: row['Bucket'] || '' }; for (const k of numCols) r[k] = toNum(row[k]); return r })
    setRawData({ rows: cleaned, numCols })
    setRunResult(null)
  }, [])

  return (
    <>
      <StyleInject />
      {!rawData && <UploadZone onData={handleData} />}
      {rawData && !runResult && <SetupScreen rawData={rawData} sheetName={sheetName} onRun={r => setRunResult(r)} onReset={() => { setRawData(null); setRunResult(null) }} />}
      {rawData && runResult && <ResultsDashboard result={runResult.result} config={runResult.config} allRows={rawData.rows} onBack={() => setRunResult(null)} onReset={() => { setRawData(null); setRunResult(null) }} />}
    </>
  )
}
