import { useState, useMemo, useContext, createContext, Component, useRef } from 'react'
import * as XLSX from 'xlsx'
import { C as DARK, compShort } from '../constants.jsx'
import { standardize, applyStd, removeOutliers, buildFixedEffects, ols, ridge, walkForwardCV, avg, computeRelativeMultiple } from '../math.js'
import {
  ComposedChart, Line, Area, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, ReferenceLine, Legend, BarChart, Bar, Cell,
  ScatterChart, Scatter, ZAxis, LabelList
} from 'recharts'


// ─── Error Boundary ────────────────────────────────────────────────────────────
class TabErrorBoundary extends Component {
  constructor(props) { super(props); this.state = { error: null } }
  static getDerivedStateFromError(e) { return { error: e } }
  render() {
    if (this.state.error) {
      const C = DARK
      return (
        <div style={{ padding: 32, textAlign: 'center', background: C.bg1, borderRadius: 16, border: `1px solid ${C.red}44` }}>
          <div style={{ fontSize: 20, marginBottom: 12 }}>⚠️</div>
          <div style={{ color: C.red, fontWeight: 600, marginBottom: 8 }}>Something went wrong in this tab</div>
          <div style={{ color: C.text3, fontSize: 12, fontFamily: 'var(--mono)', maxWidth: 400, margin: '0 auto' }}>{this.state.error?.message}</div>
          <button onClick={() => this.setState({ error: null })} style={{ marginTop: 16, padding: '8px 20px', borderRadius: 8, border: `1px solid ${C.blue}`, background: C.blueDim, color: C.blue, cursor: 'pointer', fontFamily: 'var(--font)', fontSize: 12 }}>Try again</button>
        </div>
      )
    }
    return this.props.children
  }
}

// ─── Light theme palette ──────────────────────────────────────────────────────
const LIGHT = {
  bg: '#FFFFFF', bg1: '#F8F5F2', bg2: '#F0EBE3', bg3: 'rgb(219,204,189)',
  border: 'rgba(96,0,29,0.1)', borderH: 'rgba(96,0,29,0.22)',
  blue: 'rgb(24,78,98)', blueDim: 'rgba(24,78,98,0.1)',
  green: 'rgb(24,78,98)', greenDim: 'rgba(24,78,98,0.08)',
  red: 'rgb(96,0,29)', redDim: 'rgba(96,0,29,0.08)',
  amber: 'rgb(96,0,29)', amberDim: 'rgba(96,0,29,0.06)',
  text: '#1a1a1a', text2: '#4a4a4a', text3: '#888888',
  grid: 'rgba(0,0,0,0.06)',
  accent: 'rgb(195,222,231)', // light blue for highlights
  burgundy: 'rgb(96,0,29)',
}

// Theme context so all components share current palette
const ThemeCtx = createContext(DARK)
const useTheme = () => useContext(ThemeCtx)

// ─── Dynamic axis domain from data ────────────────────────────────────────────
function getAxisDomain(values, padding = 0.1) {
  const finite = values.filter(v => isFinite(v) && v > -999)
  if (!finite.length) return ['auto', 'auto']
  // Use 2nd and 98th percentile to ignore extreme outliers
  const sorted = finite.slice().sort((a, b) => a - b)
  const p2 = sorted[Math.floor(sorted.length * 0.02)] ?? sorted[0]
  const p98 = sorted[Math.floor(sorted.length * 0.98)] ?? sorted[sorted.length - 1]
  const range = p98 - p2
  const lo = Math.floor((p2 - range * padding) * 10) / 10
  const hi = Math.ceil((p98 + range * padding) * 10) / 10
  return [lo, hi]
}

// Detect target type to format correctly throughout
function getTargetMeta(target) {
  if (!target) return { unit: 'x', isDollar: false, suffix: 'x' }
  const t = target.toLowerCase()
  if (t.includes('($mm)') || t.includes('($bn)') || t.includes('invested cap')) return { unit: '$mm', isDollar: true, suffix: 'mm' }
  if (t.includes('price ($)')) return { unit: '$', isDollar: true, suffix: '' }
  if (t.includes('(%)') || t.includes('margin') || (t.includes('growth') && !t.includes('/'))) return { unit: '%', isDollar: false, suffix: '%' }
  return { unit: 'x', isDollar: false, suffix: 'x' }
}

function formatAxisTick(v, meta) {
  if (!isFinite(v)) return ''
  const m = typeof meta === 'string' ? getTargetMeta(meta) : (meta || { unit: 'x', suffix: 'x' })
  if (m.unit === '%') return `${v.toFixed(0)}%`
  if (m.isDollar && m.unit === '$mm') {
    if (Math.abs(v) >= 100000) return `$${(v / 1000).toFixed(0)}B`
    if (Math.abs(v) >= 1000) return `$${(v / 1000).toFixed(1)}B`
    return `$${v.toFixed(0)}M`
  }
  if (m.isDollar) return `$${v.toFixed(0)}`
  if (Math.abs(v) >= 100) return `${v.toFixed(0)}x`
  if (Math.abs(v) >= 10) return `${v.toFixed(1)}x`
  return `${v.toFixed(2)}x`
}

function formatValue(v, meta) {
  if (!isFinite(v)) return '—'
  const m = typeof meta === 'string' ? getTargetMeta(meta) : (meta || { unit: 'x', suffix: 'x' })
  if (m.unit === '%') return `${v.toFixed(1)}%`
  if (m.isDollar && m.unit === '$mm') {
    if (Math.abs(v) >= 1000) return `$${(v / 1000).toFixed(1)}B`
    return `$${v.toFixed(0)}M`
  }
  if (m.isDollar) return `$${v.toFixed(0)}`
  return `${v.toFixed(2)}x`
}

// ─── Excel download utility ───────────────────────────────────────────────────
function downloadExcel(data, filename, sheetName = 'Data') {
  if (!data || !data.length) return
  const ws = XLSX.utils.json_to_sheet(data)
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, sheetName)
  XLSX.writeFile(wb, `${filename}.xlsx`)
}

function DownloadButton({ data, filename, sheetName }) {
  const C = useTheme()
  return (
    <button
      onClick={() => downloadExcel(data, filename, sheetName)}
      title="Download chart data as Excel"
      style={{ padding: '5px 10px', borderRadius: 7, border: `1px solid ${C.border}`, background: 'transparent', color: C.text3, fontSize: 11, cursor: 'pointer', fontFamily: 'var(--font)', transition: 'all 0.15s', display: 'flex', alignItems: 'center', gap: 5 }}
      onMouseEnter={e => { e.currentTarget.style.borderColor = C.blue; e.currentTarget.style.color = C.blue }}
      onMouseLeave={e => { e.currentTarget.style.borderColor = C.border; e.currentTarget.style.color = C.text3 }}>
      <svg width="11" height="11" viewBox="0 0 11 11" fill="none"><path d="M5.5 1v7M2.5 6l3 3 3-3M1 10h9" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/></svg>
      Download data
    </button>
  )
}

// ─── Chart explanation overlay ────────────────────────────────────────────────
function ExplainButton({ title, simple, advanced }) {
  const C = useTheme()
  const [open, setOpen] = useState(false)
  const [showAdv, setShowAdv] = useState(false)
  return (
    <>
      <button onClick={() => setOpen(true)}
        style={{ padding: '5px 10px', borderRadius: 7, border: `1px solid ${C.border}`, background: 'transparent', color: C.text3, fontSize: 11, cursor: 'pointer', fontFamily: 'var(--font)', transition: 'all 0.15s', display: 'flex', alignItems: 'center', gap: 5 }}
        onMouseEnter={e => { e.currentTarget.style.borderColor = C.blue; e.currentTarget.style.color = C.blue }}
        onMouseLeave={e => { e.currentTarget.style.borderColor = C.border; e.currentTarget.style.color = C.text3 }}>
        <svg width="12" height="12" viewBox="0 0 12 12" fill="none"><circle cx="6" cy="6" r="5.5" stroke="currentColor"/><text x="6" y="9" textAnchor="middle" fontSize="8" fill="currentColor" fontWeight="600">?</text></svg>
        Explain this
      </button>
      {open && (
        <div style={{ position: 'fixed', inset: 0, zIndex: 9998, display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'rgba(8,10,15,0.85)', backdropFilter: 'blur(8px)' }} onClick={() => { setOpen(false); setShowAdv(false) }}>
          <div onClick={e => e.stopPropagation()} style={{ background: C.bg1, border: `1px solid ${C.border}`, borderRadius: 20, padding: '28px 32px', maxWidth: 480, width: '90%', boxShadow: '0 24px 80px rgba(0,0,0,0.4)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 16 }}>
              <div style={{ fontSize: 15, fontWeight: 700, color: C.text, flex: 1, paddingRight: 16 }}>{title}</div>
              <button onClick={() => { setOpen(false); setShowAdv(false) }} style={{ width: 28, height: 28, borderRadius: 7, background: C.bg2, border: `1px solid ${C.border}`, color: C.text2, fontSize: 14, cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>✕</button>
            </div>
            <div style={{ fontSize: 13, color: C.text2, lineHeight: 1.9, marginBottom: 16 }}>{simple}</div>
            {advanced && (
              <>
                <button onClick={() => setShowAdv(v => !v)} style={{ fontSize: 11, color: C.blue, background: 'none', border: 'none', cursor: 'pointer', fontFamily: 'var(--font)', padding: 0, marginBottom: showAdv ? 12 : 0 }}>
                  {showAdv ? '▲ Hide advanced details' : '▼ Show advanced details'}
                </button>
                {showAdv && <div style={{ fontSize: 12, color: C.text3, lineHeight: 1.8, padding: '12px 14px', background: C.bg2, borderRadius: 10, borderLeft: `3px solid ${C.blue}` }}>{advanced}</div>}
              </>
            )}
          </div>
        </div>
      )}
    </>
  )
}

// ─── Chart wrapper with fullscreen + explain + download ───────────────────────
function ChartCard({ title, subtitle, explain, explainAdv, downloadData, downloadName, children, height = 280 }) {
  const C = useTheme()
  const [fs, setFs] = useState(false)
  const cardStyle = { background: C.bg1, border: `1px solid ${C.border}`, borderRadius: 16, padding: '20px 22px', marginBottom: 14 }
  const btnRow = (
    <div style={{ display: 'flex', gap: 7, flexShrink: 0, flexWrap: 'wrap' }}>
      {explain && <ExplainButton title={title} simple={explain} advanced={explainAdv} />}
      {downloadData && <DownloadButton data={downloadData} filename={downloadName || 'chart-data'} sheetName="Data" />}
      <button
        style={{ padding: '5px 10px', borderRadius: 7, border: `1px solid ${C.border}`, background: 'transparent', color: C.text3, fontSize: 11, cursor: 'pointer', fontFamily: 'var(--font)', transition: 'all 0.15s', display: 'flex', alignItems: 'center', gap: 5 }}
        onMouseEnter={e => { e.currentTarget.style.borderColor = C.blue; e.currentTarget.style.color = C.blue }}
        onMouseLeave={e => { e.currentTarget.style.borderColor = C.border; e.currentTarget.style.color = C.text3 }}
        onClick={() => setFs(true)}>
        <svg width="11" height="11" viewBox="0 0 11 11" fill="none"><path d="M1 4V1h3M7 1h3v3M10 7v3H7M4 10H1V7" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/></svg>
        Expand
      </button>
    </div>
  )
  return (
    <>
      <div style={cardStyle}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 16 }}>
          <div style={{ flex: 1, paddingRight: 12 }}>
            <div style={{ fontSize: 13, fontWeight: 600, color: C.text }}>{title}</div>
            {subtitle && <div style={{ fontSize: 11, color: C.text2, marginTop: 3, lineHeight: 1.5 }}>{subtitle}</div>}
          </div>
          {btnRow}
        </div>
        <ResponsiveContainer width="100%" height={height}>{children}</ResponsiveContainer>
      </div>
      {fs && (
        <div style={{ position: 'fixed', inset: 0, zIndex: 9999, background: C.bg, display: 'flex', flexDirection: 'column', animation: 'fadeUp 0.2s ease' }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '18px 28px', borderBottom: `1px solid ${C.border}`, background: C.bg1 }}>
            <div>
              <div style={{ fontSize: 15, fontWeight: 600, color: C.text }}>{title}</div>
              {subtitle && <div style={{ fontSize: 12, color: C.text2, marginTop: 2 }}>{subtitle}</div>}
            </div>
            <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
              {btnRow}
              <button onClick={() => setFs(false)} style={{ width: 34, height: 34, borderRadius: 8, background: C.bg2, border: `1px solid ${C.border}`, color: C.text2, fontSize: 16, cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>✕</button>
            </div>
          </div>
          <div style={{ flex: 1, padding: '24px 28px' }}>
            <ResponsiveContainer width="100%" height="100%">{children}</ResponsiveContainer>
          </div>
        </div>
      )}
    </>
  )
}

// ─── Shared tooltip box ───────────────────────────────────────────────────────
function TipBox({ children }) {
  const C = useTheme()
  return (
    <div style={{ background: C.bg2, border: `1px solid ${C.border}`, borderRadius: 12, padding: '12px 16px', fontSize: 11, boxShadow: '0 8px 32px rgba(0,0,0,0.3)', minWidth: 180 }}>
      {children}
    </div>
  )
}

// ─── Advanced details toggle ──────────────────────────────────────────────────
function AdvancedToggle({ show, onToggle }) {
  const C = useTheme()
  return (
    <button onClick={onToggle} style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '7px 14px', borderRadius: 8, border: `1px solid ${show ? C.blue : C.border}`, background: show ? C.blueDim : 'transparent', color: show ? C.blue : C.text3, fontSize: 11, cursor: 'pointer', fontFamily: 'var(--font)', transition: 'all 0.2s' }}>
      <svg width="12" height="12" viewBox="0 0 12 12" fill="none"><path d="M2 4h8M4 6h4M5 8h2" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/></svg>
      {show ? 'Hide advanced details' : 'Show advanced details'}
    </button>
  )
}

// ─── KPI row ─────────────────────────────────────────────────────────────────
function KPIRow({ result, method, useRelative, showAdvanced, onToggleAdvanced }) {
  const C = useTheme()
  const methodLabels = { ols: 'Pooled OLS', ridge: 'Ridge', fe: 'Fixed Effects', fe_ridge: 'FE + Ridge' }
  const sigCount = result.featureImp?.filter(f => f.sig).length || 0
  const confLabel = result.r2 > 0.55 ? 'Good' : result.r2 > 0.35 ? 'Moderate' : 'Limited'
  const confCol = result.r2 > 0.55 ? C.green : result.r2 > 0.35 ? C.amber : C.red
  const items = [
    { val: confLabel, label: 'Model confidence', col: confCol, sub: 'Based on fit quality' },
    { val: `${sigCount} of ${result.validF?.length || 0}`, label: 'Strong drivers found', col: C.blue },
    { val: result.n, label: 'Data points used', col: C.text },
    { val: methodLabels[method] || method, label: 'Method', col: C.blue },
  ]
  const advItems = [
    { val: result.r2?.toFixed(3), label: 'R² in-sample', col: result.r2 > 0.55 ? C.green : result.r2 > 0.35 ? C.amber : C.red },
    { val: result.r2Adj ? result.r2Adj.toFixed(3) : '—', label: 'Adjusted R²', col: C.text2 },
  ]
  return (
    <div style={{ marginBottom: 20 }}>
      {useRelative && (
        <div style={{ marginBottom: 12, display: 'inline-flex', alignItems: 'center', gap: 8, padding: '6px 12px', borderRadius: 8, background: C.greenDim, border: `1px solid ${C.green}33`, fontSize: 11, color: C.green }}>
          <svg width="12" height="12" viewBox="0 0 12 12" fill="none"><path d="M6 1v10M1 6h10" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/></svg>
          Relative mode — predicting premium/discount vs sector median
        </div>
      )}
      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginBottom: 10 }}>
        {items.map(({ val, label, col, sub }) => (
          <div key={label} style={{ flex: 1, minWidth: 100, padding: '14px', background: C.bg1, borderRadius: 12, border: `1px solid ${C.border}` }}>
            <div style={{ fontSize: 19, fontWeight: 700, color: col }}>{val}</div>
            <div style={{ fontSize: '9px', color: C.text3, letterSpacing: '0.1em', textTransform: 'uppercase', marginTop: 3 }}>{label}</div>
            {sub && <div style={{ fontSize: 10, color: col, opacity: 0.7, marginTop: 2 }}>{sub}</div>}
          </div>
        ))}
        {showAdvanced && advItems.map(({ val, label, col }) => (
          <div key={label} style={{ flex: 1, minWidth: 100, padding: '14px', background: C.blueDim, borderRadius: 12, border: `1px solid ${C.blue}33` }}>
            <div style={{ fontSize: 19, fontWeight: 700, color: col }}>{val}</div>
            <div style={{ fontSize: '9px', color: C.text3, letterSpacing: '0.1em', textTransform: 'uppercase', marginTop: 3 }}>{label}</div>
          </div>
        ))}
      </div>
      <AdvancedToggle show={showAdvanced} onToggle={onToggleAdvanced} />
    </div>
  )
}

// ─── Custom recharts tooltip components ──────────────────────────────────────
const ScatterTip = ({ active, payload, clientCompany, useRelative, meta }) => {
  const C = useTheme()
  if (!active || !payload?.[0]) return null
  const d = payload[0].payload
  const predKey = d.predicted ?? d.pred ?? 0
  const gap = predKey - d.actual
  const fmt = v => formatValue(v, meta || { unit: 'x', suffix: 'x' })
  return (
    <TipBox>
      <div style={{ fontWeight: 700, color: d.isClient ? C.amber : C.text, marginBottom: 6 }}>{d.isClient ? '★ ' : ''}{d.company || d.co?.split(' ')[0]} · {d.year}</div>
      {useRelative
        ? <>
            <div style={{ color: C.text2, marginBottom: 2 }}>Market premium: <strong style={{ color: C.blue }}>{d.actual > 0 ? '+' : ''}{fmt(d.actual)} vs sector</strong></div>
            <div style={{ color: C.text2, marginBottom: 6 }}>Model says fair: <strong style={{ color: C.green }}>{predKey > 0 ? '+' : ''}{fmt(predKey)} vs sector</strong></div>
          </>
        : <>
            <div style={{ color: C.text2, marginBottom: 2 }}>Actual: <strong style={{ color: C.blue }}>{fmt(d.actual)}</strong></div>
            <div style={{ color: C.text2, marginBottom: 6 }}>Model: <strong style={{ color: C.green }}>{fmt(predKey)}</strong></div>
          </>}
      <div style={{ fontWeight: 700, color: gap > 0 ? C.green : C.red, fontSize: 12 }}>{gap > 0.05 ? `↑ Underpriced` : gap < -0.05 ? `↓ Overpriced` : '≈ Fair'}</div>
    </TipBox>
  )
}

const LineTip = ({ active, payload, label, pivotYear, clientName, meta }) => {
  const C = useTheme()
  if (!active || !payload?.length) return null
  const fmt = v => isFinite(v) ? formatValue(v, meta || { unit: 'x' }) : '—'
  return (
    <TipBox>
      <div style={{ fontWeight: 700, marginBottom: 8, color: C.text }}>{clientName?.split(',')[0]} · {label} {label > pivotYear ? <span style={{ color: C.green, fontSize: 10, fontWeight: 400 }}>out-of-sample</span> : <span style={{ color: C.text3, fontSize: 10 }}>training</span>}</div>
      {payload.map(p => p.dataKey !== 'upper' && p.dataKey !== 'lower' && (
        <div key={p.dataKey} style={{ color: C.text2, marginBottom: 3 }}>{p.name}: <strong style={{ color: p.color }}>{fmt(p.value)}</strong></div>
      ))}
    </TipBox>
  )
}

// ─── Model builder util ───────────────────────────────────────────────────────
function buildModel(allRows, target, features, filterOutliers, method, useRelative = false) {
  const validF = features.filter(f => f !== target)
  const effectiveTarget = useRelative ? '_relativeMultiple' : target
  let workRows = useRelative ? computeRelativeMultiple(allRows, target) : allRows
  let train = workRows.filter(r => isFinite(r[effectiveTarget]) && validF.every(f => isFinite(r[f])))
  if (filterOutliers) train = removeOutliers(train, effectiveTarget)
  if (train.length < validF.length + 3) return null

  if (method === 'fe' || method === 'fe_ridge') {
    const fe = buildFixedEffects(train, validF, effectiveTarget)
    const wr = train.map((r, i) => { const o = {}; validF.forEach((f, j) => { o[f] = fe.XWithin[i][j] }); return o })
    const { ms, ss } = standardize(wr, validF)
    const XwStd = fe.XWithin.map(row => [1, ...row.map((v, i) => (v - ms[i]) / ss[i])])
    const res = method === 'fe_ridge' ? ridge(XwStd, fe.yWithin, 5) : ols(XwStd, fe.yWithin)
    if (!res) return null
    return { beta: res.beta, ms, ss, se: res.se, r2: res.r2, isFE: true, companyMeansY: fe.companyMeansY, companyMeansX: fe.companyMeansX, validF, effectiveTarget, useRelative }
  } else {
    const { ms, ss } = standardize(train, validF)
    const Xs = train.map(r => [1, ...validF.map((f, i) => (r[f] - ms[i]) / ss[i])])
    const y = train.map(r => r[effectiveTarget])
    const res = method === 'ridge' ? ridge(Xs, y, 5) : ols(Xs, y)
    if (!res) return null
    return { beta: res.beta, ms, ss, se: res.se, r2: res.r2, isFE: false, validF, effectiveTarget, useRelative }
  }
}

function predict(row, model) {
  if (!model) return NaN
  if (model.isFE) {
    const co = row._company
    const coMX = model.companyMeansX?.[co] || model.validF.map(() => 0)
    const dem = model.validF.map((f, i) => (row[f] ?? 0) - coMX[i])
    const x = [1, ...dem.map((v, i) => (v - model.ms[i]) / model.ss[i])]
    const wp = x.reduce((s, v, i) => s + v * model.beta[i], 0)
    const base = model.companyMeansY?.[co] ?? avg(Object.values(model.companyMeansY || {}))
    return wp + base
  }
  const x = [1, ...model.validF.map((f, i) => ((row[f] ?? 0) - model.ms[i]) / model.ss[i])]
  return x.reduce((s, v, i) => s + v * model.beta[i], 0)
}

// Convert relative prediction back to absolute using sector median
function relToAbs(relPred, sectorMedian) {
  return isFinite(relPred) && isFinite(sectorMedian) ? relPred + sectorMedian : NaN
}

// ─── Smart range computation ─────────────────────────────────────────────────
function computeSmartRange(pointEst, target, allRows) {
  const vals = allRows.map(r => r[target]).filter(v => isFinite(v) && v > -50 && v < 500)
  if (!vals.length || !isFinite(pointEst)) return null
  const sorted = vals.slice().sort((a, b) => a - b)
  const q1 = sorted[Math.floor(sorted.length * 0.25)]
  const q3 = sorted[Math.floor(sorted.length * 0.75)]
  const iqr = q3 - q1
  const meta = getTargetMeta(target)
  let halfWidth
  const tl = target?.toLowerCase() || ''
  if (tl.includes('ev/ebitda') || tl.includes('evebitda')) halfWidth = Math.min(iqr * 0.32, 3.5)
  else if (tl.includes('ev/sales') || tl.includes('evsales')) halfWidth = Math.min(iqr * 0.32, 0.75)
  else if (meta.unit === 'x') halfWidth = Math.min(iqr * 0.32, 4.0)
  else if (meta.unit === '%') halfWidth = Math.min(iqr * 0.32, 3.0)
  else if (meta.isDollar) halfWidth = Math.abs(pointEst) * 0.12
  else halfWidth = iqr * 0.3
  return { lo: pointEst - halfWidth, hi: pointEst + halfWidth, halfWidth }
}

// ─── Valuation Summary Card ───────────────────────────────────────────────────
function ValuationSummary({ result, target, clientCompany, useRelative, allRows, features, method, filterOutliers }) {
  const C = useTheme()
  const meta = getTargetMeta(target)
  const validF = features.filter(f => f !== target)

  const { absPred, currentActual, latestYear, sectorMedian, confidence, range } = useMemo(() => {
    const model = buildModel(allRows, target, features, filterOutliers, method, useRelative)
    if (!model) return {}
    const clientRows = allRows.filter(r => r._company === clientCompany && isFinite(r[target]) && validF.every(f => isFinite(r[f]))).sort((a, b) => b._year - a._year)
    if (!clientRows.length) return {}
    const latestRow = clientRows[0]
    const rawPred = predict({ ...latestRow, _company: clientCompany }, model)
    const allYearVals = allRows.filter(r => r._year === latestRow._year && isFinite(r[target])).map(r => r[target]).sort((a, b) => a - b)
    const sectorMed = allYearVals.length ? allYearVals[Math.floor(allYearVals.length / 2)] : null
    const absPred = useRelative && isFinite(sectorMed) ? rawPred + sectorMed : rawPred
    const range = computeSmartRange(absPred, target, allRows)
    const confLabel = result.r2 > 0.55 ? 'High' : result.r2 > 0.35 ? 'Moderate' : 'Limited'
    return { absPred, currentActual: latestRow[target], latestYear: latestRow._year, sectorMedian: sectorMed, confidence: confLabel, range }
  }, [allRows, target, features, filterOutliers, method, useRelative, clientCompany, validF, result])

  if (!isFinite(absPred) || !range) return null
  const fmt = v => formatValue(v, meta)
  const gap = absPred - currentActual
  const upsidePct = Math.abs(currentActual) > 0.1 ? ((absPred - currentActual) / Math.abs(currentActual)) * 100 : null
  const confCol = confidence === 'High' ? C.green : confidence === 'Moderate' ? C.amber : C.red
  const inRange = currentActual >= range.lo && currentActual <= range.hi
  const convText = result.r2 > 0.5
    ? 'Based on historical patterns across this peer group, companies have typically reached this range within 1–3 years. Macro conditions (rates, sentiment) can delay or accelerate this.'
    : 'With moderate model confidence, expect gradual alignment over 2–4 years. Use alongside DCF and other valuation approaches for a complete picture.'

  return (
    <div style={{ background: C.bg1, border: `1px solid ${C.border}`, borderRadius: 16, padding: '22px 24px', marginBottom: 14 }}>
      <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.12em', textTransform: 'uppercase', color: C.text3, marginBottom: 14 }}>
        Valuation Summary — {clientCompany?.split(',')[0]} · Based on {latestYear} fundamentals
      </div>
      <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginBottom: 14 }}>
        <div style={{ flex: 2, minWidth: 220, padding: '20px', background: gap > 0 ? C.greenDim : C.redDim, borderRadius: 12, border: `1px solid ${gap > 0 ? C.green + '44' : C.red + '44'}` }}>
          <div style={{ fontSize: 10, color: C.text3, letterSpacing: '0.1em', textTransform: 'uppercase', marginBottom: 10 }}>Model-implied fair value range — {target}</div>
          <div style={{ fontSize: 32, fontWeight: 700, color: gap > 0 ? C.green : C.red, letterSpacing: '-0.02em', marginBottom: 6 }}>
            {fmt(range.lo)} – {fmt(range.hi)}
          </div>
          <div style={{ fontSize: 12, color: C.text2, lineHeight: 1.7 }}>
            Currently at <strong style={{ color: C.blue }}>{fmt(currentActual)}</strong>
            {isFinite(upsidePct) && gap !== 0 && <> · <strong style={{ color: gap > 0 ? C.green : C.red }}>{gap > 0 ? '+' : ''}{upsidePct?.toFixed(0)}% to midpoint</strong></>}
            {inRange && <span style={{ marginLeft: 8, padding: '2px 8px', borderRadius: 4, background: C.amberDim, color: C.amber, fontSize: 10, fontWeight: 600 }}>WITHIN RANGE</span>}
          </div>
        </div>
        {isFinite(sectorMedian) && (
          <div style={{ flex: 1, minWidth: 120, padding: '16px', background: C.bg2, borderRadius: 12, border: `1px solid ${C.border}` }}>
            <div style={{ fontSize: 10, color: C.text3, letterSpacing: '0.1em', textTransform: 'uppercase', marginBottom: 8 }}>Sector median ({latestYear})</div>
            <div style={{ fontSize: 22, fontWeight: 700, color: C.text }}>{fmt(sectorMedian)}</div>
            <div style={{ fontSize: 11, color: currentActual > sectorMedian ? C.green : C.red, marginTop: 4 }}>
              Client {currentActual > sectorMedian ? '+' : ''}{fmt(currentActual - sectorMedian)} vs peers
            </div>
          </div>
        )}
        <div style={{ flex: 1, minWidth: 120, padding: '16px', background: C.bg2, borderRadius: 12, border: `1px solid ${C.border}` }}>
          <div style={{ fontSize: 10, color: C.text3, letterSpacing: '0.1em', textTransform: 'uppercase', marginBottom: 8 }}>Confidence</div>
          <div style={{ fontSize: 22, fontWeight: 700, color: confCol }}>{confidence}</div>
          <div style={{ fontSize: 10, color: C.text3, marginTop: 4 }}>R² = {result.r2?.toFixed(2)}</div>
        </div>
      </div>
      <div style={{ padding: '12px 16px', background: C.bg2, borderRadius: 10, fontSize: 12, color: C.text2, lineHeight: 1.8, borderLeft: `3px solid ${C.amber}` }}>
        <strong style={{ color: C.amber }}>Timing outlook:</strong> {convText}
        <div style={{ marginTop: 6, fontSize: 10, color: C.text3 }}>⚠ Directional signal only — {result.n} observations. Do not anchor deal numbers on this range alone.</div>
      </div>
    </div>
  )
}

// ─── Tab: Overview ────────────────────────────────────────────────────────────
function TabOverview({ result, clientCompany, useRelative, target, showAdvanced, allRows, features, method, filterOutliers }) {
  const C = useTheme()
  const meta = getTargetMeta(target)
  const data = result.chartData?.filter(d => isFinite(d.actual) && isFinite(d.predicted)) || []
  const allVals = data.map(d => d.actual).concat(data.map(d => d.predicted))
  const domain = useMemo(() => getAxisDomain(allVals), [data])
  const downloadData = data.map(d => ({
    Company: d.company, Year: d.year,
    [useRelative ? 'Actual Premium vs Sector' : `Actual ${target}`]: parseFloat(d.actual?.toFixed(3)),
    'Model Prediction': parseFloat(d.predicted?.toFixed(3)),
    'Error': parseFloat((d.predicted - d.actual)?.toFixed(3)),
    'Client': d.isClient ? 'Yes' : 'No'
  }))
  return (
    <>
      <ValuationSummary result={result} target={target} clientCompany={clientCompany} useRelative={useRelative} allRows={allRows} features={features} method={method} filterOutliers={filterOutliers} />
      <ChartCard
        title="Model accuracy — predicted vs actual"
        subtitle="Each dot = one company-year. Closer to the diagonal = better prediction. Orange = your client."
        explain={useRelative
          ? `This shows how accurately the model predicts whether a company should trade above or below its sector peers. Dots on the diagonal = perfect prediction. Above the line = market undervalued. Below = overvalued.`
          : `Each dot is one company in one year. On the diagonal = perfect prediction. The more spread out the dots, the less precise the model.`}
        explainAdv={`Scatter of predicted vs actual (Ŷ vs Y). SE = ${result.se?.toFixed(2)} · R² = ${result.r2?.toFixed(3)} (${(result.r2 * 100).toFixed(0)}% of variance explained).`}
        downloadData={downloadData}
        downloadName={`model-fit-${(target || 'data').replace(/[^a-z0-9]/gi, '-').toLowerCase()}`}
        height={280}>
        <ScatterChart margin={{ top: 8, right: 16, left: 0, bottom: 28 }}>
          <CartesianGrid stroke={C.grid} strokeDasharray="3 3" />
          <XAxis dataKey="actual" type="number" stroke={C.text3} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }} domain={domain} tickFormatter={v => formatAxisTick(v, meta)} label={{ value: useRelative ? 'Actual premium vs sector' : `Actual ${target}`, position: 'insideBottom', offset: -14, fill: C.text3, fontSize: 10 }} />
          <YAxis dataKey="predicted" type="number" stroke={C.text3} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }} tickFormatter={v => formatAxisTick(v, meta)} domain={domain} label={{ value: 'Model prediction', angle: -90, position: 'insideLeft', fill: C.text3, fontSize: 10, dx: -4 }} />
          <ZAxis range={[30, 30]} />
          <Tooltip content={<ScatterTip clientCompany={clientCompany} useRelative={useRelative} meta={meta} />} cursor={false} />
          <ReferenceLine segment={[{ x: domain[0], y: domain[0] }, { x: domain[1], y: domain[1] }]} stroke={C.green + '55'} strokeDasharray="5 3" label={{ value: 'Perfect prediction', fill: C.green, fontSize: 9, position: 'insideTopLeft' }} />
          <Scatter data={data.filter(d => !d.isClient)} fill={C.blue} opacity={0.5} />
          <Scatter data={data.filter(d => d.isClient)} fill={C.amber} />
        </ScatterChart>
      </ChartCard>
      {showAdvanced && (
        <div style={{ padding: '14px 18px', background: C.blueDim, border: `1px solid ${C.blue}22`, borderRadius: 12, fontSize: 12, color: C.text2, lineHeight: 1.8 }}>
          <strong style={{ color: C.blue }}>Technical:</strong> In-sample R² = {result.r2?.toFixed(3)} · Adj. R² = {result.r2Adj?.toFixed(3) || '—'} · SE = {result.se?.toFixed(3)} · n = {result.n}. Use the Validation tab for honest out-of-sample R².
        </div>
      )}
    </>
  )
}

// ─── Tab: Historical Accuracy ─────────────────────────────────────────────────
function TabAsOf({ allRows, target, features, clientCompany, filterOutliers, method, useRelative }) {
  const C = useTheme()
  const validF = features.filter(f => f !== target)
  const years = useMemo(() => [...new Set(allRows.map(r => r._year).filter(isFinite))].sort(), [allRows])
  const [pivot, setPivot] = useState(Math.min(2021, years[years.length - 2] || 2021))

  const { clientData, flatPred, se, r2, trainN, allCoData, scatterData, avgConv, convCount } = useMemo(() => {
    let train = allRows.filter(r => r._year <= pivot && isFinite(r[target]) && validF.every(f => isFinite(r[f])))
    if (filterOutliers) train = removeOutliers(train, target)
    if (train.length < validF.length + 3) return {}

    const model = buildModel(train, target, validF.map(f => f), false, method)
    if (!model) return {}

    const companies = [...new Set(allRows.map(r => r._company).filter(Boolean))]
    const allCoData = {}
    const convDeltas = []

    for (const co of companies) {
      const pivotRows = allRows.filter(r => r._company === co && r._year <= pivot && validF.every(f => isFinite(r[f]))).sort((a, b) => b._year - a._year)
      if (!pivotRows.length) continue
      const fp = predict({ ...pivotRows[0], _company: co }, { ...model, validF: features.filter(f => f !== target) })
      const actuals = allRows.filter(r => r._company === co && isFinite(r[target])).sort((a, b) => a._year - b._year)
      const rows = actuals.map(r => ({ year: r._year, actual: r[target], pred: fp, upper: fp + 1.96 * model.se, lower: fp - 1.96 * model.se, inTrain: r._year <= pivot, converged: Math.abs(fp - r[target]) <= model.se }))
      const firstConv = rows.find(r => !r.inTrain && r.converged)
      if (firstConv) convDeltas.push(firstConv.year - pivot)
      allCoData[co] = { rows, flatPred: fp }
    }

    const clientRows = allCoData[clientCompany]?.rows || []
    const scatterData = Object.entries(allCoData).flatMap(([co, { rows, flatPred }]) =>
      rows.filter(r => isFinite(r.actual) && r.actual > -5 && r.actual < 50).map(r => ({ ...r, co, isClient: co === clientCompany, pred: flatPred }))
    )
    const avgC = convDeltas.length ? (convDeltas.reduce((s, v) => s + v, 0) / convDeltas.length).toFixed(1) : null

    return { clientData: clientRows, flatPred: allCoData[clientCompany]?.flatPred, se: model.se, r2: model.r2, trainN: train.length, allCoData, scatterData, avgConv: avgC, convCount: convDeltas.length }
  }, [allRows, target, validF, pivot, filterOutliers, method, clientCompany, features])

  const otherCos = Object.keys(allCoData || {}).filter(c => c !== clientCompany)
  const clientName = clientCompany?.split(',')[0]

  return (
    <div>
      {/* Pivot selector */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16, flexWrap: 'wrap', gap: 12 }}>
        <div>
          <div style={{ fontSize: 11, color: C.text3, letterSpacing: '0.1em', textTransform: 'uppercase', marginBottom: 8 }}>Train model using data up to</div>
          <div style={{ display: 'flex', gap: 6 }}>
            {years.slice(1).map(y => <div key={y} className={`chip ${pivot === y ? 'on-green' : ''}`} onClick={() => setPivot(y)}><span style={{ width: 6, height: 6, borderRadius: '50%', background: 'currentColor', opacity: pivot === y ? 1 : 0.3 }} />{y}</div>)}
          </div>
        </div>
        <div style={{ display: 'flex', gap: 10 }}>
          {[{ val: trainN, label: 'Training obs' }, { val: r2?.toFixed(2), label: 'Train R²', col: r2 > 0.5 ? C.green : C.amber }, { val: avgConv ? `~${avgConv}yr` : '—', label: 'Avg convergence', col: C.amber }].map(({ val, label, col }) => (
            <div key={label} className="kpi" style={{ minWidth: 80 }}>
              <div className="kpi-val" style={{ color: col || C.text, fontSize: 18 }}>{val}</div>
              <div className="kpi-label">{label}</div>
            </div>
          ))}
        </div>
      </div>

      {clientData?.length > 0 && (() => {
        const meta = getTargetMeta(target)
        const clientVals = clientData.map(d => d.actual).filter(isFinite)
        const clientDomain = getAxisDomain([...clientVals, isFinite(flatPred) ? flatPred : NaN].filter(isFinite))
        const clientDl = clientData.map(r => ({ Year: r.year, [`Actual ${target}`]: parseFloat(r.actual?.toFixed(4)), [`Fair Value (as of ${pivot})`]: parseFloat(flatPred?.toFixed(4)), 'In Training': r.inTrain ? 'Yes' : 'No', 'Converged': r.converged ? 'Yes' : 'No' }))
        return (
          <ChartCard
            title={`${clientName} — Actual vs model fair value`}
            subtitle={`Flat dashed line = model's one-time prediction using ${pivot} fundamentals. When orange touches green, the model was confirmed.`}
            explain={`The dashed green line is a FLAT line — the model made one prediction based on ${pivot} data and it never changes. The orange line is what the market actually paid each year. When they meet, the market came to agree with the model. Green dots = years where they converged.`}
            downloadData={clientDl}
            downloadName={`${clientName?.replace(/\s+/g, '-').toLowerCase()}-historical`}
            height={260}>
            <ComposedChart data={clientData} margin={{ top: 8, right: 20, left: 0, bottom: 8 }}>
              <defs>
                <linearGradient id="ciG" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor={C.green} stopOpacity={0.12} />
                  <stop offset="95%" stopColor={C.green} stopOpacity={0.01} />
                </linearGradient>
              </defs>
              <CartesianGrid stroke={C.grid} strokeDasharray="3 3" />
              <XAxis dataKey="year" stroke={C.text3} tick={{ fontSize: 10 }} />
              <YAxis stroke={C.text3} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }} tickFormatter={v => formatAxisTick(v, meta)} domain={clientDomain} />
              <Tooltip content={<LineTip pivotYear={pivot} clientName={clientName} meta={meta} />} />
              <Legend wrapperStyle={{ fontSize: 11, color: C.text2, paddingTop: 8 }} />
              <Area type="monotone" dataKey="upper" stroke="none" fill="url(#ciG)" legendType="none" />
              <Area type="monotone" dataKey="lower" stroke="none" fill={C.bg} legendType="none" />
              <ReferenceLine x={pivot} stroke={C.green + '66'} strokeDasharray="6 3" label={{ value: `← ${pivot} cutoff`, fill: C.green, fontSize: 9, position: 'insideTopLeft' }} />
              <Line type="monotone" dataKey="pred" stroke={C.green} strokeWidth={2} strokeDasharray="8 4" dot={false} name={`Fair value (as of ${pivot}): ${formatValue(flatPred, meta)}`} />
              <Line type="monotone" dataKey="actual" stroke={C.amber} strokeWidth={2.5}
                dot={props => { const { cx, cy, payload } = props; if (!isFinite(cx) || !isFinite(cy)) return null; const c = payload.converged ? C.green : payload.inTrain ? C.text3 : C.amber; return <circle key={cx} cx={cx} cy={cy} r={5} fill={c} stroke={C.bg} strokeWidth={2} /> }}
                activeDot={{ r: 8 }} name={`Actual ${target}`} />
            </ComposedChart>
          </ChartCard>
        )
      })()}

      {/* All companies overlay */}
      {(() => {
        const meta = getTargetMeta(target)
        const allActuals = Object.values(allCoData || {}).flatMap(({ rows }) => rows.map(r => r.actual)).filter(isFinite)
        const overlayDomain = getAxisDomain(allActuals, 0.05)
        const overlayDownload = Object.entries(allCoData || {}).flatMap(([co, { rows, flatPred }]) =>
          rows.map(r => ({ Company: co, Year: r.year, [`Actual ${target}`]: parseFloat(r.actual?.toFixed(4)), 'Model Fair Value': parseFloat(flatPred?.toFixed(4)), 'In Training': r.inTrain ? 'Yes' : 'No', 'Client': co === clientCompany ? 'Yes' : 'No' }))
        )
        return (
          <ChartCard title="All comparables over time" subtitle={`Grey = peers · Orange = ${clientName} · Dashed green = model fair value for client`}
            explain="All companies' values over time on one chart. Your client stands out in orange. The dashed green line is the flat model prediction for your client — fixed at the training cutoff. When orange meets green, the model was confirmed."
            downloadData={overlayDownload} downloadName="all-companies-over-time" height={240}>
            <ComposedChart margin={{ top: 8, right: 20, left: 0, bottom: 8 }}>
              <CartesianGrid stroke={C.grid} strokeDasharray="3 3" />
              <XAxis dataKey="year" type="number" domain={[years[0], years[years.length - 1]]} tickCount={years.length} stroke={C.text3} tick={{ fontSize: 10 }} />
              <YAxis stroke={C.text3} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }} tickFormatter={v => formatAxisTick(v, meta)} domain={overlayDomain} />
              <Tooltip contentStyle={{ background: C.bg2, border: `1px solid ${C.border}`, borderRadius: 10, fontSize: 11 }} formatter={(v, n) => [isFinite(v) ? formatAxisTick(parseFloat(v), meta) : '—', n]} />
              <ReferenceLine x={pivot} stroke={C.green + '44'} strokeDasharray="5 3" />
              {otherCos.map(co => <Line key={co} data={(allCoData[co]?.rows || []).map(r => ({ year: r.year, [co]: r.actual }))} dataKey={co} stroke={C.text3 + '55'} strokeWidth={1} dot={false} legendType="none" connectNulls />)}
              {clientData && <Line data={clientData.map(r => ({ year: r.year, actual: r.actual }))} dataKey="actual" stroke={C.amber} strokeWidth={3} dot={{ r: 4, fill: C.amber, stroke: C.bg, strokeWidth: 2 }} name={clientName} connectNulls />}
              {flatPred && <Line data={clientData?.map(r => ({ year: r.year, pred: flatPred }))} dataKey="pred" stroke={C.green} strokeWidth={2} strokeDasharray="7 4" dot={false} name="Model fair value" connectNulls />}
            </ComposedChart>
          </ChartCard>
        )
      })()}

      {/* Scatter predicted vs actual */}
      {(() => {
        const meta = getTargetMeta(target)
        const scatterVals = (scatterData || []).filter(d => isFinite(d.actual) && isFinite(d.pred)).flatMap(d => [d.actual, d.pred])
        const sDomain = getAxisDomain(scatterVals)
        const scatterDl = (scatterData || []).filter(d => isFinite(d.actual)).map(d => ({ Company: d.co, Year: d.year, [`Actual ${target}`]: parseFloat(d.actual?.toFixed(4)), 'Predicted': parseFloat(d.pred?.toFixed(4)), 'Client': d.isClient ? 'Yes' : 'No', 'Out-of-sample': d.isOOS ? 'Yes' : 'No' }))
        return (
          <ChartCard title="Predicted vs Actual — all companies, all years" subtitle="Points above diagonal = market undervalued. Points below = overpriced. Orange = client."
            explain="Each dot is one company in one year. The diagonal = perfect prediction. Above = model said higher value than market paid. Below = market overpaid. Orange = your client across all years."
            downloadData={scatterDl} downloadName="predicted-vs-actual" height={280}>
            <ScatterChart margin={{ top: 8, right: 20, left: 0, bottom: 28 }}>
              <CartesianGrid stroke={C.grid} strokeDasharray="3 3" />
              <XAxis dataKey="actual" type="number" stroke={C.text3} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }} label={{ value: `Actual ${target} →`, position: 'insideBottom', offset: -14, fill: C.text3, fontSize: 10 }} domain={sDomain} tickFormatter={v => formatAxisTick(v, meta)} />
              <YAxis dataKey="pred" type="number" stroke={C.text3} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }} tickFormatter={v => formatAxisTick(v, meta)} domain={sDomain} />
              <ZAxis range={[28, 28]} />
              <Tooltip content={<ScatterTip clientCompany={clientCompany} useRelative={useRelative} meta={meta} />} cursor={false} />
              <ReferenceLine segment={[{ x: sDomain[0], y: sDomain[0] }, { x: sDomain[1], y: sDomain[1] }]} stroke={C.green + '55'} strokeDasharray="5 3" label={{ value: 'Perfect fit', fill: C.green, fontSize: 9, position: 'insideTopLeft' }} />
              <Scatter data={(scatterData || []).filter(d => !d.isClient && isFinite(d.actual) && isFinite(d.pred))} fill={C.blue} opacity={0.5} />
              <Scatter data={(scatterData || []).filter(d => d.isClient && isFinite(d.actual) && isFinite(d.pred))} fill={C.amber} />
            </ScatterChart>
          </ChartCard>
        )
      })()}
    </div>
  )
}

// ─── Tab: Investment Case ──────────────────────────────────────────────────────
function TabInvestment({ allRows, target, features, clientCompany, filterOutliers, method, useRelative, showAdvanced }) {
  const C = useTheme()
  const meta = getTargetMeta(target)
  const fmt = v => formatValue(v, meta)
  const model = useMemo(() => buildModel(allRows, target, features, filterOutliers, method, useRelative), [allRows, target, features, filterOutliers, method, useRelative])
  const validF = features.filter(f => f !== target)

  const { upsideData, clientUpside, clientTrend } = useMemo(() => {
    if (!model) return {}
    const companies = [...new Set(allRows.map(r => r._company).filter(Boolean))]

    // Pre-compute sector medians per year for relative→absolute conversion
    const sectorMedians = {}
    if (useRelative) {
      const yearGroups = {}
      for (const r of allRows) {
        if (!isFinite(r[target])) continue
        if (!yearGroups[r._year]) yearGroups[r._year] = []
        yearGroups[r._year].push(r[target])
      }
      for (const [yr, vals] of Object.entries(yearGroups)) {
        const s = vals.slice().sort((a, b) => a - b)
        sectorMedians[yr] = s[Math.floor(s.length / 2)]
      }
    }

    const rows = []
    for (const co of companies) {
      const latest = allRows.filter(r => r._company === co && isFinite(r[target]) && validF.every(f => isFinite(r[f]))).sort((a, b) => b._year - a._year)[0]
      if (!latest) continue

      const rawPred = predict({ ...latest, _company: co }, model)
      const rawActual = latest[target]

      // Always work in absolute terms for upside calculation
      const sectorMed = sectorMedians[latest._year] ?? 0
      const absPred = useRelative ? rawPred + sectorMed : rawPred
      const absActual = useRelative ? rawActual : rawActual  // rawActual is already absolute

      // Use absolute values for % upside — avoid division by near-zero
      const gap = absPred - absActual
      const denominator = Math.max(Math.abs(absActual), 0.5)  // floor at 0.5x to avoid inf
      const upsidePct = (gap / denominator) * 100

      rows.push({
        co, shortName: compShort(co),
        actual: absActual,    // absolute actual
        pred: absPred,        // absolute prediction
        relActual: rawActual, // relative actual (for relative display)
        relPred: rawPred,     // relative prediction (for relative display)
        upsidePct: isFinite(upsidePct) ? upsidePct : 0,
        gap: isFinite(gap) ? gap : 0,
        year: latest._year, isClient: co === clientCompany
      })
    }
    rows.sort((a, b) => b.upsidePct - a.upsidePct)

    // Client trend — always absolute
    const trend = allRows.filter(r => r._company === clientCompany && isFinite(r[target]) && validF.every(f => isFinite(r[f]))).sort((a, b) => a._year - b._year).map(r => {
      const rawP = predict({ ...r, _company: clientCompany }, model)
      const sectorMed = sectorMedians[r._year] ?? 0
      const absP = useRelative ? rawP + sectorMed : rawP
      return { year: r._year, actual: r[target], pred: absP }
    })
    return { upsideData: rows, clientUpside: rows.find(r => r.isClient), clientTrend: trend }
  }, [model, allRows, target, clientCompany, validF, useRelative])

  if (!model) return <div style={{ padding: 32, textAlign: 'center', color: C.text2 }}>Run the regression first.</div>

  const clientRankIdx = upsideData?.findIndex(r => r.isClient) ?? -1
  const clientRank = clientRankIdx >= 0 ? clientRankIdx + 1 : '?'
  const clientName = clientCompany?.split(',')[0]

  return (
    <div>
      {/* Headline card */}
      {clientUpside && (
        <div style={{ background: clientUpside.gap > 0 ? C.greenDim : C.redDim, border: `1px solid ${clientUpside.gap > 0 ? C.green + '44' : C.red + '44'}`, borderRadius: 16, padding: '20px 24px', marginBottom: 14 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: 16 }}>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 11, color: C.text3, letterSpacing: '0.1em', textTransform: 'uppercase', marginBottom: 8 }}>Investment Case — {clientName}</div>
              <div style={{ fontSize: 14, color: C.text, lineHeight: 1.8 }}>
                At today's <strong style={{ color: C.blue }}>{fmt(clientUpside.actual)}</strong>, the model estimates fair value at <strong style={{ color: C.green }}>{fmt(clientUpside.pred)}</strong>.
                {useRelative && <span style={{ fontSize: 11, color: C.text3 }}> (relative premium of {clientUpside.relPred > 0 ? '+' : ''}{fmt(clientUpside.relPred)} vs sector median)</span>}
                {' '}That gap implies a <strong style={{ color: clientUpside.gap > 0 ? C.green : C.red, fontSize: 16 }}>{clientUpside.upsidePct > 0 ? '+' : ''}{isFinite(clientUpside.upsidePct) ? clientUpside.upsidePct.toFixed(0) : '?'}% {clientUpside.gap > 0 ? 're-rating opportunity' : 'downside risk'}</strong>.
                {' '}Client ranks <strong style={{ color: C.amber }}>#{clientRank}</strong> of {upsideData?.length} by model-implied upside.
              </div>
              <div style={{ marginTop: 10, padding: '6px 12px', background: C.amberDim, border: `1px solid ${C.amber}25`, borderRadius: 8, display: 'inline-block', fontSize: 10, color: C.amber }}>
                ⚠ Directional signal — use to support narrative, not as a precise valuation
              </div>
            </div>
            <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
              {[{ v: fmt(clientUpside.actual), l: 'Current', c: C.blue }, { v: fmt(clientUpside.pred), l: 'Model fair value', c: C.green }, { v: `${(clientUpside.upsidePct ?? 0) > 0 ? '+' : ''}${isFinite(clientUpside.upsidePct) ? (clientUpside.upsidePct ?? 0).toFixed(0) : '?'}%`, l: 'Implied upside', c: clientUpside.gap > 0 ? C.green : C.red }, { v: `#${clientRank}`, l: 'Rank vs peers', c: C.amber }]
              .map(({ v, l, c }) => (
                <div key={l} className="kpi" style={{ minWidth: 80 }}><div className="kpi-val" style={{ color: c, fontSize: 20 }}>{v}</div><div className="kpi-label">{l}</div></div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* Upside chart */}
      <ChartCard
        title={useRelative ? 'Model-implied sector premium by company' : 'Model-implied upside by company'}
        subtitle={useRelative ? 'How much above/below sector median each company should trade, vs where they trade today.' : 'Current market multiple vs model fair value for each company.'}
        explain={`This ranks all companies by how much upside the model sees. ${useRelative ? 'Green bars = companies the model thinks deserve to trade at a higher premium to peers than they currently do.' : 'Green bars = companies the model thinks are undervalued.'} Your client is shown in orange. The longer the bar, the bigger the gap between current price and model estimate. Remember this is a directional signal, not a precise target.`}
        explainAdv={`The upside % is computed as (predicted - actual) / |actual| × 100. In relative mode, "predicted" is the model's estimate of the company's premium/discount vs the annual sector median, and "actual" is the observed premium/discount.`}
        height={Math.max(300, (upsideData?.length || 0) * 36)}>
        <BarChart data={upsideData || []} layout="vertical" margin={{ top: 4, right: 70, left: 150, bottom: 4 }}>
          <CartesianGrid stroke={C.grid} strokeDasharray="3 3" horizontal={false} />
          <XAxis type="number" stroke={C.text3} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }} tickFormatter={v => `${v > 0 ? '+' : ''}${v.toFixed(0)}%`} domain={[-80, 220]} />
          <YAxis type="category" dataKey="shortName" stroke={C.text3} tick={{ fontSize: 10, fill: C.text2 }} width={145} />
          <Tooltip content={({ active, payload }) => {
            if (!active || !payload?.[0]) return null
            const d = payload[0].payload
            return <TipBox>
              <div style={{ fontWeight: 700, color: d.isClient ? C.amber : C.text, marginBottom: 5 }}>{d.co?.split(',')[0]}</div>
              <div style={{ color: C.text2 }}>Current: <strong style={{ color: C.blue }}>{fmt(d.actual)}{useRelative ? ' vs sector' : ''}</strong></div>
              <div style={{ color: C.text2 }}>Model fair: <strong style={{ color: C.green }}>{fmt(d.pred)}{useRelative ? ' vs sector' : ''}</strong></div>
              <div style={{ color: d.gap > 0 ? C.green : C.red, fontWeight: 700, marginTop: 5 }}>{d.gap > 0 ? '↑ Upside' : '↓ Downside'}: {d.upsidePct > 0 ? '+' : ''}{isFinite(d.upsidePct) ? d.upsidePct.toFixed(0) : '?'}%</div>
            </TipBox>
          }} />
          <ReferenceLine x={0} stroke={C.text3} />
          <Bar dataKey="upsidePct" radius={[0, 4, 4, 0]} barSize={20}>
            <LabelList dataKey="upsidePct" position="right" style={{ fill: C.text3, fontSize: 9, fontFamily: 'var(--mono)' }} formatter={v => `${v > 0 ? '+' : ''}${v.toFixed(0)}%`} />
            {(upsideData || []).map((d, i) => <Cell key={i} fill={d.isClient ? C.amber : d.gap > 0 ? C.green + 'bb' : C.red + '99'} />)}
          </Bar>
        </BarChart>
      </ChartCard>

      {/* Client trend */}
      {clientTrend?.length > 0 && (
        <ChartCard title={`${clientName} — How model fair value has evolved`} subtitle="If green rises faster than orange, the investment case is strengthening. Gap = re-rating opportunity." height={240}>
          <ComposedChart data={clientTrend} margin={{ top: 8, right: 20, left: 0, bottom: 8 }}>
            <CartesianGrid stroke={C.grid} strokeDasharray="3 3" />
            <XAxis dataKey="year" stroke={C.text3} tick={{ fontSize: 10 }} />
            <YAxis stroke={C.text3} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }} tickFormatter={v => formatAxisTick(v, meta)} />
            <Tooltip contentStyle={{ background: C.bg2, border: `1px solid ${C.border}`, borderRadius: 10, fontSize: 11 }} formatter={(v, n) => [isFinite(v) ? fmt(v) : '—', n]} />
            <Legend wrapperStyle={{ fontSize: 11, color: C.text2 }} />
            <Line type="monotone" dataKey="pred" stroke={C.green} strokeWidth={2} strokeDasharray="7 3" dot={false} name="Model fair value" />
            <Line type="monotone" dataKey="actual" stroke={C.amber} strokeWidth={2.5} dot={{ r: 4, fill: C.amber, stroke: C.bg, strokeWidth: 2 }} name="Actual market multiple" />
          </ComposedChart>
        </ChartCard>
      )}
    </div>
  )
}

// ─── Tab: Validation ──────────────────────────────────────────────────────────
function TabValidation({ allRows, target, features, filterOutliers }) {
  const C = useTheme()
  const [results, setResults] = useState(null)
  const [running, setRunning] = useState(false)
  const validF = features.filter(f => f !== target)
  let rows = allRows.filter(r => isFinite(r[target]) && validF.every(f => isFinite(r[f])))
  if (filterOutliers) rows = removeOutliers(rows, target)

  const run = meth => {
    setRunning(meth)
    setTimeout(() => { setResults({ wf: walkForwardCV(rows, validF, target, meth, 5), method: meth }); setRunning(false) }, 80)
  }
  const r2Col = v => v > 0.4 ? C.green : v > 0 ? C.amber : C.red
  const methodLabels = { ols: 'Pooled OLS', ridge: 'Pooled Ridge', fe: 'Fixed Effects', fe_ridge: 'FE + Ridge' }

  return (
    <div>
      <div style={{ background: C.bg1, border: `1px solid ${C.border}`, borderRadius: 16, padding: '20px 22px', marginBottom: 14 }}>
        <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 4 }}>Walk-Forward Validation</div>
        <div style={{ fontSize: 12, color: C.text2, marginBottom: 16, lineHeight: 1.7 }}>Trains on all years up to a cutoff, predicts the next year only. Honest out-of-sample test — no data leakage. <strong style={{ color: C.amber }}>Negative R² years (2022) are due to rate shock, not model failure.</strong></div>
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
          {Object.entries(methodLabels).map(([k, v]) => (
            <button key={k} onClick={() => run(k)} disabled={!!running}
              style={{ padding: '9px 18px', borderRadius: 10, border: `1.5px solid ${results?.method === k ? C.blue : C.border}`, background: results?.method === k ? C.blueDim : C.bg2, color: results?.method === k ? C.blue : C.text2, fontSize: 12, fontWeight: 500, cursor: running ? 'wait' : 'pointer', fontFamily: 'var(--font)', transition: 'all 0.15s', display: 'flex', alignItems: 'center', gap: 8 }}>
              {running === k && <div style={{ width: 12, height: 12, border: '2px solid var(--bg3)', borderTopColor: C.blue, borderRadius: '50%', animation: 'spin 0.7s linear infinite' }} />}
              {running === k ? 'Running…' : `Run ${v}`}
            </button>
          ))}
        </div>
      </div>

      {results && (
        <>
          <div style={{ display: 'flex', gap: 10, marginBottom: 14, flexWrap: 'wrap' }}>
            <div className="kpi"><div className="kpi-val" style={{ color: C.blue, fontSize: 16 }}>{methodLabels[results.method]}</div><div className="kpi-label">Method</div></div>
            <div className="kpi"><div className="kpi-val" style={{ color: C.green }}>{results.wf.filter(r => r.r2 > 0).length}/{results.wf.length}</div><div className="kpi-label">Years R² &gt; 0</div></div>
            <div className="kpi"><div className="kpi-val" style={{ color: C.amber }}>{results.wf.length ? (results.wf.reduce((s, r) => s + r.mae, 0) / results.wf.length).toFixed(1) + 'x' : '—'}</div><div className="kpi-label">Avg error (MAE)</div></div>
            <div className="kpi"><div className="kpi-val" style={{ color: r2Col(results.wf.reduce((s, r) => s + r.r2, 0) / results.wf.length) }}>{(results.wf.reduce((s, r) => s + r.r2, 0) / results.wf.length).toFixed(2)}</div><div className="kpi-label">Avg R² out-of-sample</div></div>
          </div>
          <ChartCard
            title="How well the model predicts — year by year"
            subtitle="Each bar = one year tested. Green = model predicted well. Red = poor year (often a macro shock like rate hikes)."
            explain="This is the honest test. For each year, we train the model on all past data, then predict that year without looking at it. The bars show R² — how much of the variation the model explained. Above 0 = better than random. The orange line is the average error in x terms. 2022 is typically red because rapid rate increases caused market-wide multiple compression that no fundamentals model can predict — that's normal, not a model failure."
            explainAdv="Walk-forward validation (expanding window). R² < 0 means predictions were worse than simply guessing the mean. MAE = Mean Absolute Error in the units of the target variable.">
            <ComposedChart data={results.wf} margin={{ top: 8, right: 20, left: 0, bottom: 8 }}>
              <CartesianGrid stroke={C.grid} strokeDasharray="3 3" />
              <XAxis dataKey="testYear" stroke={C.text3} tick={{ fontSize: 10 }} />
              <YAxis yAxisId="r2" stroke={C.text3} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }} tickFormatter={v => v.toFixed(2)} label={{ value: 'R²', angle: -90, position: 'insideLeft', fill: C.text3, fontSize: 10 }} />
              <YAxis yAxisId="mae" orientation="right" stroke={C.amber} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }} tickFormatter={v => `${v.toFixed(1)}x`} label={{ value: 'Avg error', angle: 90, position: 'insideRight', fill: C.amber, fontSize: 10 }} />
              <Tooltip contentStyle={{ background: C.bg2, border: `1px solid ${C.border}`, borderRadius: 10, fontSize: 11 }}
                content={({ active, payload, label }) => {
                  if (!active || !payload?.length) return null
                  const r2p = payload.find(p => p.dataKey === 'r2')
                  const maep = payload.find(p => p.dataKey === 'mae')
                  return (
                    <TipBox>
                      <div style={{ fontWeight: 700, color: C.text, marginBottom: 6 }}>{label}</div>
                      {r2p && <div style={{ color: C.text2, marginBottom: 3 }}>Model accuracy (R²): <strong style={{ color: r2p.value > 0 ? C.green : C.red }}>{r2p.value.toFixed(3)}{r2p.value > 0.4 ? ' ✓ Strong' : r2p.value > 0 ? ' ~ Moderate' : ' ✗ Weak'}</strong></div>}
                      {maep && <div style={{ color: C.text2 }}>Average error: <strong style={{ color: C.amber }}>{maep.value.toFixed(2)}x off</strong></div>}
                    </TipBox>
                  )
                }} />
              <Legend wrapperStyle={{ fontSize: 11, color: C.text2 }} />
              <ReferenceLine y={0} yAxisId="r2" stroke={C.text3} strokeDasharray="4 2" />
              <Bar yAxisId="r2" dataKey="r2" name="Prediction accuracy (R²)" radius={[4, 4, 0, 0]} barSize={28}>
                {results.wf.map((r, i) => <Cell key={i} fill={r.r2 > 0 ? C.green + 'dd' : C.red + 'dd'} />)}
              </Bar>
              <Line yAxisId="mae" type="monotone" dataKey="mae" stroke={C.amber} strokeWidth={2} dot={{ r: 4, fill: C.amber, stroke: C.bg, strokeWidth: 2 }} name="Avg error in units (MAE)" />
            </ComposedChart>
          </ChartCard>
        </>
      )}
    </div>
  )
}

// ─── Tab: Key Drivers ─────────────────────────────────────────────────────────
function TabDrivers({ result, showAdvanced }) {
  const C = useTheme()
  if (!result?.featureImp) return null
  return (
    <ChartCard
      title="What drives the multiple"
      subtitle="Teal bars = variables we are confident about. Hover each bar to see the effect."
      explain="This chart shows which company characteristics have the strongest link to the multiple. The longer the bar, the more confident we are that this variable genuinely affects how the market prices companies. Teal bars cross the orange threshold line — those are the ones worth talking about. Grey bars don't have enough statistical evidence yet."
      explainAdv={`Each bar shows the absolute t-statistic for that variable's coefficient. The orange line at 1.96 is the threshold for statistical significance at the 5% level (p < 0.05). Variables above this threshold have coefficients statistically different from zero with high confidence. The direction of the effect (positive or negative) is visible in the hover tooltip via the beta coefficient.`}
      height={Math.max(280, result.featureImp.length * 54)}>
      <BarChart data={result.featureImp} layout="vertical" margin={{ top: 4, right: 80, left: 30, bottom: 24 }}>
        <CartesianGrid stroke={C.grid} strokeDasharray="3 3" horizontal={false} />
        <XAxis type="number" stroke={C.text3} tick={{ fontSize: 9 }} label={{ value: 'Statistical strength', position: 'insideBottom', offset: -14, fill: C.text3, fontSize: 9 }} />
        <YAxis type="category" dataKey="short" stroke={C.text3}
          tick={({ x, y, payload }) => <text x={x - 4} y={y} textAnchor="end" fill={C.text2} fontSize={10} dominantBaseline="middle" fontFamily="var(--font)">{payload.value}</text>}
          width={240} />
        <Tooltip content={({ active, payload }) => {
          if (!active || !payload?.[0]) return null
          const d = payload[0].payload
          return <TipBox>
            <div style={{ fontWeight: 700, color: d.sig ? C.green : C.text, marginBottom: 6 }}>{d.name}</div>
            <div style={{ color: C.text2, marginBottom: 3 }}>Effect: <strong style={{ color: d.beta > 0 ? C.green : C.red }}>{d.beta > 0 ? 'Higher → higher multiple' : 'Higher → lower multiple'}</strong></div>
            <div style={{ color: C.text2, marginBottom: 3 }}>Confidence: <strong style={{ color: d.pVal < 0.01 ? C.green : d.pVal < 0.05 ? C.amber : C.text3 }}>{d.pVal < 0.01 ? 'High' : d.pVal < 0.05 ? 'Medium' : 'Low'}</strong></div>
            {showAdvanced && <div style={{ color: C.text3, fontSize: 10, marginTop: 4, borderTop: `1px solid ${C.border}`, paddingTop: 4 }}>β = {d.beta.toFixed(4)} · t = {d.tStat.toFixed(3)} · p = {d.pVal < 0.0001 ? '<0.0001' : d.pVal.toFixed(4)}</div>}
          </TipBox>
        }} />
        <ReferenceLine x={1.96} stroke={C.amber} strokeDasharray="4 2" label={{ value: 'Confidence threshold', fill: C.amber, fontSize: 9, position: 'insideTopRight' }} />
        <Bar dataKey="tStat" radius={[0, 5, 5, 0]} barSize={26}>
          <LabelList dataKey="tStat" position="right" style={{ fill: C.text3, fontSize: 9, fontFamily: 'var(--mono)' }} formatter={v => v.toFixed(1)} />
          {result.featureImp.map((e, i) => <Cell key={i} fill={e.sig ? C.green + 'cc' : C.bg3} />)}
        </Bar>
      </BarChart>
    </ChartCard>
  )
}

// ─── Tab: Scenarios ───────────────────────────────────────────────────────────
function TabScenarios({ allRows, target, features, clientCompany, filterOutliers, method, useRelative }) {
  const C = useTheme()
  const model = useMemo(() => buildModel(allRows, target, features, filterOutliers, method, useRelative), [allRows, target, features, filterOutliers, method, useRelative])
  const validF = features.filter(f => f !== target)
  const clientRows = allRows.filter(r => r._company === clientCompany && validF.every(f => isFinite(r[f]))).sort((a, b) => b._year - a._year)
  const lastRow = clientRows[0]
  const init = useMemo(() => lastRow ? Object.fromEntries(validF.map(f => [f, lastRow[f]])) : {}, [lastRow, validF])
  const [inputs, setInputs] = useState({})
  const merged = { ...init, ...inputs, _company: clientCompany }

  const { base, bull, bear, topF, baseAbs, bullAbs, bearAbs } = useMemo(() => {
    if (!model || !lastRow) return {}
    const b = predict(merged, model)
    // Pick the variable with the largest standardized effect (highest |beta * ss|) as scenario driver
    const tfIdx = validF.reduce((bestIdx, f, i) => {
      const score = Math.abs((model.beta?.[i + 1] ?? 0) * (model.ss?.[i] ?? 1))
      const bestScore = Math.abs((model.beta?.[bestIdx + 1] ?? 0) * (model.ss?.[bestIdx] ?? 1))
      return score > bestScore ? i : bestIdx
    }, 0)
    const tf = validF[tfIdx]
    const fsd = model.ss?.[tfIdx] || 1
    const bullVal = predict({ ...merged, [tf]: (merged[tf] || 0) + fsd }, model)
    const bearVal = predict({ ...merged, [tf]: (merged[tf] || 0) - fsd }, model)

    // If relative mode, convert back to absolute using latest sector median
    let sectorMed = null
    if (useRelative) {
      const latestYearRows = allRows.filter(r => r._year === lastRow._year && isFinite(r[target]))
      const vals = latestYearRows.map(r => r[target]).sort((a, b) => a - b)
      sectorMed = vals.length ? vals[Math.floor(vals.length / 2)] : 0
    }
    const toAbs = v => useRelative && isFinite(sectorMed) ? v + sectorMed : v
    return { base: b, bull: bullVal, bear: bearVal, topF: tf, baseAbs: toAbs(b), bullAbs: toAbs(bullVal), bearAbs: toAbs(bearVal) }
  }, [model, merged, validF, lastRow, useRelative, allRows, target])

  if (!model) return <div style={{ padding: 32, textAlign: 'center', color: C.text2 }}>Run the regression first.</div>
  if (!lastRow) return <div style={{ padding: 32, color: C.text2, fontSize: 12 }}>No data found for {clientCompany}.</div>

  return (
    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
      <div style={{ background: C.bg1, border: `1px solid ${C.border}`, borderRadius: 16, padding: '20px 22px' }}>
        <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 4 }}>Adjust inputs</div>
        <div style={{ fontSize: 11, color: C.text2, marginBottom: 16 }}>Based on {clientCompany?.split(',')[0]}'s latest data ({lastRow._year}). Yellow border = modified.</div>
        {validF.map(f => (
          <div key={f} style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 10, color: C.text3, letterSpacing: '0.1em', textTransform: 'uppercase', marginBottom: 5 }}>{f}</div>
            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <input type="number" step="0.1" value={parseFloat((merged[f] ?? 0).toFixed(4))}
                onChange={e => setInputs(p => ({ ...p, [f]: parseFloat(e.target.value) || 0 }))}
                className={`num-input ${inputs[f] !== undefined ? 'mod' : ''}`} />
              {inputs[f] !== undefined && <button className="reset-btn" onClick={() => setInputs(p => { const n = { ...p }; delete n[f]; return n })}>↺ Reset</button>}
            </div>
          </div>
        ))}
      </div>

      <div style={{ background: C.bg1, border: `1px solid ${C.border}`, borderRadius: 16, padding: '20px 22px' }}>
        <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 16 }}>Output</div>
        {baseAbs !== undefined && (() => {
          const meta = getTargetMeta(target)
          const fmt = v => formatValue(v, meta)
          const rangeSpread = Math.abs(bullAbs - bearAbs)
          const sigmaVal = model.ss?.[validF.indexOf(topF)] || 1
          return (
            <>
              {/* Scenario range bar */}
              <div style={{ padding: '16px', background: C.bg2, borderRadius: 12, marginBottom: 14 }}>
                <div style={{ fontSize: 10, color: C.text3, letterSpacing: '0.1em', textTransform: 'uppercase', marginBottom: 12 }}>
                  Scenario range — {target}
                </div>
                {/* Range bar */}
                <div style={{ position: 'relative', height: 48, marginBottom: 10 }}>
                  <div style={{ position: 'absolute', top: '50%', left: 0, right: 0, height: 6, background: C.bg3, borderRadius: 3, transform: 'translateY(-50%)' }} />
                  <div style={{ position: 'absolute', top: '50%', left: '8%', right: '8%', height: 6, background: `linear-gradient(90deg, ${C.red}88, ${C.green}88)`, borderRadius: 3, transform: 'translateY(-50%)' }} />
                  {/* Bear marker */}
                  <div style={{ position: 'absolute', left: '8%', top: 0, transform: 'translateX(-50%)', textAlign: 'center' }}>
                    <div style={{ width: 12, height: 12, borderRadius: '50%', background: C.red, margin: '0 auto 2px', border: `2px solid ${C.bg}` }} />
                    <div style={{ fontSize: 9, color: C.red, fontWeight: 700, whiteSpace: 'nowrap' }}>BEAR</div>
                    <div style={{ fontSize: 11, color: C.red, fontWeight: 700 }}>{fmt(bearAbs)}</div>
                  </div>
                  {/* Base marker */}
                  <div style={{ position: 'absolute', left: '50%', top: 0, transform: 'translateX(-50%)', textAlign: 'center' }}>
                    <div style={{ width: 14, height: 14, borderRadius: '50%', background: C.green, margin: '0 auto 2px', border: `2px solid ${C.bg}` }} />
                    <div style={{ fontSize: 9, color: C.green, fontWeight: 700, whiteSpace: 'nowrap' }}>BASE</div>
                    <div style={{ fontSize: 11, color: C.green, fontWeight: 700 }}>{fmt(baseAbs)}</div>
                  </div>
                  {/* Bull marker */}
                  <div style={{ position: 'absolute', right: '8%', top: 0, transform: 'translateX(50%)', textAlign: 'center' }}>
                    <div style={{ width: 12, height: 12, borderRadius: '50%', background: C.blue, margin: '0 auto 2px', border: `2px solid ${C.bg}` }} />
                    <div style={{ fontSize: 9, color: C.blue, fontWeight: 700, whiteSpace: 'nowrap' }}>BULL</div>
                    <div style={{ fontSize: 11, color: C.blue, fontWeight: 700 }}>{fmt(bullAbs)}</div>
                  </div>
                </div>
                <div style={{ fontSize: 10, color: C.text3, marginTop: 4 }}>
                  Range spread: {fmt(bearAbs)} → {fmt(bullAbs)} (total {fmt(Math.abs(bullAbs - bearAbs))} gap)
                </div>
              </div>

              {/* What drives the scenarios */}
              <div style={{ padding: '12px 14px', background: C.bg2, borderRadius: 10, fontSize: 11, color: C.text2, lineHeight: 1.8, marginBottom: 12 }}>
                <strong style={{ color: C.text }}>What changes:</strong> Scenarios shift <em style={{ color: C.text }}>{topF}</em> (the variable with the highest impact on {target}) by one standard deviation (±{sigmaVal.toFixed(2)} units). Bear = worse fundamentals, Bull = better fundamentals.
                {useRelative && <span style={{ color: C.text3 }}> Values shown as absolute {target} (relative model + sector median).</span>}
              </div>

              {/* Currently at indicator */}
              <div style={{ padding: '10px 14px', background: C.blueDim, border: `1px solid ${C.blue}22`, borderRadius: 10, fontSize: 12, color: C.text2 }}>
                Currently trading at: <strong style={{ color: C.blue }}>{fmt(lastRow[target])}</strong>
                {' '}· Base case implies <strong style={{ color: baseAbs > lastRow[target] ? C.green : C.red }}>
                  {baseAbs > lastRow[target] ? '+' : ''}{Math.abs(lastRow[target]) > 0.1 ? (((baseAbs - lastRow[target]) / Math.abs(lastRow[target])) * 100).toFixed(0) + '%' : 'n/a'}
                </strong> move to base
              </div>
            </>
          )
        })()}
      </div>
    </div>
  )
}

// ─── Tab: Technical ───────────────────────────────────────────────────────────
function TabTechnical({ result, method, target }) {
  const C = useTheme()
  const meta = getTargetMeta(target)
  const ml = { ols: 'Pooled OLS', ridge: 'Ridge (α=5)', fe: 'Fixed Effects OLS', fe_ridge: 'FE + Ridge (α=5)' }
  if (!result) return null
  return (
    <>
      <div style={{ padding: '10px 16px', background: C.bg2, border: `1px solid ${C.border}`, borderRadius: 10, fontSize: 12, color: C.text2, marginBottom: 14, display: 'flex', gap: 24, flexWrap: 'wrap' }}>
        <span>Method: <strong style={{ color: C.blue }}>{ml[method]}</strong></span>
        <span>In-sample R²: <strong style={{ color: C.green }}>{result.r2?.toFixed(3)}</strong></span>
        <span>Adj R²: <strong>{result.r2Adj?.toFixed(3) || '—'}</strong></span>
        <span>SE: <strong>{result.se?.toFixed(3)}x</strong></span>
        <span>n = <strong>{result.n}</strong></span>
      </div>
      <div style={{ background: C.bg1, border: `1px solid ${C.border}`, borderRadius: 16, padding: '20px 22px', marginBottom: 14 }}>
        <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 12, color: C.text }}>Coefficient table</div>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead><tr style={{ borderBottom: `1px solid ${C.border}` }}>
              {['Variable', 'β', 'Std Error', 't-stat', 'p-value', ''].map(h => <th key={h} style={{ padding: '8px 12px', textAlign: 'left', color: C.text3, fontSize: 9, letterSpacing: '0.1em', textTransform: 'uppercase', fontWeight: 400 }}>{h}</th>)}
            </tr></thead>
            <tbody>{result.beta?.map((b, i) => {
              const sig = result.pV?.[i] < 0.05
              return <tr key={i} style={{ borderBottom: `1px solid ${C.grid}` }}>
                <td style={{ padding: '9px 12px', color: sig ? C.green : C.text, fontWeight: sig ? 600 : 400, fontFamily: 'var(--font)', fontSize: 12 }}>{result.labels?.[i]}</td>
                <td style={{ padding: '9px 12px', color: b > 0 ? C.green : C.red, fontWeight: 600, fontFamily: 'var(--mono)', fontSize: 11 }}>{b > 0 ? '+' : ''}{b.toFixed(5)}</td>
                <td style={{ padding: '9px 12px', color: C.text2, fontFamily: 'var(--mono)', fontSize: 11 }}>{result.bSE?.[i]?.toFixed(5) || '—'}</td>
                <td style={{ padding: '9px 12px', fontFamily: 'var(--mono)', fontSize: 11 }}>{result.tS?.[i]?.toFixed(3) || '—'}</td>
                <td style={{ padding: '9px 12px', color: result.pV?.[i] < 0.01 ? C.green : result.pV?.[i] < 0.05 ? C.amber : C.text3, fontFamily: 'var(--mono)', fontSize: 11 }}>{result.pV?.[i] < 0.0001 ? '<0.0001' : result.pV?.[i]?.toFixed(4) || '—'}</td>
                <td style={{ padding: '9px 12px' }}><span style={{ padding: '2px 8px', borderRadius: 4, fontSize: 9, fontWeight: 700, background: sig ? C.greenDim : C.redDim, color: sig ? C.green : C.red, border: `1px solid ${sig ? C.green + '33' : C.red + '33'}` }}>{sig ? '● SIG' : '○'}</span></td>
              </tr>
            })}</tbody>
          </table>
        </div>
      </div>
      <ChartCard title="Residual plot" subtitle="Points should be randomly scattered around zero. Patterns suggest model misspecification.">
        <ScatterChart margin={{ top: 4, right: 16, left: 0, bottom: 22 }}>
          <CartesianGrid stroke={C.grid} strokeDasharray="3 3" />
          <XAxis dataKey="predicted" stroke={C.text3} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }} label={{ value: 'Predicted', position: 'insideBottom', offset: -14, fill: C.text3, fontSize: 10 }} tickFormatter={v => formatAxisTick(v, meta)} />
          <YAxis dataKey="residual" stroke={C.text3} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }} />
          <Tooltip contentStyle={{ background: C.bg2, border: `1px solid ${C.border}`, borderRadius: 10, fontSize: 11 }} content={({ active, payload }) => { if (!active || !payload?.[0]) return null; const d = payload[0].payload; return <TipBox><div style={{ fontWeight: 700, color: d.isClient ? C.amber : C.text }}>{d.isClient ? '★ ' : ''}{d.company} · {d.year}</div><div style={{ color: C.text2, marginTop: 3 }}>Residual: {d.residual?.toFixed(2)}x</div></TipBox> }} />
          <ReferenceLine y={0} stroke={C.green + '66'} strokeDasharray="4 2" />
          <Scatter data={result.chartData?.filter(d => !d.isClient)} fill={C.blue} opacity={0.5} />
          <Scatter data={result.chartData?.filter(d => d.isClient)} fill={C.amber} />
        </ScatterChart>
      </ChartCard>
    </>
  )
}

// ─── AI Chatbot ───────────────────────────────────────────────────────────────
function buildSystemPrompt(config, result) {
  const { target, useRelative, clientCompany, features, method, selectedCos } = config
  const validF = features?.filter(f => f !== target) || []
  const topDrivers = result?.featureImp?.slice(0, 3).map(f => `${f.name} (${f.sig ? 'significant' : 'weak'})`).join(', ') || 'unknown'
  const r2 = result?.r2?.toFixed(3) || 'unknown'
  const se = result?.se?.toFixed(2) || 'unknown'
  const n = result?.n || 'unknown'
  const methodLabels = { ols: 'Pooled OLS', ridge: 'Ridge (α=5)', fe: 'Fixed Effects OLS', fe_ridge: 'Fixed Effects + Ridge' }

  return `You are a senior M&A analyst embedded inside a regression-based valuation tool called ValuationEngine. You help deal teams interpret statistical results in plain English.

CURRENT MODEL CONTEXT:
- Client company: ${clientCompany}
- Target variable: ${target}${useRelative ? ' (relative to sector median — removing macro noise)' : ' (absolute)'}
- Method: ${methodLabels[method] || method}
- Features used: ${validF.join(', ')}
- Comparables: ${selectedCos?.length || 'all'} companies
- Model fit: R² = ${r2} (${parseFloat(r2) > 0.55 ? 'strong' : parseFloat(r2) > 0.35 ? 'moderate' : 'limited'} explanatory power)
- Typical prediction error: ±${se}x
- Training observations: ${n}
- Top drivers: ${topDrivers}

YOUR ROLE:
- Translate statistics into deal language. No jargon without explanation.
- Be direct and honest about model limitations. Never overstate precision.
- Answer questions about why a company is valued the way it is, what drives premiums, what the range means, how to use this in a pitch.
- When asked about a specific company, reference what you know from the model context.
- Keep responses concise — this is a fast-moving deal environment.
- If asked something outside your context (e.g. specific news), say you don't have that data.

IMPORTANT: This is a private internal tool. All data stays in this session. Be candid and specific.`
}

function AIChatbot({ config, result, isLight }) {
  const C = isLight ? LIGHT : DARK
  const [open, setOpen] = useState(false)
  const [messages, setMessages] = useState([])
  const [input, setInput] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState(null)
  const messagesEndRef = useRef(null)
  const inputRef = useRef(null)

  const systemPrompt = useMemo(() => buildSystemPrompt(config, result), [config, result])

  const scrollToBottom = () => messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })

  useEffect(() => { if (open) { scrollToBottom(); inputRef.current?.focus() } }, [messages, open])

  const send = async () => {
    const text = input.trim()
    if (!text || loading) return
    setInput('')
    setError(null)
    const newMsg = { role: 'user', content: text }
    const updated = [...messages, newMsg]
    setMessages(updated)
    setLoading(true)
    try {
      const res = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: 'claude-sonnet-4-20250514',
          max_tokens: 1024,
          system: systemPrompt,
          messages: updated.map(m => ({ role: m.role, content: m.content }))
        })
      })
      const data = await res.json()
      if (data.error) throw new Error(data.error.message)
      const reply = data.content?.[0]?.text || '(no response)'
      setMessages(prev => [...prev, { role: 'assistant', content: reply }])
    } catch (e) {
      setError(e.message)
    }
    setLoading(false)
  }

  const suggestions = [
    'What does this valuation range mean in practice?',
    'Why is my client trading at a discount to peers?',
    'How confident should I be in these results?',
    'What are the strongest drivers of the multiple?',
  ]

  const panelBg = isLight ? '#FFFFFF' : '#0d1117'
  const borderCol = isLight ? 'rgba(96,0,29,0.12)' : 'rgba(255,255,255,0.08)'
  const accentCol = isLight ? 'rgb(24,78,98)' : '#3b82f6'

  return (
    <>
      {/* Floating button */}
      <button
        onClick={() => setOpen(v => !v)}
        style={{
          position: 'fixed', bottom: 28, right: 28, zIndex: 500,
          width: 52, height: 52, borderRadius: '50%',
          background: accentCol, border: 'none',
          boxShadow: `0 4px 20px ${accentCol}55`,
          cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center',
          transition: 'all 0.2s ease', transform: open ? 'scale(0.9)' : 'scale(1)'
        }}
        title="Ask the AI analyst"
      >
        {open
          ? <svg width="18" height="18" viewBox="0 0 18 18" fill="none"><path d="M2 2l14 14M16 2L2 16" stroke="white" strokeWidth="2.5" strokeLinecap="round"/></svg>
          : <svg width="20" height="20" viewBox="0 0 20 20" fill="none"><path d="M18 10c0 4.418-3.582 8-8 8a7.97 7.97 0 01-4-.01L2 19l1.01-4A8 8 0 1118 10z" stroke="white" strokeWidth="1.8" fill="none" strokeLinejoin="round"/><circle cx="7" cy="10" r="1" fill="white"/><circle cx="10" cy="10" r="1" fill="white"/><circle cx="13" cy="10" r="1" fill="white"/></svg>}
      </button>

      {/* Chat panel */}
      {open && (
        <div style={{
          position: 'fixed', bottom: 90, right: 28, zIndex: 499,
          width: 380, height: 540, display: 'flex', flexDirection: 'column',
          background: panelBg, border: `1px solid ${borderCol}`,
          borderRadius: 20, boxShadow: '0 24px 80px rgba(0,0,0,0.4)',
          animation: 'fadeUp 0.2s ease', overflow: 'hidden'
        }}>
          {/* Header */}
          <div style={{ padding: '14px 18px', borderBottom: `1px solid ${borderCol}`, display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexShrink: 0 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <div style={{ width: 32, height: 32, borderRadius: '50%', background: accentCol, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                <span style={{ fontSize: 14 }}>✦</span>
              </div>
              <div>
                <div style={{ fontSize: 13, fontWeight: 700, color: C.text }}>AI Analyst</div>
                <div style={{ fontSize: 10, color: C.text3 }}>Knows your model · {config?.clientCompany?.split(',')[0]}</div>
              </div>
            </div>
            <button
              onClick={() => { setMessages([]); setError(null) }}
              title="Clear conversation"
              style={{ padding: '4px 10px', borderRadius: 7, border: `1px solid ${borderCol}`, background: 'transparent', color: C.text3, fontSize: 10, cursor: 'pointer', fontFamily: 'var(--font)', transition: 'all 0.15s' }}
            >
              Clear ↺
            </button>
          </div>

          {/* Messages */}
          <div style={{ flex: 1, overflowY: 'auto', padding: '14px 16px', display: 'flex', flexDirection: 'column', gap: 12 }}>
            {messages.length === 0 && (
              <div>
                <div style={{ fontSize: 12, color: C.text2, lineHeight: 1.7, marginBottom: 14 }}>
                  Ask me anything about <strong style={{ color: C.text }}>{config?.target}</strong> or the valuation results for <strong style={{ color: accentCol }}>{config?.clientCompany?.split(',')[0]}</strong>. I have full context on the model.
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                  {suggestions.map(s => (
                    <button key={s} onClick={() => { setInput(s); inputRef.current?.focus() }}
                      style={{ textAlign: 'left', padding: '8px 12px', borderRadius: 9, border: `1px solid ${borderCol}`, background: 'transparent', color: C.text2, fontSize: 11, cursor: 'pointer', fontFamily: 'var(--font)', transition: 'all 0.15s', lineHeight: 1.4 }}
                      onMouseEnter={e => { e.currentTarget.style.borderColor = accentCol; e.currentTarget.style.color = C.text }}
                      onMouseLeave={e => { e.currentTarget.style.borderColor = borderCol; e.currentTarget.style.color = C.text2 }}>
                      {s}
                    </button>
                  ))}
                </div>
              </div>
            )}
            {messages.map((m, i) => (
              <div key={i} style={{ display: 'flex', justifyContent: m.role === 'user' ? 'flex-end' : 'flex-start' }}>
                <div style={{
                  maxWidth: '85%', padding: '10px 14px', borderRadius: m.role === 'user' ? '16px 16px 4px 16px' : '16px 16px 16px 4px',
                  background: m.role === 'user' ? accentCol : isLight ? '#F0EBE3' : '#1c2333',
                  color: m.role === 'user' ? 'white' : C.text,
                  fontSize: 12, lineHeight: 1.7, whiteSpace: 'pre-wrap'
                }}>
                  {m.content}
                </div>
              </div>
            ))}
            {loading && (
              <div style={{ display: 'flex', gap: 5, padding: '8px 4px' }}>
                {[0, 0.15, 0.3].map((d, i) => (
                  <div key={i} style={{ width: 7, height: 7, borderRadius: '50%', background: accentCol, animation: 'spin 1s ease infinite', animationDelay: `${d}s`, opacity: 0.7 }} />
                ))}
              </div>
            )}
            {error && <div style={{ fontSize: 11, color: '#f43f5e', padding: '8px 12px', background: 'rgba(244,63,94,0.1)', borderRadius: 8 }}>Error: {error}</div>}
            <div ref={messagesEndRef} />
          </div>

          {/* Input */}
          <div style={{ padding: '10px 14px', borderTop: `1px solid ${borderCol}`, display: 'flex', gap: 8, flexShrink: 0 }}>
            <input
              ref={inputRef}
              value={input}
              onChange={e => setInput(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && !e.shiftKey && send()}
              placeholder="Ask about the results…"
              style={{
                flex: 1, padding: '9px 12px', borderRadius: 10,
                background: isLight ? '#F0EBE3' : '#1c2333',
                border: `1px solid ${borderCol}`,
                color: C.text, fontSize: 12, fontFamily: 'var(--font)', outline: 'none'
              }}
            />
            <button onClick={send} disabled={loading || !input.trim()}
              style={{
                width: 36, height: 36, borderRadius: 10, border: 'none',
                background: loading || !input.trim() ? (isLight ? '#ddd' : '#1c2333') : accentCol,
                cursor: loading || !input.trim() ? 'not-allowed' : 'pointer',
                display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, transition: 'all 0.15s'
              }}>
              <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M1 7h12M8 3l5 4-5 4" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></svg>
            </button>
          </div>
        </div>
      )}
    </>
  )
}

// ─── Share Analysis — Apollo-style PDF report ─────────────────────────────────
function buildReportHTML({ config, result, allRows }) {
  const { target, useRelative, clientCompany, features, method, selectedCos } = config
  const validF = features?.filter(f => f !== target) || []
  const methodLabels = { ols: 'Pooled OLS', ridge: 'Ridge Regression', fe: 'Fixed Effects OLS', fe_ridge: 'Fixed Effects + Ridge' }
  const clientShort = clientCompany?.split(',')[0] || 'Client'

  // Build smart range
  const meta = getTargetMeta(target)
  const fmtVal = v => formatValue(v, meta)
  const model = buildModel(allRows, target, features, config.filterOutliers, method, useRelative)
  let absPred = null, currentActual = null, latestYear = null, sectorMedian = null, range = null
  if (model) {
    const clientRows = allRows.filter(r => r._company === clientCompany && isFinite(r[target]) && validF.every(f => isFinite(r[f]))).sort((a, b) => b._year - a._year)
    if (clientRows.length) {
      const latestRow = clientRows[0]
      const rawPred = predict({ ...latestRow, _company: clientCompany }, model)
      const yearVals = allRows.filter(r => r._year === latestRow._year && isFinite(r[target])).map(r => r[target]).sort((a, b) => a - b)
      sectorMedian = yearVals.length ? yearVals[Math.floor(yearVals.length / 2)] : null
      absPred = useRelative && isFinite(sectorMedian) ? rawPred + sectorMedian : rawPred
      currentActual = latestRow[target]
      latestYear = latestRow._year
      range = computeSmartRange(absPred, target, allRows)
    }
  }

  // Build all companies ranking
  const companyRanking = []
  if (model) {
    const companies = [...new Set(allRows.map(r => r._company).filter(Boolean))]
    for (const co of companies) {
      const latest = allRows.filter(r => r._company === co && isFinite(r[target]) && validF.every(f => isFinite(r[f]))).sort((a, b) => b._year - a._year)[0]
      if (!latest) continue
      const yv = allRows.filter(r => r._year === latest._year && isFinite(r[target])).map(r => r[target]).sort((a, b) => a - b)
      const sm = yv.length ? yv[Math.floor(yv.length / 2)] : 0
      const rawP = predict({ ...latest, _company: co }, model)
      const absP = useRelative ? rawP + sm : rawP
      const actual = latest[target]
      const denom = Math.max(Math.abs(actual), 0.5)
      companyRanking.push({ co: compShort(co), actual, pred: absP, gap: absP - actual, upsidePct: ((absP - actual) / denom) * 100, isClient: co === clientCompany })
    }
    companyRanking.sort((a, b) => b.upsidePct - a.upsidePct)
  }

  const topDrivers = result?.featureImp?.filter(f => f.sig).slice(0, 5) || []
  const allDrivers = result?.featureImp || []
  const gap = isFinite(absPred) && isFinite(currentActual) ? absPred - currentActual : 0
  const upsidePct = isFinite(currentActual) && Math.abs(currentActual) > 0.1 ? ((absPred - currentActual) / Math.abs(currentActual)) * 100 : 0
  const clientRank = companyRanking.findIndex(r => r.isClient) + 1
  const r2 = result?.r2 || 0
  const confidence = r2 > 0.55 ? 'High' : r2 > 0.35 ? 'Moderate' : 'Limited'

  // Investment thesis paragraph
  const direction = gap > 0 ? 'undervalued' : 'overvalued'
  const directionColor = gap > 0 ? '#0a6b3f' : '#9b1c2d'
  const thesisText = isFinite(upsidePct)
    ? `Based on ${result?.n || 0} comparable observations across ${selectedCos?.length || 0} companies, our regression model estimates ${clientShort} should trade in the ${fmtVal(range?.lo)}–${fmtVal(range?.hi)} range for ${target}. Currently at ${fmtVal(currentActual)}, this implies a directional ${upsidePct > 0 ? '+' : ''}${upsidePct.toFixed(0)}% re-rating opportunity to the midpoint of the range. ${clientShort} ranks ${clientRank} of ${companyRanking.length} comparables when sorted by model-implied upside. ${confidence === 'High' ? 'Model confidence is strong, with R² of ' + r2.toFixed(2) + ' indicating the fundamentals captured here explain most of the cross-company variation in valuation.' : confidence === 'Moderate' ? 'Model confidence is moderate (R² = ' + r2.toFixed(2) + '); supplement these findings with qualitative analysis and other valuation methods.' : 'Model confidence is limited (R² = ' + r2.toFixed(2) + '); treat these results as directional only.'}`
    : 'Insufficient data to generate thesis.'

  const today = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Valuation Analysis — ${clientShort}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,300;8..60,400;8..60,600;8..60,700&family=Inter:wght@300;400;500;600;700&display=swap');
  * { margin: 0; padding: 0; box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  body { font-family: 'Inter', -apple-system, sans-serif; color: #1a1a1a; background: #fff; line-height: 1.6; font-size: 11pt; }
  .page { max-width: 8.5in; margin: 0 auto; padding: 0.7in 0.85in; min-height: 11in; position: relative; page-break-after: always; }
  .page:last-child { page-break-after: auto; }
  .page-num { position: absolute; bottom: 0.4in; right: 0.85in; font-size: 9pt; color: #999; }
  .page-footer { position: absolute; bottom: 0.4in; left: 0.85in; font-size: 9pt; color: #999; letter-spacing: 0.08em; text-transform: uppercase; }

  h1, h2, h3 { font-family: 'Source Serif 4', Georgia, serif; font-weight: 600; letter-spacing: -0.01em; }
  h1 { font-size: 32pt; line-height: 1.1; margin-bottom: 16px; color: #0a1929; }
  h2 { font-size: 18pt; margin-bottom: 14px; color: #0a1929; padding-bottom: 8px; border-bottom: 2px solid #0a1929; }
  h3 { font-size: 13pt; margin: 22px 0 10px; color: #0a1929; }
  p { margin-bottom: 12px; }

  .firm-mark { font-size: 9pt; letter-spacing: 0.25em; text-transform: uppercase; color: #0a1929; font-weight: 700; margin-bottom: 12px; }
  .doc-type { font-size: 10pt; letter-spacing: 0.18em; text-transform: uppercase; color: #6b6b6b; margin-bottom: 80px; font-weight: 500; }
  .cover-meta { position: absolute; bottom: 1in; left: 0.85in; right: 0.85in; padding-top: 18px; border-top: 1px solid #ccc; display: flex; justify-content: space-between; font-size: 9pt; color: #666; }
  .cover-rule { width: 60px; height: 4px; background: #0a1929; margin: 28px 0; }
  .cover-subtitle { font-size: 13pt; color: #4a4a4a; font-weight: 300; line-height: 1.5; max-width: 5in; margin-top: 18px; }
  .cover-confidential { display: inline-block; padding: 4px 12px; background: #fde8ec; color: #9b1c2d; font-size: 8pt; font-weight: 600; letter-spacing: 0.15em; text-transform: uppercase; border-radius: 3px; margin-top: 40px; }

  /* Range block */
  .range-block { padding: 24px 28px; background: #f8f6f2; border-left: 4px solid #0a1929; margin: 16px 0 24px; }
  .range-label { font-size: 9pt; letter-spacing: 0.18em; text-transform: uppercase; color: #6b6b6b; margin-bottom: 8px; font-weight: 600; }
  .range-value { font-family: 'Source Serif 4', Georgia, serif; font-size: 36pt; font-weight: 700; color: #0a1929; line-height: 1; margin-bottom: 10px; letter-spacing: -0.02em; }
  .range-detail { font-size: 11pt; color: #4a4a4a; line-height: 1.7; }
  .range-detail strong { color: #0a1929; }
  .range-direction { display: inline-block; padding: 3px 11px; border-radius: 4px; font-size: 9pt; font-weight: 600; letter-spacing: 0.06em; text-transform: uppercase; margin-left: 6px; }

  /* KPI grid */
  .kpi-grid { display: grid; grid-template-columns: 1fr 1fr 1fr 1fr; gap: 12px; margin: 18px 0; }
  .kpi-card { padding: 14px 16px; border: 1px solid #e5e0d8; border-radius: 4px; }
  .kpi-label { font-size: 8pt; letter-spacing: 0.14em; text-transform: uppercase; color: #888; margin-bottom: 6px; font-weight: 600; }
  .kpi-value { font-family: 'Source Serif 4', Georgia, serif; font-size: 22pt; font-weight: 700; color: #0a1929; line-height: 1; }
  .kpi-sub { font-size: 9pt; color: #999; margin-top: 4px; }

  /* Thesis */
  .thesis { padding: 20px 24px; background: #fff; border: 1px solid #e0dcd2; border-radius: 4px; margin: 12px 0 20px; }
  .thesis p { font-size: 11.5pt; line-height: 1.8; color: #2a2a2a; }
  .thesis strong { color: #0a1929; }

  /* Tables */
  table { width: 100%; border-collapse: collapse; margin: 14px 0; font-size: 10pt; }
  th { text-align: left; padding: 10px 12px; background: #0a1929; color: white; font-weight: 600; font-size: 9pt; letter-spacing: 0.06em; text-transform: uppercase; }
  td { padding: 9px 12px; border-bottom: 1px solid #e8e3d8; }
  tr:last-child td { border-bottom: none; }
  .num { font-family: 'JetBrains Mono', monospace; font-size: 10pt; text-align: right; }
  .sig { display: inline-block; padding: 2px 7px; border-radius: 3px; font-size: 8pt; font-weight: 700; }
  .sig-yes { background: #d4f4dd; color: #0a6b3f; }
  .sig-no { background: #fde8ec; color: #9b1c2d; }
  .row-client { background: #fff8e6; font-weight: 600; }

  /* Caveat */
  .caveat { padding: 12px 16px; background: #fdf4d6; border-left: 3px solid #d4a814; font-size: 9.5pt; color: #5a4a14; margin: 18px 0; line-height: 1.7; border-radius: 0 4px 4px 0; }
  .caveat strong { color: #4a3a08; }

  /* Bar chart */
  .bar-row { display: grid; grid-template-columns: 1.5fr 3fr 0.6fr; align-items: center; gap: 10px; padding: 5px 0; font-size: 9.5pt; }
  .bar-row.client { font-weight: 700; }
  .bar-name { color: #2a2a2a; }
  .bar-track { position: relative; height: 18px; background: #f0ece4; border-radius: 2px; overflow: hidden; }
  .bar-track .zero-line { position: absolute; left: 50%; top: 0; bottom: 0; width: 1px; background: #999; }
  .bar-fill { position: absolute; top: 0; bottom: 0; }
  .bar-pct { font-family: 'JetBrains Mono', monospace; font-size: 9pt; text-align: right; }

  /* Drivers section */
  .driver-row { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #eee; font-size: 10pt; }
  .driver-row:last-child { border-bottom: none; }
  .driver-name { flex: 1; }
  .driver-effect { color: #6b6b6b; font-size: 9.5pt; }

  /* Print rules */
  @media print {
    body { font-size: 10pt; }
    .page { padding: 0.6in 0.75in; }
    .no-print { display: none; }
    h2 { page-break-after: avoid; }
  }
  @page { size: letter; margin: 0; }

  /* Print button (hidden when printing) */
  .print-bar { position: fixed; top: 16px; right: 16px; z-index: 1000; display: flex; gap: 10px; }
  .print-btn { padding: 10px 18px; border-radius: 8px; border: none; background: #0a1929; color: white; font-family: 'Inter', sans-serif; font-size: 12px; font-weight: 600; cursor: pointer; box-shadow: 0 4px 14px rgba(0,0,0,0.2); }
  .print-btn.secondary { background: white; color: #0a1929; border: 1px solid #0a1929; }
</style>
</head>
<body>
  <div class="print-bar no-print">
    <button class="print-btn secondary" onclick="window.close()">Close</button>
    <button class="print-btn" onclick="window.print()">Print / Save as PDF →</button>
  </div>

  <!-- COVER PAGE -->
  <div class="page">
    <div class="firm-mark">TierOne M&A · Valuation Engine</div>
    <div class="doc-type">Comparables Analysis Outlook</div>

    <h1>${clientShort}<br/>Valuation Outlook</h1>
    <div class="cover-rule"></div>
    <div class="cover-subtitle">${useRelative ? 'Relative-to-sector regression analysis' : 'Regression-based comparable analysis'} of ${target}, anchored on ${latestYear || 'latest'} fundamentals across ${selectedCos?.length || 0} peer companies.</div>

    <div class="cover-confidential">For internal use · Confidential</div>

    <div class="cover-meta">
      <div>
        <strong style="color: #0a1929;">Prepared by</strong><br/>
        TierOne M&A Advisory<br/>
        Valuation Engine — automated analysis
      </div>
      <div style="text-align: right;">
        <strong style="color: #0a1929;">Date</strong><br/>
        ${today}<br/>
        Document version 1.0
      </div>
    </div>
    <div class="page-num">1</div>
  </div>

  <!-- EXECUTIVE SUMMARY -->
  <div class="page">
    <h2>Executive Summary</h2>

    <h3 style="margin-top: 0;">Model-implied valuation range</h3>
    <div class="range-block">
      <div class="range-label">${target} · Based on ${latestYear} fundamentals</div>
      <div class="range-value">${fmtVal(range?.lo)} – ${fmtVal(range?.hi)}</div>
      <div class="range-detail">
        ${clientShort} is currently trading at <strong>${fmtVal(currentActual)}</strong>, implying
        <span class="range-direction" style="background: ${gap > 0 ? '#d4f4dd' : '#fde8ec'}; color: ${directionColor};">${isFinite(upsidePct) ? (upsidePct > 0 ? '+' : '') + upsidePct.toFixed(0) + '% ' + direction : 'within range'}</span>
        relative to the midpoint of our model's fair value range.
      </div>
    </div>

    <h3>Key metrics</h3>
    <div class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-label">Confidence</div>
        <div class="kpi-value">${confidence}</div>
        <div class="kpi-sub">R² = ${r2.toFixed(2)}</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Peer rank</div>
        <div class="kpi-value">#${clientRank}</div>
        <div class="kpi-sub">of ${companyRanking.length} by upside</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Strong drivers</div>
        <div class="kpi-value">${topDrivers.length}</div>
        <div class="kpi-sub">of ${allDrivers.length} tested</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Sector median</div>
        <div class="kpi-value">${fmtVal(sectorMedian)}</div>
        <div class="kpi-sub">${latestYear} cohort</div>
      </div>
    </div>

    <h3>Investment thesis</h3>
    <div class="thesis">
      <p>${thesisText}</p>
    </div>

    <h3>Methodology snapshot</h3>
    <table>
      <tr><td style="width: 35%; color: #6b6b6b;">Target variable</td><td><strong>${target}</strong>${useRelative ? ' <em>(relative to sector median)</em>' : ''}</td></tr>
      <tr><td style="color: #6b6b6b;">Regression method</td><td>${methodLabels[method] || method}</td></tr>
      <tr><td style="color: #6b6b6b;">Independent variables</td><td>${validF.join(' · ')}</td></tr>
      <tr><td style="color: #6b6b6b;">Sample size</td><td>${result?.n || 0} company-year observations</td></tr>
      <tr><td style="color: #6b6b6b;">Coverage</td><td>${selectedCos?.length || 0} comparable companies</td></tr>
    </table>

    <div class="caveat"><strong>Important:</strong> This analysis is a directional signal. Treat all figures as analytical inputs to support a broader valuation narrative — not as precise transaction values. Combine with DCF, precedent transaction analysis, and qualitative due diligence.</div>

    <div class="page-footer">${clientShort} · Valuation Outlook</div>
    <div class="page-num">2</div>
  </div>

  <!-- ANALYSIS DETAIL -->
  <div class="page">
    <h2>Analysis Detail</h2>

    <h3>What drives the multiple</h3>
    <p style="font-size: 10.5pt; color: #4a4a4a;">Of the ${allDrivers.length} variables tested, ${topDrivers.length} show statistically significant influence on ${target} across the peer set. Variables marked significant have <em>p &lt; 0.05</em> — meaning we are at least 95% confident the relationship is real.</p>

    <table>
      <thead>
        <tr>
          <th>Variable</th>
          <th style="text-align: right;">Effect direction</th>
          <th style="text-align: right;">Confidence</th>
          <th style="text-align: right;">Status</th>
        </tr>
      </thead>
      <tbody>
        ${allDrivers.slice(0, 8).map(d => `
          <tr>
            <td>${d.name}</td>
            <td class="num" style="color: ${d.beta > 0 ? '#0a6b3f' : '#9b1c2d'};">${d.beta > 0 ? 'Higher → higher multiple' : 'Higher → lower multiple'}</td>
            <td class="num">${d.pVal < 0.01 ? 'High' : d.pVal < 0.05 ? 'Medium' : 'Low'}</td>
            <td class="num"><span class="sig ${d.sig ? 'sig-yes' : 'sig-no'}">${d.sig ? '● SIG' : '○ WEAK'}</span></td>
          </tr>
        `).join('')}
      </tbody>
    </table>

    <h3>Peer ranking — model-implied upside</h3>
    <p style="font-size: 10.5pt; color: #4a4a4a;">Companies ranked by the gap between current valuation and model fair value. Positive = market trading below model estimate (potential re-rating).</p>

    <div style="margin: 14px 0;">
      ${companyRanking.slice(0, 14).map(d => {
        const w = Math.min(Math.abs(d.upsidePct), 100)
        const isClientCls = d.isClient ? ' client' : ''
        const fillCol = d.isClient ? '#d4a814' : d.gap > 0 ? '#5a9a72' : '#c47a82'
        const leftPos = d.upsidePct > 0 ? 50 : 50 - w / 2
        return `
        <div class="bar-row${isClientCls}">
          <div class="bar-name">${d.isClient ? '★ ' : ''}${d.co}</div>
          <div class="bar-track">
            <div class="zero-line"></div>
            <div class="bar-fill" style="left: ${leftPos}%; width: ${w / 2}%; background: ${fillCol};"></div>
          </div>
          <div class="bar-pct" style="color: ${d.gap > 0 ? '#0a6b3f' : '#9b1c2d'};">${d.upsidePct > 0 ? '+' : ''}${isFinite(d.upsidePct) ? d.upsidePct.toFixed(0) : '—'}%</div>
        </div>
        `
      }).join('')}
    </div>

    ${companyRanking.length > 14 ? `<p style="font-size: 9pt; color: #999; font-style: italic;">Showing top 14 of ${companyRanking.length} comparables. Full ranking available in the live tool.</p>` : ''}

    <div class="caveat" style="margin-top: 22px;"><strong>Reading this chart:</strong> Bars are model-implied upside relative to current market price. ${clientShort} (★) sits at rank #${clientRank}. Companies with the largest positive bars are those the model believes most undervalued; the largest negative bars are those most overvalued.</div>

    <div class="page-footer">${clientShort} · Valuation Outlook</div>
    <div class="page-num">3</div>
  </div>

  <!-- CONCLUSION -->
  <div class="page">
    <h2>Recommendation & Next Steps</h2>

    <h3>How to use this analysis</h3>
    <p>This regression-based valuation should be one input among several. The output gives a fundamentals-driven baseline against which to test:</p>
    <ul style="margin-left: 22px; margin-bottom: 18px;">
      <li style="margin-bottom: 8px;"><strong>Pitch narrative.</strong> Use the range and peer rank to anchor positioning vs. comparables in initial conversations.</li>
      <li style="margin-bottom: 8px;"><strong>Bid range setup.</strong> Reference the model midpoint as a credibility check against DCF-derived values.</li>
      <li style="margin-bottom: 8px;"><strong>Counterparty negotiation.</strong> Use the driver analysis to argue why ${clientShort} deserves a specific premium or discount vs. peers.</li>
      <li style="margin-bottom: 8px;"><strong>Sensitivity work.</strong> Pair this with scenarios where the top drivers move ±1σ to bracket bull/bear cases.</li>
    </ul>

    <h3>Model strengths</h3>
    <p style="margin-bottom: 6px;">${confidence === 'High' ? '<strong>Strong fit.</strong> The model captures most of the cross-company variation in ' + target + '. Findings can be referenced with confidence.' : confidence === 'Moderate' ? '<strong>Reasonable fit.</strong> The model explains a meaningful share of variation but leaves room for unmodeled factors. Use as supportive evidence rather than the primary anchor.' : '<strong>Limited fit.</strong> The model captures a small share of variation. Treat as a rough directional check only — supplement with strong qualitative arguments and other methods.'}</p>

    <h3>Model limitations</h3>
    <p>What this model does <strong>not</strong> capture:</p>
    <ul style="margin-left: 22px; margin-bottom: 18px;">
      <li style="margin-bottom: 6px;">Forward-looking strategic catalysts (new product cycles, management changes, M&A optionality)</li>
      <li style="margin-bottom: 6px;">Capital structure differences (leverage, buybacks, dividend policy)</li>
      <li style="margin-bottom: 6px;">Macro regime shifts (rate environment, sector rotation, sentiment)</li>
      <li style="margin-bottom: 6px;">Liquidity, float, and analyst-coverage effects</li>
    </ul>

    <div class="caveat" style="background: #f0f5f9; border-left-color: #1e4a6b; color: #1a2a3a;">
      <strong>Important disclosure:</strong> This document is generated by a regression-based valuation tool and is for internal use only. Figures are directional analytical outputs, not investment advice or transaction recommendations. All decisions should incorporate judgement, qualitative diligence, and other valuation approaches.
    </div>

    <div style="margin-top: 60px; padding-top: 18px; border-top: 1px solid #ddd; font-size: 9pt; color: #999; display: flex; justify-content: space-between;">
      <div>
        TierOne M&A Advisory<br/>
        Valuation Engine — automated comparables analysis
      </div>
      <div style="text-align: right;">
        Generated ${today}<br/>
        Confidential — for internal use only
      </div>
    </div>

    <div class="page-footer">${clientShort} · Valuation Outlook</div>
    <div class="page-num">4</div>
  </div>
</body>
</html>`
}

function ShareAnalysisButton({ config, result, allRows, isLight }) {
  const C = isLight ? LIGHT : DARK
  const [generating, setGenerating] = useState(false)
  const handle = () => {
    setGenerating(true)
    setTimeout(() => {
      try {
        const html = buildReportHTML({ config, result, allRows })
        const w = window.open('', '_blank')
        if (w) { w.document.write(html); w.document.close() }
      } catch (e) {
        alert('Error generating report: ' + e.message)
      }
      setGenerating(false)
    }, 100)
  }
  return (
    <button onClick={handle} disabled={generating}
      style={{
        display: 'flex', alignItems: 'center', gap: 7,
        padding: '7px 14px', borderRadius: 8,
        border: `1px solid ${C.amber}`, background: C.amberDim,
        color: C.amber, fontSize: 12, fontWeight: 600, cursor: generating ? 'wait' : 'pointer',
        fontFamily: 'var(--font)', transition: 'all 0.15s'
      }}>
      {generating
        ? <><div style={{ width: 12, height: 12, border: `2px solid ${C.amber}40`, borderTopColor: C.amber, borderRadius: '50%', animation: 'spin 0.7s linear infinite' }} />Generating…</>
        : <><svg width="13" height="13" viewBox="0 0 13 13" fill="none"><path d="M2 1h7l3 3v8H2V1z" stroke="currentColor" strokeWidth="1.4" strokeLinejoin="round" fill="none"/><path d="M9 1v3h3M4 7h5M4 9h5M4 11h3" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round"/></svg>Share analysis</>}
    </button>
  )
}

// ─── Theme Toggle Button ──────────────────────────────────────────────────────
function ThemeToggle({ isLight, onToggle }) {
  return (
    <button onClick={onToggle}
      title={isLight ? 'Switch to dark mode' : 'Switch to light mode'}
      style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '7px 12px', borderRadius: 8, border: `1px solid ${isLight ? 'rgba(96,0,29,0.2)' : DARK.border}`, background: isLight ? 'rgb(219,204,189)' : DARK.bg2, color: isLight ? 'rgb(96,0,29)' : DARK.text2, fontSize: 11, cursor: 'pointer', fontFamily: 'var(--font)', transition: 'all 0.2s', fontWeight: 500 }}>
      {isLight ? (
        <><svg width="12" height="12" viewBox="0 0 12 12" fill="none"><circle cx="6" cy="6" r="2.5" fill="currentColor"/><path d="M6 1v1M6 10v1M1 6h1M10 6h1M2.5 2.5l.7.7M8.8 8.8l.7.7M2.5 9.5l.7-.7M8.8 3.2l.7-.7" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round"/></svg>Dark mode</>
      ) : (
        <><svg width="12" height="12" viewBox="0 0 12 12" fill="none"><path d="M10 6.5A4.5 4.5 0 015.5 2a4.5 4.5 0 100 9A4.5 4.5 0 0010 6.5z" stroke="currentColor" strokeWidth="1.2" fill="none"/></svg>Light mode</>
      )}
    </button>
  )
}

// ─── Relative Mode: Positioning Map ──────────────────────────────────────────
// Shows each company's current premium vs model-implied premium
// Quadrant: X = actual premium, Y = model says it should be
// Quadrant I (top-right): high actual + high model = justified premium
// Quadrant II (top-left): low actual + high model = UNDERVALUED (buy signal)
// Quadrant III (bottom-left): low actual + low model = justified discount
// Quadrant IV (bottom-right): high actual + low model = OVERVALUED (sell signal)
function RelativePositioningMap({ allRows, target, features, clientCompany, filterOutliers, method }) {
  const C = useTheme()
  const model = useMemo(() => buildModel(allRows, target, features, filterOutliers, method, true), [allRows, target, features, filterOutliers, method])
  const validF = features.filter(f => f !== target)

  const data = useMemo(() => {
    if (!model) return []
    const companies = [...new Set(allRows.map(r => r._company).filter(Boolean))]
    const rows = []
    // Get latest year per company
    for (const co of companies) {
      const latest = allRows.filter(r => r._company === co && isFinite(r[target]) && validF.every(f => isFinite(r[f]))).sort((a, b) => b._year - a._year)[0]
      if (!latest) continue
      // Get sector median for this year
      const yearVals = allRows.filter(r => r._year === latest._year && isFinite(r[target])).map(r => r[target])
      const sorted = yearVals.slice().sort((a, b) => a - b)
      const sectorMed = sorted[Math.floor(sorted.length / 2)]
      const relActual = latest[target] - sectorMed  // how much above/below peers today
      const rawPred = predict({ ...latest, _company: co }, model)  // model's relative prediction
      rows.push({
        co, shortName: compShort(co),
        relActual: parseFloat(relActual.toFixed(2)),
        relPred: parseFloat(rawPred.toFixed(2)),
        gap: rawPred - relActual,
        isClient: co === clientCompany,
        year: latest._year
      })
    }
    return rows
  }, [model, allRows, target, validF, clientCompany])

  if (!model || !data.length) return <div style={{ padding: 32, color: C.text2, textAlign: 'center' }}>No data available.</div>

  const allVals = data.flatMap(d => [d.relActual, d.relPred]).filter(isFinite)
  const axisDom = getAxisDomain(allVals, 0.2)
  const zero = 0

  return (
    <ChartCard
      title="Relative positioning map — where each company stands vs where the model says it should"
      subtitle="X axis = actual premium/discount vs sector today · Y axis = model's estimate of deserved premium/discount"
      explain="This is the most important chart in relative mode. Each dot = one company's latest data. The X position shows where the market prices it relative to peers. The Y position shows where the model thinks it should be. Companies in the top-left are UNDERVALUED (deserve a premium but not getting one). Companies in the bottom-right are OVERVALUED. Your client is shown as a large orange dot. The closer to the diagonal, the more fairly priced."
      explainAdv="This is essentially a residual plot in sector-relative space. X = actual relative multiple (company - sector median). Y = model's predicted relative multiple. Diagonal = fair. Top-left quadrant: model predicts positive premium, market gives discount → potentially undervalued. Bottom-right: model predicts discount, market gives premium → potentially overvalued."
      height={360}>
      <ScatterChart margin={{ top: 20, right: 20, left: 20, bottom: 28 }}>
        {/* Quadrant background zones */}
        <defs>
          <pattern id="hatch-green" width="8" height="8" patternUnits="userSpaceOnUse" patternTransform="rotate(45)">
            <line x1="0" y1="0" x2="0" y2="8" stroke={C.green} strokeWidth="1" strokeOpacity="0.15" />
          </pattern>
          <pattern id="hatch-red" width="8" height="8" patternUnits="userSpaceOnUse" patternTransform="rotate(45)">
            <line x1="0" y1="0" x2="0" y2="8" stroke={C.red} strokeWidth="1" strokeOpacity="0.15" />
          </pattern>
        </defs>
        <CartesianGrid stroke={C.grid} strokeDasharray="3 3" />
        <XAxis dataKey="relActual" type="number" name="Actual premium" stroke={C.text3} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }}
          domain={axisDom} tickFormatter={v => `${v > 0 ? '+' : ''}${v.toFixed(1)}x`}
          label={{ value: '← Discount vs peers  |  Premium vs peers →', position: 'insideBottom', offset: -14, fill: C.text3, fontSize: 10 }} />
        <YAxis dataKey="relPred" type="number" name="Model says" stroke={C.text3} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }}
          domain={axisDom} tickFormatter={v => `${v > 0 ? '+' : ''}${v.toFixed(1)}x`}
          label={{ value: 'Model says should be', angle: -90, position: 'insideLeft', fill: C.text3, fontSize: 9, dx: -4 }} />
        <ZAxis range={[40, 40]} />
        <Tooltip content={({ active, payload }) => {
          if (!active || !payload?.[0]) return null
          const d = payload[0].payload
          const zone = d.relPred > 0 && d.relActual < 0 ? 'UNDERVALUED' :
                       d.relPred < 0 && d.relActual > 0 ? 'OVERVALUED' :
                       d.relPred > 0 && d.relActual > 0 ? 'PREMIUM (justified)' : 'DISCOUNT (justified)'
          const zoneCol = zone === 'UNDERVALUED' ? C.green : zone === 'OVERVALUED' ? C.red : C.amber
          return (
            <TipBox>
              <div style={{ fontWeight: 700, color: d.isClient ? C.amber : C.text, marginBottom: 6 }}>{d.isClient ? '★ ' : ''}{d.co?.split(',')[0]}</div>
              <div style={{ color: C.text2, marginBottom: 2 }}>Trades at: <strong style={{ color: C.blue }}>{d.relActual > 0 ? '+' : ''}{d.relActual?.toFixed(2)}x vs sector</strong></div>
              <div style={{ color: C.text2, marginBottom: 8 }}>Should trade at: <strong style={{ color: C.green }}>{d.relPred > 0 ? '+' : ''}{d.relPred?.toFixed(2)}x vs sector</strong></div>
              <div style={{ padding: '4px 8px', borderRadius: 5, background: zoneCol + '22', color: zoneCol, fontWeight: 700, fontSize: 11, display: 'inline-block', border: `1px solid ${zoneCol}44` }}>{zone}</div>
            </TipBox>
          )
        }} cursor={false} />
        {/* Diagonal = fair value line */}
        <ReferenceLine segment={[{ x: axisDom[0], y: axisDom[0] }, { x: axisDom[1], y: axisDom[1] }]} stroke={C.amber + '88'} strokeDasharray="6 3" label={{ value: 'Fair value', fill: C.amber, fontSize: 9, position: 'insideTopLeft' }} />
        {/* Zero lines */}
        <ReferenceLine x={zero} stroke={C.text3 + '66'} />
        <ReferenceLine y={zero} stroke={C.text3 + '66'} />
        {/* Quadrant labels */}
        <ReferenceLine x={axisDom[0] * 0.7} y={axisDom[1] * 0.7} stroke="none" label={{ value: 'UNDERVALUED ↑', fill: C.green, fontSize: 9, fontWeight: 700 }} />
        <ReferenceLine x={axisDom[1] * 0.5} y={axisDom[0] * 0.7} stroke="none" label={{ value: 'OVERVALUED ↓', fill: C.red, fontSize: 9, fontWeight: 700 }} />
        <Scatter data={data.filter(d => !d.isClient)} fill={C.blue} opacity={0.65} name="Comparables" />
        <Scatter data={data.filter(d => d.isClient)} fill={C.amber} name={compShort(clientCompany)} />
      </ScatterChart>
    </ChartCard>
  )
}

// ─── Relative Mode: Premium/Discount over time ───────────────────────────────
function RelativePremiumOverTime({ allRows, target, features, clientCompany, filterOutliers, method }) {
  const C = useTheme()
  const model = useMemo(() => buildModel(allRows, target, features, filterOutliers, method, true), [allRows, target, features, filterOutliers, method])
  const validF = features.filter(f => f !== target)

  const { chartData, years } = useMemo(() => {
    if (!model) return { chartData: [], years: [] }
    const yrs = [...new Set(allRows.map(r => r._year).filter(isFinite))].sort()
    // Build sector medians
    const sectorMeds = {}
    for (const yr of yrs) {
      const vals = allRows.filter(r => r._year === yr && isFinite(r[target])).map(r => r[target]).sort((a, b) => a - b)
      if (vals.length) sectorMeds[yr] = vals[Math.floor(vals.length / 2)]
    }
    // Client's relative actual and model prediction over time
    const clientRows = allRows.filter(r => r._company === clientCompany && isFinite(r[target]) && validF.every(f => isFinite(r[f]))).sort((a, b) => a._year - b._year)
    const data = clientRows.map(r => {
      const smed = sectorMeds[r._year] ?? 0
      const relActual = r[target] - smed
      const rawPred = predict({ ...r, _company: clientCompany }, model)
      return { year: r._year, relActual: parseFloat(relActual.toFixed(3)), relPred: parseFloat(rawPred.toFixed(3)), gap: parseFloat((rawPred - relActual).toFixed(3)), sectorMedian: parseFloat(smed.toFixed(3)) }
    })
    return { chartData: data, years: yrs }
  }, [model, allRows, target, validF, clientCompany])

  if (!chartData.length) return null

  const gapDomain = getAxisDomain(chartData.flatMap(d => [d.relActual, d.relPred]))

  const downloadData = chartData.map(d => ({ Year: d.year, 'Actual premium vs sector': d.relActual, 'Model says fair premium': d.relPred, 'Gap (model - actual)': d.gap, 'Sector median': d.sectorMedian }))

  return (
    <ChartCard
      title={`${compShort(clientCompany)} — actual vs deserved premium/discount over time`}
      subtitle="How much above or below sector median the company has traded, and what the model says it should have been"
      explain="The orange line shows how much above or below the sector median this company has traded each year (0 = exactly at sector median). The green dashed line shows what the model thinks the premium/discount should have been. When orange is below green, the company was undervalued relative to peers. When orange is above green, it was overvalued relative to peers."
      downloadData={downloadData}
      downloadName={`${compShort(clientCompany)}-relative-premium`}
      height={280}>
      <ComposedChart data={chartData} margin={{ top: 8, right: 20, left: 0, bottom: 8 }}>
        <defs>
          <linearGradient id="gapFill" x1="0" y1="0" x2="0" y2="1">
            <stop offset="5%" stopColor={C.green} stopOpacity={0.08} />
            <stop offset="95%" stopColor={C.green} stopOpacity={0.01} />
          </linearGradient>
        </defs>
        <CartesianGrid stroke={C.grid} strokeDasharray="3 3" />
        <XAxis dataKey="year" stroke={C.text3} tick={{ fontSize: 10 }} />
        <YAxis stroke={C.text3} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }} tickFormatter={v => `${v > 0 ? '+' : ''}${v.toFixed(1)}x`} domain={gapDomain} />
        <Tooltip contentStyle={{ background: C.bg2, border: `1px solid ${C.border}`, borderRadius: 10, fontSize: 11 }}
          content={({ active, payload, label }) => {
            if (!active || !payload?.length) return null
            const actual = payload.find(p => p.dataKey === 'relActual')
            const pred = payload.find(p => p.dataKey === 'relPred')
            if (!actual || !pred) return null
            const gap = pred.value - actual.value
            return (
              <TipBox>
                <div style={{ fontWeight: 700, color: C.text, marginBottom: 6 }}>{label}</div>
                <div style={{ color: C.text2, marginBottom: 2 }}>Market placed at: <strong style={{ color: C.amber }}>{actual.value > 0 ? '+' : ''}{actual.value?.toFixed(2)}x vs sector</strong></div>
                <div style={{ color: C.text2, marginBottom: 6 }}>Model says fair at: <strong style={{ color: C.green }}>{pred.value > 0 ? '+' : ''}{pred.value?.toFixed(2)}x vs sector</strong></div>
                <div style={{ fontWeight: 700, color: gap > 0 ? C.green : C.red }}>
                  {gap > 0.1 ? `↑ Market undervalued by ${gap.toFixed(2)}x` : gap < -0.1 ? `↓ Market overvalued by ${Math.abs(gap).toFixed(2)}x` : '≈ Fairly priced'}
                </div>
              </TipBox>
            )
          }} />
        <Legend wrapperStyle={{ fontSize: 11, color: C.text2 }} />
        <ReferenceLine y={0} stroke={C.text3 + '88'} strokeDasharray="4 2" label={{ value: 'Sector median', fill: C.text3, fontSize: 9, position: 'insideTopLeft' }} />
        <Line type="monotone" dataKey="relPred" stroke={C.green} strokeWidth={2} strokeDasharray="7 3" dot={false} name="Model: fair premium/discount" />
        <Line type="monotone" dataKey="relActual" stroke={C.amber} strokeWidth={2.5}
          dot={props => {
            const { cx, cy, payload } = props
            if (!isFinite(cx) || !isFinite(cy)) return null
            const g = payload.relPred - payload.relActual
            const col = g > 0.5 ? C.green : g < -0.5 ? C.red : C.amber
            return <circle key={cx} cx={cx} cy={cy} r={5} fill={col} stroke={C.bg} strokeWidth={2} />
          }}
          activeDot={{ r: 8 }} name="Actual: market premium/discount" />
      </ComposedChart>
    </ChartCard>
  )
}

// ─── Relative Mode: Valuation Zone Chart ─────────────────────────────────────
// Shows all companies as a dot plot: X = company name, Y = gap (model - actual)
// Color-coded zones: green (undervalued) / grey (neutral) / red (overvalued)
function RelativeZoneChart({ allRows, target, features, clientCompany, filterOutliers, method }) {
  const C = useTheme()
  const model = useMemo(() => buildModel(allRows, target, features, filterOutliers, method, true), [allRows, target, features, filterOutliers, method])
  const validF = features.filter(f => f !== target)

  const data = useMemo(() => {
    if (!model) return []
    const companies = [...new Set(allRows.map(r => r._company).filter(Boolean))]
    const rows = []
    for (const co of companies) {
      const latest = allRows.filter(r => r._company === co && isFinite(r[target]) && validF.every(f => isFinite(r[f]))).sort((a, b) => b._year - a._year)[0]
      if (!latest) continue
      const yearVals = allRows.filter(r => r._year === latest._year && isFinite(r[target])).map(r => r[target])
      const s = yearVals.slice().sort((a, b) => a - b)
      const sectorMed = s[Math.floor(s.length / 2)]
      const relActual = latest[target] - sectorMed
      const rawPred = predict({ ...latest, _company: co }, model)
      const gap = rawPred - relActual  // positive = undervalued, negative = overvalued
      const se = model.se || 1
      const zone = gap > se * 0.5 ? 'undervalued' : gap < -se * 0.5 ? 'overvalued' : 'neutral'
      rows.push({ co, shortName: compShort(co), gap: parseFloat(gap.toFixed(3)), relActual: parseFloat(relActual.toFixed(3)), relPred: parseFloat(rawPred.toFixed(3)), zone, isClient: co === clientCompany })
    }
    return rows.sort((a, b) => b.gap - a.gap)
  }, [model, allRows, target, validF, clientCompany])

  if (!data.length) return null

  const gapDomain = getAxisDomain(data.map(d => d.gap), 0.15)
  const se = model?.se || 1
  const downloadData = data.map(d => ({ Company: d.co, 'Gap (model - market)': d.gap, 'Actual vs sector': d.relActual, 'Model says fair vs sector': d.relPred, 'Zone': d.zone }))

  return (
    <ChartCard
      title="Valuation gap vs peers — who is under/overvalued relative to fundamentals"
      subtitle="Gap = model's fair premium minus actual premium. Positive = undervalued. Negative = overvalued."
      explain="This chart shows the valuation gap for each company: how far is the market from where the model thinks it should be (relative to peers). Bars above zero = company deserves a higher premium than it's getting (undervalued). Bars below zero = company is getting a bigger premium than fundamentals justify (overvalued). The shaded zone in the middle is the neutral range — within normal model error."
      explainAdv={`The neutral zone (grey band) spans ±${(se * 0.5).toFixed(2)}x which is half the model's standard error. Companies outside this band have statistically notable mispricings.`}
      downloadData={downloadData}
      downloadName="relative-valuation-gap"
      height={Math.max(280, data.length * 38)}>
      <BarChart data={data} layout="vertical" margin={{ top: 4, right: 70, left: 150, bottom: 4 }}>
        <CartesianGrid stroke={C.grid} strokeDasharray="3 3" horizontal={false} />
        <XAxis type="number" stroke={C.text3} tick={{ fontSize: 9, fontFamily: 'var(--mono)' }}
          tickFormatter={v => `${v > 0 ? '+' : ''}${v.toFixed(1)}x`}
          label={{ value: '← Overvalued vs peers  |  Undervalued vs peers →', position: 'insideBottom', offset: -14, fill: C.text3, fontSize: 9 }} />
        <YAxis type="category" dataKey="shortName" stroke={C.text3} tick={{ fontSize: 10, fill: C.text2 }} width={145} />
        <Tooltip content={({ active, payload }) => {
          if (!active || !payload?.[0]) return null
          const d = payload[0].payload
          return (
            <TipBox>
              <div style={{ fontWeight: 700, color: d.isClient ? C.amber : C.text, marginBottom: 6 }}>{d.isClient ? '★ ' : ''}{d.co?.split(',')[0]}</div>
              <div style={{ color: C.text2, marginBottom: 2 }}>Currently at: <strong style={{ color: C.blue }}>{d.relActual > 0 ? '+' : ''}{d.relActual?.toFixed(2)}x vs sector</strong></div>
              <div style={{ color: C.text2, marginBottom: 6 }}>Model says fair: <strong style={{ color: C.green }}>{d.relPred > 0 ? '+' : ''}{d.relPred?.toFixed(2)}x vs sector</strong></div>
              <div style={{ fontWeight: 700, color: d.zone === 'undervalued' ? C.green : d.zone === 'overvalued' ? C.red : C.amber }}>
                {d.zone === 'undervalued' ? '↑ UNDERVALUED' : d.zone === 'overvalued' ? '↓ OVERVALUED' : '≈ NEUTRAL'}
              </div>
            </TipBox>
          )
        }} />
        <ReferenceLine x={0} stroke={C.text3} />
        <ReferenceLine x={se * 0.5} stroke={C.green + '44'} strokeDasharray="4 2" />
        <ReferenceLine x={-se * 0.5} stroke={C.red + '44'} strokeDasharray="4 2" />
        <Bar dataKey="gap" radius={[0, 4, 4, 0]} barSize={22}>
          <LabelList dataKey="gap" position="right" style={{ fill: C.text3, fontSize: 9, fontFamily: 'var(--mono)' }} formatter={v => `${v > 0 ? '+' : ''}${v.toFixed(2)}x`} />
          {data.map((d, i) => <Cell key={i} fill={d.isClient ? C.amber : d.zone === 'undervalued' ? C.green + 'cc' : d.zone === 'overvalued' ? C.red + 'aa' : C.text3 + '44'} />)}
        </Bar>
      </BarChart>
    </ChartCard>
  )
}

// ─── Relative Overview wrapper ─────────────────────────────────────────────────
function RelativeOverview({ result, allRows, target, features, clientCompany, filterOutliers, method, showAdvanced }) {
  const C = useTheme()
  const meta = getTargetMeta(target)
  return (
    <div>
      <ValuationSummary result={result} target={target} clientCompany={clientCompany} useRelative={true} allRows={allRows} features={features} method={method} filterOutliers={filterOutliers} />
      <RelativeZoneChart allRows={allRows} target={target} features={features} clientCompany={clientCompany} filterOutliers={filterOutliers} method={method} />
      <RelativePremiumOverTime allRows={allRows} target={target} features={features} clientCompany={clientCompany} filterOutliers={filterOutliers} method={method} />
      <RelativePositioningMap allRows={allRows} target={target} features={features} clientCompany={clientCompany} filterOutliers={filterOutliers} method={method} />
      {showAdvanced && (
        <div style={{ padding: '14px 18px', background: C.blueDim, border: `1px solid ${C.blue}22`, borderRadius: 12, fontSize: 12, color: C.text2, lineHeight: 1.8 }}>
          <strong style={{ color: C.blue }}>Technical:</strong> R² = {result.r2?.toFixed(3)} · Adj R² = {result.r2Adj?.toFixed(3) || '—'} · SE = {result.se?.toFixed(3)} · n = {result.n}
        </div>
      )}
    </div>
  )
}

// ─── Results Dashboard ────────────────────────────────────────────────────────
export default function ResultsDashboard({ result, config, allRows, onBack, onReset }) {
  const [tab, setTab] = useState('overview')
  const [showAdvanced, setShowAdvanced] = useState(false)
  const [isLight, setIsLight] = useState(false)
  const C = isLight ? LIGHT : DARK
  const { target, useRelative, clientCompany, features, method, filterOutliers, selectedCos, yearRange } = config

  // FIX: filter allRows to only selected companies and year range — same as what was used in regression
  const filteredAllRows = useMemo(() =>
    allRows.filter(r =>
      (selectedCos ? selectedCos.includes(r._company) : true) &&
      (!yearRange || (r._year >= yearRange[0] && r._year <= yearRange[1]))
    ),
    [allRows, selectedCos, yearRange]
  )
  const TABS = [['overview','Overview'],['asof','📅 Historical'],['investment','💼 Investment Case'],['validation','🧪 Validation'],['scenarios','🔮 Scenarios'],['drivers','Drivers'],['technical','Technical']]
  const methodLabels = { ols: 'Pooled OLS', ridge: 'Ridge', fe: 'Fixed Effects', fe_ridge: 'FE + Ridge' }

  // Inject light theme CSS variables when in light mode
  const lightCSS = isLight ? `
    body { background: #FFFFFF !important; }
    .chip { background: rgb(219,204,189) !important; border-color: rgba(96,0,29,0.15) !important; color: #4a4a4a !important; }
    .chip.on { background: rgba(24,78,98,0.12) !important; border-color: rgb(24,78,98) !important; color: rgb(24,78,98) !important; }
    .chip.on-green { background: rgba(24,78,98,0.1) !important; border-color: rgb(24,78,98) !important; color: rgb(24,78,98) !important; }
    .chip.on-amber { background: rgba(96,0,29,0.1) !important; border-color: rgb(96,0,29) !important; color: rgb(96,0,29) !important; }
    .tab { color: #888 !important; }
    .tab.active { background: rgb(219,204,189) !important; color: rgb(96,0,29) !important; box-shadow: 0 1px 4px rgba(96,0,29,0.15) !important; }
    .tab-bar { background: #F0EBE3 !important; }
    .kpi { background: #F8F5F2 !important; border-color: rgba(96,0,29,0.1) !important; }
    .num-input { background: #F0EBE3 !important; border-color: rgba(96,0,29,0.15) !important; color: #1a1a1a !important; }
    .num-input.mod { border-color: rgb(96,0,29) !important; }
    .run-btn { background: rgb(24,78,98) !important; }
    .run-btn:hover { background: rgb(18,60,76) !important; box-shadow: 0 8px 24px rgba(24,78,98,0.3) !important; }
  ` : ''

  return (
    <ThemeCtx.Provider value={C}>
      {isLight && <style>{lightCSS}</style>}
      <div style={{ minHeight: '100vh', display: 'flex', flexDirection: 'column', background: C.bg }}>
        <div style={{ borderBottom: `1px solid ${C.border}`, padding: '14px 28px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: C.bg1, position: 'sticky', top: 0, zIndex: 100 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
            <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.15em', color: isLight ? 'rgb(96,0,29)' : C.blue, textTransform: 'uppercase' }}>ValuationEngine</span>
            <span style={{ width: 1, height: 14, background: C.border, display: 'inline-block' }} />
            <span style={{ fontSize: 12, color: C.text2 }}>Predicting: <strong style={{ color: C.text }}>{useRelative ? `${target} (relative)` : target}</strong></span>
            <span style={{ width: 1, height: 14, background: C.border, display: 'inline-block' }} />
            <span style={{ fontSize: 12, color: C.amber }}>★ {compShort(clientCompany)}</span>
            <span style={{ padding: '3px 9px', borderRadius: 6, background: C.blueDim, border: `1px solid ${C.blue}33`, fontSize: 10, color: C.blue, fontWeight: 500 }}>{methodLabels[method]}</span>
            {useRelative && <span style={{ padding: '3px 9px', borderRadius: 6, background: C.greenDim, border: `1px solid ${C.green}33`, fontSize: 10, color: C.green, fontWeight: 500 }}>Relative mode</span>}
          </div>
          <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
            <ShareAnalysisButton config={config} result={result} allRows={filteredAllRows} isLight={isLight} />
            <ThemeToggle isLight={isLight} onToggle={() => setIsLight(v => !v)} />
            <button onClick={onBack} style={{ fontSize: 12, padding: '7px 14px', borderRadius: 8, background: C.bg2, border: `1px solid ${C.border}`, color: C.text2, cursor: 'pointer', fontFamily: 'var(--font)' }}>← Edit config</button>
            <button onClick={onReset} style={{ fontSize: 12, color: C.text3, background: 'none', border: 'none', cursor: 'pointer', fontFamily: 'var(--font)' }}>Change file</button>
          </div>
        </div>
        <div style={{ flex: 1, padding: '24px 28px', maxWidth: 1200, margin: '0 auto', width: '100%' }}>
          <KPIRow result={result} method={method} useRelative={useRelative} showAdvanced={showAdvanced} onToggleAdvanced={() => setShowAdvanced(v => !v)} />
          <div className="tab-bar" style={{ marginBottom: 16 }}>
            {TABS.map(([id, label]) => <button key={id} className={`tab ${tab === id ? 'active' : ''}`} onClick={() => setTab(id)}>{label}</button>)}
          </div>
          <div className="fade-up">
            <TabErrorBoundary key={tab}>
            {tab === 'overview' && !useRelative && <TabOverview result={result} clientCompany={clientCompany} useRelative={false} target={target} showAdvanced={showAdvanced} allRows={filteredAllRows} features={features} method={method} filterOutliers={filterOutliers} />}
            {tab === 'overview' && useRelative && <RelativeOverview result={result} allRows={filteredAllRows} target={target} features={features} clientCompany={clientCompany} filterOutliers={filterOutliers} method={method} showAdvanced={showAdvanced} />}
            {tab === 'asof' && <TabAsOf allRows={filteredAllRows} target={target} features={features} clientCompany={clientCompany} filterOutliers={filterOutliers} method={method} useRelative={useRelative} />}
            {tab === 'investment' && <TabInvestment allRows={filteredAllRows} target={target} features={features} clientCompany={clientCompany} filterOutliers={filterOutliers} method={method} useRelative={useRelative} showAdvanced={showAdvanced} />}
            {tab === 'validation' && <TabValidation allRows={filteredAllRows} target={target} features={features} filterOutliers={filterOutliers} />}
            {tab === 'scenarios' && <TabScenarios allRows={filteredAllRows} target={target} features={features} clientCompany={clientCompany} filterOutliers={filterOutliers} method={method} useRelative={useRelative} />}
            {tab === 'drivers' && <TabDrivers result={result} showAdvanced={showAdvanced} />}
            {tab === 'technical' && <TabTechnical result={result} method={method} target={target} />}
            </TabErrorBoundary>
          </div>
        </div>
      </div>
      <AIChatbot config={config} result={result} isLight={isLight} />
    </ThemeCtx.Provider>
  )
}
