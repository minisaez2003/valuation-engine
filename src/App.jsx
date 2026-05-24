import { useState, useCallback, useMemo, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { DEFAULT_Y, DEFAULT_Y_CANDIDATES, DEFAULT_FEATURES, DEFAULT_FEATURES_CANDIDATES, META_SKIP, compShort } from './constants.jsx'
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
@keyframes chatPulse { 0%,100%{box-shadow:0 4px 24px var(--blue,#3b82f6)66,0 0 0 0 var(--blue,#3b82f6)44} 50%{box-shadow:0 4px 24px var(--blue,#3b82f6)88,0 0 0 8px transparent} }
@keyframes tickerScroll { from{transform:translateX(0)} to{transform:translateX(-50%)} }
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

/* Responsive grid utilities */
.grid-2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.grid-4{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
.hero-grid{display:grid;grid-template-columns:1fr 1fr;gap:64px;align-items:center}
.dashboard-pad{padding:24px 28px}
.setup-pad{padding:28px 28px}
.upload-pad{padding:48px 32px}
.header-pad{padding:14px 28px}
.results-max{max-width:1200px;margin:0 auto;width:100%}
.setup-max{max-width:1080px;margin:0 auto;width:100%}
.hero-max{max-width:900px;width:100%}

/* Mobile: < 768px */
@media (max-width: 768px) {
  body{font-size:14px}
  .grid-2{grid-template-columns:1fr;gap:10px}
  .grid-4{grid-template-columns:1fr 1fr;gap:8px}
  .hero-grid{grid-template-columns:1fr;gap:32px}
  .dashboard-pad{padding:14px 14px}
  .setup-pad{padding:18px 14px}
  .upload-pad{padding:24px 16px}
  .header-pad{padding:10px 14px}
  .glass{padding:16px}
  .kpi{padding:12px;min-width:80px}
  .kpi-val{font-size:18px}
  .mcard{min-width:120px;padding:12px}
  .chip{padding:6px 10px;font-size:11px}
  .tab{padding:7px 11px;font-size:11px}
  .run-btn{width:100%;padding:14px 24px;font-size:13px}
  .sel{font-size:14px}
  .slabel{font-size:9px}
  .num-input{font-size:13px}
  .header-pad>div{flex-wrap:wrap !important;gap:6px !important}
  .header-pad span{font-size:10px !important}
  .tab-bar{overflow-x:auto;-webkit-overflow-scrolling:touch}
  .mobile-hide{display:none !important}
  .mobile-stack{flex-direction:column !important;align-items:stretch !important}
  .mobile-full{width:100% !important}
  .recharts-wrapper{font-size:9px}
  /* Chat panel: full width on mobile */
  .ai-chat-panel{width:calc(100vw - 16px) !important;height:calc(100vh - 100px) !important;right:8px !important;bottom:80px !important;max-width:380px;max-height:540px}
  .ai-chat-fab{bottom:16px !important;right:16px !important;width:48px !important;height:48px !important}
}

/* Small mobile: < 480px */
@media (max-width: 480px) {
  .grid-4{grid-template-columns:1fr}
  h1{font-size:28px !important}
  .upload-pad h1{font-size:28px !important}
}
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
  const [pendingData, setPendingData] = useState(null)
  const [showWarning, setShowWarning] = useState(false)
  const [fileSensitivityWarning, setFileSensitivityWarning] = useState(null)

  // I: audit log — save filename + date to localStorage (no data, just metadata)
  const logFileAccess = (filename, companies, target) => {
    try {
      const log = JSON.parse(localStorage.getItem('ve_audit_log') || '[]')
      log.unshift({ filename, companies, target, date: new Date().toISOString(), confirmed: true })
      if (log.length > 50) log.pop() // keep last 50 entries
      localStorage.setItem('ve_audit_log', JSON.stringify(log))
    } catch {}
  }

  const processFile = f => {
    if (!f) return
    // J: Check filename for potential confidential data keywords
    const filename = f.name || ''
    const sensitivePatterns = /confidential|NDA|nonpublic|non-public|MNPI|private|project\s*[a-z]/i
    const dealPatterns = /\b(deal|target|acquire|acquisition|merger|M&A|DD|due.?dilig|IM\b|CIM\b|pitch.?book)\b/i
    const isSensitive = sensitivePatterns.test(filename) || dealPatterns.test(filename)
    setFileSensitivityWarning(isSensitive ? filename : null)
    setLoading(true)
    const r = new FileReader()
    r.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: false })
      const SKIP_SHEETS = /^__|^#|^Sheet\d*$|cache/i
      const dataSheets = wb.SheetNames.filter(n => !SKIP_SHEETS.test(n))
      const preferredNames = ['Annualized_Panel', 'Annualized_Used', 'Annualized', 'Panel', 'Data', 'Sheet1']
      const name = preferredNames.find(p => wb.SheetNames.includes(p)) || dataSheets[0] || wb.SheetNames[0]
      const data = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: null, raw: true })
      setTimeout(() => {
        setPendingData({ data, name })
        setLoading(false)
        setShowWarning(true) // Show compliance warning before loading
      }, 200)
    }
    r.readAsArrayBuffer(f)
  }

  const confirmAndLoad = () => {
    if (pendingData) {
      // I: log the file access
      logFileAccess(pendingData.name, pendingData.data?.length || 0, 'pending')
      onData(pendingData.data, pendingData.name)
      setShowWarning(false)
      setPendingData(null)
      setFileSensitivityWarning(null)
    }
  }

  const claims = [
    { icon: '⚡', text: 'Valuation range in under 60 seconds' },
    { icon: '📐', text: 'OLS, Ridge & Fixed Effects — pick your model' },
    { icon: '🎯', text: 'Relative mode removes macro noise' },
    { icon: '💬', text: 'Ask the AI anything about your results' },
  ]

  return (
    <div style={{ minHeight: '100vh', display: 'flex', flexDirection: 'column', position: 'relative', overflow: 'hidden', background: '#060810' }}>

      {/* Wall Street background — animated grid + glow */}
      <div style={{ position: 'absolute', inset: 0, zIndex: 0, pointerEvents: 'none' }}>
        {/* Grid lines */}
        <svg width="100%" height="100%" style={{ position: 'absolute', inset: 0, opacity: 0.07 }}>
          <defs>
            <pattern id="grid" width="60" height="60" patternUnits="userSpaceOnUse">
              <path d="M 60 0 L 0 0 0 60" fill="none" stroke="#3b82f6" strokeWidth="0.5"/>
            </pattern>
          </defs>
          <rect width="100%" height="100%" fill="url(#grid)" />
        </svg>
        {/* Blue glow top-left */}
        <div style={{ position: 'absolute', top: '-20%', left: '-10%', width: '60%', height: '70%', background: 'radial-gradient(ellipse, rgba(59,130,246,0.12) 0%, transparent 70%)', borderRadius: '50%' }} />
        {/* Gold glow bottom-right */}
        <div style={{ position: 'absolute', bottom: '-20%', right: '-10%', width: '50%', height: '60%', background: 'radial-gradient(ellipse, rgba(184,160,74,0.08) 0%, transparent 70%)', borderRadius: '50%' }} />
        {/* Animated ticker-style numbers — purely decorative */}
        <div style={{ position: 'absolute', bottom: 0, left: 0, right: 0, height: 36, background: 'rgba(255,255,255,0.02)', borderTop: '1px solid rgba(255,255,255,0.04)', display: 'flex', alignItems: 'center', overflow: 'hidden' }}>
          <div style={{ display: 'flex', gap: 48, fontSize: 11, fontFamily: 'var(--mono)', color: 'rgba(255,255,255,0.2)', whiteSpace: 'nowrap', animation: 'tickerScroll 30s linear infinite' }}>
            {['EV/EBITDA 12.4x ▲', 'EV/Sales 1.8x ▼', 'R² 0.61 —', 'EBITDA Margin 18.2%', 'Upside +34%', 'Peers 20', 'ROIC 14.3%', 'Growth 8.1%', 'EV/EBITDA 12.4x ▲', 'EV/Sales 1.8x ▼', 'R² 0.61 —', 'EBITDA Margin 18.2%'].map((t, i) => (
              <span key={i}>{t}</span>
            ))}
          </div>
        </div>
      </div>

      {/* Compliance warning modal */}
      {showWarning && (
        <div style={{ position: 'fixed', inset: 0, zIndex: 9999, background: 'rgba(0,0,0,0.85)', backdropFilter: 'blur(12px)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 24 }}>
          <div style={{ background: '#111520', border: '1px solid rgba(255,255,255,0.1)', borderRadius: 20, padding: '32px 36px', maxWidth: 520, width: '100%', boxShadow: '0 24px 80px rgba(0,0,0,0.8)' }}>
            <div style={{ fontSize: 28, marginBottom: 14 }}>⚠️</div>
            <div style={{ fontSize: 17, fontWeight: 700, color: 'white', marginBottom: 14 }}>Public Data Only</div>
            {/* J: Sensitive filename alert */}
            {fileSensitivityWarning && (
              <div style={{ padding: '12px 16px', background: 'rgba(244,63,94,0.15)', border: '1px solid rgba(244,63,94,0.4)', borderRadius: 10, marginBottom: 16, fontSize: 12, color: '#f87171', lineHeight: 1.7 }}>
                <strong>⚠ Filename alert:</strong> "<em>{fileSensitivityWarning}</em>" may contain sensitive deal data. Are you certain this file contains only public market information?
              </div>
            )}
            <div style={{ fontSize: 13, color: '#94a3b8', lineHeight: 1.9, marginBottom: 24 }}>
              This tool is designed for use with <strong style={{ color: 'white' }}>publicly available market data only</strong>. Do not upload:
              <ul style={{ marginTop: 10, marginLeft: 20, display: 'flex', flexDirection: 'column', gap: 6 }}>
                <li>Confidential client information</li>
                <li>Non-public financial projections</li>
                <li>Data subject to NDA or confidentiality agreement</li>
                <li>Material non-public information (MNPI)</li>
              </ul>
            </div>
            <div style={{ fontSize: 11, color: '#475569', marginBottom: 24, padding: '10px 14px', background: 'rgba(255,255,255,0.04)', borderRadius: 8, lineHeight: 1.7 }}>
              By continuing you confirm this file contains only publicly available market information.
            </div>
            <div style={{ display: 'flex', gap: 12 }}>
              <button onClick={() => { setShowWarning(false); setPendingData(null) }}
                style={{ flex: 1, padding: '13px', borderRadius: 10, border: '1px solid rgba(255,255,255,0.1)', background: 'transparent', color: '#94a3b8', fontSize: 13, cursor: 'pointer', fontFamily: 'var(--font)' }}>
                Cancel
              </button>
              <button onClick={confirmAndLoad}
                style={{ flex: 1, padding: '13px', borderRadius: 10, border: 'none', background: '#3b82f6', color: 'white', fontSize: 13, fontWeight: 700, cursor: 'pointer', fontFamily: 'var(--font)' }}>
                I confirm — continue →
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Main hero */}
      <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', position: 'relative', zIndex: 1, padding: '48px 32px 80px' }}>
        <div style={{ maxWidth: 1000, width: '100%' }}>
          <div className="hero-grid">
            {/* Left: epic copy */}
            <div className="fade-up">
              {/* Badge */}
              <div style={{ display: 'inline-flex', alignItems: 'center', gap: 8, padding: '5px 14px', borderRadius: 20, border: '1px solid rgba(59,130,246,0.4)', background: 'rgba(59,130,246,0.08)', marginBottom: 28 }}>
                <div style={{ width: 6, height: 6, borderRadius: '50%', background: '#3b82f6', boxShadow: '0 0 8px #3b82f6' }} />
                <span style={{ fontSize: 11, fontWeight: 600, letterSpacing: '0.14em', color: '#3b82f6', textTransform: 'uppercase' }}>ValuationEngine · by Pablo Saez</span>
              </div>

              <h1 style={{ fontSize: 52, fontWeight: 800, lineHeight: 1.05, marginBottom: 24, letterSpacing: '-0.03em', color: 'white' }}>
                The future of<br />
                <span style={{ background: 'linear-gradient(135deg, #3b82f6 0%, #818cf8 50%, #60a5fa 100%)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent', backgroundClip: 'text' }}>M&A valuation</span><br />
                is statistical.
              </h1>

              <p style={{ fontSize: 15, color: '#64748b', lineHeight: 1.85, marginBottom: 36, maxWidth: 440 }}>
                Regression-based comparable analysis. Tell your client <em style={{ color: '#94a3b8' }}>exactly</em> which companies are mispriced, why, and by how much — in under 60 seconds.
              </p>

              {/* Stats row */}
              <div style={{ display: 'flex', gap: 28, marginBottom: 36, flexWrap: 'wrap' }}>
                {[['60s', 'To valuation range'], ['4', 'Regression methods'], ['100%', 'Browser-private']].map(([val, label]) => (
                  <div key={label}>
                    <div style={{ fontSize: 26, fontWeight: 800, color: 'white', letterSpacing: '-0.02em', lineHeight: 1 }}>{val}</div>
                    <div style={{ fontSize: 11, color: '#475569', marginTop: 4, letterSpacing: '0.04em' }}>{label}</div>
                  </div>
                ))}
              </div>

              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                {['OLS · Ridge · Fixed Effects', 'AI Analyst built-in', 'Apollo-style reports', 'Public data only'].map(t => (
                  <span key={t} style={{ padding: '5px 12px', borderRadius: 20, border: '1px solid rgba(255,255,255,0.08)', fontSize: 11, color: '#64748b', background: 'rgba(255,255,255,0.03)' }}>{t}</span>
                ))}
              </div>
            </div>

            {/* Right: upload card */}
            <div className="fade-up d2">
              <div
                onDragOver={e => { e.preventDefault(); setDrag(true) }}
                onDragLeave={() => setDrag(false)}
                onDrop={e => { e.preventDefault(); setDrag(false); processFile(e.dataTransfer.files[0]) }}
                onClick={() => !loading && document.getElementById('xlsxIn').click()}
                style={{
                  border: `1.5px dashed ${drag ? '#3b82f6' : 'rgba(255,255,255,0.12)'}`,
                  borderRadius: 24, padding: '44px 32px',
                  cursor: loading ? 'wait' : 'pointer',
                  background: drag ? 'rgba(59,130,246,0.06)' : 'rgba(255,255,255,0.03)',
                  backdropFilter: 'blur(12px)',
                  transition: 'all 0.2s ease',
                  transform: drag ? 'scale(1.02)' : 'scale(1)',
                  textAlign: 'center',
                  position: 'relative',
                  boxShadow: drag ? '0 0 40px rgba(59,130,246,0.2)' : 'inset 0 1px 0 rgba(255,255,255,0.06)',
                }}>
                {loading
                  ? <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 14 }}>
                      <div style={{ width: 40, height: 40, border: '3px solid rgba(255,255,255,0.08)', borderTopColor: '#3b82f6', borderRadius: '50%', animation: 'spin 0.7s linear infinite' }} />
                      <div style={{ color: '#64748b', fontSize: 14 }}>Reading your data…</div>
                    </div>
                  : <>
                      {/* Icon */}
                      <div style={{ width: 64, height: 64, borderRadius: 20, background: 'linear-gradient(135deg, rgba(59,130,246,0.2), rgba(99,102,241,0.1))', border: '1px solid rgba(59,130,246,0.2)', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 20px', fontSize: 28 }}>📊</div>
                      <div style={{ fontSize: 18, fontWeight: 700, color: 'white', marginBottom: 8 }}>Drop your comps file</div>
                      <div style={{ fontSize: 13, color: '#64748b', marginBottom: 24 }}>Excel · .xlsx or .xls</div>
                      <div style={{ display: 'inline-block', padding: '12px 32px', background: '#3b82f6', borderRadius: 12, fontSize: 13, fontWeight: 700, color: 'white', boxShadow: '0 4px 16px rgba(59,130,246,0.4)', letterSpacing: '0.02em' }}>
                        Upload file →
                      </div>
                      <div style={{ marginTop: 20, fontSize: 11, color: '#334155', lineHeight: 1.8 }}>
                        Sheet "Annualized_Panel" or "Annualized_Used"<br />
                        Your data never leaves this browser
                      </div>
                    </>}
              </div>
              <input id="xlsxIn" type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={e => processFile(e.target.files[0])} />
            </div>
          </div>
        </div>
      </div>

      {/* Bottom footer strip */}
      <div style={{ borderTop: '1px solid rgba(255,255,255,0.04)', padding: '12px 32px', display: 'flex', justifyContent: 'center', gap: 32, flexWrap: 'wrap', position: 'relative', zIndex: 1 }}>
        {['Pooled OLS', 'Ridge Regression', 'Fixed Effects', 'Walk-forward CV', 'AI Analyst', 'PDF Reports'].map(t => (
          <span key={t} style={{ fontSize: 10, color: '#1e293b', letterSpacing: '0.1em', textTransform: 'uppercase', fontWeight: 500 }}>{t}</span>
        ))}
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

  // Smart defaults — pick first candidate that exists in this file's columns
  const smartDefaultY = useMemo(() =>
    DEFAULT_Y_CANDIDATES.find(c => numCols.includes(c)) || targetCols[0] || DEFAULT_Y,
    [numCols, targetCols]
  )
  const smartDefaultFeatures = useMemo(() => {
    const picked = []
    for (const group of DEFAULT_FEATURES_CANDIDATES) {
      const match = group.find(c => numCols.includes(c))
      if (match) picked.push(match)
    }
    // fallback: use DEFAULT_FEATURES if nothing matched (old file format)
    return picked.length >= 3 ? picked : DEFAULT_FEATURES.filter(f => numCols.includes(f))
  }, [numCols])

  const [target, setTarget] = useState(() => {
    // D: restore from localStorage if available
    try {
      const saved = JSON.parse(localStorage.getItem('ve_config') || '{}')
      if (saved.target && numCols.includes(saved.target)) return saved.target
    } catch {}
    return DEFAULT_Y_CANDIDATES.find(c => numCols.includes(c)) || targetCols[0] || DEFAULT_Y
  })
  const featureCols = useMemo(() => numCols.filter(c => !META_SKIP.has(c) && c !== target), [numCols, target])

  const [client, setClient] = useState(() => {
    try {
      const saved = JSON.parse(localStorage.getItem('ve_config') || '{}')
      if (saved.client && allCompanies.includes(saved.client)) return saved.client
    } catch {}
    return allCompanies[0] || ''
  })
  const [selCos, setSelCos] = useState(allCompanies)
  const [yearRange, setYearRange] = useState([2018, 2025])
  const [features, setFeatures] = useState(() => {
    try {
      const saved = JSON.parse(localStorage.getItem('ve_config') || '{}')
      if (Array.isArray(saved.features) && saved.features.every(f => numCols.includes(f))) return saved.features
    } catch {}
    const picked = []
    for (const group of DEFAULT_FEATURES_CANDIDATES) {
      const match = group.find(c => numCols.includes(c))
      if (match) picked.push(match)
    }
    return picked.length >= 3 ? picked : DEFAULT_FEATURES.filter(f => numCols.includes(f)).length >= 3
      ? DEFAULT_FEATURES.filter(f => numCols.includes(f))
      : numCols.slice(0, 6)
  })
  const [method, setMethod] = useState(() => {
    try { const s = JSON.parse(localStorage.getItem('ve_config') || '{}'); return s.method || 'ols' } catch { return 'ols' }
  })
  const [filterOut, setFilterOut] = useState(true)
  const [useRelative, setUseRelative] = useState(() => {
    try { return JSON.parse(localStorage.getItem('ve_config') || '{}').useRelative || false } catch { return false }
  })
  const [running, setRunning] = useState(false)
  const [warning, setWarning] = useState('')
  // E: advanced config collapsed by default — MD sees Quick Run first
  const [showAdvanced, setShowAdvanced] = useState(false)

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
    // D: save config so next session pre-fills
    try { localStorage.setItem('ve_config', JSON.stringify({ target, client, features, method, useRelative })) } catch {}
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
      <div style={{ borderBottom: '1px solid var(--border)', display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: 'var(--bg1)' }} className="header-pad">
        <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
          <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.15em', color: 'var(--blue)', textTransform: 'uppercase' }}>ValuationEngine</span>
          <span style={{ width: 1, height: 14, background: 'var(--border)', display: 'inline-block' }} />
          <span style={{ fontSize: 12, color: 'var(--text2)' }}>{sheetName} · {allCompanies.length} companies</span>
        </div>
        <button onClick={onReset} style={{ fontSize: 12, color: 'var(--text3)', background: 'none', border: 'none', cursor: 'pointer', fontFamily: 'var(--font)' }}>← Change file</button>
      </div>
      <div style={{ flex: 1 }} className="setup-pad setup-max">

        {/* E: Quick Run hero — visible first, no scrolling needed */}
        <div className="glass fade-up" style={{ marginBottom: 20, padding: '24px 28px', borderLeft: '4px solid var(--blue)' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 16 }}>
            <div>
              <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 4 }}>Ready to run</div>
              <div style={{ fontSize: 13, color: 'var(--text2)', lineHeight: 1.6 }}>
                <strong style={{ color: 'var(--text)' }}>{client?.split(',')[0] || 'Client'}</strong> · {target} · {method.toUpperCase()} · {obsCount} observations
              </div>
              <div style={{ fontSize: 11, color: 'var(--text3)', marginTop: 4 }}>Valuation range + peer ranking + AI analyst ready in seconds</div>
            </div>
            <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
              <button
                className="run-btn"
                onClick={() => {
                  setTarget(smartDefaultY); setFeatures(smartDefaultFeatures); setMethod('ols')
                  setFilterOut(true); setSelCos(allCompanies)
                  const yrNums = allYears.map(Number)
                  setYearRange([Math.min(...yrNums), Math.max(...yrNums)])
                  setTimeout(handleRun, 50)
                }}
                disabled={running || obsCount < 5}
                style={{ background: 'var(--green, #10b981)', minWidth: 160 }}>
                {running ? <><div className="spinner" />Running…</> : '⚡ Quick run'}
              </button>
              <button className="run-btn" onClick={handleRun} disabled={running || obsCount < 5}
                style={{ background: 'var(--blue)', minWidth: 140 }}>
                {running ? <><div className="spinner" />Running…</> : '▶ Run with my settings'}
              </button>
            </div>
          </div>
          {warning && <div className="warn" style={{ marginTop: 12 }}>⚠ {warning}</div>}
        </div>

        {/* E: Advanced config — collapsed by default */}
        <div style={{ marginBottom: 14 }}>
          <button
            onClick={() => setShowAdvanced(v => !v)}
            style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '10px 16px', borderRadius: 10, border: '1px solid var(--border)', background: 'var(--bg2)', color: 'var(--text2)', fontSize: 12, cursor: 'pointer', fontFamily: 'var(--font)', width: '100%', justifyContent: 'space-between' }}>
            <span style={{ fontWeight: 500 }}>⚙ Advanced configuration · {showAdvanced ? 'collapse' : 'expand to customise'}</span>
            <span style={{ color: 'var(--text3)', fontSize: 16 }}>{showAdvanced ? '▲' : '▼'}</span>
          </button>
        </div>

        {showAdvanced && <>
          <div className="grid-2" style={{ marginBottom: 14 }}>
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
                <div className="mdesc">Predicts the raw multiple directly. Wider prediction range but simpler.</div>
              </div>
              <div className={`mcard ${useRelative ? 'sel' : ''}`} onClick={() => setUseRelative(true)} style={{ maxWidth: 300 }}>
                <div className="radio" />
                <div className="mname" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>Relative to sector <span style={{ fontSize: 10, padding: '2px 7px', borderRadius: 5, background: 'var(--green-d, rgba(16,185,129,0.12))', color: 'var(--green)', border: '1px solid rgba(16,185,129,0.25)', fontWeight: 500 }}>Recommended</span></div>
                <div className="mdesc">Predicts premium/discount vs sector median. Removes macro noise. Cleaner signal.</div>
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

        <div className="glass fade-up d4" style={{ marginBottom: 14 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
            <div className="slabel" style={{ margin: 0 }}>Independent variables X ({validF.length} selected)</div>
            <button onClick={() => setFeatures(smartDefaultFeatures)} style={{ fontSize: 11, color: 'var(--blue)', background: 'none', border: 'none', cursor: 'pointer', fontFamily: 'var(--font)' }}>Reset to recommended</button>
          </div>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 7 }}>
            {featureCols.map(f => { const on = features.includes(f); const rec = DEFAULT_FEATURES.includes(f) || smartDefaultFeatures.includes(f); return <div key={f} className={`chip ${on ? 'on' : ''}`} onClick={() => setFeatures(p => p.includes(f) ? p.filter(x => x !== f) : [...p, f])} style={{ opacity: on ? 1 : 0.4 }}>{rec && <span style={{ fontSize: 9, opacity: 0.6 }}>★</span>}{f}</div> })}
          </div>
          <div style={{ marginTop: 10, fontSize: 11, color: 'var(--text3)', lineHeight: 1.6 }}>★ = recommended · Keep 4–7 variables</div>
        </div>
        </> }

      </div>
    </div>
  )
}

export default function App() {
  const [rawData, setRawData] = useState(null)
  const [sheetName, setSheetName] = useState('')
  const [runResult, setRunResult] = useState(null)

  const handleData = useCallback((data, sheet) => {
    if (!data || data.length === 0) return
    const headers = Object.keys(data[0] || {})

    // ── Auto-detect key columns ──────────────────────────────────────────────
    // Company name: try known variants
    const companyCol = headers.find(h =>
      /^company\s*name$/i.test(h) || /^company$/i.test(h) || /^firm$/i.test(h) || /^issuer$/i.test(h)
    ) || headers.find(h => /company/i.test(h)) || headers[0]

    // Year column: try known variants
    const yearCol = headers.find(h =>
      /^as.?of\s*year$/i.test(h) || /^year$/i.test(h) || /^fiscal.?year$/i.test(h) || /^fy$/i.test(h)
    ) || headers.find(h => /year/i.test(h))

    // Bucket column
    const bucketCol = headers.find(h => /^bucket$/i.test(h)) || ''

    // ── Build META_SKIP for this file ────────────────────────────────────────
    const alwaysSkip = new Set([
      'Include', 'Importance', 'Source Row', 'Source Slot', 'As-of Date',
      'FactSet ID', 'Bucket', 'Theme Bucket', 'RBICS Industry Group',
      'RBICS Sub-Industry', 'RBICS Sector',
    ])
    if (companyCol) alwaysSkip.add(companyCol)
    if (yearCol) alwaysSkip.add(yearCol)
    if (bucketCol) alwaysSkip.add(bucketCol)

    // ── Normalise rows ───────────────────────────────────────────────────────
    // Some files store percentages as decimals (0.07 instead of 7.0).
    // Detect by checking if a "margin" or "growth" column has values < 1.5
    const isDecimalCol = col => {
      const vals = data.map(r => parseFloat(r[col])).filter(v => isFinite(v) && v !== 0)
      if (!vals.length) return false
      const allSmall = vals.every(v => Math.abs(v) < 2)
      return allSmall && (/(decimal|pct|ratio)/i.test(col) || /margin|growth|roic|capex.?sales/i.test(col))
    }

    // ── Compute derived columns ──────────────────────────────────────────────
    const withCalc = data.map(row => {
      const out = { ...row }
      // EV/Sales — try existing column first, then compute
      if (!isFinite(parseFloat(out['EV / Sales (x)'])) && !isFinite(parseFloat(out['EV/Sales']))) {
        const ev = parseFloat(out['EV ($mm)'] ?? out['Enterprise Value ($mm)'] ?? out['EV ($M)'])
        const sales = parseFloat(out['Sales ($mm)'] ?? out['Revenue ($mm)'])
        if (isFinite(ev) && isFinite(sales) && sales > 0) out['EV/Sales (calc)'] = ev / sales
      }
      // Approx EV/EBITDA — alias if needed
      if (!isFinite(parseFloat(out['Approx EV/EBITDA'])) && isFinite(parseFloat(out['EV / EBITDA (x)']))) {
        out['Approx EV/EBITDA'] = out['EV / EBITDA (x)']
      }
      // Normalise decimal columns to percentage for readability
      for (const col of headers) {
        if (isDecimalCol(col) && isFinite(parseFloat(out[col]))) {
          const pctName = col.replace(/\(decimal\)/i, '(%)').replace(/\s*\(decimal\)\s*/i, ' (%)').trim()
          if (pctName !== col) out[pctName] = parseFloat(out[col]) * 100
        }
      }
      return out
    })

    // ── Detect numeric columns ───────────────────────────────────────────────
    const numCols = Object.keys(withCalc[0] || {}).filter(k =>
      !alwaysSkip.has(k) && isNumCol(withCalc, k)
    )

    // ── Clean rows with standardised _company / _year fields ────────────────
    const cleaned = withCalc.map(row => {
      const r = {
        _company: String(row[companyCol] || '').trim(),
        _year: toNum(row[yearCol]),
        _bucket: String(row[bucketCol] || '').trim(),
      }
      for (const k of numCols) r[k] = toNum(row[k])
      return r
    }).filter(r => r._company && isFinite(r._year))

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
