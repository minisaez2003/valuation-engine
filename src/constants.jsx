export const DEFAULT_Y = 'Approx EV/EBITDA'  // overridden dynamically if not found

// Ordered priority lists — first match found in the file wins
export const DEFAULT_Y_CANDIDATES = [
  'Approx EV/EBITDA', 'EV / EBITDA (x)', 'EV/EBITDA', 'EV/EBITDA (x)',
]

export const DEFAULT_FEATURES_CANDIDATES = [
  // Growth
  ['Sales Growth (%)', 'Sales Growth (decimal)', 'Revenue Growth (%)', 'Sales Growth (%)'],
  // Margin
  ['EBITDA Margin (%)', 'EBITDA Margin (decimal)', 'EBITDA Margin'],
  // ROIC
  ['ROIC Proxy', 'ROIC Proxy (decimal)', 'ROIC (%)'],
  // Scale
  ['Scale / Network Advantage (1-5)'],
  // Cyclicality
  ['Cyclicality (1=Low,5=High)'],
  // Theme
  ['Theme Premium Score (1-5)'],
]

export const DEFAULT_FEATURES = [
  'Sales Growth (%)', 'EBITDA Margin (%)', 'ROIC Proxy',
  'Scale / Network Advantage (1-5)', 'Cyclicality (1=Low,5=High)', 'Theme Premium Score (1-5)'
]

export const META_SKIP = new Set([
  'Include', 'Importance', 'Source Row', 'Source Slot',
  'As-of Date', 'As-of Year', 'FactSet ID', 'Bucket', 'Theme Bucket',
  'Year', 'Company', 'RBICS Industry Group', 'RBICS Sub-Industry', 'RBICS Sector',
])
export const compShort = c =>
  c.replace(/,?\s*(Inc\.|Corp\.|Co\.|Ltd\.|LLC|Class\s+[A-Z]|Holdings|Companies|Enterprises)\.?/gi, '')
   .trim().split(' ').slice(0, 2).join(' ')

// Colour tokens (JS access for recharts)
export const C = {
  bg: '#0c0f14', bg1: '#111520', bg2: '#161b26', bg3: '#1c2333',
  border: 'rgba(255,255,255,0.07)',
  blue: '#3b82f6', blueDim: 'rgba(59,130,246,0.15)',
  green: '#10b981', greenDim: 'rgba(16,185,129,0.12)',
  red: '#f43f5e', redDim: 'rgba(244,63,94,0.12)',
  amber: '#f59e0b', amberDim: 'rgba(245,158,11,0.12)',
  text: '#f1f5f9', text2: '#94a3b8', text3: '#475569',
  grid: 'rgba(255,255,255,0.04)'
}
