// ─── Number utilities ─────────────────────────────────────────────────────────
export const toNum = v => {
  if (v == null || v === '') return NaN
  if (typeof v === 'number') return v
  if (v instanceof Date) return NaN
  const n = Number(String(v).replace(/[,%$]/g, '').trim())
  return isNaN(n) ? NaN : n
}

export const isNumCol = (data, k) => {
  let g = 0, t = 0
  for (const r of data) {
    const v = r[k]
    if (v == null || v === '') continue
    t++
    if (isFinite(toNum(v))) g++
  }
  return t > 0 && g / t >= 0.8
}

export const avg = arr => arr.reduce((s, v) => s + v, 0) / arr.length
export const sdv = arr => {
  const m = avg(arr)
  const s = Math.sqrt(arr.reduce((s, v) => s + (v - m) ** 2, 0) / arr.length)
  return s < 1e-10 ? 1 : s
}

// ─── Standardization ──────────────────────────────────────────────────────────
export const standardize = (data, feats) => {
  const ms = feats.map(f => avg(data.map(r => r[f])))
  const ss = feats.map(f => sdv(data.map(r => r[f])))
  const Xs = data.map(r => [1, ...feats.map((f, i) => (r[f] - ms[i]) / ss[i])])
  return { Xs, ms, ss }
}

export const applyStd = (row, feats, ms, ss) =>
  [1, ...feats.map((f, i) => ((row[f] ?? 0) - ms[i]) / ss[i])]

// ─── Matrix inversion (Gauss-Jordan) ─────────────────────────────────────────
export function matInv(m) {
  const n = m.length
  const aug = m.map((row, i) => {
    const r = row.slice()
    for (let j = 0; j < n; j++) r.push(j === i ? 1 : 0)
    return r
  })
  for (let col = 0; col < n; col++) {
    let mx = col
    for (let row = col + 1; row < n; row++)
      if (Math.abs(aug[row][col]) > Math.abs(aug[mx][col])) mx = row
    ;[aug[col], aug[mx]] = [aug[mx], aug[col]]
    const piv = aug[col][col]
    if (Math.abs(piv) < 1e-10) return null
    for (let j = 0; j < 2 * n; j++) aug[col][j] /= piv
    for (let row = 0; row < n; row++) {
      if (row !== col) {
        const f = aug[row][col]
        for (let j = 0; j < 2 * n; j++) aug[row][j] -= f * aug[col][j]
      }
    }
  }
  return aug.map(r => r.slice(n))
}

// ─── OLS ──────────────────────────────────────────────────────────────────────
export function ols(X, y) {
  const n = y.length, k = X[0].length
  if (n < k + 2) return null
  const XtX = Array.from({ length: k }, (_, i) =>
    Array.from({ length: k }, (_, j) => X.reduce((s, r) => s + r[i] * r[j], 0))
  )
  const inv = matInv(XtX)
  if (!inv) return null
  const Xty = Array.from({ length: k }, (_, i) =>
    X.reduce((s, r, ri) => s + r[i] * y[ri], 0)
  )
  const beta = inv.map(row => row.reduce((s, v, j) => s + v * Xty[j], 0))
  const yHat = X.map(r => r.reduce((s, v, i) => s + v * beta[i], 0))
  const res = y.map((v, i) => v - yHat[i])
  const sse = res.reduce((s, r) => s + r * r, 0)
  const ym = avg(y), sst = y.reduce((s, v) => s + (v - ym) ** 2, 0)
  const r2 = sst < 1e-14 ? 0 : 1 - sse / sst
  const r2Adj = 1 - (1 - r2) * (n - 1) / (n - k)
  const mse = sse / Math.max(n - k, 1), se = Math.sqrt(mse)
  const bSE = inv.map((row, i) => Math.sqrt(Math.abs(row[i] * mse)))
  const tS = beta.map((b, i) => bSE[i] < 1e-14 ? 0 : b / bSE[i])
  const pV = tS.map(t => approxP(Math.abs(t), n - k))
  return { beta, bSE, tS, pV, r2, r2Adj, se, yHat, residuals: res, n }
}

// ─── Ridge regression ─────────────────────────────────────────────────────────
// X should already have intercept as first column
// alpha penalizes all coefficients except intercept
export function ridge(X, y, alpha = 1.0) {
  const n = y.length, k = X[0].length
  if (n < k + 2) return null
  // XtX + alpha * I (don't penalize intercept → index 0)
  const XtX = Array.from({ length: k }, (_, i) =>
    Array.from({ length: k }, (_, j) => {
      const base = X.reduce((s, r) => s + r[i] * r[j], 0)
      return base + (i === j && i > 0 ? alpha : 0)
    })
  )
  const inv = matInv(XtX)
  if (!inv) return null
  const Xty = Array.from({ length: k }, (_, i) =>
    X.reduce((s, r, ri) => s + r[i] * y[ri], 0)
  )
  const beta = inv.map(row => row.reduce((s, v, j) => s + v * Xty[j], 0))
  const yHat = X.map(r => r.reduce((s, v, i) => s + v * beta[i], 0))
  const res = y.map((v, i) => v - yHat[i])
  const sse = res.reduce((s, r) => s + r * r, 0)
  const ym = avg(y), sst = y.reduce((s, v) => s + (v - ym) ** 2, 0)
  const r2 = sst < 1e-14 ? 0 : 1 - sse / sst
  const r2Adj = 1 - (1 - r2) * (n - 1) / (n - k)
  const mse = sse / Math.max(n - k, 1), se = Math.sqrt(mse)
  return { beta, r2, r2Adj, se, yHat, residuals: res, n }
}

// ─── Ridge alpha selection via k-fold CV ──────────────────────────────────────
export function selectRidgeAlpha(X, y, alphas = null, k = 5) {
  if (!alphas) alphas = [0.01, 0.1, 0.5, 1, 2, 5, 10, 20, 50, 100, 200, 500]
  const n = y.length
  const foldSize = Math.floor(n / k)
  let bestAlpha = 1.0, bestMSE = Infinity

  for (const alpha of alphas) {
    let totalMSE = 0, count = 0
    for (let fold = 0; fold < k; fold++) {
      const valStart = fold * foldSize
      const valEnd = fold === k - 1 ? n : valStart + foldSize
      const trainIdx = [...Array(n).keys()].filter(i => i < valStart || i >= valEnd)
      const valIdx = [...Array(n).keys()].filter(i => i >= valStart && i < valEnd)
      if (trainIdx.length < X[0].length + 2) continue
      const Xtr = trainIdx.map(i => X[i]), ytr = trainIdx.map(i => y[i])
      const Xval = valIdx.map(i => X[i]), yval = valIdx.map(i => y[i])
      const res = ridge(Xtr, ytr, alpha)
      if (!res) continue
      const pred = Xval.map(r => r.reduce((s, v, i) => s + v * res.beta[i], 0))
      totalMSE += pred.reduce((s, p, i) => s + (p - yval[i]) ** 2, 0)
      count += yval.length
    }
    if (count > 0 && totalMSE / count < bestMSE) {
      bestMSE = totalMSE / count
      bestAlpha = alpha
    }
  }
  return bestAlpha
}

// ─── Fixed Effects (within estimator) ────────────────────────────────────────
// Returns the within-transformed data and company mean lookup
export function buildFixedEffects(rows, feats, target, companyKey = '_company') {
  // Compute per-company means from this dataset
  const companyMeansY = {}
  const companyMeansX = {}
  const companies = [...new Set(rows.map(r => r[companyKey]))]

  for (const co of companies) {
    const coRows = rows.filter(r => r[companyKey] === co)
    companyMeansY[co] = avg(coRows.map(r => r[target]))
    companyMeansX[co] = feats.map(f => avg(coRows.map(r => r[f])))
  }

  // Demean
  const yWithin = rows.map(r => r[target] - companyMeansY[r[companyKey]])
  const XWithin = rows.map(r => {
    const coMeans = companyMeansX[r[companyKey]] || feats.map(() => 0)
    return feats.map((f, i) => r[f] - coMeans[i])
  })

  return { yWithin, XWithin, companyMeansY, companyMeansX, companies }
}

// Predict using fixed effects model:
// prediction = within_prediction + company_mean_y (from training)
export function predictFE(row, feats, beta, ms, ss, companyMeansY, companyMeansX, companyKey = '_company') {
  const co = row[companyKey]
  const coMeansX = companyMeansX[co] || feats.map(() => 0)
  // Demean using training company means, then standardize
  const demeaned = feats.map((f, i) => (row[f] ?? 0) - coMeansX[i])
  const stdX = [1, ...demeaned.map((v, i) => (v - ms[i]) / ss[i])]
  const withinPred = stdX.reduce((s, v, i) => s + v * beta[i], 0)
  const baseline = companyMeansY[co] ?? avg(Object.values(companyMeansY))
  return withinPred + baseline
}

// ─── Walk-forward cross-validation ────────────────────────────────────────────
export function walkForwardCV(rows, feats, target, method = 'ols', ridgeAlpha = 5) {
  const years = [...new Set(rows.map(r => r._year).filter(isFinite))].sort()
  const results = []

  for (let i = 1; i < years.length; i++) {
    const cutoff = years[i - 1]
    const testYear = years[i]
    const trainRows = rows.filter(r => r._year <= cutoff && isFinite(r[target]) && feats.every(f => isFinite(r[f])))
    const testRows = rows.filter(r => r._year === testYear && isFinite(r[target]) && feats.every(f => isFinite(r[f])))
    if (trainRows.length < feats.length + 3 || testRows.length < 2) continue

    let predictions = []

    if (method === 'fe' || method === 'fe_ridge') {
      const { yWithin, XWithin, companyMeansY, companyMeansX } = buildFixedEffects(trainRows, feats, target)
      const { ms, ss } = standardize(trainRows.map((r, i) => {
        const obj = {}; feats.forEach((f, j) => { obj[f] = XWithin[i][j] }); return obj
      }), feats)
      // Standardize the within-transformed features
      const XwStd = XWithin.map(row => [1, ...row.map((v, i) => (v - ms[i]) / ss[i])])
      const fitRes = method === 'fe_ridge'
        ? ridge(XwStd, yWithin, ridgeAlpha)
        : ols(XwStd, yWithin)
      if (!fitRes) continue

      predictions = testRows.map(r => {
        const coMeansX = companyMeansX[r._company] || feats.map(() => 0)
        const demeaned = feats.map((f, i) => (r[f] ?? 0) - coMeansX[i])
        const stdX = [1, ...demeaned.map((v, i) => (v - ms[i]) / ss[i])]
        const withinPred = stdX.reduce((s, v, i) => s + v * fitRes.beta[i], 0)
        const baseline = companyMeansY[r._company] ?? avg(Object.values(companyMeansY))
        return { pred: withinPred + baseline, actual: r[target], company: r._company, year: r._year }
      })
    } else {
      // Pooled OLS or Ridge
      const { Xs, ms, ss } = standardize(trainRows, feats)
      const y = trainRows.map(r => r[target])
      const fitRes = method === 'ridge'
        ? ridge(Xs, y, ridgeAlpha)
        : ols(Xs, y)
      if (!fitRes) continue
      predictions = testRows.map(r => {
        const x = applyStd(r, feats, ms, ss)
        return { pred: x.reduce((s, v, i) => s + v * fitRes.beta[i], 0), actual: r[target], company: r._company, year: r._year }
      })
    }

    if (!predictions.length) continue
    const actuals = predictions.map(p => p.actual)
    const preds = predictions.map(p => p.pred)
    const yMean = avg(actuals)
    const sst = actuals.reduce((s, v) => s + (v - yMean) ** 2, 0)
    const sse = actuals.reduce((s, v, i) => s + (v - preds[i]) ** 2, 0)
    const r2 = sst < 1e-10 ? 0 : 1 - sse / sst
    const mae = avg(predictions.map(p => Math.abs(p.pred - p.actual)))
    results.push({ cutoff, testYear, r2, mae, n: predictions.length, predictions })
  }
  return results
}

// ─── p-value approximation ────────────────────────────────────────────────────
export function approxP(t, df) {
  if (!isFinite(t) || df <= 0) return 1
  const z = df > 120 ? t : t * (1 - 1 / (4 * df))
  return Math.max(1e-6, Math.min(1, 2 * (1 - normCDF(z))))
}
function normCDF(z) {
  const a = [0.319381530, -0.356563782, 1.781477937, -1.821255978, 1.330274429]
  const t2 = 1 / (1 + 0.2316419 * Math.abs(z))
  let poly = 0, tp = t2
  for (const c of a) { poly += c * tp; tp *= t2 }
  const p = Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI) * poly
  return z >= 0 ? 1 - p : p
}

// ─── Outlier filter ───────────────────────────────────────────────────────────
export function removeOutliers(rows, target, multiplier = 3) {
  const vals = rows.map(r => r[target]).sort((a, b) => a - b)
  const q1 = vals[Math.floor(vals.length * 0.25)]
  const q3 = vals[Math.floor(vals.length * 0.75)]
  const iqr = q3 - q1
  return rows.filter(r => r[target] >= q1 - multiplier * iqr && r[target] <= q3 + multiplier * iqr)
}

// ─── Relative multiple computation ────────────────────────────────────────────
// Computes each row's premium/discount vs the sector median for that year
// Returns enriched rows with _sectorMedian and _relativeMultiple fields
export function computeRelativeMultiple(rows, target) {
  // Group by year, compute median
  const yearGroups = {}
  for (const r of rows) {
    const y = r._year
    if (!isFinite(r[target])) continue
    if (!yearGroups[y]) yearGroups[y] = []
    yearGroups[y].push(r[target])
  }
  const yearMedians = {}
  for (const [yr, vals] of Object.entries(yearGroups)) {
    const sorted = vals.slice().sort((a, b) => a - b)
    const mid = Math.floor(sorted.length / 2)
    yearMedians[yr] = sorted.length % 2 === 0
      ? (sorted[mid - 1] + sorted[mid]) / 2
      : sorted[mid]
  }
  return rows.map(r => ({
    ...r,
    _sectorMedian: yearMedians[r._year] ?? NaN,
    _relativeMultiple: isFinite(r[target]) && isFinite(yearMedians[r._year])
      ? r[target] - yearMedians[r._year]
      : NaN,
  }))
}

// Given a predicted relative multiple + the sector median for a year,
// convert back to absolute multiple
export function relativeToAbsolute(relativePred, sectorMedian) {
  return relativePred + sectorMedian
}
