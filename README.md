# ValuationEngine — Regression-Based Valuation Support

M&A advisory tool for regression-based comparable company analysis.

## Stack
- React 18 + Vite
- Recharts (charts)
- SheetJS / xlsx (Excel parsing)
- No backend — runs entirely in the browser

## Project structure
```
valuation-engine/
├── index.html
├── package.json
├── vite.config.js
└── src/
    ├── main.jsx          # entry point
    ├── App.jsx           # main app + config panel
    ├── math.js           # OLS, Ridge, Fixed Effects, Walk-forward CV
    ├── constants.jsx     # colours, shared UI components
    └── components/
        ├── TabOverview.jsx
        ├── TabAsOf.jsx
        └── Tabs.jsx      # Investment, Validation, Drivers, Forecast, Technical
```

## Deploy to Vercel (instructions for Manus)

1. Install dependencies:
```bash
npm install
```

2. Verify it builds:
```bash
npm run build
```

3. Deploy to Vercel:
```bash
npx vercel --prod
```
When prompted:
- Project name: `valuation-engine`
- Framework: Vite
- Build command: `npm run build`
- Output directory: `dist`
- No environment variables needed

## Features
- **4 regression methods**: Pooled OLS, Ridge, Fixed Effects OLS, Fixed Effects + Ridge
- **Walk-forward validation**: honest out-of-sample test — train on past years, predict next year
- **Historical Accuracy tab**: flat-line as-of perspective with convergence tracking
- **Investment Case tab**: model-implied upside ranking vs peers, re-rating story
- **Scenarios tab**: bear/base/bull by adjusting any fundamental
- **Key Drivers tab**: feature importance ranked by statistical strength
- **Confidence labels**: built-in caveats so outputs aren't misused

## Known limitations
- With 20 companies, out-of-sample R² is low (walk-forward MAE ≈ ±4–5x)
- 2022 predictions are poor due to rate shock — macro events not captured by fundamentals
- Use as directional signal only — not for anchoring specific valuation numbers
- Fixed Effects method is recommended for panel data

## Data format
Excel file with sheet named `Annualized_Panel` containing:
- `Company Name` (string)
- `As-of Year` (number)
- Financial metrics as numeric columns
- Qualitative scores as numeric columns (1-5 scales)
