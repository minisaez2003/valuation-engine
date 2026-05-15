# ValuationEngine

Regression-based comparable company valuation tool for M&A advisory.

## Quick Start

```bash
npm install
npm run dev
```

Then open `http://localhost:5173` and drop your Excel file.

## Data Format

Excel file with a sheet named `Annualized_Panel` (or first sheet used as fallback). Required columns:

- `Company Name`
- `As-of Year`

Plus any numeric columns you want as target/features. The tool auto-detects numeric columns.

## Deployment to Vercel

```bash
npm run build
vercel
```

### Enabling AI Chat (Optional)

The AI chatbot needs a backend proxy to keep your Anthropic API key secure.

1. Get an API key at [console.anthropic.com](https://console.anthropic.com)
2. In your Vercel project settings → Environment Variables, add:
   ```
   ANTHROPIC_API_KEY = sk-ant-...
   ```
3. The `/api/chat.js` serverless function included in this repo will handle requests automatically. Your key stays server-side; the browser never sees it.

If you skip this step, the app still works fine — only the chat panel will show a setup message.

## Mobile

The app is fully responsive. Headers wrap, grids collapse to single column, chat panel goes near-fullscreen on phones.

## Methods

- **Pooled OLS** — standard cross-sectional regression
- **Ridge** — handles multicollinearity (α = 5)
- **Fixed Effects** — controls for company baseline (panel data)
- **FE + Ridge** — most robust for small datasets

## Validation

- **Walk-forward CV** — trains on year ≤ T, predicts year T+1 only (honest out-of-sample)

## Features

- Auto-detects target & feature columns from any Excel file
- Relative-to-sector mode (removes macro noise)
- Smart valuation ranges capped per target type
- AI analyst chatbot (with secure backend proxy)
- Apollo-style PDF report export (Share Analysis)
- Dark/light theme
- Mobile responsive
