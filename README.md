# Monster 16oz Single — Sales Forecast Tool

A Flask web app for RGM sales forecasting. Upload last year's data, configure
a growth rate and promo plan, and generate a formatted Excel forecast in one click.

---

## Run locally

```bash
pip install -r requirements.txt
python app.py
```
Then open http://localhost:5000

---

## Deploy to Render (free, ~5 min)

1. Push this folder to a GitHub repo
2. Go to https://render.com and sign up (free)
3. Click **New → Web Service**
4. Connect your GitHub repo
5. Set these fields:
   - **Build command:** `pip install -r requirements.txt`
   - **Start command:** `gunicorn app:app`
   - **Environment:** Python 3
6. Click **Deploy**

Render gives you a public URL like `https://your-app.onrender.com` that you can
share with anyone on your team.

---

## Features

- Upload LY Excel or CSV — auto-filters for Monster 16oz Single
- Accepts month as number (1-12) or full name (January)
- Handles styled Excel files with banner rows (like the cleaned output)
- Base growth rate applied to all 12 months preserving seasonality
- Promo lift engine: auto-derives lift from LY data for known promo types
- Manual lift entry for new promo types with no LY history
- Live preview: bar chart, monthly detail table, 6 KPI cards
- One-click Excel generation: Dashboard + Monthly Detail + Promo Summary tabs

---

## File format accepted

Your LY file needs:
- `Month` column — number (1-12) or name (January, Jan)
- `Package` column — filters automatically for "Monster 16oz Single"
- `Forecast Volume` column
- `Actual + Forecast` column (optional)
