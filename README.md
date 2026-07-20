# El Paso Case Validator (Web version)

A Streamlit web app that compares scraped court-case HTML files against Excel
output files and (optionally) a client input list. This replaces the original
Tkinter desktop app so it can run headlessly on Render.

The validation logic itself (`core.py`) is unchanged from the desktop version —
only the UI layer was rewritten.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the URL Streamlit prints (usually http://localhost:8501).

## Deploy on Render

1. Push this folder to a GitHub repo.
2. In Render: **New +** → **Blueprint**, point it at the repo. Render will read
   `render.yaml` and set everything up automatically (build command, start
   command, Python version).
   - Alternatively, create a **Web Service** manually with:
     - Build Command: `pip install -r requirements.txt`
     - Start Command: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
3. Deploy. Render assigns a public URL you can share.

## Using the app

1. Upload the `list_*.html` and `detail_*.html` files (HTML files).
2. Upload the Excel output file(s) (`.xlsx`).
3. Optionally upload a Client Input file.
4. Pick which fields to validate (all selected by default).
5. Click **Run Validation**.
6. Review results in the **Results** tab, and download the full formatted
   workbook (same sheets/coloring as the desktop version: HTML, Excel, Result,
   Observation, FillRate_Comparison, ClientInput, ClientInput_vs_Excel).

## Notes on differences from the desktop app

- No local folder browsing — files are uploaded through the browser instead.
- `os.startfile(...)` (Windows-only "open file after saving") was removed;
  use the **Download Result Workbook** button instead.
- Minor bug fix: the desktop version's Excel color-coding for the mismatch
  detail sheet checked for a sheet named `"Errors"`, but the sheet is actually
  written as `"Observation"`, so that coloring never ran. This version checks
  the correct name.
