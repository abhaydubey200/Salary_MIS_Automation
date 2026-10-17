# APIS India Limited — Monthly Salary MIS (HO & Sales)

Streamlit app that builds the monthly Salary MIS from the HO and Sales salary
sheets, runs the full set of payroll cross-checks, shows a dashboard, and
exports the approved-format MIS workbook (with live Excel formulas) plus a
dashboard-summary workbook.

This is a 1:1 port of the original browser tool
(`APIS_India_Limited___Salary_MIS_Automation.html`) to Python + Streamlit so it
can be deployed on **streamlit.io** (Streamlit Community Cloud) instead of being
opened as a local HTML file.

## What it does

1. **Upload** the HO and Sales salary sheets (`.xlsx`). Password-protected
   files are supported — enter the password(s) in the two password boxes.
2. **Parse** — the header row is located by finding *Employee Code* in column B,
   and every column is mapped by header name (so the tool keeps working if a
   column shifts position month to month).
3. **Build the MIS** in the approved layout: HO by department (fixed 13-dept
   order + any new departments), Directors as block B, Sales by channel
   (Alternate Channels / GT + any new ones), and the manual memo block
   (Promotor JK, HRA, Hiring Fees, Service Charges @3%, EPF @13%, ESIC @3.25%).
4. **Two formula sets** — *Corrected (Variable counted once)* and
   *Senior's template (as given)*.
5. **Hold review** — rows remarked Resigned / Absconding / Hold / ATL /
   Terminate / Accident are held by default; tick/untick to match payroll.
6. **Cross-checks** — every tie-out (gross = earnings + variable + arrears, net
   = gross − deductions, rows tie to the sheet's own total row, employer PF/ESIC
   tie-outs, duplicates, negatives, junk text in numeric columns, ESIC-rule
   inference, etc.).
7. **Dashboard** — Overview, HO Departments, Sales, Deductions & Hold.
8. **Export** — the MIS workbook (MIS + HOLD&TDS + Checks sheets, with live
   formulas) and a Dashboard Summary workbook.

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit UI (upload, inputs, hold review, tabs, exports) |
| `mis_core.py` | Parsing, the MIS model, cross-checks, dashboard data |
| `mis_export.py` | Excel exporters (MIS workbook + Dashboard summary) |
| `requirements.txt` | Python dependencies |
| `.streamlit/config.toml` | Theme + 50 MB upload limit |

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open http://localhost:8501.

## Deploy on streamlit.io

1. Create a **GitHub repo** and push these files (keep the folder structure,
   including `.streamlit/config.toml`).
2. Go to <https://share.streamlit.io>, sign in with GitHub, click
   **New app**.
3. Pick the repo, branch, and set **Main file path** to `app.py`.
4. Click **Deploy**. Streamlit Cloud reads `requirements.txt` automatically.

> **Note on password-protected sheets:** decryption happens on the Streamlit
> server (via `msoffcrypto-tool`). On Streamlit Community Cloud the uploaded
> file and password are processed in the app's server session and not stored.
> If your payroll policy requires files never to leave your own machine, run the
> app locally (the command above) instead of on the public cloud, or remove the
> password in Excel before uploading.

## Notes

- All money is shown with Indian digit grouping (e.g. `92,83,564`).
- The exported MIS keeps **live formulas** (SUM ranges, per-row derivations,
  memo-block % formulas), so the workbook recalculates in Excel.
- Employer PF / ESIC are carried as employer-contribution columns (ER PF /
  ER ESIC) and are not part of the payout totals, exactly as in the approved MIS.
