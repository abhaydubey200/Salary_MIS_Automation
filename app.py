"""
APIS India Limited — Monthly Salary MIS (HO & Sales)
Streamlit port of the browser tool. Deployable on streamlit.io.

Run locally:   streamlit run app.py
"""

from __future__ import annotations

import datetime as _dt

import pandas as pd
import streamlit as st

import mis_core as C
import mis_export as X

st.set_page_config(page_title="APIS Salary MIS", page_icon="📊", layout="wide")

# ---------------------------------------------------------------------------
# Brand styling (deep brown header + amber accent, mirrors the original tool)
# ---------------------------------------------------------------------------
st.markdown("""
<style>
:root{--apis-brown:#2b1d0e;--apis-amber:#c27c00;}
.apis-top{display:flex;justify-content:space-between;align-items:center;gap:12px;
  flex-wrap:wrap;background:#2b1d0e;color:#fff;border-bottom:4px solid #f2b84b;
  border-radius:10px;padding:14px 18px;margin-bottom:8px}
.apis-brand{display:flex;align-items:center;gap:12px}
.apis-logo{background:#f2b84b;color:#2b1d0e;font-weight:800;letter-spacing:1px;
  border-radius:8px;padding:8px 10px;font-size:18px}
.apis-brand b{display:block;font-size:16px;letter-spacing:.6px}
.apis-brand small{color:#e8d9b8}
.apis-top .t{font-size:15px;font-weight:600;color:#f2b84b}
.stApp h2,.stApp h3{border-left:4px solid var(--apis-amber);padding-left:8px}
div[data-testid="stMetric"]{border:1px solid #e3d9c6;border-left:4px solid var(--apis-amber);
  border-radius:8px;padding:8px 12px;background:rgba(242,184,75,.04)}
.msg{padding:8px 10px;border-radius:6px;margin:5px 0;font-size:13px;border:1px solid #e3d9c6}
.msg.ok{color:#157a3d;background:rgba(21,122,61,.06)}
.msg.wr{color:#9a5b00;background:rgba(154,91,0,.06)}
.msg.er{color:#b42318;background:rgba(180,35,24,.06)}
</style>
<div class="apis-top"><div class="apis-brand"><span class="apis-logo">APIS</span>
<div><b>APIS INDIA LIMITED</b><small>People &amp; Culture · Payroll &amp; Compliances</small></div></div>
<div class="t">Monthly Salary MIS — HO &amp; Sales</div></div>
""", unsafe_allow_html=True)

st.caption("Upload the HO and Sales salary sheets (same format every month). "
           "The MIS is built in the approved layout, checked, and exported to Excel.")

# ---------------------------------------------------------------------------
# Session defaults
# ---------------------------------------------------------------------------
ss = st.session_state
ss.setdefault("hold", None)          # set[str] once built
ss.setdefault("hold_touched", False)


def _msg(kind: str, text: str):
    st.markdown(f'<div class="msg {kind}">{text}</div>', unsafe_allow_html=True)


def _title(d: _dt.date) -> dict:
    mon = C.MON[d.month - 1]
    y = str(d.year)
    return {"mon": mon, "y": y, "full": f"{mon}, {y}", "sheet": f"{mon}, {y[2:]}"}


def _dedupe(cols):
    seen, out = {}, []
    for c in cols:
        if c in seen:
            seen[c] += 1
            out.append(f"{c} ({seen[c]})")
        else:
            seen[c] = 0
            out.append(c)
    return out


def _mis_dataframe(M: C.Model, mode: str) -> pd.DataFrame:
    cache: dict = {}
    headers = _dedupe([C.NH[k] for k in C.NC])   # 18 unique display headers (B..S)
    rows, index = [], []
    for l in M.L:
        if l is M.sRows[0]:
            rows.append([""] * len(C.NC))
            index.append("── SALES ──")
        v = C.line_vals(M, l, mode, cache)
        cols = C.cols_of(l)
        rows.append([C.fmt_in(v[k]) if k in cols else "" for k in C.NC])
        index.append(l.label)
    return pd.DataFrame(rows, index=index, columns=headers)


# ---------------------------------------------------------------------------
# 1. Upload
# ---------------------------------------------------------------------------
st.subheader("1. Upload salary sheets")
c1, c2 = st.columns(2)
with c1:
    ho_file = st.file_uploader("HO Employee's Salary Sheet (.xlsx)", type=["xlsx", "xlsm"], key="ho_up")
with c2:
    sales_file = st.file_uploader('Sales Salary Sheet (.xlsx, sheet "APIS")', type=["xlsx", "xlsm"], key="sl_up")

c3, c4, c5, c6 = st.columns([1, 1, 1, 1])
with c3:
    # Default month from filenames if detectable
    default_month = ""
    if ho_file:
        default_month = C.detect_month(ho_file.name)
    if not default_month and sales_file:
        default_month = C.detect_month(sales_file.name)
    today = _dt.date.today()
    if default_month:
        y, m = default_month.split("-")
        init_date = _dt.date(int(y), int(m), 1)
    else:
        init_date = today.replace(day=1)
    mis_date = st.date_input("MIS month", value=init_date, format="YYYY-MM-DD")
    month_val = f"{mis_date.year}-{str(mis_date.month).zfill(2)}"
with c4:
    fset_label = st.selectbox("Formula set",
                              ["Corrected (Variable counted once)", "Senior's template (as given)"])
    mode = "fix" if fset_label.startswith("Corrected") else "tpl"
with c5:
    pw1 = st.text_input("File password (if protected)", type="password", key="pw1")
with c6:
    pw2 = st.text_input("2nd password (optional)", type="password", key="pw2")

# ---------------------------------------------------------------------------
# 2. Manual inputs
# ---------------------------------------------------------------------------
st.subheader("2. Manual inputs — not part of the salary sheets (optional)")
m1, m2, m3, m4, m5 = st.columns(5)
with m1:
    jk = st.number_input("Promotor (JK) payout", min_value=0.0, value=0.0, step=1000.0, format="%.0f")
with m2:
    jkm = st.number_input("JK manpower count", min_value=0.0, value=0.0, step=1.0, format="%.0f")
with m3:
    hra = st.number_input("HRA", min_value=0.0, value=0.0, step=1000.0, format="%.0f")
with m4:
    hire = st.number_input("Hiring fees", min_value=0.0, value=0.0, step=1000.0, format="%.0f")
with m5:
    dirc = st.number_input("Directors: manual Arrears/Var/Reim", value=0.0, step=1000.0, format="%.0f")
st.caption("Leave blank if not available — the Sales Total then excludes this block and the "
           "export carries a visible warning.")

manual = C.ManualInputs(jk=jk, jkm=jkm, hra=hra, hire=hire, dirc=dirc)

build = st.button("Submit & Build MIS", type="primary", disabled=not (ho_file and sales_file))

# ---------------------------------------------------------------------------
# Parse (cached on file bytes + passwords)
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _parse(data: bytes, name: str, kind: str, pws: tuple[str, ...]):
    return C.load_sheet(data, name, kind, list(pws))


def parse_both():
    errs = {}
    ho_p = sl_p = None
    try:
        ho_p = _parse(ho_file.getvalue(), ho_file.name, "ho", (pw1, pw2))
    except Exception as e:  # noqa: BLE001
        errs["ho"] = str(e)
    try:
        sl_p = _parse(sales_file.getvalue(), sales_file.name, "sales", (pw1, pw2))
    except Exception as e:  # noqa: BLE001
        errs["sales"] = str(e)
    return ho_p, sl_p, errs


if build or (ho_file and sales_file and ss.get("built")):
    if build:
        ss["built"] = True
        ss["hold"] = None
        ss["hold_touched"] = False

if ss.get("built") and ho_file and sales_file:
    ho_p, sl_p, errs = parse_both()

    for k, lbl in (("ho", "HO"), ("sales", "Sales")):
        if k in errs:
            _msg("er", f"<b>{lbl} sheet failed:</b> {errs[k]}")

    if ho_p and sl_p:
        _msg("ok", f"HO: {len(ho_p.emps)} employees (sheet \"{ho_p.sheet}\") · "
                   f"Sales: {len(sl_p.emps)} employees (sheet \"{sl_p.sheet}\")")

        # Build the auto-hold set on first build; keep user edits afterwards.
        all_e = ho_p.emps + sl_p.emps
        if ss["hold"] is None:
            ss["hold"] = {e.code for e in all_e if C.HOLD_RE.search(e.rem or "")}

        # ---------------- Hold review ----------------
        rem_emps = [e for e in all_e if e.rem]
        if rem_emps:
            st.subheader("Hold review")
            st.caption("Employees with a Remark in the salary sheet. Resigned / Absconding / "
                       "Hold / ATL / Terminate / Accident are held by default (Net Pay withheld). "
                       "Tick/untick to match payroll's actual holds.")
            hdf = pd.DataFrame([{
                "Hold": e.code in ss["hold"], "Code": e.code, "Name": e.name,
                "Department": e.dept, "Remark": e.rem, "Net Pay": round(e.net),
            } for e in rem_emps])
            edited = st.data_editor(
                hdf, hide_index=True, width="stretch", key="hold_editor",
                column_config={
                    "Hold": st.column_config.CheckboxColumn("Hold", width="small"),
                    "Net Pay": st.column_config.NumberColumn("Net Pay", format="%d"),
                },
                disabled=["Code", "Name", "Department", "Remark", "Net Pay"])
            ss["hold"] = {row["Code"] for _, row in edited.iterrows() if row["Hold"]}

        hold = ss["hold"]

        # ---------------- Build model + checks ----------------
        M = C.build_model(ho_p.emps, sl_p.emps, hold, manual)
        checks = C.run_checks(M, ho_p, sl_p, hold, manual, mode, month_val)
        title = _title(mis_date)

        # ---------------- KPIs ----------------
        D = C.dash_data(M, ho_p, sl_p, hold, manual)
        G, SS, tot = D["G"], D["SS"], D["tot"]
        st.subheader("Key figures")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total Manpower", f"{int(tot('M'))}")
        k2.metric("Total Gross", C.fmt_in(tot("Q")))
        k3.metric("Total Payout", C.fmt_in(tot("R")))
        k4.metric("Hold (withheld)", C.fmt_in(tot("K") + tot("P")), f"{D['heldN']} employees")

        # ---------------- Tabs ----------------
        t_over, t_dept, t_sales, t_ded, t_chk, t_mis = st.tabs(
            ["Overview", "HO Departments", "Sales", "Deductions & Hold", "Checks", "MIS Report"])

        with t_over:
            o1, o2 = st.columns(2)
            with o1:
                st.markdown("**Total payout by HO department**")
                ho_df = pd.DataFrame(D["ho"]).rename(
                    columns={"label": "Department", "R": "Total Payout"}).set_index("Department")
                st.bar_chart(ho_df["Total Payout"], horizontal=True, color="#c27c00")
            with o2:
                st.markdown("**HO vs Sales total payout**")
                st.bar_chart(pd.DataFrame({"Total Payout": {"HO & Directors": G["R"], "Sales": SS["R"]}}),
                             horizontal=True, color="#2b1d0e")
                st.markdown("**Deductions & hold (HO + Sales)**")
                ded_df = pd.DataFrame(D["ded"], columns=["Deduction", "Amount"])
                ded_df.loc[len(ded_df)] = ["Hold (withheld)", tot("K") + tot("P")]
                st.bar_chart(ded_df.set_index("Deduction")["Amount"], horizontal=True, color="#9a5b00")

        with t_dept:
            rows = [{"Department": x["label"], "Manpower": int(x["M"]),
                     "Net Earning": round(x["B"]), "Total Gross": round(x["Q"]),
                     "Total Payout": round(x["R"]),
                     "Avg payout/head": round(x["R"] / x["M"]) if x["M"] else 0,
                     "Share": f"{x['R'] / G['R'] * 100:.1f}%" if G["R"] else "-"}
                    for x in D["ho"]]
            rows.append({"Department": "Total HO & Directors", "Manpower": int(G["M"]),
                         "Net Earning": round(G["B"]), "Total Gross": round(G["Q"]),
                         "Total Payout": round(G["R"]),
                         "Avg payout/head": round(G["R"] / G["M"]) if G["M"] else 0, "Share": "100%"})
            st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")

        with t_sales:
            rows = [{"Channel": x["label"], "Manpower": int(x["M"]), "Net Earning": round(x["B"]),
                     "Total Gross": round(x["Q"]), "Total Payout": round(x["R"]),
                     "Avg payout/head": round(x["R"] / x["M"]) if x["M"] else 0,
                     "Hold": round(x["K"])} for x in D["sl"]]
            rows.append({"Channel": "Total (A)", "Manpower": int(SS["M"]), "Net Earning": round(SS["B"]),
                         "Total Gross": round(SS["Q"]), "Total Payout": round(SS["R"]),
                         "Avg payout/head": round(SS["R"] / SS["M"]) if SS["M"] else 0,
                         "Hold": round(C.n(SS.get("K")) + C.n(SS.get("P")))})
            st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
            st.caption("Promotor (JK): " + ("entered — included in MIS Total (B) and Total."
                       if C.n(jk) else "not entered — Sales Total (B) and Total exclude it."))

        with t_ded:
            d1, d2 = st.columns(2)
            with d1:
                st.markdown("**Deductions**")
                dsum = sum(x[1] for x in D["ded"])
                rows = [{"Deduction": a, "Amount": round(b),
                         "% of gross": f"{b / D['gross'] * 100:.1f}%" if D["gross"] else "-"}
                        for a, b in D["ded"]]
                rows.append({"Deduction": "Total deductions", "Amount": round(dsum),
                             "% of gross": f"{dsum / D['gross'] * 100:.1f}%" if D["gross"] else "-"})
                st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
            with d2:
                st.markdown("**Hold by reason**")
                if D["hold"]:
                    rows = [{"Hold reason": r[0], "Employees": r[1], "Net pay held": round(r[2])}
                            for r in D["hold"]]
                    rows.append({"Hold reason": "Total held", "Employees": D["heldN"],
                                 "Net pay held": round(sum(r[2] for r in D["hold"]))})
                    st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
                else:
                    st.caption("No employees held.")

        with t_chk:
            for c in checks:
                _msg("ok" if c.ok else "wr",
                     ("✔ " if c.ok else "⚠ ") + c.t + (" — " + c.d if c.d else ""))

        with t_mis:
            if mode == "tpl":
                d = C.tpl_diff(M)
                _msg("wr", "Formula set: <b>senior's template (as given)</b>. HO Total Gross is "
                           f"{C.fmt_in(d['hoP'])} higher (Variable counted twice) and Sales Net "
                           f"Payout / Total Payout are {C.fmt_in(d['sQ'])} lower (Variable left out "
                           "of Sales Gross Earning). Choose <b>Corrected</b> to remove both.")
            else:
                _msg("ok", "Formula set: corrected — Variable is counted once in every row.")

            mis_df = _mis_dataframe(M, mode)
            st.dataframe(mis_df, width="stretch", height=560)

        # ---------------- Exports ----------------
        st.subheader("Export")
        e1, e2 = st.columns(2)
        with e1:
            mis_bytes = X.build_mis_workbook(M, ho_p, sl_p, hold, manual, mode, title, checks)
            st.download_button("⬇ Export MIS to Excel", data=mis_bytes,
                               file_name=f"Salary_MIS_HO_Sales_{title['mon']}_{title['y']}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               type="primary")
        with e2:
            dash_bytes = X.build_dashboard_workbook(M, ho_p, sl_p, hold, manual, title)
            st.download_button("⬇ Export Dashboard Summary (Excel)", data=dash_bytes,
                               file_name=f"Dashboard_Summary_{title['mon']}_{title['y']}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("<p style='text-align:center;color:#8a8a8a;margin-top:24px'>APIS India Limited · "
            "Confidential — for internal use only</p>", unsafe_allow_html=True)
