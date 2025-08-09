
import streamlit as st
import pandas as pd
import math
from io import BytesIO

st.set_page_config(page_title="Shotcraft Case-Based Inventory (Upload + Edit On_Hand)", layout="wide")
st.title("Shotcraft — Upload, Edit On_Hand, Calculate")
st.caption("Upload your Excel, adjust On_Hand in-app, enter Cases Sold, and see Required & Remaining per material.")

@st.cache_data(show_spinner=False)
def read_excel(file_bytes):
    xls = pd.ExcelFile(file_bytes)
    return xls.sheet_names, {name: pd.read_excel(xls, sheet_name=name) for name in xls.sheet_names}

def load_components_and_onhand(sheets_dict):
    # Try to find a formula-like sheet
    pick = None
    for name in sheets_dict.keys():
        if name.lower() in ("formula_695_cases","formula"):
            pick = name
            break
    if pick is None:
        pick = list(sheets_dict.keys())[0]
    df_formula = sheets_dict[pick].copy()

    # Normalize columns
    cols = {c.lower().strip(): c for c in df_formula.columns}
    comp_col = cols.get("component") or list(df_formula.columns)[0]
    per_case_col = cols.get("per_case")
    uom_col = cols.get("uom")
    if per_case_col is None:
        st.error("Couldn't find a Per_Case column in your formula sheet. Please include Per_Case.")
        return None, None, None

    components = df_formula[[comp_col, per_case_col]].copy()
    components.columns = ["Component","Per_Case"]
    components["UOM"] = df_formula[uom_col] if uom_col else ""

    # On_Hand from INVENTORY sheet if present
    onhand_df = None
    for name, df in sheets_dict.items():
        if name.lower() == "inventory":
            inv_cols = {c.lower().strip(): c for c in df.columns}
            if "component" in inv_cols and "on_hand" in inv_cols:
                onhand_df = df[[inv_cols["component"], inv_cols["on_hand"]]].rename(columns={inv_cols["component"]:"Component", inv_cols["on_hand"]:"On_Hand"})
            break

    return components.reset_index(drop=True), (onhand_df.reset_index(drop=True) if onhand_df is not None else None), pick

def compute_results(components, onhand_df, cases_sold):
    df = components.copy()
    if onhand_df is not None:
        df = df.merge(onhand_df, on="Component", how="left")
    if "On_Hand" not in df.columns:
        df["On_Hand"] = 0.0

    df["Per_Case"] = pd.to_numeric(df["Per_Case"], errors="coerce").fillna(0.0)
    df["On_Hand"] = pd.to_numeric(df["On_Hand"], errors="coerce").fillna(0.0)

    df["Required"] = df["Per_Case"] * float(cases_sold)
    df["Remaining"] = df["On_Hand"] - df["Required"]

    candidates = df[(df["Per_Case"]>0)]
    if not candidates.empty:
        df["MaxCasesByItem"] = df.apply(lambda r: (r["On_Hand"]/r["Per_Case"]) if r["Per_Case"]>0 else float("inf"), axis=1)
        max_sellable = math.floor(df["MaxCasesByItem"].min())
    else:
        max_sellable = 0

    shortages = df[df["Remaining"] < 0][["Component","UOM","On_Hand","Per_Case","Required","Remaining"]].copy()
    display = df[["Component","UOM","On_Hand","Per_Case","Required","Remaining"]].sort_values("Component")
    return display, int(max_sellable), shortages

def make_snapshot(formula_name, display_df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        formula = display_df[["Component","UOM","Per_Case"]].copy()
        formula.to_excel(writer, sheet_name=formula_name or "FORMULA", index=False)
        inv = display_df.copy()
        inv.to_excel(writer, sheet_name="INVENTORY", index=False)
    out.seek(0)
    return out

# --- UI ---
uploaded = st.file_uploader("Upload your Shotcraft Excel (.xlsx)", type=["xlsx"])

if "edited_onhand" not in st.session_state:
    st.session_state.edited_onhand = None

if uploaded is not None:
    sheet_names, sheets = read_excel(uploaded)
    comps, onhand, formula_name = load_components_and_onhand(sheets)
    if comps is None:
        st.stop()

    st.success(f"Loaded {len(comps)} components from sheet: {formula_name}")
    st.write("Per-case usage (from your file):")
    st.dataframe(comps, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("Step 1 — Edit your On_Hand (only this column is editable)")

    # Build an editable table seeded with any On_Hand we found
    base = comps.merge(onhand if onhand is not None else pd.DataFrame(columns=["Component","On_Hand"]),
                       on="Component", how="left")
    base["On_Hand"] = pd.to_numeric(base["On_Hand"], errors="coerce").fillna(0.0)

    edited = st.data_editor(
        base[["Component","UOM","On_Hand","Per_Case"]],
        hide_index=True,
        column_config={
            "Component": st.column_config.TextColumn(disabled=True),
            "UOM": st.column_config.TextColumn(disabled=True),
            "Per_Case": st.column_config.NumberColumn(format="%.6f", disabled=True),
            "On_Hand": st.column_config.NumberColumn(help="Type your current stock here"),
        },
        use_container_width=True,
        key="edit_table"
    )
    # Save edited On_Hand to session
    st.session_state.edited_onhand = edited[["Component","On_Hand"]].copy()

    st.subheader("Step 2 — Enter order size (cases)")
    cases = st.number_input("Cases Sold (e.g., LCBO order)", min_value=0.0, step=1.0, value=0.0)

    # Compute using edited on hand
    results, max_sellable, shortages = compute_results(comps, st.session_state.edited_onhand, cases)

    st.markdown("### Results")
    c1, c2 = st.columns(2)
    with c1:
        st.metric("Max sellable cases from current stock", max_sellable)
    with c2:
        st.metric("Order size entered (cases)", int(cases))

    st.dataframe(results, use_container_width=True, hide_index=True)

    if not shortages.empty:
        st.warning("Shortages for this order:")
        st.dataframe(shortages, use_container_width=True, hide_index=True)
    else:
        st.info("No shortages detected for this order.")

    st.markdown("### Download updated snapshot")
    buf = make_snapshot(formula_name, results)
    st.download_button("Download Excel snapshot", buf, file_name="Shotcraft_Inventory_Snapshot.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Upload your Excel to begin.")
