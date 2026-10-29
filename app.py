"""
z/OS Product Compatibility Dashboard (Streamlit)

Tab 1  z/OS Version      : choose the z/OS version -> compatibility matrix for IBM, Broadcom, BMC,
                           Rocket, Precisely, ASG (compatible? PTF required? EOS, latest version) -> Excel
Tab 2  Check my products : pick vendor / product / your existing VRM (or upload an inventory) -> verdict -> Excel
Tab 3  Ask               : plain-English questions answered from the matrix rows
Tab 4  Sources           : official vendor links (all clickable)
Maintenance details      : fetches linked vendor pages for PTF / USERMOD / SYSMOD / prerequisite text

Run locally:   pip install -r requirements.txt
               streamlit run app.py
Written answers (optional): IBM watsonx.ai Granite model. Set WATSONX_APIKEY and WATSONX_PROJECT_ID
as environment variables / Code Engine secrets (never commit them). Without them the app shows matching rows.
"""

import os
from datetime import datetime

import pandas as pd
import streamlit as st

from compatibility_db import (
    SOURCES, VENDORS, VER, ZOS_VERSIONS, CompatibilityDB,
    build_assessment_excel, build_excel, to_output,
)
from rag_engine import RagIndex, generate_answer, watsonx_configured
from vendor_details import enrich_vendor_details
from zos_learning_lab import render_learning_lab

st.set_page_config(page_title="z/OS Product Compatibility", page_icon="🖥️", layout="wide")

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
LINK_COL = st.column_config.LinkColumn("Open link", display_text="Open ↗")
ALL_STATUS = ["Yes", "No", "In Progress", "Unverified", "N/A", "Unknown"]

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');
    :root {
        --cds-blue: #0f62fe;
        --cds-blue-hover: #0353e9;
        --cds-text: #161616;
        --cds-muted: #525252;
        --cds-border: #e0e0e0;
        --cds-layer: #f4f4f4;
    }
    html, body, [class*="st-"], [data-testid="stAppViewContainer"] {
        font-family: "IBM Plex Sans", "Helvetica Neue", Arial, sans-serif;
    }
    .stApp { background: #f4f4f4; color: var(--cds-text); }
    [data-testid="stHeader"] { background: rgba(244, 244, 244, 0.92); }
    [data-testid="stSidebar"] { background: #ffffff; border-right: 1px solid var(--cds-border); }
    [data-testid="stSidebar"] [data-testid="stMetric"] {
        background: var(--cds-layer); border-left: 3px solid var(--cds-blue);
        padding: 0.85rem 1rem; border-radius: 0;
    }
    .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1600px; }
    .dashboard-hero {
        background: #161616; color: #ffffff; padding: 1.6rem 2rem;
        border-left: 5px solid var(--cds-blue); margin: 0 0 1.5rem 0;
    }
    .dashboard-hero .eyebrow {
        color: #a6c8ff; font-size: 0.75rem; font-weight: 600;
        letter-spacing: 0.1em; text-transform: uppercase; margin-bottom: 0.55rem;
    }
    .dashboard-hero h1 {
        color: #ffffff; font-size: 2rem; line-height: 1.2; font-weight: 500;
        margin: 0 0 0.45rem 0;
    }
    .dashboard-hero p { color: #c6c6c6; margin: 0; font-size: 0.95rem; }
    .section-kicker {
        color: var(--cds-muted); font-size: 0.75rem; font-weight: 600;
        letter-spacing: 0.08em; text-transform: uppercase; margin: 0 0 0.35rem 0;
    }
    [data-testid="stMetric"] {
        background: #ffffff; border: 1px solid var(--cds-border);
        border-radius: 0; padding: 1rem 1.1rem; min-height: 108px;
    }
    [data-testid="stMetricLabel"] { color: var(--cds-muted); font-size: 0.85rem; }
    [data-testid="stMetricValue"] { color: var(--cds-text); font-weight: 600; }
    [data-testid="stTabs"] [role="tablist"] {
        gap: 0.35rem; border-bottom: 1px solid var(--cds-border);
    }
    [data-testid="stTabs"] button[role="tab"] {
        color: var(--cds-muted); font-weight: 500; padding: 0.8rem 1rem;
    }
    [data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
        color: var(--cds-blue); border-bottom: 2px solid var(--cds-blue);
    }
    [data-testid="stDownloadButton"] button, [data-testid="stLinkButton"] a {
        border-radius: 0;
    }
    [data-testid="stDataFrame"] { border: 1px solid var(--cds-border); }
    </style>
    """,
    unsafe_allow_html=True,
)


# ------------------------------------------------------------------ data
@st.cache_resource
def load_base() -> CompatibilityDB:
    return CompatibilityDB.load(DATA_DIR)


def get_db() -> CompatibilityDB:
    return load_base().merged_with(st.session_state.get("extra"))


def show_table(d: pd.DataFrame, version: str, height: int = 420):
    output = to_output(d, with_link_column=True, version=version)
    styles = pd.DataFrame("", index=output.index, columns=output.columns)
    status_column = f"R{version} Compatible (Yes/No)"
    status_styles = {
        "Yes": "background-color: #c6efce; color: #155724;",
        "No": "background-color: #ffc7ce; color: #721c24;",
        "In Progress": "background-color: #ffeb9c; color: #665500;",
        "Unverified": "background-color: #e5e7eb; color: #374151;",
        "Unknown": "background-color: #e5e7eb; color: #374151;",
        "N/A": "background-color: #f3f4f6; color: #525252;",
    }
    for status, style in status_styles.items():
        styles.loc[output[status_column] == status, status_column] = style
    ptf_column = "If Yes, PTF Required?"
    needs_ptf = output[ptf_column].astype(str).str.startswith("Yes")
    styles.loc[needs_ptf, ptf_column] = "background-color: #ffdf6e; color: #161616; font-weight: 600;"

    st.dataframe(
        output.style.apply(lambda _: styles, axis=None),
        width="stretch", hide_index=True, height=height,
        column_config={"Open link": LINK_COL,
                       "Comments / Links": st.column_config.TextColumn(width="large"),
                       "PTF / USERMOD / SYSMOD / Prerequisites": st.column_config.TextColumn(width="large")},
    )


def metrics(d: pd.DataFrame, db: CompatibilityDB, version: str):
    s = db.summarize(d, version)
    c = st.columns(6)
    c[0].metric("Products", s["total"])
    c[1].metric("✅ Compatible", s["yes"])
    c[2].metric("🔧 Need PTFs", s["ptf"])
    c[3].metric("⛔ Not compatible", s["no"])
    c[4].metric("⏳ In progress", s["in_progress"])
    c[5].metric("❔ Unverified", s["unverified"])


try:
    db = get_db()
except FileNotFoundError as e:
    st.error(f"Data files not found. {e}")
    st.stop()

with st.spinner("Checking linked vendor pages for PTF, USERMOD, SYSMOD, and prerequisite details..."):
    enriched_rows, maintenance_fetch = enrich_vendor_details(db.df)
db = CompatibilityDB(df=enriched_rows)

use_llm = watsonx_configured()

# ------------------------------------------------------------------ header + sidebar
st.markdown(
    """
    <div class="dashboard-hero">
        <div class="eyebrow">IBM Z · Migration readiness</div>
        <h1>z/OS product compatibility</h1>
        <p>A single view of vendor support, required maintenance, and migration risks.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("### Data coverage")
    _sv = st.session_state.get("zos_ver", ZOS_VERSIONS[0])
    _scol = VER[_sv][0]
    verified = int((~db.df[_scol].isin(["Unverified", "Unknown", ""])).sum())
    st.caption(f"Rows with a z/OS {_sv} statement")
    st.metric("Vendor statements", f"{verified:,} / {len(db.df):,}")
    st.progress(verified / len(db.df) if len(db.df) else 0)
    st.caption(
        "Rows without a sourced vendor statement are marked Unverified; "
        "they are not assumed compatible."
    )
    st.divider()
    st.markdown("**Maintenance key**")
    st.markdown("🟨 PTF / USERMOD / SYSMOD details need review")
    st.markdown("⚪ Unverified compatibility needs vendor confirmation")
    st.divider()
    st.markdown("### Sources")
    st.caption(f"{db.df['vendor'].nunique()} vendors · {len(SOURCES)} reference pages")
    st.metric(
        "Readable vendor pages",
        f"{maintenance_fetch['urls_readable']} / {maintenance_fetch['urls_total']}",
    )
    if maintenance_fetch["urls_unavailable"]:
        st.warning(
            f"{maintenance_fetch['urls_unavailable']} page(s) were blocked, unavailable, "
            "or had no readable text; check the source links."
        )
    st.caption("Unverified means no vendor statement has been sourced for that release.")
    st.caption(
        "Maintenance details are best-effort and cached for 6 hours. "
        "No extracted detail does not mean no maintenance is required."
    )
    with st.expander("watsonx.ai assistant"):
        st.caption("Granite enabled" if use_llm else "Answer generation is not configured")

overview_tab, tab_ver, tab_check, tab_ask, tab_learn, tab_src = st.tabs(
    ["Overview", "Compatibility matrix", "Check my products", "Ask watsonx", "MVS learning lab", "Sources"]
)

# ================================================================== OVERVIEW
with overview_tab:
    version_col, note_col = st.columns([1, 4], vertical_alignment="bottom")
    ver = version_col.selectbox(
        "Target z/OS release",
        ZOS_VERSIONS,
        key="zos_ver",
        help="Choose a release to update the overview, compatibility matrix, and inventory check.",
    )
    scol = VER[ver][0]
    pcol = VER[ver][1]
    overview_stats = db.summarize(db.df, ver)
    note_col.markdown(
        f"<div class='section-kicker'>Portfolio snapshot · {len(db.df):,} product records "
        f"across {db.df['vendor'].nunique()} vendors</div>",
        unsafe_allow_html=True,
    )

    st.markdown(f"#### Portfolio snapshot · z/OS {ver}")
    c = st.columns(6)
    c[0].metric("Products", f"{overview_stats['total']:,}")
    c[1].metric("Compatible", f"{overview_stats['yes']:,}")
    c[2].metric("Need PTFs", f"{overview_stats['ptf']:,}")
    c[3].metric("Not compatible", f"{overview_stats['no']:,}")
    c[4].metric("In progress", f"{overview_stats['in_progress']:,}")
    c[5].metric("Unverified", f"{overview_stats['unverified']:,}")
    st.caption(
        "Results use the explicit vendor statements in the compatibility matrix. "
        "Compatibility is not predicted; unverified records need vendor confirmation."
    )

    st.markdown("#### All products")
    st.caption(f"{len(db.df):,} product/version records across {db.df['vendor'].nunique()} vendors.")
    show_table(db.df, ver, height=560)

    ptf_worklist = db.df[
        (db.df[scol] == "Yes") & db.df[pcol].str.startswith("Yes")
    ].sort_values(["vendor", "product", "vrm"])
    with st.expander(
        f"🟨 PTF / usermod worklist for z/OS {ver}: "
        f"{len(ptf_worklist):,} product/version record(s)"
    ):
        if ptf_worklist.empty:
            st.success(f"No compatible products are marked as requiring PTFs for z/OS {ver}.")
        else:
            st.caption(
                "Review the vendor-listed maintenance before migration. Extracted details are best-effort; "
                "confirm the required fixes on the linked vendor page."
            )
            show_table(ptf_worklist, ver, height=420)

# ================================================================== TAB 1 : z/OS version
with tab_ver:
    st.markdown("#### Compatibility matrix")
    st.caption(f"Filter the vendor statements for the selected target release: z/OS {ver}.")
    st.caption(
        "IBM maintenance details come from public SPCR/APAR pages; no IBM API key is needed for lookup. "
        "PTF IDs may also be SMP/E SYSMOD IDs. USERMODs are shown only when IBM explicitly publishes them."
    )
    vendors = st.multiselect("Vendors", VENDORS, default=VENDORS, key="ver_vendors")

    f1, f2, f3 = st.columns([2, 2, 3])
    status = f1.multiselect(f"R{ver} compatible", ALL_STATUS, default=ALL_STATUS, key="ver_status")
    ptf = f2.selectbox("PTF", ["Any", "PTF required", "No PTF needed"], key="ver_ptf")
    q = f3.text_input("Search", key="ver_q", placeholder="product, VRM, comment…")
    breaking = False
    prev_ver = VER[ver][3]
    if prev_ver:
        breaking = st.checkbox(f"Only products that work on z/OS {prev_ver} but not on {ver}", key=f"ver_break_{ver}")

    view = db.filter(vendors, status, None if ptf == "Any" else ptf, breaking, q, version=ver)
    st.markdown(f"#### Products for z/OS {ver}")
    metrics(view, db, ver)
    show_table(view, ver, height=520)

    book = {f"zOS {ver} - All Vendors": db.df}
    for v in VENDORS:
        d = db.df[db.df["vendor"] == v]
        if not d.empty:
            book[v] = d
    c1, c2 = st.columns(2)
    c1.download_button(f"⬇️ Download z/OS {ver} workbook (all vendors)", build_excel(book, version=ver),
                       file_name=f"zOS_{ver}_compatibility.xlsx", mime=XLSX, type="primary", key="ver_dl_all")
    c2.download_button("⬇️ Download current filtered view", build_excel({"Filtered": view}, version=ver),
                       file_name=f"zOS_{ver}_filtered.xlsx", mime=XLSX, key="ver_dl_view")
    st.caption("Excel links are real hyperlinks: click the Comments / Links cell (or the Sources sheet) to open the vendor page. "
               "If Excel opens the file in Protected View, click Enable Editing first.")

    with st.expander("➕ Add or correct data (upload a CSV with the same columns as data/*.csv)"):
        st.write("Rows are matched on Vendor + Product + VRM and override the built-in data for this session. "
                 "To make them permanent, edit the CSV files in `data/` and redeploy.")
        upd = st.file_uploader("CSV", type=["csv"], key="ver_upd")
        if upd is not None and st.button("Apply uploaded data", key="ver_apply"):
            try:
                st.session_state["extra"] = pd.read_csv(upd, dtype=str, keep_default_na=False)
                st.rerun()
            except Exception as e:
                st.error(f"Could not read the file: {e}")

# ================================================================== TAB 2 : check my products
with tab_check:
    cver = st.session_state.get("zos_ver", ZOS_VERSIONS[0])
    st.markdown(f"#### Check your existing products against z/OS {cver}")
    st.caption(
        "The z/OS version follows your choice on the Overview tab. "
        "Compatibility is looked up from the vendor matrix, not predicted by training a model."
    )
    inventory_template = pd.DataFrame(
        {"Vendor": VENDORS, "Product": ["" for _ in VENDORS], "VRM": ["" for _ in VENDORS]}
    )
    st.download_button(
        "Download inventory template (fill Product and VRM)",
        inventory_template.to_csv(index=False).encode("utf-8-sig"),
        file_name="zOS_product_inventory_template.csv",
        mime="text/csv",
        key="chk_template_dl",
    )
    vendors2 = st.multiselect("Vendors", VENDORS, default=VENDORS, key="chk_vendors")
    pool = db.df[db.df["vendor"].isin(vendors2)]
    picked = st.multiselect(f"Products and existing versions ({len(pool)} available)", pool["label"].tolist(),
                            placeholder="Search e.g. 'Endevor', 'CICS', 'MFX' ...  (Vendor | Product | VRM)",
                            key="chk_picked")
    up = st.file_uploader("…or upload your product inventory (CSV/XLSX with columns Vendor, Product, VRM)",
                          type=["csv", "xlsx"], key="chk_inv")

    chosen = db.df[db.df["label"].isin(picked)]
    misses = []
    inventory_assessment = None
    if up is not None:
        try:
            inv = (pd.read_csv(up, dtype=str, keep_default_na=False) if up.name.lower().endswith(".csv")
                   else pd.read_excel(up, dtype=str, keep_default_na=False))
            checked_at = datetime.now().astimezone().isoformat(timespec="seconds")
            inventory_assessment = db.assess_inventory(inv, cver, checked_at)
            matched, misses = db.match_inventory(inv)
            chosen = pd.concat([chosen, matched], ignore_index=True).drop_duplicates(subset=["vendor", "product", "vrm"])
        except Exception as e:
            st.error(f"Could not read the inventory: {e}")

    if inventory_assessment is not None:
        st.markdown(f"#### Inventory assessment · z/OS {cver}")
        st.caption(
            "The assessment timestamp uses this app server's local timezone. "
            "Unverified or unmatched products are not treated as compatible."
        )
        st.dataframe(inventory_assessment, width="stretch", hide_index=True)
        csv_result = inventory_assessment.to_csv(index=False).encode("utf-8-sig")
        dl1, dl2 = st.columns(2)
        dl1.download_button(
            "Download timestamped assessment (Excel)",
            build_assessment_excel(inventory_assessment),
            file_name=f"zOS_{cver}_inventory_assessment.xlsx",
            mime=XLSX,
            type="primary",
            key="chk_assessment_xlsx",
        )
        dl2.download_button(
            "Download timestamped assessment (CSV)",
            csv_result,
            file_name=f"zOS_{cver}_inventory_assessment.csv",
            mime="text/csv",
            key="chk_assessment_csv",
        )

    if chosen.empty:
        st.info("Select at least one product above (or upload an inventory) to see the answer.")
    else:
        st.markdown("#### Result")
        metrics(chosen, db, cver)
        s = db.summarize(chosen, cver)
        scol = VER[cver][0]
        blockers = chosen[chosen[scol] == "No"]
        if not blockers.empty:
            st.error(f"{len(blockers)} selected product(s) are NOT compatible with z/OS {cver} and need an upgrade first: "
                     + "; ".join(f"{r.product} {r.vrm}" for r in blockers.itertuples()))
        if s["ptf"]:
            st.warning(f"{s['ptf']} compatible product(s) need PTFs applied.")
        if s["unverified"]:
            st.info(f"{s['unverified']} product(s) have no vendor statement in the data yet. Open their link and confirm.")
        if s["no"] == 0 and s["unverified"] == 0 and s["in_progress"] == 0:
            st.success(f"All selected products are compatible with z/OS {cver}." +
                       (" Apply the listed PTFs first." if s["ptf"] else " No PTFs required."))
        show_table(chosen, cver)
        for m in misses:
            st.caption("⚠️ " + m)
        xl = build_excel(
            {f"My Products z{cver}": chosen, "Full Matrix": db.df},
            notes=[f"Generated {datetime.now():%Y-%m-%d %H:%M}. First sheet = your selection; 'Full Matrix' = everything."],
            version=cver,
        )
        st.download_button("⬇️ Download Excel (my products + full matrix)", xl,
                           file_name=f"zOS_{cver}_my_products.xlsx", mime=XLSX, type="primary", key="chk_dl")

# ================================================================== TAB 3 : ask
with tab_ask:
    st.markdown("Ask in plain English, e.g. *Which Broadcom products are not supported on z/OS 3.2?* "
                "or *Does Endevor 18.1 need PTFs?*")
    question = st.text_input("Question", key="ask_q")
    k = st.slider("Rows to retrieve", 5, 30, 12, key="ask_k")
    ask_ver = st.session_state.get("zos_ver", ZOS_VERSIONS[0])
    if question:
        rows = RagIndex.build(db.df).retrieve(question, k)
        if rows.empty:
            st.warning("No matching rows. Try a product name or vendor.")
        else:
            answered = False
            if use_llm:
                with st.spinner("Preparing answer…"):
                    try:
                        st.markdown(generate_answer(question, rows))
                        answered = True
                    except Exception:
                        st.info("The written answer is unavailable right now. The matching rows are shown below.")
            else:
                st.info("Showing the matching rows.")
            with st.expander("Matching rows", expanded=not answered):
                show_table(rows, ask_ver, height=320)

# ================================================================== TAB 4 : MVS learning lab
with tab_learn:
    render_learning_lab()

# ================================================================== TAB 5 : sources
with tab_src:
    st.markdown("Official places to confirm compatibility, PTFs, EOS dates and latest versions.")
    for v in VENDORS:
        items = [s for s in SOURCES if s[0] == v]
        if not items:
            continue
        st.markdown(f"**{v}**")
        for _, name, url, what in items:
            a, b = st.columns([6, 1])
            a.markdown(f"{name}  \n<small>{what}</small>", unsafe_allow_html=True)
            b.link_button("Open ↗", url)
        st.divider()

st.caption("Always confirm on the live vendor page before production changes.")
