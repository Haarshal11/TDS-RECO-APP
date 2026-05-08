"""
TDS Reconciliation Pro — Streamlit Application
A CFO-grade TDS reconciliation tool with interactive dashboards.
Supports 24Q / 26Q / 27Q forms.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from reco_engine import (
    load_file, prepare_books, prepare_returns, prepare_challans,
    run_reconciliation, run_challan_reco, generate_summary,
    check_data_quality, detect_duplicates, export_to_excel
)

# ═══════════════════════════════════════════════════════════════
# PAGE CONFIG
# ═══════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="TDS Reconciliation Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ═══════════════════════════════════════════════════════════════
# CUSTOM CSS
# ═══════════════════════════════════════════════════════════════

st.markdown("""
<style>
    /* ─── Global ─── */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    .stApp {
        font-family: 'Inter', sans-serif;
    }

    /* ─── Sidebar ─── */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f1419 0%, #1a1f2e 100%);
    }

    [data-testid="stSidebar"] .stMarkdown h1 {
        background: linear-gradient(135deg, #6C63FF, #4ECDC4);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 700;
        font-size: 1.5rem;
    }

    /* ─── KPI Cards ─── */
    .kpi-card {
        background: linear-gradient(135deg, #1a1f2e 0%, #252b3b 100%);
        border: 1px solid rgba(108, 99, 255, 0.2);
        border-radius: 16px;
        padding: 24px;
        text-align: center;
        transition: all 0.3s ease;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
    }

    .kpi-card:hover {
        border-color: rgba(108, 99, 255, 0.5);
        transform: translateY(-2px);
        box-shadow: 0 8px 30px rgba(108, 99, 255, 0.15);
    }

    .kpi-value {
        font-size: 2rem;
        font-weight: 700;
        margin: 8px 0;
        line-height: 1.2;
    }

    .kpi-label {
        font-size: 0.8rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #8b8fa3;
        font-weight: 500;
    }

    .kpi-green { color: #4ECDC4; }
    .kpi-blue { color: #6C63FF; }
    .kpi-orange { color: #FF6B6B; }
    .kpi-yellow { color: #FFE66D; }
    .kpi-purple { color: #A78BFA; }

    /* ─── Section Headers ─── */
    .section-header {
        font-size: 1.3rem;
        font-weight: 600;
        color: #FAFAFA;
        margin: 2rem 0 1rem 0;
        padding-bottom: 8px;
        border-bottom: 2px solid rgba(108, 99, 255, 0.3);
    }

    /* ─── Status Badges ─── */
    .badge-match {
        background: rgba(78, 205, 196, 0.15);
        color: #4ECDC4;
        padding: 4px 12px;
        border-radius: 12px;
        font-size: 0.75rem;
        font-weight: 600;
    }

    .badge-mismatch {
        background: rgba(255, 107, 107, 0.15);
        color: #FF6B6B;
        padding: 4px 12px;
        border-radius: 12px;
        font-size: 0.75rem;
        font-weight: 600;
    }

    /* ─── Upload Zone ─── */
    .upload-zone {
        background: linear-gradient(135deg, #1a1f2e 0%, #252b3b 100%);
        border: 2px dashed rgba(108, 99, 255, 0.4);
        border-radius: 16px;
        padding: 32px;
        text-align: center;
        margin: 12px 0;
        transition: all 0.3s ease;
    }

    .upload-zone:hover {
        border-color: rgba(108, 99, 255, 0.8);
        background: linear-gradient(135deg, #1e2433 0%, #2a3045 100%);
    }

    /* ─── Tabs ─── */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 12px;
        padding: 8px 20px;
        font-weight: 500;
    }

    /* ─── Data Tables ─── */
    .stDataFrame {
        border-radius: 12px;
        overflow: hidden;
    }

    /* ─── Divider ─── */
    hr {
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(108, 99, 255, 0.3), transparent);
        margin: 2rem 0;
    }

    /* ─── Hero Banner ─── */
    .hero-banner {
        background: linear-gradient(135deg, #1a1040 0%, #0f1419 50%, #0a1628 100%);
        border: 1px solid rgba(108, 99, 255, 0.2);
        border-radius: 20px;
        padding: 48px;
        text-align: center;
        margin-bottom: 2rem;
        position: relative;
        overflow: hidden;
    }

    .hero-banner::before {
        content: '';
        position: absolute;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background: radial-gradient(circle at 30% 50%, rgba(108, 99, 255, 0.08), transparent 50%),
                    radial-gradient(circle at 70% 50%, rgba(78, 205, 196, 0.05), transparent 50%);
    }

    .hero-title {
        font-size: 2.2rem;
        font-weight: 700;
        background: linear-gradient(135deg, #6C63FF, #4ECDC4, #A78BFA);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 8px;
        position: relative;
    }

    .hero-subtitle {
        font-size: 1rem;
        color: #8b8fa3;
        position: relative;
    }

    /* ─── Footer ─── */
    .footer {
        text-align: center;
        color: #555;
        font-size: 0.75rem;
        padding: 2rem 0 1rem 0;
    }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# HELPER: FORMAT CURRENCY
# ═══════════════════════════════════════════════════════════════

def fmt_inr(val):
    """Format number as Indian Rupee string."""
    if pd.isna(val) or val == 0:
        return "₹0"
    sign = "" if val >= 0 else "-"
    val = abs(val)
    if val >= 10000000:
        return f"{sign}₹{val/10000000:,.2f} Cr"
    elif val >= 100000:
        return f"{sign}₹{val/100000:,.2f} L"
    else:
        return f"{sign}₹{val:,.0f}"


def render_kpi(label, value, color_class="kpi-blue", prefix="", suffix=""):
    """Render a KPI card."""
    return f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value {color_class}">{prefix}{value}{suffix}</div>
    </div>
    """


# ═══════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("# 📊 TDS Reco Pro")
    st.markdown("---")

    st.markdown("### 📁 Upload Files")

    books_file = st.file_uploader(
        "Books / Deductions",
        type=["xlsx", "xls", "csv"],
        help="Upload your books of accounts TDS deduction data",
        key="books"
    )

    returns_file = st.file_uploader(
        "TDS Returns (24Q/26Q/27Q)",
        type=["xlsx", "xls", "csv"],
        help="Upload your filed TDS return data",
        key="returns"
    )

    challans_file = st.file_uploader(
        "Challans (Optional)",
        type=["xlsx", "xls", "csv"],
        help="Upload challan deposit data for 3-way reconciliation",
        key="challans"
    )

    st.markdown("---")
    st.markdown("### ⚙️ Settings")

    fy_options = [f"FY {y}-{y+1}" for y in range(2020, 2027)]
    fy_selected = st.selectbox("Financial Year", fy_options, index=5)

    form_type = st.selectbox("Form Type", ["26Q (Non-Salary)", "24Q (Salary)", "27Q (NRI)"])

    tolerance = st.number_input(
        "Tolerance (₹)",
        min_value=0.0, max_value=100.0, value=1.0, step=0.5,
        help="Differences within this amount will be treated as MATCH"
    )

    st.markdown("---")

    run_btn = st.button(
        "🚀 Run Reconciliation",
        use_container_width=True,
        type="primary",
        disabled=(books_file is None or returns_file is None)
    )

    st.markdown("---")
    st.markdown(
        '<div class="footer">TDS Reconciliation Pro v2.0<br>'
        'Built for CFO-grade compliance</div>',
        unsafe_allow_html=True
    )


# ═══════════════════════════════════════════════════════════════
# MAIN AREA
# ═══════════════════════════════════════════════════════════════

# Session state init
if "reco_done" not in st.session_state:
    st.session_state.reco_done = False

if run_btn and books_file and returns_file:
    with st.spinner("🔄 Running 3-Way TDS Reconciliation..."):
        try:
            # Load files
            books_raw = load_file(books_file.getvalue(), books_file.name)
            returns_raw = load_file(returns_file.getvalue(), returns_file.name)

            challans_raw = None
            if challans_file:
                challans_raw = load_file(challans_file.getvalue(), challans_file.name)

            # Prepare data
            books_df = prepare_books(books_raw)
            returns_df = prepare_returns(returns_raw)
            challans_df = prepare_challans(challans_raw) if challans_raw is not None else None

            # Run reconciliation
            reco_df = run_reconciliation(books_df, returns_df, tolerance=tolerance)

            # Challan reco
            challan_reco_df = run_challan_reco(books_df, challans_df) if challans_df is not None else None

            # Summary
            summary, qtr_summary, status_breakdown, sec_summary = generate_summary(
                reco_df, books_df, returns_df, challans_df
            )

            # Data quality
            quality_df = check_data_quality(books_df, returns_df)
            duplicate_df = detect_duplicates(books_df, returns_df)

            # Store in session
            st.session_state.reco_done = True
            st.session_state.reco_df = reco_df
            st.session_state.books_df = books_df
            st.session_state.returns_df = returns_df
            st.session_state.challans_df = challans_df
            st.session_state.challan_reco_df = challan_reco_df
            st.session_state.summary = summary
            st.session_state.qtr_summary = qtr_summary
            st.session_state.status_breakdown = status_breakdown
            st.session_state.sec_summary = sec_summary
            st.session_state.quality_df = quality_df
            st.session_state.duplicate_df = duplicate_df
            st.session_state.fy = fy_selected
            st.session_state.form_type = form_type

        except Exception as e:
            st.error(f"❌ Error during reconciliation: {str(e)}")
            st.exception(e)
            st.session_state.reco_done = False


# ═══════════════════════════════════════════════════════════════
# LANDING PAGE (Before reconciliation)
# ═══════════════════════════════════════════════════════════════

if not st.session_state.reco_done:
    st.markdown("""
    <div class="hero-banner">
        <div class="hero-title">TDS Reconciliation Pro</div>
        <div class="hero-subtitle">
            CFO-Grade 3-Way TDS Reconciliation Engine &nbsp;•&nbsp;
            Books ↔ Returns ↔ Challans
        </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="kpi-card">
            <div class="kpi-label">STEP 1</div>
            <div class="kpi-value kpi-blue">📁</div>
            <div class="kpi-label">Upload Files</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="kpi-card">
            <div class="kpi-label">STEP 2</div>
            <div class="kpi-value kpi-green">⚙️</div>
            <div class="kpi-label">Configure Settings</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="kpi-card">
            <div class="kpi-label">STEP 3</div>
            <div class="kpi-value kpi-purple">🚀</div>
            <div class="kpi-label">Run Reconciliation</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    st.markdown("### ✨ Features")
    feat_col1, feat_col2, feat_col3 = st.columns(3)
    with feat_col1:
        st.markdown("""
        - 🔄 **3-Way Reconciliation**
        - 📊 **Executive Dashboard**
        - 🎯 **PAN Validation**
        """)
    with feat_col2:
        st.markdown("""
        - 📋 **Challan Matching**
        - 🔍 **Duplicate Detection**
        - ⚡ **Section Normalization**
        """)
    with feat_col3:
        st.markdown("""
        - 📥 **Excel & CSV Export**
        - 🛡️ **Data Quality Audit**
        - 📈 **Interactive Charts**
        """)

    st.markdown("---")
    st.info("👈 **Upload your Books and Returns files** in the sidebar to get started.")

else:
    # ═══════════════════════════════════════════════════════════
    # RESULTS PAGES (After reconciliation)
    # ═══════════════════════════════════════════════════════════

    summary = st.session_state.summary
    reco_df = st.session_state.reco_df
    books_df = st.session_state.books_df
    returns_df = st.session_state.returns_df
    challans_df = st.session_state.challans_df
    challan_reco_df = st.session_state.challan_reco_df
    qtr_summary = st.session_state.qtr_summary
    status_breakdown = st.session_state.status_breakdown
    sec_summary = st.session_state.sec_summary
    quality_df = st.session_state.quality_df
    duplicate_df = st.session_state.duplicate_df

    # ─── Tabs ───
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Executive Dashboard",
        "🔍 Detailed Reconciliation",
        "💳 Challan Reconciliation",
        "🛡️ Data Quality",
        "📥 Download Reports"
    ])

    # ═══════════════════════════════════════════════════════════
    # TAB 1: EXECUTIVE DASHBOARD
    # ═══════════════════════════════════════════════════════════

    with tab1:
        st.markdown(f"""
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 1rem;">
            <div>
                <span style="font-size: 1.4rem; font-weight: 700; color: #FAFAFA;">
                    Executive Dashboard
                </span>
                <span style="color: #8b8fa3; margin-left: 12px;">
                    {st.session_state.fy} &nbsp;•&nbsp; {st.session_state.form_type}
                </span>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # KPI Row 1
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(render_kpi(
                "Total Books TDS", fmt_inr(summary["total_books_tds"]), "kpi-blue"
            ), unsafe_allow_html=True)
        with k2:
            st.markdown(render_kpi(
                "Total Returns TDS", fmt_inr(summary["total_returns_tds"]), "kpi-green"
            ), unsafe_allow_html=True)
        with k3:
            st.markdown(render_kpi(
                "Net Difference", fmt_inr(summary["net_difference"]),
                "kpi-green" if summary["net_difference"] == 0 else "kpi-orange"
            ), unsafe_allow_html=True)
        with k4:
            st.markdown(render_kpi(
                "Match Rate", f"{summary['match_percentage']:.1f}%",
                "kpi-green" if summary["match_percentage"] >= 90 else "kpi-orange"
            ), unsafe_allow_html=True)

        st.markdown("")

        # KPI Row 2
        k5, k6, k7, k8 = st.columns(4)
        with k5:
            st.markdown(render_kpi(
                "Total Deductees (Books)", str(summary["total_deductees_books"]), "kpi-purple"
            ), unsafe_allow_html=True)
        with k6:
            st.markdown(render_kpi(
                "Total Deductees (Returns)", str(summary["total_deductees_returns"]), "kpi-purple"
            ), unsafe_allow_html=True)
        with k7:
            st.markdown(render_kpi(
                "Books Transactions", str(summary["total_books_txns"]), "kpi-blue"
            ), unsafe_allow_html=True)
        with k8:
            challan_val = fmt_inr(summary["total_challans_deposited"]) if summary["total_challans_deposited"] > 0 else "N/A"
            st.markdown(render_kpi(
                "Challans Deposited", challan_val, "kpi-yellow"
            ), unsafe_allow_html=True)

        st.markdown("---")

        # Charts Row
        chart1, chart2 = st.columns([3, 2])

        with chart1:
            st.markdown('<div class="section-header">Quarter-wise Comparison</div>', unsafe_allow_html=True)
            fig_bar = go.Figure()
            fig_bar.add_trace(go.Bar(
                name="Books TDS",
                x=qtr_summary["Quarter"],
                y=qtr_summary["Books_TDS"],
                marker_color="#6C63FF",
                marker_line_color="#5A52E0",
                marker_line_width=1,
            ))
            fig_bar.add_trace(go.Bar(
                name="Returns TDS",
                x=qtr_summary["Quarter"],
                y=qtr_summary["Returns_TDS"],
                marker_color="#4ECDC4",
                marker_line_color="#3DBDB5",
                marker_line_width=1,
            ))
            fig_bar.update_layout(
                barmode="group",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#FAFAFA", family="Inter"),
                legend=dict(
                    orientation="h", yanchor="bottom", y=1.02,
                    xanchor="right", x=1, bgcolor="rgba(0,0,0,0)"
                ),
                margin=dict(l=20, r=20, t=40, b=20),
                height=380,
                xaxis=dict(gridcolor="rgba(255,255,255,0.05)"),
                yaxis=dict(gridcolor="rgba(255,255,255,0.08)", title="Amount (₹)"),
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        with chart2:
            st.markdown('<div class="section-header">Match Status Distribution</div>', unsafe_allow_html=True)
            # Custom colors for each status
            color_map = {
                "PERFECT_MATCH": "#4ECDC4",
                "SHORT_IN_RETURN": "#FF6B6B",
                "EXCESS_IN_RETURN": "#FFE66D",
                "ONLY_IN_BOOKS": "#FF9F43",
                "ONLY_IN_RETURNS": "#A78BFA",
                "SECTION_MISMATCH": "#54A0FF",
                "PAN_ISSUE": "#EE5A24",
            }
            colors = [color_map.get(s, "#888") for s in status_breakdown["Status"]]
            fig_donut = go.Figure(data=[go.Pie(
                labels=status_breakdown["Status"],
                values=status_breakdown["Count"],
                hole=0.55,
                marker=dict(colors=colors, line=dict(color="#0E1117", width=2)),
                textinfo="label+percent",
                textfont=dict(size=11),
            )])
            fig_donut.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#FAFAFA", family="Inter"),
                margin=dict(l=10, r=10, t=40, b=10),
                height=380,
                showlegend=False,
                annotations=[dict(
                    text=f"{summary['match_percentage']:.0f}%",
                    x=0.5, y=0.5, font_size=28, showarrow=False,
                    font=dict(color="#4ECDC4", family="Inter", weight=700 if hasattr(dict, 'weight') else None)
                )]
            )
            st.plotly_chart(fig_donut, use_container_width=True)

        st.markdown("---")

        # Quarter-wise table
        st.markdown('<div class="section-header">Quarter-wise Summary</div>', unsafe_allow_html=True)
        st.dataframe(
            qtr_summary.style.format({
                "Books_TDS": "₹{:,.2f}",
                "Returns_TDS": "₹{:,.2f}",
                "Difference": "₹{:,.2f}",
            }),
            use_container_width=True,
            hide_index=True
        )

        # Section-wise table
        st.markdown('<div class="section-header">Section-wise Summary</div>', unsafe_allow_html=True)
        st.dataframe(
            sec_summary.style.format({
                "Books_TDS": "₹{:,.2f}",
                "Returns_TDS": "₹{:,.2f}",
                "Difference": "₹{:,.2f}",
            }),
            use_container_width=True,
            hide_index=True
        )

    # ═══════════════════════════════════════════════════════════
    # TAB 2: DETAILED RECONCILIATION
    # ═══════════════════════════════════════════════════════════

    with tab2:
        st.markdown('<div class="section-header">Detailed PAN-Level Reconciliation</div>', unsafe_allow_html=True)

        # Filters
        fil1, fil2, fil3, fil4 = st.columns(4)
        with fil1:
            qtr_filter = st.multiselect(
                "Quarter",
                options=sorted(reco_df["Quarter"].unique()),
                default=sorted(reco_df["Quarter"].unique()),
                key="reco_qtr"
            )
        with fil2:
            status_filter = st.multiselect(
                "Match Status",
                options=sorted(reco_df["Match_Status"].unique()),
                default=sorted(reco_df["Match_Status"].unique()),
                key="reco_status"
            )
        with fil3:
            pan_search = st.text_input("Search PAN", "", key="pan_search")
        with fil4:
            name_search = st.text_input("Search Name", "", key="name_search")

        # Apply filters
        filtered = reco_df[
            (reco_df["Quarter"].isin(qtr_filter)) &
            (reco_df["Match_Status"].isin(status_filter))
        ]
        if pan_search:
            filtered = filtered[filtered["PAN"].str.contains(pan_search.upper(), na=False)]
        if name_search:
            filtered = filtered[filtered["Deductee_Name"].str.contains(name_search, case=False, na=False)]

        # Summary of filtered data
        f1, f2, f3, f4 = st.columns(4)
        with f1:
            st.metric("Filtered Entries", len(filtered))
        with f2:
            st.metric("Books TDS", fmt_inr(filtered["Books_TDS_Total"].sum()))
        with f3:
            st.metric("Returns TDS", fmt_inr(filtered["Returns_TDS_Total"].sum()))
        with f4:
            st.metric("Net Difference", fmt_inr(filtered["Difference"].sum()))

        # Apply color coding
        def highlight_status(val):
            colors = {
                "PERFECT_MATCH": "background-color: rgba(78, 205, 196, 0.15); color: #4ECDC4",
                "SHORT_IN_RETURN": "background-color: rgba(255, 107, 107, 0.15); color: #FF6B6B",
                "EXCESS_IN_RETURN": "background-color: rgba(255, 230, 109, 0.15); color: #FFE66D",
                "ONLY_IN_BOOKS": "background-color: rgba(255, 159, 67, 0.15); color: #FF9F43",
                "ONLY_IN_RETURNS": "background-color: rgba(167, 139, 250, 0.15); color: #A78BFA",
                "SECTION_MISMATCH": "background-color: rgba(84, 160, 255, 0.15); color: #54A0FF",
                "PAN_ISSUE": "background-color: rgba(238, 90, 36, 0.15); color: #EE5A24",
            }
            return colors.get(val, "")

        styled_df = filtered.style.map(
            highlight_status, subset=["Match_Status"]
        ).format({
            "Books_TDS_Total": "₹{:,.2f}",
            "Returns_TDS_Total": "₹{:,.2f}",
            "Difference": "₹{:,.2f}",
            "Books_Base_Total": "₹{:,.2f}",
        })

        st.dataframe(styled_df, use_container_width=True, hide_index=True, height=500)

        # Drill-down
        st.markdown("---")
        st.markdown('<div class="section-header">🔍 Transaction Drill-Down</div>', unsafe_allow_html=True)
        drill_pan = st.selectbox(
            "Select PAN to view transactions",
            options=[""] + sorted(filtered["PAN"].unique().tolist()),
            key="drill_pan"
        )

        if drill_pan:
            d1, d2 = st.columns(2)
            with d1:
                st.markdown("**📖 Books Transactions**")
                book_txns = books_df[books_df["pan"] == drill_pan][
                    ["quarter", "name", "section", "tds", "base_amount"]
                ].sort_values("quarter")
                st.dataframe(book_txns, use_container_width=True, hide_index=True)

            with d2:
                st.markdown("**📑 Returns Transactions**")
                ret_txns = returns_df[returns_df["pan"] == drill_pan][
                    ["quarter", "name", "section", "tds"]
                ].sort_values("quarter")
                st.dataframe(ret_txns, use_container_width=True, hide_index=True)


    # ═══════════════════════════════════════════════════════════
    # TAB 3: CHALLAN RECONCILIATION
    # ═══════════════════════════════════════════════════════════

    with tab3:
        st.markdown('<div class="section-header">Challan Reconciliation</div>', unsafe_allow_html=True)

        if challan_reco_df is not None and not challan_reco_df.empty:
            # KPIs
            ck1, ck2, ck3, ck4 = st.columns(4)
            total_challan_amt = challan_reco_df["Challan_Total"].sum()
            total_books_amt = challan_reco_df["Books_TDS_Total"].sum()
            short_deposit = challan_reco_df[challan_reco_df["Status"] == "SHORT_DEPOSIT"]
            excess_deposit = challan_reco_df[challan_reco_df["Status"] == "EXCESS_DEPOSIT"]

            with ck1:
                st.markdown(render_kpi(
                    "Total Deposited", fmt_inr(total_challan_amt), "kpi-green"
                ), unsafe_allow_html=True)
            with ck2:
                st.markdown(render_kpi(
                    "Total Deducted", fmt_inr(total_books_amt), "kpi-blue"
                ), unsafe_allow_html=True)
            with ck3:
                st.markdown(render_kpi(
                    "Short Deposits", str(len(short_deposit)), "kpi-orange"
                ), unsafe_allow_html=True)
            with ck4:
                st.markdown(render_kpi(
                    "Excess Deposits", str(len(excess_deposit)), "kpi-yellow"
                ), unsafe_allow_html=True)

            st.markdown("")

            # Challan chart
            fig_chal = go.Figure()
            fig_chal.add_trace(go.Bar(
                name="Books TDS", x=challan_reco_df["Section"],
                y=challan_reco_df["Books_TDS_Total"],
                marker_color="#6C63FF"
            ))
            fig_chal.add_trace(go.Bar(
                name="Challan Deposited", x=challan_reco_df["Section"],
                y=challan_reco_df["Challan_Total"],
                marker_color="#4ECDC4"
            ))
            fig_chal.update_layout(
                barmode="group",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#FAFAFA", family="Inter"),
                margin=dict(l=20, r=20, t=30, b=20),
                height=350,
                yaxis=dict(gridcolor="rgba(255,255,255,0.08)", title="Amount (₹)"),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            )
            st.plotly_chart(fig_chal, use_container_width=True)

            # Challan table
            st.dataframe(
                challan_reco_df.style.format({
                    "Books_TDS_Total": "₹{:,.2f}",
                    "Challan_Total": "₹{:,.2f}",
                    "Difference": "₹{:,.2f}",
                }),
                use_container_width=True, hide_index=True
            )

            if not short_deposit.empty:
                st.warning(f"⚠️ **{len(short_deposit)} section(s) have SHORT DEPOSITS** — total short: ₹{short_deposit['Difference'].sum():,.2f}")
        else:
            st.info("💡 Upload a **Challans file** in the sidebar to enable 3-way challan reconciliation.")


    # ═══════════════════════════════════════════════════════════
    # TAB 4: DATA QUALITY
    # ═══════════════════════════════════════════════════════════

    with tab4:
        st.markdown('<div class="section-header">Data Quality Report</div>', unsafe_allow_html=True)

        dq1, dq2 = st.columns(2)

        with dq1:
            if quality_df is not None and not quality_df.empty:
                st.markdown(f"### ⚠️ Quality Issues ({len(quality_df)})")
                st.dataframe(quality_df, use_container_width=True, hide_index=True, height=400)
            else:
                st.success("✅ No data quality issues detected!")

        with dq2:
            if duplicate_df is not None and not duplicate_df.empty:
                st.markdown(f"### 🔁 Potential Duplicates ({len(duplicate_df)})")
                st.dataframe(duplicate_df, use_container_width=True, hide_index=True, height=400)
            else:
                st.success("✅ No potential duplicates detected!")


    # ═══════════════════════════════════════════════════════════
    # TAB 5: DOWNLOAD REPORTS
    # ═══════════════════════════════════════════════════════════

    with tab5:
        st.markdown('<div class="section-header">Download Reports</div>', unsafe_allow_html=True)

        dl1, dl2, dl3 = st.columns(3)

        with dl1:
            st.markdown("""
            <div class="kpi-card">
                <div class="kpi-value kpi-blue">📊</div>
                <div class="kpi-label">Full Excel Report</div>
            </div>
            """, unsafe_allow_html=True)
            st.markdown("")

            # Generate Excel
            try:
                excel_bytes = export_to_excel(
                    reco_df, books_df, returns_df, challan_reco_df,
                    quality_df, duplicate_df, qtr_summary, sec_summary, summary
                )
                st.download_button(
                    label="📥 Download Excel Report",
                    data=excel_bytes,
                    file_name=f"TDS_Reconciliation_{st.session_state.fy.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
            except Exception as e:
                st.error(f"Excel generation error: {e}")

        with dl2:
            st.markdown("""
            <div class="kpi-card">
                <div class="kpi-value kpi-green">🔍</div>
                <div class="kpi-label">Reco Detail CSV</div>
            </div>
            """, unsafe_allow_html=True)
            st.markdown("")

            csv_data = reco_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Download Reco CSV",
                data=csv_data,
                file_name=f"TDS_Reco_Detail_{st.session_state.fy.replace(' ', '_')}.csv",
                mime="text/csv",
                use_container_width=True
            )

        with dl3:
            st.markdown("""
            <div class="kpi-card">
                <div class="kpi-value kpi-purple">📋</div>
                <div class="kpi-label">Mismatches Only</div>
            </div>
            """, unsafe_allow_html=True)
            st.markdown("")

            mismatches = reco_df[reco_df["Match_Status"] != "PERFECT_MATCH"]
            mismatch_csv = mismatches.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Download Mismatches",
                data=mismatch_csv,
                file_name=f"TDS_Mismatches_{st.session_state.fy.replace(' ', '_')}.csv",
                mime="text/csv",
                use_container_width=True
            )

        st.markdown("---")
        st.markdown(f"""
        <div style="text-align: center; color: #555; font-size: 0.85rem; padding: 1rem;">
            📊 TDS Reconciliation Pro v2.0 &nbsp;•&nbsp; {st.session_state.fy} &nbsp;•&nbsp;
            Generated with {summary['total_books_txns']} books transactions and
            {summary['total_returns_txns']} return entries
        </div>
        """, unsafe_allow_html=True)
