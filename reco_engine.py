"""
TDS Reconciliation Engine v2.0
3-Way Reconciliation: Books ↔ Returns ↔ Challans
Supports 24Q / 26Q / 27Q forms
"""

import pandas as pd
import numpy as np
import re
import io
from datetime import datetime, date

# ═══════════════════════════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════════════════════════

PAN_REGEX = re.compile(r'^[A-Z]{5}[0-9]{4}[A-Z]$')

SECTION_ALIASES = {
    "194IA": "194IA", "194-IA": "194IA", "194 IA": "194IA",
    "194IB": "194IB", "194-IB": "194IB", "194 IB": "194IB",
    "194IC": "194IC", "194-IC": "194IC", "194 IC": "194IC",
    "194I(A)": "194I(A)", "194IA(A)": "194I(A)", "194I (A)": "194I(A)",
    "194I(B)": "194I(B)", "194IA(B)": "194I(B)", "194I (B)": "194I(B)",
    "194J(A)": "194J(A)", "194J (A)": "194J(A)",
    "194J(B)": "194J(B)", "194J (B)": "194J(B)",
}

INTEREST_RATE_LATE_DEDUCTION = 0.01   # 1% per month u/s 201(1A)(i)
INTEREST_RATE_LATE_DEPOSIT   = 0.015  # 1.5% per month u/s 201(1A)(ii)


# ═══════════════════════════════════════════════════════════════
# HEADER AUTO-DETECT
# ═══════════════════════════════════════════════════════════════

def detect_header_row(file_bytes, filename, max_rows=25):
    """Detect header row in an Excel file by scanning for keywords."""
    if filename.endswith(('.xlsx', '.xls')):
        preview = pd.read_excel(io.BytesIO(file_bytes), header=None, nrows=max_rows)
    else:
        preview = pd.read_csv(io.BytesIO(file_bytes), header=None, nrows=max_rows)

    keywords = ["pan", "tds", "section", "amount", "quarter", "name",
                 "challan", "deductee", "date", "voucher"]
    for i in range(len(preview)):
        row = preview.iloc[i].astype(str).str.lower()
        if any(k in cell for cell in row for k in keywords):
            return i
    return 0  # default to first row


# ═══════════════════════════════════════════════════════════════
# COLUMN UTILITIES
# ═══════════════════════════════════════════════════════════════

def clean_cols(df):
    """Normalize column names to lowercase with underscores."""
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r'\s+', '_', regex=True)
        .str.replace(r'[^\w]', '_', regex=True)
    )
    return df


def find_col(df, options, required=True):
    """Find a column from a list of possible names."""
    for o in options:
        if o in df.columns:
            return o
    # Fuzzy: check if any column contains any option as substring
    for o in options:
        for c in df.columns:
            if o in c:
                return c
    if required:
        raise ValueError(f"Required column missing. Tried {options}. Found {list(df.columns)}")
    return None


def find_col_safe(df, options):
    """Find column without raising exception."""
    return find_col(df, options, required=False)


# ═══════════════════════════════════════════════════════════════
# DATA NORMALIZATION
# ═══════════════════════════════════════════════════════════════

def normalize_section(s):
    """Normalize a TDS section string for consistent comparison."""
    if not s or pd.isna(s):
        return ""
    s = str(s).strip().upper()
    s = re.sub(r'\s+', '', s)       # remove all spaces
    s = re.sub(r'-', '', s)         # remove hyphens
    # Normalize parentheses: "194JA" -> stay, "194J(A)" -> stay
    if s in SECTION_ALIASES:
        return SECTION_ALIASES[s]
    # Generic normalization: add parens around trailing single letter if missing
    m = re.match(r'^(194[A-Z]*)([A-Z])$', s)
    if m and len(m.group(2)) == 1 and m.group(1) not in ('194IA', '194IB', '194IC'):
        return f"{m.group(1)}({m.group(2)})"
    return s


def normalize_section_set(section_str):
    """Normalize a comma-separated section string to a set."""
    if not section_str or pd.isna(section_str):
        return set()
    return {normalize_section(s) for s in str(section_str).split(",")}


def validate_pan(pan):
    """Validate PAN format and return status."""
    if not pan or pd.isna(pan) or str(pan).strip() == "":
        return "MISSING"
    pan = str(pan).strip().upper()
    if PAN_REGEX.match(pan):
        return "VALID"
    return "INVALID"


def quarter_from_date(series):
    """Derive fiscal quarter from date series."""
    d = pd.to_datetime(series, errors="coerce", dayfirst=True)
    m = d.dt.month
    return np.select(
        [m.isin([4, 5, 6]), m.isin([7, 8, 9]), m.isin([10, 11, 12]), m.isin([1, 2, 3])],
        ["Q1", "Q2", "Q3", "Q4"],
        default=""
    )


def best_name(series):
    """Pick the longest non-null name from a series."""
    series = series.dropna().astype(str).str.strip()
    return max(series, key=len) if not series.empty else ""


def join_unique(series):
    """Join unique non-null values with comma."""
    return ", ".join(sorted(series.dropna().astype(str).unique()))


# ═══════════════════════════════════════════════════════════════
# FILE LOADING
# ═══════════════════════════════════════════════════════════════

def load_file(file_bytes, filename):
    """Load an uploaded file (Excel or CSV) with header auto-detection."""
    header_row = detect_header_row(file_bytes, filename)
    if filename.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(io.BytesIO(file_bytes), header=header_row)
    else:
        df = pd.read_csv(io.BytesIO(file_bytes), header=header_row)
    return clean_cols(df)


# ═══════════════════════════════════════════════════════════════
# DATA PREPARATION
# ═══════════════════════════════════════════════════════════════

def prepare_books(df):
    """Standardize books (deductions) dataframe."""
    b_pan = find_col(df, ["pan"])
    b_name = find_col(df, ["deductee_name", "name", "name_"])
    b_sec = find_col(df, ["tds_section", "section"])
    b_amt = find_col(df, ["tds_amount", "tds", "tax_amount"])
    b_base = find_col_safe(df, ["base_amount", "gross_amount", "amount_paid", "amount"])
    b_rate = find_col_safe(df, ["rate", "tds_rate"])

    # Quarter handling
    if "quarter" in df.columns or "quarters" in df.columns:
        qcol = find_col(df, ["quarter", "quarters"])
        df["quarter"] = df[qcol].astype(str).str.upper().str.strip()
    else:
        b_date = find_col(df, ["date", "voucher_date", "date_of_deduction"])
        df["quarter"] = quarter_from_date(df[b_date])

    df["pan"] = df[b_pan].fillna("").astype(str).str.upper().str.strip()
    df["name"] = df[b_name].astype(str).str.strip()
    df["section"] = df[b_sec].astype(str).str.strip().apply(normalize_section)
    df["tds"] = pd.to_numeric(df[b_amt], errors="coerce").fillna(0)

    if b_base:
        df["base_amount"] = pd.to_numeric(df[b_base], errors="coerce").fillna(0)
    else:
        df["base_amount"] = 0

    if b_rate:
        df["rate"] = pd.to_numeric(
            df[b_rate].astype(str).str.replace('%', ''), errors="coerce"
        ).fillna(0)
    else:
        df["rate"] = 0

    df["pan_status"] = df["pan"].apply(validate_pan)
    df["source"] = "BOOKS"
    return df


def prepare_returns(df):
    """Standardize returns dataframe."""
    r_pan = find_col(df, ["pan"])
    r_name = find_col(df, ["deductee_name", "name"])
    r_sec = find_col(df, ["tds_section", "section"])
    r_amt = find_col(df, ["tds_amount", "tds", "tax_amount"])
    r_qtr = find_col(df, ["quarter"])

    df["pan"] = df[r_pan].fillna("").astype(str).str.upper().str.strip()
    df["name"] = df[r_name].astype(str).str.strip()
    df["section"] = df[r_sec].astype(str).str.strip().apply(normalize_section)
    df["tds"] = pd.to_numeric(df[r_amt], errors="coerce").fillna(0)
    df["quarter"] = df[r_qtr].astype(str).str.upper().str.strip()

    df["pan_status"] = df["pan"].apply(validate_pan)
    df["source"] = "RETURNS"
    return df


def prepare_challans(df):
    """Standardize challans dataframe."""
    c_date = find_col(df, ["challan_date", "date", "deposit_date"])
    c_no = find_col_safe(df, ["challan_no", "challan_number", "bsr_code"])
    c_sec = find_col(df, ["tds_section", "section", "tds_section_"])
    c_amt = find_col(df, ["challan_amount", "amount", "challan_amount_"])

    df["challan_date"] = pd.to_datetime(df[c_date], errors="coerce", dayfirst=True)
    df["section"] = df[c_sec].astype(str).str.strip().apply(normalize_section)
    df["challan_amount"] = pd.to_numeric(df[c_amt], errors="coerce").fillna(0)

    if c_no:
        df["challan_no"] = df[c_no].astype(str).str.strip()
    else:
        df["challan_no"] = ""

    # Derive quarter from challan date
    m = df["challan_date"].dt.month
    df["quarter"] = np.select(
        [m.isin([4, 5, 6]), m.isin([7, 8, 9]), m.isin([10, 11, 12]), m.isin([1, 2, 3])],
        ["Q1", "Q2", "Q3", "Q4"],
        default=""
    )

    df["source"] = "CHALLANS"
    return df


# ═══════════════════════════════════════════════════════════════
# DATA QUALITY CHECKS
# ═══════════════════════════════════════════════════════════════

def check_data_quality(books_df, returns_df):
    """Run data quality checks and return a list of issues."""
    issues = []

    # Invalid PANs
    invalid_pans_books = books_df[books_df["pan_status"] == "INVALID"][["pan", "name", "tds"]].drop_duplicates(subset=["pan"])
    if not invalid_pans_books.empty:
        for _, r in invalid_pans_books.iterrows():
            issues.append({
                "Source": "Books",
                "Type": "Invalid PAN",
                "PAN": r["pan"],
                "Name": r["name"],
                "Details": f"PAN format invalid — does not match AAAAA9999A pattern"
            })

    invalid_pans_ret = returns_df[returns_df["pan_status"] == "INVALID"][["pan", "name", "tds"]].drop_duplicates(subset=["pan"])
    if not invalid_pans_ret.empty:
        for _, r in invalid_pans_ret.iterrows():
            issues.append({
                "Source": "Returns",
                "Type": "Invalid PAN",
                "PAN": r["pan"],
                "Name": r["name"],
                "Details": f"PAN format invalid"
            })

    # Missing PANs
    missing_books = books_df[books_df["pan_status"] == "MISSING"]
    if not missing_books.empty:
        total_tds = missing_books["tds"].sum()
        issues.append({
            "Source": "Books",
            "Type": "Missing PAN",
            "PAN": "",
            "Name": f"{len(missing_books)} entries",
            "Details": f"Total TDS with missing PAN: ₹{total_tds:,.2f}"
        })

    # Negative amounts
    neg_books = books_df[books_df["tds"] < 0]
    if not neg_books.empty:
        for _, r in neg_books.iterrows():
            issues.append({
                "Source": "Books",
                "Type": "Negative TDS",
                "PAN": r["pan"],
                "Name": r["name"],
                "Details": f"TDS amount is negative: ₹{r['tds']:,.2f}"
            })

    neg_ret = returns_df[returns_df["tds"] < 0]
    if not neg_ret.empty:
        for _, r in neg_ret.iterrows():
            issues.append({
                "Source": "Returns",
                "Type": "Negative TDS",
                "PAN": r["pan"],
                "Name": r["name"],
                "Details": f"TDS amount is negative: ₹{r['tds']:,.2f}"
            })

    # Blank sections
    blank_sec_books = books_df[books_df["section"].isin(["", "NAN", "NONE"])]
    if not blank_sec_books.empty:
        issues.append({
            "Source": "Books",
            "Type": "Missing Section",
            "PAN": "",
            "Name": f"{len(blank_sec_books)} entries",
            "Details": f"TDS section is blank for {len(blank_sec_books)} deduction entries"
        })

    return pd.DataFrame(issues) if issues else pd.DataFrame(
        columns=["Source", "Type", "PAN", "Name", "Details"]
    )


def detect_duplicates(books_df, returns_df):
    """Detect potential duplicates within books and returns."""
    dupes = []

    # Books duplicates: same PAN + same TDS amount + same date/quarter
    book_cols = ["pan", "tds", "quarter"]
    if "date" in books_df.columns:
        book_cols.append("date")
    books_dups = books_df[books_df.duplicated(subset=book_cols, keep=False)]
    if not books_dups.empty:
        for pan, grp in books_dups.groupby("pan"):
            if pan:
                dupes.append({
                    "Source": "Books",
                    "PAN": pan,
                    "Name": grp["name"].iloc[0],
                    "Count": len(grp),
                    "Total_TDS": grp["tds"].sum(),
                    "Details": f"{len(grp)} entries with same PAN + TDS amount in same quarter"
                })

    # Returns duplicates
    ret_cols = ["pan", "tds", "quarter"]
    ret_dups = returns_df[returns_df.duplicated(subset=ret_cols, keep=False)]
    if not ret_dups.empty:
        for pan, grp in ret_dups.groupby("pan"):
            if pan:
                dupes.append({
                    "Source": "Returns",
                    "PAN": pan,
                    "Name": grp["name"].iloc[0],
                    "Count": len(grp),
                    "Total_TDS": grp["tds"].sum(),
                    "Details": f"{len(grp)} entries with same PAN + TDS amount in same quarter"
                })

    return pd.DataFrame(dupes) if dupes else pd.DataFrame(
        columns=["Source", "PAN", "Name", "Count", "Total_TDS", "Details"]
    )


# ═══════════════════════════════════════════════════════════════
# CORE RECONCILIATION
# ═══════════════════════════════════════════════════════════════

def run_reconciliation(books_df, returns_df, tolerance=0.0):
    """
    PAN + Quarter level reconciliation between Books and Returns.
    Returns detailed reconciliation dataframe.
    """
    # Build name master
    name_master = (
        pd.concat([books_df[["pan", "name"]], returns_df[["pan", "name"]]])
        .groupby("pan")["name"]
        .apply(best_name)
        .to_dict()
    )

    # Group books
    books_grp = books_df.groupby(["pan", "quarter"]).agg(
        Books_TDS_Total=("tds", "sum"),
        Books_Base_Total=("base_amount", "sum"),
        Books_Sections=("section", join_unique),
        Books_Txn_Count=("tds", "count"),
    ).reset_index()

    # Group returns
    returns_grp = returns_df.groupby(["pan", "quarter"]).agg(
        Returns_TDS_Total=("tds", "sum"),
        Returns_Sections=("section", join_unique),
        Returns_Txn_Count=("tds", "count"),
    ).reset_index()

    # Merge
    reco = books_grp.merge(returns_grp, on=["pan", "quarter"], how="outer")

    # Fill NaN with 0 for numeric, "" for string
    for c in ["Books_TDS_Total", "Returns_TDS_Total", "Books_Base_Total",
              "Books_Txn_Count", "Returns_Txn_Count"]:
        if c in reco.columns:
            reco[c] = reco[c].fillna(0)
    for c in ["Books_Sections", "Returns_Sections"]:
        reco[c] = reco[c].fillna("")

    reco["Deductee_Name"] = reco["pan"].map(name_master).fillna("")
    reco["Difference"] = reco["Books_TDS_Total"] - reco["Returns_TDS_Total"]
    reco["Abs_Difference"] = reco["Difference"].abs()
    reco["PAN_Status"] = reco["pan"].apply(validate_pan)

    # Mismatch reason
    def classify(row):
        if row["PAN_Status"] in ("MISSING", "INVALID"):
            return "PAN_ISSUE"

        b_sec = normalize_section_set(row["Books_Sections"])
        r_sec = normalize_section_set(row["Returns_Sections"])
        diff = abs(row["Difference"])

        books_zero = row["Books_TDS_Total"] == 0
        ret_zero = row["Returns_TDS_Total"] == 0

        if books_zero and not ret_zero:
            return "ONLY_IN_RETURNS"
        if ret_zero and not books_zero:
            return "ONLY_IN_BOOKS"

        if diff <= tolerance and b_sec == r_sec:
            return "PERFECT_MATCH"
        if diff <= tolerance:
            return "SECTION_MISMATCH"
        if row["Difference"] > 0:
            return "SHORT_IN_RETURN"
        return "EXCESS_IN_RETURN"

    reco["Match_Status"] = reco.apply(classify, axis=1)

    # Reorder columns
    reco = reco.rename(columns={"pan": "PAN", "quarter": "Quarter"})
    col_order = [
        "Quarter", "PAN", "Deductee_Name", "PAN_Status",
        "Books_TDS_Total", "Returns_TDS_Total", "Difference",
        "Books_Base_Total", "Books_Txn_Count", "Returns_Txn_Count",
        "Books_Sections", "Returns_Sections", "Match_Status"
    ]
    col_order = [c for c in col_order if c in reco.columns]
    reco = reco[col_order].sort_values(["Quarter", "Match_Status", "PAN"])

    return reco


# ═══════════════════════════════════════════════════════════════
# CHALLAN RECONCILIATION
# ═══════════════════════════════════════════════════════════════

def run_challan_reco(books_df, challans_df):
    """
    Section + Quarter level reconciliation of challans vs books deductions.
    """
    if challans_df is None or challans_df.empty:
        return pd.DataFrame(columns=[
            "Quarter", "Section", "Books_TDS_Total",
            "Challan_Total", "Difference", "Status"
        ])

    # Group books by section + quarter
    books_sec = books_df.groupby(["quarter", "section"]).agg(
        Books_TDS_Total=("tds", "sum")
    ).reset_index()

    # Group challans by section + quarter
    chal_sec = challans_df.groupby(["quarter", "section"]).agg(
        Challan_Total=("challan_amount", "sum")
    ).reset_index()

    # Merge
    chal_reco = books_sec.merge(chal_sec, on=["quarter", "section"], how="outer")
    chal_reco["Books_TDS_Total"] = chal_reco["Books_TDS_Total"].fillna(0)
    chal_reco["Challan_Total"] = chal_reco["Challan_Total"].fillna(0)
    chal_reco["Difference"] = chal_reco["Books_TDS_Total"] - chal_reco["Challan_Total"]

    def challan_status(row):
        if row["Difference"] == 0:
            return "FULLY_DEPOSITED"
        elif row["Difference"] > 0:
            return "SHORT_DEPOSIT"
        else:
            return "EXCESS_DEPOSIT"

    chal_reco["Status"] = chal_reco.apply(challan_status, axis=1)
    chal_reco = chal_reco.rename(columns={"quarter": "Quarter", "section": "Section"})
    return chal_reco.sort_values(["Quarter", "Section"])


# ═══════════════════════════════════════════════════════════════
# SUMMARY GENERATION
# ═══════════════════════════════════════════════════════════════

def generate_summary(reco_df, books_df, returns_df, challans_df=None):
    """Generate executive summary metrics."""
    total_books_tds = reco_df["Books_TDS_Total"].sum()
    total_returns_tds = reco_df["Returns_TDS_Total"].sum()
    total_diff = total_books_tds - total_returns_tds

    matched = reco_df[reco_df["Match_Status"] == "PERFECT_MATCH"]
    match_pct = (len(matched) / len(reco_df) * 100) if len(reco_df) > 0 else 0

    total_challans = 0
    if challans_df is not None and not challans_df.empty:
        total_challans = challans_df["challan_amount"].sum()

    summary = {
        "total_books_tds": total_books_tds,
        "total_returns_tds": total_returns_tds,
        "net_difference": total_diff,
        "total_challans_deposited": total_challans,
        "total_deductees_books": books_df["pan"].nunique(),
        "total_deductees_returns": returns_df["pan"].nunique(),
        "total_reco_entries": len(reco_df),
        "matched_entries": len(matched),
        "match_percentage": match_pct,
        "mismatched_entries": len(reco_df) - len(matched),
        "total_books_txns": len(books_df),
        "total_returns_txns": len(returns_df),
        "short_in_return": len(reco_df[reco_df["Match_Status"] == "SHORT_IN_RETURN"]),
        "excess_in_return": len(reco_df[reco_df["Match_Status"] == "EXCESS_IN_RETURN"]),
        "only_in_books": len(reco_df[reco_df["Match_Status"] == "ONLY_IN_BOOKS"]),
        "only_in_returns": len(reco_df[reco_df["Match_Status"] == "ONLY_IN_RETURNS"]),
        "section_mismatch": len(reco_df[reco_df["Match_Status"] == "SECTION_MISMATCH"]),
        "pan_issues": len(reco_df[reco_df["Match_Status"] == "PAN_ISSUE"]),
    }

    # Quarter-wise summary
    qtr_summary = reco_df.groupby("Quarter").agg(
        Books_TDS=("Books_TDS_Total", "sum"),
        Returns_TDS=("Returns_TDS_Total", "sum"),
        Difference=("Difference", "sum"),
        Deductees=("PAN", "nunique"),
        Matched=("Match_Status", lambda x: (x == "PERFECT_MATCH").sum()),
        Mismatched=("Match_Status", lambda x: (x != "PERFECT_MATCH").sum()),
    ).reset_index()

    # Status breakdown
    status_breakdown = reco_df["Match_Status"].value_counts().reset_index()
    status_breakdown.columns = ["Status", "Count"]

    # Section-wise summary
    section_summary = reco_df.copy()
    # Combine sections from both sources
    all_sections_books = books_df.groupby("section")["tds"].sum().reset_index()
    all_sections_books.columns = ["Section", "Books_TDS"]
    all_sections_ret = returns_df.groupby("section")["tds"].sum().reset_index()
    all_sections_ret.columns = ["Section", "Returns_TDS"]

    sec_summary = all_sections_books.merge(all_sections_ret, on="Section", how="outer").fillna(0)
    sec_summary["Difference"] = sec_summary["Books_TDS"] - sec_summary["Returns_TDS"]
    sec_summary = sec_summary.sort_values("Books_TDS", ascending=False)

    return summary, qtr_summary, status_breakdown, sec_summary


# ═══════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ═══════════════════════════════════════════════════════════════

def export_to_excel(reco_df, books_df, returns_df, challan_reco_df,
                    quality_df, duplicate_df, qtr_summary, sec_summary, summary_dict):
    """Export all reconciliation data to a styled Excel file and return bytes."""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        # ── Formats ──
        header_fmt = workbook.add_format({
            'bold': True, 'font_color': 'white', 'bg_color': '#4F81BD',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter',
            'align': 'center', 'font_size': 11
        })
        currency_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        number_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1})
        text_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        pct_fmt = workbook.add_format({'num_format': '0.0%', 'border': 1})
        match_fmt = workbook.add_format({
            'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1
        })
        mismatch_fmt = workbook.add_format({
            'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1
        })
        title_fmt = workbook.add_format({
            'bold': True, 'font_size': 14, 'font_color': '#4F81BD'
        })
        subtitle_fmt = workbook.add_format({
            'bold': True, 'font_size': 11, 'font_color': '#333333'
        })
        kpi_value_fmt = workbook.add_format({
            'bold': True, 'font_size': 16, 'num_format': '#,##0.00',
            'font_color': '#4F81BD'
        })
        kpi_label_fmt = workbook.add_format({
            'font_size': 10, 'font_color': '#666666'
        })

        # ═══════ Sheet 1: Executive Summary ═══════
        ws_sum = workbook.add_worksheet("Executive_Summary")
        writer.sheets["Executive_Summary"] = ws_sum

        ws_sum.set_column('A:A', 25)
        ws_sum.set_column('B:B', 20)
        ws_sum.set_column('C:C', 20)
        ws_sum.set_column('D:D', 20)
        ws_sum.set_column('E:F', 15)

        row = 0
        ws_sum.write(row, 0, "TDS RECONCILIATION — EXECUTIVE SUMMARY", title_fmt)
        row += 1
        ws_sum.write(row, 0, f"Generated: {datetime.now().strftime('%d-%b-%Y %H:%M')}", subtitle_fmt)
        row += 3

        # KPI Cards
        ws_sum.write(row, 0, "TOTAL BOOKS TDS", kpi_label_fmt)
        ws_sum.write(row, 1, "TOTAL RETURNS TDS", kpi_label_fmt)
        ws_sum.write(row, 2, "NET DIFFERENCE", kpi_label_fmt)
        ws_sum.write(row, 3, "MATCH %", kpi_label_fmt)
        row += 1
        ws_sum.write(row, 0, summary_dict["total_books_tds"], kpi_value_fmt)
        ws_sum.write(row, 1, summary_dict["total_returns_tds"], kpi_value_fmt)
        ws_sum.write(row, 2, summary_dict["net_difference"], kpi_value_fmt)
        ws_sum.write(row, 3, f"{summary_dict['match_percentage']:.1f}%", kpi_value_fmt)
        row += 3

        # Status breakdown
        ws_sum.write(row, 0, "RECONCILIATION STATUS BREAKDOWN", subtitle_fmt)
        row += 1
        labels = [
            ("Perfect Match", summary_dict["matched_entries"]),
            ("Short in Return", summary_dict["short_in_return"]),
            ("Excess in Return", summary_dict["excess_in_return"]),
            ("Only in Books", summary_dict["only_in_books"]),
            ("Only in Returns", summary_dict["only_in_returns"]),
            ("Section Mismatch", summary_dict["section_mismatch"]),
            ("PAN Issues", summary_dict["pan_issues"]),
        ]
        for label, val in labels:
            ws_sum.write(row, 0, label, text_fmt)
            ws_sum.write(row, 1, val, number_fmt)
            row += 1

        row += 2
        # Quarter-wise summary table
        ws_sum.write(row, 0, "QUARTER-WISE SUMMARY", subtitle_fmt)
        row += 1
        qtr_summary.to_excel(writer, sheet_name="Executive_Summary",
                             startrow=row, index=False)
        row += len(qtr_summary) + 3

        # Section-wise summary
        ws_sum.write(row, 0, "SECTION-WISE SUMMARY", subtitle_fmt)
        row += 1
        sec_summary.to_excel(writer, sheet_name="Executive_Summary",
                             startrow=row, index=False)

        # ═══════ Sheet 2: Detailed Reco ═══════
        reco_df.to_excel(writer, sheet_name="Detailed_Deductee_Reco", index=False)
        ws_reco = writer.sheets["Detailed_Deductee_Reco"]
        ws_reco.freeze_panes(1, 0)
        ws_reco.autofilter(0, 0, len(reco_df), len(reco_df.columns) - 1)
        for i, col in enumerate(reco_df.columns):
            max_len = max(reco_df[col].astype(str).str.len().max(), len(col)) + 2
            ws_reco.set_column(i, i, min(max_len, 30))

        # ═══════ Sheet 3: Challan Reco ═══════
        if challan_reco_df is not None and not challan_reco_df.empty:
            challan_reco_df.to_excel(writer, sheet_name="Challan_Reconciliation", index=False)
            ws_chal = writer.sheets["Challan_Reconciliation"]
            ws_chal.freeze_panes(1, 0)
            ws_chal.autofilter(0, 0, len(challan_reco_df), len(challan_reco_df.columns) - 1)
            for i, col in enumerate(challan_reco_df.columns):
                max_len = max(challan_reco_df[col].astype(str).str.len().max(), len(col)) + 2
                ws_chal.set_column(i, i, min(max_len, 25))

        # ═══════ Sheet 4: Data Quality ═══════
        if quality_df is not None and not quality_df.empty:
            quality_df.to_excel(writer, sheet_name="Data_Quality_Report", index=False)

        # ═══════ Sheet 5: Duplicates ═══════
        if duplicate_df is not None and not duplicate_df.empty:
            duplicate_df.to_excel(writer, sheet_name="Potential_Duplicates", index=False)

        # ═══════ Sheet 6: Raw Books ═══════
        books_df.to_excel(writer, sheet_name="Raw_Books_Data", index=False)

        # ═══════ Sheet 7: Raw Returns ═══════
        returns_df.to_excel(writer, sheet_name="Raw_Returns_Data", index=False)

    output.seek(0)
    return output.getvalue()
