"""Quick test script — run this to verify engine logic."""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from reco_engine import (
    load_file, prepare_books, prepare_returns, prepare_challans,
    run_reconciliation, run_challan_reco, generate_summary,
    check_data_quality, detect_duplicates, export_to_excel
)

DATA_DIR = r"C:\Users\harshal.bankar\Desktop\Automation\TDS 26Q RECO"

def main():
    with open(os.path.join(DATA_DIR, "deductions.xlsx"), "rb") as f:
        books_bytes = f.read()
    with open(os.path.join(DATA_DIR, "returns.xlsx"), "rb") as f:
        returns_bytes = f.read()
    with open(os.path.join(DATA_DIR, "challans.xlsx"), "rb") as f:
        challans_bytes = f.read()

    books_raw = load_file(books_bytes, "deductions.xlsx")
    returns_raw = load_file(returns_bytes, "returns.xlsx")
    challans_raw = load_file(challans_bytes, "challans.xlsx")

    books_df = prepare_books(books_raw)
    returns_df = prepare_returns(returns_raw)
    challans_df = prepare_challans(challans_raw)

    print(f"Books: {len(books_df)} rows")
    print(f"Returns: {len(returns_df)} rows")
    print(f"Challans: {len(challans_df)} rows")

    reco_df = run_reconciliation(books_df, returns_df, tolerance=1.0)
    print(f"\nReco entries: {len(reco_df)}")
    print("\nMatch status breakdown:")
    print(reco_df["Match_Status"].value_counts().to_string())

    summary, qtr_summary, status_breakdown, sec_summary = generate_summary(
        reco_df, books_df, returns_df, challans_df
    )

    print("\n=== EXECUTIVE SUMMARY ===")
    for k, v in summary.items():
        if isinstance(v, float):
            print(f"  {k}: ₹{v:,.2f}" if "tds" in k or "diff" in k or "challan" in k else f"  {k}: {v:.1f}")
        else:
            print(f"  {k}: {v}")

    challan_reco = run_challan_reco(books_df, challans_df)
    print(f"\nChallan reco entries: {len(challan_reco)}")

    quality = check_data_quality(books_df, returns_df)
    print(f"Quality issues: {len(quality)}")

    dupes = detect_duplicates(books_df, returns_df)
    print(f"Potential duplicates: {len(dupes)}")

    excel_bytes = export_to_excel(
        reco_df, books_df, returns_df, challan_reco,
        quality, dupes, qtr_summary, sec_summary, summary
    )
    print(f"\nExcel report: {len(excel_bytes):,} bytes")
    print("\n✅ ALL TESTS PASSED")

if __name__ == "__main__":
    main()
