import sys
from pathlib import Path

def safe_imports():
    try:
        import pandas as pd
        import openpyxl  # noqa: F401  (engine for .xlsx)
        return pd
    except ImportError:
        print("Required packages not found.\n"
              "Please install them with:\n"
              "  py -m pip install --user pandas openpyxl")
        sys.exit(1)

def normalize_columns(df):
    """
    Trim whitespace and make a mapping that finds columns case-insensitively.
    Returns a helper that resolves canonical names -> actual df columns.
    """
    df.columns = [str(c).strip() for c in df.columns]
    lowered = {c.lower(): c for c in df.columns}

    def resolve(*candidates):
        for cand in candidates:
            cand_l = cand.lower().strip()
            if cand_l in lowered:
                return lowered[cand_l]
        return None

    return resolve

def compare_and_update_excel_files(pd):
    """
    Compare two Excel files and update 'User Name' in Laptop Hostname.xlsx
    based on matching 'Login ID' from HC Report.xlsx.
    """
    temp_dir = Path(r"C:\Temp")
    hc_report_path = temp_dir / "HC Report.xlsx"
    hostname_path = temp_dir / "Laptop Hostname.xlsx"
    backup_path = temp_dir / "Laptop Hostname_backup.xlsx"

    # Basic file checks
    if not hc_report_path.exists():
        print(f"Error: HC Report.xlsx not found at {hc_report_path}")
        return False
    if not hostname_path.exists():
        print(f"Error: Laptop Hostname.xlsx not found at {hostname_path}")
        return False

    try:
        print("Reading Excel files...")
        hc = pd.read_excel(hc_report_path, sheet_name=0, engine="openpyxl")
        host = pd.read_excel(hostname_path, sheet_name=0, engine="openpyxl")

        print(f"HC Report rows: {len(hc)}")
        print(f"Hostname file rows: {len(host)}")

        # Resolve columns flexibly
        hc_resolve = normalize_columns(hc)
        host_resolve = normalize_columns(host)

        col_login_hc = hc_resolve("Login ID", "LoginID", "UserID", "User ID")
        col_user_hc = hc_resolve("User Name", "Username", "User", "Display Name", "Full Name")

        col_login_host = host_resolve("Login ID", "LoginID", "UserID", "User ID")
        col_user_host = host_resolve("User Name", "Username", "User", "Display Name", "Full Name")

        # Validate mandatory columns
        missing = []
        if not col_login_hc:  missing.append("HC Report: 'Login ID'")
        if not col_user_hc:   missing.append("HC Report: 'User Name'")
        if not col_login_host: missing.append("Hostname: 'Login ID'")
        if missing:
            print("Error: Required column(s) not found -> " + "; ".join(missing))
            print("Tip: Check for exact column headers or trailing spaces in Excel.")
            return False

        # Ensure 'User Name' column exists in host; create if absent
        if not col_user_host:
            col_user_host = "User Name"
            if col_user_host not in host.columns:
                host[col_user_host] = ""

        # Clean up key columns (normalize case and strip)
        hc[col_login_hc] = hc[col_login_hc].astype(str).str.strip().str.lower()
        hc[col_user_hc]  = hc[col_user_hc].astype(str).str.strip()
        host[col_login_host] = host[col_login_host].astype(str).str.strip().str.lower()

        # Drop rows in HC with missing keys/usernames
        hc_clean = hc.dropna(subset=[col_login_hc, col_user_hc])
        hc_clean = hc_clean[(hc_clean[col_login_hc] != "") & (hc_clean[col_user_hc] != "")]

        # Build mapping via merge (left join to keep all host rows)
        merged = host.merge(
            hc_clean[[col_login_hc, col_user_hc]].drop_duplicates(subset=[col_login_hc]),
            how="left",
            left_on=col_login_host,
            right_on=col_login_hc,
            suffixes=("", "_hc"),
        )

        # Count matches before writing
        matches = merged[col_user_hc].notna().sum()

        # Update/overwrite the host User Name column with values from HC where available
        merged[col_user_host] = merged[col_user_hc].combine_first(merged[col_user_host])

        # Stats
        total_login = merged[col_login_host].ne("").sum()
        no_match = total_login - matches
        print("\nUpdate Summary:")
        print(f"- Records with Login ID: {total_login}")
        print(f"- Records updated from HC: {int(matches)}")
        print(f"- Login IDs with no match: {int(no_match)}")

        # Create backup if it doesn't already exist
        if backup_path.exists():
            print(f"Backup file already exists: {backup_path}")
        else:
            host.to_excel(backup_path, index=False, engine="openpyxl")
            print(f"Backup created: {backup_path}")

        # Save updated file (drop merge helper columns if present)
        cols_to_drop = [c for c in [col_login_hc, col_user_hc] if c in merged.columns and c != col_login_host and c != col_user_host]
        merged.drop(columns=[c for c in cols_to_drop if c in merged.columns], inplace=True, errors="ignore")

        merged.to_excel(hostname_path, index=False, engine="openpyxl")
        print(f"\nUpdated file saved: {hostname_path}")

        # Show a small sample of updated rows
        if matches > 0:
            sample = merged.loc[merged[col_user_hc].notna(), [col_login_host, col_user_host]].head()
            print("\nSample of updated records:")
            print(sample.to_string(index=False))

        return True

    except Exception as e:
        print(f"An error occurred: {e}")
        return False

def main():
    pd = safe_imports()
    print("Excel File Comparison and Update Tool")
    print("=" * 40)
    ok = compare_and_update_excel_files(pd)
    if ok:
        print("\n✓ Process completed successfully!")
    else:
        print("\n✗ Process failed. Please check the messages above.")

if __name__ == "__main__":
    main()
