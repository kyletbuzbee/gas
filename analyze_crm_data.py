import pandas as pd
import os
import json
import glob
from datetime import datetime

def analyze_csvs(directory):
    report = {
        "files": {},
        "cross_references": {},
        "potential_issues": []
    }

    csv_files = glob.glob(os.path.join(directory, "*.csv"))

    all_dfs = {}

    for file_path in csv_files:
        filename = os.path.basename(file_path)
        try:
            # Read first few lines to detect headers and skip empty trailing rows
            df = pd.read_csv(file_path).dropna(how='all')
            all_dfs[filename] = df

            report["files"][filename] = {
                "row_count": len(df),
                "columns": list(df.columns),
                "null_counts": df.isnull().sum().to_dict(),
                "sample_ids": df.iloc[:, 0].head(5).tolist() if not df.empty else []
            }
        except Exception as e:
            report["potential_issues"].append(f"Error reading {filename}: {str(e)}")

    # Check for Company ID consistency
    company_ids = {}
    for filename, df in all_dfs.items():
        id_cols = [col for col in df.columns if 'ID' in col and 'Company' in col]
        name_cols = [col for col in df.columns if 'Name' in col and 'Company' in col] or [col for col in df.columns if col == 'Company']

        if id_cols and name_cols:
            id_col = id_cols[0]
            name_col = name_cols[0]
            for _, row in df.iterrows():
                cid = str(row[id_col]).strip()
                name = str(row[name_col]).strip()
                if cid != 'nan' and name != 'nan':
                    if cid not in company_ids:
                        company_ids[cid] = set()
                    company_ids[cid].add(name)

    # Report name mismatches for same ID
    for cid, names in company_ids.items():
        if len(names) > 1:
            report["potential_issues"].append({
                "issue": "Mismatched names for ID",
                "id": cid,
                "names": list(names)
            })

    # Detect duplicate IDs in Outreach
    if "Outreach.csv" in all_dfs:
        outreach_df = all_dfs["Outreach.csv"]
        duplicates = outreach_df[outreach_df.duplicated(['Outreach ID'], keep=False)]
        if not duplicates.empty:
            report["potential_issues"].append({
                "issue": "Duplicate Outreach IDs",
                "ids": duplicates['Outreach ID'].unique().tolist()
            })

    return report

if __name__ == "__main__":
    results = analyze_csvs("csv")
    with open("crm_analysis_blueprint.json", "w") as f:
        json.dump(results, f, indent=4)
    print("Analysis complete. Check crm_analysis_blueprint.json for details.")