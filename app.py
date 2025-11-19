import io
from datetime import datetime
import re

import numpy as np
import pandas as pd
import streamlit as st

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ------------------------
# LSC Buyer Journey Intelligence Engine
# ------------------------

REQUIRED_COLUMNS = [
    "email",
    "first_name",
    "last_name",
    "company",
    "engagement_date",
    "engagement_type",
]

STAGE_ORDER = ["Awareness", "Problem Definition", "Solution Exploration", "Vendor Evaluation"]

STAGE_COLOR = {
    "Vendor Evaluation": "FF0000",  # red
    "Solution Exploration": "FFA500",  # orange
    "Problem Definition": "FFFF00",  # yellow
    "Awareness": "C0C0C0",  # gray
}


def normalize_columns(df):
    # Lowercase and replace spaces
    mapping = {}
    cols = list(df.columns)
    lowered = [c.strip().lower().replace(" ", "_") for c in cols]
    for orig, low in zip(cols, lowered):
        mapping[orig] = low
    df = df.rename(columns=mapping)
    return df


def find_required_columns(df):
    df = normalize_columns(df)
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    return missing, df


def parse_dates(df, col="engagement_date"):
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def stage_from_text(text: str):
    if pd.isna(text):
        return None
    t = str(text).lower()
    # heuristics
    if re.search(r"(demo|trial|pricing|rfp|quote|proposal|purchase|evaluation|vendor|vendor eval)", t):
        return "Vendor Evaluation"
    if re.search(r"(solution|product|datasheet|whitepaper|case study|webinar|feature|use case)", t):
        return "Solution Exploration"
    if re.search(r"(problem|pain|challenge|need|gap|issue|requirement|define|diagnos)", t):
        return "Problem Definition"
    if re.search(r"(newsletter|social|blog|press|awareness|announcement|advertis|marketing)", t):
        return "Awareness"
    # default fallback: Awareness
    return "Awareness"


@st.cache_data
def process_dataframe(df: pd.DataFrame, min_engagements: int = 5):
    # normalize
    df = normalize_columns(df)
    df = parse_dates(df, "engagement_date")

    # Ensure necessary columns exist with defaults
    for c in ["email", "first_name", "last_name", "company", "engagement_type", "product"]:
        if c not in df.columns:
            df[c] = np.nan

    # derive stage per row
    df["stage"] = df["engagement_type"].apply(stage_from_text)

    # engagement count per person
    df["email_clean"] = df["email"].astype(str).str.lower().str.strip()
    grouped = df.groupby("email_clean")

    individuals = grouped.agg(
        first_name=("first_name", lambda s: s.dropna().astype(str).iloc[0] if s.dropna().shape[0] > 0 else ""),
        last_name=("last_name", lambda s: s.dropna().astype(str).iloc[0] if s.dropna().shape[0] > 0 else ""),
        company=("company", lambda s: s.dropna().astype(str).iloc[0] if s.dropna().shape[0] > 0 else ""),
        email=("email", lambda s: s.dropna().astype(str).iloc[0] if s.dropna().shape[0] > 0 else s.name),
        total_engagements=("engagement_date", lambda s: s.count()),
        first_seen=("engagement_date", lambda s: s.min()),
        last_seen=("engagement_date", lambda s: s.max()),
    ).reset_index(drop=True)

    # compute counts by stage for each individual
    stage_counts = df.pivot_table(index="email_clean", columns="stage", values="engagement_date", aggfunc="count", fill_value=0)
    stage_counts = stage_counts.reset_index().rename_axis(None, axis=1)

    individuals = individuals.merge(stage_counts, left_on="email", right_on="email_clean", how="left")
    if "email_clean" in individuals.columns:
        individuals = individuals.drop(columns=["email_clean"])  # keep normalized email in 'email'

    # Fill missing stage columns
    for s in STAGE_ORDER:
        if s not in individuals.columns:
            individuals[s] = 0

    # determine buyer stage using a simple priority: if any Vendor Evaluation -> Vendor, else if Solution -> Solution, else Problem, else Awareness
    def pick_stage(row):
        for s in reversed(STAGE_ORDER):
            if row.get(s, 0) > 0:
                return s
        return "Awareness"

    individuals["buyer_stage"] = individuals.apply(pick_stage, axis=1)

    # filter by engagements
    individuals_filtered = individuals[individuals["total_engagements"] >= min_engagements].copy()

    # Company-level aggregations
    company_grp = individuals_filtered.groupby("company")
    company_combined = company_grp.agg(
        num_individuals=("email", "nunique"),
        total_engagements=("total_engagements", "sum"),
        first_seen=("first_seen", "min"),
        last_seen=("last_seen", "max"),
    ).reset_index()

    # counts of buyer stages by company
    company_stages = individuals_filtered.pivot_table(index="company", columns="buyer_stage", values="email", aggfunc="count", fill_value=0).reset_index()
    company_summary = company_combined.merge(company_stages, on="company", how="left")

    # Product map: from original df filter to only filtered emails
    emails_keep = set(individuals_filtered["email"].dropna().astype(str).str.lower())
    df["email_clean"] = df["email"].astype(str).str.lower().str.strip()
    df_products = df[df["email_clean"].isin(emails_keep)]
    if "product" not in df_products.columns:
        df_products["product"] = "(unknown)"
    product_map = df_products.groupby(["product"]).agg(
        mentions=("product", "count"),
        unique_companies=("company", lambda s: s.dropna().nunique()),
    ).reset_index().sort_values(by="mentions", ascending=False)

    # Sales hot list: companies with a lot of vendor-eval engagements and high total engagements
    # compute vendor eval per company from original df
    vendor_df = df[df["stage"] == "Vendor Evaluation"].copy()
    vendor_df["company"] = vendor_df["company"].fillna("(unknown)")
    vendor_counts = vendor_df.groupby("company").size().reset_index(name="vendor_eval_engagements")
    sales_hot = company_summary.merge(vendor_counts, on="company", how="left").fillna({"vendor_eval_engagements": 0})
    sales_hot = sales_hot.sort_values(by=["vendor_eval_engagements", "total_engagements"], ascending=False)

    # return all
    return {
        "individuals": individuals_filtered,
        "company_combined": company_combined,
        "company_summary": company_summary,
        "product_map": product_map,
        "sales_hot_list": sales_hot,
        "raw": df,
    }


def excel_with_styles(dfs: dict, filename: str = "report.xlsx") -> bytes:
    # Writes multiple dataframes to an in-memory workbook and applies color-coding
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet, df in dfs.items():
            # sanitize sheet name
            sheet_name = sheet[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        writer.save()

    output.seek(0)
    wb = load_workbook(output)

    # apply coloring rules
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # find header -> find buyer_stage or stage or total_engagements
        headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        header_index = {h: i + 1 for i, h in enumerate(headers) if h}

        # color rows by buyer_stage column if present
        if "buyer_stage" in header_index:
            col_idx = header_index["buyer_stage"]
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                stage = cell.value
                if stage in STAGE_COLOR:
                    fill = PatternFill(start_color=STAGE_COLOR[stage], end_color=STAGE_COLOR[stage], fill_type="solid")
                    cell.fill = fill

        # highlight high engagement rows (>=12) if total_engagements column present
        if "total_engagements" in header_index:
            te_idx = header_index["total_engagements"]
            for row in range(2, ws.max_row + 1):
                val = ws.cell(row=row, column=te_idx).value
                try:
                    if val is not None and int(val) >= 12:
                        # apply green to the whole row
                        green = PatternFill(start_color="00C853", end_color="00C853", fill_type="solid")
                        for c in range(1, ws.max_column + 1):
                            ws.cell(row=row, column=c).fill = green
                except Exception:
                    pass

    # write back to bytes
    out2 = io.BytesIO()
    wb.save(out2)
    return out2.getvalue()


def enrich_companies_stub(companies: pd.DataFrame) -> pd.DataFrame:
    # Placeholder: integrate external trigger APIs here. Adds a column with a sample trigger flag.
    companies = companies.copy()
    companies["external_triggers"] = "(stub - no data)"
    return companies


def main():
    st.set_page_config(page_title="LSC Buyer Journey Intelligence Engine", layout="wide")
    st.title("LSC Buyer Journey Intelligence Engine")

    st.markdown(
        "Upload an XLSX or CSV export from Life Science Connect. The app validates columns, aggregates engagements to individuals, computes buyer journey stages, and produces an Excel report."
    )

    uploaded = st.file_uploader("Upload Life Science Connect XLSX/CSV", type=["xlsx", "csv"], accept_multiple_files=False)

    min_eng = st.sidebar.number_input("Minimum engagements per individual (filter)", value=5, min_value=1, step=1)
    enrich = st.sidebar.checkbox("Enrich high-intent companies (stub)", value=False)

    if uploaded is None:
        st.info("Upload a .xlsx or .csv file to begin. You can also open the sample below.")
        st.download_button("Download sample CSV", data=sample_csv(), file_name="sample_lsc_export.csv")
        return

    # Read file
    try:
        if uploaded.type == "text/csv" or uploaded.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded)
        else:
            df = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        return

    missing, df = find_required_columns(df)
    if missing:
        st.error(f"Missing required columns (after normalization): {missing}")
        st.write("Columns found:", list(df.columns))
        return

    # Process
    with st.spinner("Processing data — this may take a moment for tens of thousands of rows..."):
        results = process_dataframe(df, min_engagements=min_eng)

    st.success("Processing complete")

    # Show previews
    st.subheader("Individuals (filtered)")
    st.dataframe(results["individuals"].head(200))

    st.subheader("Company Summary")
    st.dataframe(results["company_summary"].head(200))

    st.subheader("Top Products")
    st.dataframe(results["product_map"].head(200))

    st.subheader("Sales Hot List")
    st.dataframe(results["sales_hot_list"].head(200))

    # optional enrichment
    if enrich:
        with st.spinner("Running enrichment (stub)..."):
            enriched = enrich_companies_stub(results["company_summary"])  # placeholder
            st.write(enriched.head(50))
            results["company_summary"] = enriched

    # Prepare Excel
    sheets = {
        "Company_Combined": results["company_combined"],
        "Company_Summary": results["company_summary"],
        "Individuals": results["individuals"],
        "Product_Map": results["product_map"],
        "Sales_Hot_List": results["sales_hot_list"],
    }

    excel_bytes = excel_with_styles(sheets, filename="LSC_Buyer_Journey_Report.xlsx")

    st.download_button(label="Download Excel Report", data=excel_bytes, file_name="LSC_Buyer_Journey_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def sample_csv() -> bytes:
    sample = pd.DataFrame(
        {
            "First Name": ["Alice", "Bob"],
            "Last Name": ["A", "B"],
            "Email": ["alice@example.com", "bob@example.com"],
            "Company": ["Acme Inc", "Beta LLC"],
            "Engagement Date": [datetime.now().isoformat(), datetime.now().isoformat()],
            "Engagement Type": ["webinar - product features", "requested demo"],
            "Product": ["Product X", "Product Y"],
        }
    )
    return sample.to_csv(index=False).encode("utf-8")


if __name__ == "__main__":
    main()
