import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="GST Report Analyzer", layout="wide")

st.title("GST Excel Report Analyzer")

uploaded_file = st.file_uploader(
    "Upload GST Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file:

    # =========================
    # READ EXCEL
    # =========================
    # Actual headers start from row 5
    # =========================
    df = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        header=4
    )

    # Remove fully empty rows
    df = df.dropna(how="all")

    # Clean column names
    df.columns = [str(col).strip() for col in df.columns]

    st.success("File Uploaded Successfully")

    # =========================
    # EXACT COLUMN NAMES
    # =========================
    gstin_col = "GSTIN of supplier"
    trade_col = "Trade/Legal name"
    taxable_col = "Taxable Value (₹)"
    igst_col = "Integrated Tax(₹)"
    cgst_col = "Central Tax(₹)"
    sgst_col = "State/UT Tax(₹)"

    # =========================
    # CHECK REQUIRED COLUMNS
    # =========================
    required_cols = [
        gstin_col,
        trade_col,
        taxable_col,
        igst_col,
        cgst_col,
        sgst_col
    ]

    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        st.error(f"Missing columns: {missing_cols}")

        st.write("Detected Columns:")
        st.write(list(df.columns))

        st.stop()

    # =========================
    # NUMERIC CONVERSION
    # =========================
    numeric_cols = [
        taxable_col,
        igst_col,
        cgst_col,
        sgst_col
    ]

    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # =========================
    # CREATE TYPE COLUMN
    # =========================
    df["Type"] = df[gstin_col].astype(str).apply(
        lambda x: "Intrastate"
        if x.startswith("33")
        else "Interstate"
    )

    # =========================
    # CREATE RATE COLUMN
    # =========================
    df["Rate"] = (
        (
            df[igst_col]
            + df[cgst_col]
            + df[sgst_col]
        )
        / df[taxable_col].replace(0, 1)
    ) * 100

    # Round Rate
    df["Rate"] = df["Rate"].round().astype(int)

    # =========================
    # INSERT TYPE COLUMN
    # BETWEEN COLUMN 2 & 3
    # =========================
    cols = list(df.columns)

    cols.remove("Type")
    cols.insert(2, "Type")

    df = df[cols]

    # =========================
    # INSERT RATE COLUMN
    # BETWEEN COLUMN 8 & 9
    # =========================
    cols = list(df.columns)

    cols.remove("Rate")
    cols.insert(9, "Rate")

    df = df[cols]

    st.subheader("Processed Data Preview")
    st.dataframe(df.head(20), use_container_width=True)

    # =========================
    # REPORT GENERATION
    # =========================
    target_rates = [5, 12, 18, 28]

    report_groups = {}

    for report_type in ["Intrastate", "Interstate"]:

        type_df = df[df["Type"] == report_type]

        # Standard Rates
        for rate in target_rates:

            filtered = type_df[type_df["Rate"] == rate]

            if not filtered.empty:
                report_groups[
                    f"{report_type} {rate}%"
                ] = filtered

        # Mixed Rates
        mixed = type_df[
            ~type_df["Rate"].isin(target_rates)
        ]

        if not mixed.empty:
            report_groups[
                f"{report_type} Mixed"
            ] = mixed

    # =========================
    # CREATE ZIP
    # =========================
    zip_buffer = BytesIO()

    with ZipFile(zip_buffer, "a") as zip_file:

        for report_name, report_df in report_groups.items():

            output = BytesIO()

            wb = Workbook()

            # =========================
            # DETAILED REPORT
            # =========================
            ws1 = wb.active
            ws1.title = "Detailed Report"

            for row in dataframe_to_rows(
                report_df,
                index=False,
                header=True
            ):
                ws1.append(row)

            # =========================
            # SUMMARY REPORT
            # =========================
            summary = (
                report_df.groupby(
                    [
                        gstin_col,
                        trade_col,
                        "Type",
                        "Rate"
                    ],
                    dropna=False
                )
                .agg({
                    taxable_col: "sum",
                    igst_col: "sum",
                    cgst_col: "sum",
                    sgst_col: "sum"
                })
                .reset_index()
            )

            ws2 = wb.create_sheet(
                title="Summary"
            )

            for row in dataframe_to_rows(
                summary,
                index=False,
                header=True
            ):
                ws2.append(row)

            wb.save(output)

            zip_file.writestr(
                f"{report_name}.xlsx",
                output.getvalue()
            )

    zip_buffer.seek(0)

    st.success("Reports Generated Successfully")

    st.download_button(
        label="Download All Reports (ZIP)",
        data=zip_buffer,
        file_name="GST_Reports.zip",
        mime="application/zip"
    )
