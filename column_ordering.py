import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel Column Reorder Tool")

# Upload rule file
rule_file = st.file_uploader(
    "Upload Rule File (CSV or Excel)",
    type=["csv", "xlsx"]
)

# Upload excel file
excel_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx"]
)

if rule_file is not None and excel_file is not None:

    # ----------------------
    # Load rule file sheets
    # ----------------------
    try:
        if rule_file.name.endswith('.xlsx'):
            rule_excel = pd.ExcelFile(rule_file)
            rule_sheets = rule_excel.sheet_names

            selected_rule_sheet = st.selectbox(
                "Select Rule Sheet",
                rule_sheets
            )

            rule_df = pd.read_excel(
                rule_file,
                sheet_name=selected_rule_sheet
            )
        else:
            rule_df = pd.read_csv(rule_file)
            selected_rule_sheet = "CSV Rule"

        # Validate rule file
        if 'new_order' not in rule_df.columns:
            st.error("Rule file must contain a 'new_order' column")
            st.stop()

    except Exception as e:
        st.error(f"Error reading rule file: {e}")
        st.stop()

    # ----------------------
    # Load uploaded Excel sheets
    # ----------------------
    try:
        excel_data = pd.ExcelFile(excel_file)
        excel_sheets = excel_data.sheet_names

        selected_data_sheet = st.selectbox(
            "Select Excel Sheet to Rearrange",
            excel_sheets
        )

        df = pd.read_excel(
            excel_file,
            sheet_name=selected_data_sheet
        )

    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        st.stop()

    # ----------------------
    # Extract reorder rule
    # ----------------------
    desired_order = rule_df['new_order'].dropna().astype(str).tolist()

    # ----------------------
    # Reorder columns
    # Allow missing columns
    # ----------------------
    reordered_df = pd.DataFrame()

    for col in desired_order:
        if col in df.columns:
            reordered_df[col] = df[col]
        else:
            reordered_df[col] = ""

    # ----------------------
    # Success message
    # ----------------------
    st.success("Columns reordered successfully")

    st.write(f"Rule Sheet Used: {selected_rule_sheet}")
    st.write(f"Excel Sheet Rearranged: {selected_data_sheet}")

    # ----------------------
    # Safe display for duplicate columns
    # ----------------------
    display_df = reordered_df.copy()

    cols = pd.Series(display_df.columns)

    for dup in cols[cols.duplicated()].unique():
        dup_indexes = cols[cols == dup].index.tolist()

        for i, idx in enumerate(dup_indexes):
            if i > 0:
                cols[idx] = f"{dup}_{i}"

    display_df.columns = cols

    st.dataframe(display_df)

    # ----------------------
    # Create downloadable Excel
    # ----------------------
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        reordered_df.to_excel(
            writer,
            sheet_name=selected_data_sheet,
            index=False
        )

    output.seek(0)

    # ----------------------
    # Download button
    # ----------------------
    st.download_button(
        label="Download Reordered Excel",
        data=output,
        file_name="reordered_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_reordered_excel"
    )
