import streamlit as st
import pandas as pd

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

    # Read rule file
    if rule_file.name.endswith('.csv'):
        rule_df = pd.read_csv(rule_file)
    else:
        rule_df = pd.read_excel(rule_file)

    # Read Excel file
    df = pd.read_excel(excel_file)

    # Extract desired column order
    desired_order = rule_df['new_order'].tolist()

    # Check if columns exist
    missing_cols = [col for col in desired_order if col not in df.columns]

    if missing_cols:
        st.error(f"Missing columns in Excel file: {missing_cols}")
    else:
        # Reorder columns
        reordered_df = df[desired_order]

        st.success("Columns reordered successfully")

        st.dataframe(reordered_df)

        # Download button
        output_file = "reordered_output.xlsx"

        reordered_df.to_excel(output_file, index=False)

        with open(output_file, "rb") as file:
            st.download_button(
                label="Download Reordered Excel",
                data=file,
                file_name="reordered_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
