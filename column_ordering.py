import streamlit as st
import pandas as pd
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
    if rule_file.name.endswith('.xlsx'):
        rule_excel = pd.ExcelFile(rule_file)
        rule_sheets = rule_excel.sheet_names

        selected_rule_sheet = st.selectbox(
            "Select Rule Sheet",
            rule_sheets
        )

        rule_df = pd.read_excel(rule_file, sheet_name=selected_rule_sheet)

    else:
        rule_df = pd.read_csv(rule_file)

    # ----------------------
    # Load uploaded Excel sheets
    # ----------------------
    excel_data = pd.ExcelFile(excel_file)
    excel_sheets = excel_data.sheet_names

    selected_data_sheet = st.selectbox(
        "Select Excel Sheet to Rearrange",
        excel_sheets
    )

    df = pd.read_excel(excel_file, sheet_name=selected_data_sheet)

    # ----------------------
    # Extract reorder rule
    # ----------------------
    desired_order = rule_df['new_order'].tolist()

    # Validate columns
    missing_cols = [col for col in desired_order if col not in df.columns]

    if missing_cols:
        st.error(f"Missing columns in Excel file: {missing_cols}")
    else:
        reordered_df = df[desired_order]

        st.success("Columns reordered successfully")

        st.write(f"Rule Sheet Used: {selected_rule_sheet if rule_file.name.endswith('.xlsx') else 'CSV Rule'}")
        st.write(f"Excel Sheet Rearranged: {selected_data_sheet}")

        st.dataframe(reordered_df)

        # Download output
        output_file = "reordered_output.xlsx"

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            reordered_df.to_excel(
                writer,
                sheet_name=selected_data_sheet,
                index=False
            )

        with open(output_file, "rb") as file:
            st.download_button(
                label="Download Reordered Excel",
                data=file,
                file_name="reordered_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
