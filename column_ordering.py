import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel Column Reorder Tool")

# ----------------------
# Helper functions
# ----------------------
def safe_display_df(df):
    """
    Streamlit dataframe display has problems when duplicate column names exist.
    This creates temporary unique column names only for display.
    """
    display_df = df.copy()
    cols = pd.Series(display_df.columns)

    for dup in cols[cols.duplicated()].unique():
        dup_indexes = cols[cols == dup].index.tolist()

        for i, idx in enumerate(dup_indexes):
            if i > 0:
                cols[idx] = f"{dup}_{i}"

    display_df.columns = cols
    return display_df


def clean_sheet_name(name):
    """
    Excel sheet names cannot contain some characters and must be <= 31 characters.
    """
    invalid_chars = ["\\", "/", "*", "?", ":", "[", "]"]

    for ch in invalid_chars:
        name = str(name).replace(ch, "_")

    name = name.strip()

    if name == "":
        name = "Sheet"

    return name[:31]


def make_unique_sheet_name(base_name, used_names):
    """
    Avoid duplicate sheet names in output Excel file.
    """
    base_name = clean_sheet_name(base_name)
    sheet_name = base_name
    counter = 1

    while sheet_name in used_names:
        suffix = f"_{counter}"
        sheet_name = f"{base_name[:31-len(suffix)]}{suffix}"
        counter += 1

    used_names.add(sheet_name)
    return sheet_name


def load_rule_df(rule_bytes, rule_file_name, rule_sheet_name):
    """
    Load selected rule sheet or CSV rule file.
    """
    if rule_file_name.endswith(".xlsx"):
        return pd.read_excel(
            BytesIO(rule_bytes),
            sheet_name=rule_sheet_name
        )
    else:
        return pd.read_csv(BytesIO(rule_bytes))


def reorder_columns(df, rule_df, append_extra_columns=False):
    """
    Reorder dataframe columns based on the 'new_order' column in rule file.
    Missing columns are created as blank columns.
    """
    if "new_order" not in rule_df.columns:
        raise ValueError("Rule file must contain a 'new_order' column")

    desired_order = rule_df["new_order"].dropna().astype(str).tolist()

    reordered_columns = []

    for col in desired_order:
        if col in df.columns:
            selected_col = df[col]

            # If duplicate columns exist in source file, take the first matching one
            if isinstance(selected_col, pd.DataFrame):
                selected_col = selected_col.iloc[:, 0]

            reordered_columns.append(selected_col.reset_index(drop=True))
        else:
            blank_col = pd.Series([""] * len(df), name=col)
            reordered_columns.append(blank_col)

    reordered_df = pd.concat(reordered_columns, axis=1)
    reordered_df.columns = desired_order

    if append_extra_columns:
        extra_columns = [col for col in df.columns if col not in desired_order]

        if extra_columns:
            reordered_df = pd.concat(
                [reordered_df, df[extra_columns].reset_index(drop=True)],
                axis=1
            )

    missing_columns = [col for col in desired_order if col not in df.columns]
    extra_columns = [col for col in df.columns if col not in desired_order]

    return reordered_df, missing_columns, extra_columns


# ----------------------
# Upload files
# ----------------------
rule_file = st.file_uploader(
    "Upload Rule File (CSV or Excel)",
    type=["csv", "xlsx"]
)

excel_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx"]
)

if rule_file is not None and excel_file is not None:

    rule_bytes = rule_file.getvalue()
    excel_bytes = excel_file.getvalue()

    # ----------------------
    # Load rule sheet names
    # ----------------------
    try:
        if rule_file.name.endswith(".xlsx"):
            rule_excel = pd.ExcelFile(BytesIO(rule_bytes))
            rule_sheets = rule_excel.sheet_names
        else:
            rule_sheets = ["CSV Rule"]

    except Exception as e:
        st.error(f"Error reading rule file: {e}")
        st.stop()

    # ----------------------
    # Load Excel sheet names
    # ----------------------
    try:
        excel_data = pd.ExcelFile(BytesIO(excel_bytes))
        excel_sheets = excel_data.sheet_names

    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        st.stop()

    # ----------------------
    # Multiple pairing UI
    # ----------------------
    st.subheader("Pair Rule Sheets with Excel Sheets")

    pair_count = st.number_input(
        "How many rule-sheet and Excel-sheet pairs do you want?",
        min_value=1,
        max_value=20,
        value=1,
        step=1
    )

    pairs = []

    for i in range(pair_count):
        st.markdown(f"### Pair {i + 1}")

        col1, col2 = st.columns(2)

        with col1:
            selected_rule_sheet = st.selectbox(
                f"Select Rule Sheet for Pair {i + 1}",
                rule_sheets,
                key=f"rule_sheet_{i}"
            )

        with col2:
            selected_data_sheet = st.selectbox(
                f"Select Excel Sheet to Reorder for Pair {i + 1}",
                excel_sheets,
                key=f"data_sheet_{i}"
            )

        pairs.append({
            "rule_sheet": selected_rule_sheet,
            "data_sheet": selected_data_sheet
        })

    keep_unselected_sheets = st.checkbox(
        "Keep unselected Excel sheets in the output file",
        value=True
    )

    append_extra_columns = st.checkbox(
        "Append extra columns not listed in the rule at the end",
        value=False
    )

    # ----------------------
    # Process button
    # ----------------------
    if st.button("Reorder Columns"):

        try:
            output = BytesIO()
            used_output_sheet_names = set()

            reordered_results = {}
            summary_rows = []

            # ----------------------
            # Reorder selected sheets
            # ----------------------
            for pair in pairs:
                rule_sheet = pair["rule_sheet"]
                data_sheet = pair["data_sheet"]

                rule_df = load_rule_df(
                    rule_bytes,
                    rule_file.name,
                    rule_sheet
                )

                df = pd.read_excel(
                    BytesIO(excel_bytes),
                    sheet_name=data_sheet
                )

                reordered_df, missing_columns, extra_columns = reorder_columns(
                    df,
                    rule_df,
                    append_extra_columns=append_extra_columns
                )

                output_sheet_name = make_unique_sheet_name(
                    data_sheet,
                    used_output_sheet_names
                )

                reordered_results[output_sheet_name] = reordered_df

                summary_rows.append({
                    "Rule Sheet Used": rule_sheet,
                    "Excel Sheet Reordered": data_sheet,
                    "Output Sheet Name": output_sheet_name,
                    "Missing Columns Created as Blank": len(missing_columns),
                    "Extra Columns Not in Rule": len(extra_columns)
                })

            # ----------------------
            # Create output Excel
            # ----------------------
            with pd.ExcelWriter(output, engine="openpyxl") as writer:

                # Write reordered selected sheets first
                for sheet_name, reordered_df in reordered_results.items():
                    reordered_df.to_excel(
                        writer,
                        sheet_name=sheet_name,
                        index=False
                    )

                # Keep unselected sheets if selected
                if keep_unselected_sheets:
                    selected_data_sheets = [pair["data_sheet"] for pair in pairs]

                    for sheet in excel_sheets:
                        if sheet not in selected_data_sheets:
                            original_df = pd.read_excel(
                                BytesIO(excel_bytes),
                                sheet_name=sheet
                            )

                            output_sheet_name = make_unique_sheet_name(
                                sheet,
                                used_output_sheet_names
                            )

                            original_df.to_excel(
                                writer,
                                sheet_name=output_sheet_name,
                                index=False
                            )

            output.seek(0)

            st.success("Columns reordered successfully")

            # ----------------------
            # Summary
            # ----------------------
            st.subheader("Reordering Summary")

            summary_df = pd.DataFrame(summary_rows)
            st.dataframe(summary_df)

            # ----------------------
            # Preview reordered sheets
            # ----------------------
            st.subheader("Preview of Reordered Sheets")

            for sheet_name, reordered_df in reordered_results.items():
                with st.expander(f"Preview: {sheet_name}", expanded=False):
                    st.dataframe(safe_display_df(reordered_df))

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

        except Exception as e:
            st.error(f"Error while reordering columns: {e}")
