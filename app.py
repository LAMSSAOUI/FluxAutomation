import streamlit as st
import pandas as pd
import io

st.title("📊 Fluctuation Analysis Dashboard")

# Helper function to clean column headers
def clean_headers(df):
    df.columns = df.columns.str.strip()
    return df

# --- Upload and select two sheets from Fluctuation Report
st.subheader("📁 Upload Fluctuation Report Workbook (with two sheets)")
fluct_file = st.file_uploader("Upload Fluctuation Report Excel", type=["xlsx"])

if fluct_file:
    fluct_xls = pd.ExcelFile(fluct_file)
    st.write("Sheets found:", fluct_xls.sheet_names)

    # Select 2 sheets
    last_week_sheet_name = st.selectbox("Select sheet for last week's data", fluct_xls.sheet_names, key="last_week")
    fluct_calc_sheet_name = st.selectbox("Select sheet for fluctuation calculations", fluct_xls.sheet_names, key="fluct_calc")

    # Read both sheets
    last_week_df = clean_headers(pd.read_excel(fluct_xls, sheet_name=last_week_sheet_name))
    fluct_calc_df = clean_headers(pd.read_excel(fluct_xls, sheet_name=fluct_calc_sheet_name))

    st.markdown("### 🧾 Last Week Data Preview")
    st.dataframe(last_week_df.head())

    st.markdown("### 🔁 Fluctuation Calculation Sheet Preview")
    st.dataframe(fluct_calc_df.head())
else:
    last_week_df = None
    fluct_calc_df = None

# --- Upload SPEEDI and Delivery Files
st.subheader("📁 Upload SPEEDI and Delivery Files")
speedi_df = None
delivery_df = None

speedi_file = st.file_uploader("Upload SPEEDI Extraction Excel", type=["xlsx"])
if speedi_file:
    xls = pd.ExcelFile(speedi_file)
    selected_sheet = st.selectbox("Select sheet from SPEEDI file", xls.sheet_names, key="speedi_sheet")
    speedi_df = clean_headers(pd.read_excel(xls, sheet_name=selected_sheet))
    st.markdown("### ⚙️ SPEEDI Data Preview")
    st.dataframe(speedi_df.head())
    st.markdown("### ⚙️ SPEEDI Data Preparation")
    columns_to_drop = [
        'Show demands', 'Sales document', 'Item (SD)', 'sales document type',
        'Material type', 'Customer Material', 'Sold-To Party', 'Net price', 'Currency Key', 'Name sold-to party'
    ]

    speedi_df.drop(
        columns=[col for col in speedi_df.columns if col in columns_to_drop or 'Sales' in col],
        inplace=True
    )
    st.markdown("### ⚙️ SPEEDI Data Prepared")
    st.dataframe(speedi_df.head())

    if last_week_df is not None and 'Material' in last_week_df.columns and 'Material' in speedi_df.columns:
        # Strip spaces and ensure same data type
        last_week_df['Material'] = last_week_df['Material'].astype(str).str.strip()
        speedi_df['Material'] = speedi_df['Material'].astype(str).str.strip()

        # # Filter speedi_df to only include materials found in last_week_df (optional safety)
        speedi_df = speedi_df[speedi_df['Material'].isin(last_week_df['Material'])]

        # Set 'Material' as index to preserve order from last_week_df
        last_week_df['Material'] = last_week_df['Material'].astype(str).str.strip()
        speedi_df['Material'] = speedi_df['Material'].astype(str).str.strip()

        # Split materials into two groups: common and new
        materials_in_both = last_week_df['Material'][last_week_df['Material'].isin(speedi_df['Material'])]
        materials_only_in_speedi = speedi_df[~speedi_df['Material'].isin(materials_in_both)]

        # Reorder speedi_df to match last_week_df first, then add unmatched materials
        speedi_matched = speedi_df.set_index('Material').loc[materials_in_both].reset_index()
        speedi_df_reordered = pd.concat([speedi_matched, materials_only_in_speedi], ignore_index=True)

        # Preview
        st.markdown("### ✅ SPEEDI Data Sorted (Matching Last Week on Top)")
        st.dataframe(speedi_df_reordered.head())


        columns_to_copy = ['Sales document', 'Name sold-to party', 'Project', 'Material']

        # Filter last_week_df to keep only needed columns
        copy_from_last = last_week_df[columns_to_copy]
        
    

        # # Merge into speedi_df based on 'Material'
        speedi_df_organized = speedi_df_reordered.merge(copy_from_last, on='Material', how='left')

        # st.markdown("### ✅ speedi_df_organized")
        # st.dataframe(speedi_df_organized.head())

        # # Now reorder columns so 'Project' comes right after 'Material'
        # # Start with these 4 columns in the desired order:
        ordered_cols = ['Sales document', 'Name sold-to party', 'Material', 'Project']

        # # Then add other columns from speedi_df that are NOT in the above list
        other_cols = [col for col in speedi_df.columns if col not in ordered_cols]

        # # Final column order
        final_cols = ordered_cols + other_cols

        # Reorder dataframe columns
        speedi_df_organized = speedi_df_organized[final_cols]

        st.markdown("### ✅ SPEEDI Data Organized (Matching Last Week on Top)")
        st.dataframe(speedi_df_organized.head())
        # add the first colunms project 

delivery_file = st.file_uploader("Upload Delivery Extraction Excel", type=["xlsx"])
if delivery_file:
    xls = pd.ExcelFile(delivery_file)
    selected_sheet = st.selectbox("Select sheet from Delivery file", xls.sheet_names, key="delivery_sheet")
    delivery_df = clean_headers(pd.read_excel(xls, sheet_name=selected_sheet))
    st.markdown("### 🚚 Delivery Data Preview")
    st.dataframe(delivery_df.head())

