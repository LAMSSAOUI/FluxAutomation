import streamlit as st
import pandas as pd
import io
import re

st.title("📊 Fluctuation Analysis Dashboard")

# Helper function to clean column headers
def clean_headers(df):
    df.columns = df.columns.str.strip()
    return df


def rename_quantity_week_columns(df):
    new_cols = []
    for col in df.columns:
        # Check if column contains 'Quantity' and a week/year pattern like 31/2025 or 31/2026
        if 'Quantity' in col:
            match = re.search(r'(\d{1,2})/\d{4}', col)  # Extract week number before slash and 4-digit year
            if match:
                week_num = match.group(1)
                new_col = f"wk {week_num}"
                new_cols.append(new_col)
            else:
                new_cols.append(col)
        else:
            new_cols.append(col)
    df.columns = new_cols
    return df

def clean_material_id(x):
    try:
        # Convert to float first (to handle '123.0')
        f = float(x)
        # Convert to int if no decimal part, then to string
        if f.is_integer():
            return str(int(f))
        else:
            return str(f)
    except:
        # If conversion fails (e.g., already string), just strip and return
        return str(x).strip()

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


        ordered_cols = ['Sales document', 'Name sold-to party', 'Material', 'Project']

        # # Then add other columns from speedi_df that are NOT in the above list
        other_cols = [col for col in speedi_df.columns if col not in ordered_cols]

        # # Final column order
        final_cols = ordered_cols + other_cols

        # Reorder dataframe columns
        speedi_df_organized = speedi_df_organized[final_cols]

        speedi_df_organized = rename_quantity_week_columns(speedi_df_organized)

        st.markdown("### ✅ SPEEDI Data Organized (Matching Last Week on Top)")
        st.dataframe(speedi_df_organized.head())

        # Get the delevery and prepare the data 


delivery_file = st.file_uploader("Upload Delivery Extraction Excel", type=["xlsx"])
if delivery_file:
    xls = pd.ExcelFile(delivery_file)
    selected_sheet = st.selectbox("Select sheet from Delivery file", xls.sheet_names, key="delivery_sheet")
    delivery_df = clean_headers(pd.read_excel(xls, sheet_name=selected_sheet))
    st.markdown("### 🚚 Delivery Data Preview")
    st.dataframe(delivery_df.head())

    columns_to_drop = [
        'Material description', 'Batch', 'Plant', 'Storage location',
        'Movement type', 'Movement Type Text', 'Material Document', 'Material Doc.Item', 'Special Stock', 'Unit of Entry', 'Amt.in Loc.Cur.',
        'Posting Date', 'Document Date', 'Cost Center', 'Order', 'Purchase order', 'Sales order', 'Customer', 'Supplier', 'Reference', 'User Name', 'Entry Date', 'Time of Entry'
    ]

    delivery_df.drop(
        columns=[col for col in delivery_df.columns if col in columns_to_drop],
        inplace=True
    )
    delivery_df['Qty in unit of entry'] = delivery_df['Qty in unit of entry'].abs()
    st.write(f"Number of rows before : {len(delivery_df)}")
    delivery_df = delivery_df.groupby('Material', as_index=False).agg({
        'Qty in unit of entry': 'sum',
    })
    st.write(f"Number of rows after : {len(delivery_df)}")
    st.markdown("### ⚙️ delivery_df Data Prepared")
    st.dataframe(delivery_df.head())
    if delivery_df is not None and speedi_df_organized is not None:
        # Clean material IDs on both dataframes (your existing helper)
        speedi_df_organized['Material'] = speedi_df_organized['Material'].apply(clean_material_id)
        delivery_df['Material'] = delivery_df['Material'].apply(clean_material_id)

        # Define ordered columns as before
        ordered_cols = ['Sales document', 'Name sold-to party', 'Material', 'Project']

        # Start Flux_Calc_df by copying those columns from speedi_df_organized (preserves order)
        Flux_Calc_df = speedi_df_organized[ordered_cols].copy()

        # Merge delivery quantities onto this DataFrame by 'Material'
        Flux_Calc_df = Flux_Calc_df.merge(
            delivery_df[['Material', 'Qty in unit of entry']],
            on='Material',
            how='left'
        )

        # Fill missing delivery quantities with 0
        Flux_Calc_df['Qty in unit of entry'] = Flux_Calc_df['Qty in unit of entry'].fillna(0)

        # Final column order: ordered_cols + delivery qty column
        final_cols = ordered_cols + ['Qty in unit of entry']
        Flux_Calc_df = Flux_Calc_df[final_cols]

        st.markdown("### 📦 Delivery Quantities in SPEEDI Order with Metadata")
        st.dataframe(Flux_Calc_df)


        week_columns = [col for col in speedi_df_organized.columns if col.startswith('wk ')]

        # 2. Sort week columns numerically by their week number
        # week_columns = sorted(week_columns, key=lambda x: int(x.split(' ')[1]))

        # 3. Show only these weeks in the selectbox — no repeats, no extras
        selected_week = st.selectbox("Select starting week", week_columns)

        # Now selected_week will always be one of the actual columns from your data.

        # 4. Use selected_week to filter or calculate differences:
        #    For example, select all weeks >= selected_week
        # start_index = week_columns.index(selected_week)
        # selected_weeks_range = week_columns[start_index:]
        # # Step 2: Select the week from Streamlit dropdown
        # # selected_week = st.selectbox("Select a starting week", week_columns)

        # # Step 3: Find index of selected week in the sorted list
        # # start_idx = week_columns.index(selected_week)

        # # Step 4: Define the weeks range from selected week to the end
        # # selected_weeks_range = week_columns[start_idx:]

        # # Step 5: Prepare last_week_df and speedi_df_organized for the operation
        # # Make sure 'Material' columns are string and stripped
        # # Make sure 'Material' columns are string and stripped (to align keys correctly)
        # last_week_df['Material'] = last_week_df['Material'].astype(str).str.strip()
        # speedi_df_organized['Material'] = speedi_df_organized['Material'].astype(str).str.strip()

        # # selected_weeks_range is a list of columns starting from selected_week, e.g. ['wk31', 'wk32', 'wk33', ...]
        # # Make sure you have it sorted and filtered as needed before this step

        # # Set index on 'Material' for alignment and slicing only selected weeks
        # last_week_weeks = last_week_df.set_index('Material')[selected_weeks_range]
        # speedi_weeks = speedi_df_organized.set_index('Material')[selected_weeks_range]

        # # Calculate difference for every material and week
        # # This subtracts speedi data from last_week data
        # diff_weeks = last_week_weeks.subtract(speedi_weeks, fill_value=0)

        # Find the starting index of the selected week
        # Step 3: Weeks to process from selected week onward
        start_idx = week_columns.index(selected_week)
        weeks_to_process = week_columns[start_idx:]

        # st.markdown("### weeks_to_process")
        # st.dataframe(weeks_to_process)
        # Ensure 'Material' columns are clean and set as index for both DataFrames
        last_week_df['Material'] = last_week_df['Material'].apply(clean_material_id)
        speedi_df_organized['Material'] = speedi_df_organized['Material'].apply(clean_material_id)

        last_week_indexed = last_week_df.set_index('Material')
        speedi_indexed = speedi_df_organized.set_index('Material')

        # Only keep the weeks_to_process columns (plus 'Material' as index)
        last_week_weeks = last_week_indexed[weeks_to_process]
        speedi_weeks = speedi_indexed[weeks_to_process]

        # Align indexes (materials) for subtraction
        last_week_weeks, speedi_weeks = last_week_weeks.align(speedi_weeks, join='outer', fill_value=0)

        # Calculate the difference: (speedi - last_week) for each material and week
        diff_weeks = speedi_weeks.subtract(last_week_weeks, fill_value=0)

        # Reset index to bring 'Material' back as a column
        diff_weeks = diff_weeks.reset_index()

        Flux_Calc_df = Flux_Calc_df.merge(
            diff_weeks,
            on='Material',
            how='left'
        )

        # Optional: Fill NaN in new week columns with 0
        Flux_Calc_df[weeks_to_process] = Flux_Calc_df[weeks_to_process].fillna(0)

        st.markdown("### 📊 Flux_Calc_df with Weekly Differences")
        st.dataframe(Flux_Calc_df)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            last_week_df.to_excel(writer, index=False, sheet_name='LastWeek')
            speedi_df_organized.to_excel(writer, index=False, sheet_name='CurrentWeek')
            Flux_Calc_df.to_excel(writer, index=False, sheet_name='Fluctuation (pcs)')
        output.seek(0)

        st.download_button(
            label="📥 Download Fluctuation Workbook",
            data=output,
            file_name="Fluctuation_Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )




























        # # Prepare 'Material' columns
        # last_week_df['Material'] = last_week_df['Material'].astype(str)
        # speedi_df_organized['Material'] = speedi_df_organized['Material'].astype(str)

        # # Set index on Material for easier access
        # last_week_indexed = last_week_df.set_index('Material')
        # print('last week index',last_week_indexed)
        # speedi_indexed = speedi_df_organized.set_index('Material')
        # print('new week index ',speedi_indexed)

        # # Create empty DataFrame to collect differences
        # diff_data = pd.DataFrame()
        # diff_data['Material'] = last_week_df['Material'].unique()  # all materials

        # # for wk in weeks_to_process:
        # #     # Calculate difference series: last_week - speedi
        # #     diff_series = last_week_indexed[wk].subtract(speedi_indexed[wk], fill_value=0)
        # #     diff_series.name = wk  # name the series with the week
            
        # #     # Reset index to get 'Material' back and merge or join
        # #     diff_df = diff_series.reset_index()
            
        # #     if diff_data.empty or 'Material' not in diff_data:
        # #         diff_data = diff_df
        # #     else:
        # #         # Merge the new week difference column into diff_data on 'Material'
        # #         diff_data = diff_data.merge(diff_df, on='Material', how='outer')





        # diff_dict = {'Material': last_week_df['Material'].unique()}

        # for wk in weeks_to_process:
        #     diff_series = last_week_indexed[wk].subtract(speedi_indexed[wk], fill_value=0)
        #     # diff_dict[wk] = diff_series.reindex(diff_dict['Material']).fillna(0).values  

        # diff_data = pd.DataFrame(diff_series)
        
        # st.markdown(f"### diff_data")
        # st.dataframe(diff_data)
        # # # Now merge this diff_data with Flux_Calc_df on 'Material'
        # # Flux_Calc_df = Flux_Calc_df.merge(diff_data, on='Material', how='left')

        # # # Fill any NaN with 0
        # # Flux_Calc_df[weeks_to_process] = Flux_Calc_df[weeks_to_process].fillna(0)

        # # st.markdown(f"### Differences per week from {selected_week} onwards (last_week_df - speedi_df_organized)")
        # # st.dataframe(Flux_Calc_df)


