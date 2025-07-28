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
    # st.write("Sheets found:", fluct_xls.sheet_names)

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
    # st.markdown("### ⚙️ SPEEDI Data Preview")
    # st.dataframe(speedi_df.head())
    # st.markdown("### ⚙️ SPEEDI Data Preparation")
    columns_to_drop = [
        'Show demands', 'Sales document', 'Item (SD)', 'sales document type',
        'Material type', 'Customer Material', 'Sold-To Party', 'Net price', 'Currency Key', 'Name sold-to party'
    ]

    speedi_df.drop(
        columns=[col for col in speedi_df.columns if col in columns_to_drop or 'Sales' in col],
        inplace=True
    )
    # st.markdown("### ⚙️ SPEEDI Data Prepared")
    # st.dataframe(speedi_df.head())

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
        # st.markdown("### ✅ SPEEDI Data Sorted (Matching Last Week on Top)")
        # st.dataframe(speedi_df_reordered.head())


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
    # st.markdown("### 🚚 Delivery Data Preview")
    # st.dataframe(delivery_df.head())

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
    # st.markdown("### ⚙️ delivery_df Data Prepared")
    # st.dataframe(delivery_df.head())
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

        # st.markdown("### 📦 Delivery Quantities in SPEEDI Order with Metadata")
        # st.dataframe(Flux_Calc_df)


        week_columns = [col for col in speedi_df_organized.columns if col.startswith('wk ')]

        selected_week = st.selectbox("Select starting week", week_columns)

       
        start_idx = week_columns.index(selected_week)
        weeks_to_process = week_columns[start_idx:]

   
        last_week_df['Material'] = last_week_df['Material'].apply(clean_material_id)
        speedi_df_organized['Material'] = speedi_df_organized['Material'].apply(clean_material_id)

        last_week_indexed = last_week_df.set_index('Material')
        speedi_indexed = speedi_df_organized.set_index('Material')

   
        last_week_weeks = last_week_indexed[weeks_to_process]
        speedi_weeks = speedi_indexed[weeks_to_process]


        last_week_weeks, speedi_weeks = last_week_weeks.align(speedi_weeks, join='outer', fill_value=0)
      
        diff_weeks = speedi_weeks.subtract(last_week_weeks, fill_value=0)

        diff_weeks = diff_weeks.reset_index()

        Flux_Calc_df = Flux_Calc_df.merge(
            diff_weeks,
            on='Material',
            how='left'
        )

        # Optional: Fill NaN in new week columns with 0
        Flux_Calc_df[weeks_to_process] = Flux_Calc_df[weeks_to_process].fillna(0)


        
        # Find the previous week column
        selected_week_idx = week_columns.index(selected_week)  
        prev_week = week_columns[selected_week_idx - 1]
 


        current_deficit = speedi_df_organized.set_index('Material')['Deficit quantity'].groupby(level=0).sum()
        last_deficit = last_week_df.set_index('Material')['Deficit quantity'].groupby(level=0).sum()
        last_week_demand = last_week_df.set_index('Material')[prev_week].groupby(level=0).sum()
 


        # # deficit fluctuatuion 
        # current_deficit = speedi_df_organized.set_index('Material')[selected_week].groupby(level=0).sum()
        
        # last_deficit = last_week_df.set_index('Material')[selected_week].groupby(level=0).sum()
        
        # last_week_demand = last_week_df.set_index('Material')[prev_week].groupby(level=0).sum()
   
        Flux_Calc_df['Deficit quantity'] = (
            Flux_Calc_df['Material'].map(current_deficit).fillna(0)
            - Flux_Calc_df['Material'].map(last_deficit).fillna(0)
            + Flux_Calc_df['Qty in unit of entry']
            - Flux_Calc_df['Material'].map(last_week_demand).fillna(0)
        )


        cols = list(Flux_Calc_df.columns)
        qty_idx = cols.index('Qty in unit of entry')
        # Remove 'Deficit quantity' if already present
        cols.remove('Deficit quantity')
        # Insert after 'Qty in unit of entry'
        cols.insert(qty_idx + 1, 'Deficit quantity')
        Flux_Calc_df = Flux_Calc_df[cols]    

     

        fluct_pct_df = Flux_Calc_df[['Sales document', 'Name sold-to party', 'Material', 'Project']].copy()
        for week in weeks_to_process:
            last_week_vals = last_week_indexed[week] if week in last_week_indexed.columns else pd.Series(0, index=last_week_indexed.index)
            curr_week_vals = speedi_indexed[week] if week in speedi_indexed.columns else pd.Series(0, index=speedi_indexed.index)
            last_week_vals, curr_week_vals = last_week_vals.align(curr_week_vals, join='outer', fill_value=0)
            pct = []
            for mat in Flux_Calc_df['Material']:
                lw = last_week_vals.get(mat, 0)
                cw = curr_week_vals.get(mat, 0)
                if isinstance(lw, pd.Series): lw = lw.sum()
                if isinstance(cw, pd.Series): cw = cw.sum()
                if lw == 0:
                    pct.append(1 if cw > 0 else 0)
                else:
                    pct.append((cw - lw) / lw)
            # Format as percentage for Excel (not string)
            fluct_pct_df[week] = pct


        fluct_pct_df = Flux_Calc_df[['Sales document', 'Name sold-to party', 'Material', 'Project']].copy()
        for week in weeks_to_process:
            last_week_vals = last_week_indexed[week] if week in last_week_indexed.columns else pd.Series(0, index=last_week_indexed.index)
            curr_week_vals = speedi_indexed[week] if week in speedi_indexed.columns else pd.Series(0, index=speedi_indexed.index)
            last_week_vals, curr_week_vals = last_week_vals.align(curr_week_vals, join='outer', fill_value=0)
            pct = []
            for mat in Flux_Calc_df['Material']:
                lw = last_week_vals.get(mat, 0)
                cw = curr_week_vals.get(mat, 0)
                if isinstance(lw, pd.Series): lw = lw.sum()
                if isinstance(cw, pd.Series): cw = cw.sum()
                if lw == 0:
                    pct.append(1 if cw > 0 else 0)
                else:
                    pct.append((cw - lw) / lw)
            fluct_pct_df[week] = pct

        # Calculate last_week_demand and last_deficit for each material
        last_week_demand_series = last_week_df.set_index('Material')[prev_week].groupby(level=0).sum()
        last_deficit_series = last_week_df.set_index('Material')['Deficit quantity'].groupby(level=0).sum()
        deficit_quantity_series = Flux_Calc_df.set_index('Material')['Deficit quantity']

        # Compute deficit pourcentage for each material
        deficit_pct = []
        for mat in Flux_Calc_df['Material']:
            lw_demand = last_week_demand_series.get(mat, 0)
            lw_deficit = last_deficit_series.get(mat, 0)
            deficit = deficit_quantity_series.get(mat, 0)
            denominator = lw_demand + lw_deficit
            if denominator == 0:
                deficit_pct.append(0)
            else:
                deficit_pct.append(deficit / denominator)
        fluct_pct_df['Deficit Pourcentage'] = deficit_pct

        # Reorder columns to put 'Deficit Pourcentage' after 'Project'
        cols = list(fluct_pct_df.columns)
        cols.remove('Deficit Pourcentage')
        proj_idx = cols.index('Project')
        cols.insert(proj_idx + 1, 'Deficit Pourcentage')
        fluct_pct_df = fluct_pct_df[cols + [col for col in fluct_pct_df.columns if col not in cols]]


        percent_cols = list(weeks_to_process) + ['Deficit Pourcentage']
        for col in percent_cols:
            if col in fluct_pct_df.columns:
                fluct_pct_df[col] = pd.to_numeric(fluct_pct_df[col], errors='coerce')
                fluct_pct_df[col] = (fluct_pct_df[col] * 100).round(2).astype(str) + '%'

        st.markdown("### 📊 Fluctuation File")
        st.dataframe(Flux_Calc_df)

        output = io.BytesIO()
        # ...existing code for Excel export...
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Write sheets
            last_week_df.to_excel(writer, index=False, sheet_name='LastWeek')
            speedi_df_organized.to_excel(writer, index=False, sheet_name='CurrentWeek')
            Flux_Calc_df.to_excel(writer, index=False, sheet_name='Fluctuation (pcs)')
            fluct_pct_df.to_excel(writer, index=False, sheet_name='Fluctuation (Pourcentage)')

            workbook  = writer.book


            worksheet = writer.sheets['Fluctuation (pcs)']
            worksheet.activate()
            worksheet.set_first_sheet()

            # Header format: gray background, bold
            header_format = workbook.add_format({
                'bg_color': '#B7B7B7',
                'font_color': 'black',
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True
            })
            header_height = 22

            # Week value formats
            yellow_bold = workbook.add_format({'bg_color': '#FFFF00', 'bold': True})
            orange_yellow = workbook.add_format({'bg_color': '#FFC000'})

            # Sheet tab colors (choose your own)
            tab_colors = {
                'LastWeek': '#CBF3F0',         # Gray
                'CurrentWeek': '#FFBF69',      # Light green
                'Fluctuation (pcs)': '#9D69A3', # Light yellow
                'Fluctuation (Pourcentage)': '#E85D75' # Light greenish
            }



            # Apply header style and tab color to all sheets
            for sheet_name, tab_color in tab_colors.items():
                worksheet = writer.sheets[sheet_name]
                worksheet.set_tab_color(tab_color)
                columns = (
                    Flux_Calc_df.columns if sheet_name == 'Fluctuation (pcs)' else
                    fluct_pct_df.columns if sheet_name == 'Fluctuation (Pourcentage)' else
                    last_week_df.columns if sheet_name == 'LastWeek' else
                    speedi_df_organized.columns
                )
                for col_num, value in enumerate(columns):
                    worksheet.write(0, col_num, value, header_format)

            # Conditional formatting for week columns in Fluctuation (pcs)
            fluct_ws = writer.sheets['Fluctuation (pcs)']
            start_row = 1  # 0-based, so 1 is first data row
            end_row = len(Flux_Calc_df)
            for col in weeks_to_process:
                col_idx = Flux_Calc_df.columns.get_loc(col)
                # >0: yellow bold
                fluct_ws.conditional_format(start_row, col_idx, end_row, col_idx, {
                    'type': 'cell',
                    'criteria': '>',
                    'value': 0,
                    'format': yellow_bold
                })
                # <0: light yellow
                fluct_ws.conditional_format(start_row, col_idx, end_row, col_idx, {
                    'type': 'cell',
                    'criteria': '<',
                    'value': 0,
                    'format': orange_yellow
                })


            green_format = workbook.add_format({'bg_color': '#AAAE7F'})  # > 20%
            purple_format = workbook.add_format({'bg_color': '#9000B3'})  # between -20% and 20%
            red_format = workbook.add_format({'bg_color': '#F0544F'})    # < -20%

            fluct_pct_ws = writer.sheets['Fluctuation (Pourcentage)']
            start_row = 1
            end_row = len(fluct_pct_df)

            # Create number formats that include percentage symbol
            green_format = workbook.add_format({
                'bg_color': '#AAAE7F',
                'num_format': '0.00%'
            })
            purple_format = workbook.add_format({
                'bg_color': '#9000B3',
                'num_format': '0.00%'
            })
            red_format = workbook.add_format({
                'bg_color': '#F0544F',
                'num_format': '0.00%'
            })

            percent_cols = list(weeks_to_process) + ['Deficit Pourcentage']
            for col in percent_cols:
                if col in fluct_pct_df.columns:
                    col_idx = fluct_pct_df.columns.get_loc(col)
                    
                    # Convert percentage strings back to numbers for Excel
                    fluct_pct_df[col] = fluct_pct_df[col].str.rstrip('%').astype(float) / 100
                    
                    # Apply conditional formatting using decimal values
                    # > 20%: Green
                    fluct_pct_ws.conditional_format(start_row, col_idx, end_row, col_idx, {
                        'type': 'cell',
                        'criteria': '>',
                        'value': 20,
                        'format': green_format
                    })
                    
                    # Between 0% and 20%: Purple
                    fluct_pct_ws.conditional_format(start_row, col_idx, end_row, col_idx, {
                        'type': 'cell',
                        'criteria': 'between',
                        'minimum': 0,
                        'maximum': 20,
                        'format': purple_format
                    })
                    
                    # Between -20% and 0%: Purple
                    fluct_pct_ws.conditional_format(start_row, col_idx, end_row, col_idx, {
                        'type': 'cell',
                        'criteria': 'between',
                        'minimum': -20,
                        'maximum': 0,
                        'format': purple_format
                    })
                    
                    # < -20%: Red
                    fluct_pct_ws.conditional_format(start_row, col_idx, end_row, col_idx, {
                        'type': 'cell',
                        'criteria': '<',
                        'value': -20,
                        'format': red_format
                    })


        output.seek(0)

        st.download_button(
            label="📥 Download Fluctuation Workbook",
            data=output,
            file_name="Fluctuation_Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )



