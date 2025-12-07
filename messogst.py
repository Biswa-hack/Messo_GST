def process_zip_and_combine_data(zip_file_uploader, combo_template_file):
    """Handles zip extraction, file identification, processing, and Excel output."""
    
    sales_file_data = None
    returns_file_data = None
    
    # 1. Unzip and Identify Files (Remains unchanged)
    try:
        with zipfile.ZipFile(io.BytesIO(zip_file_uploader.read())) as z:
            for name in z.namelist():
                if name.endswith('.xlsx') or name.endswith('.xls'):
                    if 'return' in name.lower() or 'rtn' in name.lower():
                        returns_file_data = z.open(name)
                    else:
                        sales_file_data = z.open(name)
            
            if not sales_file_data or not returns_file_data:
                 st.error("❌ Could not identify both 'Sales' and 'Returns' files inside the ZIP.")
                 return None

    except zipfile.BadZipFile:
        st.error("❌ The uploaded file is not a valid ZIP archive.")
        return None
    except Exception as e:
        st.error(f"❌ An unexpected error occurred during zip processing: {e}")
        return None

    # 2. Process Data and Merge (Remains unchanged)
    df_sales = process_file(sales_file_data, 'Sale')
    df_returns = process_file(returns_file_data, 'Return')
    
    merged_dfs = []
    if df_sales is not None: merged_dfs.append(df_sales)
    if df_returns is not None: merged_dfs.append(df_returns)

    if not merged_dfs:
        st.error("❌ No valid sales or returns data was processed.")
        return None
        
    df_merged = pd.concat(merged_dfs, ignore_index=True)
    df_merged = df_merged.dropna(axis=1, how='all')

    # 3. Insert Merged Data into Combo Template
    try:
        wb = load_workbook(combo_template_file)
        if 'raw' not in wb.sheetnames:
            st.error("❌ Error: The 'combo' file must contain a sheet named 'raw'.")
            return None
            
        ws = wb['raw']

        # --- A. CLEAR CELLS B3:I[MAX_ROW] (Preserve Formulas in J-O) ---
        start_row_to_clear = 3
        if ws.max_row >= start_row_to_clear:
            rows_to_clear = list(range(start_row_to_clear, ws.max_row + 1))
            for row_idx in rows_to_clear:
                for col_idx in range(2, 10): # Columns 2 (B) through 9 (I)
                    ws.cell(row=row_idx, column=col_idx).value = None
            st.info(f"Cleared old data from columns B:I starting at row 3.")

        # --- B. Paste Merged Data starting at B3 ---
        for r_idx, row in enumerate(dataframe_to_rows(df_merged, header=False, index=False)):
            for c_idx, value in enumerate(row):
                ws.cell(row=start_row_to_clear + r_idx, column=2 + c_idx, value=value)
        
        new_max_row = start_row_to_clear + len(df_merged) - 1
        st.success(f"Successfully pasted {len(df_merged)} rows (B3 to I{new_max_row}).")
        
        # --- C. HARDCODED FORMULAS (K3:O[LastRow]) ---
        formula_target_start_row = 3
        
        # Hardcoded formulas (using {0} as a placeholder for the row index)
        # Note: E1% has been replaced with E{0}/100 for correct openpyxl syntax.
        RAW_FORMULAS = [
            # K: =IF(J1=$X$22,F1*E1%/2,0)
            '=IF(J{0}=$X$22,F{0}*E{0}/100/2,0)', 
            # L: =IF(J1=$X$22,F1*E1%/2,0)
            '=IF(J{0}=$X$22,F{0}*E{0}/100/2,0)',
            # M: =IF(J1=$X$22,0,F1*E1%)
            '=IF(J{0}=$X$22,0,F{0}*E{0}/100)',
            # N: =K1+L1+M1+F1 
            '=K{0}+L{0}+M{0}+F{0}',
            # O: =(M1+L1+K1)/F1 
            '=(M{0}+L{0}+K{0})/F{0}'
        ]
        
        # Formulas will be pasted starting at Column K (index 11)
        formula_start_col = 11 

        if len(df_merged) > 0:
            for row_idx in range(formula_target_start_row, new_max_row + 1): 
                for col_idx, formula_template in enumerate(RAW_FORMULAS):
                    
                    # 1. Format the formula string with the current row index
                    formatted_formula = formula_template.format(row_idx)

                    # 2. Paste the formula into the cell
                    target_col = formula_start_col + col_idx
                    ws.cell(row=row_idx, column=target_col, value=formatted_formula)
        
            st.info(f"Hardcoded formulas (K:O) applied to {new_max_row - formula_target_start_row + 1} rows.")
            
        st.warning("⚠️ **Pivot Table Refresh:** Please ensure the Pivot Tables are set to 'Refresh data when opening the file' in Excel.")
        
        output = io.BytesIO()
        wb.save(output)
        
        return output.getvalue()
        
    except Exception as e:
        st.error(f"❌ An error occurred during file manipulation: {e}")
        return None
