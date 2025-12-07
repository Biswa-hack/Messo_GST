def process_zip_and_combine_data(zip_file_uploader, combo_template_file):
    # ... (Sections 1 and 2 remain unchanged) ...

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
        
        # --- C. COPY FORMULAS DOWN (K1:O1 to K3:O[LastRow]) ---
        master_formula_row = 1 # 🟢 UPDATED: Source row is now 1
        formula_start_col = 11 # K 
        formula_end_col = 15   # O
        
        if len(df_merged) > 0:
            for row_idx in range(start_row_to_clear, new_max_row + 1): # Start copying from Row 3 down
                for col_idx in range(formula_start_col, formula_end_col + 1):
                    
                    # Get the value (which should be the formula string) from the master row (Row 1)
                    source_value = ws.cell(row=master_formula_row, column=col_idx).value
                    
                    if isinstance(source_value, str) and source_value.startswith('='):
                        ws.cell(row=row_idx, column=col_idx).value = source_value
        
            st.info(f"Copied formulas from K1:O1 down to K{new_max_row}:O{new_max_row}.")
            
        st.warning("⚠️ **Pivot Table Refresh:** Please ensure the Pivot Tables are set to 'Refresh data when opening the file' in Excel.")
        
        output = io.BytesIO()
        wb.save(output)
        
        return output.getvalue()
        
    except Exception as e:
        # Keep this for debugging the file manipulation error
        st.error(f"❌ An error occurred during file manipulation: {e}")
        return None
