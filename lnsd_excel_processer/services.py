import pandas as pd
import os
import uuid
from django.conf import settings
import numpy as np # For np.nan

TARGET_COLUMNS = ['组织', '资源集', '地域']
EXCLUDE_SHEETS = ['项目']
TEMP_UPLOAD_DIR = os.path.join(settings.MEDIA_ROOT, 'temp_uploads')

def get_excel_metadata(excel_file):
    """保存上传的Excel文件，提取其元数据（工作表、列名）。"""
    try:
        os.makedirs(TEMP_UPLOAD_DIR, exist_ok=True)
        temp_file_id = f"temp_{uuid.uuid4().hex[:12]}.xlsx"
        temp_file_path = os.path.join(TEMP_UPLOAD_DIR, temp_file_id)
        
        # 保存上传的文件到临时位置
        with open(temp_file_path, 'wb+') as destination:
            for chunk in excel_file.chunks():
                destination.write(chunk)

        xls = pd.ExcelFile(temp_file_path)
        sheet_names = [name for name in xls.sheet_names if name not in EXCLUDE_SHEETS]
        
        sheets_metadata = []
        all_columns_set = set()
        for sheet_name in sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, nrows=0) # Read only headers
                columns = df.columns.tolist()
                sheets_metadata.append({'name': sheet_name, 'columns': columns})
                all_columns_set.update(columns)
            except Exception as e:
                print(f"Error reading columns from sheet {sheet_name}: {e}")
                # Optionally skip sheet or return error for this sheet
                sheets_metadata.append({'name': sheet_name, 'columns': [], 'error': str(e)})

        if not sheets_metadata:
            os.remove(temp_file_path) # Clean up if no valid sheets found
            return {'success': False, 'message': '未在Excel文件中找到有效的工作表。'}

        return {
            'success': True,
            'data': {
                'temp_file_id': temp_file_id,
                'sheets': sheets_metadata,
                'all_available_columns': sorted(list(all_columns_set)) # For selecting stat columns
            }
        }
    except Exception as e:
        error_message = f"提取Excel元数据时出错: {str(e)}"
        print(error_message)
        # Attempt to clean up if a temp file was partially created and an error occurred
        if 'temp_file_path' in locals() and os.path.exists(temp_file_path):
            try:
                os.remove(temp_file_path)
            except Exception as cleanup_error:
                print(f"Error cleaning up temp file {temp_file_path}: {cleanup_error}")
        return {'success': False, 'message': error_message}

def perform_final_analysis(temp_file_id, selected_sheet_names, additional_stat_columns):
    """根据选择的表和字段处理Excel，进行分组统计和额外统计。"""
    temp_file_path = os.path.join(TEMP_UPLOAD_DIR, temp_file_id)
    if not os.path.exists(temp_file_path):
        return {'success': False, 'message': '临时文件不存在或已过期。'}

    try:
        xls = pd.ExcelFile(temp_file_path)
        
        all_sheet_count_results = []
        # To store aggregated sums for each additional_stat_column across all selected sheets
        # Key: original stat column name, Value: DataFrame of its sum grouped by TARGET_COLUMNS
        aggregated_stat_sums = {}

        for sheet_name in selected_sheet_names:
            if sheet_name not in xls.sheet_names or sheet_name in EXCLUDE_SHEETS:
                continue
            
            df = pd.read_excel(xls, sheet_name=sheet_name)
            sheet_df_columns = df.columns.tolist()

            # 1. Process sheet counts (original logic)
            # Column name is now just the sheet_name
            sheet_count_df = df.groupby(TARGET_COLUMNS).size().reset_index(name=sheet_name)
            all_sheet_count_results.append(sheet_count_df)

            # 2. Process additional stat columns for this sheet
            if additional_stat_columns:
                valid_numeric_stat_cols_for_sheet = []
                for col_name in additional_stat_columns:
                    if col_name in sheet_df_columns and pd.api.types.is_numeric_dtype(df[col_name]):
                        valid_numeric_stat_cols_for_sheet.append(col_name)
                
                if valid_numeric_stat_cols_for_sheet:
                    # Calculate sums for valid columns for the current sheet
                    current_sheet_sums_df = df.groupby(TARGET_COLUMNS, as_index=False)[valid_numeric_stat_cols_for_sheet].sum()
                    
                    # Aggregate these sums into the global aggregated_stat_sums DataFrames
                    for col_name in valid_numeric_stat_cols_for_sheet:
                        # Extract the series for the current stat col and rename it to its original name
                        stat_series_df = current_sheet_sums_df[TARGET_COLUMNS + [col_name]].copy()
                        # stat_series_df now holds TARGET_COLUMNS and the sum for col_name from current sheet

                        if col_name not in aggregated_stat_sums:
                            aggregated_stat_sums[col_name] = stat_series_df
                        else:
                            # Merge with existing aggregated sums for this col_name, summing the values
                            # Suffixes are added to temporarily resolve the sum column name conflict during merge
                            merged = pd.merge(aggregated_stat_sums[col_name], stat_series_df, on=TARGET_COLUMNS, how='outer', suffixes=['_agg', '_new'])
                            # Sum up the old aggregated sum and the new sum from the current sheet
                            merged[col_name] = merged[f'{col_name}_agg'].fillna(0) + merged[f'{col_name}_new'].fillna(0)
                            # Keep only TARGET_COLUMNS and the final summed col_name
                            aggregated_stat_sums[col_name] = merged[TARGET_COLUMNS + [col_name]]
        
        if not all_sheet_count_results and not aggregated_stat_sums:
            if os.path.exists(temp_file_path): os.remove(temp_file_path)
            return {'success': False, 'message': '没有选择任何有效的工作表或没有可处理的数据。'}

        # Merge all sheet count results first
        final_df = None
        if all_sheet_count_results:
            final_df = all_sheet_count_results[0]
            for i in range(1, len(all_sheet_count_results)):
                final_df = pd.merge(final_df, all_sheet_count_results[i], on=TARGET_COLUMNS, how='outer')
            # For count columns, NaN means the group didn't exist; leave as NaN (or np.nan for consistency)
            # No fillna(0) here for count columns

        # Merge aggregated stat sums
        for stat_col_name, stat_df in aggregated_stat_sums.items():
            if final_df is None: # If there were no count results (e.g., all sheets excluded or empty)
                final_df = stat_df
            else:
                final_df = pd.merge(final_df, stat_df, on=TARGET_COLUMNS, how='outer')
            # For summed stat columns, fill NaN with 0 as requested
            final_df[stat_col_name] = final_df[stat_col_name].fillna(0)

        if final_df is None or final_df.empty:
            if os.path.exists(temp_file_path): os.remove(temp_file_path)
            return {'success': False, 'message': '最终处理结果为空。'}
        
        # Ensure TARGET_COLUMNS are first, followed by sheet count columns (sorted by name), then stat columns (sorted by name)
        if all_sheet_count_results:
            count_col_names = sorted([res.columns[-1] for res in all_sheet_count_results if res.columns[-1] not in TARGET_COLUMNS])
        else:
            count_col_names = []    
        
        stat_col_names_sorted = sorted(list(aggregated_stat_sums.keys()))
        
        ordered_columns = TARGET_COLUMNS + count_col_names + stat_col_names_sorted
        # Filter out any columns that might not be in final_df if some parts were empty
        ordered_columns = [col for col in ordered_columns if col in final_df.columns]
        final_df = final_df[ordered_columns]

        output_filename = f"analysis_result_{uuid.uuid4().hex[:8]}.xlsx"
        output_file_path = os.path.join(settings.MEDIA_ROOT, output_filename)
        os.makedirs(settings.MEDIA_ROOT, exist_ok=True)
        final_df.to_excel(output_file_path, index=False)
        
        if os.path.exists(temp_file_path): os.remove(temp_file_path)

        return {
            'success': True,
            'data': {
                'rows': len(final_df),
                'columns': len(final_df.columns),
                'column_names': final_df.columns.tolist(),
                'download_filename': output_filename
            }
        }
    except Exception as e:
        error_message = f"执行最终分析时出错: {str(e)}"
        print(error_message)
        if os.path.exists(temp_file_path):
            try: os.remove(temp_file_path)
            except Exception as cleanup_error: print(f"Error cleaning up temp file {temp_file_path} during error handling: {cleanup_error}")
        return {'success': False, 'message': error_message} 