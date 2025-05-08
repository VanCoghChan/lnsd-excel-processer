import pandas as pd
import os
import uuid
from django.conf import settings

TARGET_COLUMNS = ['组织', '资源集', '地域']
EXCLUDE_SHEETS = ['项目']

def process_excel_file(excel_file):
    """处理上传的Excel文件，进行分组统计并保存结果。"""
    try:
        sheet_names = pd.ExcelFile(excel_file).sheet_names
        all_grouped_results = []
        processed_sheets_info = {}

        for sheet_name in sheet_names:
            if sheet_name not in EXCLUDE_SHEETS:
                sheet_df = pd.read_excel(excel_file, sheet_name=sheet_name)
                sheet_df_columns = sheet_df.columns.tolist()

                print(f"Sheet: {sheet_name}, columns: {sheet_df_columns}")

                missing_columns = [col for col in TARGET_COLUMNS if col not in sheet_df_columns]
                if missing_columns:
                    error_message = f"工作表 '{sheet_name}' 缺少必要的列: {', '.join(missing_columns)}"
                    print(error_message)
                    return {
                        'success': False,
                        'message': error_message,
                        'data': {
                            'sheet_name': sheet_name,
                            'missing_columns': missing_columns,
                            'available_columns': sheet_df_columns
                        }
                    }
                
                try:
                    grouped = sheet_df.groupby(TARGET_COLUMNS).size().reset_index(name=sheet_name)
                    all_grouped_results.append(grouped)
                    processed_sheets_info[sheet_name] = {
                        'group_count': len(grouped),
                        'total_rows': len(sheet_df)
                    }
                    print(f"Sheet '{sheet_name}' 处理成功，共 {len(grouped)} 个分组")
                except Exception as e:
                    error_message = f"工作表 '{sheet_name}' 分组统计错误: {str(e)}"
                    print(error_message)
                    return {
                        'success': False,
                        'message': error_message,
                        'data': {'sheet_name': sheet_name}
                    }

        if not all_grouped_results:
            return {
                'success': False,
                'message': '没有找到可供处理的工作表。',
                'data': {}
            }

        combined_df = all_grouped_results[0]
        for i in range(1, len(all_grouped_results)):
            combined_df = pd.merge(
                combined_df,
                all_grouped_results[i],
                on=TARGET_COLUMNS,
                how='outer'
            )
        
        filename = f"result_{uuid.uuid4().hex[:8]}.xlsx"
        # 确保media目录存在
        os.makedirs(settings.MEDIA_ROOT, exist_ok=True)
        file_path = os.path.join(settings.MEDIA_ROOT, filename)
        combined_df.to_excel(file_path, index=False)

        return {
            'success': True,
            'data': {
                'rows': len(combined_df),
                'columns': len(combined_df.columns),
                'column_names': combined_df.columns.tolist(),
                'download_filename': filename,
                'processed_sheets_info': processed_sheets_info
            }
        }

    except Exception as e:
        error_message = f"处理Excel文件时出错: {str(e)}"
        print(error_message)
        return {
            'success': False,
            'message': error_message
        } 