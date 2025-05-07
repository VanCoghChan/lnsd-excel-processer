# excel_processor/views.py
import pandas as pd
from django.http import JsonResponse
from django.shortcuts import render
from rest_framework.decorators import api_view

TARGET_COLUMNS = ['组织', '资源集', '地域']
EXCLUDE_SHEETS = ['项目']

def index(request):
    """显示上传页面"""
    return render(request, 'excel_processor/index.html')


@api_view(['POST'])
def upload_excel(request):
    """处理文件上传"""
    if request.FILES.get('excel_file') is not None:
        excel_file = request.FILES['excel_file']
        try:
            # 使用pandas读取文件
            sheet_names = pd.ExcelFile(excel_file).sheet_names

            # 存储处理结果
            results = {}

            # 遍历所有的 sheet
            for sheet_name in sheet_names:
                if sheet_name not in EXCLUDE_SHEETS:
                    # 读取每个sheet的数据
                    sheet_df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    sheet_df_columns = sheet_df.columns.tolist()

                    # 打印sheet的名称和数据
                    print(f"Sheet: {sheet_name}, columns: {sheet_df_columns}")

                    # 检查当前sheet是否包含目标列
                    missing_columns = [col for col in TARGET_COLUMNS if col not in sheet_df_columns]

                    if missing_columns:
                        # 如果缺少目标列，立即返回错误信息
                        error_message = f"工作表 '{sheet_name}' 缺少必要的列: {', '.join(missing_columns)}"
                        print(error_message)
                        return JsonResponse({
                            'success': False,
                            'message': error_message,
                            'data': {
                                'sheet_name': sheet_name,
                                'missing_columns': missing_columns,
                                'available_columns': sheet_df_columns
                            }
                        })

                    try:
                        # 根据目标列进行分组并统计每组的行数
                        grouped = sheet_df.groupby(TARGET_COLUMNS).size().reset_index(name='count')

                        # 将分组结果转换为字典形式，方便JSON序列化
                        group_data = grouped.to_dict('records')

                        # 存储处理结果
                        results[sheet_name] = {
                            'status': 'success',
                            'message': '处理成功',
                            'data': {
                                'group_count': len(group_data),
                                'total_rows': len(sheet_df),
                                'groups': group_data
                            }
                        }

                        print(f"Sheet '{sheet_name}' 处理成功，共 {len(group_data)} 个分组")

                    except Exception as e:
                        # 处理分组过程中可能出现的错误，立即返回错误信息
                        error_message = f"工作表 '{sheet_name}' 分组统计错误: {str(e)}"
                        print(error_message)
                        return JsonResponse({
                            'success': False,
                            'message': error_message,
                            'data': {
                                'sheet_name': sheet_name
                            }
                        })

            # 返回所有sheet的处理结果
            return JsonResponse({
                'success': True,
                'data': {
                    'sheet_count': len(sheet_names),
                    'sheets': results,
                    'column_names': sheet_df.columns.tolist() if 'sheet_df' in locals() else []
                }
            })

        except Exception as e:
            error_message = f"处理Excel文件时出错: {str(e)}"
            print(error_message)
            return JsonResponse({
                'success': False,
                'message': error_message
            })
    else:
        return JsonResponse({
            'success': False,
            'message': "未提供Excel文件"
        })