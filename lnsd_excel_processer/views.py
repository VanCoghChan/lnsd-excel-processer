# excel_processor/views.py
import pandas as pd
from django.shortcuts import render
from django.http import JsonResponse
from rest_framework.decorators import api_view
from .service.excel_service import sheet_process

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
            sheet_names = pd.ExcelFile(excel_file).sheet_names
            # 存储所有工作表的分组结果
            all_grouped_results = []
            # 遍历所有的 sheet
            for sheet_name in sheet_names:
                if sheet_name not in EXCLUDE_SHEETS:
                    # 读取每个sheet的数据
                    sheet_df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    sheet_df_columns = sheet_df.columns.tolist()
                    # 打印sheet的名称和数据
                    print(f"Sheet: {sheet_name}, columns: {sheet_df_columns}")
                    # TODO 处理每一个 sheet
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
                                'sheet': sheet_name,
                                '缺少的列': missing_columns
                            }
                        })
                    # 根据目标列进行分组并统计每组的行数
                    grouped = sheet_df.groupby(TARGET_COLUMNS).size().reset_index(name=sheet_name)
                    all_grouped_results.append(grouped)

            # 将所有分组结果纵向拼接
            if all_grouped_results:
                combined_df = all_grouped_results[0]
                for i in range(1, len(all_grouped_results)):
                    combined_df = pd.merge(
                        combined_df,
                        all_grouped_results[i],
                        on=TARGET_COLUMNS,
                        how='outer'
                    )
                combined_df.to_csv("res.csv", encoding="utf-8")
            # 返回处理结果
            return JsonResponse({
                'success': True,
                'message': '文件上传并处理成功',
                'data': {
                    'rows': 1,
                    'columns':1 ,
                    'column_names': 1
                }
            })
        except Exception as e:
            return JsonResponse({'success': False, 'message': f'处理文件时出错: {str(e)}'})

    return JsonResponse({'success': False, 'message': '请选择一个Excel文件上传'})