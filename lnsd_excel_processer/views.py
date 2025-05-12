# excel_processor/views.py
from django.http import JsonResponse, FileResponse
from django.shortcuts import render
from rest_framework.decorators import api_view
import os
import json # For parsing JSON from request body
from django.conf import settings
from .services import get_excel_metadata, perform_final_analysis

def index(request):
    """显示上传页面"""
    return render(request, 'excel_processor/index.html')

@api_view(['POST'])
def upload_and_get_metadata_view(request):
    """处理文件上传，提取元数据。"""
    if request.FILES.get('excel_file') is not None:
        excel_file = request.FILES['excel_file']
        result = get_excel_metadata(excel_file)
        return JsonResponse(result)
    else:
        return JsonResponse({'success': False, 'message': "未提供Excel文件"}, status=400)

@api_view(['POST'])
def trigger_final_analysis_view(request):
    """触发最终的Excel分析和处理。"""
    try:
        data = json.loads(request.body)
        temp_file_id = data.get('temp_file_id')
        selected_sheets = data.get('selected_sheets') # List of sheet names
        additional_stat_columns = data.get('additional_stat_columns') # List of column names for stats

        if not temp_file_id or not isinstance(selected_sheets, list) or not isinstance(additional_stat_columns, list):
            return JsonResponse({
                'success': False, 
                'message': '请求参数无效或缺失 (temp_file_id, selected_sheets, additional_stat_columns)。'
            }, status=400)
        
        if not selected_sheets:
            return JsonResponse({
                'success': False, 
                'message': '必须至少选择一个工作表进行处理。'
            }, status=400)

        result = perform_final_analysis(temp_file_id, selected_sheets, additional_stat_columns)
        return JsonResponse(result)
    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'message': '无效的JSON请求体。'}, status=400)
    except Exception as e:
        # Catch any other unexpected errors during view processing
        print(f"Error in trigger_final_analysis_view: {str(e)}")
        return JsonResponse({'success': False, 'message': f'处理请求时发生服务器内部错误: {str(e)}'}, status=500)

@api_view(['GET'])
def download_result(request, filename):
    """下载处理结果"""
    file_path = os.path.join(settings.MEDIA_ROOT, filename) # Results are in MEDIA_ROOT directly
    if os.path.exists(file_path):
        try:
            response = FileResponse(open(file_path, 'rb'))
            response['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            response['Content-Disposition'] = f'attachment; filename="{filename}"'
            return response
        except Exception as e:
            print(f"Error serving file {filename}: {e}")
            return JsonResponse({'success': False, 'message': '下载文件时出错。'}, status=500)
    else:
        return JsonResponse({'success': False, 'message': '文件不存在或无法访问。'}, status=404)