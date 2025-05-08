# excel_processor/views.py
from django.http import JsonResponse, FileResponse
from django.shortcuts import render
from rest_framework.decorators import api_view
import os
from django.conf import settings
from .services import process_excel_file # 导入服务

def index(request):
    """显示上传页面"""
    return render(request, 'excel_processor/index.html')

@api_view(['GET'])
def download_result(request, filename):
    """下载处理结果"""
    file_path = os.path.join(settings.MEDIA_ROOT, filename)
    if os.path.exists(file_path):
        response = FileResponse(open(file_path, 'rb'))
        response['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response
    return JsonResponse({'success': False, 'message': '文件不存在'})

@api_view(['POST'])
def upload_excel(request):
    """处理文件上传，调用服务层进行Excel处理"""
    if request.FILES.get('excel_file') is not None:
        excel_file = request.FILES['excel_file']
        
        # 调用服务层处理文件
        result = process_excel_file(excel_file)
        
        return JsonResponse(result)
    else:
        return JsonResponse({
            'success': False,
            'message': "未提供Excel文件"
        })