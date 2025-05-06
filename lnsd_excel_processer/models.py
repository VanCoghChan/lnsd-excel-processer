# excel_processor/models.py
from django.db import models

class ExcelFile(models.Model):
    file = models.FileField(upload_to='excel_files/')
    uploaded_at = models.DateTimeField(auto_now_add=True)
    processed = models.BooleanField(default=False)

    def __str__(self):
        return f"Excel文件 - 上传于 {self.uploaded_at.strftime('%Y-%m-%d %H:%M')}"