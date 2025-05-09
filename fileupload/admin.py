from django.contrib import admin
from django.http import HttpResponse
from .models import Document
import os
from .iif_converter import convert_excel_to_iif 

def download_iif(request, queryset):
    if queryset.count() == 1:
        document = queryset.first()
        excel_file_path = document.file.path
        
        output_iif_path = os.path.join("media", f"{document.title}_timesheet_data.iif")

        try:
            iif_content = convert_excel_to_iif(excel_file_path, output_iif_path)
            
            if "Error" in iif_content:
                return HttpResponse(iif_content)  
            
            with open(output_iif_path, 'r') as iif_file:
                response = HttpResponse(iif_file.read(), content_type="text/plain")
                response['Content-Disposition'] = f'attachment; filename={os.path.basename(output_iif_path)}'
                return response

        except Exception as e:
            return HttpResponse(f"An error occurred: {str(e)}")

    else:
        return HttpResponse("Please select one document to download the IIF file.")



@admin.register(Document)
class DocumentAdmin(admin.ModelAdmin):
    list_display = ('title', 'file')  
    search_fields = ('title', 'file') 
    actions = ['download_iif'] 

    def download_iif(self, request, queryset):
        return download_iif(request, queryset)

