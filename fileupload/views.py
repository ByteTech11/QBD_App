from rest_framework import viewsets
from rest_framework.response import Response
from rest_framework.parsers import MultiPartParser, FormParser
from .models import Document
from .serializers import DocumentSerializer
from django.http import HttpResponse, FileResponse 
from .iif_converter import convert_excel_to_iif 

class DocumentViewSet(viewsets.ModelViewSet):
    queryset = Document.objects.all()
    serializer_class = DocumentSerializer
    parser_classes = (MultiPartParser, FormParser)  

    def create(self, request, *args, **kwargs):
        file = request.FILES.get('file') 
        
        if file and file.name.endswith('.xlsx'):
            iif_content = convert_excel_to_iif(file)
            
            response = HttpResponse(iif_content, content_type='text/plain')
            response['Content-Disposition'] = f'attachment; filename={file.name.replace(".xlsx", ".iif")}'
            return response
        else:
            return Response({"error": "Invalid file type. Please upload an Excel file."}, status=400)
# DONT TRY IT WITHOUT PERMISSION BY ME AYYAZ
##from django.http import HttpResponse, FileResponse #####
##from .models import Document

## View for downloading a file based on its ID
##def download_file(request, file_id):
#    try:
        # Fetch the document object by its ID
#        document = Document.objects.get(id=file_id)
        
        # Return the file as an attachment
#       response = FileResponse(document.file.open(), content_type='application/octet-stream')
#        response['Content-Disposition'] = f'attachment; filename={document.file.name}'
        
#        return response
#    except Document.DoesNotExist:
##        return HttpResponse("File not found", status=404) ###

from django.http import HttpResponse

def home(request):
    return HttpResponse("Hello Wellcome To The Django_App !")


    


