from django.contrib import admin
from django.urls import path
from django.shortcuts import render
from django import forms
from django.http import FileResponse, HttpResponseRedirect
from django.contrib import messages
from zipfile import ZipFile
import os
import tempfile
from .models import Document
from .iif_converter import convert_excel_to_iif
from django.forms.widgets import ClearableFileInput


class MultipleFileInput(forms.ClearableFileInput):
    allow_multiple_selected = True


class MultipleFileField(forms.FileField):
    widget = MultipleFileInput

    def clean(self, data, initial=None):
        if not data:
            raise forms.ValidationError("No files were submitted.")
        if not isinstance(data, list):
            data = [data]
        return data


class MultiFileUploadForm(forms.Form):
    title = forms.CharField(required=False)
    files = MultipleFileField()

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['files'].widget.attrs['name'] = 'files'


@admin.register(Document)
class DocumentAdmin(admin.ModelAdmin):
    list_display = ('title', 'file')
    actions = ['download_iif_files']
    change_list_template = "admin/upload_button.html"

    def save_model(self, request, obj, form, change):
        """
        Called when saving a Document in admin (single upload).
        This will generate the .iif file right after the Excel file is saved.
        """
        super().save_model(request, obj, form, change) 

        input_path = obj.file.path
        output_path = os.path.splitext(input_path)[0] + '.iif'

        result = convert_excel_to_iif(input_path, output_path)

        if "Error" in result:
            self.message_user(request, f"IIF file generation failed for {obj.title or obj.file.name}.", level=messages.ERROR)
        else:
            self.message_user(request, f"IIF file generated for {obj.title or obj.file.name}.", level=messages.SUCCESS)

    @admin.action(description='Download IIF files')
    def download_iif_files(self, request, queryset):
        selected_docs = list(queryset)
        if not selected_docs:
            self.message_user(request, "No documents selected.", level=messages.ERROR)
            return HttpResponseRedirect(request.get_full_path())

        iif_files = []
        for doc in selected_docs:
            iif_path = os.path.splitext(doc.file.path)[0] + '.iif'
            if os.path.exists(iif_path):
                iif_files.append(iif_path)
            else:
                self.message_user(request, f"IIF file not found for {doc.title}.", level=messages.WARNING)

        if not iif_files:
            self.message_user(request, "No IIF files found for the selected documents.", level=messages.ERROR)
            return HttpResponseRedirect(request.get_full_path())

        if len(iif_files) == 1:
            iif_file_path = iif_files[0]
            return FileResponse(open(iif_file_path, 'rb'), as_attachment=True, filename=os.path.basename(iif_file_path))
        else:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.zip') as tmp:
                with ZipFile(tmp, 'w') as zipf:
                    for file_path in iif_files:
                        zipf.write(file_path, arcname=os.path.basename(file_path))
                tmp.seek(0)
                return FileResponse(open(tmp.name, 'rb'), as_attachment=True, filename='iif_files.zip')

    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path(
                'upload-multiple/',
                self.admin_site.admin_view(self.upload_multiple_view),
                name='upload-multiple-excel'
            ),
        ]
        return custom_urls + urls

    def upload_multiple_view(self, request):
        if request.method == 'POST':
            form = MultiFileUploadForm(request.POST, request.FILES)
            if form.is_valid():
                title = form.cleaned_data.get('title')
                files = form.cleaned_data['files']

                zip_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
                with ZipFile(zip_temp, 'w') as zipf:
                    for f in files:
                        doc = Document.objects.create(title=title or f.name, file=f)
                        input_path = doc.file.path
                        output_path = os.path.splitext(input_path)[0] + '.iif'
                        result = convert_excel_to_iif(input_path, output_path)
                        if "Error" not in result:
                            zipf.write(output_path, arcname=os.path.basename(output_path))
                        else:
                            self.message_user(request, f"IIF file not found for {f.name}.", level=messages.WARNING)

                zip_temp.seek(0)
                return FileResponse(zip_temp, as_attachment=True, filename="converted_iif_files.zip")
        else:
            form = MultiFileUploadForm()

        return render(request, 'admin/multi_upload_form.html', {'form': form})
