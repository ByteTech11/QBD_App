from django.test import TestCase
from django.core.files.uploadedfile import SimpleUploadedFile
from .models import Document

class DocumentUploadTestCase(TestCase):
    
    def setUp(self):
        """Set up any data needed for the tests"""
        self.valid_file = SimpleUploadedFile("test_file.txt", b"File content goes here", content_type="text/plain")
        self.invalid_file = SimpleUploadedFile("test_file.exe", b"Malicious file content", content_type="application/octet-stream")

    def test_valid_file_upload(self):
        """Test that a valid file can be uploaded."""
        response = self.client.post('/api/documents/', {'title': 'Test Document', 'file': self.valid_file})
        self.assertEqual(response.status_code, 201)
        document = Document.objects.get(title='Test Document')
        self.assertEqual(document.title, 'Test Document')
        self.assertTrue(document.file.name.endswith('test_file.txt'))

    def test_invalid_file_upload(self):
        """Test that an invalid file type is rejected."""
        response = self.client.post('/api/documents/', {'title': 'Invalid Document', 'file': self.invalid_file})
        self.assertEqual(response.status_code, 400)

    def test_large_file_upload(self):
        """Test that a file larger than the size limit is rejected."""
        large_file = SimpleUploadedFile("large_file.txt", b"A" * (6 * 1024 * 1024), content_type="text/plain")
        response = self.client.post('/api/documents/', {'title': 'Large File', 'file': large_file})
        self.assertEqual(response.status_code, 400)

