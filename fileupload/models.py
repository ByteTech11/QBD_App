#fom django.db import models

#cass Document(models.Model):
from django.db import models

class Document(models.Model):
    file = models.FileField(upload_to='uploads/')
    title = models.CharField(max_length=255, blank=True, null=True)

    def __str__(self):
        return self.title if self.title else 'No title'




#   title = models.CharField(max_length=100)
#   file = models.FileField(upload_to='uploads/')
#   def __str__(self):
#       return self.title