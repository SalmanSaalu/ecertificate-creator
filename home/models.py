# from distutils.command.upload import upload
from pickle import TRUE
from django.db import models
from django.utils import timezone

# Create your models here.
class excel(models.Model):
    num=models.IntegerField(null=True)
    name=models.CharField(max_length=100,null=True)
    email=models.CharField(max_length=100,null=True)
    excel_file_upload=models.FileField(upload_to='excel',default=0)

class beforefetch(models.Model):
    
    uniquename=models.CharField(max_length=100)
    excel_file_upload=models.FileField(upload_to='before_fetch')
    template_file_upload=models.FileField(upload_to='before_fetch')

class status(models.Model):
    statusname=models.CharField(max_length=10000)

class statusdetails(models.Model):
    statusnamecheck=models.CharField(max_length=10000,default=0)
    email=models.CharField(max_length=100)
    send=models.IntegerField()
    dateadded=models.DateTimeField(default=timezone.now)