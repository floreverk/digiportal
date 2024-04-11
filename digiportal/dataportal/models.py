from django.db import models
from django.utils import timezone
import uuid

# Create your models here.
class bruikleenmodel(models.Model):
    bruikleennummer = models.CharField(max_length=100, blank=True, unique=True, default=uuid.uuid4)
    date = models.DateField(default=timezone.now, unique=True)
    instelling = models.CharField(max_length=100)
    contactpersoon = models.CharField(max_length=100)
    straat = models.CharField(max_length=300)
    huisnummer = models.IntegerField()
    postcode = models.IntegerField()
    stad = models.CharField(max_length=30)
    telefoon = models.IntegerField()
    email = models.EmailField()
    periode_start = models.DateField(default=timezone.now)
    periode_end = models.DateField(default=timezone.now)

    def __str__(self):
        return str(self.date), str(self.bruikleennummer), str(self.huisnummer), str(self.postcode), str(self.telefoon), str(self.periode_start), str(self.periode_end)

