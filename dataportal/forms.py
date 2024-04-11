from django import forms
from .models import bruikleenmodel

class bruikleenform(forms.ModelForm):
    class Meta:
        model = bruikleenmodel
        fields = '__all__'
