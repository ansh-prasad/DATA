from django import forms

class UploadFileForm(forms.Form):
    file = forms.FileField()
    action = forms.ChoiceField(choices=[('action1', 'ODD'), ('action2', 'EVEN')])