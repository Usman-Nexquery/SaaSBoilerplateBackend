from importlib.metadata import requires

from rest_framework import serializers

class ReportSerializer(serializers.Serializer):
    report_title = serializers.CharField(max_length=255,required = True)
    patient_name = serializers.CharField(max_length=255, required = True)
    pronouns = serializers.ChoiceField(choices=[
        ('male', 'Male'),
        ('female', 'Female'),
        ('other', 'Other'),
    ] , required = True)
    bascsrp = serializers.FileField(required = False)
    bascprs1 = serializers.FileField(required = False)
    bascprs2 = serializers.FileField(required = False)
    basctrs = serializers.FileField(required = False)
    wisc = serializers.FileField(required = False)
    wais = serializers.FileField(required = False)
    woodcock = serializers.FileField(required = False)
    brown = serializers.FileField(required = False)