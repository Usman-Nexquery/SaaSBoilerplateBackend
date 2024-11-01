import os
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from .serializers import ReportSerializer
from django.http import FileResponse
from apps.report.services.main import ReportWriter
from pathlib import Path
from config.settings import UPLOAD_FOLDER,BASE_DIR2

# Define the upload folder and log directory similar to your Flask app
LOG_DIRECTORY = os.path.join(BASE_DIR2, "logs")

# Ensure the directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(LOG_DIRECTORY, exist_ok=True)


class ReportUploadApi(APIView):
    def post(self, request):
        serializer = ReportSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)

        report_title = serializer.validated_data['report_title']
        patient_name = serializer.validated_data['patient_name']
        pronouns = serializer.validated_data['pronouns']

        # Retrieve and save uploaded files
        uploaded_files = [
            request.FILES.get('bascsrp', None),
            request.FILES.get('bascprs1', None),
            request.FILES.get('bascprs2', None),
            request.FILES.get('basctrs', None),
            request.FILES.get('wisc', None),
            request.FILES.get('wais', None),
            request.FILES.get('woodcock', None),
            request.FILES.get('brown', None),
        ]

        for f in uploaded_files:
            if f is not None:
                file_path = os.path.join(UPLOAD_FOLDER, f.name)
                with open(file_path, 'wb+') as destination:
                    for chunk in f.chunks():
                        destination.write(chunk)

        # Initialize ReportWriter and generate the report
        report_writer = ReportWriter()
        report_file_path = os.path.join(UPLOAD_FOLDER, f"{report_title}.docx")

        report_writer.start(report_file_path, patient_name, pronouns)

        # Return the report file as a downloadable response
        response = FileResponse(open(report_file_path, 'rb'), as_attachment=True, filename=f"{report_title}.docx")
        return response
