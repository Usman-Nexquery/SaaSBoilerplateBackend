import shutil
from pathlib import Path

from . import bascIndividual
from . import wisc
from . import woodcock
from . import wais
from . import demographics
from . import brown

import glob
import os
from config.settings import BASE_DIR2


class ReportWriter:
    def __init__(self):
        self.reportFile = os.path.join(BASE_DIR2, "files/report.docx")
        #self.reportFile = 'report.docx'
        self.templateFile = os.path.join(BASE_DIR2, "files/template.docx")
        #self.templateFile = 'template.docx'
        self.bascFile = [os.path.join(BASE_DIR2, 'files/bascprs1.docx'), os.path.join(BASE_DIR2, 'files/bascprs2.docx'),
                         os.path.join(BASE_DIR2, 'files/basctrs.docx'), os.path.join(BASE_DIR2, 'files/bascsrp.docx')]
        #self.bascFile = ['bascprs1.docx', 'bascprs2.docx',
        #                 'basctrs.docx', 'bascsrp.docx']
        self.wiscFile = os.path.join(BASE_DIR2, 'files/wisc.docx')
        #self.wiscFile = 'wisc.pdf'
        self.woodcockFile = os.path.join(BASE_DIR2, 'files/woodcock.docx')
        #self.woodcockFile = 'woodcock.docx'
        self.waisFile = os.path.join(BASE_DIR2, 'files/wais.docx')
        #self.waisFile = 'wais.pdf'
        self.brownFile = os.path.join(BASE_DIR2, 'files/brown.pdf')

    def start(self, reportFile, firstName, pronouns, templateFile = "", bascFile = [], wiscFile = "", woodcockFile = "", waisFile = "", brownFile = ""):
        import os
        if not templateFile:
            templateFile = self.templateFile
        if not bascFile:
            bascFile = self.bascFile
        if not wiscFile:
            wiscFile = self.wiscFile
        if not woodcockFile:
            woodcockFile = self.woodcockFile
        if not waisFile:
            waisFile = self.waisFile
        if not brownFile:
            brownFile = self.brownFile
        if any(os.path.exists(file) for file in bascFile) or os.path.exists(wiscFile) or os.path.exists(woodcockFile):
            templateFile = os.path.join(BASE_DIR2, 'files/template.docx')
        else:
            templateFile = os.path.join(BASE_DIR2, 'files/ADHD.docx')
        print("BASC Exists: ", any(os.path.exists(file) for file in bascFile))
        print("WISC Exists: ", os.path.exists(wiscFile))
        print("Woodcock Exists: ", os.path.exists(woodcockFile))
        print("Report file is: " + reportFile)
        print("Template file is: " + templateFile)
        print("Basc file is: ")
        print(self.bascFile)
        print("Copying template file...")
        shutil.copy(templateFile, reportFile)
        if os.path.exists(self.brownFile):
            try:
                brownText = brown.Brown(reportFile)
            except:
                brownText = None
                print('Error with Brown')
        else:
            brownText = None
        print("Inserting patient name and pronouns")
        demographics.replace_placeholders(reportFile, firstName, pronouns, brownText)
        print("Attempting BASC")
        #basc.update_template_with_averages(bascFile, templateFile, reportFile)
        for filename in bascFile:
            if os.path.exists(filename):
                if 'prs1' in filename:
                    type = 'prs1'
                elif 'prs2' in filename:
                    type = 'prs2'
                elif 'trs' in filename:
                    type = 'trs'
                elif 'srp' in filename:
                    type = 'srp'
                else:
                    print('Invalid filename')
                    return
                t_scores = bascIndividual.extract_t_scores(filename, [3, 7, 9])  # Tables are 0-indexed
                bascIndividual.update_template(reportFile, t_scores, type, reportFile)

        print("Attempting WISC")
        if os.path.exists(wiscFile):
            # Extract table data from PDF
            table_content = wisc.extract_table_from_docx(wiscFile)
            print(table_content)

            # Insert table data into Word document
            wisc.insert_table_into_word(reportFile, table_content, reportFile)

            # Update the paragraph to insert result values
            wisc.update_document(reportFile, table_content, reportFile)

        else:
            print("failed to find the wisc my brutha")
        print("Attempting WAIS")
        # Extract table data from PDF
        if os.path.exists(waisFile):
            table_content = wais.extract_table_from_docx(waisFile)

            # Insert table data into Word document
            wais.insert_table_into_word(reportFile, table_content, reportFile)

            # Update the paragraph to insert result values
            wais.update_document(reportFile, table_content, reportFile)

        print("Attempting Woodcock")
        # Adjust the file paths as necessary
        if os.path.exists(woodcockFile):
            data = woodcock.extract_table_data(woodcockFile)
            woodcock.insert_table_into_word(reportFile, data, reportFile)

            woodcock.update_document(reportFile, data, reportFile)
    def delete_docx_files(self, directory, exclude_file):
        # Construct the file pattern to match .docx files
        file_pattern = os.path.join(directory, "*.docx")

        # Iterate through all .docx files in the directory
        for file_path in glob.glob(file_pattern):
            # Check if the file is not the one to exclude
            if os.path.basename(file_path) not in exclude_file:
                try:
                    # Delete the file
                    os.remove(file_path)
                    print(f"Deleted: {file_path}")
                except Exception as e:
                    print(f"Failed to delete {file_path}: {e}")


#ReportWriter().start('Owen G.docx')
