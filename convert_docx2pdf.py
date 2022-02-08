"""
Created: 2022-01-04
Author: Jeffrey.Russell
"""

from docx2pdf import convert
from shutil import copy2
import os


def docx_to_pdf(data):
    print(f"\n---- [2/5] CONVERTING .DOCX FILES TO ACCEPTABLE FORMAT... ----\n")

    output_folder = data['output_folder']
    # CREATE TEMP FOLDER TO STORE COPIED WORD FILES
    temp_folder = os.path.join(output_folder,"temp")
    if os.path.isdir(temp_folder):
        for f in os.listdir(temp_folder):
            try:
                os.remove(os.path.join(temp_folder, f))
            except Exception:
                print('\n ------------------ Error in clearing temp folder, continue')
    else:
        os.makedirs(temp_folder)

    # COPY FILES FROM ORIGIN TO TEMP FOLDER
    for value in data['list_docx_files']:
        print(f"------------ Copying to \\temp: {value}")
        docx_path = value

        docx_name = docx_path.rsplit('\\')[-1]
        pdf_name = str(docx_name.rstrip('docx') + 'pdf')
        data['list_names'].append(docx_name.rsplit('.')[0])
        pdf_path_loc = os.path.join(output_folder, pdf_name)
        temp_docx_path = os.path.join(temp_folder, docx_name)
        copy2(src=docx_path, dst=temp_docx_path)  # create copy of file to avoid overwriting mod dates.
        data['list_pdf_files'].append(pdf_path_loc)

    # CONVERT ALL FILES IN TEMP FOLDER TO PDF, AND PLACE PDFs IN OUTPUT FOLDER
    print(f"\n----------------- ***Converting to pdf. This will take a while, so sit tight and this dialogue will change when it's done.***")
    try:
        convert(input_path=temp_folder, output_path=output_folder)  # convert docx file to pdf. This overwrites existing files with same name
    except AttributeError:
        print('\n ------------------ Attribute error in docx2pdf, continue')
    except Exception:
        print('\n ------------------ Other error in docx2pdf, continue')

    # DELETE ALL FILES IN TEMP FOLDER
    for f in os.listdir(temp_folder):
        try:
            os.remove(os.path.join(temp_folder, f))
        except Exception:
            print('\n ------------------ Error in deleting docx files from temp folder, continue')

    return data
