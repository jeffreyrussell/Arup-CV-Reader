"""
Created: 2022-01-04
Author: Jeffrey.Russell
"""

from docx2pdf import convert
from shutil import copy2
import os, math, time, shutil


def docx_to_pdf(data):
    print(f"\n---- [2/5] CONVERTING .DOCX FILES TO ACCEPTABLE FORMAT... ----\n")

    output_folder = data['output_folder']
    # CREATE TEMP FOLDER TO STORE COPIED WORD FILES
    temp_folder = os.path.join(output_folder,"temp")
    if os.path.isdir(temp_folder):
        for filename in os.listdir(temp_folder):
            file_path = os.path.join(temp_folder, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print('Failed to delete %s. Reason: %s' % (file_path, e))
    else:
        os.makedirs(temp_folder)

    # COPY FILES FROM ORIGIN TO TEMP FOLDER
    dir_capacity = 10  # limits the number of files that docx2pdf will convert at once
    quantity = len(data['list_docx_files'])
    num_buckets = math.ceil(quantity/dir_capacity)
    temp_dir_list = []
    for i in range(num_buckets):  # create subfolders that contain a maximum number of files. This eases the docx2pdf conversion.
        x = str(os.path.join(temp_folder,str(i)))
        temp_dir_list.append(x)
    idx = 0
    for value in data['list_docx_files']:
        print(f"------------ Copying to \\temp: {value}")
        docx_path = value
        docx_name = docx_path.rsplit('\\')[-1]
        pdf_name = str(docx_name.rstrip('docx') + 'pdf')
        data['list_names'].append(docx_name.rsplit('.')[0])
        pdf_path_loc = os.path.join(output_folder, pdf_name)

        dir_index = math.floor(idx/dir_capacity)
        temp_docx_path = os.path.join(temp_n[dir_index], docx_name)

        if os.path.exists(docx_path):
            if os.path.exists(temp_n[dir_index]):
                copy2(src=docx_path, dst=temp_docx_path)  # create copy of file to avoid overwriting mod dates.
            else:
                os.makedirs(temp_n[dir_index])
                copy2(src=docx_path, dst=temp_docx_path)  # create copy of file to avoid overwriting mod dates.
            data['list_pdf_files'].append(pdf_path_loc)
        else:
            print(f"------------------- File recently moved/renamed, ignored: {value}")
            data['list_pdf_files'].append("File moved/renamed")
        idx += 1

    # CONVERT ALL FILES IN DIVIDED TEMP FOLDERS TO PDF, AND PLACE PDFs IN OUTPUT FOLDER
    print(f"\n---- ***Converting to pdf. This will take a while, so sit tight and this dialogue will change when it's done.***")
    idx_1 = 1
    for item in os.listdir(temp_folder):
        print(f"-------- Segment {idx_1}/{len(temp_n)} (***Watch computer for pop-ups from Word. Always select 'Yes' when asked to save, and select 'General' for confidentiality pop-ups.***):")
        if os.path.isdir(os.path.join(temp_folder,str(item))):
            try:
                xyz = convert(input_path=os.path.join(temp_folder,str(item)),output_path=output_folder)  # convert docx file to pdf. This overwrites existing files with same name
                del xyz
            except AttributeError:
                print('------------------ Attribute error in docx2pdf, continue')
                time.sleep(5)
            except Exception:
                print('------------------ Other error in docx2pdf, continue')
                time.sleep(5)
        else:
            print("No directories found. Clear subdirectories within the output folder and try again.")
        idx_1 += 1
        time.sleep(5)

    # DELETE ALL FILES IN TEMP FOLDER
    for filename in os.listdir(temp_folder):
        file_path = os.path.join(temp_folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))
    return data
