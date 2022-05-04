"""
Created: 2022-01-04
Author: Jeffrey.Russell
"""

from pdf2image import convert_from_path
from PIL import Image
import os


def pdf_to_jpeg(data, res=500):
    print(f"\n---- [3/5] FILE CONVERSION CONTINUED... ----\n")
    quantity = len(data['list_docx_files'])
    idx_2 = 1

    for value in data['list_pdf_files']:
        print(f"------------ ({idx_2}/{quantity}) - Converting to .jpeg: {value}")
        pdf_path1 = value
        output_folder = data['output_folder']

        pdf_name = pdf_path1.rsplit('\\')[-1]
        img_name = str(pdf_name.rstrip('pdf') + 'jpeg')
        img_path = f'{output_folder}\\{img_name}'

        if os.path.exists(pdf_path1):
            try:
                pages = convert_from_path(pdf_path1, res, poppler_path=r'poppler-0.68.0\bin')  # convert pdf file to jpeg
            except Exception:
                print('\n ------------------ Error in pdf2jpeg, continue')

            for idx, page in enumerate(pages):
                if idx == 0:  # only saving the first page
                    page.save(f'{img_path.rstrip(".jpeg")}_{idx}.jpeg','JPEG')  # move this line above if() to save all pages, remove 'continue' below
                    img_path_p1 = f'{img_path.rstrip(".jpeg")}_{idx}.jpeg'
                    img = Image.open(img_path_p1)
                    crop_img = img.crop(box=(0, 0, 1450, 5000))
                    crop_img.save(img_path_p1,'JPEG')
                    data['list_img_files'].append(img_path_p1)
                    continue

            os.remove(pdf_path1)  # delete the temporary file
        else:
            data['list_img_files'].append("Error")
        idx_2 += 1

    print(f"\n---- FILE CONVERSION COMPLETE ----\n")
    return data