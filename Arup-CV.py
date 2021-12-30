"""
Created: 2021-11-08
Author: Jeffrey.Russell
"""

## IMPORT MODULES
# Gooey:
from gooey import Gooey, GooeyParser
# Start:
import os
from datetime import datetime
# docx2pdf:
from docx2pdf import convert
from shutil import copy2
import os
# pdf2image:
from pdf2image import convert_from_path
from PIL import Image
# tesseract:
import pytesseract
from PIL import Image
# xlsx:
import xlsxwriter

toggle_gooey = True #set to false for debugging

@Gooey(program_name="Arup CV Template")
def parse_args():
    tesseract_path = r"C:\Users\Jeffrey.Russell\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
    parser = GooeyParser(description="Read with Python")
    parser.add_argument("Parent_Folder", widget="DirChooser",help="Select parent folder that contains CVs")
    parser.add_argument("Output_Folder", widget="DirChooser", help="Select output folder to place results")
    args = parser.parse_args()
    parent_folder = args.Parent_Folder
    output_folder = args.Output_Folder
    start(parent_folder, output_folder, tesseract_path)

def main():
    # User Inputs. Use r"[path]" format
    parent_folder = r"C:\Users\Jeffrey.Russell\..."
    output_folder = r"C:\Users\Jeffrey.Russell\..."

    # Tesseract OCR must be saved on the computer. Add path here to the 'tesseract.exe' file. Use r"[path]" format
    tesseract_path = r"C:\Users\Jeffrey.Russell\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
    start(parent_folder, output_folder, tesseract_path)

def start(parent_folder, output_folder, tesseract_path):
    incorrect_folders = ['Assets', 'Headshots', 'Licences and Certifications', 'SF330', 'ss', 'SS', 'Superseded',
                         'superseded', 'sub', 'sub2']
    for root, dirs, files in os.walk(parent_folder):
        for file in files:
            if root.endswith(tuple(incorrect_folders)):  # ignores files in directories that contain old CVs
                continue
            if file == "LastName_FirstName_MasterYEAR_.docx":  # ignores template files
                continue
            if file.endswith(".docx"):
                path = os.path.join(root, file)
                list_docx_files.append(path)
                mod_time = datetime.fromtimestamp(os.path.getmtime(path))
                list_docx_modtime.append(mod_time.strftime("%Y/%m/%d, %H:%M:%S"))
                now = datetime.now()
                delta = now - mod_time
                yrs_old = delta.days // 365
                list_file_age.append(yrs_old)

    for item in list_docx_files:
        pdf_path = docx_to_pdf(item, output_folder)
        img_path = pdf_to_jpeg(pdf_path, output_folder)
        tesseract_OCR(img_path, tesseract_path)

    create_xlsx(output_folder)


def docx_to_pdf(docx_path, output_folder):
    docx_name = docx_path.rsplit('\\')[-1]
    pdf_name = str(docx_name.rstrip('docx') + 'pdf')
    list_names.append(docx_name.rsplit('.')[0])
    pdf_path_loc = str(output_folder + "\\" + pdf_name)
    copied_file = copy2(src=docx_path, dst=str(output_folder + "\\filecopy.docx"))  # create copy of file to avoid overwriting mod dates. Appears in same folder as .py script
    convert(input_path=copied_file, output_path=pdf_path_loc)  # convert docx file to pdf. This overwrites existing files with same name
    os.remove(copied_file) # delete the temporary file
    list_pdf_files.append(pdf_path_loc)
    return pdf_path_loc


def pdf_to_jpeg(pdf_path1, output_folder, res=500):
    pdf_name = pdf_path1.rsplit('\\')[-1]
    img_name = str(pdf_name.rstrip('pdf') + 'jpeg')
    img_path = f'{output_folder}\\{img_name}'
    pages = convert_from_path(pdf_path1, res)  # convert pdf file to jpeg

    for idx, page in enumerate(pages):
        if idx == 0:  # only saving the first page
            page.save(f'{img_path.rstrip(".jpeg")}_{idx}.jpeg','JPEG')  # move this line above if() to save all pages, remove 'continue' below
            img_path_p1 = f'{img_path.rstrip(".jpeg")}_{idx}.jpeg'
            img = Image.open(img_path_p1)
            crop_img = img.crop(box=(0, 0, 1450, 5000))
            crop_img.save(img_path_p1,'JPEG')
            list_img_files.append(img_path_p1)
            continue

    os.remove(pdf_path1)  # delete the temporary file
    return img_path_p1


def tesseract_OCR(img, tess_path):
    pytesseract.pytesseract.tesseract_cmd = tess_path  # better alternative?
    full_text = pytesseract.image_to_string(Image.open(img))

    ## NAME
    # name_line = 10000
    # fullname = "N/A"
    # for idx, line in enumerate(full_text.split("\n")):
    #     print(f"idx = {idx}")
    #     print(f"line = {line}")
    #     if idx == 1:
    #         fullname = line
    #         print(f"fullname = {fullname}")
    #         break

    ## PROFESSION
    profession_line = 10000
    profession = "N/A"
    for idx, line in enumerate(full_text.split("\n")):
        if line == "Profession":
            profession_line = idx + 1
        if idx == profession_line:
            profession = line
            if profession == "":  # If the OCR function doesn't read any value, type "Error"
                list_profession.append("Error")
            else:  # If the OCR function works correctly...
                list_profession.append(profession)
            break

    ## CURRENT POSITION
    position_line = 10000
    position = "N/A"
    for idx, line in enumerate(full_text.split("\n")):
        if line == "Current Position":
            position_line = idx + 1
        if idx == position_line:
            position = line
            if position == "":  # If the OCR function doesn't read any value, type "Error"
                list_current_position.append("Error")
            else:  # If the OCR function works correctly...
                list_current_position.append(position)
            break

    ## JOINED ARUP
    joinedArup_line = 10000
    joinedArup = "N/A"
    for idx, line in enumerate(full_text.split("\n")):
        if line == "Joined Arup":
            joinedArup_line = idx + 1
        if idx == joinedArup_line:
            joinedArup = line
            if joinedArup == "":  # If the OCR function doesn't read any value, type "Error"
                list_JoinedArup.append("Error")
            else:  # If the OCR function works correctly...
                list_JoinedArup.append(joinedArup)
            break

    ## YEARS OF EXPERIENCE
    num_years_line = 10000
    num_years = 0
    for idx, line in enumerate(full_text.split("\n")):
        if line == "Years of Experience":
            num_years_line = idx + 1
        if idx == num_years_line:
            num_years = line
            if num_years == "":  # If the OCR function doesn't read any value, type "Error"
                list_YoE.append("Error")
            else: # If the OCR function works correctly...
                list_YoE.append(num_years)
            break
    # Prevent errors due to other unrelated .docx files in the folder structure (eg. not a resume):
    if num_years == 0:  # if the number of years hasn't changed
        if num_years_line == 10000:  # if the line value hasn't changed
            list_YoE.append("No Info")
        else:
            list_YoE.append("Error")

    os.remove(img)  # delete the temporary file

def create_xlsx(output_folder):
    xlsx_headers = ['Name', 'Profession', 'Current Position', 'Joined Arup', 'Years of Experience', 'Years Since CV Update',
                    'CV Last Modified', 'Word File Path']
    now = datetime.now()
    timestamp = now.strftime("%Y-%m-%d_%Hh-%Mm-%Ss")
    xlsx_path = f'{output_folder}\\CV_Summary_{timestamp}.xlsx'
    outWorkbook = xlsxwriter.Workbook(xlsx_path)
    outSheet = outWorkbook.add_worksheet()
    outSheet.write_row(0, 0, xlsx_headers)
    outSheet.write_column(1, 0, list_names)
    outSheet.write_column(1, 1, list_profession)
    outSheet.write_column(1, 2, list_current_position)
    outSheet.write_column(1, 3, list_JoinedArup)
    outSheet.write_column(1, 4, list_YoE)
    outSheet.write_column(1, 5, list_file_age)
    outSheet.write_column(1, 6, list_docx_modtime)
    outSheet.write_column(1, 7, list_docx_files)
    # outSheet.write_column(1, 8, list_pdf_files)
    # outSheet.write_column(1, 9, list_img_files)

    # highlight if file age is greater than 0
    format1 = outWorkbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})  # Light red fill with dark red text
    outSheet.conditional_format('F2:F100', {'type': 'cell', 'criteria': '>', 'value': 0, 'format': format1})
    # Bold headers
    formatheader = outWorkbook.add_format({'bold': True, 'font_color': 'blue'})
    outSheet.set_row(0,30,formatheader)
    formatwrap = outWorkbook.add_format({'text_wrap': True})
    outSheet.set_column(0,10,15, formatwrap)
    # outSheet.write()
    outWorkbook.close()


# Create blank lists to be filled with results. These lists will be used to create an excel table of all results.
list_names = []
list_docx_files = []
list_docx_modtime = []
list_file_age = []
list_pdf_files = []
list_img_files = []
# list_firstname = []
# list_lastname = []
list_profession = []
list_current_position = []
list_JoinedArup = []
list_YoE = []
# list_qualifications = []
# list_professional_associations = []


# Function call to start the process
if __name__ == '__main__':
    if toggle_gooey:
        parse_args()
    else:
        main()
