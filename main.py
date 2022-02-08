"""
Created: 2022-01-04
Author: Jeffrey.Russell
"""

# IMPORT REFERENCED SCRIPTS
from search_dir import search
from convert_docx2pdf import docx_to_pdf
from convert_pdf2image import pdf_to_jpeg
from ocr_tesseract import tesseract_OCR
from summary_xlsx import create_xlsx

from gooey import Gooey, GooeyParser

toggle_gooey = True #set to false for debugging

@Gooey(program_name="Read Arup CV", image_dir='icon' ,timing_options = {'show_time_remaining':True})
def parse_args(data):  # this function runs the script with the Gooey GUI. Note that this function is used if toggle_gooey = True.
    parser = GooeyParser(description="Create a summary table of Arup CV info.")
    parser.add_argument("Parent_Folder", widget="DirChooser",help="Select parent folder that contains CVs")
    parser.add_argument("Output_Folder", widget="DirChooser", help="Select output folder to place results. **THIS SHOULD BE ON YOUR LOCAL COMPUTER**")
    args = parser.parse_args()
    data['parent_folder'] = args.Parent_Folder
    data['output_folder'] = args.Output_Folder
    return data

def main(data):  # For debugging. This function runs the script without the Gooey GUI. Note that this function is used if toggle_gooey = False.
    data['parent_folder'] = r"C:\Users\..."  # Use r"[path]" format
    data['output_folder'] = r"C:\Users\..."  # Use r"[path]" format
    return data

# Function call to start the process
if __name__ == '__main__':
    # DEFINE DICTIONARY. ALL DATA IS STORED HERE.
    data = {'parent_folder': [], 'output_folder': [], 'tesseract_path': r"Tesseract-OCR\tesseract.exe", 'list_names': [], 'list_docx_files': [], 'list_docx_modtime': [], 'list_file_age': [], 'list_pdf_files': [],
            'list_img_files': [], 'list_profession': [], 'list_current_position': [], 'list_JoinedArup': [],
            'list_YoE': [], 'xlsx_name': [], 'xlsx_headers': []}

    if toggle_gooey:
        data = parse_args(data)  # run with gooey GUI
    else:
        data = main(data) # run without gooey GUI, for debugging

    # CONTROL THE OPERATIONS
    data = search(data)
    data = docx_to_pdf(data)
    data = pdf_to_jpeg(data)
    data = tesseract_OCR(data)
    data = create_xlsx(data)

