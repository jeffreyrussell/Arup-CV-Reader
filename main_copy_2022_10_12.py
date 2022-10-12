"""
Created: 2022-10-12
Author: Jeffrey.Russell
"""

# IMPORT REFERENCED SCRIPTS
from search_dir_copy_2022_10_12 import search
from convert_docx2pdf_copy_2022_10_12 import docx_to_pdf
# from convert_pdf2image import pdf_to_jpeg
# from ocr_tesseract import tesseract_OCR
# from summary_xlsx import create_xlsx
import sys
import codecs
from gooey import Gooey, GooeyParser

toggle_gooey = False #set to false for debugging

# prevent UTF errors when running through Gooey
if sys.stdout.encoding != 'UTF-8':
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
if sys.stderr.encoding != 'UTF-8':
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

# used to show on Gooey print statements in console
class Unbuffered(object):
    def __init__(self, stream):
        self.stream = stream

    def write(self, data):
        self.stream.write(data)
        self.stream.flush()

    def writelines(self, datas):
        self.stream.writelines(datas)
        self.stream.flush()

    def __getattr__(self, attr):
        return getattr(self.stream, attr)

sys.stdout = Unbuffered(sys.stdout)


@Gooey(program_name="Read Arup CV", image_dir='icon' ,timing_options = {'show_time_remaining':True}, default_size=(1300, 530))
def parse_args(data):  # this function runs the script with the Gooey GUI. Note that this function is used if toggle_gooey = True.
    parser = GooeyParser(description="Create a summary table of Arup CV info.")
    parser.add_argument("Parent_Folder", widget="DirChooser",help="Select parent folder that contains CVs")
    parser.add_argument("Output_Folder", widget="DirChooser", help="Select output folder to place results. **THIS SHOULD BE ON YOUR LOCAL COMPUTER**")
    args = parser.parse_args()
    data['parent_folder'] = args.Parent_Folder
    data['output_folder'] = args.Output_Folder
    return data

def main(data):  # For debugging. This function runs the script without the Gooey GUI. Note that this function is used if toggle_gooey = False.
    # data['parent_folder'] = r"R:\Marketing + Communications\Resumes\Toronto"  # Use r"[path]" format
    data['parent_folder'] = r"R:\Marketing + Communications\Resumes\New York"  # Use r"[path]" format
    data['output_folder'] = r"C:\Users\Jeffrey.Russell\Arup\Sean Walker - 01 - iKaun CV Automation\Offices\New York"  # Use r"[path]" format
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
    # data = pdf_to_jpeg(data)
    # data = tesseract_OCR(data)
    # data = create_xlsx(data)