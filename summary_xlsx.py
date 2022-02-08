"""
Created: 2022-01-04
Author: Jeffrey.Russell
"""

import xlsxwriter
from datetime import datetime


def create_xlsx(data):
    print(f"\n---- [5/5] CREATING SUMMARIZED .XLSX FILE... ----\n")
    output_folder = data['output_folder']
    data['xlsx_headers'] = ['Name', 'Profession', 'Current Position', 'Joined Arup', 'Years of Experience', 'Years Since CV Update',
                    'CV Last Modified', 'Word File Path']
    now = datetime.now()
    timestamp = now.strftime("%Y-%m-%d_%Hh-%Mm-%Ss")
    xlsx_path = f'{output_folder}\\Arup_CV_Summary_{timestamp}.xlsx'
    outWorkbook = xlsxwriter.Workbook(xlsx_path)
    outSheet = outWorkbook.add_worksheet()
    outSheet.write_row(0, 0, data['xlsx_headers'])
    outSheet.write_column(1, 0, data['list_names'])
    outSheet.write_column(1, 1, data['list_profession'])
    outSheet.write_column(1, 2, data['list_current_position'])
    outSheet.write_column(1, 3, data['list_JoinedArup'])
    outSheet.write_column(1, 4, data['list_YoE'])
    outSheet.write_column(1, 5, data['list_file_age'])
    outSheet.write_column(1, 6, data['list_docx_modtime'])
    outSheet.write_column(1, 7, data['list_docx_files'])

    # Conditional formatting for cell if file age is greater than 0
    format1 = outWorkbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})  # Light red fill with dark red text
    outSheet.conditional_format('F2:F1000', {'type': 'cell', 'criteria': '>', 'value': 0, 'format': format1})
    # Bold headers
    formatheader = outWorkbook.add_format({'bold': True, 'font_color': 'blue'})
    outSheet.set_row(0,30,formatheader)
    formatwrap = outWorkbook.add_format({'text_wrap': True})
    outSheet.set_column(0,10,15, formatwrap)

    outWorkbook.close()
    data['xlsx_name'] = xlsx_path.rsplit('\\')[-1]
    print(f"\n---- SUCCESS! REFER TO '{data['xlsx_name']}' IN OUTPUT FOLDER FOR AN EXCEL TABLE OF SUMMARIZED RESULTS ----\n")

    return data
