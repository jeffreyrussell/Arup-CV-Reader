"""
Created: 2022-01-04
Author: Jeffrey.Russell
"""

import pytesseract
from PIL import Image
import os


def tesseract_OCR(data):
    print(f"\n---- [4/5] READING CV CONTENTS ----\n")
    pytesseract.pytesseract.tesseract_cmd = data['tesseract_path']
    quantity = len(data['list_docx_files'])
    idx_2 = 1

    for value in data['list_img_files']:
        print(f"------------ ({idx_2}/{quantity}) - Reading: {value}")
        img = value

        if os.path.exists(img):
            try:
                full_text = pytesseract.image_to_string(Image.open(img))
            except Exception:
                print('\n ------------------ Error in Tesseract, continue')

            ## PROFESSION
            profession_line = 10000
            profession = "N/A"
            countA = 0
            for idx, line in enumerate(full_text.split("\n")):
                if line == "Profession":
                    profession_line = idx + 1
                if idx == profession_line:
                    profession = line
                    if profession == "":  # If Tesseract thinks there's a blank row between title and value
                        if countA <= 2:
                            profession_line = idx + 1
                            countA += 1
                            continue
                        else:
                            data['list_profession'].append("Error")
                    else:  # If the OCR function works correctly...
                        data['list_profession'].append(profession)
                    break
            if profession == "N/A":
                data['list_profession'].append(profession)

            ## CURRENT POSITION
            position_line = 10000
            position = "N/A"
            countB = 0
            for idx, line in enumerate(full_text.split("\n")):
                if line == "Current Position" or line == "Current position":
                    position_line = idx + 1
                if idx == position_line:
                    position = line
                    if position == "":  # If the OCR function doesn't read any value, type "Error"
                        if countB <= 2:
                            position_line = idx + 1
                            countB += 1
                            continue
                        else:
                            data['list_current_position'].append("Error")
                    else:  # If the OCR function works correctly...
                        data['list_current_position'].append(position)
                    break
            if position == "N/A":
                data['list_current_position'].append(position)

            ## JOINED ARUP
            joinedArup_line = 10000
            joinedArup = "N/A"
            countC = 0
            for idx, line in enumerate(full_text.split("\n")):
                if line == "Joined Arup":
                    joinedArup_line = idx + 1
                if idx == joinedArup_line:
                    joinedArup = line
                    if joinedArup == "":  # If the OCR function doesn't read any value, type "Error"
                        if countC <= 2:
                            joinedArup_line = idx + 1
                            countC += 1
                            continue
                        else:
                            data['list_JoinedArup'].append("Error")
                    else:  # If the OCR function works correctly...
                        data['list_JoinedArup'].append(joinedArup)
                    break
            if joinedArup == "N/A":
                data['list_JoinedArup'].append(joinedArup)

            ## YEARS OF EXPERIENCE
            num_years_line = 10000
            num_years = 0
            countD = 0
            for idx, line in enumerate(full_text.split("\n")):
                if line == "Years of Experience" or line == "Years of experience":
                    num_years_line = idx + 1
                if idx == num_years_line:
                    num_years = line
                    if num_years == "":  # If the OCR function doesn't read any value, type "Error"
                        if countD <= 2:
                            num_years_line = idx + 1
                            countD += 1
                            continue
                        else:
                            data['list_YoE'].append("Error")
                    else: # If the OCR function works correctly...
                        data['list_YoE'].append(num_years)
                    break
            # Prevent errors due to other unrelated .docx files in the folder structure (eg. not a resume):
            if num_years == 0:  # if the number of years hasn't changed
                if num_years_line == 10000:  # if the line value hasn't changed
                    data['list_YoE'].append("No Info")
                else:
                    data['list_YoE'].append("Error")
            os.remove(img)  # delete the temporary file

        else:
            data['list_profession'].append("File not found")
            data['list_current_position'].append("File not found")
            data['list_JoinedArup'].append("File not found")
            data['list_YoE'].append("File not found")

        idx_2 += 1

    return data