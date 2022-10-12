"""
Created: 2022-10-12
Author: Jeffrey.Russell
"""

import os
from datetime import datetime


def search(data):
    parent_folder = data['parent_folder']
    print(f"\n---- [1/5] SEARCHING PARENT FOLDER AND SUBFOLDERS... ----\n")
    incorrect_folders = ['Assets', 'assets', 'Headshots', '_Headshots', 'Licences and Certifications', 'SF330', 'ss', 'SS', '_ss', '_Ss', 'Superseded',
                         'superseded', 'Superceded', 'superceded', '_Superseded', 'sub', 'sub2', '_Admin', '_Resume Folder Setup', '_Skills', 'Aviation',
                         'CVs for Processing', 'Processed', 'Email', 'proposals with detailed CVs', 'zz_leavers', '_Leavers', '_leavers', 'Previous Versions', 
                         'Tailored CVs', 'French CV', 'Spanish CV', 'Licences and Certifications', 'Licenses and Certifications', '__Folder Structure',
                         '_Licensure', '_Other offices', '_Consultant Resumes', '_Conversion Emails', '_Correspondence', '_New Hires', '_zz_Departed', 
                         'Bogota Resumes', '_Boston Resume Descriptions', '_Transfers', '_Leavers_Departed', '__Leavers_Departed', '__Guidance and Templates',
                         '_2018 Headshots', '_Newhire Resumes (working)', '_Tracking', '_headshot tracker', '_SS Resume Updates', '_Resume Folder Setup', 
                         '_emails', '__DC Project Writeups', '_Bogota Resumes']

    #  FIND DIRECTORIES THAT CONTAIN CVs OF PREVIOUS ARUP STAFF
    contents = os.listdir(parent_folder)
    subdirs = []
    for item in contents:
        joined = os.path.join(parent_folder, item)
        if os.path.isdir(joined):
            subdirs.append(item)

    # CREATE LIST OF STAFF THAT HAVE LEFT ARUP TO PREVENT SCRIPT FROM READING THESE FILES
    leavers = []
    occur_once = False
    if "_Leavers" in subdirs and occur_once == False:
        occur_once = True
        leavercontents = os.listdir(os.path.join(parent_folder, "_Leavers"))
        for item in leavercontents:
            joined = os.path.join(parent_folder, "_Leavers", item)
            if os.path.isdir(joined):
                leavers.append(item)

    # FIND ACCEPTABLE FILES
    for root, dirs, files in os.walk(parent_folder):
        # print(f"root = {root}")
        for file in files:
            # print(f"file = {file}")
            if file.endswith(".docx"):
                if root.find("_Leavers") != -1 or root.find("zz_leavers") != -1 or root.find("Assets") != -1 or root.find("_Superseded") != -1 or root.find("zz_Departed") != -1 or root.find("_Other offices") != -1 or root.find("_Transfers") != -1 or root.find("__Leavers_Departed") != -1 or root.find("_Newhire Resumes (working)") != -1 or root.find("_Leavers_Departed") != -1 or root.find("_Consultant Resumes") != -1 or root.find("_ss") != -1 or root.find("_leavers") != -1 or root.find("_Boston Resume Descriptions") != -1:
                    # print(f"------------ This directory is ignored (Leaver): {root}")
                    continue
                elif root.endswith(tuple(leavers)):
                    # print(f"------------ This directory is ignored (Leaver): {root}")
                    continue
                elif root.endswith(tuple(incorrect_folders)):  # ignores files in directories that contain old CVs
                    # print(f"------------ This directory is ignored (Not primary CV): {root}")
                    continue
                elif file == "LastName_FirstName_MasterYEAR_.docx":  # ignores template files
                    # print(f"------------ This file is ignored (Template): {file}")
                    continue
                elif file == "LastName_FirstName_MasterYEAR.docx":
                    # print(f"------------ This file is ignored (Template): {file}")
                    continue
                elif file.find("LastName") != -1:
                    # print(f"------------ This file is ignored (Template): {file}")
                    continue
                elif file.find("FirstName") != -1:
                    # print(f"------------ This file is ignored (Template): {file}")
                    continue
                elif file.find("~$") != -1:
                    # print(f"------------ This file is ignored (Open file): {file}")
                    continue
                else:
                    path = os.path.join(root, file)
                    print(f"------------ Valid File Found: {path}")
                    data['list_docx_files'].append(path)
                    mod_time = datetime.fromtimestamp(os.path.getmtime(path))
                    data['list_docx_modtime'].append(mod_time.strftime("%Y/%m/%d, %H:%M:%S"))
                    now = datetime.now()
                    delta = now - mod_time
                    yrs_old = delta.days // 365
                    data['list_file_age'].append(yrs_old)
            else:
                continue
    return data
