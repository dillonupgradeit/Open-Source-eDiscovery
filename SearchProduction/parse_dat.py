import os
import csv
import re
from os import listdir
from os.path import isfile, join
from shutil import copyfile
from shutil import copyfile
from os import listdir,path
import random
import string
from datetime import datetime

# User Prompts
search_after = datetime.strptime("04/13/2018", "%m/%d/%Y")
search_terms = ["Lorem","lorem.IPSUM@gmail.com"]
prefix = "JSCO"
vol_nums = [1]                                                  # Default: [1,2]
prod_num = 2                                                    # Default: 1
include_if_no_date = True                                       # Default: True
search_email_recipients_senders_only = False                    # Default: False

cur = os.getcwd()

# THEY MAY CHANGE BELOW from PRO to PROD
prod_fol = f"{prefix}_PROD{format(prod_num, '04d')}"
folders = []
texts = []
for vol_num in vol_nums:
    vol_fol = f"VOL{format(vol_num, '05d')}"
    if os.path.exists(f"{cur}\\{prod_fol}\\{vol_fol}"):
        n_path = f"\\{prod_fol}\\{vol_fol}\\NATIVES"    # \\NATIVE00001
        i_path = f"\\{prod_fol}\\{vol_fol}\\IMAGES"     # \\IMAGES00001
        m_path = f"\\{prod_fol}\\{vol_fol}\\TEXT"       # \\TEXT00001
        images = f"{cur}{i_path}"
        natives = f"{cur}{n_path}"
        text = f"{cur}{m_path}"
        texts.append(text)
        folders.append(text)
        folders.append(images)
        folders.append(natives)

def get_random_string(length):
    letters = string.ascii_lowercase
    result_str = ''.join(random.choice(letters) for i in range(length))
    return result_str

search_fol = f"RESULTS_{format(prod_num, '03d')}_{search_terms[0].upper().replace(' ', '_')}_{get_random_string(5)}"
dest = f"{cur}\\{search_fol}"
print(search_fol)

filtered_docs = []
list_of_docs = []
# //CREATE A LIST OF FILENAMES TO SEARCH FOR IN PRODUCTION BASED ON DAT FILE
dat_path = f"{cur}\\{prod_fol}\\{prefix}_PRO{format(prod_num, '03d')}.dat"
if os.path.exists(dat_path):
    with open(dat_path) as csvfile:
        spamreader = csvfile.readlines()
        list_of_docs = []
        index = 0
        for rowz in spamreader:
            if index != 0:
                row = rowz.split("Ã¾")
                if len(row) == 1:
                    row = rowz.split("þ")
                # print(row)
                docs = []
                if row[5] == '':
                    start = row[1]
                    end = row[3]
                else:
                    start = row[5]
                    end = row[7]
                m = re.search(r"\d", start)
                if m is not None:
                    start_num = int(start[m.start():])
                    end_num = int(end[m.start():])
                    pre = start[:m.start()]
                    # print("start end:",start_num,end_num)
                    for a in range(start_num,end_num+1):
                        tmp_name = pre + str(format(a, '08d'))
                        docs.append(tmp_name)
                    # print(docs)
                file_date_str = ""
                if prod_num == 2:
                    if row[35] == '':
                        file_date_str = row[39]
                    else:
                        file_date_str = row[35]
                else:
                    if row[29] == '':
                        file_date_str = row[31]
                    else:
                        file_date_str = row[29]
                if file_date_str != "":
                    # print("tmp hide becasue just added")
                    file_date = datetime.strptime(file_date_str, "%m/%d/%Y")
                    if file_date >= search_after:
                        is_filtered = False
                        if prod_num == 2:
                            file_details = [row[21],row[27],row[9],row[11],row[13],row[15],row[17]]
                        else:
                            file_details = [row[11],row[13],row[17],row[19],row[21],row[23],row[25]]
                        # print(file_details)
                        for term in search_terms:
                            for detail in file_details:
                                if term in detail.lower():
                                    if docs not in filtered_docs:
                                        filtered_docs.append(docs)
                                    is_filtered = True
                        if is_filtered == False:
                            list_of_docs.append(docs)
                elif include_if_no_date:
                    list_of_docs.append(docs)
            index += 1
        print("done with .dat")
        # print(list_of_docs)
# GO THROUGH FILES IN PRODUCTION AND COPY OUT FILES IN LIST OF NAME (list_of_docs)
# ["1 start","3 end","5 att_start","7 att_end","9 imageCount", "11 Cusodian","13 FileName","15 Folder","17 From","19 To","21 CC","23 Bcc","25 Subject","27 Created", "29 Modified", "31 Sent", "33 TXT", "35 FILEPATH"]
 
# START
filtered_list = []
for docs in filtered_docs:
    for doc in docs:
        filtered_list.append(doc)
# iterate through text and search
for texter in texts:
    dirs = [f for f in listdir(texter) if isfile(join(texter, f))]
    if len(dirs)>0:
        for file_batch in list_of_docs:
            for file_pre in file_batch:
                for img in dirs:
                    if file_pre in img:
                        with open(f"{texter}\\{img}","r",encoding="utf-8") as txtfile:
                            text_in_file = txtfile.read()
                            for term in search_terms:
                                if term.lower() in text_in_file.lower():
                                    pre, _ = os.path.splitext(img)
                                    if pre not in filtered_list:
                                        file_batch_pre = []
                                        for filer in file_batch:
                                            filer_pre, _ = os.path.splitext(filer)
                                            file_batch_pre.append(filer_pre)
                                        # print(file_batch_pre)
                                        filtered_list=filtered_list+file_batch_pre
        for img in dirs:
            pre, _ = os.path.splitext(img)
            if pre not in filtered_list:
                print(f"{texter}\\{img}")
                with open(f"{texter}\\{img}","r",encoding="utf-8") as txtfile:
                    text_in_file = txtfile.read()
                    for term in search_terms:
                        if term.lower() in text_in_file.lower():                       
                            filtered_list.append(pre)
print("done with search")
# print(filtered_list)

if not path.exists(f"{dest}"):
    os.mkdir(f"{dest}")
# copy texts
for folder in folders:
    dirs = [f for f in listdir(folder) if isfile(join(folder, f))]
    for pre in filtered_list:
        for img in dirs:
            if pre in img:
                copyfile(f"{folder}\\{img}", f"{dest}\\{img}")

print("done")


    
