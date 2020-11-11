#pip install -r requirements.txt
import os
from os import listdir,path
from os.path import isfile, join
import shutil
import win32com.client
from win32com.client import gencache
import time
import PyPDF2
import pprint
import csv 
from datetime import datetime
from PIL import Image
from tika import parser
#in repo
from create_image import split_pdf, image_to_pdf, doc_to_pdf, html_to_pdf, natives_to_pdf, ppt2jpg
from reorder_dat import write_dat
from create_opt import write_opt
from md5_hashing import hash_file
from get_office_properties import get_doc_props_old,get_open_doc_props
import pre_processing

# -----------------INSTRUCTIONS------------------------
# 1. Fill in User Prompts Below
# 2. UPLOAD COLLECTION/COLLECTED DOCUMENTS TO 'input' folder. Extract all compressed files. Make sure files are within a directory inside 'input' (ex. '\input\Test_Collection)
# 3. Run 'email_parser'
# 4. Find output in Open-Source-eDiscovery Diectory (ex. Open-Source_eDiscovery\JSCO_PROD001\VOL0001\NATIVES\JSCO_00000001.xls)

# --------------- USER PROMPTS-------------------------
defaultConfidential = True
defaultCustodian = r"John Smith"
defaultAuthor = r"John Smith"
vol_num = 1     # Default: 1
prod_num = 2    # Default: 1
prefix = "JSCO" # Default: Client Name
start_index = 1 # Default: 1

# ------------------CODEBASE---------------------------
cur = os.getcwd()
folder = f"{cur}\\input"
prod_fol = f"{prefix}_PROD{format(prod_num, '04d')}"
vol_fol = f"VOL{format(vol_num, '05d')}"
n_path = f"\\{prod_fol}\\{vol_fol}\\NATIVES"
i_path = f"\\{prod_fol}\\{vol_fol}\\IMAGES"
t_path = f"\\temp"
m_path = f"\\{prod_fol}\\{vol_fol}\\TEXT"

def setup_output():
    print(f"{cur}\\{prod_fol}\\{vol_fol}")
    if not path.exists(f"{cur}\\{prod_fol}"):
        os.mkdir(f"{cur}\\{prod_fol}")
    if not path.exists(f"{cur}\\{prod_fol}\\{vol_fol}"):
        os.mkdir(f"{cur}\\{prod_fol}\\{vol_fol}")
    if not path.exists(f"{cur}\\{n_path}"):
        os.mkdir(f"{cur}\\{n_path}")
    if not path.exists(f"{cur}\\{t_path}"):
        os.mkdir(f"{cur}\\{t_path}")
    if not path.exists(f"{cur}\\{i_path}"):
        os.mkdir(f"{cur}\\{i_path}")
    if not path.exists(f"{cur}\\{m_path}"):
        os.mkdir(f"{cur}\\{m_path}")

setup_output()
datfile = open(f"{cur}{t_path}\\{prefix}_PRO{format(prod_num, '03d')}.dat","w+",encoding="latin-1")
datfile.write("þProduction::Begin BatesþþProduction::End BatesþþProduction::Begin AttachmentþþProduction::End AttachmentþþProduction::Image CountþþCustodianþþFile NameþþDocument TitleþþAuthorþþED FolderþþEmail FromþþEmail ToþþEmail CCþþEmail BCCþþEmail SubjectþþDate CreatedþþDate Last ModifiedþþDate AccessedþþDate SentþþConfidentiality DesignationþþMD5þþText PrecedenceþþFILE_PATHþ\n")
   
def loop_through_files(index):
    from_email = False
    subfolders = [ f.path for f in os.scandir(folder) if f.is_dir()]
    for sub in subfolders:
        files = [f for f in listdir(sub) if isfile(join(sub, f))]
        for filer in files:
            print(index,filer)
            _, ext = os.path.splitext(filer)
            if ext != ".msg" and ext != ".eml":
                if ext == ".zip":
                    continue
                if "Attachment - " not in filer:
                    if from_email:
                        index = index-1
                        from_email = False
                    index = loop_through_reg_files(filer,sub,index)
            else:
                index = parse_email(sub,filer,index)
                index += 1
                from_email = True
    natives_to_pdf(defaultConfidential,prefix,vol_num,prod_num)
    return index  

def loop_through_reg_files(filer,sub,index):
    m_index = index
    m_index += 1
    n_index = m_index
    imageCount = 1
    nat_bool = False
    pre, ext = os.path.splitext(filer)
    ext = ext.lower()
    full_path = f"{sub}\\{filer}"
    metaData = {}
    metaData["created"] = datetime.fromtimestamp(os.path.getctime(full_path)).strftime("%m/%d/%Y %H:%M:%S")
    metaData["modified"] = datetime.fromtimestamp(os.path.getmtime(full_path)).strftime("%m/%d/%Y %H:%M:%S")
    metaData["accessed"] = datetime.fromtimestamp(os.path.getatime(full_path)).strftime("%m/%d/%Y %H:%M:%S")
    print(metaData)
    mode,ino,dev,nlink,uid,gid,suze,atime,mtime,ctime = os.stat(full_path)
    print( datetime.fromtimestamp(ctime).strftime("%m/%d/%Y %H:%M:%S"), datetime.fromtimestamp(mtime).strftime("%m/%d/%Y %H:%M:%S"),datetime.fromtimestamp(atime).strftime("%m/%d/%Y %H:%M:%S"))
    metaData["e"] = {}
    metaData["e"]["from"] = ""
    metaData["e"]["to"] = ""
    metaData["e"]["cc"] = ""
    metaData["e"]["bcc"] = ""
    metaData["e"]["subject"] = ""
    metaData["e"]["sent_date"] = ""
    metaData["file_name"] = filer.encode('latin-1', 'replace').decode('latin-1')
    metaData["doc_name"] = pre.encode('latin-1', 'replace').decode('latin-1')
    metaData["confidential"] = defaultConfidential
    metaData["custodian"] = defaultCustodian
    metaData["author"] = ""
    if ext == ".pdf":
        if(save_tmp_file(full_path,m_index)):
            pdf_file = PyPDF2.PdfFileReader(full_path,'rb')
            pdf_info = pdf_file.getDocumentInfo()
            pp = pprint.PrettyPrinter(indent=4)
            for i in pdf_info:
                print(i,pdf_info[i])
            c_date = pdf_info['/CreationDate']
            m_date = pdf_info['/ModDate']
            print(c_date,m_date)
            metaData["author"] = pdf_info('/Author') if pdf_info('/Author') != '' else defaultAuthor # This could be pulled from a default prompt
            metaData["created"] = datetime.strptime(c_date, '%Y/%m/%d %H:%M').strftime("%m/%d/%Y %H:%M:%S")
            metaData["modified"] = datetime.strptime(m_date, "D:%Y%m%d%H%M%S-05'00'").strftime("%m/%d/%Y %H:%M:%S")
            n_index, _ = split_pdf(filer,m_index,defaultConfidential,prefix,vol_num,prod_num)
            if n_index > m_index:
                imageCount = (n_index+1) - m_index
    elif ext == ".jpg" or ext == ".jpeg" or ext == ".png" or ext == ".gif":
        if(save_tmp_file(full_path,m_index)):
            image_to_pdf(filer,m_index,defaultConfidential,prefix,vol_num,prod_num)
            attrs = {}
            attrs = get_doc_props_old(sub,filer)
            print(filer,": ",attrs)
            if "Author" in attrs.keys():
                metaData["author"] = attrs['Author']
            if "Creation Date" in attrs.keys():
                metaData["created"] = attrs['Creation Date'].strftime("%m/%d/%Y %H:%M:%S")
            if "Last Save Time" in attrs.keys():
                metaData["modified"] = attrs['Last Save Time'].strftime("%m/%d/%Y %H:%M:%S")
    elif ext == ".doc" or ext == ".docx":
        if(save_tmp_file(full_path,m_index)):
            print("sleep")
            time.sleep(5)
            n_index, _ = doc_to_pdf(filer,m_index,defaultConfidential,prefix,vol_num,prod_num)
            if n_index > m_index:
                imageCount = (n_index+1) - m_index
            attrs = {}
            attrs = get_doc_props_old(sub,filer)
            if "Author" in attrs.keys():
                metaData["author"] = attrs['Author']
            if "Creation Date" in attrs.keys():
                metaData["created"] = attrs['Creation Date'].strftime("%m/%d/%Y %H:%M:%S")
            if "Last Save Time" in attrs.keys():
                metaData["modified"] = attrs['Last Save Time'].strftime("%m/%d/%Y %H:%M:%S")
    elif ext == ".xls" or ext == ".xlsx":
        nat_bool = True
        save_native_file(full_path,m_index)
        attrs = {}
        attrs = get_doc_props_old(sub,filer)
        if "Author" in attrs.keys():
            metaData["author"] = attrs['Author']
        if "Creation Date" in attrs.keys():
            metaData["created"] = attrs['Creation Date'].strftime("%m/%d/%Y %H:%M:%S")
        if "Last Save Time" in attrs.keys():
            metaData["modified"] = attrs['Last Save Time'].strftime("%m/%d/%Y %H:%M:%S")
    elif ext == ".mov" or ext == ".mp4":
        nat_bool = True
        save_native_file(full_path,m_index)
    elif ext == ".html":
        if(save_tmp_file(full_path,m_index)):
            html_to_pdf(filer,m_index,defaultConfidential,prefix,vol_num,prod_num)
    elif ext == ".txt":
        if(save_tmp_file(full_path,m_index)):
            html_to_pdf(filer,m_index,defaultConfidential,prefix,vol_num,prod_num)
    elif ext == ".pptx" or ext == ".ppt":
        nat_bool = True
        save_native_file(full_path,m_index)
    hasher = hash_file(full_path)
    metaData["image_count"] = f"{imageCount}"
    save_metadata(sub,filer,m_index,nat_bool,metaData)
    save_in_dat(sub,filer,m_index,m_index,n_index,False,nat_bool,imageCount,hasher,metaData)
    if n_index>m_index:
        m_index = n_index
    return m_index

def find_email_pdf(name):
    # IMPORTANT - YOU MUST SAVE INDIVIDUAL MSG AS PDF FROM OUTLOOK INTO direr
    pre, _ = os.path.splitext(name)
    direr = f"{cur}\\temp\\EXTEMAIL"
    files = [f for f in listdir(direr) if isfile(join(direr, f))]
    for filer in files:
        p_pre, _ = os.path.splitext(filer)
        if p_pre == pre:
            return direr,filer
    return False,False
def find_email_eml(name):
    # IMPORTANT - YOU MUST SAVE INDIVIDUAL MSG AS PDF FROM OUTLOOK INTO direr
    pre, _ = os.path.splitext(name)
    direr = f"{cur}\\temp\\EMAILS"
    files = [f for f in listdir(direr) if isfile(join(direr, f))]
    for filer in files:
        p_pre, _ = os.path.splitext(filer)
        if(len(filer)==79):
            p_pre = p_pre[:-3]
        if p_pre == pre or p_pre in pre:
            return direr,filer
    return False,False

def parse_email(path,mes,index):
    print(index,mes)
    n_index = index
    orig_path = f"{path}\\{mes}"
    pre, ext = os.path.splitext(mes)
    hasher = hash_file(orig_path)
    metaData = {}
    metaData["created"] = datetime.fromtimestamp(os.path.getctime(orig_path)).strftime("%m/%d/%Y %H:%M:%S")
    metaData["modified"] = datetime.fromtimestamp(os.path.getmtime(orig_path)).strftime("%m/%d/%Y %H:%M:%S")
    metaData["accessed"] = datetime.fromtimestamp(os.path.getatime(orig_path)).strftime("%m/%d/%Y %H:%M:%S")
    metaData["file_name"] = mes.encode('latin-1', 'replace').decode('latin-1')
    metaData["doc_name"] = pre
    metaData["confidential"] = "None"
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    msg = outlook.OpenSharedItem(orig_path)
    metaData["e"] = {}
    metaData["e"]["from"] = msg.SenderName
    metaData["e"]["to"] = msg.To
    metaData["e"]["cc"] = msg.CC
    metaData["e"]["bcc"] = msg.BCC
    metaData["e"]["subject"] = msg.Subject
    metaData["e"]["sent_date"] = msg.SentOn.strftime("%m/%d/%Y %H:%M:%S")
    metaData["author"] = msg.SenderName
    metaData["custodian"] = msg.To.split(";")[0]

    eml_path,eml_file = find_email_eml(mes)
    if eml_file:
        eml_full_path = f"{eml_path}\\{eml_file}"
        metaData["accessed"] = datetime.fromtimestamp(os.path.getatime(eml_full_path)).strftime("%m/%d/%Y %H:%M:%S")
        metaData["created"] = datetime.fromtimestamp(os.path.getctime(eml_full_path)).strftime("%m/%d/%Y %H:%M:%S")
        metaData["modified"] = datetime.fromtimestamp(os.path.getmtime(eml_full_path)).strftime("%m/%d/%Y %H:%M:%S")
        hasher = hash_file(eml_full_path)
        metaData["custodian"] =  "Sadura, Tonya"
    else:
        attrs = get_open_doc_props(msg)
        print(mes,": ", attrs)
        if "Author" in attrs.keys():
            metaData["author"] = attrs['Author']
        if "Creation Dat" in attrs.keys():
            metaData["created"] = attrs['Creation Date'].strftime("%m/%d/%Y %H:%M:%S")
        if "Last Save Time" in attrs.keys():
            metaData["modified"] = attrs['Last Save Time'].strftime("%m/%d/%Y %H:%M:%S")

    e_att = msg.Attachments
    att_count = len(e_att)
    p_sub, p_name = find_email_pdf(mes)
    e_index = index
    if p_sub and p_name:
        if(save_tmp_file(p_sub+'\\'+p_name,index)):
            e_index, _ = split_pdf(p_name,index,defaultConfidential,prefix,vol_num,prod_num)
    
    imageCount = 1
    if e_index > index:
        imageCount = (e_index+1) - index
    metaData["image_count"] = f"{imageCount}"
    save_email_metadata(orig_path,index,msg)
    hasher = hash_file(orig_path)
    if att_count > 0:
        n_index = loop_through_attachments(e_att,index,e_index,att_count)
    else:
        n_index = e_index
    msg.Close(1)
    save_in_dat(path,mes,index,index,n_index,False,False,imageCount,hasher,metaData)
    return n_index

def loop_through_attachments(atts,start_index,index,att_count):
    m_index = index
    att_arrays = []
    for att in atts:
        imageCount = 1
        nat_bool = False
        m_index += 1
        n_index = 0
        pre, ext = os.path.splitext(att.FileName)
        ext = ext.lower()
        tmp_path = f"{cur}\\temp"
        a_dest = f"\\{prefix}{format(m_index, '08d')}"
        PROD_FILENAME = f"{a_dest}{ext}"
        full_path = f"{tmp_path}\\{PROD_FILENAME}"
        metaData = {}
        metaData["author"] = ""
        metaData["file_name"] = att.FileName.encode('latin-1', 'replace').decode('latin-1')
        metaData["doc_name"] = pre.encode('latin-1', 'replace').decode('latin-1')
        metaData["confidential"] = "None"
        metaData["custodian"] = f"Masdal, Annette"
        if ext == ".pdf":
            save_tmp_attachment(att,m_index)
            n_index, _ = split_pdf(att.FileName,m_index,defaultConfidential,prefix,vol_num,prod_num)
            if n_index > m_index:
                imageCount = n_index - m_index
        elif ext == ".jpg" or ext == ".jpeg" or ext == ".png" or ext == ".gif":
            save_tmp_attachment(att,m_index)
            attrs = {}
            attrs = get_doc_props_old(tmp_path,PROD_FILENAME)
            if "Author" in attrs.keys():
                metaData["author"] = attrs['Author']
            if "Creation Date" in attrs.keys():
                metaData["created"] = attrs['Creation Date'].strftime("%m/%d/%Y %H:%M:%S")
            if "Last Save Time" in attrs.keys():
                metaData["modified"] = attrs['Last Save Time'].strftime("%m/%d/%Y %H:%M:%S")
            image_to_pdf(att.FileName,m_index,defaultConfidential,prefix,vol_num,prod_num)
        elif ext == ".doc" or ext == ".docx":
            if(save_tmp_attachment(att,m_index)):
                print("sleep")
                attrs = {}
                attrs = get_doc_props_old(tmp_path,PROD_FILENAME)
                if "Author" in attrs.keys():
                    metaData["author"] = attrs['Author']
                if "Creation Date" in attrs.keys():
                    metaData["created"] = attrs['Creation Date'].strftime("%m/%d/%Y %H:%M:%S")
                if "Last Save Time" in attrs.keys():
                    metaData["modified"] = attrs['Last Save Time'].strftime("%m/%d/%Y %H:%M:%S")
                n_index, _ = doc_to_pdf(att.FileName,m_index,defaultConfidential,prefix,vol_num,prod_num)
                if n_index > m_index:
                    imageCount = n_index - m_index
            
        elif ext == ".xls" or ext == ".xlsx":
            nat_bool = True
            save_native_att(att,m_index)
            attrs = {}
            attrs = get_doc_props_old(tmp_path,PROD_FILENAME)
            if "Author" in attrs.keys():
                metaData["author"] = attrs['Author']
            if "Creation Date" in attrs.keys():
                metaData["created"] = attrs['Creation Date'].strftime("%m/%d/%Y %H:%M:%S")
            if "Last Save Time" in attrs.keys():
                metaData["modified"] = attrs['Last Save Time'].strftime("%m/%d/%Y %H:%M:%S")
        elif ext == ".mov" or ext == ".mp4":
            nat_bool = True
            save_native_att(att,m_index)
        elif ext == ".html":
            save_tmp_attachment(att,m_index)
            html_to_pdf(att.FileName,m_index,defaultConfidential,prefix,vol_num,prod_num)
        elif ext == ".txt":
            save_tmp_attachment(att,m_index)
            html_to_pdf(att.FileName,m_index,defaultConfidential,prefix,vol_num,prod_num)
        elif ext == ".eml":
            save_tmp_attachment(att,m_index)
            #  NEED TO CONVERT TO PDF
        elif ext == ".msg":
            save_tmp_attachment(att,m_index)
            m_index = parse_email(cur+t_path,PROD_FILENAME,m_index)
        
        metaData["image_count"] = f"{imageCount}"
        metaData["created"] = datetime.fromtimestamp(os.path.getctime(full_path)).strftime("%m/%d/%Y %H:%M:%S")
        metaData["modified"] = datetime.fromtimestamp(os.path.getmtime(full_path)).strftime("%m/%d/%Y %H:%M:%S")
        metaData["accessed"] = datetime.fromtimestamp(os.path.getatime(full_path)).strftime("%m/%d/%Y %H:%M:%S")
        metaData["e"] = {}
        metaData["e"]["from"] = ""
        metaData["e"]["to"] = ""
        metaData["e"]["cc"] = ""
        metaData["e"]["bcc"] = ""
        metaData["e"]["subject"] = ""
        metaData["e"]["sent_date"] = ""
        hasher = hash_file(full_path)
        save_metadata(tmp_path,att.FileName,m_index,nat_bool,metaData)
        att_array = [att.FileName,m_index,nat_bool,imageCount,hasher]
        att_arrays.append(att_array)
        if n_index>m_index:
            m_index = n_index
    for arr in att_arrays:
        save_in_dat(tmp_path,arr[0],start_index,arr[1],m_index,True,arr[2],arr[3],arr[4],metaData)
    return m_index

def save_tmp_file(filer,index):
    _, ext = os.path.splitext(filer)
    f_dest = f"\\{prefix}{format(index, '08d')}"
    shutil.copy2(filer,cur+t_path+f_dest+ext)
    if ext == ".png":
        im = Image.open(cur+t_path+f_dest+ext)
        rgb_im = im.convert('RGB')
        rgb_im.save(cur+t_path+f_dest+'.jpg')
    return True

def save_native_file(filer,index):
    _, ext = os.path.splitext(filer)
    f_dest = f"\\{prefix}{format(index, '08d')}"
    shutil.copy2(filer,cur+n_path+f_dest+ext)

def save_tmp_attachment(att,index):
    _, ext = os.path.splitext(att.FileName)
    a_dest = f"\\{prefix}{format(index, '08d')}"
    PROD_FILENAME = f"{a_dest}{ext}"
    att.SaveAsFile(cur+t_path+PROD_FILENAME)
    if ext == ".png":
        im = Image.open(cur+t_path+PROD_FILENAME)
        rgb_im = im.convert('RGB')
        rgb_im.save(cur+t_path+a_dest+'.jpg')
    return True

def save_native_att(att,index):
    _, ext = os.path.splitext(att.FileName)
    prod_filename = f"\\{prefix}{format(index, '08d')}{ext}"
    att.SaveAsFile(cur+n_path+prod_filename)

def save_email(orig_path,index):
    _, ext = os.path.splitext(orig_path)
    e_dest = f"\\{prefix}{format(index, '08d')}"
    shutil.copy2(orig_path,cur+n_path+e_dest+ext)

def save_metadata(orig_path,filer,index,nat_bool,metaData):
    _,ext = os.path.splitext(filer)
    if nat_bool:
        sub = n_path
        new_ext = ext
        nt_path = n_path
    else:
        sub = i_path
        new_ext = ".jpg"
        nt_path = t_path
    
    i_dest = f"\\{prefix}{format(index, '08d')}"
    meta_filename = f"{i_dest}.txt"
    metafile = open(cur+m_path+meta_filename,"w+",encoding="utf-8")
    size = os.path.getsize(cur+nt_path+i_dest+ext)
    PROD_FILENAME = f"{i_dest}{ext}"
    e_create = metaData["created"]
    e_modified = metaData["modified"]
    e_accessed = metaData["accessed"]

    if ext == ".jpg" or ext == ".jpeg" or ext == ".png" or ext == ".gif":
        im = Image.open(cur+nt_path+i_dest+ext)
        w,h  = im.size
    else:
        w = ""
        h= ""
    metafile.write(f"Original Filename: {filer}\n")
    metafile.write(f"File Type: {ext}\n")
    if ext == ".jpg" or ext == ".jpeg" or ext == ".png" or ext == ".gif":
        metafile.write(f"Dimensions: {w} x {h}\n")
    metafile.write(f"File Size: {size/100} KB\n")
    metafile.write(f"File Location: {sub+i_dest+new_ext}\n")
    metafile.write(f"File Description: \n")
    metafile.write(f"Date Created: {e_create}\n")
    metafile.write(f"Last Modified: {e_modified}\n")
    metafile.write(f"Last Accessed: {e_accessed}\n")
    if ext == ".pdf":
        new_path = cur+t_path+i_dest+ext
        rawText = parser.from_file(new_path)
        if rawText['content']:
            rawList = rawText['content'].splitlines()
            if len(rawList) > 0:
                metafile.write(f"Extracted Text: \n")
                for line in rawList:
                    if line.strip() != "":
                        metafile.write(f"{line.strip()}\n")

    if ext == ".html" or ext == ".txt":
        with open(cur+t_path+i_dest+ext, "r", encoding='utf-8') as f:
            text= f.read()
            if text.strip() != "" and text:
                metafile.write(f"Extracted Text Page: {text}\n")
    if ext ==".docx" or ext == ".doc":
        app = win32com.client.Dispatch('Word.Application')
        doc = app.Documents.Open(cur+t_path+i_dest+ext)
        text = doc.Content.Text
        if text and text.strip() != "":
            metafile.write(f"Extracted Text Page: {text}\n")
        app.Quit()

def save_email_metadata(orig_path,index,msg):
    _,tail = os.path.split(orig_path)
    i_dest = f"\\{prefix}{format(index, '08d')}"
    meta_filename = f"{i_dest}.txt"
    metafile = open(cur+m_path+meta_filename,"w+",encoding="utf-8")
    e_from = msg.SenderName
    e_to = msg.To
    e_cc = msg.CC
    e_bcc = msg.BCC
    e_sub = msg.Subject
    e_sent = msg.SentOn.strftime("%m/%d/%Y %H:%M:%S")
    e_body = msg.Body
    metafile.write(f"Original File Name: {tail}\n")
    metafile.write(f"From: {e_from}\n")
    metafile.write(f"To: {e_to}\n")
    metafile.write(f"CC: {e_cc}\n")
    metafile.write(f"BCC: {e_bcc}\n")
    metafile.write(f"Subject: {e_sub}\n")
    metafile.write(f"Sent: {e_sent}\n")
    for att in msg.Attachments:
        metafile.write(f"Attachment: {att.FileName}\n")
    metafile.write(f"Body: \n {e_body}\n")
    metafile.close()
 

def save_in_dat(orig_path,filer,start_index,m_index,end_index,att_bool,nat_bool,imageCount,hasher,metaData):
    orig_full_path = f"{orig_path}\\{filer}"
    _,tail = os.path.split(orig_full_path)
    _, ext = os.path.splitext(tail)
    a_dest = f"\\{prefix}{format(m_index, '08d')}"
    sub = ""
    new_ext = ""
    if nat_bool:
        sub = f"\\{vol_fol}\\NATIVE"
        new_ext = ext
    else:
        sub = f"\\{vol_fol}\\IMAGES"
        new_ext = ".jpg"

    if start_index == m_index:
        if imageCount == 1:
            b_begin = f"{prefix}{format(start_index, '08d')}"
            b_end = f"{prefix}{format(start_index, '08d')}"
        else:
            b_begin = f"{prefix}{format(start_index, '08d')}"
            b_end = f"{prefix}{format(start_index+(imageCount-1), '08d')}"
    else:
        if imageCount == 1:
            b_begin = f"{prefix}{format(m_index, '08d')}"
            b_end = f"{prefix}{format(m_index, '08d')}"
        else:
            b_begin = f"{prefix}{format(m_index, '08d')}"
            b_end = f"{prefix}{format(m_index+(imageCount-1), '08d')}"

    if end_index > start_index:
        b_attach_begin = f"{prefix}{format(start_index, '08d')}"
        b_attach_end = f"{prefix}{format(end_index, '08d')}"
    else:
        b_attach_begin = ""
        b_attach_end = ""
    dest_path = f"{sub}\\{a_dest}{new_ext}"
    text_path = f".\\{vol_fol}\\TEXT\\{a_dest}.txt"
    edFolder = ""
    if ext == ".msg":
        FILE_PATH = dest_path  
    text_prec = text_path
    if metaData["image_count"] == "1":
        FILE_PATH = F".{dest_path}"
    else:
        FILE_PATH = ""
    DAT = [b_begin,b_end,b_attach_begin,b_attach_end,metaData["image_count"],metaData["custodian"],metaData["file_name"],metaData["doc_name"],metaData["author"],edFolder,metaData["e"]["from"],metaData["e"]["to"],metaData["e"]["cc"],metaData["e"]["bcc"],metaData["e"]["subject"],metaData["created"],metaData["modified"],metaData["accessed"],metaData["e"]["sent_date"],str(metaData["confidential"]),hasher,text_prec,FILE_PATH]
    datString = "þþ".join(DAT)
    datfile.write("þ"+datString.encode('latin-1', 'replace').decode('latin-1')+"þ\n")

def main():
    pre_processing.main()
    print("start")
    end_index = loop_through_files(start_index)
    datfile.close()
    print("finished: ",end_index)
    print("create final dat")
    write_dat(prefix,vol_num,prod_num)
    print("write opt")
    write_opt(prefix,prod_num,vol_num,start_index,end_index+1)

main()
