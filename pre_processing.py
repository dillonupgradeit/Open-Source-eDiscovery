import os
from os import listdir,path
from os.path import isfile, join
import shutil
import win32com.client
from win32com.client import gencache
from PIL import Image
from tika import parser
from email.parser import Parser
from email.mime.base import MIMEBase
from msg_parser import MsOxMessage
import codecs 
# in Repo
from md5_hashing import hash_file

cur = os.getcwd()
folder = f"{cur}\\preprocessing"
a_path = f"\\input\\ATTACHMENTS"
msg_path = f"\\input\\EMAILS"
elm_path = f"\\temp\\EMAILS"
tmp_path = f"\\input\\Test_Collection"
t_path = f"\\temp"

def loop_through_files():
    subfolders = [ f.path for f in os.scandir(folder) if f.is_dir()]
    for sub in subfolders:
        files = [f for f in listdir(sub) if isfile(join(sub, f))]
        for filer in files:
            _, ext = os.path.splitext(filer)
            if ext == ".msg":
                parse_email(sub,filer)
            elif ext == ".eml":
                run(sub+"\\"+filer, cur+elm_path+"\\")

# ELM MESSAGES
def parse_message(filename):
    with open(filename,encoding="utf-8") as f:
        return Parser().parse(f)

def find_attachments(message):
    # Return a tuple of parsed content-disposition dict, message object
    # for each attachment found.
    found = []
    for part in message.walk():
        if 'content-disposition' not in part:
            continue
        cdisp = part['content-disposition'].split(';')
        cdisp = [x.strip() for x in cdisp]
        if cdisp[0].lower() != 'attachment':
            continue
        parsed = {}
        for kv in cdisp[1:]:
            key, val = kv.split('=')
            if val.startswith('"'):
                val = val.strip('"')
            elif val.startswith("'"):
                val = val.strip("'")
            parsed[key] = val
        found.append((parsed, part))
    return found

def run(eml_filename, output_dir):
    msg = parse_message(eml_filename)
    print(msg['From'])
    attachments = find_attachments(msg)
    print ("Found {0} attachments...".format(len(attachments)))
    if not os.path.isdir(output_dir):
        os.mkdir(output_dir)
    for cdisp, part in attachments:
        cdisp_filename = os.path.normpath(cdisp['filename'])
        # prevent malicious crap
        if os.path.isabs(cdisp_filename):
            cdisp_filename = os.path.basename(cdisp_filename)
        towrite = os.path.join(output_dir, cdisp_filename)
        print( "Writing " + towrite)
        with open(towrite, 'wb') as fp:
            data = part.get_payload(decode=True)
            fp.write(data)

#  MSG FILES
def parse_email(path,mes):
    orig_path = f"{path}\\{mes}"
    print(orig_path)
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    msg = outlook.OpenSharedItem(orig_path)
    e_att = msg.Attachments
    print(e_att)
    att_count = len(e_att)
    if att_count > 0:
        loop_through_attachments(e_att)
    
def loop_through_attachments(atts):
    for att in atts:
        _, ext = os.path.splitext(att.FileName)
        ext = ext.lower()
        save_tmp_attachment(att,ext)
        
def save_tmp_attachment(att,ext):
    PROD_FILENAME = f"{att.FileName}"
    print(PROD_FILENAME,ext)
    if ext == '.msg':
        print("save email")
        att.SaveAsFile(cur+msg_path+"\\"+att.FileName)
    elif ext == '.eml':
        att.SaveAsFile(cur+elm_path+"\\"+att.FileName)

def main():
    print("start preprocessing")
    if not path.exists(f"{cur}{a_path}"):
        os.mkdir(f"{cur}{a_path}")
    if not path.exists(f"{cur}{elm_path}"):
        os.mkdir(f"{cur}{elm_path}")
    if not path.exists(f"{cur}{msg_path}"):
        os.mkdir(f"{cur}{msg_path}")
    loop_through_files()
