import os
from os import listdir,path
from os.path import isfile, join
import shutil
from PIL import Image, ExifTags
import win32com.client
from win32com.client import gencache, Dispatch

cur = os.getcwd()
folder = f"{cur}\\input"
BUILTIN_XLS_ATTRS = ['Title', 'Subject', 'Author', 'Keywords', 'Comments', 'Template', 'Last Author', 'Revision Number',
                     'Application Name', 'Last Print Date', 'Creation Date', 'Last Save Time', 'Total Editing Time',
                     'Number of Pages', 'Number of Words', 'Number of Characters', 'Security', 'Category', 'Format',
                     'Manager', 'Company', 'Number of Btyes', 'Number of Lines', 'Number of Paragraphs',
                     'Number of Slides', 'Number of Notes', 'Number of Hidden Slides', 'Number of Multimedia Clips',
                     'Hyperlink Base', 'Number of Characters (with spaces)']

def loop_old():
    subfolders = [ f.path for f in os.scandir(folder) if f.is_dir()]
    for sub in subfolders:
        files = [f for f in listdir(sub) if isfile(join(sub, f))]
        for filer in files:
           get_doc_props_old(sub,filer)

def get_doc_props_old(sub,filer):
    _, ext = os.path.splitext(filer)
    attrs = {}
    if ext == ".xls" or ext == ".xlsx":
        xl = win32com.client.DispatchEx('Excel.Application')
        # Open the workbook
        wb = xl.Workbooks.Open(sub+"\\"+filer)
        # Save the attributes in a dictionary
        for attrname in BUILTIN_XLS_ATTRS:
            try:
                val = wb.BuiltinDocumentProperties(attrname).Value
                if val:
                    attrs[attrname] = val
            except:
                pass
        wb.Close(SaveChanges=False)
        xl.Quit()
    elif ext == '.doc' or ext == '.docx':
        print("word")
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(sub+"\\"+filer)
        for attrname in BUILTIN_XLS_ATTRS:
            try:
                val = doc.BuiltinDocumentProperties(attrname).Value
                if val:
                    attrs[attrname] = val
            except:
                pass
        doc.Close(SaveChanges=False)
        word.Quit()
    elif ext == ".jpg" or ext ==".jpeg" or ext == ".png" or ext == ".gif":
        img = Image.open(sub+"\\"+filer)
        attrs = { ExifTags.TAGS[k]: v for k, v in img._getexif().items() if k in ExifTags.TAGS }
    return attrs

def get_open_doc_props(o_file):
    attrs = {}
    for attrname in BUILTIN_XLS_ATTRS:
        try:
            val = o_file.BuiltinDocumentProperties(attrname).Value
            if val:
                attrs[attrname] = val
        except:
            pass
    return attrs
    
def get_doc_props(sub,filer):
    _, ext = os.path.splitext(filer)
    if ext == ".xls" or ext == ".xlsx":
        xl = gencache.EnsureDispatch('Excel.Application')
        xl.visible = True
        ss = xl.Workbooks.Open(sub+"\\"+filer)
        props = ss.BuiltInDocumentProperties()
        print(props)
        info = os.stat(sub+"\\"+filer)
        print(info)
        ss.Close()

def loop_through_orig():
    subfolders = [ f.path for f in os.scandir(folder) if f.is_dir()]
    for sub in subfolders:
        shell = Dispatch("Shell.Application")
        _dict = {}
        # enter directory where your file is located
        ns = shell.NameSpace(sub)
        for i in ns.Items():
            # Check here with the specific filename
            if ".xls" in str(i) or ".xlsx" in str(i):
                for j in range(0,49):
                    _dict[ns.GetDetailsOf(j,j)] = ns.GetDetailsOf(i,j)
        print(_dict)
        files = [f for f in listdir(sub) if isfile(join(sub, f))]
        for filer in files:
            get_doc_props(sub,filer)
