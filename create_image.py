from fpdf import FPDF
import os
from os import listdir
from os.path import isfile, join
import ghostscript
import locale
import win32com.client
from pdf2image import convert_from_path
import tempfile
import time
from PIL import Image, ImageDraw, ImageFont

cur = os.getcwd()

def start(prefix,vol_num,prod_num):
    prod_fol = f"{prefix}_PROD{format(prod_num, '04d')}"
    vol_fol = f"VOL{format(vol_num, '05d')}"
    n_path = f"{cur}\\{prod_fol}\\{vol_fol}\\NATIVES"
    i_path = f"{cur}\\{prod_fol}\\{vol_fol}\\IMAGES"
    t_path = f"{cur}\\temp"
    m_path = f"{cur}\\{prod_fol}\\{vol_fol}\\TEXT"
    return prod_fol,vol_fol,n_path,i_path,t_path,m_path

def pdf2jpeg(pdf_input_path, jpeg_output_path):
    ghostscript.cleanup()
    args = ["pef2jpeg", # actual value doesn't matter
            "-dNOPAUSE",
            "-sDEVICE=jpeg",
            "-r144",
            "-sOutputFile=" + jpeg_output_path,
            pdf_input_path]

    encoding = locale.getpreferredencoding()
    args = [a.encode(encoding) for a in args]
    ghostscript.Ghostscript(*args)

def natives_to_pdf(confidential,prefix,vol_num,prod_num): 
    prod_fol,vol_fol,n_path,i_path,t_path,m_path = start(prefix,vol_num,prod_num)
    dirs = [f for f in listdir(n_path) if isfile(join(n_path, f))]      
    for filer in dirs:
        pre, ext = os.path.splitext(filer)
        if ext != ".msg":
            pdf = FPDF()
            pdf.l_margin = 10
            pdf.t_margin = 2
            pdf.add_page()
            pdf.set_auto_page_break(True, margin = 0.25)
            pdf.set_font('arial','B', 13.0)
            if confidential:
                pdf.cell(ln=0, h=5.0, align='C', w=0, txt="CONFIDENTIAL                                                                                 "+pre, border=0)
            else:
                pdf.cell(ln=0, h=5.0, align='C', w=0, txt="                                                                                             "+pre, border=0)
            pdf.set_font('arial','B', 20.0)
            pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
            pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
            pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
            pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
            pdf.cell(ln=1, h=10.0, align='C', w=0, txt="FILE PRODUCED NATIVELY", border=0)
            pdf.output(f"{t_path}\\{pre}.pdf", 'F')
            pdf_to_image(pre,t_path,i_path)

def images_to_pdf(confidential,prefix,vol_num,prod_num): 
    prod_fol,vol_fol,n_path,i_path,t_path,m_path = start(prefix,vol_num,prod_num)
    i_dirs = [f for f in listdir(i_path) if isfile(join(i_path, f))]
    for filer in i_dirs:
        pre, _ = os.path.splitext(filer)
        pdf = FPDF()
        pdf.l_margin = 30
        pdf.add_page()
        pdf.set_auto_page_break(True, margin = 0.25)
        pdf.set_font('arial','B', 10.0)
        if confidential:
            pdf.cell(ln=0, h=5.0, align='C', w=0, txt="CONFIDENTIAL                                                                                        "+pre, border=0)
        else:
            pdf.cell(ln=0, h=5.0, align='C', w=0, txt="                                                                                                    "+pre, border=0)
        pdf.set_font('arial','B', 20.0)
        pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
        pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
        pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
        pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
        name = f"{t_path}\\{filer}"
        pdf.image(name, x = 45, y = None,w=120, type = '', link = '')
        pdf.output(f"{t_path}{pre}.pdf", 'F')
        pdf_to_image(pre,t_path,i_path)

def image_to_pdf(filer,index,confidential,prefix,vol_num,prod_num): 
    prod_fol,vol_fol,n_path,i_path,t_path,m_path = start(prefix,vol_num,prod_num)
    _, ext = os.path.splitext(filer)
    if ext == ".png":
        n_ext = ".jpg"
    else:
        n_ext = ext
    mn_index = format(index, '08d')
    dest = f"{prefix}{mn_index}"
    pdf = FPDF()
    pdf.l_margin = 30
    pdf.add_page()
    pdf.set_auto_page_break(True, margin = 0.25) 
    pdf.set_font('arial','B', 10.0)   
    if confidential:
        pdf.cell(ln=0, h=5.0, align='C', w=0, txt=f"CONFIDENTIAL                                                                                        {dest}", border=0)
    else :
        pdf.cell(ln=0, h=5.0, align='C', w=0, txt=f"                                                                                                    {dest}", border=0)
    pdf.set_font('arial','B', 20.0)
    pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
    pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
    pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
    pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
    name = f"{t_path}\\{dest}{n_ext}"
    typer = n_ext[1:]
    pdf.image(name, x = 45, y = None,w=120, type = typer, link = '')
    pdf.output(f"{t_path}\\{dest}.pdf", 'F')
    pdf_to_image(dest,t_path,i_path)

# PUBLIC AND PRIVATE FUNCTION
def split_pdf(orig_path,m_start_index,confidential,prefix,vol_num,prod_num): 
    prod_fol,vol_fol,n_path,i_path,t_path,m_path = start(prefix,vol_num,prod_num)
    _, ext = os.path.splitext(orig_path)
    a_dest = f"\\{prefix}{format(m_start_index, '08d')}"
    PROD_FILENAME = F"{a_dest}{ext}"
    pages = convert_from_path(t_path+PROD_FILENAME, thread_count=6, output_folder=t_path,fmt='jpeg')
    m_index = m_start_index
    for i in range(len(pages)):
        mn_index = format(m_index, '08d')
        dest = f"{prefix}{mn_index}.jpg"
        full_dest = f"{i_path}\\{dest}"
        page = pages[i]
        page.save(full_dest, 'JPEG')
        if os.path.exists(full_dest):
            add_text_to_pdf_image(dest,confidential,i_path)
        else:
            time.sleep(1)
            add_text_to_pdf_image(dest,confidential,i_path)
        m_index += 1 
    return m_index-1,dest

def doc_to_pdf(orig_path,m_index,confidential,prefix,vol_num,prod_num): 
    prod_fol,vol_fol,n_path,i_path,t_path,m_path = start(prefix,vol_num,prod_num)
    _, ext = os.path.splitext(orig_path)
    mn_index = format(m_index, '08d')
    n_file = f"\\{prefix}{mn_index}"
    dest = f"{n_file}.pdf"
    wdFormatPDF = 17
    in_file = os.path.abspath(t_path+n_file+ext)
    print(t_path+n_file)
    out_file = os.path.abspath(t_path+dest)
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
    return split_pdf(t_path+dest,m_index,confidential,prefix,vol_num,prod_num)

def html_to_pdf(orig_path,index,confidential,prefix,vol_num,prod_num): 
    prod_fol,vol_fol,n_path,i_path,t_path,m_path = start(prefix,vol_num,prod_num)
    _, ext = os.path.splitext(orig_path)
    in_index = format(index, '08d')
    n_path = f"{prefix}{in_index}"
    text = open(t_path+"\\"+n_path+ext).read()
    pdf = FPDF()
    pdf.l_margin = 30
    pdf.add_page()
    pdf.set_auto_page_break(True, margin = 0.25)
    pdf.set_font('arial','B', 10.0)
    if confidential:
        pdf.cell(ln=1, h=10.0, align='C', w=0, txt="CONFIDENTIAL                                                                                        "+n_path, border=0)
    else:
        pdf.cell(ln=1, h=10.0, align='C', w=0, txt="                                                                                                    "+n_path, border=0)
    pdf.set_font('arial','', 13.0)
    pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
    pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
    pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
    pdf.cell(ln=1, h=10.0, align='C', w=0, txt="", border=0)
    pdf.multi_cell(0, 5, text.encode('latin-1', 'replace').decode('latin-1'))
    pdf.output(f"{t_path}\\{n_path}.pdf", 'F')
    pdf_to_image(n_path,t_path,i_path) 


def add_text_to_image(file_name,confidential,prefix,vol_num,prod_num):
    # print("ADD TEXT TO IMAGE",file_name)
    prod_fol,vol_fol,n_path,i_path,t_path,m_path = start(prefix,vol_num,prod_num)
    pre, _ = os.path.splitext(file_name)
    file_path = f"{i_path}\\{file_name}"
    image = Image.open(file_path)
    w,h = image.size
    image2 = Image.new('RGB', (w,h+20), (255,255,255))
    image2.paste(image, (0,20))
    draw = ImageDraw.Draw(image2)
    font = ImageFont.truetype('arialbd.ttf', size=8)
    (x, y) = (10, 5)
    if confidential:
        message = "      CONFIDENTIAL                                                                                        "+pre
    else:
        message = "                                                                                                          "+pre
    color = 'rgb(0, 0, 0)' # black color
    draw.text((x, y), message, fill=color, font=font)
    image2.save(file_path)
    
def ppt2jpg(orig_path,index,confidential,prefix,vol_num,prod_num): 
    prod_fol,vol_fol,n_path,i_path,t_path,m_path = start(prefix,vol_num,prod_num)
    ppt_path = input(orig_path)
    in_index = format(index, '08d')
    n_path = f"{prefix}{in_index}"
    ppt_app = win32com.client.Dispatch('PowerPoint.Application')
    # ppt_app.Visible = False
    ppt = ppt_app.Presentations.Open(ppt_path, ReadOnly= False)
    ppt.SaveAs(f"{t_path}\\{n_path}.pdf", 32)  
    # ppt.SaveAs(f"{i_path}\\{n_path}.jpg", 17)  
    ppt_app.Quit()
    ppt =  None
    ppt_app = None
    split_pdf(f"{t_path}\\{n_path}.pdf",index,confidential,prefix,vol_num,prod_num)

# -----------------------PRIVATE FUNCTIONS------------------------

def pdf_to_image(pre,t_path,i_path):
    orig_path = f"{t_path}\\{pre}.pdf"
    image_dest = f"{i_path}\\{pre}.jpg"
    pdf2jpeg(orig_path,image_dest)


def add_text_to_pdf_image(file_name,confidential,i_path):
    pre, _ = os.path.splitext(file_name)
    file_path = f"{i_path}\\{file_name}"
    image = Image.open(file_path)
    w,h = image.size
    image2 = Image.new('RGB', (w,h+50), (255,255,255))
    image2.paste(image, (0,50))
    draw = ImageDraw.Draw(image2)
    font = ImageFont.truetype('arialbd.ttf', size=35)
    (x, y) = (30, 5)
    if confidential:
        message = "      CONFIDENTIAL                                                                                        "+pre
    else:
        message = "                                                                                                          "+pre
    color = 'rgb(0, 0, 0)' # black color
    draw.text((x, y), message, fill=color, font=font)
    image2.save(file_path)