# Author: Manuel Rosa
# Description: Turns dwg files into PDFs automatically, merges pdfs for easier printing and creates a list of all drawings
import os
import glob
import win32com.client as win32
import tkinter as tk
from tkinter import Tk, ttk
from tkinter.filedialog import askdirectory
import fitz
from tkinter import messagebox
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
import subprocess
import psutil
import shutil
from PyPDF4 import PdfFileWriter, PdfFileReader
import sys
import time


script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
parent_dir = os.path.dirname(script_dir)


def convert_to_pdf(file_path, output_folder):
    sw = win32.Dispatch("SldWorks.Application")
    doc = sw.OpenDoc(file_path, 3)
    pdf_filename = os.path.splitext(os.path.basename(file_path))[0] + ".pdf"
    pdf_path = os.path.join(output_folder, pdf_filename)
    try:
        time.sleep(5)
        doc.SaveAs3(pdf_path, 0, 0)
        print(f"Convertido {file_path} para PDF: {pdf_path}")
    except AttributeError:
        print("Terminado")


def merge_a4_pdfs(pdf_folder, output_path):
    pdf_files = glob.glob(os.path.join(pdf_folder, "*.pdf"))
    pdf_files.sort()

    merged_doc = fitz.open()

    for pdf_file in pdf_files:
        doc = fitz.open(pdf_file)
        for page in doc:
            mediabox = page.mediabox
            page_width = mediabox[2] - mediabox[0]
            page_height = mediabox[3] - mediabox[1]
            is_a4 = (
                page_width < 700 and page_height < 900
            )
            if is_a4:
                merged_doc.insert_pdf(doc, from_page=page.number, to_page=page.number)
    merged_doc.save(output_path)
    merged_doc.close()


def merge_rest_pdfs(pdf_folder, output_path):
    pdf_files = glob.glob(os.path.join(pdf_folder, "*.pdf"))
    pdf_files.sort()

    merged_doc = fitz.open()
    create_rest = False
    for pdf_file in pdf_files:
        doc = fitz.open(pdf_file)
        for page in doc:
            mediabox = page.mediabox
            page_width = mediabox[2] - mediabox[0]
            is_a4 = page_width > 700
            if is_a4:
                create_rest = True
                merged_doc.insert_pdf(doc, from_page=page.number, to_page=page.number)
    if create_rest == True:
        merged_doc.save(output_path)
        merged_doc.close()


def check_sldworks_running():
    for proc in psutil.process_iter():
        try:
            if proc.name() == "SLDWORKS.exe":
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False


def convert_folder_to_pdf(folder_path):
    subprocess.call(["taskkill", "/f", "/im", "SLDWORKS.exe"])
    subprocess.call(["taskkill", "/f", "/im", "Acrobat.exe"])
    os.chdir(folder_path)
    files = glob.glob("*.slddrw")

    if not files:
        print("Nenhum ficheiro de desenho na pasta")
        return

    output_folder = os.path.join(folder_path, "pdf")
    os.makedirs(output_folder, exist_ok=True)
    run_wait(exe)
    for file in files:
        file_path = os.path.join(folder_path, file)
        print(f"Converter {file_path} to PDF...")
        convert_to_pdf(file_path, output_folder)
    data = []

    drawing_files = glob.glob(os.path.join(folder_path, "*.slddrw"))
    processed_filenames = set()

    for drawing_file in drawing_files:
        filename = os.path.splitext(os.path.basename(drawing_file))[0]
        if not filename.startswith("~$") and filename not in processed_filenames:
            data.append([filename])
            processed_filenames.add(filename)

    pdf_file = os.path.join(output_folder, "!lista_desenhos.pdf")
    doc = SimpleDocTemplate(pdf_file, pagesize=(595, 842))
    font_size = 8
    heading_style = ParagraphStyle(
        "Heading1",
        fontSize=14,  
        alignment=1,  
        spaceAfter=12,  
    )
    heading = Paragraph("Lista de Desenhos", style=heading_style)
    table_data = [[]]
    drawing_files = glob.glob(os.path.join(folder_path, "*.slddrw"))
    processed_filenames = set()

    for drawing_file in drawing_files:
        filename = os.path.splitext(os.path.basename(drawing_file))[0]
        if not filename.startswith("~$") and filename not in processed_filenames:
            if len(table_data[-1]) == 1:
                table_data[-1].append(filename)
            else:
                table_data.append([filename])
            processed_filenames.add(filename)

    del table_data[0]

    num_rows = len(table_data)
    num_cols = max(len(row) for row in table_data)

    table = Table(table_data, colWidths=[200] * num_cols)
    table.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, 0), font_size),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 5),
                ("TOPPADDING", (0, 0), (-1, 0), 5),
                ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ]
        )
    )

    for i in range(1, num_rows):
        table.setStyle(
            TableStyle(
                [
                    ("SPAN", (1, i), (1, i)),
                    ("FONTNAME", (0, i), (0, i), "Helvetica"),
                    ("FONTSIZE", (0, i), (0, i), font_size),
                    ("FONTSIZE", (1, i), (1, i), font_size),
                    ("TOPPADDING", (0, i), (0, i), 5),
                    ("BOTTOMPADDING", (0, i), (0, i), 5),
                    ("BOTTOMPADDING", (1, i), (1, i), 5),
                    ("BACKGROUND", (0, i), (-1, i), colors.white),
                    ("GRID", (0, i), (-1, i), 1, colors.black),
                ]
            )
        )
    line_width = 250 
    line_height = 1  
    line_color = colors.black  

    line_style = TableStyle(
        [
            (
                "FONTNAME",
                (0, 0),
                (-1, -1),
                "Helvetica",
            ),  
            ("FONTSIZE", (0, 0), (-1, -1), 12),  
        ]
    )

    received_line = Table(
        [["RECEBI EM:___________________"]], colWidths=[line_width], hAlign="LEFT"
    )
    received_line.setStyle(line_style)

    ass_line = Table(
        [["ASS.  ________________________"]], colWidths=[line_width], hAlign="LEFT"
    )
    ass_line.setStyle(line_style)

    elements = [heading, table, Spacer(1, 60), received_line, Spacer(1, 40), ass_line]

    doc.build(elements)
    logo_image_path = os.path.join(parent_dir, "extras", "Logo.pdf")
    print(logo_image_path)
    put_watermark(pdf_file, pdf_file, logo_image_path)
    new_pdf_file = os.path.join(output_folder, "zlista_desenhos.pdf")
    shutil.copyfile(pdf_file, new_pdf_file)

    subprocess.call(["taskkill", "/f", "/im", "Acrobat.exe"])

    merged_pdf_path_a4 = os.path.join(output_folder, "!imprimir_a4.pdf")
    merge_a4_pdfs(output_folder, merged_pdf_path_a4)

    merged_pdf_path_rest = os.path.join(output_folder, "!imprimir_resto.pdf")
    merge_rest_pdfs(output_folder, merged_pdf_path_rest)


def put_watermark(input_pdf, output_pdf, watermark):
    watermark_instance = PdfFileReader(watermark)
    watermark_page = watermark_instance.getPage(0)

    scale_factor = 1 / 8
    scaled_watermark_width = watermark_page.mediaBox.getUpperRight_x() * scale_factor
    scaled_watermark_height = watermark_page.mediaBox.getUpperRight_y() * scale_factor

    watermark_page.scaleBy(scale_factor)

    pdf_reader = PdfFileReader(input_pdf)
    pdf_writer = PdfFileWriter()
    for page_num in range(pdf_reader.getNumPages()):
        page = pdf_reader.getPage(page_num)

        page_width = page.mediaBox.getUpperRight_x()
        page_height = page.mediaBox.getUpperRight_y()

        page_width = float(page_width)
        page_height = float(page_height)

        x = 47  
        y = (
            page_height - scaled_watermark_height
        )
        page.mergeTranslatedPage(watermark_page, x, y)
        pdf_writer.addPage(page)
    with open(output_pdf, "wb") as out:
        pdf_writer.write(out)


options = os.path.join(parent_dir, "extras", "op.txt")
exe = "wait_window.exe"


def run_wait(exe):
    exe_to_run = os.path.join(parent_dir, "exe", exe)
    file_path = exe_to_run
    subprocess.Popen(file_path, shell=True)


with open(options, "r") as file:
    lines = file.readlines()
if lines:
    first_line = lines[
        0
    ].strip() 
    if first_line:
        macro_to_run = first_line[-1] 

if lines:
    second_line = lines[1].strip()
    if second_line:
        path_sld = second_line[7:]
converted_path_sld = path_sld.replace("\\", r"\\")

selected_folder = askdirectory(title="Selecione a pasta com os ficheiros de desenho.")

if selected_folder:
    convert_folder_to_pdf(selected_folder)
    subprocess.call(["taskkill", "/f", "/im", "SLDWORKS.exe"])
    messagebox.showinfo("Completo", "Processo Terminado")
else:
    messagebox.showinfo("Erro", "Selecione uma pasta com dwgs")
