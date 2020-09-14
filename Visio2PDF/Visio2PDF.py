import os
import random
import subprocess
from datetime import datetime
from pathlib import Path
from time import sleep
from tkinter import Tk, filedialog

import eel
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
from reportlab.lib.colors import red
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas

"""
TODO: Create Option for subdirectory converting
TODO: Dynamically set watermark page size
TODO: Add Support for addiitonal file types in the Microsoft family ( Probably want to include subdirectory support first)
"""


cwd = Path(__file__).parent

eel_path = os.path.join(cwd, "web")
eel.init(eel_path)
external_converter = os.path.join(cwd, "OfficeToPDF.exe")


@eel.expose
def btn_SelectDir():
    root = Tk()
    root.withdraw()
    root.wm_attributes("-topmost", 1)
    folder = filedialog.askdirectory()
    return folder


@eel.expose
def btn_SelectCoverSheet():
    root = Tk()
    root.withdraw()
    root.wm_attributes("-topmost", 1)
    cover_file = filedialog.askopenfilename()
    return cover_file


@eel.expose
def pick_file(folder):
    if os.path.isdir(folder):
        return random.choice(os.listdir(folder))
    else:
        return "Not valid folder"


@eel.expose
def find_cover(dir_path):
    dir_path = Path(dir_path)

    for files in os.listdir(dir_path):
        if "coversheet" in files.lower():
            return os.path.join(dir_path, files)
        if "cover sheet" in files.lower():
            return os.path.join(dir_path, files)

    return "none"


def create_watermark(engineer_name, version_tag, system_name):
    dateTimeObj = datetime.now()
    timestampstr = dateTimeObj.strftime("%m/%d/%Y")
    watermark = f"          VERSION: {version_tag}           SYSTEM: {system_name}           AUTHOR: {engineer_name}           DATE: {timestampstr}"

    c = canvas.Canvas(os.path.join(cwd, "watermark.pdf"))
    c.setPageSize(landscape((792, 1224)))
    c.setFontSize(22)
    c.setFont("Helvetica", 8)
    c.setFillColor(red)
    c.drawString(15, 15, watermark)
    c.save()


def mark_pdf(watermark_file, input_pdf):
    output = input_pdf
    watermark_obj = PdfFileReader(watermark_file)
    watermark_page = watermark_obj.getPage(0)

    pdf_reader = PdfFileReader(input_pdf)
    pdf_writer = PdfFileWriter()

    for page in range(pdf_reader.getNumPages()):
        page = pdf_reader.getPage(page)
        page.mergePage(watermark_page)
        pdf_writer.addPage(page)

    with open(output, "wb") as out:
        pdf_writer.write(out)


def convert_visio(file_dir, save_dir, enable_tagging, tag):
    watermark = os.path.join(cwd, "watermark.pdf")
    for files in os.listdir(file_dir):
        name, extension = os.path.splitext(files)
        if extension == ".vsdx" or extension == ".vsd":
            log(f"Converting: {name}{extension}")
            target = os.path.join(file_dir, files)
            if enable_tagging:
                save_name = os.path.join(save_dir, f"{name} {tag}.pdf")
            else:
                save_name = os.path.join(save_dir, f"{name}.pdf")

            subprocess.run([external_converter, target, save_name], check=True)

            if enable_tagging:
                mark_pdf(watermark, save_name)
            else:
                sleep(0.5)


def merge_pdfs(pdf_dir, new_name, coversheet, enable_tagging, tag):
    log("Merging PDFs...")

    merger = PdfFileMerger()

    if os.path.isfile(coversheet):
        merger.append(coversheet)

    for files in os.listdir(pdf_dir):
        if files.endswith("pdf"):
            if files == "CoverSheet.pdf":
                break
            else:
                merger.append(os.path.join(pdf_dir, files))

    if enable_tagging:
        final_pdf_path = os.path.join(pdf_dir, f"{new_name} {tag}.pdf")

    else:
        final_pdf_path = os.path.join(pdf_dir, f"{new_name}.pdf")

    merger.write(final_pdf_path)
    merger.close()

    log(f"PDFs Merged: {new_name}.pdf")

    return final_pdf_path


def log(message):
    dateTimeObj = datetime.now()
    timestampStr = dateTimeObj.strftime("%d-%b (%H:%M:%S)")
    eel.putMessageInOutput(f"{timestampStr} | {message}")


@eel.expose
def main(
    visio_dir,
    coversheet,
    sys_name="merged",
    insert_version_tag=False,
    version_tag=None,
    engineer_name="undefined",
    include_subdir=False,
):
    log("Starting...")

    if insert_version_tag:
        create_watermark(engineer_name, version_tag, sys_name)

    visio_dir = Path(visio_dir)

    # Set Save Path
    if insert_version_tag == True:
        save_dir = os.path.join(visio_dir, "PDFs", version_tag)
    else:
        dateTimeObj = datetime.now()
        timestampstr = dateTimeObj.strftime("%m-%d-%Y")
        save_dir = os.path.join(visio_dir, "PDFs", timestampstr)

    if not os.path.isdir(save_dir):
        os.makedirs(save_dir)
        log(f"Creating Save Directory: {save_dir}")

    if include_subdir == True:
        visio_dir_list = []

        for item in os.walk(visio_dir):
            if "PDFs" in item[0]:
                continue
            else:
                visio_dir_list.append(item[0])

        for visio_dir in visio_dir_list:
            convert_visio(
                visio_dir,
                save_dir,
                insert_version_tag,
                version_tag,
            )

    elif include_subdir == False:
        convert_visio(
            visio_dir,
            save_dir,
            insert_version_tag,
            version_tag,
        )

    # Cover Sheet
    cover_path = "not a file"
    if os.path.isfile(coversheet):
        log(f"Setting Coversheet: {coversheet}")
        cover_path = os.path.join(save_dir, "CoverSheet.pdf")

        subprocess.run([external_converter, coversheet, cover_path], check=True)
    else:
        log("No Coversheet included, Skipping...")

    merged_pdf = merge_pdfs(
        save_dir, sys_name, cover_path, insert_version_tag, version_tag
    )

    log("Process Complete")

    eel.setMsgVisible()

    os.startfile(merged_pdf)


eel.start("main.html", size=(550, 700), port=0)
