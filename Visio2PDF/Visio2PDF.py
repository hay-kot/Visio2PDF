import json
import os
import random
import subprocess
import urllib.request
from datetime import datetime
from pathlib import Path
from time import sleep
from tkinter import Tk, filedialog

import eel
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
from PyPDF2.generic import Bookmark
from reportlab.lib.colors import red
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas

# TODO: Dynamically set watermark page size
# TODO: Create Formatted Excel Table of Contents


# SECTION: Global Variables
APP_VERSION = "v0.41"
CWD = Path(__file__).parent
EXTERNAL_CONVERTER = os.path.join(CWD, "OfficeToPDF.exe")
PDF_DIR = "PDFs"
REPO_URL = "https://api.github.com/repos/hay-kot/Visio2PDF/releases/latest"

# !SECTION: Global Variables


def get_app_version():
    with urllib.request.urlopen(REPO_URL) as response:
        json_values = json.loads(response.read())

    up_to_date = "True"
    repo_version = json_values["tag_name"]
    if APP_VERSION != repo_version:
        up_to_date = "False"

    return up_to_date, repo_version


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


def log(message):
    dateTimeObj = datetime.now()
    timestampStr = dateTimeObj.strftime("%d-%b (%H:%M:%S)")
    eel.putMessageInOutput(f"{timestampStr} | {message}")


def check_create_dir(dir_path: Path):
    if not os.path.isdir(dir_path):
        os.makedirs(dir_path)
        log(f"Creating Save Directory: {dir_path}")


class Watermark:
    def __init__(self, author_name: str, version_tag: str, system_name: str):
        self.author_name = author_name
        self.version_tag = version_tag
        self.system_name = system_name

        dateTimeObj = datetime.now()
        timestampstr = dateTimeObj.strftime("%m/%d/%Y")
        watermark = f"          VERSION: {version_tag}           SYSTEM: {system_name}           AUTHOR: {author_name}           DATE: {timestampstr}"

        c = canvas.Canvas(os.path.join(CWD, "watermark.pdf"))
        c.setPageSize(landscape((792, 1224)))
        c.setFontSize(22)
        c.setFont("Helvetica", 8)
        c.setFillColor(red)
        c.drawString(15, 15, watermark)
        c.save()

        self.watermark = os.path.join(CWD, "watermark.pdf")

    def write(self, PDF: Path):
        output = PDF
        watermark_obj = PdfFileReader(self.watermark)
        watermark_page = watermark_obj.getPage(0)

        pdf_reader = PdfFileReader(PDF)
        pdf_writer = PdfFileWriter()

        for page in range(pdf_reader.getNumPages()):
            page = pdf_reader.getPage(page)
            page.mergePage(watermark_page)
            pdf_writer.addPage(page)

        with open(output, "wb") as out:
            pdf_writer.write(out)


class ConverterTarget:
    def __init__(
        self,
        main_dir: str,
        coversheet: str,
        merged_name: str,
        tag: bool,
        include_subdir: bool,
        author_name: str,
        v_tag: str,
        filetypes: dict,
    ):
        # Main Parameters
        self.main_dir = Path(main_dir)
        self.dir_list = [self.main_dir]
        self.name = merged_name
        self.include_tagging = tag
        self.include_subdir = include_subdir

        if os.path.isfile(coversheet):
            self.coversheet = Path(coversheet)

        # Optional Version Parameters
        if self.include_tagging:
            self.author_name = author_name
            self.tag = v_tag

            self.watermark = Watermark(author_name, tag, self.name)
            self.save_dir = os.path.join(main_dir, PDF_DIR, self.tag)
        else:
            dateTimeObj = datetime.now()
            timestampstr = dateTimeObj.strftime("%m-%d-%Y")
            self.save_dir = os.path.join(main_dir, PDF_DIR, timestampstr)

        check_create_dir(self.save_dir)

        if self.include_subdir:

            for directories in os.walk(self.main_dir):
                if PDF_DIR in directories[0]:
                    continue
                else:
                    self.dir_list.append(directories[0])

        self.file_types = []

        # Create File Type List
        VISIO = [".vsd", ".vsdx", ".vsdm", ".svg"]
        WORD = [".doc", ".dot", ".docx", ".dotx", ".docm", ".dotm", ".rtf", ".wpd"]
        EXCEL = [".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm", ".cs"]
        POWERPOINT = [
            ".ppt",
            ".pptx",
            ".pptm",
            ".pps",
            ".ppsx",
            ".ppsm",
            ".pot",
            ".potx",
            ".potm",
        ]
        PUBLISHER = [".pub"]
        OUTLOOK = [".msg", ".vcf", ".ics"]
        PROJECT = [".mpp"]
        OPENOFFICE = [".odt", ".odp", ".ods"]

        if filetypes["visio"]:
            self.file_types.extend(VISIO)

        if filetypes["word"]:
            self.file_types.extend(WORD)

        if filetypes["excel"]:
            self.file_types.extend(EXCEL)

        if filetypes["powerpoint"]:
            self.file_types.extend(POWERPOINT)

        if filetypes["project"]:
            self.file_types.extend(PROJECT)

        if filetypes["publisher"]:
            self.file_types.extend(PUBLISHER)

        if filetypes["openOffice"]:
            self.file_types.extend(OPENOFFICE)

        if filetypes["outlook"]:
            self.file_types.extend(OUTLOOK)


class Convert:
    @staticmethod
    def coversheet(target: ConverterTarget):
        cover_path = "not a file"
        if os.path.isfile(target.coversheet):
            log(f"Setting Coversheet: {target.coversheet}")
            cover_path = os.path.join(target.save_dir, "CoverSheet.pdf")

            subprocess.run(
                [EXTERNAL_CONVERTER, target.coversheet, cover_path], check=True
            )

            target.coversheetPDF = cover_path
        else:
            log("No Coversheet included, Skipping...")

    @staticmethod
    def files(target: ConverterTarget):
        for dir in target.dir_list:
            save_dir = os.path.join(target.save_dir, os.path.basename(dir))
            for files in os.listdir(dir):
                name, extension = os.path.splitext(files)
                if extension in target.file_types:
                    check_create_dir(save_dir)
                    log(f"Converting: {name}{extension}")
                    visio_file = os.path.join(dir, files)
                    if target.include_tagging:
                        save_name = os.path.join(save_dir, f"{name} {target.tag}.pdf")
                    else:
                        save_name = os.path.join(save_dir, f"{name}.pdf")

                    subprocess.run(
                        [EXTERNAL_CONVERTER, visio_file, save_name], check=True
                    )

                    if target.include_tagging:
                        watermark = Watermark(
                            target.author_name, target.tag, target.name
                        )
                        watermark.write(save_name)
                    else:
                        sleep(0.5)

    @staticmethod
    def preview(target: ConverterTarget):
        preview_files = []
        x = 1

        try:
            if target.coversheet:
                filename = os.path.basename(target.coversheet)
                name, extension = os.path.splitext(filename)
                temp_dict = {"Page": x, "Name": name, "Type": extension}
                preview_files.append(temp_dict)
                x += 1
        except:
            pass

        for dir in target.dir_list:
            for files in os.listdir(dir):
                name, extension = os.path.splitext(files)
                if extension in target.file_types:
                    temp_dict = {"Page": x, "Name": name, "Type": extension}
                    preview_files.append(temp_dict)
                    x += 1

        return json.dumps(preview_files)


def merge_pdfs(target: ConverterTarget):
    log("Merging PDFs...")

    merger = PdfFileMerger()

    try:
        if os.path.isfile(target.coversheetPDF):
            merger.append(target.coversheetPDF)
    except:
        pass

    all_pdfs = []

    for directories in os.walk(target.save_dir):
        all_pdfs.append(directories[0])

    for dir in all_pdfs:
        x = 0
        for files in os.listdir(dir):
            if x == 0:
                bookmark = os.path.basename(dir)
            else:
                bookmark = None
            if files.endswith("pdf"):
                if files == "CoverSheet.pdf":
                    break
                else:
                    merger.append(
                        os.path.join(dir, files),
                        bookmark=bookmark,
                        import_bookmarks=False,
                    )
            x += 1

    if target.include_tagging:
        final_pdf_path = os.path.join(
            target.save_dir, f"{target.name} {target.tag}.pdf"
        )

    else:
        final_pdf_path = os.path.join(target.save_dir, f"{target.name}.pdf")

    merger.write(final_pdf_path)
    merger.close()

    log(f"PDFs Merged: {target.name}.pdf")

    return final_pdf_path


@eel.expose
def getPreview(
    visio_dir,
    coversheet,
    sys_name="Merged",
    insert_version_tag=False,
    version_tag=None,
    author_name="undefined",
    include_subdir=False,
    filetypes=None,
):

    target = ConverterTarget(
        visio_dir,
        coversheet,
        sys_name,
        insert_version_tag,
        include_subdir,
        author_name,
        version_tag,
        filetypes,
    )

    return Convert.preview(target)


@eel.expose
def main(
    visio_dir,
    coversheet,
    sys_name="Merged",
    insert_version_tag=False,
    version_tag=None,
    author_name="undefined",
    include_subdir=False,
    filetypes=None,
):
    log("Starting...")

    target = ConverterTarget(
        visio_dir,
        coversheet,
        sys_name,
        insert_version_tag,
        include_subdir,
        author_name,
        version_tag,
        filetypes,
    )

    Convert.files(target)

    try:
        if target.coversheet:
            Convert.coversheet(target)
    except:
        pass

    # Merge all PDFs into one File
    merged_pdf = merge_pdfs(target)

    # Final Cleanup
    log("Process Complete")

    eel.setMsgVisible()
    eel.toggleButtons()

    os.startfile(merged_pdf)


if __name__ == "__main__":
    eel_path = os.path.join(CWD, "web")
    eel.init(eel_path)
    eel.setVersion(APP_VERSION)

    # Update Version Header
    up_to_date, _repo_version = get_app_version()
    eel.setRepoVersionNotify(up_to_date)

    eel.start("main.html", size=(570, 800), port=0)
