import json
import os
import subprocess
import urllib.request
from datetime import datetime
from pathlib import Path
from time import sleep
from tkinter import Tk, filedialog

import eel
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
from reportlab.lib.colors import red
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas

# TODO: Dynamically set watermark page size
# TODO: Create Formatted Excel Table of Contents
# TODO: Fix error on long file path/name
# TODO: Automatically Generate Settings if none Avaiable


# SECTION: Global Variables
APP_VERSION = "v0.6.0"
REPO_URL = "https://api.github.com/repos/hay-kot/Visio2PDF/releases/latest"

CWD = Path(__file__).parent
SETTINGS_DIR = os.path.join(CWD, "settings")
SETTINGS = os.path.join(CWD, "settings", "settings.json")
HISTORY_DIR = os.path.join(CWD, "settings", "history")
BATCH_JOBS_DIR = os.path.join(CWD, "settings", "saved-jobs")
SAVED_JOBS_DIR = os.path.join(CWD, "settings", "saved-jobs")

EXTERNAL_CONVERTER = os.path.join(CWD, "OfficeToPDF.exe")

PDF_DIR = "PDFs"

# !SECTION: Global Variables


def setup():
    dirs = [SETTINGS_DIR, HISTORY_DIR, BATCH_JOBS_DIR, SAVED_JOBS_DIR]

    for dir in dirs:
        check_create_dir(dir)

    if not os.path.isfile(SETTINGS):
        defaults = {"defaultMode": "light"}

        with open(SETTINGS, "w") as f:
            f.write(json.dumps(defaults, indent=4))


def get_app_version():
    with urllib.request.urlopen(REPO_URL) as response:
        json_values = json.loads(response.read())

    up_to_date = "True"
    repo_version = json_values["tag_name"]
    if APP_VERSION != repo_version:
        up_to_date = "False"

    return up_to_date, repo_version


@eel.expose
def write_settings(settings: dict):
    with open(SETTINGS, "r") as f:
        saved_settings = json.load(f)

        saved_settings.update(settings)

    with open(SETTINGS, "w") as f:
        f.write(json.dumps(saved_settings))


def load_settings():
    with open(SETTINGS, "r") as f:
        settings_dict = json.load(f)

    eel.setSettings(settings_dict)


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
def find_cover(dir_path):
    dir_path = Path(dir_path)

    for files in os.listdir(dir_path):
        if "coversheet" in files.lower():
            return os.path.join(dir_path, files)
        if "cover sheet" in files.lower():
            return os.path.join(dir_path, files)

    return


def log(message):
    dateTimeObj = datetime.now()
    timestampStr = dateTimeObj.strftime("%d-%b (%H:%M:%S)")
    eel.putMessageInOutput(f"{timestampStr} | {message}")


def check_create_dir(dir_path: Path):
    if not os.path.isdir(dir_path):
        os.makedirs(dir_path)
        log(f"Creating Directory: {dir_path}")


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


class ConverterJob:
    def __init__(self, job_data: dict):

        # Main Parameters
        self.job_data = job_data
        self.job_dir = Path(job_data["directory"])
        self.dir_list = [self.job_dir]
        self.name = job_data["name"]
        self.include_tagging = job_data["versionTagging"]
        self.include_subdir = job_data["includeSubDir"]

        try:
            self.coversheet = Path(job_data["coverSheet"])
        except:
            pass

        # Optional Version Parameters
        if self.include_tagging:
            self.author_name = job_data["versionData"]["authorName"]
            self.tag = job_data["versionData"]["versionTag"]

            self.watermark = Watermark(self.author_name, self.tag, self.name)
            self.save_dir = os.path.join(self.job_dir, PDF_DIR, self.tag)
        else:
            dateTimeObj = datetime.now()
            timestampstr = dateTimeObj.strftime("%m-%d-%Y")
            self.save_dir = os.path.join(self.job_dir, PDF_DIR, timestampstr)

            self.author_name = ""
            self.tag = ""

        check_create_dir(self.save_dir)

        if self.include_subdir:

            for directories in os.walk(self.job_dir):
                if PDF_DIR in directories[0]:
                    continue
                else:
                    self.dir_list.append(directories[0])

        self.file_types = []

        # Supported File Types
        VISIO = [".vsd", ".vsdx", ".vsdm", ".svg"]
        WORD = [".doc", ".dot", ".docx", ".dotx", ".docm", ".dotm", ".rtf", ".wpd"]
        EXCEL = [".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm", ".cs"]
        POWERPOINT = [".ppt", ".pptx", ".pps", ".ppsx", ".pot", ".potx"]
        PUBLISHER = [".pub"]
        OUTLOOK = [".msg", ".vcf", ".ics"]
        PROJECT = [".mpp"]
        OPENOFFICE = [".odt", ".odp", ".ods"]

        #
        self.visio = job_data["fileTypes"]["visio"]
        self.word = job_data["fileTypes"]["word"]
        self.excel = job_data["fileTypes"]["excel"]
        self.powerpoint = job_data["fileTypes"]["powerpoint"]
        self.project = job_data["fileTypes"]["project"]
        self.publisher = job_data["fileTypes"]["publisher"]
        self.openoffice = job_data["fileTypes"]["openOffice"]
        self.outlook = job_data["fileTypes"]["outlook"]

        if self.visio:
            self.file_types.extend(VISIO)

        if self.word:
            self.file_types.extend(WORD)

        if self.excel:
            self.file_types.extend(EXCEL)

        if self.powerpoint:
            self.file_types.extend(POWERPOINT)

        if self.project:
            self.file_types.extend(PROJECT)

        if self.publisher:
            self.file_types.extend(PUBLISHER)

        if self.openoffice:
            self.file_types.extend(OPENOFFICE)

        if self.outlook:
            self.file_types.extend(OUTLOOK)

        self.lastran = job_data["lastRan"]

        dateTimeObj = datetime.now()
        timestampstr = dateTimeObj.strftime("%d-%b %H:%M")
        self.timestamp = timestampstr

        self.savepath = ""

        self.include_saving = job_data["saveJob"]
        if job_data["saveJob"]:
            self.save_job()

    def returnDICT(self):
        return {
            "name": self.name,
            "directory": str(self.job_dir),
            "includeSubDir": self.include_subdir,
            "coverSheet": str(self.coversheet),
            "versionTagging": self.include_tagging,
            "versionData": {"authorName": self.author_name, "versionTag": self.tag},
            "fileTypes": {
                "visio": self.visio,
                "excel": self.excel,
                "word": self.word,
                "powerpoint": self.powerpoint,
                "publisher": self.publisher,
                "outlook": self.outlook,
                "project": self.project,
                "openOffice": self.openoffice,
            },
            "saveJob": self.include_saving,
            "savePath": str(self.savepath),
            "lastRan": self.timestamp,
        }

    def returnJSON(self):
        return json.dumps(self.returnDICT(), indent=4)

    def save_job(self):
        self.savepath = Path(os.path.join(SAVED_JOBS_DIR, self.name + ".json"))

        data = self.returnJSON()

        with open(self.savepath, "w") as f:
            f.write(data)

    def log_job(self):
        # History Cleanup
        history_list = os.listdir(HISTORY_DIR)

        if len(history_list) >= 4:
            job_history = []

            for file in history_list:
                file = os.path.join(HISTORY_DIR, file)
                job_history.append(file)

            job_history = sorted(job_history, key=os.path.getctime)

            while len(job_history) >= 4:
                os.remove(os.path.join(HISTORY_DIR, job_history[0]))
                del job_history[0]

        # Actual Logging
        save_data = self.returnJSON()

        dateTimeObj = datetime.now()
        timestampstr = dateTimeObj.strftime("%d-%b_%H-%M-%S")

        save_file = os.path.join(HISTORY_DIR, f"{timestampstr} {self.name}.json")

        with open(save_file, "w") as f:
            f.write(save_data)

    @staticmethod
    @eel.expose
    def get_history():
        jobs = ConverterJob.get_job_folder(HISTORY_DIR, "history")
        eel.setHistory(jobs)

    @staticmethod
    @eel.expose
    def get_saved():
        jobs = ConverterJob.get_job_folder(SAVED_JOBS_DIR, "saved")
        eel.setSaved(jobs)

    @staticmethod
    @eel.expose
    def get_job_folder(dir: Path, button: str):
        job_list = []
        x = 1

        for job in os.listdir(dir):
            with open(os.path.join(dir, job), "r") as f:
                details = json.load(f)

                try:
                    coversheet = os.path.isfile(details["coverSheet"])
                except TypeError:
                    coversheet = ""

            data = {
                "#": x,
                "Name": details["name"],
                "Last Ran": details["lastRan"],
                "Directory": details["directory"],
                "Subdirectories": details["includeSubDir"],
                "Cover Sheet": coversheet,
                "Tagging": details["versionTagging"],
                "Author": details["versionData"]["authorName"],
                "Tag": details["versionData"]["versionTag"],
                "Run": button,
                "Import": f"import-{button}",
                "meta": {"path": (os.path.join(dir, job)), "jobType": button},
            }
            job_list.append(data)

            x += 1

        return json.dumps(job_list)

    @staticmethod
    @eel.expose
    def run_file(job_path: str):
        job_path = Path(job_path)

        with open(job_path, "r") as f:
            data = json.load(f)

        main(data)

    @staticmethod
    @eel.expose
    def import_file(job_path: str):
        job_path = Path(job_path)

        with open(job_path, "r") as f:
            data = json.load(f)

        return data


class Convert:
    @staticmethod
    def coversheet(target: ConverterJob):
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
    def files(target: ConverterJob):
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
    def preview(target: ConverterJob):
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


def merge_pdfs(target: ConverterJob):
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

    target.log_job()

    return final_pdf_path


@eel.expose
def getPreview(job_data: dict):
    job = ConverterJob(job_data)

    return Convert.preview(job)


@eel.expose
def main(jobData: dict):
    log("Starting...")

    eel.setWorking()

    job = ConverterJob(jobData)

    Convert.files(job)

    try:
        if job.coversheet:
            Convert.coversheet(job)
    except:
        pass

    # Merge all PDFs into one File
    merged_pdf = merge_pdfs(job)

    # Final Cleanup
    log("Process Complete")

    eel.setMsgVisible()
    eel.toggleButtons()

    ConverterJob.get_history()
    ConverterJob.get_saved()

    os.startfile(merged_pdf)


# Functions to call prior to eel running
def pre_start():
    setup()
    load_settings()
    ConverterJob.get_history()
    ConverterJob.get_saved()
    up_to_date, _repo_version = get_app_version()
    eel.setRepoVersionNotify(up_to_date)


if __name__ == "__main__":
    eel_path = os.path.join(CWD, "web")
    eel.init(eel_path)
    eel.setVersion(APP_VERSION)

    pre_start()

    eel.start("main.html", size=(1080, 725), port=0)
