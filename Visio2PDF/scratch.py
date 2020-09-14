import os
from pathlib import Path

VISIO_PATH = Path(r"D:\Programming\Lab\visio_files")


for item in os.walk(VISIO_PATH):
    print(item[0])
    if "PDFs" in item[0]:
        print("TRUE")
