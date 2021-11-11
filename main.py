import argparse
import glob
import docx2txt
import datetime
from docx import Document

DATETIME = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")

def extract_bullets(raw):
    batch = glob.glob(raw, recursive=True)
    document = Document()
    for file in batch:
        print(str(file))
        data = docx2txt.process(file)
        data = "".join(data)
        data = data.split("\n\n")
        for bullet in data:
            if bullet == "":
                continue
            bullet = bullet.strip()
            if bullet.startswith("(") or bullet.startswith("-"):
                document.add_paragraph(bullet)
    document.save(f"{str(args.output)}.docx")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Extracts bullets from collection of files.')
    parser.add_argument("-i", "--input", type=str, help='Enter folder location. Use /filepath\**\*.docx format')
    parser.add_argument("-o", "--output", type=str, help='Enter desired file name', default=f"bullet extract {DATETIME}")
    parser.add_argument("-f", "--filepath", type=str, help='Enter desired file path', default=None)
    args = parser.parse_args()
    extract_bullets(f"{str(args.input)}" + "\**\*.docx")


