import os
import sys
import io
from zipfile import ZipFile
import csv


def escape_val(v):
    return v\
           .replace("&","&amp;")\
           .replace(">","&gt;")\
           .replace("<","&lt;")\
           .replace("\'","&apos;")\
           .replace("\"","&quot;")\
           .replace("\r\n","\n")\
           .replace("\r","\n") if v else ""


sheetdata = io.BytesIO()
sheetdata.write("<?xml version=\"1.0\" encoding=\"utf-8\"?>")
sheetdata.write("<x:worksheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">")
sheetdata.write("<x:sheetData>")

csv_reader = csv.reader(sys.stdin, delimiter=',', quotechar='"')
for row in csv_reader:
    sheetdata.write("<x:row>")
    for cell in row:
        sheetdata.write("<x:c t=\"str\"><x:v>" + escape_val(cell) + "</x:v></x:c>")
    
    sheetdata.write("</x:row>")

sheetdata.write("</x:sheetData>")
sheetdata.write("</x:worksheet>")

zipdata = io.BytesIO()

with ZipFile(zipdata,"a") as zf:
    for root, dirs, files in os.walk("template"):
        for ff in files:
            fullname = os.path.join(root, ff)
            zf.write(fullname, fullname[9:])

    zf.writestr("xl/worksheets/sheet.xml",sheetdata.getvalue())

if sys.platform == 'win32':
    import msvcrt
    msvcrt.setmode(sys.stdout.fileno(), os.O_BINARY)

sys.stdout.write(zipdata.getvalue())