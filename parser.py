
from pathlib import Path
import csv, sys
import fitz
from PIL import Image
import win32com.client

if len(sys.argv)<2:
    print("Usage: python verify_archive.py <folder>")
    raise SystemExit(1)
ROOT=Path(sys.argv[1])
OUT="verification_results.csv"

def start_word():
    w=win32com.client.DispatchEx("Word.Application");w.Visible=False;w.DisplayAlerts=0;return w
def start_excel():
    e=win32com.client.DispatchEx("Excel.Application");e.Visible=False;e.DisplayAlerts=False;return e
word=start_word(); excel=start_excel()

files=[f for f in ROOT.rglob("*") if f.is_file()]
total=len(files)
with open(OUT,"w",newline="",encoding="utf-8-sig") as f:
    wr=csv.writer(f)
    wr.writerow(["File","Extension","SizeMB","Status","Error"])
    for i,file in enumerate(files,1):
        ext=file.suffix.lower()
        if ext not in [".doc",".docx",".xls",".xlsx",".pdf",".png",".jpg",".jpeg",".tif",".tiff",".gif",".bmp"]:
            continue
        status="OK";err=""
        try:
            if ext in [".doc",".docx"]:
                try:
                    d=word.Documents.Open(str(file),ReadOnly=True,AddToRecentFiles=False,ConfirmConversions=False,Visible=False)
                except Exception:
                    try: word.Quit()
                    except: pass
                    word=start_word()
                    d=word.Documents.Open(str(file),ReadOnly=True,AddToRecentFiles=False,ConfirmConversions=False,Visible=False)
                d.Close(False)
            elif ext in [".xls",".xlsx"]:
                try:
                    wb=excel.Workbooks.Open(str(file),ReadOnly=True,UpdateLinks=False,IgnoreReadOnlyRecommended=True)
                except Exception:
                    try: excel.Quit()
                    except: pass
                    excel=start_excel()
                    wb=excel.Workbooks.Open(str(file),ReadOnly=True,UpdateLinks=False,IgnoreReadOnlyRecommended=True)
                wb.Close(False)
            elif ext==".pdf":
                p=fitz.open(file); p.close()
            else:
                with Image.open(file) as im:
                    im.verify()
        except Exception as e:
            msg=str(e)
            if "password" in msg.lower():
                status="PASSWORD_PROTECTED"
            else:
                status="FAILED"
            err=msg
        wr.writerow([str(file),ext,round(file.stat().st_size/1048576,2),status,err])
        print(f"{i}/{total} {status}: {file}")
try: word.Quit()
except: pass
try: excel.Quit()
except: pass
print("Done. Results:",OUT)
