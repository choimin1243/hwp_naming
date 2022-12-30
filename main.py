import win32com.client as win32
import os
import PyQt5


hwp=win32.gencache.EnsureDispatch("hwpframe.hwpobject")
hwp.XHwpWindows.Item(0).Visible=True
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
path=os.getcwd()
hwp.Open(path+"\\필드.hwp")

excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open(path+"\\필드.xlsx")
ws = wb.Worksheets(1)

xlsx_values = [list(i) for i in ws.UsedRange()]

hwp.Run("CopyPage")

def insert(index, value):
    field_list = list(
            ws.Range(ws.Cells(1,1),
                     ws.Cells(1, 4)).Value[0]
        )
    for idx, field in enumerate(field_list):
        hwp.PutFieldText(f"{field}{{{{{index}}}}}", value[idx])


row = 2
while True:
    if not ws.Cells(row, 1).Value:
        hwp.Run("DeletePage")
        break
    else:
        data = list(
            ws.Range(ws.Cells(row,1),
                     ws.Cells(row, 4)).Value[0]
        )
        insert(row-2, data)
        hwp.Run("PastePage")
        row += 1



