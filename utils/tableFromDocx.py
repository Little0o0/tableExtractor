from docx import Document
import pandas as pd
from win32com import client as wc
import os
import re
def extractTablesFromDocx(path):
    # extract tables from *.docx file
    filename = re.search(r"[\w-]+\.\w+$",path).group().split(".")[0]

    with open(path, 'rb') as f:
        file = Document(f)

    DFList = []
    for table in file.tables:
        tmpTable = [[cell.text for cell in row.cells] for row in table.rows]
        DFList.append(tmpTable)

    tmpDFList  = [DFList[0]]
    j = 0
    for i in range(1, len(DFList)):
        if len(DFList[i][0]) == len(tmpDFList[j][0]):
            tmpDFList[j] = tmpDFList[j] + DFList[i]
        else:
            j += 1
            tmpDFList.append(DFList[i])

    DFList = [pd.DataFrame(x) for x in tmpDFList]

    with pd.ExcelWriter(f"table/{filename}.xlsx") as xlsx:
        for i, df in enumerate(DFList):
            df.to_excel(xlsx, sheet_name=f"{i}_sheet", index=False)

def extractTablesFromDoc(path):
    # 转成docx
    filename = re.search(r"[\w-]+\.\w+$",path).group().split(".")[0]
    word = wc.Dispatch("Word.Application")
    doc = word.Documents.Open(os.path.abspath(path))
    doc.SaveAs(os.path.abspath(f"{filename}.docx"), 12)  # 12表示docx格式
    doc.Close()
    word.Quit()
    extractTablesFromDocx(f"{filename}.docx")
    os.remove(f"{filename}.docx")

if __name__ == '__main__':
    extractTablesFromDocx("1.docx")