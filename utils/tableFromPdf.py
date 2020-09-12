import pdfplumber
import pandas as pd
import re
def extractTablesFromPdf(path):
    # extract tables from *.pdf file
    filename = re.search(r"[\w-]+\.\w+$",path).group().split(".")[0]
    DFList = []
    i = -1
    flag = True
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            if page.extract_tables() == []:
                flag = True
                continue

            elif flag:
                flag = False
                i += 1
                DFList.append([])

            for table in page.extract_tables():
                tmpTable = [row for row in table]
                DFList[i] = DFList[i] + tmpTable

    DFList = [pd.DataFrame(x) for x in DFList]
    with pd.ExcelWriter(f"table/{filename}.xlsx") as xlsx:
        for i, df in enumerate(DFList):
            df.to_excel(xlsx, sheet_name=f"{i}_sheet", index=False)

if __name__ == '__main__':
    extractTablesFromPdf("1.pdf")