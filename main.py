from utils.tableFromDocx import extractTablesFromDocx,extractTablesFromDoc
from utils.tableFromImage import extractTablesFromImage
from utils.tableFromPdf import extractTablesFromPdf
import os
import re
class tableExtractor:

    def __init__(self):
        self.SecretId = input("请输入SecretId:")
        self.SecretKey = input("请输入SecretKey:")

        if not ('table') in os.listdir():
            os.mkdir("table")

        if not ('rawFile') in os.listdir():
            os.mkdir("rawFile")

    def extract(self):
        try:
            while True:
                path = input("请输入文件夹路径(相对or绝对):")
                if path == "":
                    path = "rawFile/"
                files = os.listdir(path)
                for filename in files:
                    self.extractone(path + filename)
        except Exception as e:
            print(e)

    def extractone(self, path):
        filename = re.search(r"[\w-]+\.\w+$",path).group()
        fileType = filename.split(".")[-1]
        print(f"正在处理: {filename} ----")
        if fileType in ["pdf","PDF"]:
            extractTablesFromPdf(path)
        elif fileType in ["docx","DOCX"]:
            extractTablesFromDocx(path)
        elif fileType in ["doc","DOC"]:
            extractTablesFromDoc(path)
        elif fileType in ["png","PNG","jpg","JPG"]:
            extractTablesFromImage(path, self.SecretId , self.SecretKey)


if __name__ == '__main__':
    t = tableExtractor()
    t.extract()

#AKIDtXhStGv9xnOkpjlGuW5wVnuqZ1hClmMg
#DBtwayTjldGrB4iCXdS1L094lIdjqkTv