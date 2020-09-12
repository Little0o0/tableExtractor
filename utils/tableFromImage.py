# from PIL import Image
# import pytesseract
import numpy as np
import pandas as pd
import os
import json
import re
import base64
# import matplotlib.pylab as plt
# import cv2
##导入腾讯AI api
from tencentcloud.common import credential
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.ocr.v20181119 import ocr_client, models

#定义函数
def extractTablesFromImage(picture,SecretId,SecretKey):
    filename = re.search(r"[\w-]+\.\w+$",picture ).group().split(".")[0]
    try:
        with open(picture,"rb") as f:
                img_data = f.read()

        img_base64 = base64.b64encode(img_data)
        cred = credential.Credential(SecretId, SecretKey)  #ID和Secret从腾讯云申请
        httpProfile = HttpProfile()
        httpProfile.endpoint = "ocr.tencentcloudapi.com"

        clientProfile = ClientProfile()
        clientProfile.httpProfile = httpProfile
        client = ocr_client.OcrClient(cred, "ap-shanghai", clientProfile)

        req = models.TableOCRRequest()
        params = '{"ImageBase64":"' + str(img_base64, 'utf-8') + '"}'
        req.from_json_string(params)
        resp = client.TableOCR(req)
        #     print(resp.to_json_string())

    except TencentCloudSDKException as err:
        print(err)

    ##提取识别出的数据，并且生成json
    result1 = json.loads(resp.to_json_string())

    rowIndex = []
    colIndex = []
    content = []

    for item in result1['TextDetections']:
        rowIndex.append(item['RowTl'])
        colIndex.append(item['ColTl'])
        content.append(item['Text'])

    ##导出Excel
    ##ExcelWriter方案
    rowIndex = pd.Series(rowIndex)
    colIndex = pd.Series(colIndex)

    index = rowIndex.unique()
    index.sort()

    columns = colIndex.unique()
    columns.sort()

    data = pd.DataFrame(index = index, columns = columns)
    for i in range(len(rowIndex)):
        data.loc[rowIndex[i],colIndex[i]] = re.sub(" ","",content[i])

    # writer = pd.ExcelWriter("../tables/" + re.match(".*\.",f.name).group() + "xlsx", engine='xlsxwriter')
    writer = pd.ExcelWriter(f"table/{filename}.xlsx", engine='xlsxwriter')

    data.to_excel(writer,sheet_name = 'Sheet1', index=False,header = False)
    writer.save()



# Address=input("请输入图片所在文件夹地址：") #输入自己的图片所在地址
# os.chdir(Address)
# pictures = os.listdir()
# for pic in pictures:
#     if ".jpg" in pic:
#         extractTablesFromImage(pic,"AKIDd4tTnstFXZMwGKBliQs68xZwfAj4B3jE","bnqFfqaIkgOIyb4TeZAXp4mLbcx72g5F") #填腾讯云的secretID和secretKey
#         print("已经完成" + pic + "的表格提取.")
#     if ".png" in pic:
#         extractTablesFromImage(pic, "AKIDd4tTnstFXZMwGKBliQs68xZwfAj4B3jE", "bnqFfqaIkgOIyb4TeZAXp4mLbcx72g5F")#填腾讯云的secretID和secretKey
#         print("已经完成" + pic + "的表格提取.")