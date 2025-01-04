import pypdf
import cv2
import numpy as np
import os
import re
import pandas as pd
import shutil
from pdf2image import convert_from_path
from pyzbar.pyzbar import decode

from PIL import ImageFile
ImageFile.LOAD_TRUNCATED_IMAGES = True

def check_credit_code(code):
    codeDict = {
                '0':0,'1':1,'2':2,'3':3,'4':4,'5':5,'6':6,'7':7,'8':8,'9':9,
                'A':10,'B':11,'C':12, 'D':13, 'E':14, 'F':15, 'G':16, 'H':17, 'J':18, 'K':19, 'L':20, 'M':21, 'N':22, 'P':23, 'Q':24,
                'R':25, 'T':26, 'U':27, 'W':28, 'X':29, 'Y':30}
    
    weights = ['1', '3', '9', '27', '19', '26', '16', '17', '20', '29', '25', '13', '8', '24', '10', '30', '28']
    sum = 0
    
    for i in range(len(code) - 1):
        #print(codeDict[code[i]]);
        sum += codeDict[code[i]] * int(weights[i])
    
    mod = 31-sum % 31    
    
    if(mod == codeDict[code[-1]]) or ((mod == 31) and codeDict[code[-1]] == 0):
        return True
    else:
        return False



def pdf_invoice(pdf_folder_path):

    all_data = []
    headers = ['文件名', '发票种类代码', '发票代码', '发票号码', '含税金额', '开票日期', '校验码', '加密字符', '销售方']
    

    pdf_files = [os.path.join(pdf_folder_path, file) for file in os.listdir(pdf_folder_path) if file.endswith('.pdf')]

    for pdf_file in pdf_files:

        image = convert_from_path(pdf_file)[0]
        image = image.crop((0, 0, image.size[0]/4, image.size[1]/4))

        
        # 插入二维码提取数据
        text_list = decode(image)[0].data.decode('utf-8').split(',')
        data = [text_list[1], text_list[2], text_list[3], text_list[4], text_list[5], text_list[6], text_list[7]]
            
        try:
            # 插入卖方
            reader = pypdf.PdfReader(pdf_file)  # 创建 PDF 阅读器对象
            page = reader.pages[0]
            text = page.extract_text()
            text = text.replace("中山百灵生物技术股份有限公司", "")
            text_lines = text.splitlines()
            text_lines = [line for line in text_lines if "公司" in line]
            text_lines = [line for line in text_lines if "银行" not in line]
            seller = text_lines[0].lstrip("名称：").lstrip("：").lstrip()

            data.append(seller)

        except Exception as e:
            print(f"{pdf_file}没有识别出销售方")
            data.append("")
            pass


        file_name = os.path.splitext(os.path.basename(pdf_file))[0]
        data.insert(0, file_name)
        all_data.append(data)

    df = pd.DataFrame(all_data, columns=headers)

    for index, row in df.iterrows():
        # 电子发票（普通发票）显示含税额，增值税显示的未含税额
        if row['发票种类代码'] != '32':
            df.at[index, '含税金额'] = None

    df.to_excel('./发票信息.xlsx', index=False)

def rename(user_name,old_file_folder,new_file_folder):
    # 读取Excel文件，这里假设文件名为data.xlsx，根据实际情况修改文件名和路径
    df = pd.read_excel('./发票信息.xlsx')

    for index, row in df.iterrows():
        old_name = row['文件名'] + '.pdf'
        new_name = str(row['销售方']) + '-' + str(row['发票号码']) + '-' + str(row['含税金额']) + '元-' + str(user_name) +'.pdf'
        old_file_path = os.path.join(old_file_folder, old_name)
        new_file_path = os.path.join(new_file_folder, new_name)
        shutil.copy(old_file_path, new_file_path)
