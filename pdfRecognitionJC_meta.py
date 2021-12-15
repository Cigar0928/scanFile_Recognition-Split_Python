# coding=utf-8
import os
import fitz
import pytesseract
import cv2
import numpy as np
import xlwt
import re
import numpy
from PIL import Image
import socket


def ocr09(imagepath):
    cropped = cv2.imdecode(np.fromfile(imagepath, dtype=np.uint8), -1)
    os.remove(imagepath)
    # 识别封面
    if re.search("0.jpg", imagepath):
        # 移除红章
        red_minus_blue = cropped[:, :, 2] - cropped[:, :, 0]
        red_minus_green = cropped[:, :, 2] - cropped[:, :, 1]
        red_minus_blue = red_minus_blue >= 20
        red_minus_green = red_minus_green >= 20
        red = cropped[:, :, 2] >= numpy.mean(cropped[:, :, 2]) / 2
        mask = red_minus_green & red_minus_blue & red
        cropped[mask, :] = 255
        # 膨胀
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
        imgerode = cv2.erode(cropped, kernel)
        # 二值化
        gray = cv2.cvtColor(imgerode, cv2.COLOR_BGR2GRAY)
        ret, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        string0 = pytesseract.image_to_string(binary, lang='chi_sim', config='--psm 7 -c tessedit_char_blacklist=O〇′_〔〕任〉《')
        string0 = string0.replace(' ', '').replace('〔', '').replace('′', '').replace('_', '')
        string0 = string0.replace('O', '0').replace('〇', '0').replace('任', '4').replace('〕', 'J')
        # print(string0)
        number = re.findall(r"\d+", string0)
        # print(number)
        research1 = re.search("430*", string0)
        if research1:
            span1 = research1.span()
            num = string0[span1[0]:span1[0]+19]
            print(num)
            return num
        else:
            return ''
    # 识别审批表
    else:

        # 二值化
        gray = cv2.cvtColor(cropped, cv2.COLOR_BGR2GRAY)
        ret, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        cv2.imwrite(imagepath, binary)
        # 膨胀
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
        imgerode = cv2.erode(binary, kernel)

        string1 = pytesseract.image_to_string(cropped, 'chi_sim', False, 'mywhitelist8')
        string1 = string1.replace(' ', '').replace('|', '').replace('“', '').replace('\n', '')
        # print(string)
        research1 = re.search('申请人', string1)
        research0 = re.search('在册人口', string1)
        if research1 and research0:
            span1 = research1.span()
            span0 = research0.span()
            qlr = string1[span1[1]:span0[0]-2]
            print(qlr)
            return qlr
        elif research1:
            span1 = research1.span()
            qlr = string1[span1[1]:]
            print(qlr)
            return qlr
        elif research0:
            span0 = research0.span()
            qlr = string1[:span0[0]-2]
            print(qlr)
            return qlr
        else:
            return ''


def pdf2jpg(pdf0, index1):
    if index1 == 0:
        box = (0, 4000, 5950, 4900)
    else:
        box = (0, 800, 5950, 1150)
        # box = (0, 0, 5950, 1150)
    # 另存为jpg
    page = pdf0[index1]
    zoom = int(1000)
    trans = fitz.Matrix(zoom / 100.0, zoom / 100.0)
    pm = page.getPixmap(matrix=trans, alpha=False)
    jpgname = pdf0.name.split(".")
    jpgname = jpgname[0] + str(index1)
    pm.save(jpgname + ".jpg")
    # 提高分辨率
    enhanceimg = Image.open(jpgname + '.jpg')
    enhanceimg1 = enhanceimg.crop(box)
    enhanceimg1.save(jpgname + '.jpg', dpi=(600.0, 600.0))
    return jpgname + '.jpg'


def writexls(renlist0, index0):
    # 创建一个workbook 设置编码
    workbook = xlwt.Workbook(encoding='utf-8')
    # 创建一个worksheet
    worksheet = workbook.add_sheet('QiDongP9')

    if index0 == '1':
        # 写入excel
        # 参数对应 行, 列, 值
        for j in range(len(renlist0)):
            if renlist0[j]:
                worksheet.write(j, 0, label=renlist0[j][0])
                worksheet.write(j, 1, label=renlist0[j][1])
                worksheet.write(j, 2, label=renlist0[j][2])
                # worksheet.write(j, 3, label=renlist0[j][3])
                # worksheet.write(j, 4, label=renlist0[j][4])
                worksheet.write(j, 5, label=".pdf")
        # 保存
        workbook.save('签字表识别宗地代码.xls')
    else:
        # 写入excel
        # 参数对应 行, 列, 值
        for j in range(len(renlist0)):
            if renlist0[j]:
                worksheet.write(j, 0, label=renlist0[j][0])
                worksheet.write(j, 1, label=renlist0[j][1])
                worksheet.write(j, 2, label=renlist0[j][2])
                worksheet.write(j, 3, label=renlist0[j][3])
                # worksheet.write(j, 4, label=renlist0[j][4])
                worksheet.write(j, 5, label=".pdf")
        # 保存
        workbook.save('签字表识别宗地代码+户主.xls')


def setPytesseract():
    # 设置路径
    pytesseract.pytesseract.tesseract_cmd = r'F:\Program Files\Tesseract-OCR\tesseract.exe'
    myname = socket.getfqdn(socket.gethostname())
    if myname != "DESKTOP-LSJ":
        pytesseract.pytesseract.tesseract_cmd = r'D:\Program Files\Tesseract-OCR\tesseract.exe'


if __name__ == "__main__":
    setPytesseract()

    print("请输入识别要求，1：宗地代码，8：宗地代码+户主")
    index = input()
    print("开始识别，请稍候。。。")

    renlist = []
    mainPath = os.getcwd()
    for f in os.listdir(mainPath):
        if f.endswith('.pdf'):
            pdf = fitz.open(f)
            pages = pdf.page_count
            if pages > 8:
                record = []
                record.append("REN ")
                oldname = f
                record.append(oldname)
                if index == '1':
                    # 封面
                    jpgpath1 = pdf2jpg(pdf, 0)
                    record.append(ocr09(jpgpath1))
                    renlist.append(record)
                else:
                    # 封面
                    jpgpath1 = pdf2jpg(pdf, 0)
                    record.append(ocr09(jpgpath1))
                    # 审批表
                    huzhunum = (pages - 7)//2
                    jpgpath8 = pdf2jpg(pdf, huzhunum+6)
                    record.append(ocr09(jpgpath8))
                    renlist.append(record)
    writexls(renlist, index)
    if index == '1':
        input("请在当前路径查看表格：'签字表识别宗地代码.xls'")
    else:
        input("请在当前路径查看表格：'签字表识别宗地代码+户主.xls'")
