# coding=utf-8
import PyPDF2
import os
import shutil


def pdf_1_7(huzhu):
    # if pages == 9:
    for i in range(0, 6 + huzhu):
        output17.addPage(input1.getPage(i))

    newpdf = os.path.join(root_path, newfile[:-4], newfile[:-4] + "#不动产权籍调查表.pdf")
    outputStream = open(newpdf, "wb")
    output17.write(outputStream)


def pdf_8(huzhu):
    output8.addPage(input1.getPage(6 + huzhu))

    newpdf = os.path.join(root_path, newfile[:-4], newfile[:-4] + "#农村宅基地及房屋补办审批表.pdf")
    outputStream = open(newpdf, "wb")
    output8.write(outputStream)


def pdf_9(huzhu, pages0):
    for i in range(7 + huzhu, pages0):
        output9.addPage(input1.getPage(i))

    newpdf = os.path.join(root_path, newfile[:-4], newfile[:-4] + "#农村宅基地实地调查申请表.pdf")
    outputStream = open(newpdf, "wb")
    output9.write(outputStream)


root_path = os.getcwd()
dijiNum = input("请输入地籍编码，例如：430426105247" + "\n")
for file in os.listdir(root_path):
    if file.endswith(".pdf"):
        # 重命名
        newfile = dijiNum + "JC" + file
        os.rename(file, newfile)

        # 归档入文件夹
        oldpath = os.path.join(root_path, newfile)
        os.mkdir(newfile[:-4])
        newpath = os.path.join(root_path, newfile[:-4], newfile)
        shutil.move(oldpath, newpath)

        # 分割pdf
        output17 = PyPDF2.PdfFileWriter()
        output8 = PyPDF2.PdfFileWriter()
        output9 = PyPDF2.PdfFileWriter()
        input1 = PyPDF2.PdfFileReader(open(newpath, "rb"))
        pages = input1.getNumPages()
        huzhuNum = (pages - 7) // 2
        pdf_1_7(huzhuNum)
        pdf_8(huzhuNum)
        pdf_9(huzhuNum, pages)
