from PyPDF2 import PdfFileReader, PdfFileWriter
import os 
from PIL import Image
import pytesseract
from pdf2image import convert_from_path
import re
import xlsxwriter
import argparse


def func(num, files, S):
    workbook = xlsxwriter.Workbook(f"{num}-{num+S}data_parsed.xlsx")
    worksheet = workbook.add_worksheet()

    row = 0

    worksheet.write(row, 0, "name")
    for i in range(1,100):
        worksheet.write(row, i+1, str(i))
    row += 1


    for f in files[num:num+S]:
        # print(f)
        if ".pdf" in f:
            # print(row)
            fff = f'{f}test.pdf'

            with open(os.path.join(FOLDER, f), 'rb') as infile:
                reader = PdfFileReader(infile)
                writer = PdfFileWriter()
                writer.addPage(reader.getPage(0))

                with open(fff, 'wb') as outfile:
                    writer.write(outfile)

            images = convert_from_path(fff)

            fff_png = fff+'.png'

            images[0].save(fff_png)
            # print("response text")
            text = pytesseract.image_to_string(Image.open(fff_png))
    
            text = text.replace("\n", ' ')
            text = text.replace("\r", ' ')

            with open("parsed.txt", "w", encoding="utf-8") as g:
                g.write(text)

            temp = re.split(pattern, text)[1:]
        

            worksheet.write(row, 0, f[:-4])

            while len(temp) > 0:
                try:
                    if len(temp) == 0:
                        break
                    
                    col = int(temp.pop(0).strip())
                    
                    if col == 57:
                        break
                    
                    part_text = temp.pop(0)

                    if "siehe" in part_text:
                        if len(temp) == 0:
                            break
                        part_text += temp.pop(0)

                    worksheet.write(row, col+1, part_text)
                except ValueError:
                    # print("end parse")
                    continue

            row += 1

            os.remove(fff)
            os.remove(fff_png)

    workbook.close()



if __name__ == '__main__':
    FOLDER = "."

    S = 500

    pattern =  r'\(([\d)]+)\)'

    files = os.listdir(FOLDER)

    for num in range(0, len(files), S):
        func(num, files, S)
