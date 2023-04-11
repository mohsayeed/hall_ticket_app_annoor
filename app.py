import os
from PyPDF2 import  PdfReader, PdfWriter
from flask import Flask, render_template, request, url_for
from flask import redirect
import pypdfium2 as pdfium
from PIL import Image
import os

import pandas as pd

app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # prepraing to write the
        f = request.files['file']

        # sheet evaluation
        data_xls = pd.read_excel(f)
        data_dict = {'exam_1': None, 'exam_2': None, 'exam_3': None, 'exam_4': None, 'exam_5': None, 'exam_6': None, 'exam_7': None, 'exam_8': None, 'exam_9': None, 'exam_10': None, 'sname_1': None, 'sname_2': None, 'sname_3': None, 'sname_4': None, 'sname_5': None, 'sname_6': None, 'sname_7': None, 'sname_8': None, 'sname_9': None, 'sname_10': None,
                     'fname_2': None, 'fname_3': None, 'fname_4': None, 'fname_5': None, 'fname_6': None, 'fname_7': None, 'fname_8': None, 'fname_9': None, 'fname_10': None, 'class_9': None, 'class_10': None, 'class_8': None, 'class_7': None, 'class_5': None, 'class_6': None, 'class_4': None, 'class_3': None, 'class_1': None, 'class_2': None, 'fname_1': None}
        list_dict = getDict(data_xls, data_dict)
        print(list_dict)
        reader = PdfReader("form.pdf")

        for i in range(len(list_dict)):
            writer = PdfWriter()
            writer.addPage(reader.getPage(0))
            writer.updatePageFormFieldValues(
            writer.getPage(0), list_dict[i])
            with open("temp/"+str(i)+".pdf", "wb") as s:
                writer.write(s)
            pdf = pdfium.PdfDocument("temp/"+str(i)+".pdf")
            n_pages = len(pdf)
            for page_number in range(n_pages):
                page = pdf.get_page(page_number)
                pil_image = page.render_topil(
                    scale=200/72, 
                    rotation=0,
                    crop=(0, 0, 0, 0),
                    greyscale=False,
                    optimise_mode=pdfium.OptimiseMode.NONE,
                )
                pil_image.save(f"temp/image"+str(i)+".png")

        img = []
        for i in range(len(list_dict)):
            path = "temp/image"+str(i)+".png"
            x = Image.open(path)
            im_1 = x.convert('RGB')
            img.append(im_1)
            x.close()

        img1st = img[0]
        img.pop(0)
        img1st.save(r'static/final_res.pdf',
                    save_all=True, append_images=img)            
        return redirect(url_for('static', filename='final_res.pdf'))
    return render_template('index.html')


@app.route('/delete')
def delete():
    temp_list = os.listdir('temp')
    for i in temp_list:
        os.unlink('temp/'+i)
    # os.chmod('temp', stat.S_IRWXU | stat.S_IRWXG | stat.S_IRWXO)
    # shutil.rmtree(path='temp')
    return redirect(url_for('upload_file'))

def getDict(excel_sheet, dict):
    list_dict_res = []
    length_sheet = len(excel_sheet)
    number_of_pages = length_sheet//10
    extra_sheet_values = length_sheet % 10
    count = 0
    for i in range(number_of_pages):
        dict_temp = dict.copy()
        for j in range(10):
            dict_temp['class_'+str(j+1)] = excel_sheet['Class'][count]
            dict_temp['exam_'+str(j+1)] = '  '+excel_sheet['Exam'][count]
            dict_temp['sname_' +
                      str(j+1)] = format_len(excel_sheet['Student Name'][count])
            dict_temp['fname_' +
                      str(j+1)] = format_len(excel_sheet['Father / Guardian'][count])
            count+=1
        list_dict_res.append(dict_temp)
    if (extra_sheet_values != 0):
        dict_temp = dict.copy()
        for l in range(10):
            dict_temp['class_'+str(l+1)] = excel_sheet['Class'][count]
            dict_temp['exam_'+str(l+1)] = '  '+excel_sheet['Exam'][count]
            dict_temp['sname_' +
                      str(l+1)] = format_len(excel_sheet['Student Name'][count])
            dict_temp['fname_' +
                      str(l+1)] = format_len(excel_sheet['Father / Guardian'][count])
            count += 1
            if (count == length_sheet):
                list_dict_res.append(dict_temp)
                return list_dict_res
    return list_dict_res

def format_len(s):
    s = ' '+s
    if(len(s)<30):
        t = 30-len(s)
        for i in range(t):
            s = s+' '
    return s



if __name__ == "__main__":
    app.run()
