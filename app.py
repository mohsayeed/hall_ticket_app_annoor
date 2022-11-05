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
        data_dict = {
            'CLASS_4': None, 'FESS_4': None, 'BOOKS_4': None, 'OTHERS_4': None, 'TOTAL_4': None, 'CONTACT_4': None, 'PAY_4': None, 'FEES_NOTICE_4': None,
            'IDNO_4': None, 'STUDENT_NAME_4': None, 'FATHER_NAME_4': None, 'OLD_BALANCE_4': None, 'PVT_TUTION_4': None,

            'FATHERGUARDIAN NAME_1': None, 'CLASS_1': None, 'FESS_1': None, 'BOOKS_1': None, 'OTHERS_1': None, 'PAY_1': None, 'TOTAL_1': None, 'CONTACT_1': None,
            'FEES_NOTICE_1': None, 'STUDENT_NAME_1': None, 'IDNO_1': None, 'OLD_BALANCE_1': None, 'PVT_TUTION_1': None,

            'CONTACT_3': None, 'PAY_3': None, 'TOTAL_3': None, 'OTHERS_3': None, 'PVTTUTION_3': None, 'BOOKS_3': None, 'FESS_3': None, 'OLDBALANCE_3': None, 'CLASS_3': None,
            'IDNO_3': None, 'FATHER_NAME_3': None, 'STUDENTNAME_3': None, 'FEES_NOTICE_3': None,

            'FEES_NOTICE_2': None, 'STUDENT_NAME_2': None, 'FATHER_NAME_2': None, 'IDNO_2': None, 'CLASS_2': None, 'OLD_BALANCE_2': None, 'FESS_2': None,
            'BOOKS_2': None, 'PVT_TUTION_2': None, 'OTHERS_2': None, 'TOTAL_2': None, 'PAY_2': None, 'CONTACT_2': None


        }
        list_dict = getDict(data_xls, data_dict)
        reader = PdfReader("FEE_NOTICE_LST_V1.pdf")

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
    number_of_pages = length_sheet//4
    extra_sheet_values = length_sheet % 4
    count = 0
    for i in range(number_of_pages):
        dict_temp = dict.copy()
        dict_temp['FATHERGUARDIAN NAME_1'] = excel_sheet['Father/GuardianName'][count]
        dict_temp['CLASS_1'] = excel_sheet['Class'][count]
        dict_temp['FESS_1'] = excel_sheet['Fees'][count]
        dict_temp['BOOKS_1'] = excel_sheet['Books'][count]
        dict_temp['OTHERS_1'] = excel_sheet['Other'][count]
        dict_temp['CONTACT_1'] = excel_sheet['Contact'][count]
        dict_temp['PAY_1'] = (excel_sheet['Pay Amount']
                              [count].date().strftime("%d-%m-%Y"))+"        "
        dict_temp['TOTAL_1'] = excel_sheet['Total'][count]
        dict_temp['FEES_NOTICE_1'] = excel_sheet['Fee Notice Date'][count].date(
        ).strftime("%d-%m-%Y")
        dict_temp['STUDENT_NAME_1'] = excel_sheet['StudentName'][count]
        dict_temp['IDNO_1'] = excel_sheet['IDNo'][count]
        dict_temp['OLD_BALANCE_1'] = excel_sheet['Old Bal'][count]
        dict_temp['PVT_TUTION_1'] = excel_sheet['Pvt Tution'][count]
        count = count + 1

        dict_temp['FEES_NOTICE_2'] = excel_sheet['Fee Notice Date'][count].date(
        ).strftime("%d-%m-%Y")
        dict_temp['STUDENT_NAME_2'] = excel_sheet['StudentName'][count]
        dict_temp['FATHER_NAME_2'] = excel_sheet['Father/GuardianName'][count]
        dict_temp['IDNO_2'] = excel_sheet['IDNo'][count]
        dict_temp['CLASS_2'] = excel_sheet['Class'][count]
        dict_temp['OLD_BALANCE_2'] = excel_sheet['Old Bal'][count]
        dict_temp['FESS_2'] = excel_sheet['Fees'][count]
        dict_temp['BOOKS_2'] = excel_sheet['Books'][count]
        dict_temp['PVT_TUTION_2'] = excel_sheet['Pvt Tution'][count]
        dict_temp['OTHERS_2'] = excel_sheet['Other'][count]
        dict_temp['TOTAL_2'] = excel_sheet['Total'][count]
        dict_temp['PAY_2'] = (excel_sheet['Pay Amount']
                              [count].date().strftime("%d-%m-%Y"))+"        "
        dict_temp['CONTACT_2'] = excel_sheet['Contact'][count]
        count += 1

        dict_temp['CONTACT_3'] = excel_sheet['Contact'][count]
        dict_temp['PAY_3'] = (excel_sheet['Pay Amount']
                              [count].date().strftime("%d-%m-%Y"))+"        "
        dict_temp['TOTAL_3'] = excel_sheet['Total'][count]
        dict_temp['OTHERS_3'] = excel_sheet['Other'][count]
        dict_temp['PVTTUTION_3'] = excel_sheet['Pvt Tution'][count]
        dict_temp['BOOKS_3'] = excel_sheet['Books'][count]
        dict_temp['FESS_3'] = excel_sheet['Fees'][count]
        dict_temp['OLDBALANCE_3'] = excel_sheet['Old Bal'][count]
        dict_temp['CLASS_3'] = excel_sheet['Class'][count]
        dict_temp['IDNO_3'] = excel_sheet['IDNo'][count]
        dict_temp['FATHER_NAME_3'] = excel_sheet['Father/GuardianName'][count]
        dict_temp['STUDENTNAME_3'] = excel_sheet['StudentName'][count]
        dict_temp['FEES_NOTICE_3'] = excel_sheet['Fee Notice Date'][count].date(
        ).strftime("%d-%m-%Y")
        count += 1

        dict_temp['CLASS_4'] = excel_sheet['Class'][count]
        dict_temp['FESS_4'] = excel_sheet['Fees'][count]
        dict_temp['BOOKS_4'] = excel_sheet['Books'][count]
        dict_temp['OTHERS_4'] = excel_sheet['Other'][count]
        dict_temp['TOTAL_4'] = excel_sheet['Total'][count]
        dict_temp['CONTACT_4'] = excel_sheet['Contact'][count]
        dict_temp['PAY_4'] = (excel_sheet['Pay Amount']
                              [count].date().strftime("%d-%m-%Y"))+"        "
        dict_temp['FEES_NOTICE_4'] = excel_sheet['Fee Notice Date'][count].date(
        ).strftime("%d-%m-%Y")
        dict_temp['IDNO_4'] = excel_sheet['IDNo'][count]
        dict_temp['STUDENT_NAME_4'] = excel_sheet['StudentName'][count]
        dict_temp['FATHER_NAME_4'] = excel_sheet['Father/GuardianName'][count]
        dict_temp['OLD_BALANCE_4'] = excel_sheet['Old Bal'][count]
        dict_temp['PVT_TUTION_4'] = excel_sheet['Pvt Tution'][count]
        count += 1
        list_dict_res.append(dict_temp)
    if (extra_sheet_values != 0):
        dict_temp = dict.copy()
        dict_temp['FATHERGUARDIAN NAME_1'] = excel_sheet['Father/GuardianName'][count]
        dict_temp['CLASS_1'] = excel_sheet['Class'][count]
        dict_temp['FESS_1'] = excel_sheet['Fees'][count]
        dict_temp['BOOKS_1'] = excel_sheet['Books'][count]
        dict_temp['OTHERS_1'] = excel_sheet['Other'][count]
        dict_temp['CONTACT_1'] = excel_sheet['Contact'][count]
        dict_temp['PAY_1'] = (excel_sheet['Pay Amount']
                              [count].date().strftime("%d-%m-%Y"))+"        "
        dict_temp['TOTAL_1'] = excel_sheet['Total'][count]
        dict_temp['FEES_NOTICE_1'] = excel_sheet['Fee Notice Date'][count].date(
        ).strftime("%d-%m-%Y")
        dict_temp['STUDENT_NAME_1'] = excel_sheet['StudentName'][count]
        dict_temp['IDNO_1'] = excel_sheet['IDNo'][count]
        dict_temp['OLD_BALANCE_1'] = excel_sheet['Old Bal'][count]
        dict_temp['PVT_TUTION_1'] = excel_sheet['Pvt Tution'][count]
        count = count + 1
        if (count == length_sheet):
            list_dict_res.append(dict_temp)
            return list_dict_res
        dict_temp['FEES_NOTICE_2'] = excel_sheet['Fee Notice Date'][count].date(
        ).strftime("%d-%m-%Y")
        dict_temp['STUDENT_NAME_2'] = excel_sheet['StudentName'][count]
        dict_temp['FATHER_NAME_2'] = excel_sheet['Father/GuardianName'][count]
        dict_temp['IDNO_2'] = excel_sheet['IDNo'][count]
        dict_temp['CLASS_2'] = excel_sheet['Class'][count]
        dict_temp['OLD_BALANCE_2'] = excel_sheet['Old Bal'][count]
        dict_temp['FESS_2'] = excel_sheet['Fees'][count]
        dict_temp['BOOKS_2'] = excel_sheet['Books'][count]
        dict_temp['PVT_TUTION_2'] = excel_sheet['Pvt Tution'][count]
        dict_temp['OTHERS_2'] = excel_sheet['Other'][count]
        dict_temp['TOTAL_2'] = excel_sheet['Total'][count]
        dict_temp['PAY_2'] = (excel_sheet['Pay Amount']
                              [count].date().strftime("%d-%m-%Y"))+"        "
        dict_temp['CONTACT_2'] = excel_sheet['Contact'][count]
        count += 1
        if (count == length_sheet):
            list_dict_res.append(dict_temp)
            return list_dict_res
        dict_temp['CONTACT_3'] = excel_sheet['Contact'][count]
        dict_temp['PAY_3'] = (excel_sheet['Pay Amount']
                              [count].date().strftime("%d-%m-%Y"))+"        "
        dict_temp['TOTAL_3'] = excel_sheet['Total'][count]
        dict_temp['OTHERS_3'] = excel_sheet['Other'][count]
        dict_temp['PVTTUTION_3'] = excel_sheet['Pvt Tution'][count]
        dict_temp['BOOKS_3'] = excel_sheet['Books'][count]
        dict_temp['FESS_3'] = excel_sheet['Fees'][count]
        dict_temp['OLDBALANCE_3'] = excel_sheet['Old Bal'][count]
        dict_temp['CLASS_3'] = excel_sheet['Class'][count]
        dict_temp['IDNO_3'] = excel_sheet['IDNo'][count]
        dict_temp['FATHER_NAME_3'] = excel_sheet['Father/GuardianName'][count]
        dict_temp['STUDENTNAME_3'] = excel_sheet['StudentName'][count]
        dict_temp['FEES_NOTICE_3'] = excel_sheet['Fee Notice Date'][count].date(
        ).strftime("%d-%m-%Y")
        count += 1
        if (count == length_sheet):
            list_dict_res.append(dict_temp)
            return list_dict_res
    return list_dict_res





if __name__ == "__main__":
    app.run()
