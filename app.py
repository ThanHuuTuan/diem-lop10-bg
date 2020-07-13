from flask import Flask
from flask import Flask, render_template, flash, request
import os
from openpyxl import Workbook
import requests
import json
from datetime import datetime
from flask import jsonify
from flask import Flask, send_file
from flask import Flask, render_template, flash, request
from wtforms import Form, TextField, TextAreaField, validators, StringField, SubmitField

app = Flask(__name__)
app.config.from_object(__name__)
app.config['SECRET_KEY'] = '7d441f27d441f27567d441f2b6176a'


class ReusableForm(Form):
    matruong = TextField('matruong:', validators=[validators.required()])
    sothisinh = TextField('sothisinh:', validators=[validators.required()])

@app.route('/download/<string:filename>')
def download_file(filename):
    path = "file_diem/{}".format(filename)
    return send_file(path, as_attachment=True)

@app.route('/danhsach')
def list_file():
    path = os.getcwd() + "/file_diem"
    list_of_files = {}

    for filename in os.listdir(path):
        list_of_files[filename] =  filename
    return json(list_of_files)

@app.route("/", methods=['GET', 'POST'])
def getdiem():
    form = ReusableForm(request.form)
    if request.method == 'POST':
        matruong = request.form['matruong']
        sothisinh = request.form['sothisinh']

        if form.validate():
            wb = Workbook()
            ws = wb.active
            data = [
                ["STT", "SBD", "HỌ VÀ TÊN", "NGÀY SINH", "UT", "KK", "VĂN", "ANH", "TOÁN", "TỔNG", "TỔNG UTK"]
            ]
            ma_truong = str(matruong)
            max_std = int(sothisinh)
            for i in range(max_std):
                try:
                    rr = requests.get(
                        'http://bacgiang.edu.vn/ssearch.ashx?iid=516&q=' + ma_truong + (
                                    '%03d' % (i + 1)) + '&pl=10').content
                    r_decode = rr.decode("utf-8")
                    print(r_decode)
                    data_r = json.loads(r_decode)['rs'][0]['r']
                    stt = i + 1
                    sbd = str(data_r[0])
                    name = str(data_r[1])
                    dob = str(data_r[2])
                    ut = str(data_r[3]).replace(",", ".")
                    kk = str(data_r[4]).replace(",", ".")
                    van = str(data_r[5]).replace(",", ".")
                    anh = str(data_r[6]).replace(",", ".")
                    toan = str(data_r[7]).replace(",", ".")
                    tong = float(toan) * 2 + float(van) * 2 + float(anh)
                    total = float(toan) * 2 + float(van) * 2 + float(anh) + float(ut) + float(kk)
                    new_data = [stt, sbd, name, dob, ut, kk, van, toan, anh, tong, total]
                    # print(new_data)
                    data.append(new_data)
                except:
                    new_data = [i + 1, "", "", "", "", "", "", "", "", "", ""]
                    data.append(new_data)

            for r in data:
                ws.append(r)

            ws.auto_filter.ref = "A:K"
            ws.column_dimensions['B'].width = 8
            ws.column_dimensions['C'].width = 30
            ws.column_dimensions['D'].width = 18
            ws.column_dimensions['K'].width = 13
            now = datetime.now()
            lte_time = now.strftime("%d-%m-%Y-%H_%M")
            filename = "file_diem/{}_{}.xlsx".format(ma_truong, lte_time)
            wb.save(filename)
            if os.path.exists(filename):
                file_create = "{}_{}.xlsx".format(ma_truong, lte_time)
                flash(file_create)
            else:
                flash('Error: Thất bại')

        else:
            flash('Error: Không được để trống các trường ')
    else:
        flash('Error: Nhập vào các thông tin bên trên để download danh sách điểm thi vào lớp 10')

    return render_template('index.html', form=form)

if __name__ == '__main__':
    app.run()
