# coding:utf-8
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Color, Font, Alignment
from openpyxl.styles.colors import BLUE, RED, GREEN, YELLOW
from flask import Flask
from flask import abort
from flask import redirect
from flask import Blueprint,render_template,send_file
from flask import Flask,render_template,request

app = Flask(__name__)

font = Font(name=u'宋体', size=10, color=RED, bold=False)


def read_excel(file_name):
    wb2 = load_workbook(file_name)
    sheet_list = wb2.get_sheet_names()
    ws = wb2.get_sheet_by_name(sheet_list[0])
    data = {}
    for row in ws.iter_rows():
        for cell in row:
            data.update({cell.coordinate: cell.value})
    return data


def data_compare(file_model, file_new):
    model_dict = read_excel(file_model)
    new_dict = read_excel(file_new)
    difference_key = []
    for k, v in new_dict.items():
        if model_dict.get(k) != v:
            difference_key.append(k)
    return difference_key


def sign_red(file_name, keys):
    wb2 = load_workbook(file_name)
    sheet_list = wb2.get_sheet_names()
    ws = wb2.get_sheet_by_name(sheet_list[0])
    for key in keys:
        cell = ws.cell(key)
        cell.font = font
    wb2.save(file_name)


def main(file_model,file_new):
    difference_key = data_compare(file_model, file_new)
    sign_red(file_new, keys=difference_key)


@app.route('/')
def index():
    return send_file('index.html')


@app.route('/compare_excel',methods=['POST'])
def main_route():
    file1 = request.args.get('file1')
    file2 = request.args.get('file2')
    main(file1,file2)
    return True


if __name__ == '__main__':
    app.run(debug=True)
