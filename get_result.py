# coding=utf-8
import argparse
import os
import shutil

from flask import Flask, request, jsonify, send_from_directory,send_file
import urllib
from tools.infer.predict_system import *
import xlwt
import xlrd   # 1.2.0版本
import re
from flask import render_template
app = Flask(__name__)


@app.route("/")
def index():
    return render_template('upload.html')

@app.route("/post", methods=["POST"])
def get_sum():

    # 首先删除上次检测的结果文件
    if os.path.exists('./inference_results'):
        file_list = os.listdir('./inference_results')
        for f in file_list:
            file_path = os.path.join('./inference_results/', f)
            if os.path.isfile(file_path):
                os.remove(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path, True)

    # 删除上次保存的检测图片
    if os.path.exists('./data'):
        file_list = os.listdir('./data')
        for f in file_list:
            file_path = os.path.join('./data/', f)
            if os.path.isfile(file_path):
                os.remove(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path, True)

    file = request.files['imgfile']
    print(type(file))
    file_name = file.filename
    print(file_name)
    file_name = file_name.split('\\')[-1]
    print(file_name)
    file.save('./data/%s' % file_name)

    # predict_system部分开始
    args = utility.parse_args()
    if args.use_mp:
        p_list = []
        total_process_num = args.total_process_num
        for process_id in range(total_process_num):
            cmd = [sys.executable, "-u"] + sys.argv + [
                "--process_id={}".format(process_id),
                "--use_mp={}".format(False)
            ]
            p = subprocess.Popen(cmd, stdout=sys.stdout, stderr=sys.stdout)
            p_list.append(p)
        for p in p_list:
            p.wait()
    else:
        main(args)
        # 从excel获取检测项目名称
        key_words = []  # 存检验项目名称
        worksheet = xlrd.open_workbook(r'./det_name_list/det_name_list.xlsx')
        sheet_names = worksheet.sheet_names()
        sheet = worksheet.sheet_by_name(sheet_names[0])  # 获取sheet对象
        cols_value = sheet.col_values(0)  # 获取第一列内容， 数据格式为此数据的原有格式（原：字符串，读取：字符串；  原：浮点数， 读取：浮点数）
        key_words = cols_value[1:]  # 获取检验项目名称
        print(key_words)

        # 获取ocr识别结果
        with open(r'./inference_results/system_results.txt',
                  'r',
                  encoding='utf-8') as f:
            result = f.read()
            result = result.split("\t")[1]
            result = json.loads(result)  # 转成list格式
            # print(result)
            result = [item['transcription'] for item in result]  # 取得所有检测框的内容
            # 最终写入excel的结果列表
            # 获取标本种类
            for item in result:
                if '标本种类' in item:
                    sample = item.split('：')[-1]
                    break
                else:
                    sample = '未知标本'

            for item in result:
                if '报告日期' in item:
                    match = re.findall('\d{4}-\d{2}-\d{2}', item)
                    date = match[-1]
                    break
                else:
                    date = '未知报告日期'

            # 获取关键字结果，存储在ans里
            ans = []
            it = iter(result)
            while True:
                try:
                    item = next(it)
                    # print(item)
                    if item in key_words:
                        for i in range(3):
                            ans.append(item)
                            item = next(it)
                        ans.append(item)
                except StopIteration:
                    break
            # print(ans)

            # 写入excel
            # 创建结果文件
            result_file = r'.\inference_results\final_result.xls'
            # 打开工作表
            wb = xlwt.Workbook()
            # 添加子sheet
            result_sheet = wb.add_sheet('ocr检测结果', cell_overwrite_ok=True)
            result_sheet.write(0, 0, '标本类型名称')
            result_sheet.write(0, 1, '报告日期')
            result_sheet.write(0, 2, '检验项目名称')
            result_sheet.write(0, 3, '检验结果值')
            result_sheet.write(0, 4, '检验结果参考值')
            result_sheet.write(0, 5, '检验结果单位')
            row = 1
            it2 = iter(ans)
            while True:
                try:
                    for col in range(4):
                        item = next(it2)
                        result_sheet.write(row, col + 2, item)  # 检验结果
                    result_sheet.write(row, 0, sample)  # 标本种类
                    result_sheet.write(row, 1, date)  # 报告日期
                    row = row + 1
                except StopIteration:
                    break
            wb.save(result_file)

    return send_file('./inference_results/final_result.xls')
    # info['name'] = "中文"
    # info["age"] = 8928
    # return jsonify(info)  # 将json格式数据转换成json字符串
    # return '检测完成'
    # return app.send_static_file('./inference_results/final_result.xls')



if __name__ == "__main__":

    app.config["JSON_AS_ASCII"] = False  # 是否以ascii编码展示响应报文
    # server = pywsgi.WSGIServer(('0.0.0.0', 8080), app)
    # server.serve_forever()
    # app.run('127.0.0.1', 8080)
    app.run(host='0.0.0.0', port=8080)
