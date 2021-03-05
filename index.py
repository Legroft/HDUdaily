from pywebio.platform.flask import webio_view
from flask import Flask, send_from_directory

from pywebio import start_server, set_env
from pywebio.input import *
from pywebio.output import *
from openpyxl import load_workbook


def check_token(p):  # 检验函数校验通过时返回None，否则返回错误消息
    if len(p) != 36:
        return 'token格式错误！'


def main():
    set_env(title='自动打卡', output_animation=False)
    put_markdown(r""" # 杭电自动健康打卡系统
    本系统尚不完善，存在许多bug,一切以实际效果为准.
    使用说明见[自动打卡使用说明](https://jinjis.cn/index.php/archives/hduzi-dong-jian-kang-da-ka-xi-tong.html)
    """,lstrip=True)

    data = input_group("信息", [
        input("请输入token", help_text='token获取方式见底部', type=TEXT, validate=check_token, name='token'),
        input("请输入省份", help_text='城市获取方式见底部 默认为浙江省杭州市江干区', type=NUMBER, value='7', name='p'),
        input("请输入城市", type=NUMBER, value='9', name='c'),
        input("请输入地区", type=NUMBER, value='4', name='a'),
        input("Qmsg酱推送KEY", type=TEXT, help_text='推送说明见底部 不推送请填0', name='qq'),
    ])

    list = []
    for value in data.values():
        list.append(value)
    wb = load_workbook('data.xlsx')
    ws = wb.active
    ws.append(list)
    wb.save('data.xlsx')
    popup('提示', '信息保存成功！请不要重复提交信息 有问题请联系管理员')


app = Flask(__name__)

# task_func 为使用PyWebIO编写的任务函数
app.add_url_rule('/tool', 'webio_view', webio_view(main),
            methods=['GET', 'POST', 'OPTIONS'])  # 接口需要能接收GET、POST和OPTIONS请求

app.run(host='127.0.0.1', port=8080)