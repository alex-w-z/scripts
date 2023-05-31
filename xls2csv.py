# -*- coding: utf-8 -*-

import pandas as pd
import os
import argparse

# 读取文件夹下所有xls文件，返回文件路劲列表
def get_xls_list(path):
    xls_list = []
    for root, dirs, files in os.walk(path):
        for file in files:
            if os.path.splitext(file)[1] == '.xls':
                xls_list.append(os.path.join(root, file))
    return xls_list

# 读取xls文件，返回xls里的所有sheet，并将sheet转换为csv文件，保存在以xls文件名命名的文件夹下
def xls2csv(xls_path, output):

        # xls_path 为xls文件路径，output为输出文件夹路径， 从xls_path 中读取sheet，转换为csv文件，保存在output文件夹下
        # 获取xls文件名，判断在本地"csv"目录下是否有xls文件名命名的文件夹，没有则创建，有则跳过，最后返回csv文件夹路径赋值给csv_path
        df = pd.read_excel(xls_path, sheet_name=None)

        # 从xls_path中获取xls文件名(不包含后缀)，赋值给xls_name, xls_name与outtput拼接，作为之后sheet转换为csv文件的输出路径
        xls_name = os.path.basename(os.path.splitext(xls_path)[0])

        #去除xls_name中的空格
        xls_name = xls_name.replace(' ', '')

        xls_output = os.path.join(output, xls_name)
        if not os.path.exists(xls_output):
            os.mkdir(xls_output)
        for sheet in df:
            csv_file = os.path.join(xls_output, sheet.replace(' ', '') + '.csv')
            df[sheet].to_csv(csv_file, index=False, encoding='utf-8')






        # # 获取xls文件名，判断在本地"csv"目录下是否有xls文件名命名的文件夹，没有则创建，有则跳过，最后返回csv文件夹路径赋值给csv_path
        # xls_name = os.path.basename(os.path.abspath(xls))
        # # print(f"xls_name: {xls}")
        # csv_path = os.path.join('./csv', xls_name)
        # if not os.path.exists(csv_path):
        #     os.mkdir(csv_path)
        # # 修改 xls cmd 权限，已管理员身份运行
        # os.system('icacls ' + xls + ' /grant everyone:F')
        # # data = pd.read_excel(xls, sheet_name=None)
        # # for sheet in data:
        # #     csv_file = os.path.join(csv_path, sheet + '.csv') # type: ignore
        #     # data[sheet].to_csv(csv_file, index=False, encoding='utf-8')


if __name__ == "__main__":
    

    dir_path = 'D:\Desktop\worldcopmare\qa2_20230517'
    # 判断dir_path目录下是否有output文件夹，没有则创建
    output = os.path.join(dir_path, 'output')
    if not os.path.exists(output):
        os.mkdir(output)
    xls_list = get_xls_list(dir_path)
    for xls in xls_list:
        xls_path = xls.replace(os.sep, '/')
        xls2csv(xls_path, output)
        print(xls_path)

