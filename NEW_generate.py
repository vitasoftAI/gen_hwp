import pandas as pd
import openpyxl
import numpy as np
import random
from datetime import datetime
from random import randrange
from datetime import timedelta
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('-f','--file-path', type=str, default="VARIAB_2.xlsx", help='load file path or file name')
opt = parser.parse_args()

wb = openpyxl.load_workbook(opt.file_path)
sheet_name_list = wb.sheetnames


data_field = pd.read_excel(io=opt.file_path, sheet_name='VDPSpec')
data_format = pd.read_excel(io=opt.file_path, sheet_name='VDataSpec')

data_field_dict = {}
data_format_dict = {}
total_dict = {}
item_dict = {}
#필드 및 포멧형식 만들기
for i in range(len(data_field)):
    data_field_dict[data_field.iloc[i][0]] = data_field.iloc[i][1]
for j in range(len(data_format)):
    data_format_dict[data_format.iloc[j][0]] = data_format.iloc[j][1]
#필드 및 포멧형식 재구성
for key,val in data_field_dict.items():
    if len(val.split('+')) == 1:
        if val in data_format_dict.keys():
            total_dict[key] = data_format_dict[val]
    #여러개를 합친거 처리 필요
    else:
        item_lst = []
        for lst in val.split('+'):
            if lst in data_format_dict.keys():
                 item_lst.append(data_format_dict[lst])
        total_dict[key] = item_lst
    #         # print(len(val.split('+')))
    #         break

# 값변환


d1 = datetime.strptime('1/1/2000', '%m/%d/%Y')
d2 = datetime.strptime('6/9/2022', '%m/%d/%Y')

def random_date(start, end):
    delta = end - start
    int_delta = (delta.days)
    random_second = randrange(int_delta)

    return start + timedelta(days=random_second)

def number_change(num_data):
    while '#' in num_data:
        targte_value = random.randint(0, 9)
        num_data = num_data.replace('#', str(targte_value), 1)
    return num_data

def check_Data(value):
    str_data = ''
    indices = pd.read_excel(opt.file_path,sheet_name=value,header=None)
    if indices.values.flatten().tolist():
        str_data = random.choice(indices.values.flatten().tolist())
    return str_data

def edit_content(data, src, target):

    return data.replace(str(src).encode(), str(target).encode(), 1)

for key, value in total_dict.items():
    if isinstance(value,list):
        sum_data = ''
        for idx,item in enumerate(value):
            if item in sheet_name_list:
                str_data = check_Data(item)
                sum_data = sum_data + str_data
            if '#' in item :
                num_data = number_change(item)
                sum_data = sum_data + num_data
        total_dict[key] = sum_data
    if 'MM/DD/YYYY' in value:
        total_dict[key] = random_date(d1, d2).strftime("%m/%d/%Y")
    if value in sheet_name_list:
        str_data = check_Data(value)
        total_dict[key] = str_data
    if '#' in total_dict[key]:
        num_data = number_change(total_dict[key])
        total_dict[key] = num_data

with open('Form_NV_0001.hml', 'rb') as f:
    data = f.read()
for key, value in total_dict.items():
    data = edit_content(data, key, value.replace('&', '&amp;'))

with open(f'sample{str(1).zfill(4)}.hml', 'wb') as f:
    f.write(data)
