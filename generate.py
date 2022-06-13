import pandas as pd
from random import randrange
from datetime import timedelta

def random_date(start, end):
    delta = end - start
    int_delta = (delta.days )
    random_second = randrange(int_delta)

    return start + timedelta(days=random_second)

from datetime import datetime
d1 = datetime.strptime('1/1/2000', '%m/%d/%Y')
d2 = datetime.strptime('6/9/2022', '%m/%d/%Y')

def remove_zeros(data):
    return str(int(data))


import random
class Fields(object):
    def __init__(self, data_list, ):
        company, address, port, kind, country = data_list

        self.Invoice = '#####'
        self.InvoiceDate = random_date(d1, d2).strftime("%m/%d/%Y")
        self.phone_number = '(##)###-####'
        self.Cosignee = company + ' ' + address + ' ' + self.phone_number
        self.Buyer = self.Cosignee
        self.Port = port
        self.OrderNumber = '#######'
        self.CustomerPONumber = self.OrderNumber
        self.TermsOfDelivery = '## Days'
        self.MarksAndNumbers = 'DLSU#######'
        self.Pkgs = '##'
        self.TotalGrossWeight = '###'
        self.TotalCube = '##.##'
        self.PackageSpec = kind +' '+ self.TotalGrossWeight +' '+ self.TotalCube
        self.PartNumber = '##'
        self.HTS ='####.##.####'
        self.Quantity = '###'
        self.PriceEach = '$#,###.##'
        self.Value = '$#,###.##'
        self.InvoiceTotal = '$#,###.##'
        self.CountryOfOrigin = country
    def get_str(self):
        dict_data = self.__dict__
        remov_zeros_list = ['PartNumber', 'Quantity']
        for key, value in dict_data.items():
            while '#' in dict_data[key]:
                targte_value = random.randint(0, 9)
                dict_data[key] = dict_data[key].replace('#', str(targte_value), 1)

        for key in remov_zeros_list:
            dict_data[key] = remove_zeros(dict_data[key])
        return dict_data
    def get_value(self):
        return 'test'
import xmltodict
#
# def hwp_to_image(hwp_path):
#     filename = hwp_path
#     hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
#     hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
#     with open('Form_NV0_0525A_0001.hml', 'rb') as f:
#         data = f.read() #.find(b'#Number, kind, weight and dimensions of packages')#.replace(b'#Cosignee', b'test')
#     hwp.Open(hwp_path)
#
#     xml = hwp.GetTextFile("HWP", "")
#     print(xml)
#     exit()
#
#     with open(hwp_path.replace('hwp', 'xml'), 'w', encoding="UTF-8") as f:
#         f.write(xml)
#     with open(hwp_path.replace('hwp', 'xml'), 'r', encoding="UTF-8") as f:
#         data = f.read()
#     import json
#     jsonString = json.dumps(xmltodict.parse(data, process_namespaces=True),
#                             ensure_ascii=False,
#                             indent=4
#                             )
#     with open(hwp_path.replace("hwp", 'json'), 'w', encoding="UTF-8") as f:
#         f.write(jsonString)
#
#     xmlString = xmltodict.unparse(json.loads(jsonString)
#                                   # , pretty=True
#                                   , encoding='utf-16'  # 16 중요 안하면 한글에서 인식 못함
#                                   )
#
#     with open(hwp_path.replace('asdf.hwp', 'asdf_json.xml') , 'w', encoding="UTF-8") as f:
#         f.write(xmlString)
#     with open(hwp_path.replace('asdf.hwp', 'asdf_json.xml') , 'r', encoding="UTF-8") as f:
#         tt = f.read()
#     hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
#     hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
#     hwp.SetTextFile(tt, "HWPML2X", "")
#     hwp.SaveAs(hwp_path.replace('asdf', "rrrr"), "HWPML2X", "")
    # with open(filename.replace('hwp', 'xml'), 'w') as f:
    #     f.write(xml)


    # hwp.CreatePageImage("{}".format(filename.replace(".hwp", "")), 0, resolution=300,depth=24)



# hwp_to_image(os.path.abspath('dataset/asdf.hwp'))

def edit_content(data, src, target):

    return data.replace(str(src).encode(), str(target).encode(), 1)

key_value_pair = [('#Invoice Total(USD)', 'InvoiceTotal'),
                  ('#Invoice Date', 'InvoiceDate'),('#Invoice', 'Invoice') , ('#Cosignee', 'Cosignee'), ('#Origin of Shipment', 'Port'),
                  ('#Buyer', 'Buyer'),
                  ('#Sales Order No.', 'OrderNumber'), ('#Customer PO No.', 'CustomerPONumber'), ('#Terms of Sale and Delivery', 'TermsOfDelivery'),
                  ('#Marks and Numbers', 'MarksAndNumbers'), ('#Number, kind, weight and dimensions of packages', 'PackageSpec'),
                  ('# # of Pkgs', 'Pkgs'), ('#Total Gross Weight', 'TotalGrossWeight'), ('#Total Cube', 'TotalCube'), ('#Part number', 'PartNumber'),
                  ('#Countryof Origin', 'CountryOfOrigin'), ('#HTS', 'HTS'), ('#Quantity', 'Quantity'), ('#Price each', 'PriceEach'), ("#Value(USD)", "Value")]

indices = pd.read_excel('OCRDBFields_.xlsx',sheet_name='3rd')
with open('Form_NV0_0525A_0001-40.hml', 'rb') as f:
    data = f.read()  # .find(b'#Number, kind, weight and dimensions of packages')#.replace(b'#Cosignee', b'test')
for idx1, item in indices.iterrows():
    if idx1 <800:
        continue
    if idx1>1001:
        break

    item = list(item[indices.columns].values)
#     # 'company', 'address', 'kind', 'country', 'tech'
#     print(item)
    fields = Fields(item)

    target_dict = fields.get_str()
    # print(data.decode())
#
    for idx, item in enumerate(key_value_pair):
        data = edit_content(data, item[0], target_dict[item[1]].replace('&', '&amp;'))
    data = edit_content(data, '#Technical Description', "")
with open(f'sample{str(905).zfill(4)}.hml', 'wb') as f:
    f.write(data)
    # hwp_to_image(f'test{21+idx1}.hml')
#
# # with open('test1.hml', 'wb') as f:
# #     f.write(data)
# import xml.etree.ElementTree as ET
#
#
#
# with open("Form_NV0_0525A_0001.hml",'r', encoding="UTF-8") as f:
#     xml2 = f.read()
# import json
# tree = ET.parse("Form_NV0_0525A_0001.hml")
# root = tree.getroot()
# print(root)
# asdf = tree.iter("IMAGE")
# import base64
# imgdata = base64.b64decode(tree.getiterator('IMAGE')[0].text.encode('utf-16'))
# filename = 'some_image.png'  # I assume you have a way of picking unique filenames
# with open(filename, 'wb') as f:
#     f.write(imgdata)

# hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
# hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
# hwp.SetTextFile(xml2, "HWPML2X", "")
# hwp.SaveAs(r"C:\Users\sunci\Desktop\busi\hhh.hwp","HWPML2X","")

# hwp.Quit()
