import math
# Read *.xlsx file
import os

import matplotlib.pyplot as plt
from matplotlib.pyplot import MultipleLocator
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from scipy import stats
from scipy.stats import pearsonr

input_name = r'CKD_ADMA_dealed.xlsx'
output_name = r'CKD_result.xlsx'
try:
    t = pd.DataFrame(pd.read_excel(input_name))  # header = 1 表示从第一行开始
except FileNotFoundError:
    print("File not exist!")
    exit()

# paramter set

t_head = t.columns
# catagory = ['control', 'case']
catagory = ['control', 'cancer']
gender = [0, 1]
gender_CN = ['男', '女']

class XYplot:
    def __init__(self, t):
        self.t = t
        self.t_head = t.columns

    def gen_protein_class(self, count_protein, count_cata):
        '''
        :param gender: gender class 0 for male / 1 for female
        :param count_protein: what kind of protein you want
        :return: a list include control/tumor/cancer + 50/60/70
        '''
        r = []
        r.append(list(t.loc[(t['catagory'] == catagory[count_cata]), self.t_head[2]]))
        r.append(list(t.loc[(t['catagory'] == catagory[count_cata]), self.t_head[count_protein]]))
        return r

def deleteSheet(ExcelName, SheetName):
    try:
        wb = load_workbook(ExcelName)
        ws = wb[SheetName]
        wb.remove(ws)
        wb.save(ExcelName)
    except KeyError:
        print("sheet not exist！")
        exit()

def genXYplot():
    plot = XYplot(t)
    # 要画这么多张图
    for count_protein in range(4, len(t_head)):
        # 每张图的散点有这么多种
        rData = []
        for class_cata in range(len(catagory)):
            rData.append(plot.gen_protein_class(count_protein, class_cata))
        fig, ax = plt.subplots()
        plt.ylabel(t_head[count_protein], fontsize=14)
        control = ax.scatter(rData[0][0], rData[0][1], color='r', s=10)
        cancer = ax.scatter(rData[1][0], rData[1][1], color='g', s=10)
        ax.legend((control, cancer), (u'control', u'cancer'), loc=2)
        plt.savefig(str(t_head[count_protein]) + '.png')
        plt.clf()
        # plot.WriteSheet(rData, output_name, t_head[count_protein] + catagory[class_cata])
    # deleteSheet(output_name, 'Sheet1')
    # data = range(-20, 20)

genXYplot()