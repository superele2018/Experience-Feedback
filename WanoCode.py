from openpyxl import Workbook
from openpyxl import load_workbook
import time
from openpyxl.chart import BarChart, Series, Reference
from pylab import *
from matplotlib.ticker import MultipleLocator
from matplotlib.ticker import FormatStrFormatter
import os
import re

class WanoCode(object):

    len_code = {}
    code = {}
    rb=None

    def __init__(self, file_code):
        self.file_code = file_code
        self.rb = load_workbook(file_code)
        for sheet in self.rb:
            # code.update({sheet.title:[]})
            child_code={}
            for row in list(sheet.rows):
                child_code.update({str(list(row)[0].value):list(row)[1].value})
            self.code.update({sheet.title:child_code})
            self.len_code.update({sheet.title:len(child_code)})

    def code_Kind_level1(self,code_str,area):
        kind=[]
        list_code_event=re.split('[,，/]+',code_str) #code_str.split(',')
        for i_code in list_code_event:
            if i_code.isdigit():
                if str(int(i_code)) in self.code[area].keys():
                    kind.append(divmod(int(i_code),int(100))[0]*100)
                else:
                    kind.append('other')
            else:
                kind.append('other')
        return list(set(kind))







