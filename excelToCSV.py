# what this script will try and accomplish:
# import excel file
# read certain cells in file
# create csv based off this information
import os
import pandas as pd
from openpyxl import load_workbook
import json
import csv

class excelToCSV():

    def __init__(self,filepath):
        self.filepath = filepath
        self.file_exists = self.isFileThere()
        self.file_info = self.fileInfo()

    def fileInfo(self,filepath=""):
        if filepath == "":
            filepath = self.filepath
        file_info = os.path.splitext(filepath)
        return file_info

    def isFileThere(self, filepath = ""):
        if filepath == "":
            filepath = self.filepath
        file_exists = os.path.isfile(filepath);
        return file_exists

    def readSheet(self, excelOBJ, sheet, range):
        dataset = []
        sheet_ranges = excelOBJ[sheet]
        for cells in sheet_ranges[range]:
            for cell in cells:
                dataset.append(cell.value)
        return dataset

    def readExcel(self, filepath = ""):
        if filepath == "":
            filepath = self.filepath
        if self.file_exists == True:
            json_str = '{'
            i = 0
            wb = load_workbook(self.filepath)
            # grab date
            date = self.readSheet(wb,'MAIN','B2:B2')
            parse_date = date[0].strftime("%Y-%m-%d")
            # grab county
            county = self.readSheet(wb,'MAIN','A4:A118')
            # grab initial claims
            ic = self.readSheet(wb,'IC','C4:C118')
            # grab rate
            rate = self.readSheet(wb,'TUR','D5:D119')
            # grab weeks claimed
            wc = self.readSheet(wb,'IC','D4:D118')
            # grab benefits paid
            paid = self.readSheet(wb,'IC','E4:E118')

            while i < len(ic):
                json_str += '"info{}":'.format(i)
                json_str += '{'
                json_str += '"date":"{}","county":"{}","ic":"{}","rate":"{}","wc":"{}","paid":"{}"'.format(parse_date,county[i],ic[i],rate[i],wc[i],paid[i])
                if i == len(ic)-1:
                    json_str += '}'
                else:
                    json_str += '},'


                i += 1
            json_str += '}'
            # parse json
            json_dict = json.loads(json_str)

            return json_dict

    def createCSV(self,dict=""):
        with open('import.csv',mode='w') as csv_file:
            fieldnames = ['date','county','ic','rate','wc','paid']
            writer = csv.DictWriter(csv_file, fieldnames = fieldnames)
            writer.writeheader()
            for key,value in dict.items():
                 writer.writerow(value)



test = excelToCSV("C:/Users/michew2/Desktop/Python/import.xlsx")
data_dict = test.readExcel()
write_csv = test.createCSV(data_dict)
