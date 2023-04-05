
import pandas as pd
from datetime import date
import csv
import re


class data_capture():

    def __init__(self, stock_nimber0):
        # initialization for the object

        # save the stock number
        self.stock_number = str(stock_nimber0)

        pass

    def reorder(self, stock_id, stock_name, date):
        datestr = date.strftime('%Y%m%d')
        datestr1 = date.strftime('%Y/%m/%d')
        yearstr = date.strftime('%Y')
        monstr = date.strftime('%m')
        inputfilename = 'c:\\py_gary\\py_code\\dailydata\\' + \
            stock_id + '_' + datestr + '.csv'
        outputfilename = 'c:\\py_gary\\py_code\\dailydata\\' + \
            stock_id + 'reorder_' + datestr + '.csv'
        print(inputfilename)
        # print(datestr1)
        header = ["日期", "證券代號", "證券名稱", "序號", "券商", "價格", "買進股數", "賣出股數"]

        csvFileToRead = open(inputfilename, 'r', encoding='big5')
        # csvFileToRead = open('3034 _20230224.csv', 'r', encoding='big5')
        csvDataToRead = csv.reader(csvFileToRead)
        next(csvDataToRead)
        next(csvDataToRead)
        dataList = list(csvDataToRead)
        csvFileToRead.close()
        total_lines = len(dataList)
        # print(total_lines)
        with open(outputfilename, 'w', newline='', encoding='utf-8') as csvFileToWrite:
            writer = csv.writer(csvFileToWrite, delimiter=',')
            writer.writerow(header)
            for cnt in range(1, total_lines):
                dataList_new1 = dataList[cnt][0:5]
                dataList_new2 = dataList[cnt][6:11]
                dataList_new1[1] = dataList_new1[1].replace("　", "")
                dataList_new1[1] = dataList_new1[1].replace(" ", "")
                dataList_new2[1] = dataList_new2[1].replace("　", "")
                dataList_new2[1] = dataList_new2[1].replace(" ", "")
                dataList_new2[4] = dataList_new2[4].strip()
                dataList_new1.insert(0, datestr1)
                dataList_new2.insert(0, datestr1)
                # dataList_new1.insert(1, "3034")
                # dataList_new2.insert(1, "3034")
                dataList_new1.insert(1, stock_id)
                dataList_new2.insert(1, stock_id)
                dataList_new1.insert(2, stock_name)
                dataList_new2.insert(2, stock_name)
                dataList_new1[4] = re.sub(
                    '[\u4e00-\u9fa5]', '', dataList_new1[4])
                dataList_new2[4] = re.sub(
                    '[\u4e00-\u9fa5]', '', dataList_new2[4])
                # print(dataList_new1)
                # print(dataList_new2)
                writer.writerow(dataList_new1)
                if (len(dataList_new2[3]) == 0):
                    continue
                else:
                    writer.writerow(dataList_new2)
        csvFileToWrite.close()


if __name__ == '__main__':

    '''
    here is to place the example code if this file is the main of excution
    '''

    d = date.fromordinal(730920)

    stock1 = data_capture('3034')
    stock1.reorder(stock1.stock_number, 'novatek', d)

    pass
