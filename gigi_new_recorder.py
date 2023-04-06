
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

    def reorder_pd(self, stock_id, stock_name, date):
        datestr = date.strftime('%Y%m%d')
        datestr1 = date.strftime('%Y/%m/%d')
        yearstr = date.strftime('%Y')
        monstr = date.strftime('%m')
        inputfilename = 'c:\\py_gary\\py_code\\dailydata\\' + '_' + datestr + '.csv'
        outputfilename = 'c:\\py_gary\\py_code\\dailydata\\' + '_' + datestr + '.csv.gz'
        # outputfilename = '../reorderdata/'+ stock_id + 'reorder_' + datestr + '.csv'
        print(inputfilename)

        # new_header = ["日期", "證券代號", "證券名稱", "序號", "券商", "價格", "買進股數", "賣出股數"]
        # old_header = ["序號", "券商", "價格", "買進股數", "賣出股數"]
        df_raw = pd.read_csv(inputfilename, delimiter=',',
                             skiprows=2, encoding='big5')
        # print(df_raw)
        df_raw1 = df_raw.iloc[:, :5]
        df_raw2 = df_raw.iloc[:, 6:]
        # print(df_raw1)
        # print(df_raw2)
        df_raw2.rename(columns={'序號.1': '序號', '券商.1': '券商', '價格.1': '價格',
                       '買進股數.1': '買進股數', '賣出股數.1': '賣出股數'}, inplace=True)
        # print(df_raw2)
        df_raw3 = pd.concat([df_raw1, df_raw2], ignore_index=True)
        # print(df_raw3)
        df_raw3.dropna(axis=0, inplace=True)
        # print(df_raw3)
        df_raw3.sort_values(by=["序號"], ascending=True,
                            inplace=True, ignore_index=True)
        df_raw3['序號'] = df_raw3['序號'].astype('uint32')
        # print(df_raw3)
        df_raw3.insert(0, "日期", datestr1)
        df_raw3.insert(1, "證券代號", stock_id)
        df_raw3.insert(1, "證券名稱", stock_name)
        df_raw3.to_csv(outputfilename, index=False, compression="gzip")


if __name__ == '__main__':

    '''
    here is to place the example code if this file is the main of excution
    '''

    d = date.fromordinal(730920)

    stock1 = data_capture('3034')
    stock1.reorder_pd(stock1.stock_number, 'novatek', d)

    pass
