from re import L
import pandas as pd
import matplotlib.pyplot as plt
import math
import os
import shutil

from math import isnan


import datetime
import seaborn as sns
import matplotlib.pyplot as plt
# from matplotlib import dates
import matplotlib.dates as mdates

from datetime import timedelta
import numpy as np
import plotly.graph_objects as go

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Pt
from pptx.util import Inches

from tqdm import tqdm

from openpyxl import load_workbook



long_dir = r'C:\Users\dean.huang\main\projects\Todd\Shift_Change_DT_Kalamazoo\shift_filter'
filtered_path = os.path.join(long_dir, "shift_filtered_data")




def save_to_excel(filename, data_dict):
    combined_df = pd.DataFrame.from_dict(data_dict)
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    combined_df.to_excel(writer, sheet_name='Monthly_Data')
    writer.save()

def save_csv(X, dest_path):
    X.to_csv(dest_path, index = False)

def dt2s(datetime_obj):
    time = datetime.time.strftime(datetime_obj,"%H:%M:%S.%f")
    seconds = int(time[:2])*60**2 + int(time[3:5])*60 + float(time[6:])
    return round(seconds,2)

def time_HHMM(in_string):
    return in_string[:in_string.find(':', 3)]
def time_MMSS(in_string):
    return in_string[in_string.find(':')+1:]


def wc_lookup(wc, lookup): #give me wc, I return Department and Shift Hours
    return lookup.iloc[lookup[lookup['Equipment Name'] == wc].index[0]].tolist()[1:]


def wc_unithr(wc, qv): #give me wc, I return Units/HR
    for index, row in qv.iterrows():
        if row['Work Center'] == wc: return round(row['Unit/Hr'], 0)

def time_check(sl, start_time):
    time = datetime.datetime.strftime(start_time.to_pydatetime(),"%H:%M:%S.%f")
    hours, mins= int(time[:2]), int(time[3:5])

    if sl == 12:
        if mins < 16:
            if (hours == 7 or hours == 19): return True
        if mins > 44:
            if (hours == 6 or hours == 18): return True
    elif sl == 8:
        if mins < 16:
            if (hours == 7 or hours == 15 or hours == 23): return True
        if mins > 44:
            if (hours == 6 or hours == 14 or hours == 22): return True

    return False


    


def dd_loader(raw_file):
    print("Starting %s filter time" % raw_file)
    #Load raw data
    raw_data = os.path.join(long_dir,"datadump",  raw_file)

    #Load the tables
    ig_table = pd.read_excel(raw_data, "ig")
    qv_table = pd.read_excel(raw_data, "qv")
    lookup_table = pd.read_excel(raw_data, "lookup")


    # Pass the Lookup and QV tables through, only process the ig_table
    # for index, row in tqdm(ig_table.head(5).iterrows()):

    shift_filtered_list = []
    for index, row in tqdm(ig_table.iterrows()):

        # Only grab the variables I need to calculate time zone
        wc = int(row['Equipment Name'])
        timestamp = row['Delta Time Stamp']
        shift_start = row['Fifteen Minute Interval']

        shift_length = wc_lookup(wc, lookup_table)[1]
        uph = wc_unithr(wc, qv_table)

        if time_check(shift_length, shift_start):
            shift_filtered_list.append(row.tolist())

    multi_sheet_excel_writer(raw_file, shift_filtered_list)



def multi_sheet_excel_writer(filename, shift_change_list):
    dd_path = os.path.join(long_dir, "datadump", filename)
    book = load_workbook(dd_path)
    
    # thresh_filename = filename.split('.')[0] + "_thresh"
    out_filename = filename.replace('.xlsx', '_filtered')
    print("Creating %s now..." % out_filename)

    col_names = ['Equipment Name', 'Line State Type', 'Delta Time Stamp', 'Shift Start Date', 'Operator', 'Fifteen Minute Interval', 
                'Shift', 'Shift Month of Year', 'Shift Week of Year', 'Shift Year', 'Line Downtime Reason']
    df10 = pd.DataFrame(shift_change_list, columns = col_names)
    # writer = pd.ExcelWriter('%s/%s.xlsx' % (thresh_path, thresh_filename), engine='xlsxwriter')
    writer = pd.ExcelWriter('%s/%s.xlsx' % (filtered_path, out_filename), engine='openpyxl')
    writer.book = book
    df10.to_excel(writer, sheet_name='filtered', index=False)
    writer.save()



def extract_int(in_string):
    try:
        if math.isnan(in_string):
            return in_string
    except:
        pass
    if type(in_string) == float or type(in_string) == int:
        return int(in_string)
    s2i = [str(i) for i in in_string if i.isdigit()]
    s2i = ''.join(s2i)
    return int(s2i)

def dt2d(datetime_obj):
    return datetime_obj.strftime('%Y-%m-%d %X')[:10]



def days_ago(n): #
    week_ago = datetime.datetime.now() - datetime.timedelta(days = n)
    return week_ago.year, week_ago.month, week_ago.day


def create_report():
    for file in os.listdir(os.path.join(long_dir, "datadump")):
        # if file.endswith('kal_apr_sep_dd.xlsx'):
        if file.endswith('.xlsx'):
            dd_loader(file)



def main():


    # create_report()

    net_path = r'C:\Users\dean.huang\OneDrive - Graphic Packaging International, LLC\shift_change'
    directory = 'test'

    path = os.path.join(net_path, directory)

    os.mkdir(path)










if __name__ == '__main__':
    main()

















'''
ddloader:
    # ig_date = []
    # month_list = []
    # df1 = pd.read_excel(dd_file)
    # month_list.append(df1)
    # combined = pd.concat(month_list, ignore_index=True)

    # qv_date = []
    # qv_table = pd.read_excel(dd_qv_path)

    # qv_date = qv_table['Date'].to_list()
    # ig_date = list(set(combined['Scheduled Shift Start Date Time'].to_list()))
    # qv_date = [dt2d(d2) for d2 in qv_date]
    # ig_date = [dt2d(d) for d in ig_date]
    # qv_date = list(set(qv_date))
    # ig_date = list(set(ig_date))


'''