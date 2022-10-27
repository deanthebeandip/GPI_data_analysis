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


# Declaring the path for folder
long_dir = r'C:\Users\dean.huang\main\projects\Joel\Joel_Kal_Compare_Base_Anomaly\thresh_calc'
thresh_path = os.path.join(long_dir, "thresh_files")
shared_dir = r'C:\Users\dean.huang\OneDrive - Graphic Packaging International, LLC\anomaly_shared\dashboard'
shared_path = os.path.join(shared_dir, "thresh_files")

# Static number used for later
thresh_const = 10**6

# Static Watch List
static_watch_list = ['PLG5 Feeder Trips   Operator', 'CIM5 Side Lay Front Lay', 'PQBW Blanket Wash Normal']



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

def print_dict(d, indent=0):
    for key, value in d.items():
        print('\t' * indent + str(key))
        if isinstance(value, dict):
            print_dict(value, indent+1)
        else:
            print('\t' * (indent+1) + str(value))




def dd_loader(raw_file):

    #Load raw data
    raw_data = os.path.join(long_dir,"datadump",  raw_file)
    # Dict Initializations Center

    # Date - WC - Qty
    qv_date_dict = {}

    # DT - WC - Date - Day Value
    dt_orig_t_dict = {}
    dt_orig_f_dict = {}
    dt_op_t_dict = {}
    dt_op_f_dict = {}

    # DT - WC - (thresh t, thresh f)
    thresh_og_dict = {}
    thresh_op_dict = {}

    # DT - WC - thresh
    thresh_og_t_dict = {}
    thresh_og_f_dict = {}
    thresh_op_t_dict = {}
    thresh_op_f_dict = {}

    # Date - DT - WC - Day Value
    # Want a list that shows DT-WC-Dates-Values
    # Basically a history book that shows every single day of each DT-WC


    # CHANGE VALUE
    thresh_perc = 80


    combined = pd.read_excel(raw_data, "ig")
    qv_table = pd.read_excel(raw_data, "qv")

    #Populate qv_date_dict so I can lookup the Date by key, then use the WC as the sub key to hold Qty 
    for index, row in tqdm(qv_table.iterrows()):
        curr_date = dt2d(row['Date'])
        if curr_date not in qv_date_dict:
            qv_date_dict[curr_date] = {}
        if row['Work Center'] not in qv_date_dict[curr_date]:
            qv_date_dict[curr_date][row['Work Center']] = int(row['Confirmed Qty'])



    for index, row in tqdm(combined.iterrows()):
        ig_date = dt2d(row['Scheduled Shift Start Date Time'])
        if row['Line Downtime Original Reason'] not in dt_orig_t_dict: dt_orig_t_dict[row['Line Downtime Original Reason']] = {}
        if row['Line Downtime Reason'] not in dt_op_t_dict: dt_op_t_dict[row['Line Downtime Reason']] = {}
        if row['Line Downtime Original Reason'] not in dt_orig_f_dict: dt_orig_f_dict[row['Line Downtime Original Reason']] = {}
        if row['Line Downtime Reason'] not in dt_op_f_dict: dt_op_f_dict[row['Line Downtime Reason']] = {}

        # now go fill the dt-wc list, if doesn't exist, create it
        wc_num = extract_int(row['Line Downtime Equipment Name'])

        # Fill in WC key for current 
        if wc_num not in dt_orig_t_dict[row['Line Downtime Original Reason']]:dt_orig_t_dict[row['Line Downtime Original Reason']][wc_num] = {}
        if wc_num not in dt_orig_f_dict[row['Line Downtime Original Reason']]:dt_orig_f_dict[row['Line Downtime Original Reason']][wc_num] = {}
        if wc_num not in dt_op_t_dict[row['Line Downtime Reason']]:dt_op_t_dict[row['Line Downtime Reason']][wc_num] = {}
        if wc_num not in dt_op_f_dict[row['Line Downtime Reason']]:dt_op_f_dict[row['Line Downtime Reason']][wc_num] = {}
        if ig_date not in dt_orig_t_dict[row['Line Downtime Original Reason']][wc_num]:dt_orig_t_dict[row['Line Downtime Original Reason']][wc_num][ig_date] = 0
        if ig_date not in dt_orig_f_dict[row['Line Downtime Original Reason']][wc_num]:dt_orig_f_dict[row['Line Downtime Original Reason']][wc_num][ig_date] = 0
        if ig_date not in dt_op_t_dict[row['Line Downtime Reason']][wc_num]:dt_op_t_dict[row['Line Downtime Reason']][wc_num][ig_date] = 0
        if ig_date not in dt_op_f_dict[row['Line Downtime Reason']][wc_num]:dt_op_f_dict[row['Line Downtime Reason']][wc_num][ig_date] = 0

        # Each DT-WC-Date combo now has a time and freq, for both og and op 
        dt_orig_t_dict[row['Line Downtime Original Reason']][wc_num][ig_date] += row['Line State Duration']
        dt_op_t_dict[row['Line Downtime Reason']][wc_num][ig_date] += row['Line State Duration']
        dt_orig_f_dict[row['Line Downtime Original Reason']][wc_num][ig_date] += 1
        dt_op_f_dict[row['Line Downtime Reason']][wc_num][ig_date] += 1



    # Remove the nan keys, so they don't screw up my data
    for key in list(dt_orig_t_dict):
        if key != key: dt_orig_t_dict.pop(key)
    for key in list(dt_orig_f_dict):
        if key != key: dt_orig_f_dict.pop(key)
    for key in list(dt_op_t_dict):
        if key != key: dt_op_t_dict.pop(key)
    for key in list(dt_op_f_dict):
        if key != key: dt_op_f_dict.pop(key)


    # Thresh OG Dict populate zone

    min_list = 3
    for dt in dt_orig_f_dict:
        if dt not in thresh_og_dict: #populate OG DT in thresh dict
            thresh_og_dict[dt] = {}
        for wc in dt_orig_f_dict[dt]:
            if wc not in thresh_og_dict[dt]:
                thresh_og_dict[dt][wc] = []
            dt_wc_t = []
            dt_wc_f = []
            # print(type(wc))
            if type(wc) == int:
                for day in dt_orig_t_dict[dt][wc]:
                    try:
                        # print(dt, wc, dt_orig_f_dict[dt][wc][day], qv_date_dict[day][wc])
                        normal_value = thresh_const * dt_orig_t_dict[dt][wc][day] / qv_date_dict[day][wc] 
                        dt_wc_t.append(round(normal_value, 2))
                        normal_value = thresh_const * dt_orig_f_dict[dt][wc][day] / qv_date_dict[day][wc] 
                        dt_wc_f.append(round(normal_value, 2))
                    except:
                        pass
            thresh_og_dict[dt][wc].append(dt_wc_t)
            thresh_og_dict[dt][wc].append(dt_wc_f)


    # Split Thresh OG Dict into Time and Freq dict for printing purpose
    for dt in thresh_og_dict:
        if dt not in thresh_og_t_dict: thresh_og_t_dict[dt] = {}
        if dt not in thresh_og_f_dict: thresh_og_f_dict[dt] = {}

        for wc in thresh_og_dict[dt]:
            if wc not in thresh_og_t_dict[dt]: thresh_og_t_dict[dt][wc] = 0
            if wc not in thresh_og_f_dict[dt]: thresh_og_f_dict[dt][wc] = 0

            if len(thresh_og_dict[dt][wc][0]) > min_list:
                # print(thresh_og_dict[dt][wc][0])
                dt_wc_t_thresh = round(np.percentile(thresh_og_dict[dt][wc][0], thresh_perc), 2)
                thresh_og_dict[dt][wc].append(dt_wc_t_thresh)
                thresh_og_t_dict[dt][wc] = dt_wc_t_thresh
                dt_wc_f_thresh = round(np.percentile(thresh_og_dict[dt][wc][1], thresh_perc), 2)
                thresh_og_dict[dt][wc].append(dt_wc_f_thresh)
                thresh_og_f_dict[dt][wc] = dt_wc_f_thresh
                thresh_og_dict[dt][wc].pop(0)
                thresh_og_dict[dt][wc].pop(0)


    # Operator DT Zone
    for dt in dt_op_f_dict:
        if dt not in thresh_op_dict: #populate op DT in thresh dict
            thresh_op_dict[dt] = {}
        for wc in dt_op_f_dict[dt]:
            if wc not in thresh_op_dict[dt]: #populate OP DT in thresh dict
                thresh_op_dict[dt][wc] = []
            dt_wc_t = []
            dt_wc_f = []
            if type(wc) == int:
                for day in dt_op_t_dict[dt][wc]:
                    try:
                        # print(dt, wc, dt_op_f_dict[dt][wc][day], qv_date_dict[day][wc])
                        normal_value = thresh_const * dt_op_t_dict[dt][wc][day] / qv_date_dict[day][wc] 
                        dt_wc_t.append(round(normal_value, 2))
                        normal_value = thresh_const * dt_op_f_dict[dt][wc][day] / qv_date_dict[day][wc] 
                        dt_wc_f.append(round(normal_value, 2))
                    except:
                        pass
            thresh_op_dict[dt][wc].append(dt_wc_t)
            thresh_op_dict[dt][wc].append(dt_wc_f)


    # Split Thresh OP Dict into Time and Freq dict for printing purpose
    for dt in thresh_op_dict:
        if dt not in thresh_op_t_dict: thresh_op_t_dict[dt] = {}
        if dt not in thresh_op_f_dict: thresh_op_f_dict[dt] = {}

        for wc in thresh_op_dict[dt]:

            if wc not in thresh_op_t_dict[dt]: thresh_op_t_dict[dt][wc] = 0
            if wc not in thresh_op_f_dict[dt]: thresh_op_f_dict[dt][wc] = 0

            if len(thresh_op_dict[dt][wc][0]) > min_list:
                # print(dt, wc, len(thresh_op_dict[dt][wc][0]), end = ' ')
                dt_wc_t_thresh = round(np.percentile(thresh_op_dict[dt][wc][0], thresh_perc), 2)
                # print(len(thresh_op_dict[dt][wc][0]), end = ' ')
                thresh_op_dict[dt][wc].append(dt_wc_t_thresh)
                thresh_op_t_dict[dt][wc] = dt_wc_t_thresh
                dt_wc_f_thresh = round(np.percentile(thresh_op_dict[dt][wc][1], thresh_perc), 2)
                thresh_op_dict[dt][wc].append(dt_wc_f_thresh)
                thresh_op_f_dict[dt][wc] = dt_wc_f_thresh
                thresh_op_dict[dt][wc].pop(0)
                thresh_op_dict[dt][wc].pop(0)



    ## Work done on 9/20/22 
    # weekly_col = ['DT Reason', 'Work Center', 'Date', 'Threshold', 'Value', '% Ratio', 'Class', 'WC-DT'] 
    # week_og_t, week_og_t_anom = week_anom_populate(dt_orig_t_dict, thresh_og_dict, 0, qv_date_dict)
    # week_og_f, week_og_f_anom = week_anom_populate(dt_orig_f_dict, thresh_og_dict, 1, qv_date_dict)
    week_op_t, week_op_t_anom = week_anom_populate(dt_op_t_dict, thresh_op_dict, 0, qv_date_dict)
    week_op_f, week_op_f_anom = week_anom_populate(dt_op_f_dict, thresh_op_dict, 1, qv_date_dict)





    # Create master list of offenders
    weekly_anomaly_list = week_op_t + week_op_f
    # for w in week_op_t:
    #     weekly_anomaly_list.append(w)
    # for w in week_op_f:
    #     weekly_anomaly_list.append(w)
    


    # DT - WC - Date - Type - Value
    watch_list = []
    watch_list = dt_watchlist(dt_op_t_dict, dt_op_f_dict, qv_date_dict)


    # Date - DT - WC - Day Value - OGOP - TF
    day_list = []
    day_list = day_list_populator(dt_orig_t_dict, dt_orig_f_dict, dt_op_t_dict, dt_op_f_dict, thresh_og_dict, thresh_op_dict, qv_date_dict) 

    #create the excel file with the file name
    # multi_sheet_excel_writer(thresh_og_t_dict, thresh_og_f_dict, thresh_op_t_dict, thresh_op_f_dict, raw_file,
    #                         week_og_t_anom, week_og_f_anom, week_op_t_anom, week_op_f_anom, weekly_anomaly_list, day_list, watch_list)
    multi_sheet_excel_writer(thresh_op_t_dict, thresh_op_f_dict, raw_file,
                            week_op_t_anom, week_op_f_anom, weekly_anomaly_list, day_list, watch_list)



def dt_watchlist(dt_op_t_dict, dt_op_f_dict, qv_date_dict):
    watch_list = []
    # print_dict(dt_op_t_dict)
    for dt in dt_op_t_dict:
        if dt in static_watch_list: # only continue if DT is on Watch List
            for wc in dt_op_t_dict[dt]:
                for day in dt_op_t_dict[dt][wc]:
                    y, m, d = days_ago(30)
                    if datetime.datetime(int(day[:4]), int(day[5:7]), int(day[8:])) > datetime.datetime(y, m, d-1):
                        try:
                            value = round(dt_op_t_dict[dt][wc][day], 2)
                            curr_watchlist = [dt, wc, day, "Time", value]
                            watch_list.append(curr_watchlist)
                        except:
                            pass

    for dt in dt_op_f_dict:
        if dt in static_watch_list: # only continue if DT is on Watch List
            for wc in dt_op_f_dict[dt]:
                for day in dt_op_f_dict[dt][wc]:
                    y, m, d = days_ago(30)
                    if datetime.datetime(int(day[:4]), int(day[5:7]), int(day[8:])) > datetime.datetime(y, m, d-1):
                        try:
                            value = round(qv_date_dict[day][wc]/dt_op_f_dict[dt][wc][day], 2)
                            curr_watchlist = [dt, wc, day, "Sheets/Freq", value]
                            watch_list.append(curr_watchlist)
                            value = round(dt_op_f_dict[dt][wc][day], 2)
                            curr_watchlist = [dt, wc, day, "Freq", value]
                            watch_list.append(curr_watchlist)
                            value = round(dt_op_t_dict[dt][wc][day]/dt_op_f_dict[dt][wc][day], 2)
                            curr_watchlist = [dt, wc, day, "avgTime", value]
                            watch_list.append(curr_watchlist)
                        except:
                            pass
    return watch_list


def day_list_single(day_list, in_dict, thresh_dict, qv_date_dict, tf, category):
    for dt in in_dict:
        for wc in in_dict[dt]:
            for day in in_dict[dt][wc].keys():
                y, m, d = days_ago(30)
                if datetime.datetime(int(day[:4]), int(day[5:7]), int(day[8:])) > datetime.datetime(y, m, d-1):
                # if datetime.datetime(int(day[:4]), int(day[5:7]), int(day[8:])) > datetime.datetime(2022, 8, 22):
                    thresh = thresh_dict[dt][wc][tf]
                    if type(thresh) == list: pass #weird error, thresh sometimes is empty list...
                    else:
                        try:
                            thresh, value = round(thresh, 2), round(thresh_const * in_dict[dt][wc][day]/qv_date_dict[day][wc], 2)
                            curr_day_list = [day, dt, wc, value, thresh, category.split()[0], category.split()[1]] 
                            day_list.append(curr_day_list)
                        except: pass
    return day_list


# Create a Date - DT - WC - Day Value - OGOP - TF
def day_list_populator(dt_orig_t_dict, dt_orig_f_dict, dt_op_t_dict, dt_op_f_dict, thresh_og_dict, thresh_op_dict, qv_date_dict):
    day_list = []

    # day_list = day_list_single(day_list, dt_orig_t_dict, thresh_og_dict, qv_date_dict, 0, "Original Time")
    # day_list = day_list_single(day_list, dt_orig_f_dict, thresh_og_dict, qv_date_dict, 1, "Original Freq")
    day_list = day_list_single(day_list, dt_op_t_dict, thresh_op_dict, qv_date_dict, 0, "Operator Time")
    day_list = day_list_single(day_list, dt_op_f_dict, thresh_op_dict, qv_date_dict, 1, "Operator Freq")

    return day_list


def week_anom_populate(in_dict, thresh_dict, tf, qv_date_dict):
    week_list, week_anom_list = [], []
    for dt in in_dict:
        for wc in in_dict[dt]:
            for day in in_dict[dt][wc].keys():
                y, m, d = days_ago(30)
                if datetime.datetime(int(day[:4]), int(day[5:7]), int(day[8:])) > datetime.datetime(y, m, d-1):
                # if datetime.datetime(int(day[:4]), int(day[5:7]), int(day[8:])) > datetime.datetime(2022, 9, 15):
                    thresh = thresh_dict[dt][wc][tf]
                    if type(thresh) == list: pass #weird error, thresh sometimes is empty list...
                    else:
                        try:

                            #Assign all the variables
                            thresh, value = round(thresh, 2), round(thresh_const * in_dict[dt][wc][day]/qv_date_dict[day][wc], 2) 
                            perc = round((value/thresh)-1, 2)
                            anom = int(1) if perc > 0 else int(0)
                            tf_val = 'Time' if tf == 0 else 'Freq'
                            wc_dt = str(wc) + ": " + str(dt)
                            y, m, d = days_ago(8)
                            past_week = "Yes" if datetime.datetime(int(day[:4]), int(day[5:7]), int(day[8:])) > datetime.datetime(y, m, d-1) else "No"

                            #Create array with values
                            day_list = [dt, wc, day, thresh, value, perc, tf_val, wc_dt, anom, past_week] 
                            week_list.append(day_list)
                            if perc > 0: week_anom_list.append(day_list)
                        except: pass
                        
    week_anom_list = sorted(week_anom_list, key=lambda l:l[5], reverse=True)
    return week_list, week_anom_list 

    # multi_sheet_excel_writer(thresh_og_t_dict, thresh_og_f_dict, thresh_op_t_dict, thresh_op_f_dict, raw_file,
    #                         week_og_t_anom, week_og_f_anom, week_op_t_anom, week_op_f_anom, weekly_anomaly_list, day_list, watch_list)
# Old function def
# def multi_sheet_excel_writer(thresh_og_t_dict, thresh_og_f_dict, thresh_op_t_dict, thresh_op_f_dict, filename, wopt, wopf, wal, dl, wl):
def multi_sheet_excel_writer(thresh_op_t_dict, thresh_op_f_dict, filename, wopt, wopf, wal, dl, wl):
    dd_path = os.path.join(long_dir, "datadump", filename)
    book = load_workbook(dd_path)
    
    thresh_filename = filename.split('.')[0] + "_thresh"
    print("Creating %s now..." % thresh_filename)
    # df1 = pd.DataFrame.from_dict(thresh_og_t_dict, orient='index')
    # df2 = pd.DataFrame.from_dict(thresh_og_f_dict, orient='index')
    df3 = pd.DataFrame.from_dict(thresh_op_t_dict, orient='index')
    df4 = pd.DataFrame.from_dict(thresh_op_f_dict, orient='index')

    weekly_col = ['DT Reason', 'Work Center', 'Date', 'Threshold', 'Value', '% Ratio', 'Class', 'WC-DT', 'Anomaly', 'PastWeek']
    # df5 = pd.DataFrame(wogt, columns = weekly_col)
    # df6 = pd.DataFrame(wogf, columns = weekly_col)
    df7 = pd.DataFrame(wopt, columns = weekly_col)
    df8 = pd.DataFrame(wopf, columns = weekly_col)
    df9 = pd.DataFrame(wal, columns = weekly_col)
    df10 = pd.DataFrame(dl, columns = ['Date', 'DT', 'WC', 'Value', 'Thresh', 'OG/OP', 'T/F'])
    df11 = pd.DataFrame(wl, columns = ['DT', 'WC', 'Date', 'Metric', 'Value'])
    # writer = pd.ExcelWriter('%s/%s.xlsx' % (thresh_path, thresh_filename), engine='xlsxwriter')

    #Print for Local
    writer = pd.ExcelWriter('%s/%s.xlsx' % (thresh_path, thresh_filename), engine='openpyxl')
    writer.book = book
    # df1.to_excel(writer, sheet_name='og_t')
    # df2.to_excel(writer, sheet_name='og_f')
    df3.to_excel(writer, sheet_name='op_t')
    df4.to_excel(writer, sheet_name='op_f')
    # df5.to_excel(writer, sheet_name='weekly_og_t', index=False)
    # df6.to_excel(writer, sheet_name='weekly_og_f', index=False)
    df7.to_excel(writer, sheet_name='weekly_op_t', index=False)
    df8.to_excel(writer, sheet_name='weekly_op_f', index=False)
    df9.to_excel(writer, sheet_name='weekly_anomalies', index=False)
    df10.to_excel(writer, sheet_name='thirty_days', index=False)
    df11.to_excel(writer, sheet_name='watch_list', index=False)
    writer.save()

    #Print for OneDrive
    #Pause since we're beta testing locally
    '''
    writer = pd.ExcelWriter('%s/%s.xlsx' % (shared_path, thresh_filename), engine='openpyxl')
    writer.book = book
    writer.save()
    '''

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
        if file.endswith('cst_apr_sep_dd.xlsx'):
        # if file.endswith('.xlsx'):
            dd_loader(file)



def main():


    create_report()




if __name__ == '__main__':
    main()

