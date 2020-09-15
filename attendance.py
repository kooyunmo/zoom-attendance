import time
import datetime
import argparse
import pandas as pd
import numpy as np


def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--input-file', type=str, required=True,
                        help='input csv file path')
    parser.add_argument('-t', '--min-time', type=int, default=120,
                        help='minimum participation time (min)')
    parser.add_argument('-s', '--start-time', type=str, default='18:30',
                        help='start time of the lecture (HH:MM)')
    parser.add_argument('-l', '--late-time', type=int, default=10,
                        help='threshold time between late and present')
    parser.add_argument('-o', '--output-file', type=str, default='output.xlsx',
                        help='output xlsx file path')
    return parser.parse_args()


if __name__ == '__main__':
    args = get_args()
    
    df = pd.read_excel(args.input_file)
    
    # assign new columns
    row_len = len(df['학번'])
    new_col_strt = pd.Series(np.zeros(row_len))
    new_col_sec = pd.Series(np.zeros(row_len))
    new_col_att = pd.Series(['N/A'] * row_len)
    df = df.assign(start_time = new_col_strt.values)
    df = df.assign(sec = new_col_sec.values)
    df = df.assign(attendance = new_col_att.values)

    # convert string time(HH:MM:SS) to second(float)
    for idx, row in df.iterrows():
        x = time.strptime(row['참여시작시간'].split(' ')[-1], '%H:%M:%S')
        sec_time = datetime.timedelta(hours=x.tm_hour, minutes=x.tm_min, seconds=x.tm_sec).total_seconds()
        df.at[idx, 'start_time'] = sec_time
        x = time.strptime(row['참여시간'], '%H:%M:%S')
        sec_time = datetime.timedelta(hours=x.tm_hour, minutes=x.tm_min, seconds=x.tm_sec).total_seconds()
        df.at[idx, 'sec'] = sec_time

    # for each duplicate student id, join the participation time
    dup_id_list = list(set(df[df.duplicated(subset=['학번'], keep=False)==True]['학번'].tolist()))
    for id in dup_id_list:
        total_part_time = 0
        part_time_list = df.loc[df['학번'] == id]['sec'].tolist()
        for part_time in part_time_list:
            total_part_time += part_time
        h = int(total_part_time) // 3600
        m = (int(total_part_time) - (h * 3600)) // 60
        s = int(total_part_time) % 60
        total_part_time_str = "{:02d}:{:02d}:{:02d}".format(h, m, s)
        df.at[df.loc[df['학번']==id].index[0], '참여시간'] = total_part_time_str
        df.at[df.loc[df['학번']==id].index[0], 'sec'] = total_part_time
        df.at[df.loc[df['학번']==id].index[0], '참여종료시간'] = df.at[df.loc[df['학번']==id].index[-1], '참여종료시간']
    df = df.drop_duplicates(subset=['학번'], keep='first')

    # check attendance
    for idx, _ in df.iterrows():
        if df.at[idx, 'sec'] > args.min_time * 60:
            df.at[idx, 'attendance'] = 'O'
        else:    
            df.at[idx, 'attendance'] = 'X'
    
    # check lateness
    x = time.strptime(args.start_time, '%H:%M')
    class_start_time_sec = datetime.timedelta(hours=x.tm_hour, minutes=x.tm_min).total_seconds()
    for idx, _ in df.iterrows():
        if df.at[idx, 'attendance'] is not 'X' and df.at[idx, 'start_time'] > class_start_time_sec + args.late_time * 60:
            df.at[idx, 'attendance'] = 'L'

    # drop useless column
    df = df.drop(columns=['start_time', 'sec'])

    # save as .xlsx file
    df.to_excel(args.output_file)

