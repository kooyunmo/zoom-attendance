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
    parser.add_argument('-o', '--output-file', type=str, default='output.xlsx',
                        help='output xlsx file path')
    return parser.parse_args()


if __name__ == '__main__':
    args = get_args()
    
    df = pd.read_excel(args.input_file)
    
    # assign new columns
    row_len = len(df['학번'])
    new_col_sec = pd.Series(np.zeros(row_len))
    new_col_att = pd.Series(['N/A'] * row_len)
    df = df.assign(sec = new_col_sec.values)
    df = df.assign(attendance = new_col_att.values)

    # convert string time(HH:MM:SS) to second(float)
    for idx, row in df.iterrows():
        x = time.strptime(row['참여시간'], '%H:%M:%S')
        sec_time = datetime.timedelta(hours=x.tm_hour, minutes=x.tm_min, seconds=x.tm_sec).total_seconds()
        df.at[idx, 'sec'] = sec_time

    # join duplicate names' participation time
    dup_name_list = list(set(df[df.duplicated(subset=['이름'], keep=False)==True]['이름'].tolist()))
    for name in dup_name_list:
        total_part_time = 0
        part_time_list = df.loc[df['이름'] == name]['sec'].tolist()
        for part_time in part_time_list:
            total_part_time += part_time
        df.at[df.loc[df['이름']==name].index[0], 'sec'] = total_part_time
    df = df.drop_duplicates(subset=['이름'], keep='first')

    # check attendance
    for idx, row in df.iterrows():
        if df.at[idx, 'sec'] > args.min_time * 60:
            df.at[idx, 'attendance'] = 'O'
        else:    
            df.at[idx, 'attendance'] = 'X'

    # save as .xlsx file
    df.to_excel(args.output_file)

