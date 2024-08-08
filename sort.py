import pandas as pd
import numpy as np
import argparse
import os
import re

input_folder = "input"
output_folder = "."
output_name = "Summary"

file_data = []
file_names = []

output = output_folder + "/" + output_name + ".xlsx"
writer = pd.ExcelWriter(output, engine='xlsxwriter')

for file in os.listdir(input_folder):
    name =  re.findall(r'(.+)\.xlsx$', file)
    if not name:
        continue
    file_names.append(name[0])

    file_path = input_folder + '/' + file

    excel_data = pd.read_excel(file_path)
    file_data.append(excel_data)

for df, name in zip(file_data, file_names):
    df.to_excel(writer, sheet_name=name[:31])

start_row = 0
for data, run in zip(file_data,file_names):
    calls = {}
    for call in data['Label']:
        if call not in calls:
            calls[call] = 1
        else:
            calls[call] += 1

    total_calls = 0
    for call_type in calls:
        total_calls+= calls[call_type]
    calls['Total'] = total_calls

    call_len_avg = data['Call Length (s)'].mean()
    principal_freq_avg = data['Principal Frequency (kHz)'].mean()
    slope_avg = data['Slope (kHz/s)'].mean()
    avgs_dict = {
        'Average Call Length (s)': [call_len_avg],
        'Average Principal Frequency (kHz)': [principal_freq_avg],
        'Average Slope (kHz/s)': [slope_avg]
    }


    pd.DataFrame(data={run:[run]}).to_excel(writer, sheet_name='Summary', index=False, startrow=start_row, header=False)
    start_row += 2

    pd.DataFrame(data={'Calls':['Calls']}).to_excel(writer, sheet_name='Summary', index=False, startrow=start_row, startcol=1)
    start_row += 1

    df = pd.DataFrame(data=calls, index=[0])
    df.to_excel(writer, sheet_name='Summary', index=False, startrow=start_row,startcol=1)
    start_row += 3

    df = pd.DataFrame(data=avgs_dict)
    df.to_excel(writer, sheet_name='Summary', index=False, startrow=start_row, startcol=1)
    start_row +=4

writer.close()