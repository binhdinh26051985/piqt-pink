import pandas as pd
import os
from datetime import datetime
import datetime as DT
import calendar
import numpy as np
import matplotlib.pyplot as plt
import datetime as dt
from datetime import timedelta
import re
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

#Orderouput = pd.read_excel(r"D:\IE\EFFCAL\HANGER INPUT & OUTPUT REPORT in Nov.xlsx")
#print(Orderouput.head(5))
wk_id = pd.read_excel(r"d:\IE\EFFCAL\Time check SMV.xlsx",sheet_name='HR')

wk_id = wk_id[['Register ID','Vietnamese name']]


line1 = pd.read_excel(r"d:\IE\EFFCAL\Time check SMV.xlsx",sheet_name='Line 1')
line2 = pd.read_excel(r"d:\IE\EFFCAL\Time check SMV.xlsx",sheet_name='Line 2')
line3 = pd.read_excel(r"d:\IE\EFFCAL\Time check SMV.xlsx",sheet_name='Line 3')
line4 = pd.read_excel(r"d:\IE\EFFCAL\Time check SMV.xlsx",sheet_name='Line 4')
line5 = pd.read_excel(r"d:\IE\EFFCAL\Time check SMV.xlsx",sheet_name='Line 5')
line6 = pd.read_excel(r"d:\IE\EFFCAL\Time check SMV.xlsx",sheet_name='Line 6')

all_lines = pd.concat([line1, line2, line3, line4, line5, line6], ignore_index=True)

all_lines = all_lines[all_lines["Date"] >= '2025-12-1']

all_lines['SMVs'] = all_lines['SMV'] *60

all_lines = all_lines[all_lines['RM'] != 'Not Sew']

all_lines1 = all_lines.groupby(['Register ID','Line'])[['TIME (pcs/sec)', 'SMVs']].sum().reset_index()

all_lines1['Eff.%'] = all_lines1['SMVs']/all_lines1['TIME (pcs/sec)']*100


# Calculate percentiles
p30 = all_lines1['Eff.%'].quantile(0.3)
p70 = all_lines1['Eff.%'].quantile(0.7)

# Assign grades A, B, C
all_lines1['GSD_Grade'] = np.where(all_lines1['Eff.%'] >= p70, 'A',
                               np.where(all_lines1['Eff.%'] >= p30, 'B', 'C'))

# Sort by efficiency
all_lines1 = all_lines1.sort_values('Eff.%', ascending=False)

# Add GSD Scores based on Grade - NEW MAPPING
gsd_mapping = {'A': 36, 'B': 21.6, 'C': 14.4}
all_lines1['GSD Scores'] = all_lines1['GSD_Grade'].map(gsd_mapping)

all_lines1 = all_lines1.merge(wk_id, on='Register ID', how='left')

# Display the DataFrame with grades
print("Employees with Efficiency Grades:")
print(all_lines1[['Register ID', 'Vietnamese name', 'TIME (pcs/sec)', 'SMVs', 'Eff.%', 'GSD_Grade']].to_string(index=False))


all_lines1.to_excel('all_lines.xlsx', index=False)
