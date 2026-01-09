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
#wk_id = pd.read_excel(r"d:\IE\EFFCAL\Time check SMV.xlsx",sheet_name='HR')

finalgrade = pd.read_excel(r"d:\IE\EFFCAL\ABC_Final report_Dec.xlsx",sheet_name='HR')


mapping = {'A': 100, 'B': 60, 'C': 40}


finalgrade['GSD Scores'] = finalgrade['GSD1'].map(mapping)
finalgrade['Leader Scores'] = finalgrade['Leader Evaluate'].map(mapping)
finalgrade['GSD 60%'] = finalgrade['GSD Scores'] * 0.6
finalgrade['Leader 40%'] = finalgrade['Leader Scores'] * 0.4
finalgrade['Total Scores'] = finalgrade['GSD 60%'] + finalgrade['Leader 40%']


def assign_grade(score):
    if score >= 84:
        return 'A'
    elif score >= 52:
        return 'B'
    else:
        return 'C'

finalgrade['Final Grade'] = finalgrade['Total Scores'].apply(assign_grade)



print(finalgrade.head(5))

finalgrade.to_excel('finalgrade.xlsx', index=False)