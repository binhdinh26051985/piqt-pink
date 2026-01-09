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

workhr = pd.read_excel(r"d:\IE\EFFCAL\Attendance.xlsx",sheet_name='Attend')
workhr = workhr[workhr["Date"] >= '2026-01-01']

workhr['WorkingTT(hrs)'] = workhr['Attendance']*workhr['Working Hrs']
workhr['SKU'] = workhr['Line'] + '_' + workhr['Date'].astype(str)
workhrs = workhr.groupby(['SKU'])['WorkingTT(hrs)'].sum().reset_index()



workhrs1 = workhrs[['SKU','WorkingTT(hrs)']]

#print(orderoutputsum.head(5))

smv = pd.read_excel(r"d:\IE\EFFCAL\Attendance.xlsx",sheet_name='SMV')    

smv = smv.drop_duplicates(subset=['Style'])

ouputhrs = pd.read_excel(r"d:\IE\EFFCAL\Attendance.xlsx",sheet_name='Output')

ouputhrs = ouputhrs[ouputhrs["Date"] >= '2026-01-01']

ouputhrs = ouputhrs.merge(smv, on='Style', how='left')
ouputhrs ['SmvTT(hrs)'] = ouputhrs['Output']*ouputhrs['SMV']/60
ouputhrs['SKU'] = ouputhrs['Line'] + '_' + ouputhrs['Date'].astype(str)

ouputhrs2 = ouputhrs.groupby(['SKU'])['SmvTT(hrs)'].sum().reset_index()


#ouputhrs2 = ouputhrs1[['SKU','SmvTT']]

efficiency = workhrs1.merge (ouputhrs2, on='SKU', how='left')

efficiency['Eff%'] =efficiency['SmvTT(hrs)']/ efficiency['WorkingTT(hrs)']*100


# Add Price column based on Line condition
efficiency['Price'] = np.where(efficiency['SKU'].str.contains('LINE 4'), 15000, 30000)


efficiency['Incentive'] = np.where(efficiency['Eff%'] > 69.9, (efficiency['Eff%'] -  69.9)* efficiency['Price'], 0)

efficiency['Incentive_65%'] = np.where(efficiency['Eff%'] > 64.9, (efficiency['Eff%'] -  64.9) * efficiency['Price'], 0)

efficiency['Incentive_60%'] = np.where(efficiency['Eff%'] > 59.9, (efficiency['Eff%'] -  59.9) * efficiency['Price'], 0)

efficiency['Incentive_55%'] = np.where(efficiency['Eff%'] > 54.9, (efficiency['Eff%'] -  54.9) * efficiency['Price'], 0)

efficiency['Incentive_50%'] = np.where(efficiency['Eff%'] > 49.9, (efficiency['Eff%'] -  49.9) * efficiency['Price'], 0)

efficiency['Line'] = efficiency['SKU'].str.split('_').str[0]




#Incentive = efficiency.groupby(['Line'])[
    #['Incentive', 'Incentive_65%', 'Incentive_60%','Incentive_55%','Incentive_50%']].sum().reset_index()


import matplotlib.pyplot as plt
import numpy as np

# Calculate incentive by line
Incentive = efficiency.groupby(['Line'])[
    ['Incentive', 'Incentive_65%', 'Incentive_60%', 'Incentive_55%', 'Incentive_50%']
].sum().reset_index()

# Set up the chart
fig, ax = plt.subplots(figsize=(15, 7))

# Create x positions for each line
x = range(len(Incentive['Line']))

# Define colors and markers for each line
colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']
markers = ['o', 's', '^', 'D', 'v']
line_styles = ['-', '--', '-.', ':', '-']
labels = ['Incentive', '65% Incentive', '60% Incentive', '55% Incentive', '50% Incentive']

# Plot each incentive type as a line
for i, col in enumerate(['Incentive', 'Incentive_65%', 'Incentive_60%', 'Incentive_55%', 'Incentive_50%']):
    # Plot the line
    ax.plot(x, Incentive[col], 
            marker=markers[i], 
            color=colors[i], 
            linestyle=line_styles[i],
            linewidth=2,
            markersize=10,
            label=labels[i],
            markerfacecolor='white',  # White fill for markers
            markeredgewidth=2)  # Thicker marker edges
    
    # Add data point labels for EVERY point (not just > 0)
    for j, value in enumerate(Incentive[col]):
        formatted_value = f"{value:,.3f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        
        # Calculate vertical offset to avoid overlap
        vertical_offset = 10 + (i * 3)  # Stagger labels for different lines
        
        # Add the label with white background for readability
        ax.annotate(formatted_value,
                    xy=(j, value),
                    xytext=(0, vertical_offset),
                    textcoords="offset points",
                    ha='center', 
                    va='bottom', 
                    size=8,
                    color=colors[i],
                    fontweight='bold',
                    bbox=dict(boxstyle='round,pad=0.3', 
                             facecolor='white', 
                             edgecolor=colors[i],
                             alpha=0.9))

# Set x-axis ticks and labels
ax.set_xticks(x)
ax.set_xticklabels(Incentive['Line'], rotation=45, ha='right', fontsize=11)

# Add grid
plt.grid(True, linewidth=1, alpha=0.7, linestyle='--')

# Labels and title
plt.xlabel("Line", fontsize=12, fontweight='bold')
plt.ylabel("Million VND", fontsize=12, fontweight='bold')
plt.title("INCENTIVE BY LINE", fontsize=14, fontweight='bold')

# Add legend outside the plot
plt.legend(loc='upper left', bbox_to_anchor=(1.02, 1), fontsize=10, framealpha=0.9)

# Adjust layout to accommodate labels
plt.tight_layout(rect=[0, 0, 0.85, 1])  # Leave space for legend

# Add value range annotation
min_val = Incentive[['Incentive', 'Incentive_65%', 'Incentive_60%', 'Incentive_55%', 'Incentive_50%']].min().min()
max_val = Incentive[['Incentive', 'Incentive_65%', 'Incentive_60%', 'Incentive_55%', 'Incentive_50%']].max().max()
ax.text(0.02, 0.98, f'Value Range: {min_val:,.3f} - {max_val:,.3f}'.replace(',', 'X').replace('.', ',').replace('X', '.'),
        transform=ax.transAxes, fontsize=9, verticalalignment='top',
        bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))

plt.savefig('incentive.png', dpi=80, bbox_inches='tight')
plt.show()


print(Incentive.head(6))

#efficiency.to_excel('efficiency_report.xlsx', index=False)



# Calculate totals for each column
total_row = ['TOTAL VND'] + Incentive[
    ['Incentive', 'Incentive_65%', 'Incentive_60%', 'Incentive_55%', 'Incentive_50%']
].sum().tolist()

# Create a DataFrame for the total row
total_df = pd.DataFrame([total_row], columns=Incentive.columns)

# Append the total row to the end
Incentive = pd.concat([Incentive, total_df], ignore_index=True)



with pd.ExcelWriter('incentive.xlsx') as writer:
    workhr.to_excel(writer, sheet_name='workhr', index=False)
    ouputhrs.to_excel(writer, sheet_name='ouputhrs',index=False)
    efficiency.to_excel(writer, sheet_name='efficiency', index=False)
    Incentive.to_excel(writer, sheet_name='Incentive', index=False)

# Now add the chart to the Excel file
wb = load_workbook('incentive.xlsx')

# Select the Incentive sheet
ws = wb['Incentive']

# Add the chart image to the sheet
img = Image('incentive.png')
# Position the image below the data (adjust the cell position as needed)
img.anchor = 'E2'  # This will place the image starting at cell E2
ws.add_image(img)

# Save the workbook with the chart
wb.save('incentive.xlsx')

# Optional: Clean up the temporary image file
os.remove('incentive.png')