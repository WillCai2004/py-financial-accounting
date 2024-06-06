import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
import os

current_directory = os.getcwd()
print("Current working directory:", current_directory)

file_A = os.path.join(current_directory, 'A.xlsx')
file_B = os.path.join(current_directory, 'B.xlsx')

if not os.path.exists(file_A):
    raise FileNotFoundError(f"File {file_A} not found.")
if not os.path.exists(file_B):
    raise FileNotFoundError(f"File {file_B} not found.")

df_A = pd.read_excel(file_A)
df_B = pd.read_excel(file_B)

df_A['总价'] = df_A['数量'] * df_A['采购单价 含税']
df_A_summary = df_A.groupby('代码')['总价'].sum().reset_index()

wb_temp = Workbook()
ws_temp = wb_temp.active
ws_temp.title = 'TEMP'
ws_temp.append(['代码', '总价'])
for index, row in df_A_summary.iterrows():
    ws_temp.append(row.tolist())

new_file_temp = os.path.join(current_directory, 'TEMP.xlsx')
wb_temp.save(new_file_temp)

df_B_summary = df_B.groupby('备注')['价税合计'].sum().reset_index()
df_TEMP = pd.read_excel(new_file_temp)

df_B_summary['备注'] = df_B_summary['备注'].astype(str)
df_TEMP['代码'] = df_TEMP['代码'].astype(str)

result = pd.merge(df_B_summary, df_TEMP, left_on='备注', right_on='代码', how='left')
result['差值'] = (result['总价'] - result['价税合计']).round(2)
result = result.rename(columns={'总价_y': '代码', '总价_x': '总价'})
pd.set_option('display.float_format', lambda x: '%.2f' % x)
result.fillna('', inplace=True)

wb_result = Workbook()
ws_result = wb_result.active
ws_result.title = 'Result'
ws_result.append(['备注', '价税合计', '代码', '采购总价', '差值'])
for index, row in result.iterrows():
    ws_result.append(row.tolist())

new_file_result = os.path.join(current_directory, 'result_new.xlsx')
wb_result.save(new_file_result)

wb_A = load_workbook(filename=file_A)
ws_A = wb_A.active
wb_result_new = load_workbook(filename=new_file_result)
ws_result_new = wb_result_new.active

codes_to_highlight = set()
for row in ws_result_new.iter_rows(min_row=2, values_only=True):
    if row[4] == 0:
        code = str(row[2])
        codes_to_highlight.add(code)

for i, code_a in enumerate(df_A['代码']):
    if str(code_a) in codes_to_highlight:
        cell = ws_A.cell(row=i+2, column=df_A.columns.get_loc('代码')+1)
        cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

wb_A.save(file_A)

print(f"Cells with a difference of 0 in '差值' column highlighted in green in file A.xlsx")
input("Press Enter to exit...")
