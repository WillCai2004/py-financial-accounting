import pandas as pd
from openpyxl import Workbook
import os

# 获取当前工作目录
current_directory = os.getcwd()
print("Current working directory:", current_directory)

# 定义文件路径（假设文件在当前目录中）
file_A = os.path.join(current_directory, 'A.xlsx')
file_B = os.path.join(current_directory, 'B.xlsx')

# 检查文件是否存在
if not os.path.exists(file_A):
    raise FileNotFoundError(f"File {file_A} not found.")
if not os.path.exists(file_B):
    raise FileNotFoundError(f"File {file_B} not found.")

# 读取Excel文件
df_A = pd.read_excel(file_A)
df_B = pd.read_excel(file_B)

# 计算A表中的“总价”列（“数量” * “采购单价 含税”）
df_A['总价'] = df_A['数量'] * df_A['采购单价 含税']

# 汇总A表中“代码”对应的总价
df_A_summary = df_A.groupby('代码')['总价'].sum().reset_index()

# 保存A表的汇总结果到TEMP.xlsx
wb_temp = Workbook()
ws_temp = wb_temp.active
ws_temp.title = 'TEMP'

# 写入标题
ws_temp.append(['代码', '总价'])

# 写入数据
for index, row in df_A_summary.iterrows():
    ws_temp.append(row.tolist())

# 保存TEMP.xlsx文件
new_file_temp = os.path.join(current_directory, 'TEMP.xlsx')
wb_temp.save(new_file_temp)

# 汇总B表中“备注”对应的“价税合计”
df_B_summary = df_B.groupby('备注')['价税合计'].sum().reset_index()

# 读取保存的TEMP文件
df_TEMP = pd.read_excel(new_file_temp)

# 确保匹配列的数据类型一致
df_B_summary['备注'] = df_B_summary['备注'].astype(str)
df_TEMP['代码'] = df_TEMP['代码'].astype(str)

# 将B表的汇总结果与TEMP表的汇总结果合并
result = pd.merge(df_B_summary, df_TEMP, left_on='备注', right_on='代码', how='left')

# 计算差值并保留两位小数
result['差值'] = (result['总价'] - result['价税合计']).round(2)

# 修改列名
result = result.rename(columns={'总价_y': '代码', '总价_x': '总价'})

# 设置科学计数法显示为正常数值显示
pd.set_option('display.float_format', lambda x: '%.2f' % x)

# 处理空值
result.fillna('', inplace=True)

# 创建result_new.xlsx并写入结果
wb_result = Workbook()
ws_result = wb_result.active
ws_result.title = 'Result'

# 写入标题
ws_result.append(['备注', '价税合计', '代码', '采购总价','差值'])

# 写入数据
for index, row in result.iterrows():
    ws_result.append(row.tolist())

# 保存result_new.xlsx文件
new_file_result = os.path.join(current_directory, 'result_new.xlsx')
wb_result.save(new_file_result)

print(f'Results saved in {new_file_result}')
