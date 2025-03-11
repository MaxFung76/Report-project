import pandas as pd
import os
from tkinter import filedialog
from tkinter import messagebox
import tkinter as tk
from openpyxl import load_workbook
from datetime import datetime, timedelta

# 檢查並建立輸出文件夾
output_folder = 'output(Tencent)'
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

def process_excel(file_path):
    # 讀取CSV或Excel文件
    if file_path.lower().endswith('.csv'):
        df = pd.read_csv(file_path, encoding='gbk')
    else:
        df = pd.read_excel(file_path)
    
    # 選擇需要的列
    columns_to_keep = ['Owner Account ID', 'ProductName', 'SubproductName', 'BillingMode', 
                      'ProjectName', 'Region', 'InstanceID', 'InstanceName', 'TransactionType', 
                      'TransactionTime', 'Usage Start Time', 'Usage End Time', 
                      'Configuration Description', 'OriginalCost']
    
    df_filtered = df[columns_to_keep]
    
    # 添加Discount Multiplier列，統一設為1
    df_filtered['Discount Multiplier'] = 1
    
    # 計算Total Cost
    df_filtered['Total Cost'] = df_filtered['OriginalCost'] * df_filtered['Discount Multiplier']
    
    # 獲取上個月的月份名稱作為工作表名稱
    last_month = datetime.now().replace(day=1) - timedelta(days=1)
    sheet_name = f"{last_month.strftime('%b_%Y')}"
    
    # 根據Owner Account ID分組並保存到不同的Excel文件
    current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
    for owner_id, data in df_filtered.groupby('Owner Account ID'):
        output_file = f'{output_folder}/output_{owner_id}_{current_time}.xlsx'
        data.to_excel(output_file, sheet_name=sheet_name, index=False)
    
    messagebox.showinfo('成功', 'Excel文件已成功處理並保存！')

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx'), ('CSV files', '*.csv')])
    if file_path:
        process_excel(file_path)

# 創建主視窗
root = tk.Tk()
root.title('Excel處理程序')

# 上傳按鈕
upload_button = tk.Button(root, text='上傳CSV/Excel文件', command=upload_file)
upload_button.pack(pady=20)

# 執行主迴圈
root.mainloop()