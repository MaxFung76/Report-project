import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os
from openpyxl import load_workbook
from datetime import datetime, timedelta

# 檢查並建立輸出文件夾
output_folder = 'output(Azure)'
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 新的要刪除的列
columns_to_delete = ['PartnerId', 'CustomerId', 'InvoiceNumber', 'MpnId', 'Bill to', 'PriceAdjustmentDescription', 'EffectiveUnitPrice']

def process_excel(file_path):
    # 讀取Excel文件
    df = pd.read_excel(file_path)

    # 找出A列有數據但F列沒有數據的行
    rows_to_delete = df[(df['CustomerName'].notnull()) & (df['Quantity'].isnull())].index

    # 刪除符合條件的行
    df.drop(rows_to_delete, inplace=True)

    # 根據CustomerName列分割數據並儲存到不同的Excel文件中，排除Bill to為Accord的數據，同時刪除指定的列，計算Subtotal和Total
    grouped = df.groupby('CustomerName')
    
    for customer, data in grouped:
        if not any(data['Bill to'] == 'Accord'):
            customer_file_name = f"{output_folder}/{customer}.xlsx"
            data.drop(columns=columns_to_delete, inplace=True)
            data['Subtotal'] = data['UnitPrice'] * data['BillableQuantity']
            data['Total'] = data['UnitPrice'] * data['BillableQuantity']
            
            # 在Total列的最後隔一個row加入總和
            total_sum = data['Total'].sum()
            total_row = pd.DataFrame(data={'Total': total_sum}, index=['Total'])
            data = pd.concat([data, total_row])
            
            # Get the name for the new sheet
            last_month = datetime.now().replace(day=1) - timedelta(days=1)
            sheet_name = f"{last_month.strftime('%b_%Y')}"

            if os.path.exists(customer_file_name):
                with pd.ExcelWriter(customer_file_name, engine='openpyxl', mode='a') as writer:
                    data.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                data.to_excel(customer_file_name, index=False)
    
    messagebox.showinfo("Success", "Excel file processed and saved successfully!")

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        process_excel(file_path)

# 創建主視窗
root = tk.Tk()
root.title("Excel Processor")

# 上傳按鈕
upload_button = tk.Button(root, text="Upload Excel File", command=upload_file)
upload_button.pack(pady=20)

# 執行主迴圈
root.mainloop()