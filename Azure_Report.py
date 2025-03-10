import pandas as pd
import os
from datetime import datetime, timedelta
import logging
from typing import Dict, Any
from pathlib import Path

# 配置日誌
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('azure_report.log'),
        logging.StreamHandler()
    ]
)

# 配置
CONFIG = {
    'output_folder': 'output(Azure)',
    'columns_to_delete': [
        'PartnerId', 'CustomerId', 'InvoiceNumber', 'MpnId',
        'Bill to', 'PriceAdjustmentDescription', 'EffectiveUnitPrice'
    ]
}

class ExcelProcessor:
    def __init__(self, config: Dict[str, Any]):
        """初始化 Excel 處理器

        Args:
            config: 配置字典，包含輸出目錄和需要刪除的列等設置
        """
        self.config = config
        self._ensure_output_folder()

    def _ensure_output_folder(self) -> None:
        """確保輸出目錄存在"""
        output_path = Path(self.config['output_folder'])
        if not output_path.exists():
            output_path.mkdir(parents=True)
            logging.info(f"創建輸出目錄：{output_path}")

    def _clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """清理數據

        Args:
            df: 輸入的 DataFrame

        Returns:
            清理後的 DataFrame
        """
        try:
            # 刪除無效行
            rows_to_delete = df[(df['CustomerName'].notnull()) & 
                               (df['Quantity'].isnull())].index
            df = df.drop(rows_to_delete)
            logging.info(f"刪除了 {len(rows_to_delete)} 行無效數據")

            # 刪除不需要的列
            df = df.drop(columns=self.config['columns_to_delete'])
            return df
        except Exception as e:
            logging.error(f"數據清理過程中發生錯誤：{str(e)}")
            raise

    def _calculate_totals(self, df: pd.DataFrame) -> pd.DataFrame:
        """計算小計和總計

        Args:
            df: 輸入的 DataFrame

        Returns:
            添加了計算結果的 DataFrame
        """
        try:
            df['Subtotal'] = df['UnitPrice'] * df['BillableQuantity']
            df['Total'] = df['UnitPrice'] * df['BillableQuantity']
            
            # 計算總和並添加匯總行
            total_sum = df['Total'].sum()
            total_row = pd.DataFrame({'Total': [total_sum]}, index=['Total'])
            df = pd.concat([df, total_row])
            
            return df
        except Exception as e:
            logging.error(f"計算總計時發生錯誤：{str(e)}")
            raise

    def _get_sheet_name(self) -> str:
        """獲取工作表名稱（上個月的年月）"""
        last_month = datetime.now().replace(day=1) - timedelta(days=1)
        return last_month.strftime('%b_%Y')

    def _save_customer_data(self, customer: str, data: pd.DataFrame) -> None:
        """保存客戶數據到 Excel 文件

        Args:
            customer: 客戶名稱
            data: 要保存的數據
        """
        try:
            file_path = Path(self.config['output_folder']) / f"{customer}.xlsx"
            sheet_name = self._get_sheet_name()

            if file_path.exists():
                with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
                    data.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                data.to_excel(file_path, index=False)

            logging.info(f"成功保存客戶 {customer} 的數據到 {file_path}")
        except Exception as e:
            logging.error(f"保存客戶 {customer} 的數據時發生錯誤：{str(e)}")
            raise

    def process_excel(self, file_path: str) -> None:
        """處理 Excel 文件

        Args:
            file_path: Excel 文件路徑
        """
        try:
            logging.info(f"開始處理文件：{file_path}")
            
            # 讀取 Excel 文件
            df = pd.read_excel(file_path)
            logging.info(f"成功讀取文件，共 {len(df)} 行數據")

            # 清理數據
            df = self._clean_data(df)

            # 按客戶分組處理數據
            for customer, customer_data in df.groupby('CustomerName'):
                if not any(customer_data['Bill to'] == 'Accord'):
                    # 計算總計
                    processed_data = self._calculate_totals(customer_data)
                    # 保存數據
                    self._save_customer_data(customer, processed_data)

            logging.info("文件處理完成")
            return True
        except Exception as e:
            logging.error(f"處理文件時發生錯誤：{str(e)}")
            raise

def main(file_path: str) -> bool:
    """主函數

    Args:
        file_path: 要處理的 Excel 文件路徑

    Returns:
        bool: 處理是否成功
    """
    try:
        processor = ExcelProcessor(CONFIG)
        return processor.process_excel(file_path)
    except Exception as e:
        logging.error(f"程序執行失敗：{str(e)}")
        return False

if __name__ == '__main__':
    import tkinter as tk
    from tkinter import filedialog, messagebox

    def upload_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                if main(file_path):
                    messagebox.showinfo("成功", "Excel 文件處理完成！")
                else:
                    messagebox.showerror("錯誤", "處理過程中發生錯誤，請查看日誌文件了解詳情。")
            except Exception as e:
                messagebox.showerror("錯誤", f"處理失敗：{str(e)}")

    # 創建主視窗
    root = tk.Tk()
    root.title("Excel 報表處理器")

    # 上傳按鈕
    upload_button = tk.Button(root, text="上傳 Excel 文件", command=upload_file)
    upload_button.pack(pady=20)

    # 執行主迴圈
    root.mainloop()