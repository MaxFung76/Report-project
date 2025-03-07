# Excel 報表處理系統

這是一個用於處理 Azure 和騰訊雲賬單報表的自動化處理系統。系統支持通過網頁界面上傳 Excel 文件，自動處理數據並生成相應的報表。

## 功能特點

- 支持處理 Azure 雲賬單報表
- 支持處理騰訊雲賬單報表
- 提供網頁界面進行文件上傳
- 自動分類並處理數據
- 支持批量下載處理後的文件

## 系統架構

- 前端：HTML + JavaScript
- 後端：Node.js (Express) + Python
- 文件處理：xlsx 庫

## 安裝依賴

```bash
npm install
```

## 啟動服務

```bash
node server.js
```

服務將在 http://localhost:3000 啟動

## 使用方法

1. 訪問系統網頁界面
2. 選擇要上傳的 Excel 文件（Azure 或騰訊雲報表）
3. 點擊上傳按鈕
4. 等待系統處理完成
5. 在文件列表中下載處理後的報表

## 文件結構

- `server.js`: 主要的服務器代碼
- `Azure_Report.py`: Azure 報表處理邏輯
- `Tencent_Report.py`: 騰訊雲報表處理邏輯
- `views/index.html`: 網頁界面
- `output(Azure)/`: Azure 報表輸出目錄
- `output(Tencent)/`: 騰訊雲報表輸出目錄

## 注意事項

- 確保上傳的文件格式正確
- 處理後的文件會自動保存到對應的輸出目錄
- 支持的文件格式：.xlsx, .csv