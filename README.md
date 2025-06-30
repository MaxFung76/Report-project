# Excel 報表處理系統 v2.0

> 現代化的雲端報表管理平台，支持Azure和騰訊雲賬單報表的智能處理

![系統版本](https://img.shields.io/badge/版本-2.0.0-blue.svg)
![技術棧](https://img.shields.io/badge/技術棧-React%20%2B%20Node.js-green.svg)
![狀態](https://img.shields.io/badge/狀態-生產就緒-success.svg)

## ✨ 功能特色

### 🎨 現代化界面
- **響應式設計**: 完美適配桌面和移動設備
- **暗色模式**: 支持系統主題自動切換
- **直觀操作**: 拖拽上傳，一鍵處理
- **實時反饋**: 處理進度和狀態實時顯示

### 🔧 智能處理
- **Azure報表**: 自動過濾、分組、計算客戶賬單
- **騰訊雲報表**: 按Owner Account ID智能分組
- **批量處理**: 支持多文件同時上傳處理
- **格式支持**: Excel (.xlsx) 和 CSV 格式

### 🛡️ 安全可靠
- **多重驗證**: 前後端雙重文件類型檢查
- **大小限制**: 智能文件大小控制
- **錯誤處理**: 完整的異常捕獲和用戶提示
- **數據保護**: 臨時文件自動清理

## 🚀 快速開始

### 環境要求
- Node.js 16.0+ 
- 4GB+ RAM
- 1GB+ 可用存儲空間

### 安裝步驟

1. **克隆項目**
```bash
git clone <repository-url>
cd excel-report-system
```

2. **安裝依賴**
```bash
# 使用 pnpm (推薦)
pnpm install

# 或使用 npm
npm install

# 或使用 yarn
yarn install
```

3. **啟動系統**
```bash
# 終端1: 啟動後端服務
node server.js

# 終端2: 啟動前端開發服務器
pnpm run dev --host
```

4. **訪問系統**
- 前端界面: http://localhost:5173
- 後端API: http://localhost:3001
- 健康檢查: http://localhost:3001/api/health

## 📖 使用指南

### Azure 報表處理
1. 點擊「文件上傳」標籤頁
2. 在Azure區域選擇或拖拽 `.xlsx` 文件
3. 系統自動處理數據：
   - 過濾無效數據
   - 按客戶名稱分組
   - 計算小計和總計
4. 在「文件管理」標籤頁下載結果

### 騰訊雲報表處理
1. 點擊「文件上傳」標籤頁
2. 在騰訊雲區域選擇或拖拽 `.xlsx/.csv` 文件
3. 系統自動處理數據：
   - 提取關鍵列信息
   - 按Owner Account ID分組
   - 計算折扣和總成本
4. 在「文件管理」標籤頁下載結果

### 批量下載
- 點擊「全部下載」按鈕下載對應類型的所有文件
- 文件自動打包為ZIP格式，便於管理

## 🏗️ 項目結構

```
excel-report-system/
├── src/                    # 前端源代碼
│   ├── components/         # React組件
│   ├── assets/            # 靜態資源
│   └── App.jsx            # 主應用組件
├── public/                # 靜態文件
├── output/                # 處理結果輸出
│   ├── azure/            # Azure報表輸出
│   └── tencent/          # 騰訊雲報表輸出
├── server.js              # 後端服務器
├── package.json           # 項目配置
└── README.md             # 項目說明
```

## 🔧 配置選項

### 環境變量
```bash
# .env 文件
PORT=3001                  # 後端端口
NODE_ENV=production        # 運行環境
UPLOAD_LIMIT=10485760     # 上傳限制 (10MB)
CORS_ORIGIN=http://localhost:5173  # CORS配置
```

### 自定義配置
```javascript
// server.js 中可修改的配置
const port = 3001;                    // 服務端口
const uploadLimit = 10 * 1024 * 1024; // 文件大小限制
const azureOutputDir = './output/azure';    // Azure輸出目錄
const tencentOutputDir = './output/tencent'; // 騰訊雲輸出目錄
```

## 📊 API 文檔

### 核心端點
- `POST /api/process-azure` - 處理Azure報表
- `POST /api/process-tencent` - 處理騰訊雲報表
- `GET /api/processed-files` - 獲取處理後文件列表
- `GET /api/download-all/:type` - 批量下載文件
- `GET /api/health` - 健康檢查

### 響應格式
```json
{
  "success": true,
  "message": "處理成功",
  "files": [...],
  "summary": {
    "totalFiles": 5,
    "totalRecords": 1250
  }
}
```

## 🐛 故障排除

### 常見問題

**Q: 文件上傳失敗**
- 檢查文件格式 (.xlsx, .csv)
- 確認文件大小 < 10MB
- 檢查磁盤空間是否充足

**Q: 前端無法連接後端**
- 確認後端服務器正在運行
- 檢查端口3001是否被占用
- 驗證CORS配置

**Q: 處理結果異常**
- 檢查Excel文件列名是否正確
- 確認數據格式符合要求
- 查看瀏覽器控制台錯誤信息

### 日誌查看
```bash
# 查看服務器日誌
tail -f server.log

# 檢查端口占用
netstat -an | grep :3001
```

## 🚀 生產部署

### 構建生產版本
```bash
# 構建前端
pnpm run build

# 啟動生產服務器
NODE_ENV=production node server.js
```

### Docker 部署
```dockerfile
FROM node:16-alpine
WORKDIR /app
COPY package*.json ./
RUN npm install
COPY . .
RUN npm run build
EXPOSE 3001
CMD ["node", "server.js"]
```

## 🤝 貢獻指南

1. Fork 本項目
2. 創建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 開啟 Pull Request

## 📄 許可證

本項目採用 MIT 許可證 - 查看 [LICENSE](LICENSE) 文件了解詳情

## 📞 技術支持

- **問題反饋**: [GitHub Issues](https://github.com/your-repo/issues)
- **功能建議**: [GitHub Discussions](https://github.com/your-repo/discussions)
- **技術文檔**: [Wiki](https://github.com/your-repo/wiki)

## 🎯 版本歷史

### v2.0.0 (2024-06-30)
- 🎉 全新現代化界面設計
- 🔧 完整技術架構重構
- 🛡️ 全面安全性加固
- 📱 響應式設計支持
- ⚡ 性能大幅提升

### v1.0.0 (原版本)
- 基礎Excel處理功能
- 簡單HTML界面
- Node.js + Python混合架構

---

**讓Excel報表處理變得簡單高效！** 🚀

如果這個項目對您有幫助，請給我們一個 ⭐ Star！

