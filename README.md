# Excel 報表處理系統 v2.1.0

> 現代化的雲端報表管理平台，支持Azure和騰訊雲賬單報表的智能處理

![系統版本](https://img.shields.io/badge/版本-2.1.0-blue.svg)
![技術棧](https://img.shields.io/badge/技術棧-React%20%2B%20Tailwind%20%2B%20Node.js-green.svg)
![狀態](https://img.shields.io/badge/狀態-生產就緒-success.svg)
![更新](https://img.shields.io/badge/最新更新-2025.08.15-orange.svg)

## 🎉 v2.1.0 重大更新

### ✨ 全新現代化界面
- **🎨 完整UI重構**: 從基礎HTML升級為React + Tailwind CSS現代化設計
- **🌙 暗色模式**: 一鍵切換深色主題，護眼體驗
- **📱 響應式設計**: 完美適配桌面、平板、移動端
- **🎪 拖拽上傳**: 現代化文件上傳體驗
- **📊 實時統計**: 動態數據展示面板

### 🛠️ 騰訊雲處理邏輯修復
- **🔧 正確分組**: 從錯誤的按產品分組改為按Owner Account ID分組
- **📋 標準化處理**: 14個核心字段 + 成本計算邏輯
- **📁 規範命名**: `output_{owner_id}_{timestamp}.xlsx`格式
- **📊 工作表命名**: 使用上個月年月格式（如Jul_2025）

### 🗑️ Azure處理優化
- **新增列刪除**: Tier2MpnId列自動刪除
- **數據清理**: 現在刪除8個不需要的列
- **處理效率**: 提高數據處理和文件生成速度

## 🚀 功能特色

### 🎨 現代化界面
- **玻璃擬態設計**: 半透明卡片效果，現代化視覺體驗
- **漸變背景**: 專業的紫色漸變設計
- **流暢動畫**: 60fps動畫效果，懸停和過渡動畫
- **統計面板**: 藍色、綠色、紫色專業配色的實時統計
- **標籤頁導航**: 清晰的功能分區和導航



### 🛡️ 安全可靠
- **多重驗證**: 前後端雙重文件類型檢查
- **大小限制**: 智能文件大小控制（10MB）
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
git clone https://github.com/MaxFung76/Report-project.git
cd Report-project
```

2. **安裝依賴**
```bash
# 使用 npm (推薦)
npm install

# 或使用 pnpm
pnpm install

# 或使用 yarn
yarn install
```

3. **啟動系統**
```bash
# 方式1: 開發模式
npm run dev        # 前端開發服務器
npm run server     # 後端服務器

# 方式2: 生產模式
npm run build      # 構建前端
npm start          # 啟動生產服務器
```

4. **訪問系統**
- 開發模式: http://localhost:5173
- 生產模式: http://localhost:3001
- API端點: http://localhost:3001/api/health

## 📖 使用指南

### Azure 報表處理
1. 點擊「文件上傳」標籤頁
2. 在Azure區域選擇或拖拽 `.xlsx` 文件
3. 系統自動處理數據：
   - 過濾無效數據（刪除Quantity為空的行）
   - 刪除8個不需要的列（包括Tier2MpnId）
   - 按客戶名稱分組
   - 計算Subtotal和Total
4. 在「文件管理」標籤頁下載結果

### 騰訊雲報表處理
1. 點擊「文件上傳」標籤頁
2. 在騰訊雲區域選擇或拖拽 `.xlsx/.csv` 文件
3. 系統自動處理數據：
   - 保留14個標準字段
   - 按Owner Account ID分組
   - 添加Discount Multiplier（統一設為1）
   - 計算Total Cost
   - 使用上個月年月作為工作表名
4. 在「文件管理」標籤頁下載結果

### 文件管理和下載
- **文件列表**: 顯示文件名、大小、處理時間、狀態
- **單個下載**: 點擊每個文件的下載按鈕
- **批量下載**: 點擊「全部下載」按鈕，自動打包為ZIP格式

## 🏗️ 項目結構

```
Report-project/
├── src/                    # 前端源代碼
│   ├── App.jsx            # 主應用組件（React）
│   ├── App.css            # Tailwind CSS樣式
│   └── main.jsx           # 應用入口
├── public/                # 靜態文件
│   └── downloads/         # 下載文件目錄
│       ├── azure/         # Azure報表輸出
│       └── tencent/       # 騰訊雲報表輸出
├── output/                # 處理結果輸出
│   ├── azure/            # Azure原始輸出
│   └── tencent/          # 騰訊雲原始輸出
├── server.js              # 後端服務器（Express）
├── package.json           # 項目配置
├── tailwind.config.js     # Tailwind配置
├── vite.config.js         # Vite配置
├── CHANGELOG.md           # 更新日誌
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
- `GET /api/download-all/:type` - 批量下載文件（ZIP格式）
- `GET /api/health` - 健康檢查

### 響應格式
```json
{
  "success": true,
  "message": "處理成功",
  "files": [
    "BlueTop_Group_processed.xlsx",
    "CCC_Kei_Yuen_College_processed.xlsx"
  ],
  "summary": {
    "totalRecords": 782,
    "processedFiles": 23
  }
}
```

### 文件列表API響應
```json
{
  "azure": [
    {
      "name": "Customer_Name_processed.xlsx",
      "size": "63.42 kB",
      "time": "8/15/2025, 9:16:22 AM",
      "path": "/downloads/azure/Customer_Name_processed.xlsx"
    }
  ],
  "tencent": [
    {
      "name": "output_{owner_id}_{timestamp}.xlsx",
      "size": "16.76 kB", 
      "time": "8/15/2025, 9:37:57 AM",
      "path": "/downloads/tencent/output_{owner_id}_{timestamp}.xlsx"
    }
  ]
}
```

## 🐛 故障排除

### 常見問題

**Q: 文件上傳失敗**
- 檢查文件格式 (.xlsx, .csv)
- 確認文件大小 < 10MB
- 檢查磁盤空間是否充足

**Q: 前端無法連接後端**
- 確認後端服務器正在運行（端口3001）
- 檢查防火牆設置
- 驗證CORS配置

**Q: 騰訊雲處理結果為空**
- 檢查CSV文件是否包含Owner Account ID列
- 確認OriginalCost列數據格式正確
- 查看瀏覽器控制台錯誤信息

**Q: Azure處理結果異常**
- 檢查Excel文件是否包含CustomerName列
- 確認UnitPrice和BillableQuantity列存在
- 驗證數據格式符合要求

### 日誌查看
```bash
# 查看服務器日誌
tail -f server.log

# 檢查端口占用
netstat -an | grep :3001

# 檢查PM2狀態（生產環境）
pm2 status
pm2 logs excel-report-system
```

## 🚀 生產部署

### 構建生產版本
```bash
# 安裝依賴
npm install

# 構建前端
npm run build

# 啟動生產服務器
NODE_ENV=production npm start
```

### PM2 部署（推薦）
```bash
# 安裝PM2
npm install -g pm2

# 啟動服務
pm2 start server.js --name excel-report-system

# 設置開機自啟
pm2 startup
pm2 save
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

```bash
# 構建鏡像
docker build -t excel-report-system .

# 運行容器
docker run -d -p 3001:3001 --name excel-report excel-report-system
```

## 🔧 技術棧

### 前端
- **React 18.2.0**: 現代化用戶界面框架
- **Tailwind CSS 3.3.0**: 實用優先的CSS框架
- **Vite 4.4.0**: 快速的構建工具
- **Radix UI**: 無障礙的UI組件庫
- **Lucide React**: 現代化圖標庫

### 後端
- **Node.js**: 服務器運行環境
- **Express.js**: Web應用框架
- **xlsx**: Excel文件處理
- **express-fileupload**: 文件上傳處理
- **archiver**: ZIP文件打包
- **cors**: 跨域資源共享

### 開發工具
- **ESLint**: 代碼質量檢查
- **Nodemon**: 開發環境自動重啟
- **PM2**: 生產環境進程管理

## 🤝 貢獻指南

1. Fork 本項目
2. 創建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 開啟 Pull Request

### 開發規範
- 使用ESLint進行代碼檢查
- 遵循React Hooks最佳實踐
- 保持組件的單一職責原則
- 添加適當的錯誤處理

## 📄 許可證

本項目採用 MIT 許可證 - 查看 [LICENSE](LICENSE) 文件了解詳情

## 📞 技術支持

- **問題反饋**: [GitHub Issues](https://github.com/MaxFung76/Report-project/issues)
- **功能建議**: [GitHub Discussions](https://github.com/MaxFung76/Report-project/discussions)
- **技術文檔**: [Wiki](https://github.com/MaxFung76/Report-project/wiki)

## 🎯 版本歷史

### v2.1.0 (2025-08-15) 🎉
- ✨ **前端界面現代化**: React + Tailwind CSS完整重構
- 🛠️ **騰訊雲處理修復**: 按Owner Account ID正確分組
- 🗑️ **Azure處理優化**: 新增Tier2MpnId列刪除規則
- 🔧 **API端點修復**: 解決前端硬編碼localhost問題
- 📁 **文件管理完善**: 完整的下載功能和文件列表
- 🎨 **用戶體驗提升**: 暗色模式、響應式設計、拖拽上傳
- ⚡ **性能優化**: 提高處理速度和系統穩定性

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

## 📊 系統截圖

### 主界面
![主界面](https://via.placeholder.com/800x400/667eea/ffffff?text=現代化主界面)

### 文件上傳
![文件上傳](https://via.placeholder.com/800x400/48bb78/ffffff?text=拖拽上傳界面)

### 文件管理
![文件管理](https://via.placeholder.com/800x400/9f7aea/ffffff?text=文件管理界面)

### 暗色模式
![暗色模式](https://via.placeholder.com/800x400/2d3748/ffffff?text=暗色模式界面)

---

**讓Excel報表處理變得簡單高效！** 🚀

如果這個項目對您有幫助，請給我們一個 ⭐ Star！

## 🔗 相關鏈接

- [在線演示](http://59.148.172.2:3001) - 體驗最新版本
- [更新日誌](CHANGELOG.md) - 查看詳細更新記錄
- [API文檔](https://github.com/MaxFung76/Report-project/wiki/API) - 完整API說明
- [部署指南](https://github.com/MaxFung76/Report-project/wiki/Deploy) - 生產環境部署

---

**最後更新**: 2025年8月15日  
**當前版本**: v2.1.0  
**維護狀態**: 🟢 積極維護中

