const express = require('express');
const fileUpload = require('express-fileupload');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const archiver = require('archiver');
const cors = require('cors');

// Initialize Express app
const app = express();
const port = 3001; // 使用不同的端口避免與React開發服務器衝突

// 定義所有輸出目錄
const azureOutputDir = path.join(__dirname, 'output', 'azure');
const tencentOutputDir = path.join(__dirname, 'output', 'tencent');
const downloadsDir = path.join(__dirname, 'public', 'downloads');

// 確保所有必要目錄存在
const ensureDirectories = () => {
    [azureOutputDir, tencentOutputDir, downloadsDir].forEach(dir => {
        if (!fs.existsSync(dir)) {
            fs.mkdirSync(dir, { recursive: true });
            console.log(`Created directory: ${dir}`);
        }
    });
};

ensureDirectories();

// Middleware setup
app.use(cors({
    origin: ['http://localhost:5173', 'http://localhost:3000', `http://59.148.172.2:3001`], // 允許React開發服務器和生產環境訪問
    credentials: true
}));
app.use(express.json());

// 服務構建後的前端文件
app.use(express.static(path.join(__dirname, 'dist')));
app.use(express.static(path.join(__dirname, 'public')));
app.use(fileUpload({
    limits: { fileSize: 10 * 1024 * 1024 }, // 10MB 限制
    abortOnLimit: true,
    responseOnLimit: "文件大小超過限制（最大10MB）"
}));

// 錯誤處理中間件
const errorHandler = (err, req, res, next) => {
    console.error('Error:', err);
    res.status(500).json({
        success: false,
        message: '服務器內部錯誤',
        error: process.env.NODE_ENV === 'development' ? err.message : undefined
    });
};

// 文件類型驗證
const validateFileType = (file, allowedTypes) => {
    const fileExt = path.extname(file.name).toLowerCase();
    return allowedTypes.includes(fileExt);
};

// 處理Azure數據的函數
const processAzureData = (data) => {
    try {
        // 需要保留的列
        const requiredColumns = [
            'CustomerName', 'UnitPrice', 'BillableQuantity', 'Quantity'
        ];
        
        // 需要刪除的列
        const columnsToDelete = [
            'PartnerId', 'CustomerId', 'InvoiceNumber', 'MpnId', 'Tier2MpnId',
            'Bill to', 'PriceAdjustmentDescription', 'EffectiveUnitPrice'
        ];
        
        // 過濾數據
        const filteredData = data
            .filter(row => row.CustomerName && row.Quantity) // 刪除 Quantity 為空的行
            .filter(row => row['Bill to'] !== 'Accord')      // 排除 Bill to 為 Accord 的數據
            .map(row => {
                // 刪除不需要的列
                columnsToDelete.forEach(col => delete row[col]);
                
                // 計算 Subtotal 和 Total
                const unitPrice = parseFloat(row.UnitPrice) || 0;
                const billableQuantity = parseFloat(row.BillableQuantity) || 0;
                
                row.Subtotal = unitPrice * billableQuantity;
                row.Total = unitPrice * billableQuantity;
                return row;
            });

        // 按客戶分組
        const groupedData = filteredData.reduce((acc, row) => {
            const customer = row.CustomerName;
            if (!acc[customer]) {
                acc[customer] = {
                    rows: [],
                    total: 0
                };
            }
            acc[customer].rows.push(row);
            acc[customer].total += row.Total;
            return acc;
        }, {});

        return groupedData;
    } catch (error) {
        console.error('處理Azure數據時出錯:', error);
        throw error;
    }
};

// 處理騰訊雲數據的函數
const processTencentData = (data) => {
    try {
        // 定義需要保留的列
        const columnsToKeep = [
            'Owner Account ID', 'ProductName', 'SubproductName', 'BillingMode',
            'ProjectName', 'Region', 'InstanceID', 'InstanceName', 'TransactionType',
            'TransactionTime', 'Usage Start Time', 'Usage End Time',
            'Configuration Description', 'OriginalCost'
        ];

        // 過濾和處理數據
        const filteredData = data
            .filter(row => {
                // 確保必要字段存在
                return row['Owner Account ID'] && row['OriginalCost'];
            })
            .map(row => {
                // 創建過濾後的行，只保留需要的列
                const filteredRow = {};
                columnsToKeep.forEach(col => {
                    filteredRow[col] = row[col] || '';
                });

                // 處理OriginalCost字段
                const originalCost = parseFloat(filteredRow['OriginalCost']) || 0;
                filteredRow['OriginalCost'] = originalCost;
                
                // 添加Discount Multiplier列，統一設為1
                filteredRow['Discount Multiplier'] = 1;
                
                // 計算Total Cost
                filteredRow['Total Cost'] = originalCost * filteredRow['Discount Multiplier'];

                return filteredRow;
            });

        // 按Owner Account ID分組
        const groupedData = filteredData.reduce((acc, row) => {
            const ownerId = row['Owner Account ID'];
            if (!acc[ownerId]) {
                acc[ownerId] = {
                    rows: [],
                    total: 0
                };
            }
            acc[ownerId].rows.push(row);
            acc[ownerId].total += row['Total Cost'];
            return acc;
        }, {});

        return groupedData;
    } catch (error) {
        console.error('處理騰訊雲數據時出錯:', error);
        throw error;
    }
};

// API路由

// 健康檢查端點
app.get('/api/health', (req, res) => {
    res.json({
        success: true,
        message: 'Server is running',
        timestamp: new Date().toISOString(),
        version: '2.0.0'
    });
});

// 獲取已處理文件列表
app.get('/api/processed-files', (req, res) => {
    try {
        const getFileDetails = (dir, files) => {
            return files.map(file => {
                const filePath = path.join(dir, file);
                const stats = fs.statSync(filePath);
                return {
                    name: file,
                    size: stats.size,
                    createdAt: stats.mtime.toISOString(),
                    path: `/downloads/${path.basename(dir)}/${file}`
                };
            });
        };

        const azureFiles = fs.existsSync(azureOutputDir) ? 
            fs.readdirSync(azureOutputDir).filter(file => file.endsWith('.xlsx')) : [];
        const tencentFiles = fs.existsSync(tencentOutputDir) ? 
            fs.readdirSync(tencentOutputDir).filter(file => file.endsWith('.xlsx')) : [];

        res.json({
            success: true,
            files: {
                azure: getFileDetails(azureOutputDir, azureFiles),
                tencent: getFileDetails(tencentOutputDir, tencentFiles)
            },
            summary: {
                totalFiles: azureFiles.length + tencentFiles.length,
                azureCount: azureFiles.length,
                tencentCount: tencentFiles.length
            }
        });
    } catch (error) {
        console.error('獲取文件列表時出錯:', error);
        res.status(500).json({
            success: false,
            message: '獲取文件列表失敗',
            error: error.message
        });
    }
});

// 處理Azure報表
app.post('/api/process-azure', (req, res) => {
    try {
        if (!req.files || !req.files.file) {
            return res.status(400).json({
                success: false,
                message: '請選擇要上傳的文件'
            });
        }

        const file = req.files.file;
        
        // 驗證文件類型
        if (!validateFileType(file, ['.xlsx'])) {
            return res.status(400).json({
                success: false,
                message: 'Azure報表僅支持 .xlsx 格式'
            });
        }

        // 讀取Excel文件
        const workbook = xlsx.read(file.data, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet);

        if (data.length === 0) {
            return res.status(400).json({
                success: false,
                message: '文件中沒有找到數據'
            });
        }

        // 處理數據
        const groupedData = processAzureData(data);
        const processedFiles = [];

        // 為每個客戶創建單獨的文件
        Object.keys(groupedData).forEach(customer => {
            const customerData = groupedData[customer];
            const newWorkbook = xlsx.utils.book_new();
            const newWorksheet = xlsx.utils.json_to_sheet(customerData.rows);
            
            xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Data');
            
            const fileName = `${customer.replace(/[^a-zA-Z0-9]/g, '_')}_processed.xlsx`;
            const filePath = path.join(azureOutputDir, fileName);
            
            xlsx.writeFile(newWorkbook, filePath);
            
            // 複製文件到下載目錄
            const downloadPath = path.join(downloadsDir, 'azure', fileName);
            fs.mkdirSync(path.dirname(downloadPath), { recursive: true });
            fs.copyFileSync(filePath, downloadPath);
            
            processedFiles.push(fileName);
        });

        res.json({
            success: true,
            message: 'Azure報表處理完成',
            files: processedFiles,
            summary: {
                totalRecords: data.length,
                processedFiles: processedFiles.length
            }
        });

    } catch (error) {
        console.error('處理Azure報表時出錯:', error);
        res.status(500).json({
            success: false,
            message: '處理Azure報表失敗',
            error: error.message
        });
    }
});

// 處理騰訊雲報表
app.post('/api/process-tencent', (req, res) => {
    try {
        if (!req.files || !req.files.file) {
            return res.status(400).json({
                success: false,
                message: '請選擇要上傳的文件'
            });
        }

        const file = req.files.file;
        
        // 驗證文件類型
        if (!validateFileType(file, ['.xlsx', '.csv'])) {
            return res.status(400).json({
                success: false,
                message: '騰訊雲報表支持 .xlsx 和 .csv 格式'
            });
        }

        let data;
        
        // 根據文件類型讀取數據
        if (path.extname(file.name).toLowerCase() === '.csv') {
            const csvData = file.data.toString('utf8');
            const workbook = xlsx.read(csvData, { type: 'string' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            data = xlsx.utils.sheet_to_json(worksheet);
        } else {
            const workbook = xlsx.read(file.data, { type: 'buffer' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            data = xlsx.utils.sheet_to_json(worksheet);
        }

        if (data.length === 0) {
            return res.status(400).json({
                success: false,
                message: '文件中沒有找到數據'
            });
        }

        // 處理數據
        const groupedData = processTencentData(data);
        const processedFiles = [];

        // 為每個Owner Account ID創建單獨的文件
        Object.keys(groupedData).forEach(ownerId => {
            const ownerData = groupedData[ownerId];
            const newWorkbook = xlsx.utils.book_new();
            
            // 獲取工作表名稱（上個月的年月）
            const lastMonth = new Date();
            lastMonth.setMonth(lastMonth.getMonth() - 1);
            const sheetName = lastMonth.toLocaleDateString('en-US', { month: 'short', year: 'numeric' }).replace(' ', '_');
            
            const newWorksheet = xlsx.utils.json_to_sheet(ownerData.rows);
            xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);
            
            // 使用時間戳命名文件
            const currentTime = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
            const fileName = `output_${ownerId}_${currentTime}.xlsx`;
            const filePath = path.join(tencentOutputDir, fileName);
            
            xlsx.writeFile(newWorkbook, filePath);
            
            // 複製文件到下載目錄
            const downloadPath = path.join(downloadsDir, 'tencent', fileName);
            fs.mkdirSync(path.dirname(downloadPath), { recursive: true });
            fs.copyFileSync(filePath, downloadPath);
            
            processedFiles.push(fileName);
        });

        res.json({
            success: true,
            message: '騰訊雲報表處理完成',
            files: processedFiles,
            summary: {
                totalRecords: data.length,
                processedFiles: processedFiles.length
            }
        });

    } catch (error) {
        console.error('處理騰訊雲報表時出錯:', error);
        res.status(500).json({
            success: false,
            message: '處理騰訊雲報表失敗',
            error: error.message
        });
    }
});

// 批量下載文件
app.get('/api/download-all/:type', (req, res) => {
    try {
        const { type } = req.params;
        let sourceDir;
        let zipName;

        if (type === 'azure') {
            sourceDir = azureOutputDir;
            zipName = 'azure_reports.zip';
        } else if (type === 'tencent') {
            sourceDir = tencentOutputDir;
            zipName = 'tencent_reports.zip';
        } else {
            return res.status(400).json({
                success: false,
                message: '無效的文件類型，請使用 azure 或 tencent'
            });
        }

        if (!fs.existsSync(sourceDir)) {
            return res.status(404).json({
                success: false,
                message: '沒有找到相關文件'
            });
        }

        const files = fs.readdirSync(sourceDir).filter(file => file.endsWith('.xlsx'));
        
        if (files.length === 0) {
            return res.status(404).json({
                success: false,
                message: '沒有找到可下載的文件'
            });
        }

        // 設置響應頭
        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition', `attachment; filename="${zipName}"`);

        // 創建zip壓縮流
        const archive = archiver('zip', {
            zlib: { level: 9 }
        });

        archive.on('error', (err) => {
            console.error('壓縮文件時出錯:', err);
            res.status(500).json({
                success: false,
                message: '壓縮文件失敗'
            });
        });

        // 將壓縮流管道到響應
        archive.pipe(res);

        // 添加文件到壓縮包
        files.forEach(file => {
            const filePath = path.join(sourceDir, file);
            archive.file(filePath, { name: file });
        });

        // 完成壓縮
        archive.finalize();

    } catch (error) {
        console.error('下載文件時出錯:', error);
        res.status(500).json({
            success: false,
            message: '下載文件失敗',
            error: error.message
        });
    }
});

// 錯誤處理中間件
app.use(errorHandler);

// Catch-all handler: 將所有非API請求重定向到前端應用
app.get('*', (req, res) => {
    // 如果是API請求，返回404
    if (req.path.startsWith('/api/')) {
        return res.status(404).json({
            success: false,
            message: 'API端點不存在'
        });
    }
    
    // 否則返回前端應用的index.html
    res.sendFile(path.join(__dirname, 'dist', 'index.html'));
});

// 啟動服務器
app.listen(port, '0.0.0.0', () => {
    console.log(`
🚀 Excel 報表處理系統已啟動
📍 端口: ${port}
🌐 前端界面: http://localhost:${port}
🔗 健康檢查: http://localhost:${port}/api/health
📁 輸出目錄:
   - Azure: ${azureOutputDir}
   - 騰訊雲: ${tencentOutputDir}
   - 下載: ${downloadsDir}
    `);
});

// 優雅關閉
process.on('SIGINT', () => {
    console.log('\n正在關閉服務器...');
    process.exit(0);
});

module.exports = app;

