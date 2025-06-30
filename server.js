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
    origin: ['http://localhost:5173', 'http://localhost:3000'], // 允許React開發服務器訪問
    credentials: true
}));
app.use(express.json());
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
            'PartnerId', 'CustomerId', 'InvoiceNumber', 'MpnId',
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
        console.error('Azure data processing error:', error);
        throw new Error('Azure數據處理失敗');
    }
};

// 處理騰訊雲數據的函數
const processTencentData = (data) => {
    try {
        // 需要保留的列
        const columnsToKeep = [
            'Owner Account ID', 'ProductName', 'SubproductName', 'BillingMode',
            'ProjectName', 'Region', 'InstanceID', 'InstanceName', 'TransactionType',
            'TransactionTime', 'Usage Start Time', 'Usage End Time',
            'Configuration Description', 'OriginalCost'
        ];

        const processedData = data.map(row => {
            const newRow = {};
            // 只保留指定的列
            columnsToKeep.forEach(col => {
                if (row[col] !== undefined) newRow[col] = row[col];
            });
            // 添加 Discount Multiplier 和計算 Total Cost
            newRow['Discount Multiplier'] = 1;
            const originalCost = parseFloat(newRow['OriginalCost']) || 0;
            newRow['Total Cost'] = originalCost * newRow['Discount Multiplier'];
            return newRow;
        });

        // 按 Owner Account ID 分組
        const groupedData = processedData.reduce((acc, row) => {
            const ownerId = row['Owner Account ID'];
            if (!acc[ownerId]) acc[ownerId] = [];
            acc[ownerId].push(row);
            return acc;
        }, {});

        return groupedData;
    } catch (error) {
        console.error('Tencent data processing error:', error);
        throw new Error('騰訊雲數據處理失敗');
    }
};

// 保存Azure數據
const saveAzureData = (processedData) => {
    const now = new Date();
    const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const sheetName = `${lastMonth.toLocaleString('en', { month: 'short' })}_${lastMonth.getFullYear()}`;
    const timestamp = now.toISOString().replace(/[:.]/g, '-').split('T')[0];

    const savedFiles = [];

    for (const [customer, data] of Object.entries(processedData)) {
        const workbook = xlsx.utils.book_new();
        
        // 添加所有行
        const worksheet = xlsx.utils.json_to_sheet(data.rows);
        
        // 添加總計行
        const totalRow = { CustomerName: 'Total', Total: data.total };
        const totalRowPos = data.rows.length + 2; // 空一行後添加總計
        xlsx.utils.sheet_add_json(worksheet, [totalRow], {
            skipHeader: true,
            origin: `A${totalRowPos}`
        });
        
        xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
        
        const fileName = `${customer}_${timestamp}.xlsx`;
        const filePath = path.join(azureOutputDir, fileName);
        xlsx.writeFile(workbook, filePath);
        
        savedFiles.push({
            name: fileName,
            path: filePath,
            customer: customer,
            recordCount: data.rows.length,
            total: data.total
        });
    }

    return savedFiles;
};

// 保存騰訊雲數據
const saveTencentData = (processedData) => {
    const now = new Date();
    const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const sheetName = `${lastMonth.toLocaleString('en', { month: 'short' })}_${lastMonth.getFullYear()}`;
    const timestamp = now.toISOString().replace(/[:.]/g, '-').split('T')[0];

    const savedFiles = [];

    for (const [ownerId, data] of Object.entries(processedData)) {
        const workbook = xlsx.utils.book_new();
        const worksheet = xlsx.utils.json_to_sheet(data);
        xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
        
        const fileName = `tencent_${ownerId}_${timestamp}.xlsx`;
        const filePath = path.join(tencentOutputDir, fileName);
        xlsx.writeFile(workbook, filePath);
        
        savedFiles.push({
            name: fileName,
            path: filePath,
            ownerId: ownerId,
            recordCount: data.length
        });
    }

    return savedFiles;
};

// Azure 報表處理路由
app.post('/api/process-azure', async (req, res) => {
    try {
        if (!req.files || !req.files.file) {
            return res.status(400).json({
                success: false,
                message: '未上傳文件'
            });
        }

        const file = req.files.file;
        
        // 驗證文件類型
        if (!validateFileType(file, ['.xlsx'])) {
            return res.status(400).json({
                success: false,
                message: '請上傳 .xlsx 格式的文件'
            });
        }

        console.log(`Processing Azure file: ${file.name}`);
        
        // 讀取Excel文件
        const workbook = xlsx.read(file.data);
        const sheetName = workbook.SheetNames[0];
        const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

        if (data.length === 0) {
            return res.status(400).json({
                success: false,
                message: '文件中沒有有效數據'
            });
        }

        // 處理數據
        const processedData = processAzureData(data);
        
        // 保存處理後的文件
        const savedFiles = saveAzureData(processedData);

        res.json({
            success: true,
            message: `Azure 報表處理成功，生成了 ${savedFiles.length} 個文件`,
            files: savedFiles,
            summary: {
                totalCustomers: savedFiles.length,
                totalRecords: savedFiles.reduce((sum, file) => sum + file.recordCount, 0),
                totalAmount: savedFiles.reduce((sum, file) => sum + file.total, 0)
            }
        });
    } catch (error) {
        console.error('Azure processing error:', error);
        res.status(500).json({
            success: false,
            message: '處理失敗：' + error.message
        });
    }
});

// 騰訊雲報表處理路由
app.post('/api/process-tencent', async (req, res) => {
    try {
        if (!req.files || !req.files.file) {
            return res.status(400).json({
                success: false,
                message: '未上傳文件'
            });
        }

        const file = req.files.file;
        
        // 驗證文件類型
        if (!validateFileType(file, ['.xlsx', '.csv'])) {
            return res.status(400).json({
                success: false,
                message: '請上傳 .xlsx 或 .csv 格式的文件'
            });
        }

        console.log(`Processing Tencent file: ${file.name}`);
        
        // 讀取文件
        let data;
        if (file.name.toLowerCase().endsWith('.csv')) {
            // 處理CSV文件
            const csvData = file.data.toString('utf8');
            const workbook = xlsx.read(csvData, { type: 'string' });
            const sheetName = workbook.SheetNames[0];
            data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
        } else {
            // 處理Excel文件
            const workbook = xlsx.read(file.data);
            const sheetName = workbook.SheetNames[0];
            data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
        }

        if (data.length === 0) {
            return res.status(400).json({
                success: false,
                message: '文件中沒有有效數據'
            });
        }

        // 處理數據
        const processedData = processTencentData(data);
        
        // 保存處理後的文件
        const savedFiles = saveTencentData(processedData);

        res.json({
            success: true,
            message: `騰訊雲報表處理成功，生成了 ${savedFiles.length} 個文件`,
            files: savedFiles,
            summary: {
                totalOwners: savedFiles.length,
                totalRecords: savedFiles.reduce((sum, file) => sum + file.recordCount, 0)
            }
        });
    } catch (error) {
        console.error('Tencent processing error:', error);
        res.status(500).json({
            success: false,
            message: '處理失敗：' + error.message
        });
    }
});

// 獲取處理後文件列表的路由
app.get('/api/processed-files', (req, res) => {
    try {
        const azureFiles = fs.readdirSync(azureOutputDir).map(file => ({
            name: file,
            type: 'azure',
            path: `/api/download/azure/${file}`,
            size: fs.statSync(path.join(azureOutputDir, file)).size,
            createdAt: fs.statSync(path.join(azureOutputDir, file)).mtime
        }));

        const tencentFiles = fs.readdirSync(tencentOutputDir).map(file => ({
            name: file,
            type: 'tencent',
            path: `/api/download/tencent/${file}`,
            size: fs.statSync(path.join(tencentOutputDir, file)).size,
            createdAt: fs.statSync(path.join(tencentOutputDir, file)).mtime
        }));

        res.json({
            success: true,
            files: {
                azure: azureFiles,
                tencent: tencentFiles
            },
            summary: {
                totalFiles: azureFiles.length + tencentFiles.length,
                azureCount: azureFiles.length,
                tencentCount: tencentFiles.length
            }
        });
    } catch (error) {
        console.error('Get files error:', error);
        res.status(500).json({
            success: false,
            message: '獲取文件列表失敗'
        });
    }
});

// 文件下載路由
app.get('/api/download/azure/:filename', (req, res) => {
    const filePath = path.join(azureOutputDir, req.params.filename);
    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).json({ success: false, message: '文件不存在' });
    }
});

app.get('/api/download/tencent/:filename', (req, res) => {
    const filePath = path.join(tencentOutputDir, req.params.filename);
    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).json({ success: false, message: '文件不存在' });
    }
});

// 批量下載路由
app.get('/api/download-all/:type', async (req, res) => {
    try {
        const type = req.params.type; // 'azure', 'tencent', or 'all'
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').split('T')[0];
        
        let zipFileName, sourceDir, files;
        
        if (type === 'azure') {
            zipFileName = `azure_reports_${timestamp}.zip`;
            files = fs.readdirSync(azureOutputDir).map(file => ({
                path: path.join(azureOutputDir, file),
                name: `Azure/${file}`
            }));
        } else if (type === 'tencent') {
            zipFileName = `tencent_reports_${timestamp}.zip`;
            files = fs.readdirSync(tencentOutputDir).map(file => ({
                path: path.join(tencentOutputDir, file),
                name: `Tencent/${file}`
            }));
        } else if (type === 'all') {
            zipFileName = `all_reports_${timestamp}.zip`;
            files = [
                ...fs.readdirSync(azureOutputDir).map(file => ({
                    path: path.join(azureOutputDir, file),
                    name: `Azure/${file}`
                })),
                ...fs.readdirSync(tencentOutputDir).map(file => ({
                    path: path.join(tencentOutputDir, file),
                    name: `Tencent/${file}`
                }))
            ];
        } else {
            return res.status(400).json({
                success: false,
                message: '無效的下載類型'
            });
        }

        if (files.length === 0) {
            return res.status(404).json({
                success: false,
                message: '沒有可下載的文件'
            });
        }

        const zipFilePath = path.join(downloadsDir, zipFileName);
        const output = fs.createWriteStream(zipFilePath);
        const archive = archiver('zip', { zlib: { level: 9 } });

        output.on('close', () => {
            console.log(`Archive created: ${zipFilePath} (${archive.pointer()} bytes)`);
            res.download(zipFilePath, zipFileName, (err) => {
                if (err) {
                    console.error('Download error:', err);
                }
                // 清理臨時文件
                setTimeout(() => {
                    if (fs.existsSync(zipFilePath)) {
                        fs.unlinkSync(zipFilePath);
                    }
                }, 5000);
            });
        });

        archive.on('error', (err) => {
            console.error('Archive error:', err);
            res.status(500).json({
                success: false,
                message: '創建壓縮文件失敗'
            });
        });

        archive.pipe(output);

        files.forEach(file => {
            if (fs.existsSync(file.path)) {
                archive.file(file.path, { name: file.name });
            }
        });

        await archive.finalize();
    } catch (error) {
        console.error('Batch download error:', error);
        res.status(500).json({
            success: false,
            message: '批量下載失敗：' + error.message
        });
    }
});

// 健康檢查路由
app.get('/api/health', (req, res) => {
    res.json({
        success: true,
        message: 'Server is running',
        timestamp: new Date().toISOString(),
        version: '2.0.0'
    });
});

// 使用錯誤處理中間件
app.use(errorHandler);

// 啟動服務器
app.listen(port, () => {
    console.log(`
🚀 Excel 報表處理系統後端服務器已啟動
📍 端口: ${port}
🌐 健康檢查: http://localhost:${port}/api/health
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

