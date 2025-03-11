const express = require('express');
const fileUpload = require('express-fileupload');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const archiver = require('archiver');

// Initialize Express app
const app = express();
const port = 3000;

// 定義所有輸出目錄
const azureOutputDir = path.join(__dirname, 'output(Azure)');
const tencentOutputDir = path.join(__dirname, 'output(Tencent)');
const downloadsDir = path.join(__dirname, 'public', 'downloads');

// 確保所有必要目錄存在
if (!fs.existsSync(azureOutputDir)) fs.mkdirSync(azureOutputDir);
if (!fs.existsSync(tencentOutputDir)) fs.mkdirSync(tencentOutputDir);
if (!fs.existsSync(downloadsDir)) fs.mkdirSync(downloadsDir, { recursive: true });

// Middleware setup
app.use(express.static(path.join(__dirname, 'public')));
app.use(fileUpload());

// 設置根路徑路由
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'views', 'index.html'));
});

// Azure 報表處理路由
app.post('/process-azure', async (req, res) => {
    try {
        if (!req.files || !req.files.file) {
            return res.status(400).json({ message: '未上傳文件' });
        }

        const file = req.files.file;
        const workbook = xlsx.read(file.data);
        const sheetName = workbook.SheetNames[0];
        const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

        // 處理數據
        const processedData = processAzureData(data);
        
        // 保存處理後的文件
        saveAzureData(processedData);

        res.json({ 
            message: 'Azure 報表處理成功'
        });
    } catch (error) {
        res.status(500).json({ message: '處理失敗：' + error.message });
    }
});

// 騰訊雲報表處理路由
app.post('/process-tencent', async (req, res) => {
    try {
        if (!req.files || !req.files.file) {
            return res.status(400).json({ message: '未上傳文件' });
        }

        const file = req.files.file;
        const workbook = xlsx.read(file.data);
        const sheetName = workbook.SheetNames[0];
        const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

        // 處理數據
        const processedData = processTencentData(data);
        
        // 保存處理後的文件
        saveTencentData(processedData);

        res.json({ 
            message: '騰訊雲報表處理成功'
        });
    } catch (error) {
        res.status(500).json({ message: '處理失敗：' + error.message });
    }
});

// 添加下載目錄的靜態服務
app.use('/downloads', express.static(downloadsDir));

// 添加獲取處理後文件列表的路由
app.get('/processed-files', (req, res) => {
    const azureFiles = fs.readdirSync(azureOutputDir).map(file => ({
        name: file,
        type: 'Azure',
        path: `/downloads/azure/${file}`
    }));

    const tencentFiles = fs.readdirSync(tencentOutputDir).map(file => ({
        name: file,
        type: 'Tencent',
        path: `/downloads/tencent/${file}`
    }));

    res.json([...azureFiles, ...tencentFiles]);
});

// 添加下載全部文件的路由
app.get('/download-all', async (req, res) => {
    try {
        // 創建一個唯一的文件名
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const zipFileName = `all_reports_${timestamp}.zip`;
        const zipFilePath = path.join(downloadsDir, zipFileName);
        
        // 創建一個可寫流
        const output = fs.createWriteStream(zipFilePath);
        const archive = archiver('zip', {
            zlib: { level: 9 } // 設置壓縮級別
        });
        
        // 監聽所有存檔數據寫入完成
        output.on('close', () => {
            console.log(`${archive.pointer()} 總字節數`); 
            console.log('存檔已完成並且輸出文件已關閉');
            res.json({
                success: true,
                message: '所有文件已打包完成',
                zipPath: `/downloads/${zipFileName}`
            });
        });
        
        // 監聽警告
        archive.on('warning', (err) => {
            if (err.code === 'ENOENT') {
                console.warn('警告:', err);
            } else {
                throw err;
            }
        });
        
        // 監聽錯誤
        archive.on('error', (err) => {
            throw err;
        });
        
        // 將存檔數據通過管道傳輸到文件
        archive.pipe(output);
        
        // 添加 Azure 報表文件到存檔
        const azureFiles = fs.readdirSync(azureOutputDir);
        azureFiles.forEach(file => {
            const filePath = path.join(azureOutputDir, file);
            archive.file(filePath, { name: `Azure/${file}` });
        });
        
        // 添加騰訊雲報表文件到存檔
        const tencentFiles = fs.readdirSync(tencentOutputDir);
        tencentFiles.forEach(file => {
            const filePath = path.join(tencentOutputDir, file);
            archive.file(filePath, { name: `Tencent/${file}` });
        });
        
        // 完成存檔
        await archive.finalize();
    } catch (error) {
        console.error('打包文件時發生錯誤:', error);
        res.status(500).json({
            success: false,
            message: '打包文件失敗: ' + error.message
        });
    }
});

// 添加下載ZIP文件的路由
app.get('/downloads/:filename', (req, res) => {
    const filePath = path.join(downloadsDir, req.params.filename);
    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).json({ message: '文件不存在' });
    }
});

// 修改文件下載路由
app.get('/downloads/azure/:filename', (req, res) => {
    const filePath = path.join(azureOutputDir, req.params.filename);
    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).json({ message: '文件不存在' });
    }
});

app.get('/downloads/tencent/:filename', (req, res) => {
    const filePath = path.join(tencentOutputDir, req.params.filename);
    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).json({ message: '文件不存在' });
    }
});

// Start the server
app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});

function processTencentData(data) {
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
        // 添加 Discount Multiplier 和计算 Total Cost
        newRow['Discount Multiplier'] = 1;
        newRow['Total Cost'] = newRow['OriginalCost'] * newRow['Discount Multiplier'];
        return newRow;
    });

    // 按 Owner Account ID 分组
    return processedData.reduce((acc, row) => {
        const ownerId = row['Owner Account ID'];
        if (!acc[ownerId]) acc[ownerId] = [];
        acc[ownerId].push(row);
        return acc;
    }, {});
}

function processAzureData(data) {
    // 删除指定列
    const columnsToDelete = ['PartnerId', 'CustomerId', 'InvoiceNumber', 'MpnId', 'Bill to', 'PriceAdjustmentDescription', 'EffectiveUnitPrice'];
    
    // 过滤数据
    const filteredData = data
        .filter(row => row.CustomerName && row.Quantity) // 删除 Quantity 为空的行
        .filter(row => row['Bill to'] !== 'Accord')      // 排除 Bill to 为 Accord 的数据
        .map(row => {
            // 删除不需要的列
            columnsToDelete.forEach(col => delete row[col]);
            
            // 计算 Subtotal 和 Total
            row.Subtotal = row.UnitPrice * row.BillableQuantity;
            row.Total = row.UnitPrice * row.BillableQuantity;
            return row;
        });

    // 按客户分组
    return filteredData.reduce((acc, row) => {
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
}

function saveAzureData(processedData) {
    const now = new Date();
    const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const sheetName = `${lastMonth.toLocaleString('en', { month: 'short' })}_${lastMonth.getFullYear()}`;

    for (const [customer, data] of Object.entries(processedData)) {
        const workbook = xlsx.utils.book_new();
        
        // 添加所有行
        const worksheet = xlsx.utils.json_to_sheet(data.rows);
        
        // 添加总计行
        const totalRow = { Total: data.total };
        const totalRowPos = data.rows.length + 2; // 空一行后添加总计
        xlsx.utils.sheet_add_json(worksheet, [totalRow], {
            skipHeader: true,
            origin: `A${totalRowPos}`
        });
        
        xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
        
        const filePath = path.join(azureOutputDir, `${customer}.xlsx`);
        xlsx.writeFile(workbook, filePath);
    }
}

function saveTencentData(processedData) {
    const now = new Date();
    const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const sheetName = `${lastMonth.toLocaleString('en', { month: 'short' })}_${lastMonth.getFullYear()}`;

    for (const [ownerId, data] of Object.entries(processedData)) {
        const workbook = xlsx.utils.book_new();
        const worksheet = xlsx.utils.json_to_sheet(data);
        xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
        
        const filePath = path.join(tencentOutputDir, `output_${ownerId}.xlsx`);
        xlsx.writeFile(workbook, filePath);
    }
}