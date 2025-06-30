const xlsx = require('xlsx');
const path = require('path');

// 創建測試用的Azure數據
const azureTestData = [
    {
        'CustomerName': 'Test Company A',
        'UnitPrice': 100,
        'BillableQuantity': 5,
        'Quantity': 5,
        'PartnerId': 'P001',
        'CustomerId': 'C001',
        'InvoiceNumber': 'INV001',
        'Bill to': 'Customer'
    },
    {
        'CustomerName': 'Test Company B',
        'UnitPrice': 200,
        'BillableQuantity': 3,
        'Quantity': 3,
        'PartnerId': 'P002',
        'CustomerId': 'C002',
        'InvoiceNumber': 'INV002',
        'Bill to': 'Customer'
    },
    {
        'CustomerName': 'Test Company C',
        'UnitPrice': 150,
        'BillableQuantity': 4,
        'Quantity': 4,
        'PartnerId': 'P003',
        'CustomerId': 'C003',
        'InvoiceNumber': 'INV003',
        'Bill to': 'Accord'  // 這個應該被過濾掉
    }
];

// 創建測試用的騰訊雲數據
const tencentTestData = [
    {
        'Owner Account ID': '12345678',
        'ProductName': 'CVM',
        'SubproductName': 'Standard Instance',
        'BillingMode': 'Monthly',
        'ProjectName': 'Default Project',
        'Region': 'ap-beijing',
        'InstanceID': 'ins-12345',
        'InstanceName': 'test-instance-1',
        'TransactionType': 'Purchase',
        'TransactionTime': '2024-01-01 10:00:00',
        'Usage Start Time': '2024-01-01 10:00:00',
        'Usage End Time': '2024-01-31 23:59:59',
        'Configuration Description': '2核4GB',
        'OriginalCost': 100.50
    },
    {
        'Owner Account ID': '87654321',
        'ProductName': 'COS',
        'SubproductName': 'Standard Storage',
        'BillingMode': 'Pay-as-you-go',
        'ProjectName': 'Test Project',
        'Region': 'ap-shanghai',
        'InstanceID': 'cos-67890',
        'InstanceName': 'test-bucket',
        'TransactionType': 'Usage',
        'TransactionTime': '2024-01-15 15:30:00',
        'Usage Start Time': '2024-01-01 00:00:00',
        'Usage End Time': '2024-01-31 23:59:59',
        'Configuration Description': '100GB Storage',
        'OriginalCost': 25.75
    }
];

// 創建Azure測試文件
const azureWorkbook = xlsx.utils.book_new();
const azureWorksheet = xlsx.utils.json_to_sheet(azureTestData);
xlsx.utils.book_append_sheet(azureWorkbook, azureWorksheet, 'Azure Data');
xlsx.writeFile(azureWorkbook, path.join(__dirname, 'test_azure_data.xlsx'));

// 創建騰訊雲測試文件
const tencentWorkbook = xlsx.utils.book_new();
const tencentWorksheet = xlsx.utils.json_to_sheet(tencentTestData);
xlsx.utils.book_append_sheet(tencentWorkbook, tencentWorksheet, 'Tencent Data');
xlsx.writeFile(tencentWorkbook, path.join(__dirname, 'test_tencent_data.xlsx'));

console.log('測試文件已創建:');
console.log('- test_azure_data.xlsx');
console.log('- test_tencent_data.xlsx');

