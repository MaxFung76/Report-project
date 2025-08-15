import { useState, useEffect } from 'react'
import { Button } from '@/components/ui/button.jsx'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card.jsx'
import { Progress } from '@/components/ui/progress.jsx'
import { Badge } from '@/components/ui/badge.jsx'
import { Alert, AlertDescription } from '@/components/ui/alert.jsx'
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs.jsx'
import { 
  Upload, 
  FileSpreadsheet, 
  Download, 
  CheckCircle, 
  AlertCircle, 
  Loader2,
  Cloud,
  FileText,
  BarChart3,
  Settings,
  Moon,
  Sun
} from 'lucide-react'
import './App.css'

function App() {
  const [isDark, setIsDark] = useState(false)
  const [files, setFiles] = useState([])
  const [uploadProgress, setUploadProgress] = useState({})
  const [processedFiles, setProcessedFiles] = useState({ azure: [], tencent: [] })
  const [alerts, setAlerts] = useState([])

  // 切換暗色模式
  const toggleDarkMode = () => {
    setIsDark(!isDark)
    document.documentElement.classList.toggle('dark')
  }

  // 添加提醒消息
  const addAlert = (message, type = 'info') => {
    const id = Date.now()
    setAlerts(prev => [...prev, { id, message, type }])
    setTimeout(() => {
      setAlerts(prev => prev.filter(alert => alert.id !== id))
    }, 5000)
  }

  // 文件上傳處理
  const handleFileUpload = async (file, type) => {
    const fileId = Date.now()
    setUploadProgress(prev => ({ ...prev, [fileId]: 0 }))
    
    try {
      const formData = new FormData()
      formData.append('file', file)
      
      // 顯示上傳進度
      setUploadProgress(prev => ({ ...prev, [fileId]: 20 }))
      
      const endpoint = type === 'azure' ? 
        '/api/process-azure' : 
        '/api/process-tencent'
      
      const response = await fetch(endpoint, {
        method: 'POST',
        body: formData
      })
      
      setUploadProgress(prev => ({ ...prev, [fileId]: 80 }))
      
      if (!response.ok) {
        const errorData = await response.json()
        throw new Error(errorData.message || '處理失敗')
      }
      
      const result = await response.json()
      setUploadProgress(prev => ({ ...prev, [fileId]: 100 }))
      
      // 更新文件列表
      await loadProcessedFiles()
      
      addAlert(`${type === 'azure' ? 'Azure' : '騰訊雲'}報表處理完成：${result.message}`, 'success')
    } catch (error) {
      console.error('Upload error:', error)
      addAlert(`處理失敗: ${error.message}`, 'error')
    } finally {
      setTimeout(() => {
        setUploadProgress(prev => {
          const newProgress = { ...prev }
          delete newProgress[fileId]
          return newProgress
        })
      }, 1000)
    }
  }

  // 載入處理後的文件列表
  const loadProcessedFiles = async () => {
    try {
      const response = await fetch('/api/processed-files')
      if (response.ok) {
        const data = await response.json()
        if (data.success) {
          setProcessedFiles({
            azure: data.files.azure.map(file => ({
              id: file.name,
              name: file.name,
              type: 'azure',
              size: file.size,
              processedAt: file.createdAt,
              status: 'completed',
              downloadUrl: `${file.path}`
            })),
            tencent: data.files.tencent.map(file => ({
              id: file.name,
              name: file.name,
              type: 'tencent',
              size: file.size,
              processedAt: file.createdAt,
              status: 'completed',
              downloadUrl: `${file.path}`
            }))
          })
        }
      }
    } catch (error) {
      console.error('Load files error:', error)
    }
  }

  // 下載文件
  const downloadFile = async (url, filename) => {
    try {
      const link = document.createElement('a')
      link.href = url
      link.download = filename
      link.target = '_blank'
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
    } catch (error) {
      console.error('Download error:', error)
      addAlert('下載失敗', 'error')
    }
  }

  // 批量下載
  const downloadAllFiles = async (type) => {
    try {
      const url = `/api/download-all/${type}`
      const link = document.createElement('a')
      link.href = url
      link.target = '_blank'
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
      addAlert('開始下載...', 'success')
    } catch (error) {
      console.error('Batch download error:', error)
      addAlert('批量下載失敗', 'error')
    }
  }

  // 組件載入時獲取文件列表
  useEffect(() => {
    loadProcessedFiles()
    // 定期更新文件列表
    const interval = setInterval(loadProcessedFiles, 10000)
    return () => clearInterval(interval)
  }, [])

  // 文件拖拽處理
  const handleDrop = (e, type) => {
    e.preventDefault()
    const droppedFiles = Array.from(e.dataTransfer.files)
    droppedFiles.forEach(file => handleFileUpload(file, type))
  }

  const handleDragOver = (e) => {
    e.preventDefault()
  }

  // 格式化文件大小
  const formatFileSize = (bytes) => {
    if (bytes === 0) return '0 Bytes'
    const k = 1024
    const sizes = ['Bytes', 'KB', 'MB', 'GB']
    const i = Math.floor(Math.log(bytes) / Math.log(k))
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i]
  }

  return (
    <div className={`min-h-screen bg-gradient-to-br from-slate-50 to-blue-50 dark:from-slate-900 dark:to-slate-800 transition-colors duration-300`}>
      {/* 頭部導航 */}
      <header className="bg-white/80 dark:bg-slate-900/80 backdrop-blur-sm border-b border-slate-200 dark:border-slate-700 sticky top-0 z-50">
        <div className="container mx-auto px-4 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              <div className="p-2 bg-blue-600 rounded-lg">
                <BarChart3 className="h-6 w-6 text-white" />
              </div>
              <div>
                <h1 className="text-2xl font-bold text-slate-900 dark:text-white">
                  Excel 報表處理系統
                </h1>
                <p className="text-sm text-slate-600 dark:text-slate-400">
                  現代化雲端報表管理平台
                </p>
              </div>
            </div>
            <div className="flex items-center space-x-4">
              <Button
                variant="ghost"
                size="icon"
                onClick={toggleDarkMode}
                className="rounded-full"
              >
                {isDark ? <Sun className="h-5 w-5" /> : <Moon className="h-5 w-5" />}
              </Button>
              <Button variant="outline" size="sm">
                <Settings className="h-4 w-4 mr-2" />
                設定
              </Button>
            </div>
          </div>
        </div>
      </header>

      {/* 提醒消息 */}
      <div className="fixed top-20 right-4 z-50 space-y-2">
        {alerts.map(alert => (
          <Alert 
            key={alert.id} 
            className={`animate-slide-in-right ${
              alert.type === 'success' ? 'border-green-500 bg-green-50 dark:bg-green-950' :
              alert.type === 'error' ? 'border-red-500 bg-red-50 dark:bg-red-950' :
              'border-blue-500 bg-blue-50 dark:bg-blue-950'
            }`}
          >
            {alert.type === 'success' ? <CheckCircle className="h-4 w-4 text-green-600" /> :
             alert.type === 'error' ? <AlertCircle className="h-4 w-4 text-red-600" /> :
             <AlertCircle className="h-4 w-4 text-blue-600" />}
            <AlertDescription>{alert.message}</AlertDescription>
          </Alert>
        ))}
      </div>

      <main className="container mx-auto px-4 py-8">
        {/* 統計卡片 */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
          <Card className="bg-gradient-to-r from-blue-500 to-blue-600 text-white border-0">
            <CardHeader className="pb-2">
              <CardTitle className="text-lg font-medium">Azure 報表</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="text-3xl font-bold">{processedFiles.azure.length}</div>
              <p className="text-blue-100 text-sm">已處理文件</p>
            </CardContent>
          </Card>
          <Card className="bg-gradient-to-r from-green-500 to-green-600 text-white border-0">
            <CardHeader className="pb-2">
              <CardTitle className="text-lg font-medium">騰訊雲報表</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="text-3xl font-bold">{processedFiles.tencent.length}</div>
              <p className="text-green-100 text-sm">已處理文件</p>
            </CardContent>
          </Card>
          <Card className="bg-gradient-to-r from-purple-500 to-purple-600 text-white border-0">
            <CardHeader className="pb-2">
              <CardTitle className="text-lg font-medium">總計</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="text-3xl font-bold">
                {processedFiles.azure.length + processedFiles.tencent.length}
              </div>
              <p className="text-purple-100 text-sm">處理文件總數</p>
            </CardContent>
          </Card>
        </div>

        {/* 主要內容區域 */}
        <Tabs defaultValue="upload" className="space-y-6">
          <TabsList className="grid w-full grid-cols-2 lg:w-[400px]">
            <TabsTrigger value="upload" className="flex items-center space-x-2">
              <Upload className="h-4 w-4" />
              <span>文件上傳</span>
            </TabsTrigger>
            <TabsTrigger value="files" className="flex items-center space-x-2">
              <FileText className="h-4 w-4" />
              <span>文件管理</span>
            </TabsTrigger>
          </TabsList>

          {/* 文件上傳標籤頁 */}
          <TabsContent value="upload" className="space-y-6">
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              {/* Azure 上傳區域 */}
              <Card className="hover:shadow-lg transition-shadow duration-300">
                <CardHeader>
                  <div className="flex items-center space-x-3">
                    <div className="p-2 bg-blue-100 dark:bg-blue-900 rounded-lg">
                      <Cloud className="h-5 w-5 text-blue-600 dark:text-blue-400" />
                    </div>
                    <div>
                      <CardTitle>Azure 報表處理</CardTitle>
                      <CardDescription>
                        上傳 Azure 雲端服務賬單報表進行處理
                      </CardDescription>
                    </div>
                  </div>
                </CardHeader>
                <CardContent>
                  <div
                    className="border-2 border-dashed border-slate-300 dark:border-slate-600 rounded-lg p-8 text-center hover:border-blue-400 dark:hover:border-blue-500 transition-colors duration-300 cursor-pointer"
                    onDrop={(e) => handleDrop(e, 'azure')}
                    onDragOver={handleDragOver}
                  >
                    <FileSpreadsheet className="h-12 w-12 text-slate-400 mx-auto mb-4" />
                    <p className="text-slate-600 dark:text-slate-400 mb-2">
                      拖拽 Excel 文件到此處或點擊選擇
                    </p>
                    <Button 
                      variant="outline" 
                      className="mb-4"
                      onClick={() => document.getElementById('azure-file-input').click()}
                    >
                      <Upload className="h-4 w-4 mr-2" />
                      選擇文件
                    </Button>
                    <input
                      id="azure-file-input"
                      type="file"
                      accept=".xlsx"
                      style={{ display: 'none' }}
                      onChange={(e) => {
                        const file = e.target.files[0]
                        if (file) handleFileUpload(file, 'azure')
                      }}
                    />
                    <p className="text-xs text-slate-500 dark:text-slate-400">
                      支持 .xlsx 格式，最大 10MB
                    </p>
                  </div>
                </CardContent>
              </Card>

              {/* 騰訊雲上傳區域 */}
              <Card className="hover:shadow-lg transition-shadow duration-300">
                <CardHeader>
                  <div className="flex items-center space-x-3">
                    <div className="p-2 bg-green-100 dark:bg-green-900 rounded-lg">
                      <Cloud className="h-5 w-5 text-green-600 dark:text-green-400" />
                    </div>
                    <div>
                      <CardTitle>騰訊雲報表處理</CardTitle>
                      <CardDescription>
                        上傳騰訊雲服務賬單報表進行處理
                      </CardDescription>
                    </div>
                  </div>
                </CardHeader>
                <CardContent>
                  <div
                    className="border-2 border-dashed border-slate-300 dark:border-slate-600 rounded-lg p-8 text-center hover:border-green-400 dark:hover:border-green-500 transition-colors duration-300 cursor-pointer"
                    onDrop={(e) => handleDrop(e, 'tencent')}
                    onDragOver={handleDragOver}
                  >
                    <FileSpreadsheet className="h-12 w-12 text-slate-400 mx-auto mb-4" />
                    <p className="text-slate-600 dark:text-slate-400 mb-2">
                      拖拽 Excel/CSV 文件到此處或點擊選擇
                    </p>
                    <Button 
                      variant="outline" 
                      className="mb-4"
                      onClick={() => document.getElementById('tencent-file-input').click()}
                    >
                      <Upload className="h-4 w-4 mr-2" />
                      選擇文件
                    </Button>
                    <input
                      id="tencent-file-input"
                      type="file"
                      accept=".xlsx,.csv"
                      style={{ display: 'none' }}
                      onChange={(e) => {
                        const file = e.target.files[0]
                        if (file) handleFileUpload(file, 'tencent')
                      }}
                    />
                    <p className="text-xs text-slate-500 dark:text-slate-400">
                      支持 .xlsx, .csv 格式，最大 10MB
                    </p>
                  </div>
                </CardContent>
              </Card>
            </div>

            {/* 上傳進度 */}
            {Object.keys(uploadProgress).length > 0 && (
              <Card>
                <CardHeader>
                  <CardTitle className="flex items-center space-x-2">
                    <Loader2 className="h-5 w-5 animate-spin" />
                    <span>處理進度</span>
                  </CardTitle>
                </CardHeader>
                <CardContent className="space-y-4">
                  {Object.entries(uploadProgress).map(([fileId, progress]) => (
                    <div key={fileId} className="space-y-2">
                      <div className="flex justify-between text-sm">
                        <span>處理中...</span>
                        <span>{progress}%</span>
                      </div>
                      <Progress value={progress} className="h-2" />
                    </div>
                  ))}
                </CardContent>
              </Card>
            )}
          </TabsContent>

          {/* 文件管理標籤頁 */}
          <TabsContent value="files" className="space-y-6">
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              {/* Azure 文件列表 */}
              <Card>
                <CardHeader>
                  <div className="flex items-center justify-between">
                    <div className="flex items-center space-x-3">
                      <div className="p-2 bg-blue-100 dark:bg-blue-900 rounded-lg">
                        <Cloud className="h-5 w-5 text-blue-600 dark:text-blue-400" />
                      </div>
                      <div>
                        <CardTitle>Azure 報表文件</CardTitle>
                        <CardDescription>
                          {processedFiles.azure.length} 個已處理文件
                        </CardDescription>
                      </div>
                    </div>
                    {processedFiles.azure.length > 0 && (
                      <Button 
                        size="sm" 
                        variant="outline"
                        onClick={() => downloadAllFiles('azure')}
                      >
                        <Download className="h-4 w-4 mr-2" />
                        全部下載
                      </Button>
                    )}
                  </div>
                </CardHeader>
                <CardContent>
                  {processedFiles.azure.length === 0 ? (
                    <div className="text-center py-8 text-slate-500 dark:text-slate-400">
                      <FileText className="h-12 w-12 mx-auto mb-4 opacity-50" />
                      <p>暫無處理完成的 Azure 報表文件</p>
                    </div>
                  ) : (
                    <div className="space-y-3">
                      {processedFiles.azure.map(file => (
                        <div key={file.id} className="flex items-center justify-between p-3 bg-slate-50 dark:bg-slate-800 rounded-lg">
                          <div className="flex items-center space-x-3">
                            <FileSpreadsheet className="h-5 w-5 text-blue-600" />
                            <div>
                              <p className="font-medium text-sm">{file.name}</p>
                              <p className="text-xs text-slate-500">
                                {formatFileSize(file.size)} • {new Date(file.processedAt).toLocaleString()}
                              </p>
                            </div>
                          </div>
                          <div className="flex items-center space-x-2">
                            <Badge variant="secondary" className="bg-green-100 text-green-800 dark:bg-green-900 dark:text-green-200">
                              <CheckCircle className="h-3 w-3 mr-1" />
                              完成
                            </Badge>
                            <Button 
                              size="sm" 
                              variant="ghost"
                              onClick={() => downloadFile(file.downloadUrl, file.name)}
                            >
                              <Download className="h-4 w-4" />
                            </Button>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </CardContent>
              </Card>

              {/* 騰訊雲文件列表 */}
              <Card>
                <CardHeader>
                  <div className="flex items-center justify-between">
                    <div className="flex items-center space-x-3">
                      <div className="p-2 bg-green-100 dark:bg-green-900 rounded-lg">
                        <Cloud className="h-5 w-5 text-green-600 dark:text-green-400" />
                      </div>
                      <div>
                        <CardTitle>騰訊雲報表文件</CardTitle>
                        <CardDescription>
                          {processedFiles.tencent.length} 個已處理文件
                        </CardDescription>
                      </div>
                    </div>
                    {processedFiles.tencent.length > 0 && (
                      <Button 
                        size="sm" 
                        variant="outline"
                        onClick={() => downloadAllFiles('tencent')}
                      >
                        <Download className="h-4 w-4 mr-2" />
                        全部下載
                      </Button>
                    )}
                  </div>
                </CardHeader>
                <CardContent>
                  {processedFiles.tencent.length === 0 ? (
                    <div className="text-center py-8 text-slate-500 dark:text-slate-400">
                      <FileText className="h-12 w-12 mx-auto mb-4 opacity-50" />
                      <p>暫無處理完成的騰訊雲報表文件</p>
                    </div>
                  ) : (
                    <div className="space-y-3">
                      {processedFiles.tencent.map(file => (
                        <div key={file.id} className="flex items-center justify-between p-3 bg-slate-50 dark:bg-slate-800 rounded-lg">
                          <div className="flex items-center space-x-3">
                            <FileSpreadsheet className="h-5 w-5 text-green-600" />
                            <div>
                              <p className="font-medium text-sm">{file.name}</p>
                              <p className="text-xs text-slate-500">
                                {formatFileSize(file.size)} • {new Date(file.processedAt).toLocaleString()}
                              </p>
                            </div>
                          </div>
                          <div className="flex items-center space-x-2">
                            <Badge variant="secondary" className="bg-green-100 text-green-800 dark:bg-green-900 dark:text-green-200">
                              <CheckCircle className="h-3 w-3 mr-1" />
                              完成
                            </Badge>
                            <Button 
                              size="sm" 
                              variant="ghost"
                              onClick={() => downloadFile(file.downloadUrl, file.name)}
                            >
                              <Download className="h-4 w-4" />
                            </Button>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </CardContent>
              </Card>
            </div>
          </TabsContent>
        </Tabs>
      </main>

      {/* 頁腳 */}
      <footer className="bg-white/80 dark:bg-slate-900/80 backdrop-blur-sm border-t border-slate-200 dark:border-slate-700 mt-16">
        <div className="container mx-auto px-4 py-6">
          <div className="flex items-center justify-between">
            <p className="text-sm text-slate-600 dark:text-slate-400">
              © 2024 Excel 報表處理系統. 現代化雲端報表管理平台
            </p>
            <div className="flex items-center space-x-4 text-sm text-slate-600 dark:text-slate-400">
              <span>版本 2.0</span>
              <span>•</span>
              <span>技術支援</span>
            </div>
          </div>
        </div>
      </footer>
    </div>
  )
}

export default App

