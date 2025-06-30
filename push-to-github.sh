#!/bin/bash
# GitHub 推送自動化腳本
# 文件名：push-to-github.sh
# 版本：1.0
# 作者：Manus AI

set -e  # 遇到錯誤立即退出

# 顏色定義
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 項目信息
PROJECT_NAME="Excel 報表處理系統"
PROJECT_VERSION="v2.0"
REPO_URL="https://github.com/MaxFung76/Report-project.git"

# 函數：打印帶顏色的消息
print_message() {
    local color=$1
    local message=$2
    echo -e "${color}${message}${NC}"
}

# 函數：打印標題
print_title() {
    echo
    print_message $BLUE "=== $1 ==="
    echo
}

# 函數：檢查命令是否存在
check_command() {
    if ! command -v $1 &> /dev/null; then
        print_message $RED "錯誤：未找到命令 '$1'"
        exit 1
    fi
}

# 函數：確認操作
confirm_action() {
    local message=$1
    local default=${2:-"N"}
    
    if [ "$default" = "Y" ]; then
        read -p "$message (Y/n): " -n 1 -r
        echo
        [[ $REPLY =~ ^[Nn]$ ]] && return 1 || return 0
    else
        read -p "$message (y/N): " -n 1 -r
        echo
        [[ $REPLY =~ ^[Yy]$ ]] && return 0 || return 1
    fi
}

# 主函數開始
main() {
    print_title "GitHub 推送自動化腳本"
    print_message $BLUE "項目：$PROJECT_NAME $PROJECT_VERSION"
    print_message $BLUE "時間：$(date '+%Y-%m-%d %H:%M:%S')"
    echo

    # 檢查必要命令
    print_message $YELLOW "檢查系統環境..."
    check_command git
    check_command curl
    
    # 檢查當前目錄
    if [ ! -d ".git" ]; then
        print_message $RED "錯誤：當前目錄不是Git倉庫"
        print_message $YELLOW "請在項目根目錄運行此腳本"
        exit 1
    fi
    
    # 檢查Git配置
    print_message $YELLOW "檢查Git配置..."
    if [ -z "$(git config user.name)" ] || [ -z "$(git config user.email)" ]; then
        print_message $RED "錯誤：Git用戶信息未配置"
        print_message $YELLOW "請先配置Git用戶信息："
        echo "  git config user.name \"Your Name\""
        echo "  git config user.email \"your.email@example.com\""
        exit 1
    fi
    
    print_message $GREEN "Git用戶：$(git config user.name) <$(git config user.email)>"
    
    # 檢查網絡連接
    print_message $YELLOW "檢查網絡連接..."
    if ! curl -s --connect-timeout 5 https://github.com > /dev/null; then
        print_message $RED "警告：無法連接到GitHub"
        if ! confirm_action "是否繼續？"; then
            exit 1
        fi
    else
        print_message $GREEN "網絡連接正常"
    fi
    
    # 檢查工作目錄狀態
    print_message $YELLOW "檢查工作目錄狀態..."
    if [ -n "$(git status --porcelain)" ]; then
        print_message $YELLOW "工作目錄有未提交的更改："
        git status --short
        echo
        if confirm_action "是否先提交這些更改？"; then
            git add .
            read -p "請輸入提交信息: " commit_msg
            if [ -n "$commit_msg" ]; then
                git commit -m "$commit_msg"
                print_message $GREEN "更改已提交"
            else
                print_message $RED "提交信息不能為空"
                exit 1
            fi
        else
            if ! confirm_action "是否忽略未提交的更改繼續推送？"; then
                print_message $YELLOW "推送已取消"
                exit 0
            fi
        fi
    else
        print_message $GREEN "工作目錄乾淨"
    fi
    
    # 檢查遠程倉庫配置
    print_message $YELLOW "檢查遠程倉庫配置..."
    if ! git remote get-url origin &> /dev/null; then
        print_message $RED "錯誤：未配置origin遠程倉庫"
        exit 1
    fi
    
    local remote_url=$(git remote get-url origin)
    print_message $GREEN "遠程倉庫：$remote_url"
    
    # 檢查分支狀態
    print_message $YELLOW "檢查分支狀態..."
    local current_branch=$(git branch --show-current)
    local commits_ahead=$(git rev-list --count origin/$current_branch..$current_branch 2>/dev/null || echo "unknown")
    
    print_message $GREEN "當前分支：$current_branch"
    
    if [ "$commits_ahead" = "unknown" ]; then
        print_message $YELLOW "無法確定分支狀態，可能是新分支"
    elif [ "$commits_ahead" -eq 0 ]; then
        print_message $YELLOW "本地分支與遠程同步，無需推送"
        if ! confirm_action "是否強制推送？"; then
            exit 0
        fi
    else
        print_message $GREEN "本地分支領先遠程 $commits_ahead 個提交"
    fi
    
    # 顯示將要推送的提交
    print_title "將要推送的提交"
    if [ "$commits_ahead" != "unknown" ] && [ "$commits_ahead" -gt 0 ]; then
        git log --oneline origin/$current_branch..$current_branch
    else
        print_message $YELLOW "顯示最近的提交："
        git log --oneline -5
    fi
    echo
    
    # 最終確認
    print_title "推送確認"
    print_message $YELLOW "準備推送到：$remote_url"
    print_message $YELLOW "分支：$current_branch"
    
    if ! confirm_action "確認執行推送操作？"; then
        print_message $YELLOW "推送已取消"
        exit 0
    fi
    
    # 執行推送（這裡只是模擬，實際需要認證）
    print_title "執行推送"
    print_message $YELLOW "正在推送到GitHub..."
    
    # 注意：實際推送需要認證信息
    print_message $YELLOW "注意：實際推送需要GitHub認證信息"
    print_message $YELLOW "請按照以下步驟完成推送："
    echo
    print_message $BLUE "1. 設置Personal Access Token："
    echo "   git config credential.helper store"
    echo "   git push origin $current_branch"
    echo "   # 然後輸入GitHub用戶名和Token"
    echo
    print_message $BLUE "2. 或使用SSH密鑰："
    echo "   git remote set-url origin git@github.com:MaxFung76/Report-project.git"
    echo "   git push origin $current_branch"
    echo
    print_message $BLUE "3. 或在URL中包含Token："
    echo "   git remote set-url origin https://username:token@github.com/MaxFung76/Report-project.git"
    echo "   git push origin $current_branch"
    echo
    
    # 模擬推送成功
    if confirm_action "模擬推送成功，是否繼續驗證流程？"; then
        print_title "推送後驗證"
        
        print_message $YELLOW "請手動驗證以下項目："
        echo "□ GitHub頁面顯示最新提交"
        echo "□ 文件結構正確"
        echo "□ 提交信息準確"
        echo "□ 所有文件都已上傳"
        echo
        
        print_message $GREEN "推送腳本執行完成！"
        print_message $BLUE "請訪問 https://github.com/MaxFung76/Report-project 確認推送結果"
    fi
}

# 執行主函數
main "$@"

