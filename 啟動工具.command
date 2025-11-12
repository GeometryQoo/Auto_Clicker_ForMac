#!/bin/bash
# 自動點擊工具啟動腳本

# 取得腳本所在目錄
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# 切換到專案目錄
cd "$DIR"

# 檢查 Python 是否安裝
if ! command -v python3 &> /dev/null; then
    echo "錯誤: 找不到 Python 3"
    echo "請先安裝 Python 3"
    exit 1
fi

# 檢查依賴套件是否安裝
if ! python3 -c "import pyautogui" &> /dev/null; then
    echo "正在安裝依賴套件..."
    pip3 install pyautogui keyboard
fi

# 啟動程式
echo "正在啟動自動點擊工具..."
python3 auto_clicker.py
