#!/bin/bash

# --- 1. 錯誤立即停止 ---
set -e

# --- 2. 設定變數 (已根據您的資訊填入) ---
PROJECT_USER="daanyoyo"                                           #
PROJECT_DIR="/home/$PROJECT_USER/TCGweb-health-checker"           # 
REPO_URL="https://github.com/dayoxiao/TCGweb-health-checker"      # 個人GitHub 待修改
PYTHON_CMD="python3"
PYTHON_SCRIPT="gcp_main.py"
LOG_FILE="/home/$PROJECT_USER/crawler_startup.log"

# --- 3. 記錄啟動 ---
echo "========================================" >> $LOG_FILE
echo "VM 啟動時間: $(date)" >> $LOG_FILE

# --- 4. 安裝系統套件 ---
echo "更新系統套件..." >> $LOG_FILE
apt-get update -y >> $LOG_FILE 2>&1
echo "安裝系統套件 (git, pip, gcloud, curl)..." >> $LOG_FILE
apt-get install -y python3-pip git curl google-cloud-sdk >> $LOG_FILE 2>&1

# --- 5. 取得/更新程式碼 ---
if [ ! -d "$PROJECT_DIR" ]; then
    echo "複製程式碼從 $REPO_URL..." >> $LOG_FILE
    # 以 PROJECT_USER 身份 clone
    sudo -u $PROJECT_USER git clone "$REPO_URL" "$PROJECT_DIR" >> $LOG_FILE 2>&1
else
    echo "重設並更新現有程式碼..." >> $LOG_FILE
    cd "$PROJECT_DIR"
    sudo -u $PROJECT_USER git reset --hard HEAD >> $LOG_FILE 2>&1
    sudo -u $PROJECT_USER git pull >> $LOG_FILE 2>&1
fi

# --- 6. 安裝 Python 依賴 ---
cd "$PROJECT_DIR"
echo "安裝 Python 依賴套件 (requirements.txt)..." >> $LOG_FILE
sudo -u $PROJECT_USER pip3 install -r requirements.txt >> $LOG_FILE 2>&1

echo "安裝 Playwright 瀏覽器 (chromium)..." >> $LOG_FILE
sudo -u $PROJECT_USER $PYTHON_CMD -m playwright install chromium >> $LOG_FILE 2>&1

# --- 7. 檢查檔案 ---
if [ ! -f "$PYTHON_SCRIPT" ]; then
    echo "錯誤：找不到 $PYTHON_SCRIPT 檔案" >> $LOG_FILE
    exit 1
fi
if [ ! -f "config/websites.csv" ]; then
    echo "錯誤：找不到 config/websites.csv 設定檔" >> $LOG_FILE
    exit 1
fi

# --- 8. 檢查程式是否已在執行 ---
if pgrep -u "$PROJECT_USER" -f "$PYTHON_SCRIPT" > /dev/null; then
    echo "檢查： $PYTHON_SCRIPT 已經在執行。啟動腳本將退出。" >> $LOG_FILE
    echo "啟動腳本執行完成 (未啟動新程序): $(date)" >> $LOG_FILE
    exit 0
fi

# --- 9. 在背景執行爬蟲程式 ---
echo "開始執行網站爬蟲 ($PYTHON_SCRIPT)..." >> $LOG_FILE
CRAWLER_ARGS="--config config/websites.csv --concurrent 2 --no-save-html --no-pagination"
echo "執行指令: $PYTHON_CMD $PYTHON_SCRIPT $CRAWLER_ARGS" >> $LOG_FILE

export GCLOUD_PATH="/usr/bin:/google/google-cloud-sdk/bin"

nohup sudo -u $PROJECT_USER \
    PATH=$PATH:$GCLOUD_PATH \
    $PYTHON_CMD $PYTHON_SCRIPT $CRAWLER_ARGS \
    > /home/$PROJECT_USER/crawler_execution.log 2>&1 &

echo "爬蟲程式已在背景啟動，PID: $!" >> $LOG_FILE
echo "可以查看 /home/$PROJECT_USER/crawler_execution.log 了解執行進度" >> $LOG_FILE
echo "啟動腳本執行完成: $(date)" >> $LOG_FILE