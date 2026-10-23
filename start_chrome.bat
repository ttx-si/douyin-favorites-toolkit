@echo off
chcp 65001 >nul
echo 正在启动具备远程调试端口 (9223) 的独立 Chrome 浏览器...

set CHROME_PATH="C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist %CHROME_PATH% (
    set CHROME_PATH="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
)

set PROFILE_DIR="%USERPROFILE%\.douyin-scraper-chrome-profile"

start "" %CHROME_PATH% --remote-debugging-port=9223 --user-data-dir=%PROFILE_DIR% "https://www.douyin.com"

echo Chrome 已拉起，请在打开的窗口中登录抖音并进入你的“收藏”页面。
echo 登录完成后，即可运行: python run.py
pause
