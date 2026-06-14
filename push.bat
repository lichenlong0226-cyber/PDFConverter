@echo off
cd /d "C:\Users\liche\Documents\Codex\temp-pdf-repo"
echo 正在推送到 GitHub...
git push origin main
if %errorlevel% neq 0 (
    echo 推送失败，可能网络或登录问题
    pause
    exit /b %errorlevel%
)
git tag -f v1.0.0
git push origin v1.0.0 --force
echo.
echo ========================
echo 推送成功！构建已触发
echo.
pause
