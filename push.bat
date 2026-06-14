@echo off
cd /d "C:\Users\liche\Documents\Codex\temp-pdf-repo"
echo 正在推送到 GitHub...
git push origin main
if %errorlevel% neq 0 (
    echo 推送失败，请检查网络
    pause
    exit /b %errorlevel%
)
git tag v1.1.1
git push origin v1.1.1
echo.
echo ========================
echo 推送成功！v1.1.1 已发布
echo.
pause
