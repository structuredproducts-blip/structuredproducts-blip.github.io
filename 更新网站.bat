@echo off
chcp 65001 >nul
cd /d "%~dp0"
title FCN Watchlist 一键更新

echo ══════════════════════════════════════════════
echo   FCN Watchlist 一键更新
echo   流程：拉取最新代码 → 生成分析 → 推送上线
echo ══════════════════════════════════════════════
echo.

echo [1/4] 拉取最新代码...
git pull --rebase --autostash
if errorlevel 1 goto :err

echo.
echo [2/4] 生成标的分析（约 25 分钟，期间请勿关闭本窗口）...
echo       提示：如果本机没有登录富途 OpenD，分析会缺少分析师评级数据，但仍可正常生成。
python generate_watchlist.py
if errorlevel 1 goto :err

echo.
echo [3/4] 提交更改...
git add -A
git commit -m "weekly update"

echo.
echo [4/4] 推送上线...
git push
if errorlevel 1 goto :err

echo.
echo ✅ 完成！网站将在几分钟内自动更新：https://structuredproducts-blip.github.io
pause
exit /b 0

:err
echo.
echo ❌ 出错了：请截图整个窗口发给 Dave。
pause
exit /b 1
