@echo off
chcp 65001 >nul
cd /d "%~dp0"
title FCN 准备询价（生成三榜单 + 执行价/敲入价）

echo ══════════════════════════════════════════════
echo   FCN 准备询价
echo   流程：重建三榜单 → 取IV → 反解执行价/敲入价 → 填入Excel
echo ══════════════════════════════════════════════
echo.

echo [检查] 富途 OpenD 是否已启动...
netstat -an | findstr ":11111" | findstr "LISTENING" >nul
if errorlevel 1 (
  echo.
  echo ❌ 没检测到富途 OpenD（端口 11111 未监听）。
  echo    请先启动并登录 Futu OpenD，再重新双击本文件。
  echo.
  pause
  exit /b 1
)
echo    OpenD 已就绪
echo.

echo [1/4] 重建三个榜单 Excel（v2 格式）...
python rebuild_lists.py
if errorlevel 1 goto :err

echo.
echo [2/4] 获取 90%% 档 6M put IV（约 10-15 分钟，期间请勿关闭窗口）...
python fetch_iv.py
if errorlevel 1 goto :err

echo.
echo [3/4] 反解建议执行价 / 敲入价...
python strike_advisor_pipeline.py
if errorlevel 1 goto :err

echo.
echo [4/4] 把执行价 / 敲入价填进三个榜单 Excel...
python fill_terms.py
if errorlevel 1 goto :err

echo.
echo ══════════════════════════════════════════════
echo ✅ 完成！三个榜单 Excel 已生成，执行价 / 敲入价已填好。
echo.
echo   接下来（你来做）：
echo   1) 拿三个榜单去交易台询价
echo   2) 把票息填进每个 Excel 的「票息」列
echo      例：18%% 就填 0.18（或直接输入 18%%）；没报价的填 NA
echo   3) 双击「更新网站.bat」推送上线
echo.
echo   若热度榜有标的显示「跳过 / 无执行价」（IV 低于 30），
echo   请找我换一个替补标的再继续。
echo ══════════════════════════════════════════════
pause
exit /b 0

:err
echo.
echo ❌ 出错了：请把整个窗口截图发给我或 Dave。
pause
exit /b 1
