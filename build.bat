@echo off
chcp 65001 >nul
echo ============================================
echo   T.QM.013 检验指导书生成器 - 打包工具
echo ============================================
echo.

:: 检查 PyInstaller
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo [1/3] 安装 PyInstaller...
    pip install pyinstaller
)

echo [2/3] 开始打包...
pyinstaller --noconfirm --onefile --windowed ^
    --name "T.QM.013检验指导书生成器" ^
    --add-data "template\T.QM.013_Template.xlsm;template" ^
    --hidden-import flask ^
    --hidden-import openpyxl ^
    --hidden-import werkzeug ^
    --hidden-import engineio ^
    --hidden-import socketio ^
    app.py

echo.
echo [3/3] 打包完成！
echo.
echo 输出文件位置: dist\T.QM.013检验指导书生成器.exe
echo.
echo 注意：将 template\T.QM.013_Template.xlsm 复制到
echo       dist\ 目录下与 .exe 同级，或打包进安装程序。
echo.
pause
