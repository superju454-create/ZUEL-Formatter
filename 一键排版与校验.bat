@echo off
:: ZUEL-Formatter 一键运行脚本 (专为 Windows 用户设计)
:: 双击此文件即可直接运行，无需了解任何命令行或 AI 知识
chcp 65001 >nul
title 中南财经政法大学学术排版与校验工具 -- ZUEL-Formatter

echo ======================================================================
echo          中南财经政法大学 (ZUEL) 学术论文一键排版与校验工具
echo ======================================================================
echo.

:: 1. 检测 Python 环境
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未在您的电脑上检测到 Python 环境！
    echo 请先前往 Python 官网 (https://www.python.org/) 下载并安装 Python 3.8 或以上版本。
    echo 安装时请务必勾选 "Add Python to PATH" (将 Python 添加到系统环境变量)。
    echo.
    pause
    exit /b
)

:: 2. 自动检查并安装所需依赖库
echo 正在检查并准备运行环境，请稍候...
python -c "import markdown, docx, pypdf, reportlab" >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在为您自动安装必要的排版支持库 (只需在首次运行时安装)...
    pip install markdown python-docx pypdf reportlab -i https://pypi.tuna.tsinghua.edu.cn/simple
    if %errorlevel% neq 0 (
        echo [警告] 依赖库自动安装失败，正在尝试默认源安装...
        pip install markdown python-docx pypdf reportlab
    )
)
cls

:menu
echo ======================================================================
echo          中南财经政法大学 (ZUEL) 学术论文一键排版与校验工具
echo ======================================================================
echo.
echo  【请选择您要进行的操作】：
echo     1. 论文一键排版 (将 Markdown 格式的论文转换为 Word + PDF)
echo     2. Word 格式合规校验 (检查写好的 .docx 论文是否符合学校排版规范)
echo     3. 退出程序
echo.
echo ======================================================================
set /p opt="请输入数字并按回车 (1/2/3): "

if "%opt%"=="1" goto convert
if "%opt%"=="2" goto validate
if "%opt%"=="3" exit
goto menu

:convert
echo.
echo ----------------------------------------------------------------------
echo  操作 1：一键排版 (Markdown 转 Word & PDF)
echo ----------------------------------------------------------------------
echo.
echo  [步骤] 请直接将您的 Markdown 格式论文文件 (.md)
echo         拖拽到本窗口内，然后按下键盘上的【回车键】。
echo.
set /p mdfile="请拖入文件: "
:: 去除拖拽可能产生的引号
set mdfile=%mdfile:"=%

if not exist "%mdfile%" (
    echo.
    echo [错误] 找不到该文件，请检查路径是否正确！
    pause
    cls
    goto menu
)

echo.
set /p header_text="请输入您的论文页眉文字 (若不需要请直接按回车): "

echo.
echo 正在为您进行高保真学术排版，请稍候...
python "%~dp0scripts\zuel_formatter.py" convert -i "%mdfile%" --header "%header_text%"
if %errorlevel% eq 0 (
    echo.
    echo [成功] 论文已顺利排版完成！
    echo 生成的 Word (.docx) 和 PDF (.pdf) 文件与您的源文件在同一个文件夹下。
) else (
    echo.
    echo [错误] 排版失败，请确保已安装 Pandoc 且 Edge 浏览器处于默认位置！
)
echo.
pause
cls
goto menu

:validate
echo.
echo ----------------------------------------------------------------------
echo  操作 2：Word 格式合规性校验 (检查 .docx 论文)
echo ----------------------------------------------------------------------
echo.
echo  [步骤] 请直接将您的 Word 论文文件 (.docx)
echo         拖拽到本窗口内，然后按下键盘上的【回车键】。
echo.
set /p docxfile="请拖入文件: "
set docxfile=%docxfile:"=%

if not exist "%docxfile%" (
    echo.
    echo [错误] 找不到该文件，请检查路径是否正确！
    pause
    cls
    goto menu
)

echo.
echo 正在对您的 Word 文档进行合规性格式扫描，请稍候...
python "%~dp0scripts\zuel_formatter.py" validate-docx -i "%docxfile%"
if %errorlevel% eq 0 (
    echo.
    echo [分析完成] 校验报告已输出！
    echo 详细的网页版报告 (Markdown 格式) 已经生成在与您 Word 同一目录下，
    echo 格式为: 文件名_validation_report.md
) else (
    echo.
    echo [错误] 校验失败，请确保文件未被其它程序占用！
)
echo.
pause
cls
goto menu
