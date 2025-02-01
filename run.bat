@echo off


REM 更新 pip
pip install --upgrade pip

REM 安装必要的依赖
pip install -r requirements.txt

REM 确保 stocks.xlsx 文件存在
if not exist "stocks.xlsx" (
    echo stocks.xlsx 不存在，正在创建...
    python -c "from openpyxl import Workbook; wb = Workbook(); wb.save('stocks.xlsx')"
)

REM 运行 app.py
python app.py

REM 暂停以查看输出
pause