#!/bin/bash
# 发票归集软件启动脚本

echo "正在启动发票归集软件..."
echo "请确保您在图形界面环境中运行此程序"

# 检查必要的依赖
echo "检查依赖..."

if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到python3，请先安装Python 3"
    exit 1
fi

# 尝试导入必要的库
python3 << EOF
try:
    import sys
    import os
    sys.path.append('/home/admin/workspace/invoice_collector')
    
    # 检查必要的模块
    import PyQt6
    import fitz
    import pandas
    import openpyxl
    
    print("✓ 所有依赖项都已安装")
    
    # 导入我们的模块
    from ui.mainwindow import MainWindow
    from data_processing.invoice_parser import InvoiceProcessor
    print("✓ 所有模块都可以正常导入")
    
except ImportError as e:
    print(f"✗ 缺少依赖: {e}")
    exit(1)
except Exception as e:
    print(f"✗ 模块导入错误: {e}")
    exit(1)
EOF

if [ $? -ne 0 ]; then
    echo "依赖检查失败，请检查安装。"
    exit 1
fi

echo "启动发票归集软件..."
cd /home/admin/workspace/invoice_collector && python3 app.py