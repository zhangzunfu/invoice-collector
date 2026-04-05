# 如何在Windows 11上将发票归集软件编译成绿色exe包

## 准备工作

1. 确保已在Windows 11系统上安装Python 3.8+
2. 下载之前提供的发票归集软件包并解压

## 创建虚拟环境

```bash
# 进入解压后的发票归集软件目录
cd invoice_collector

# 创建虚拟环境
python -m venv invoice_env

# 激活虚拟环境
invoice_env\Scripts\activate  # Windows
# 或
source invoice_env/bin/activate  # Linux/Mac
```

## 安装依赖

```bash
# 安装项目依赖
pip install PyQt6 PyQt6-WebEngine pandas openpyxl PyMuPDF

# 安装PyInstaller
pip install pyinstaller
```

## 编译exe文件

```bash
# 使用PyInstaller编译成单个exe文件
pyinstaller --onefile --windowed --icon=icon.ico --name="发票归集软件" app.py

# 或者创建完整目录结构的绿色包
pyinstaller --onedir --windowed --icon=icon.ico --name="发票归集软件" app.py
```

## 编译参数说明

- `--onefile`: 将所有内容打包成单个exe文件
- `--onedir`: 创建包含exe和依赖文件的目录
- `--windowed`: 不显示控制台窗口
- `--icon`: 指定程序图标
- `--name`: 指定生成的exe文件名

## 生成的可执行文件

编译完成后，exe文件将位于dist目录下，这是一个独立的可执行文件，无需安装即可运行。

## 分发绿色包

如果使用`--onedir`选项，整个dist目录即为绿色包，可直接分发使用。