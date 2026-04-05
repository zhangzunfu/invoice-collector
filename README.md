# 发票归集软件

一款用于扫描、识别和整理PDF格式发票的桌面应用程序，支持普通发票、增值税发票和电商发票的识别，自动生成Excel格式的发票清单，并提供智能仪表盘统计功能。

## 功能特性

- **GUI界面**: 直观易用的桌面应用程序
- **发票扫描**: 自动扫描指定文件夹中的PDF发票
- **发票识别**: 支持普通发票、增值税发票、电商发票的识别
- **数据提取**: 提取开票日期、发票号码、购方名称、销方名称、项目名称、金额、税额、价税合计等信息
- **Excel导出**: 生成结构化Excel清单，按开票日期升序排列
- **智能仪表盘**: 实时统计发票数据

## 系统要求

- Windows 11
- Python 3.8 或更高版本

## 安装步骤

1. 克隆或下载项目代码到本地
2. 在Windows 11系统上打开命令提示符或PowerShell
3. 导航到项目目录
4. 安装依赖包：

```bash
pip install PyQt6 PyQt6-WebEngine pandas openpyxl PyMuPDF
```

## 使用方法

### 方法一：直接运行GUI应用

```bash
python app.py
```

### 方法二：使用启动脚本

```bash
./start_app.sh  # Linux/Mac
# 或者在Windows上双击 start_app.bat
```

## 字段说明

生成的Excel清单包含以下字段（从左到右）：
1. 开票日期
2. 发票号码（最大20位数字）
3. 购方名称
4. 销方名称
5. 项目名称
6. 金额
7. 税额
8. 价税合计

## 注意事项

- 本软件主要针对Windows 11系统设计，GUI界面需要在Windows环境下运行
- PDF解析功能依赖于PyMuPDF库
- Excel导出功能依赖于openpyxl库
- 发票识别算法针对标准格式发票优化

## 源码结构

- `app.py`: 主应用程序入口
- `ui/mainwindow.py`: GUI界面实现
- `data_processing/invoice_parser.py`: PDF发票解析逻辑
- `start_app.sh`: 启动脚本

## 命令行演示

如果无法运行GUI版本，可以使用命令行演示版查看核心功能：

```bash
python demo_cli.py
```