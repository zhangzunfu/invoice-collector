# 发票归集软件 - 项目上下文文档

## 项目概述

发票归集软件是一款基于 Python 的桌面应用程序，用于扫描、识别和整理 PDF 格式的发票。该软件支持普通发票、增值税发票和电商发票的识别，能够自动提取发票关键信息，生成 Excel 格式的发票清单，并提供智能仪表盘统计功能。

### 主要特性

- **GUI 界面**：基于 PyQt6 的直观易用桌面应用程序
- **发票扫描**：自动扫描指定文件夹中的 PDF 发票文件
- **智能识别**：支持多种发票类型的识别（普通发票、增值税发票、电子发票）
- **数据提取**：自动提取开票日期、发票号码、购方名称、销方名称、项目名称、金额、税额、价税合计等信息
- **Excel 导出**：生成结构化 Excel 清单，按开票日期升序排列
- **智能仪表盘**：实时统计发票数据，包括总发票数、本月发票数、总金额等统计信息

## 技术栈

- **语言**：Python 3.8+
- **GUI 框架**：PyQt6 + PyQt6-WebEngine
- **PDF 解析**：PyMuPDF (fitz)
- **数据处理**：pandas
- **Excel 导出**：openpyxl

## 项目结构

```
invoice_collector/
├── app.py                          # 主应用程序入口（推荐使用）
├── main.py                         # 备用主程序（较旧版本）
├── demo_cli.py                     # 命令行演示版
├── test_core_functionality.py      # 核心功能测试
├── test_parser.py                  # 发票解析器测试
├── start_app.bat                   # Windows 启动脚本
├── start_app.sh                    # Linux/Mac 启动脚本
├── BUILD_EXE_GUIDE.md              # EXE 编译指南
├── data_processing/                # 数据处理模块
│   ├── invoice_parser.py           # PDF 发票解析逻辑
│   └── __pycache__/                # Python 缓存文件
├── ui/                             # 用户界面模块
│   ├── mainwindow.py               # 主窗口界面实现
│   └── __pycache__/                # Python 缓存文件
├── utils/                          # 工具函数模块（当前为空）
└── sample_invoices/                # 示例发票文件夹
```

### 核心文件说明

#### `app.py`
- 主应用程序入口文件
- 创建 PyQt6 应用实例并启动主窗口
- 使用 Fusion 样式提供跨平台外观

#### `ui/mainwindow.py`
- 完整的 GUI 界面实现
- 包含三个主要组件：
  - **MainWindow**：主窗口，控制整体布局和功能协调
  - **DashboardWidget**：仪表盘组件，显示发票统计信息
  - **SettingsWidget**：设置界面，配置监控和输出选项
- 使用多线程（QThread）处理发票扫描，避免界面卡顿
- 支持发票列表展示、Excel 导出、日志记录等功能

#### `data_processing/invoice_parser.py`
- 核心发票解析逻辑
- **InvoiceParser 类**：负责从 PDF 文本中提取发票信息
  - 支持多种发票类型识别
  - 使用正则表达式提取关键字段
  - 智能处理日期、金额、发票号码等格式
- **InvoiceProcessor 类**：批量处理文件夹中的发票文件

#### `demo_cli.py`
- 命令行演示版本
- 展示核心功能，无需 GUI 环境
- 适用于测试和功能演示

#### `test_core_functionality.py`
- 核心功能测试脚本
- 测试 Excel 导出功能
- 验证数据处理逻辑

## 系统要求

- **操作系统**：Windows 11（GUI 功能需要 Windows 环境）
- **Python 版本**：3.8 或更高版本

## 安装和依赖

### 安装依赖

```bash
pip install PyQt6 PyQt6-WebEngine pandas openpyxl PyMuPDF
```

### 依赖说明

- `PyQt6`：桌面 GUI 框架
- `PyQt6-WebEngine`：Web 引擎支持
- `pandas`：数据处理和分析
- `openpyxl`：Excel 文件读写
- `PyMuPDF`：PDF 文件解析

## 使用方法

### 启动 GUI 应用

#### 方法一：直接运行
```bash
python app.py
```

#### 方法二：使用启动脚本
- Windows：双击 `start_app.bat`
- Linux/Mac：`./start_app.sh`

### 运行命令行演示
```bash
python demo_cli.py
```

### 运行测试
```bash
# 核心功能测试
python test_core_functionality.py

# 发票解析器测试
python test_parser.py
```

## 开发指南

### 添加新功能

1. **发票解析增强**：在 `data_processing/invoice_parser.py` 中修改 `InvoiceParser` 类
2. **界面扩展**：在 `ui/mainwindow.py` 中添加新的 GUI 组件
3. **数据处理**：在 `data_processing/` 目录下添加新的处理模块

### 代码规范

- 使用 UTF-8 编码
- 遵循 PEP 8 代码风格
- 所有文件包含中文注释和文档字符串
- GUI 使用 PyQt6 最佳实践

### 主要类和方法

#### InvoiceParser 类
- `parse_pdf_invoice(pdf_path)`：解析单个 PDF 发票
- `_extract_invoice_info(text)`：从文本中提取发票信息
- `_identify_invoice_type(text)`：识别发票类型

#### MainWindow 类
- `select_folder()`：选择发票文件夹
- `start_scan()`：开始扫描发票
- `export_to_excel()`：导出发票数据到 Excel
- `display_invoices(invoices)`：在表格中显示发票信息

### 数据字段说明

提取的发票数据包含以下字段：

| 字段 | 说明 | 示例 |
|------|------|------|
| 开票日期 | 发票开具日期 | 2024-01-15 |
| 发票号码 | 发票编号（8-20位） | 01234567890123456789 |
| 购方名称 | 购买方公司名称 | 北京科技有限公司 |
| 销方名称 | 销售方公司名称 | 上海贸易有限公司 |
| 项目名称 | 商品或服务名称 | 办公用品 |
| 金额 | 不含税金额 | 1000.00 |
| 税额 | 税额 | 130.00 |
| 价税合计 | 含税总金额 | 1130.00 |
| 发票类型 | 发票类型分类 | 增值税发票/电子发票/普通发票 |

## 编译为 EXE

项目提供了详细的 EXE 编译指南，请参考 `BUILD_EXE_GUIDE.md`。

### 快速编译步骤

```bash
# 1. 创建虚拟环境
python -m venv invoice_env
invoice_env\Scripts\activate

# 2. 安装依赖
pip install PyQt6 PyQt6-WebEngine pandas openpyxl PyMuPDF pyinstaller

# 3. 编译为单个 EXE
pyinstaller --onefile --windowed --name="发票归集软件" app.py

# 或编译为绿色包（目录结构）
pyinstaller --onedir --windowed --name="发票归集软件" app.py
```

## 注意事项

1. **PDF 解析依赖**：PDF 解析功能依赖于 PyMuPDF 库的稳定性
2. **发票格式**：识别算法针对标准格式发票优化，非标准格式可能需要调整
3. **系统兼容性**：GUI 界面专为 Windows 11 设计，其他系统可能需要适配
4. **编码问题**：确保所有文件使用 UTF-8 编码
5. **线程安全**：GUI 使用多线程处理耗时操作，注意线程间通信

## 扩展建议

1. **OCR 集成**：添加 OCR 功能支持扫描的 PDF 图片发票
2. **数据库支持**：将发票数据存储到数据库便于管理和查询
3. **高级图表**：使用 matplotlib 或其他图表库增强仪表盘功能
4. **多格式支持**：扩展支持图片格式发票
5. **自动监控**：实现文件夹自动监控和实时更新

## 故障排查

### 常见问题

1. **PDF 解析失败**
   - 检查 PDF 文件是否可读
   - 确认 PyMuPDF 安装正确
   - 查看日志区域的错误信息

2. **GUI 无法启动**
   - 确认 PyQt6 安装正确
   - 检查 Python 版本是否符合要求
   - 尝试使用命令行演示版测试核心功能

3. **Excel 导出失败**
   - 确认 openpyxl 安装正确
   - 检查文件保存路径权限
   - 验证数据格式是否正确

## 贡献指南

如需贡献代码或报告问题，请遵循以下步骤：

1. 阅读并理解现有代码结构
2. 编写清晰的代码注释
3. 确保代码通过测试
4. 更新相关文档

## 版本信息

- 当前版本：1.0.0
- Python 要求：3.8+
- 推荐系统：Windows 11

## 许可证

请查看项目根目录下的 LICENSE 文件（如果存在）。

## 联系方式

如有问题或建议，请通过项目仓库提交 Issue。