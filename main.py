#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
发票归集软件主程序
支持扫描指定文件夹中的PDF发票，并生成Excel清单
"""

import sys
import os
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem, QHeaderView, QTextEdit, QTabWidget, QProgressBar, QMessageBox
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import pandas as pd
from datetime import datetime


class InvoiceProcessor:
    """发票处理器类"""
    
    def __init__(self):
        self.invoice_data = []
        
    def extract_invoice_info(self, pdf_path):
        """
        从PDF发票中提取信息
        返回字典包含发票信息
        """
        # 这里是模拟实现，实际需要使用PDF解析库如PyMuPDF、pdfplumber等
        # 并配合OCR技术识别发票信息
        return {
            '开票日期': '2023-01-01',
            '发票号码': '12345678901234567890',
            '购方名称': '购买方公司名称',
            '销方名称': '销售方公司名称',
            '项目名称': '商品或服务名称',
            '金额': '100.00',
            '税额': '13.00',
            '价税合计': '113.00'
        }
    
    def scan_folder(self, folder_path):
        """扫描文件夹中的PDF文件"""
        invoices = []
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.pdf'):
                    pdf_path = os.path.join(root, file)
                    invoice_info = self.extract_invoice_info(pdf_path)
                    invoices.append(invoice_info)
        return invoices


class InvoiceScannerThread(QThread):
    """发票扫描线程"""
    progress_update = pyqtSignal(int)
    status_update = pyqtSignal(str)
    scan_complete = pyqtSignal(list)
    
    def __init__(self, folder_path):
        super().__init__()
        self.folder_path = folder_path
        
    def run(self):
        processor = InvoiceProcessor()
        invoices = processor.scan_folder(self.folder_path)
        self.scan_complete.emit(invoices)


class DashboardWidget(QWidget):
    """仪表盘组件"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # 添加仪表盘内容
        title_label = QLabel("发票统计仪表盘")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin: 10px;")
        
        # 统计信息
        stats_layout = QHBoxLayout()
        
        total_label = QLabel("总发票数: 0")
        total_label.setStyleSheet("background-color: #e3f2fd; padding: 10px; border-radius: 5px; margin: 5px;")
        total_label.setMinimumWidth(120)
        
        monthly_label = QLabel("本月发票: 0")
        monthly_label.setStyleSheet("background-color: #e8f5e8; padding: 10px; border-radius: 5px; margin: 5px;")
        monthly_label.setMinimumWidth(120)
        
        amount_label = QLabel("总金额: ¥0.00")
        amount_label.setStyleSheet("background-color: #fff3e0; padding: 10px; border-radius: 5px; margin: 5px;")
        amount_label.setMinimumWidth(120)
        
        stats_layout.addWidget(total_label)
        stats_layout.addWidget(monthly_label)
        stats_layout.addWidget(amount_label)
        stats_layout.addStretch()
        
        layout.addWidget(title_label)
        layout.addLayout(stats_layout)
        
        # 添加图表占位符
        chart_placeholder = QLabel("图表显示区域")
        chart_placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        chart_placeholder.setStyleSheet("background-color: white; border: 1px solid gray; padding: 20px; margin: 10px;")
        layout.addWidget(chart_placeholder)
        
        self.setLayout(layout)


class MainWindow(QMainWindow):
    """主窗口类"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("发票归集软件")
        self.setGeometry(100, 100, 1200, 800)
        
        self.processor = InvoiceProcessor()
        self.current_invoices = []
        
        self.init_ui()
        
    def init_ui(self):
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout()
        
        # 顶部控制区域
        control_layout = QHBoxLayout()
        
        self.select_folder_btn = QPushButton("选择文件夹")
        self.select_folder_btn.clicked.connect(self.select_folder)
        
        self.scan_btn = QPushButton("开始扫描")
        self.scan_btn.clicked.connect(self.start_scan)
        self.scan_btn.setEnabled(False)
        
        self.export_btn = QPushButton("导出Excel")
        self.export_btn.clicked.connect(self.export_to_excel)
        self.export_btn.setEnabled(False)
        
        self.status_label = QLabel("就绪")
        
        control_layout.addWidget(self.select_folder_btn)
        control_layout.addWidget(self.scan_btn)
        control_layout.addWidget(self.export_btn)
        control_layout.addStretch()
        control_layout.addWidget(self.status_label)
        
        # 创建选项卡
        self.tabs = QTabWidget()
        
        # 发票列表选项卡
        self.invoice_table = QTableWidget()
        self.setup_table()
        self.tabs.addTab(self.invoice_table, "发票列表")
        
        # 仪表盘选项卡
        self.dashboard = DashboardWidget()
        self.tabs.addTab(self.dashboard, "仪表盘")
        
        # 日志区域
        self.log_area = QTextEdit()
        self.log_area.setMaximumHeight(150)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        
        # 添加到主布局
        main_layout.addLayout(control_layout)
        main_layout.addWidget(self.tabs)
        main_layout.addWidget(self.progress_bar)
        main_layout.addWidget(self.log_area)
        
        central_widget.setLayout(main_layout)
        
    def setup_table(self):
        """设置表格"""
        headers = ['开票日期', '发票号码', '购方名称', '销方名称', '项目名称', '金额', '税额', '价税合计']
        self.invoice_table.setColumnCount(len(headers))
        self.invoice_table.setHorizontalHeaderLabels(headers)
        
        # 设置列宽
        header = self.invoice_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        
    def select_folder(self):
        """选择文件夹"""
        folder_path = QFileDialog.getExistingDirectory(self, "选择发票文件夹")
        if folder_path:
            self.selected_folder = folder_path
            self.status_label.setText(f"已选择文件夹: {folder_path}")
            self.scan_btn.setEnabled(True)
            self.log_area.append(f"选择文件夹: {folder_path}")
            
    def start_scan(self):
        """开始扫描"""
        if hasattr(self, 'selected_folder'):
            self.scan_thread = InvoiceScannerThread(self.selected_folder)
            self.scan_thread.progress_update.connect(self.update_progress)
            self.scan_thread.status_update.connect(self.update_status)
            self.scan_thread.scan_complete.connect(self.on_scan_complete)
            
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            self.scan_btn.setEnabled(False)
            self.status_label.setText("正在扫描...")
            
            self.scan_thread.start()
            
    def update_progress(self, value):
        """更新进度"""
        self.progress_bar.setValue(value)
        
    def update_status(self, status):
        """更新状态"""
        self.status_label.setText(status)
        
    def on_scan_complete(self, invoices):
        """扫描完成回调"""
        self.current_invoices = invoices
        self.display_invoices(invoices)
        self.progress_bar.setVisible(False)
        self.scan_btn.setEnabled(True)
        self.export_btn.setEnabled(True)
        self.status_label.setText(f"扫描完成，找到 {len(invoices)} 张发票")
        self.log_area.append(f"扫描完成，共找到 {len(invoices)} 张发票")
        
    def display_invoices(self, invoices):
        """在表格中显示发票信息"""
        self.invoice_table.setRowCount(len(invoices))
        
        for row, invoice in enumerate(invoices):
            self.invoice_table.setItem(row, 0, QTableWidgetItem(invoice.get('开票日期', '')))
            self.invoice_table.setItem(row, 1, QTableWidgetItem(invoice.get('发票号码', '')))
            self.invoice_table.setItem(row, 2, QTableWidgetItem(invoice.get('购方名称', '')))
            self.invoice_table.setItem(row, 3, QTableWidgetItem(invoice.get('销方名称', '')))
            self.invoice_table.setItem(row, 4, QTableWidgetItem(invoice.get('项目名称', '')))
            self.invoice_table.setItem(row, 5, QTableWidgetItem(str(invoice.get('金额', ''))))
            self.invoice_table.setItem(row, 6, QTableWidgetItem(str(invoice.get('税额', ''))))
            self.invoice_table.setItem(row, 7, QTableWidgetItem(str(invoice.get('价税合计', ''))))
    
    def export_to_excel(self):
        """导出到Excel"""
        if self.current_invoices:
            file_path, _ = QFileDialog.getSaveFileName(
                self, 
                "保存Excel文件", 
                f"发票清单_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                "Excel 文件 (*.xlsx)"
            )
            
            if file_path:
                try:
                    df = pd.DataFrame(self.current_invoices)
                    # 按开票日期升序排列
                    df = df.sort_values(by=['开票日期'])
                    df.to_excel(file_path, index=False)
                    
                    self.log_area.append(f"成功导出到: {file_path}")
                    QMessageBox.information(self, "导出成功", f"发票清单已成功导出到:\n{file_path}")
                except Exception as e:
                    self.log_area.append(f"导出失败: {str(e)}")
                    QMessageBox.critical(self, "导出失败", f"导出失败:\n{str(e)}")


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()