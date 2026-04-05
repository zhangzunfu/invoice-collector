#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
发票解析模块 - V3.4 精准处理版
负责从PDF文件中提取发票信息
使用基于文本块结构的解析方法

V3.4 精准处理内容（2026-04-04）：
1. 【京东发票特殊处理】添加精准的京东发票格式检测和处理
2. 【扫描文本块】在最后阶段扫描所有文本块，查找包含两个有效公司名称的文本块
3. 【智能分配】如果找到两个有效的公司名称，第一个作为销方名称，第二个作为购方名称
4. 【保留修复】保留V3.2版本的严格验证逻辑，确保不会出现"合"、"计"等单字错误
5. 【安全优先】只在检测到特殊格式时进行特殊处理，避免引入新的bug

V3.2 紧急修复内容（2026-04-04）：
1. 【严重Bug修复】完全禁用公司名称分割逻辑 - 解决"合"、"计"被误识别为公司名称的问题
2. 【严格验证】增加_validate_company_name函数，实现严格的公司名称验证
   - 长度必须在4-50个字符之间
   - 绝对不能是单个汉字（如"合"、"计"）
   - 绝对不能是"合计"、"大写"、"小写"等常见错误词
   - 中文字符占比必须超过60%
   - 必须包含公司名称特征关键词（"有限公司"、"公司"等）

V3.1 修复内容：
1. 修复关键词匹配失败 - 移除换行符后再匹配标签位置
2. 优化位置判断 - 调整为60像素（V3.0的80像素过于宽松）
3. 增加公司名称分割逻辑 - 处理购方和销方名称在一个文本块中的情况（已禁用）
4. 改进项目名称提取逻辑 - 对多行文本块，提取第一行作为项目名称
5. 增强公司名称验证 - 添加公司名称特征关键词检查，避免发票类型信息被误识别
6. 完善排除关键词列表 - 添加更多需要排除的关键词
"""

import fitz  # PyMuPDF
import re
import json
from datetime import datetime
from typing import Dict, List, Optional, Tuple
import logging

# 配置日志
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class InvoiceParser:
    """发票解析器 - V3.0 修复版"""
    
    def __init__(self, debug=False):
        self.debug = debug
        # 发票类型关键词
        self.special_invoice_keywords = [
            '增值', '专票', '专用发票', '增值税专用发票'
        ]
        self.electronic_invoice_keywords = [
            '电子', '普通发票', '电子普通发票', '电子发票'
        ]
        self.common_invoice_keywords = [
            '普通发票', '发票'
        ]
    
    def _validate_company_name(self, name: str) -> bool:
        """
        严格验证是否为公司名称
        
        Args:
            name: 待验证的文本
            
        Returns:
            bool: 是否为公司名称
        """
        if not name or not isinstance(name, str):
            return False
        
        # 基本长度检查：公司名称必须在4-50个字符之间
        if len(name) < 4 or len(name) > 50:
            return False
        
        # 绝对不能是单个汉字
        if len(name) == 1:
            return False
        
        # 绝对不能是"合"、"计"等常见错误
        if name in ['合', '计', '合计', '大写', '小写', '整', '元', '圆']:
            return False
        
        # 计算中文字符比例
        chinese_chars = len(re.findall(r'[\u4e00-\u9fa5]', name))
        total_chars = len(name)
        chinese_ratio = chinese_chars / total_chars if total_chars > 0 else 0
        
        # 中文字符占比必须超过60%
        if chinese_ratio < 0.6:
            return False
        
        # 必须包含公司名称特征关键词
        company_keywords = ['有限公司', '股份有限公司', '有限责任公司', '公司', '集团', '企业', '厂', '店', '合作社']
        has_company_keyword = any(kw in name for kw in company_keywords)
        
        # 如果没有公司关键词，需要更严格的中文字符占比（超过80%）
        if not has_company_keyword and chinese_ratio < 0.8:
            return False
        
        return True
    
    def _simplify_item_name(self, item_name: str) -> str:
        """
        简化项目名称，只保留大类和物料名称
        移除规格型号、数量、金额等详细信息
        
        例如：
        输入: *食品添加剂*甘油13%千克441061.9557338.057.876106194690356000
        输出: 食品添加剂 甘油
        """
        if not item_name or item_name == '未识别':
            return item_name
        
        # 移除开头的*
        cleaned = item_name.lstrip('*')
        
        # 按*分割，提取前两部分（大类和物料名称）
        parts = cleaned.split('*')
        if len(parts) >= 2:
            # 取前两部分：大类和物料名称
            simplified = parts[0] + ' ' + parts[1]
        elif len(parts) == 1:
            simplified = parts[0]
        else:
            simplified = cleaned
        
        # 移除所有数字、百分号、货币符号、特殊符号（保留中文、空格、字母）
        # 使用正则表达式移除：数字、%、¥、（、）、~、<、>等
        simplified = re.sub(r'[0-9%¥（）\(\)\[\]\<\>\~\-\.\,]+', ' ', simplified)
        
        # 移除常见的规格单位（使用词边界匹配，避免删除包含这些词的中文字符）
        unit_keywords = ['千克', '克', '公斤', '吨', 'ML', 'MM', '升', '包', '箱', '瓶', '个', '条', '卷', '张', '台', '套', '只', '米', '桶', '袋', '盒', '罐', '片', '支', '双', '副', '组', '批', '次']
        for keyword in unit_keywords:
            # 使用词边界，避免删除包含这些词的中文
            simplified = re.sub(r'\b' + re.escape(keyword) + r'\b', '', simplified)
        
        # 移除连续的空格
        simplified = re.sub(r'\s+', ' ', simplified).strip()
        
        # 如果清理后为空或太短，返回原始清理后的名称
        if len(simplified) < 3:
            # 尝试从原始名称中提取第一个词组
            words = re.findall(r'[\u4e00-\u9fa5]+', item_name)
            if len(words) >= 2:
                simplified = ' '.join(words[:2])
            elif len(words) == 1:
                simplified = words[0]
            else:
                simplified = item_name
        
        return simplified
    
    def parse_pdf_invoice(self, pdf_path: str) -> Optional[Dict]:
        """
        解析PDF发票文件，提取关键信息
        使用基于文本块的方法 - V4.1 优化版
        """
        try:
            if self.debug:
                logger.info(f"开始解析文件: {pdf_path}")
            
            doc = fitz.open(pdf_path)
            
            # 处理每一页
            for page_num, page in enumerate(doc):
                if self.debug:
                    logger.info(f"正在处理第 {page_num + 1} 页")
                
                # 提取文本块
                blocks = page.get_text("blocks")
                
                # 优先使用V4.0优化版解析逻辑
                invoice_info = self._extract_invoice_info_from_blocks_v4(blocks, page)
                
                # 如果V4.0失败，回退到V3.0逻辑
                if not invoice_info or not self._has_valid_data(invoice_info):
                    if self.debug:
                        logger.info("V4.0解析失败，回退到V3.0逻辑")
                    invoice_info = self._extract_invoice_info_from_blocks(blocks, page)
                
                if invoice_info and self._has_valid_data(invoice_info):
                    if self.debug:
                        logger.info(f"成功提取发票信息: {invoice_info}")
                    doc.close()
                    return invoice_info
            
            doc.close()
            logger.warning("未找到有效的发票信息")
            return None
            
        except Exception as e:
            logger.error(f"解析PDF失败: {str(e)}", exc_info=True)
            return None
    
    def _has_valid_data(self, invoice_info: Dict) -> bool:
        """检查是否提取到有效数据"""
        valid_fields = 0
        for key in ['发票号码', '开票日期', '购方名称', '销方名称', '价税合计']:
            if invoice_info.get(key):
                valid_fields += 1
        return valid_fields >= 3
    
    def _extract_invoice_info_from_blocks_v4(self, blocks: List, page) -> Dict:
        """
        从文本块中提取发票信息 - V4.0 优化版
        使用test_blocks_parser.py中的高效解析逻辑
        基于位置和内容的精确匹配，简化判断逻辑
        """
        result = {
            '开票日期': '',
            '发票号码': '',
            '购方名称': '',
            '销方名称': '',
            '项目名称': '',
            '税率': '',
            '金额': '',
            '税额': '',
            '价税合计': '',
            '发票类型': '普通发票'
        }
        
        # 收集所有文本块
        all_blocks = []
        for block in blocks:
            # block结构: (x0, y0, x1, y1, text, block_type, block_no)
            if len(block) >= 6:
                x0, y0, x1, y1, text, block_type = block[:6]
                if text.strip():  # 只保留非空文本块
                    all_blocks.append({
                        'x0': x0,
                        'y0': y0,
                        'x1': x1,
                        'y1': y1,
                        'text': text.strip(),
                        'type': block_type
                    })
        
        # 识别发票类型
        for block in all_blocks:
            text = block['text']
            if '增值税专用发票' in text:
                result['发票类型'] = '增值税专用发票'
                break
            elif '电子发票' in text:
                result['发票类型'] = '电子发票'
                break
        
        # 提取发票号码（纯数字，通常在右上角，x坐标大于400，y坐标较小）
        for block in all_blocks:
            text = block['text']
            if (re.match(r'^\d{20}$', text) and 
                block['x0'] > 400 and block['y0'] < 60):
                result['发票号码'] = text
                break
        
        # 提取开票日期
        for block in all_blocks:
            text = block['text']
            date_match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', text)
            if date_match:
                year, month, day = date_match.groups()
                result['开票日期'] = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                break
        
        # 提取购方名称（包含"公司"，在左侧区域）
        for block in all_blocks:
            text = block['text']
            if ('公司' in text and 
                block['x0'] > 30 and block['x0'] < 300 and
                block['y0'] > 80 and block['y0'] < 150 and
                len(text) > 5):
                result['购方名称'] = text
                break
        
        # 提取销方名称（在右侧区域）
        for block in all_blocks:
            text = block['text']
            # 放宽条件：不仅包含"公司"，也接受其他组织类型
            company_keywords = ['公司', '经营部', '厂', '店', '集团', '企业', '合作社']
            if (any(kw in text for kw in company_keywords) and 
                block['x0'] > 300 and
                block['y0'] > 80 and block['y0'] < 150 and
                len(text) > 5):
                # 避免与购方名称重复
                if result.get('购方名称', '') != text:
                    result['销方名称'] = text
                    break
        
        # V4.1: 京东发票特殊处理
        # 如果购方名称包含两个公司名称（用换行符分隔），第一个是购方，第二个是销方
        if result['购方名称'] and '\n' in result['购方名称']:
            companies = result['购方名称'].split('\n')
            if len(companies) >= 2:
                # 智能判断：包含"绿微康"的是购方名称
                potential_buyer = None
                potential_seller = None
                
                for company in companies:
                    company = company.strip()
                    if not company:
                        continue
                    if '绿微康' in company:
                        potential_buyer = company
                    elif '公司' in company and len(company) > 5:
                        potential_seller = company
                
                # 如果找到了购方和销方，就设置
                if potential_buyer and potential_seller:
                    result['购方名称'] = potential_buyer
                    result['销方名称'] = potential_seller
                    if self.debug:
                        logger.debug(f"京东发票特殊处理：购方={potential_buyer}, 销方={potential_seller}")
        
        # 提取项目名称（包含"*"的商品名称）
        for block in all_blocks:
            text = block['text']
            if '*' in text and block['y0'] > 150 and block['y0'] < 250:
                # 提取*后的第一行内容（避免包含规格、单位等信息）
                parts = text.split('*')
                if len(parts) >= 2:
                    # 格式通常是"*类别*名称"，取第二个*后的内容
                    # 如果只有1个*，则取*后的内容
                    if len(parts) >= 3:
                        item_name = parts[2].split('\n')[0].strip()
                    else:
                        item_name = parts[1].split('\n')[0].strip()
                    
                    # 过滤掉纯数字或规格型号（如"0.08*580*"、"1060MM"）
                    if item_name and not re.match(r'^[\d\.*]+$', item_name):
                        result['项目名称'] = item_name
                        break
        
        # 提取金额和税额（在"合计"区域，通常包含两个¥值）
        for block in all_blocks:
            text = block['text']
            # 查找包含两个¥金额的块（金额和税额）
            amounts = re.findall(r'¥([\d,]+\.?\d*)', text)
            if len(amounts) >= 2:
                # 第一个是金额，第二个是税额
                result['金额'] = amounts[0].replace(',', '')
                result['税额'] = amounts[1].replace(',', '')
                break
        
        # 提取价税合计（包含"价税合计"或在大写金额附近）
        for block in all_blocks:
            text = block['text']
            # 查找价税合计的金额
            total_match = re.search(r'¥([\d,]+\.?\d*)', text)
            if total_match and block['y0'] > 270:
                result['价税合计'] = total_match.group(1).replace(',', '')
                break
        
        # 提取税率（百分比格式）
        for block in all_blocks:
            text = block['text']
            # 查找包含百分号的税率值
            tax_rate_match = re.search(r'(\d+(?:\.\d+)?)%', text)
            if tax_rate_match:
                tax_rate = float(tax_rate_match.group(1))
                # 只提取合理的税率值（0-20之间）
                if 0 <= tax_rate <= 20:
                    # 排除纯金额中的数字（金额通常大于100或包含小数点且不是常见税率）
                    if tax_rate > 100 or (tax_rate > 10 and '.' in str(tax_rate) and tax_rate not in [13.0, 9.0, 6.0]):
                        continue
                    result['税率'] = f"{tax_rate:.2f}%"
                    break
        
        # 如果没有找到税率，尝试从金额和税额计算
        if not result['税率'] or result['税率'] == '0.00%':
            if result['金额'] and result['税额']:
                try:
                    amount = float(result['金额'])
                    tax = float(result['税额'])
                    if amount > 0:
                        calculated_rate = (tax / amount * 100)
                        # 检查是否接近常见税率
                        common_rates = [0, 1, 3, 5, 6, 9, 11, 13]
                        matched_rate = None
                        min_diff = float('inf')
                        
                        for rate in common_rates:
                            diff = abs(calculated_rate - rate)
                            if diff < min_diff and diff <= 0.3:
                                min_diff = diff
                                matched_rate = rate
                        
                        if matched_rate is not None:
                            result['税率'] = f"{matched_rate:.2f}%"
                            if self.debug:
                                logger.debug(f"从金额税额计算税率: {calculated_rate:.2f}% -> {result['税率']}")
                except ValueError:
                    pass
        
        # 数据验证和智能计算（使用现有的验证逻辑）
        # 如果有金额和税额但没有价税合计，计算价税合计
        if result['金额'] and result['税额'] and not result['价税合计']:
            try:
                amount = float(result['金额'])
                tax = float(result['税额'])
                result['价税合计'] = f"{amount + tax:.2f}"
            except ValueError:
                pass
        
        # 如果有价税合计和税额但没有金额，计算金额
        if result['价税合计'] and result['税额'] and not result['金额']:
            try:
                total = float(result['价税合计'])
                tax = float(result['税额'])
                result['金额'] = f"{total - tax:.2f}"
            except ValueError:
                pass
        
        # 如果有价税合计和金额但没有税额，计算税额
        if result['价税合计'] and result['金额'] and not result['税额']:
            try:
                total = float(result['价税合计'])
                amount = float(result['金额'])
                result['税额'] = f"{total - amount:.2f}"
            except ValueError:
                pass
        
        # 确保所有字段都有默认值
        if not result['购方名称']:
            result['购方名称'] = '未识别'
        if not result['销方名称']:
            result['销方名称'] = '未识别'
        if not result['项目名称']:
            result['项目名称'] = '未识别'
        if not result['税率']:
            result['税率'] = '0.00%'
        if not result['金额']:
            result['金额'] = '0.00'
        if not result['税额']:
            result['税额'] = '0.00'
        if not result['价税合计']:
            result['价税合计'] = '0.00'
        
        if self.debug:
            logger.info(f"V4.0提取结果: {result}")
        
        return result
    
    def _extract_invoice_info_from_blocks(self, blocks: List, page) -> Dict:
        """
        从文本块中提取发票信息
        基于位置关系和文本内容 - V3.0 修复版
        """
        result = {
            '开票日期': '',
            '发票号码': '',
            '购方名称': '',
            '销方名称': '',
            '项目名称': '',
            '税率': '',
            '金额': '',
            '税额': '',
            '价税合计': '',
            '发票类型': '普通发票'
        }
        
        # 过滤出文本块
        text_blocks = [b for b in blocks if b[6] == 0]  # 0 表示文本块
        
        if self.debug:
            logger.debug(f"找到 {len(text_blocks)} 个文本块")
        
        # 收集所有候选数据
        invoice_numbers = []  # 发票号码候选
        dates = []  # 日期候选
        buyer_candidates = []  # 购方候选
        seller_candidates = []  # 销方候选
        item_names = []  # 项目名称候选
        amount_values = []  # 金额候选（带位置信息）
        tax_values = []  # 税额候选（带位置信息）
        total_values = []  # 价税合计候选（带位置信息）
        tax_rate_values = []  # 税率候选（带位置信息）
        
        # 先识别发票类型和收集基本信息
        all_text = ' '.join([b[4].strip() for b in text_blocks])
        
        # 识别发票类型
        if '增值税专用发票' in all_text:
            result['发票类型'] = '增值税专用发票'
        elif '电子发票' in all_text or '电子普通发票' in all_text:
            result['发票类型'] = '电子发票'
        
        # 遍历文本块收集信息
        for block in text_blocks:
            x0, y0, x1, y1, text, block_no, _ = block
            text_clean = text.strip()
            
            if not text_clean:
                continue
            
            if self.debug:
                logger.debug(f"块 {block_no}: [{x0:.2f},{y0:.2f},{x1:.2f},{y1:.2f}] {text_clean[:50]}")
            
            # 提取发票号码（20位数字）
            if re.match(r'^[0-9]{20}$', text_clean):
                invoice_numbers.append((text_clean, y0))
            elif re.match(r'^[0-9]{19,21}$', text_clean):
                # 19-21位数字也可能是发票号码
                invoice_numbers.append((text_clean, y0))
            
            # 提取开票日期（多种格式）
            if re.match(r'^[0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日$', text_clean):
                dates.append((text_clean, y0))
            elif re.match(r'^[0-9]{4}-[0-9]{1,2}-[0-9]{1,2}$', text_clean):
                dates.append((text_clean, y0))
            
            # 提取金额数值（带¥符号或不带）
            # 支持单个金额或多个金额（如"¥126.73 ¥1.27"）
            amount_pattern = r'¥?([0-9,]+\.[0-9]{2})'
            matches = re.findall(amount_pattern, text_clean)
            for match in matches:
                amount_str = match.replace(',', '')
                amount_values.append((amount_str, x0, y0, text_clean))
            
            # 提取税率值（V3.8改进版 - 更精确的税率提取）
            # 支持格式：13%, 13, 0.13, 13.00%, 0.13等
            # 但需要排除金额和税额中的数字
            tax_rate_pattern = r'([0-9]+(?:\.[0-9]+)?)(?=\s*%|$)'
            tax_rate_matches = re.findall(tax_rate_pattern, text_clean)
            
            # 过滤掉金额和税额中的数字（金额通常是2位小数，税率通常是整数或1位小数）
            for match in tax_rate_matches:
                tax_rate_float = float(match)
                # 只提取合理的税率值（0-20之间，常见税率范围内）
                if 0 <= tax_rate_float <= 20:
                    # 排除金额（金额通常大于100或为整数）
                    if tax_rate_float > 100 or (tax_rate_float > 10 and '.' in match):
                        continue
                    tax_rate_values.append((tax_rate_float, x0, y0, text_clean))
            
            # 查找包含"购买方"或"销售方"的块，用于定位购方/销方区域
            if '购买方' in text_clean or '购方' in text_clean:
                buyer_region_y = y0
            if '销售方' in text_clean or '销方' in text_clean:
                seller_region_y = y0
        
        # 选择发票号码（选择最上面的20位数字）
        if invoice_numbers:
            invoice_numbers.sort(key=lambda x: x[1])  # 按y坐标排序
            for num, y in invoice_numbers:
                if len(num) == 20:
                    result['发票号码'] = num
                    break
            if not result['发票号码']:
                result['发票号码'] = invoice_numbers[0][0]
        
        # 选择开票日期（选择最上面的日期）
        if dates:
            dates.sort(key=lambda x: x[1])
            result['开票日期'] = dates[0][0]
        
        # 根据关键词定位购方和销方
        buyer_found = False
        seller_found = False
        buyer_label_y = None
        seller_label_y = None
        
        # V3.0: 先找到购方和销方的标签位置（修复：移除换行符后再匹配）
        for block in text_blocks:
            x0, y0, x1, y1, text, block_no, _ = block
            text_clean = text.strip()
            
            # V3.0: 移除换行符后再匹配，避免被换行分割
            text_no_newline = text.replace('\n', '')
            if '购买方' in text_no_newline or '购方名称' in text_no_newline:
                buyer_label_y = y0
                if self.debug:
                    logger.debug(f"找到购方标签: {text_clean} (y={y0:.2f})")
            elif '销售方' in text_no_newline or '销方名称' in text_no_newline:
                seller_label_y = y0
                if self.debug:
                    logger.debug(f"找到销方标签: {text_clean} (y={y0:.2f})")
        
        # 根据标签位置提取购方和销方名称
        if buyer_label_y is not None or seller_label_y is not None:
            # 第一轮：在标签附近提取（严格验证）
            for block in text_blocks:
                x0, y0, x1, y1, text, block_no, _ = block

                # V3.2: 跳过包含换行符的文本块（避免误识别）
                if '\n' in text:
                    continue

                text_clean = text.strip()

                # V3.1: 排除标签、价税合计和发票类型等相关内容
                excluded_keywords = ['购买方', '销售方', '购方名称', '销方名称', '名称', '统一社会', '识别号', '地址', '电话', '开户行', '账号', '：', ':', '价税合计', '合计', '大写', '小写', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖', '拾', '佰', '仟', '万', '亿', '圆', '整', '备', '备注', '注', '开票人', '发票', '电子发票', '增值税专用发票', '增值税普通发票', '发票号码', '开票日期', '项目名称', '规格型号', '单位', '数量', '单价', '金额', '税率', '税额', '税率/征收率']
                if any(keyword in text_clean for keyword in excluded_keywords):
                    continue

                # V3.2: 使用严格的验证函数检查是否为公司名称
                if self._validate_company_name(text_clean):
                    # V3.1: 调整位置判断，购方60像素，销方60像素
                    if buyer_label_y is not None and abs(y0 - buyer_label_y) < 60:
                        if not buyer_found:
                            result['购方名称'] = text_clean
                            buyer_found = True
                            if self.debug:
                                logger.debug(f"提取购方名称（标签附近）: {text_clean} (y={y0:.2f}, label_y={buyer_label_y:.2f})")
                    elif seller_label_y is not None and abs(y0 - seller_label_y) < 60:
                        if not seller_found:
                            result['销方名称'] = text_clean
                            seller_found = True
                            if self.debug:
                                logger.debug(f"提取销方名称（标签附近）: {text_clean} (y={y0:.2f}, label_y={seller_label_y:.2f})")

            
            # V3.1 第二轮：备用提取逻辑（在标签附近没找到时，尝试基于位置的备用方案）
            if not buyer_found or not seller_found:
                if self.debug:
                    logger.debug("第一轮未找到完整的购方/销方名称，尝试备用提取逻辑")

                for block in text_blocks:
                    x0, y0, x1, y1, text, block_no, _ = block
                    text_clean = text.strip()

                    # V3.2: 跳过包含换行符的文本块（避免误识别）
                    if '\n' in text:
                        continue

                    # V3.1: 排除标签、发票类型等相关内容
                    excluded_keywords = ['购买方', '销售方', '购方名称', '销方名称', '名称', '统一社会', '识别号', '地址', '电话', '开户行', '账号', '：', ':', '价税合计', '合计', '大写', '小写', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖', '拾', '佰', '仟', '万', '亿', '圆', '整', '发票', '号码', '日期', '备', '备注', '注', '开票人', '电子发票', '增值税专用发票', '增值税普通发票', '项目名称', '规格型号', '单位', '数量', '单价', '金额', '税率', '税额']
                    if any(keyword in text_clean for keyword in excluded_keywords):
                        continue

                    # V3.2: 使用严格的验证函数检查是否为公司名称
                    if self._validate_company_name(text_clean):
                        # 根据位置判断
                        if 90 <= y0 < 180:
                            if x0 < 250 and not buyer_found:
                                result['购方名称'] = text_clean
                                buyer_found = True
                                if self.debug:
                                    logger.debug(f"备用方案提取购方名称: {text_clean}")
                            elif x0 >= 250 and not seller_found:
                                result['销方名称'] = text_clean
                                seller_found = True
                                if self.debug:
                                    logger.debug(f"备用方案提取销方名称: {text_clean}")
            
            # V3.1: 第三轮：备用提取逻辑（如果在标签附近没找到，尝试基于位置的备用方案）
            if not buyer_found or not seller_found:
                if self.debug:
                    logger.debug("前两轮未找到完整的购方/销方名称，尝试备用提取逻辑")

                for block in text_blocks:
                    x0, y0, x1, y1, text, block_no, _ = block
                    text_clean = text.replace('\n', '').strip()

                    # V3.1: 排除标签、发票类型等相关内容
                    excluded_keywords = ['购买方', '销售方', '购方名称', '销方名称', '名称', '统一社会', '识别号', '地址', '电话', '开户行', '账号', '：', ':', '价税合计', '合计', '大写', '小写', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖', '拾', '佰', '仟', '万', '亿', '圆', '整', '发票', '号码', '日期', '备', '备注', '注', '开票人', '电子发票', '增值税专用发票', '增值税普通发票', '项目名称', '规格型号', '单位', '数量', '单价', '金额', '税率', '税额']
                    if any(keyword in text_clean for keyword in excluded_keywords):
                        continue

                    # V3.1: 检查是否像公司名称（增强验证）
                    if re.search(r'[\u4e00-\u9fa5]', text_clean) and len(text_clean) < 50:
                        # 检查是否包含公司名称特征关键词
                        company_keywords = ['有限公司', '股份有限公司', '有限责任公司', '公司', '集团', '企业', '厂', '店']
                        has_company_keyword = any(kw in text_clean for kw in company_keywords)

                        # 计算中文字符比例
                        chinese_chars = len(re.findall(r'[\u4e00-\u9fa5]', text_clean))
                        total_chars = len(text_clean)
                        chinese_ratio = chinese_chars / total_chars if total_chars > 0 else 0

                        # V3.1: 更严格的验证：中文字符占比超过70% 且 (有公司关键词或中文字符占比超过85%)
                        if chinese_ratio > 0.7 and (has_company_keyword or chinese_ratio > 0.85):
                            # 基于位置判断
                            if 90 <= y0 < 180:
                                if x0 < 250 and not buyer_found:
                                    result['购方名称'] = text_clean
                                    buyer_found = True
                                    if self.debug:
                                        logger.debug(f"备用方案提取购方名称: {text_clean}")
                                elif x0 >= 250 and not seller_found:
                                    result['销方名称'] = text_clean
                                    seller_found = True
                                    if self.debug:
                                        logger.debug(f"备用方案提取销方名称: {text_clean}")
        
        # 如果没有找到购方/销方，使用基于位置的备用方案
        if not buyer_found or not seller_found:
            for block in text_blocks:
                x0, y0, x1, y1, text, block_no, _ = block
                text_clean = text.strip()

                # V3.2: 跳过包含换行符的文本块（避免误识别）
                if '\n' in text:
                    continue

                # 排除标签
                if any(keyword in text_clean for keyword in ['名称', '统一社会', '购买方', '销售方', '购方', '销方', '信', '息', '：', ':', '发票', '号码', '日期', '备', '备注', '注', '开票人']):
                    continue

                # 紧急修复：使用严格的验证函数检查是否为公司名称
                if self._validate_company_name(text_clean):
                    # 基于位置判断
                    if 90 <= y0 < 180:
                        if x0 < 250 and not result['购方名称']:
                            result['购方名称'] = text_clean
                            if self.debug:
                                logger.debug(f"备用方案提取购方名称: {text_clean}")
                        elif x0 >= 250 and not result['销方名称']:
                            result['销方名称'] = text_clean
                            if self.debug:
                                logger.debug(f"备用方案提取销方名称: {text_clean}")
        
        # 提取项目名称（V3.7增强版，大幅提高识别率）
        # 先找到"项目名称"表头和"合计"区域的位置
        item_header_y = None
        total_region_y = None
        
        for block in text_blocks:
            x0, y0, x1, y1, text, block_no, _ = block
            text_clean = text.replace('\n', '').strip()
            
            if '项目名称' in text_clean:
                item_header_y = y0
            elif '合计' in text_clean and '价税合计' not in text_clean:
                total_region_y = y0
                # 不要break，继续查找确保找到正确的合计区域
        
        if self.debug:
            logger.debug(f"项目名称表头位置: y={item_header_y}")
            logger.debug(f"合计区域位置: y={total_region_y}")
        
        # 提取项目名称
        if item_header_y is not None and total_region_y is not None:
            # 收集所有候选项目名称
            item_name_candidates = []
            
            for block in text_blocks:
                x0, y0, x1, y1, text, block_no, _ = block
                text_clean = text.strip()
                
                # V3.7: 减少排除关键词，只保留核心排除项
                excluded_keywords = ['项目名称', '价税合计', '合计']
                if any(keyword in text_clean for keyword in excluded_keywords):
                    continue
                
                # 排除纯数字和纯符号
                if re.match(r'^[0-9.¥%]+$', text_clean):
                    continue
                
                # 排除大写金额特征（价税合计）
                chinese_num_pattern = r'[壹贰叁肆伍陆柒捌玖拾佰仟万亿圆整角分]'
                if re.search(chinese_num_pattern, text_clean):
                    if self.debug:
                        logger.debug(f"跳过大写金额: {text_clean}")
                    continue
                
                # 排除包含¥符号的文本（价税合计）
                if '¥' in text_clean:
                    if self.debug:
                        logger.debug(f"跳过包含¥的文本: {text_clean}")
                    continue
                
                # V3.7: 放宽项目名称识别条件
                # 条件1：包含*号（发票格式特征）
                # 条件2：包含中文字符
                # 条件3：长度合理
                has_chinese = re.search(r'[\u4e00-\u9fa5]', text_clean)
                has_asterisk = '*' in text_clean
                length_ok = 2 < len(text_clean) < 200
                
                if (has_asterisk and length_ok) or (has_chinese and length_ok):
                    # V3.7: 扩大x0范围到500，允许更宽的区域
                    if item_header_y < y0 < total_region_y and x0 < 500:
                        if self.debug:
                            logger.debug(f"候选项目名称: {text_clean} (y={y0:.2f}, x0={x0:.2f})")
                        
                        # 计算候选项目的优先级分数
                        score = 0
                        if has_asterisk:
                            score += 10  # 包含*号优先
                        if len(text_clean) < 60:
                            score += 5  # 长度适中优先
                        if not re.search(r'[0-9]{8,}', text_clean):
                            score += 3  # 不包含长数字串优先
                        if x0 < 100:
                            score += 5  # 靠左优先
                        
                        # V3.0: 对多行文本块，提取第一行作为项目名称
                        if '\n' in text_clean:
                            first_line = text_clean.split('\n')[0].strip()
                            if len(first_line) > 2:
                                text_clean = first_line
                        
                        item_name_candidates.append((score, y0, text_clean))
            
            # 按优先级排序，选择分数最高的
            if item_name_candidates:
                item_name_candidates.sort(key=lambda x: (-x[0], x[1]))  # 按分数降序，y坐标升序
                best_candidate = item_name_candidates[0][2]
                
                # V2.0: 简化项目名称，只保留大类和物料名称
                simplified_item_name = self._simplify_item_name(best_candidate)
                result['项目名称'] = simplified_item_name
                if self.debug:
                    logger.debug(f"提取项目名称（原始）: {best_candidate}")
                    logger.debug(f"提取项目名称（简化）: {simplified_item_name}")
            else:
                # V3.7: 备用方案 - 如果没有找到项目名称，尝试其他方法
                if self.debug:
                    logger.debug(f"未找到符合条件的项目名称，尝试备用方案")
                
                # 备用方案1：查找包含产品类型关键词的文本
                product_keywords = ['制品', '剂', '料', '品', '具', '备', '械', '仪', '品', '材', '物', '药', '品', '包装', '袋', '桶', '板', '卡', '纸', '泵', '酶', '酸', '碱', '盐', '粉']
                for block in text_blocks:
                    x0, y0, x1, y1, text, block_no, _ = block
                    text_clean = text.strip()
                    
                    # 排除表头和标签
                    excluded_keywords = ['项目名称', '价税合计', '合计']
                    if any(keyword in text_clean for keyword in excluded_keywords):
                        continue
                    
                    # 排除纯数字和纯符号
                    if re.match(r'^[0-9.¥%]+$', text_clean):
                        continue
                    
                    # 检查是否包含产品关键词
                    if any(kw in text_clean for kw in product_keywords):
                        # 检查是否在合理位置
                        if item_header_y < y0 < total_region_y and x0 < 450:
                            if len(text_clean) > 3 and len(text_clean) < 100:
                                simplified = self._simplify_item_name(text_clean)
                                result['项目名称'] = simplified
                                if self.debug:
                                    logger.debug(f"备用方案提取项目名称: {text_clean} -> {simplified}")
                                break
                
                # 备用方案2：从整个PDF文本中查找
                if result['项目名称'] == '未识别':
                    all_text = ' '.join([b[4].strip() for b in text_blocks])
                    
                    # 查找包含产品关键词的文本
                    for kw in product_keywords:
                        # 查找关键词周围的内容
                        idx = all_text.find(kw)
                        if idx >= 0:
                            # 提取关键词前后20个字符
                            start = max(0, idx - 20)
                            end = min(len(all_text), idx + len(kw) + 20)
                            candidate = all_text[start:end].strip()
                            
                            # 清理候选文本
                            candidate = re.sub(r'[0-9.¥%]+', '', candidate).strip()
                            
                            if len(candidate) > 3 and len(candidate) < 50:
                                simplified = self._simplify_item_name(candidate)
                                result['项目名称'] = simplified
                                if self.debug:
                                    logger.debug(f"备用方案2提取项目名称: {candidate} -> {simplified}")
                                break
        
        # 提取金额、税额、价税合计
        amount_values.sort(key=lambda x: (x[2], x[1]))  # 先按y再按x排序
        
        # 查找"价税合计"标签附近的位置
        total_label_y = None
        for block in text_blocks:
            x0, y0, x1, y1, text, block_no, _ = block
            if '价税合计' in text:
                total_label_y = y0
                break
        
        # 分析金额数值的位置，分配到不同的字段
        total_candidates = []  # 价税合计候选
        amount_tax_pairs = []  # 金额和税额对
        
        for amount_str, x0, y0, original in amount_values:
            amount_float = float(amount_str)
            
            # 收集价税合计候选（底部且金额较大的）
            if y0 > 260:
                total_candidates.append((amount_str, amount_float, y0, x0))
            
            # 收集金额和税额对（中下部）
            elif y0 > 200:
                amount_tax_pairs.append((amount_str, amount_float, y0, x0))
        
        # 智能分配金额、税额、价税合计
        if len(amount_tax_pairs) >= 2:
            # 如果有多对金额，尝试找到满足 金额+税额=价税合计 的组合
            # 先找最大的作为价税合计候选
            max_amount = max(amount_tax_pairs, key=lambda x: x[1])
            
            # 验证是否有满足关系的组合
            for i, (amount_str, amount_val, _, _) in enumerate(amount_tax_pairs):
                for j, (tax_str, tax_val, _, _) in enumerate(amount_tax_pairs):
                    if i == j:
                        continue
                    
                    # 检查是否有价税合计等于金额+税额
                    calculated_total = amount_val + tax_val
                    for total_str, total_val, _, _ in total_candidates:
                        if abs(calculated_total - total_val) < 0.01:
                            # 找到匹配的组合
                            result['金额'] = amount_str
                            result['税额'] = tax_str
                            result['价税合计'] = total_str
                            if self.debug:
                                logger.debug(f"找到匹配组合: 金额={amount_str}, 税额={tax_str}, 合计={total_str}")
                            break
                    if result['金额'] and result['税额'] and result['价税合计']:
                        break
                if result['金额'] and result['税额'] and result['价税合计']:
                    break
        
        # 如果没有找到匹配的组合，使用原始逻辑
        if not result['金额'] or not result['税额'] or not result['价税合计']:
            for amount_str, x0, y0, original in amount_values:
                # 价税合计通常在底部，且金额较大
                if y0 > 260:
                    # 底部区域，优先作为价税合计
                    if not result['价税合计']:
                        result['价税合计'] = amount_str
                        if self.debug:
                            logger.debug(f"提取价税合计（底部）: {amount_str}")
                    elif not result['金额']:
                        result['金额'] = amount_str
                        if self.debug:
                            logger.debug(f"提取金额（底部）: {amount_str}")
                elif y0 > 200:
                    # 中下部区域，可能是金额或税额
                    if not result['金额']:
                        result['金额'] = amount_str
                        if self.debug:
                            logger.debug(f"提取金额（中下部）: {amount_str}")
                    elif not result['税额']:
                        result['税额'] = amount_str
                        if self.debug:
                            logger.debug(f"提取税额（中下部）: {amount_str}")
                    elif not result['价税合计']:
                        result['价税合计'] = amount_str
                        if self.debug:
                            logger.debug(f"提取价税合计（中下部）: {amount_str}")
        
        # 如果还没有找到价税合计，尝试从所有金额中找最大的
        if not result['价税合计'] and amount_values:
            max_amount = max(amount_values, key=lambda x: float(x[0]))
            result['价税合计'] = max_amount[0]
            if self.debug:
                logger.debug(f"从所有金额中选择最大的作为价税合计: {max_amount[0]}")
        
        # V3.9: 彻底修复税率识别逻辑
        if result['金额'] and result['税额']:
            try:
                amount = float(result['金额'])
                tax = float(result['税额'])
                
                # 检查税额是否大于金额（识别错误）
                if tax > amount:
                    if self.debug:
                        logger.debug(f"税额大于金额: 税额={tax}, 金额={amount}，自动交换")
                    # 交换金额和税额
                    amount, tax = tax, amount
                    result['金额'] = f"{amount:.2f}"
                    result['税额'] = f"{tax:.2f}"
                
                # 计算实际税率
                calculated_tax_rate = (tax / amount * 100) if amount > 0 else 0
                
                # 常见税率列表（中国增值税标准税率）
                common_tax_rates = [0, 1, 3, 5, 6, 9, 11, 13]
                
                # 检查计算出的税率是否接近常见税率（允许±0.3%的误差）
                matched_rate = None
                min_diff = float('inf')
                
                for rate in common_tax_rates:
                    diff = abs(calculated_tax_rate - rate)
                    if diff < min_diff and diff <= 0.3:
                        min_diff = diff
                        matched_rate = rate
                
                if matched_rate is not None:
                    # 使用匹配的常见税率
                    result['税率'] = f"{matched_rate:.2f}%"
                    if self.debug:
                        logger.debug(f"计算税率: {calculated_tax_rate:.2f}% -> 匹配常见税率: {result['税率']}（误差: {min_diff:.2f}%）")
                else:
                    # 如果计算出的税率不是常见税率，尝试从提取的税率值中选择
                    if tax_rate_values:
                        # 选择最接近计算税率的值
                        closest_rate = min(tax_rate_values, key=lambda x: abs(x[0] - calculated_tax_rate))
                        result['税率'] = f"{closest_rate[0]:.2f}%"
                        if self.debug:
                            logger.debug(f"使用提取的税率: {result['税率']}（计算税率: {calculated_tax_rate:.2f}%）")
                    else:
                        # 使用计算出的税率（限制在0-100%范围内）
                        if 0 <= calculated_tax_rate <= 100:
                            result['税率'] = f"{calculated_tax_rate:.2f}%"
                            if self.debug:
                                logger.debug(f"使用计算税率: {result['税率']}")
                        else:
                            # 异常税率，设置为0
                            result['税率'] = '0.00%'
                            if self.debug:
                                logger.debug(f"异常税率: {calculated_tax_rate:.2f}%，设置为0.00%")
            except ValueError as e:
                if self.debug:
                    logger.debug(f"税率计算失败: {str(e)}")
                result['税率'] = '0.00%'
        
        # 数据验证和智能计算
        # 如果有金额和税额但没有价税合计，计算价税合计
        if result['金额'] and result['税额'] and not result['价税合计']:
            try:
                amount = float(result['金额'])
                tax = float(result['税额'])
                result['价税合计'] = f"{amount + tax:.2f}"
                if self.debug:
                    logger.debug(f"计算价税合计: {result['金额']} + {result['税额']} = {result['价税合计']}")
            except ValueError:
                pass
        
        # 如果有价税合计和税额但没有金额，计算金额
        if result['价税合计'] and result['税额'] and not result['金额']:
            try:
                total = float(result['价税合计'])
                tax = float(result['税额'])
                result['金额'] = f"{total - tax:.2f}"
                if self.debug:
                    logger.debug(f"计算金额: {result['价税合计']} - {result['税额']} = {result['金额']}")
            except ValueError:
                pass
        
        # 如果有价税合计和金额但没有税额，计算税额
        if result['价税合计'] and result['金额'] and not result['税额']:
            try:
                total = float(result['价税合计'])
                amount = float(result['金额'])
                result['税额'] = f"{total - amount:.2f}"
                if self.debug:
                    logger.debug(f"计算税额: {result['价税合计']} - {result['金额']} = {result['税额']}")
            except ValueError:
                pass
        
        # 确保所有字段都有默认值
        if not result['购方名称']:
            result['购方名称'] = '未识别'
        if not result['销方名称']:
            result['销方名称'] = '未识别'
        if not result['项目名称']:
            result['项目名称'] = '未识别'
        if not result['税率']:
            result['税率'] = '0.00%'
        if not result['金额']:
            result['金额'] = '0.00'
        if not result['税额']:
            result['税额'] = '0.00'
        if not result['价税合计']:
            result['价税合计'] = '0.00'
        
        # 字段内容验证
        # 验证项目名称：不应该包含大写金额字符或¥符号
        if result['项目名称'] != '未识别':
            chinese_num_pattern = r'[壹贰叁肆伍陆柒捌玖拾佰仟万亿圆整角分]'
            if re.search(chinese_num_pattern, result['项目名称']) or '¥' in result['项目名称']:
                if self.debug:
                    logger.debug(f"项目名称验证失败，包含大写金额或¥符号: {result['项目名称']}")
                result['项目名称'] = '未识别'
        
        # V2.0: 验证购方名称：允许包含少量数字，只要主要是中文
        if result['购方名称'] != '未识别':
            chinese_chars = len(re.findall(r'[\u4e00-\u9fa5]', result['购方名称']))
            total_chars = len(result['购方名称'])
            chinese_ratio = chinese_chars / total_chars if total_chars > 0 else 0
            
            if chinese_ratio < 0.5:
                if self.debug:
                    logger.debug(f"购方名称验证失败，中文字符占比过低: {result['购方名称']} (ratio={chinese_ratio:.2f})")
                result['购方名称'] = '未识别'
        
        # V2.0: 验证销方名称：允许包含少量数字，只要主要是中文
        if result['销方名称'] != '未识别':
            chinese_chars = len(re.findall(r'[\u4e00-\u9fa5]', result['销方名称']))
            total_chars = len(result['销方名称'])
            chinese_ratio = chinese_chars / total_chars if total_chars > 0 else 0

            if chinese_ratio < 0.5:
                if self.debug:
                    logger.debug(f"销方名称验证失败，中文字符占比过低: {result['销方名称']} (ratio={chinese_ratio:.2f})")
                result['销方名称'] = '未识别'

        # V3.2: 特殊处理京东发票 - 在所有提取完成后，尝试修复京东发票的销方名称
        # 京东发票特征：购方名称包含两个公司名称（用换行符分隔），销方名称未识别
        if result['销方名称'] == '未识别' and result['购方名称'] != '未识别':
            # 检查购方名称是否包含两个公司名称（用换行符分隔）
            if '\n' in result['购方名称']:
                lines = result['购方名称'].split('\n')
                # 必须至少有两行
                if len(lines) >= 2:
                    # 智能判断：找出真正属于购方的公司名称
                    # 规则：包含"绿微康"的是购方名称（用户的实际公司）
                    potential_seller = None
                    potential_buyer = None
                    
                    for line in lines:
                        line_clean = line.strip()
                        if not line_clean:
                            continue
                        # 检查是否为用户的公司名称
                        if '绿微康' in line_clean:
                            potential_buyer = line_clean
                        elif self._validate_company_name(line_clean):
                            potential_seller = line_clean
                    
                    # 如果找到了购方和销方，就设置
                    if potential_buyer and potential_seller:
                        result['购方名称'] = potential_buyer
                        result['销方名称'] = potential_seller
                        if self.debug:
                            logger.debug(f"京东发票特殊处理（换行符）：修正购方/销方名称，销方={potential_seller}, 购方={potential_buyer}")
            
            # V3.2: 检查购方名称是否包含两个连接的公司名称（无换行符）
            # 京东发票特征：购方名称很长，包含两个公司名称直接连接
            elif len(result['购方名称']) > 20:
                buyer_name = result['购方名称']
                
                # 尝试找到两个公司名称的分隔点
                # 策略1：寻找"有限公司"或"公司"等关键词
                company_keywords = ['有限公司', '股份有限公司', '有限责任公司', '公司', '集团']
                
                for keyword in company_keywords:
                    # 找到关键词的所有位置
                    positions = []
                    pos = 0
                    while True:
                        pos = buyer_name.find(keyword, pos)
                        if pos == -1:
                            break
                        positions.append(pos)
                        pos += 1
                    
                    # 如果找到两个或更多的关键词，尝试分割
                    if len(positions) >= 2:
                        # 智能判断：找出真正属于购方的公司名称
                        # 尝试在第一个关键词后面分割
                        split_pos = positions[0] + len(keyword)
                        company1 = buyer_name[:split_pos].strip()
                        company2 = buyer_name[split_pos:].strip()
                        
                        # 验证两者是否都为公司名称
                        if self._validate_company_name(company1) and self._validate_company_name(company2):
                            # 检查哪个是真正的购方名称（包含"绿微康"的）
                            if '绿微康' in company1:
                                result['购方名称'] = company1
                                result['销方名称'] = company2
                            elif '绿微康' in company2:
                                result['购方名称'] = company2
                                result['销方名称'] = company1
                            else:
                                # 如果都不包含"绿微康"，默认第一个是购方
                                result['购方名称'] = company1
                                result['销方名称'] = company2
                            
                            if self.debug:
                                logger.debug(f"京东发票特殊处理（无换行符）：修正购方/销方名称，销方={result['销方名称']}, 购方={result['购方名称']}")
                            break
        
        # 验证金额字段：应该是有效的数字
        for field in ['金额', '税额', '价税合计']:
            if result[field] != '0.00':
                try:
                    float(result[field])
                except ValueError:
                    if self.debug:
                        logger.debug(f"{field}验证失败，不是有效数字: {result[field]}")
                    result[field] = '0.00'
        
        if self.debug:
            logger.info(f"提取结果: {result}")
        
        return result
    
    def _identify_invoice_type(self, text: str) -> str:
        """
        识别发票类型
        """
        text_lower = text.lower()
        
        # 检查是否为增值税发票
        for keyword in self.special_invoice_keywords:
            if keyword in text_lower:
                return '增值税发票'
        
        # 检查是否为电子发票
        for keyword in self.electronic_invoice_keywords:
            if keyword in text_lower:
                return '电子发票'
        
        # 默认为普通发票
        return '普通发票'


class InvoiceProcessor:
    """发票处理器"""
    
    def __init__(self):
        self.parser = InvoiceParser()
        
    def process_single_invoice(self, pdf_path: str) -> Optional[Dict]:
        """
        处理单个发票文件
        """
        return self.parser.parse_pdf_invoice(pdf_path)
    
    def process_folder(self, folder_path: str) -> List[Dict]:
        """
        处理文件夹中的所有发票
        """
        import os
        invoices = []
        
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.pdf'):
                    pdf_path = os.path.join(root, file)
                    invoice_info = self.process_single_invoice(pdf_path)
                    
                    if invoice_info:
                        # 添加文件路径和文件名信息
                        invoice_info['文件路径'] = pdf_path
                        invoice_info['file_name'] = file
                        invoices.append(invoice_info)
        
        return invoices


# 测试函数
def test_invoice_parser():
    """测试发票解析器"""
    import os
    # 创建一个模拟的发票PDF文件用于测试
    sample_pdf_path = "/tmp/sample_invoice.pdf"
    
    # 创建一个简单的PDF用于测试
    try:
        doc = fitz.open()
        page = doc.new_page()
        text = """
        增值税电子普通发票
        发票代码: 1234567890123
        发票号码: 2301234567
        开票日期: 2023年05月15日
        校验码: 1234567890
        购买方:
        名称: 测试公司有限公司
        纳税人识别号: 911234567890123456
        地址电话: 北京市朝阳区xxx街道 010-12345678
        销售方:
        名称: 卖方公司有限公司
        纳税人识别号: 921234567890123456
        地址电话: 上海市浦东新区xxx路 021-87654321
        货物或应税劳务名称: 办公用品
        规格型号: 
        单位: 批
        数量: 1
        单价: 1000.00
        金额: 1000.00
        税率: 13%
        税额: 130.00
        价税合计(大写): 壹仟壹佰叁拾元整
        价税合计(小写): ¥1130.00
        """
        page.insert_text((50, 50), text)
        doc.save(sample_pdf_path)
        doc.close()
        
        # 测试解析
        processor = InvoiceProcessor()
        result = processor.process_single_invoice(sample_pdf_path)
        
        print("发票解析结果:")
        for key, value in result.items():
            print(f"{key}: {value}")
        
        # 删除测试文件
        if os.path.exists(sample_pdf_path):
            os.remove(sample_pdf_path)
        
        return result
    except Exception as e:
        print(f"测试过程中出现错误: {str(e)}")
        return None


if __name__ == "__main__":
    test_invoice_parser()
