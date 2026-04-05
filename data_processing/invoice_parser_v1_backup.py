#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
发票解析模块 - 改进版
负责从PDF文件中提取发票信息
使用基于文本块结构的解析方法
"""

import fitz  # PyMuPDF
import re
import json
from datetime import datetime
from typing import Dict, List, Optional
import logging

# 配置日志
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class InvoiceParser:
    """发票解析器 - V2.0 改进版"""
    
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
        使用基于文本块的方法
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
                
                # 提取发票信息
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
    
    def _extract_invoice_info_from_blocks(self, blocks: List, page) -> Dict:
        """
        从文本块中提取发票信息
        基于位置关系和文本内容 - 改进版
        """
        result = {
            '开票日期': '',
            '发票号码': '',
            '购方名称': '',
            '销方名称': '',
            '项目名称': '',
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
        
        # 先找到购方和销方的标签位置
        for block in text_blocks:
            x0, y0, x1, y1, text, block_no, _ = block
            text_clean = text.strip()
            
            if '购买方' in text_clean or '购方名称' in text_clean:
                buyer_label_y = y0
            elif '销售方' in text_clean or '销方名称' in text_clean:
                seller_label_y = y0
        
        # 根据标签位置提取购方和销方名称
        if buyer_label_y is not None or seller_label_y is not None:
            # 第一轮：在标签附近提取（严格验证）
            for block in text_blocks:
                x0, y0, x1, y1, text, block_no, _ = block
                text_clean = text.replace('\n', '').strip()
                
                # 排除标签本身和价税合计相关内容
                excluded_keywords = ['购买方', '销售方', '购方名称', '销方名称', '名称', '统一社会', '识别号', '地址', '电话', '开户行', '账号', '：', ':', '价税合计', '合计', '大写', '小写', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖', '拾', '佰', '仟', '万', '亿', '圆', '整', '备', '备注', '注', '开票人']
                if any(keyword in text_clean for keyword in excluded_keywords):
                    continue
                
                # V2.0: 放宽验证条件，只要主要是中文就可以
                # 检查是否像公司名称（主要是中文，长度合适）
                if re.search(r'[\u4e00-\u9fa5]', text_clean) and len(text_clean) < 50:
                    # 计算中文字符比例
                    chinese_chars = len(re.findall(r'[\u4e00-\u9fa5]', text_clean))
                    total_chars = len(text_clean)
                    chinese_ratio = chinese_chars / total_chars if total_chars > 0 else 0
                    
                    # V2.0: 只要中文字符占比超过50%就认为是公司名称
                    if chinese_ratio > 0.5:
                        # 根据y坐标判断是购方还是销方
                        if buyer_label_y is not None and abs(y0 - buyer_label_y) < 50:
                            if not buyer_found:
                                result['购方名称'] = text_clean
                                buyer_found = True
                                if self.debug:
                                    logger.debug(f"提取购方名称（标签附近）: {text_clean} (y={y0:.2f}, label_y={buyer_label_y:.2f})")
                        elif seller_label_y is not None and abs(y0 - seller_label_y) < 50:
                            if not seller_found:
                                result['销方名称'] = text_clean
                                seller_found = True
                                if self.debug:
                                    logger.debug(f"提取销方名称（标签附近）: {text_clean} (y={y0:.2f}, label_y={seller_label_y:.2f})")
            
            # V2.0: 第二轮：备用提取逻辑（如果在标签附近没找到，尝试基于位置的备用方案）
            if not buyer_found or not seller_found:
                if self.debug:
                    logger.debug("第一轮未找到完整的购方/销方名称，尝试备用提取逻辑")
                
                for block in text_blocks:
                    x0, y0, x1, y1, text, block_no, _ = block
                    text_clean = text.replace('\n', '').strip()
                    
                    # 排除标签本身和价税合计相关内容
                    excluded_keywords = ['购买方', '销售方', '购方名称', '销方名称', '名称', '统一社会', '识别号', '地址', '电话', '开户行', '账号', '：', ':', '价税合计', '合计', '大写', '小写', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖', '拾', '佰', '仟', '万', '亿', '圆', '整', '发票', '号码', '日期', '备', '备注', '注', '开票人']
                    if any(keyword in text_clean for keyword in excluded_keywords):
                        continue
                    
                    # 检查是否像公司名称
                    if re.search(r'[\u4e00-\u9fa5]', text_clean) and len(text_clean) < 50:
                        # 计算中文字符比例
                        chinese_chars = len(re.findall(r'[\u4e00-\u9fa5]', text_clean))
                        total_chars = len(text_clean)
                        chinese_ratio = chinese_chars / total_chars if total_chars > 0 else 0
                        
                        # V2.0: 只要中文字符占比超过50%就认为是公司名称
                        if chinese_ratio > 0.5:
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
                
                # 排除标签
                if any(keyword in text_clean for keyword in ['名称', '统一社会', '购买方', '销售方', '购方', '销方', '信', '息', '：', ':', '发票', '号码', '日期', '备', '备注', '注', '开票人']):
                    continue
                
                # 检查是否像公司名称
                if re.match(r'^[\u4e00-\u9fa5]{4,}', text_clean) and len(text_clean) < 50:
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
        
        # 提取项目名称（改进版，排除价税合计干扰）
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
            for block in text_blocks:
                x0, y0, x1, y1, text, block_no, _ = block
                text_clean = text.strip()
                
                # 排除表头和标签
                excluded_keywords = ['项目名称', '货物', '商品', '规格', '型号', '单位', '数量', '单价', '金额', '税率', '税额', '合计', '价税合计', '名称', '日期', '号码', '购买方', '销售方']
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
                
                # 排除地址、电话等信息
                if any(keyword in text_clean for keyword in ['地址', '电话', '开户', '账号', '识别号', '省', '市', '区', '路', '街道', '银行']):
                    continue
                
                # 检查是否像项目名称（包含*号或者有合理的长度）
                if (('*' in text_clean and len(text_clean) > 5) or 
                    (len(text_clean) > 5 and len(text_clean) < 100 and 
                     not re.match(r'^[0-9,\.¥%]+$', text_clean))):
                    
                    # 根据位置判断：在项目名称表头下方，合计区域上方，且在左侧
                    if item_header_y < y0 < total_region_y and x0 < 250:
                        if self.debug:
                            logger.debug(f"候选项目名称: {text_clean} (y={y0:.2f}, x0={x0:.2f})")
                        if not result['项目名称']:
                            raw_item_name = text_clean.replace('\n', '').strip()
                            # V2.0: 简化项目名称，只保留大类和物料名称
                            simplified_item_name = self._simplify_item_name(raw_item_name)
                            result['项目名称'] = simplified_item_name
                            if self.debug:
                                logger.debug(f"提取项目名称（原始）: {raw_item_name}")
                                logger.debug(f"提取项目名称（简化）: {simplified_item_name} (y={y0:.2f})")
                            break
            if self.debug and not result['项目名称']:
                logger.debug(f"未找到符合条件的项目名称")
        
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
                        # 添加文件路径信息
                        invoice_info['文件路径'] = pdf_path
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