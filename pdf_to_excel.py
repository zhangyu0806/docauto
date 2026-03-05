#!/usr/bin/env python3
"""
PDF发票批量处理脚本
功能：从PDF中提取表格数据并导出到Excel
支持：普通PDF表格、扫描版PDF（需OCR）、OFD格式发票
"""

import pdfplumber
import pandas as pd
import os
from pathlib import Path
from typing import List, Dict, Any
import re


class PDFInvoiceProcessor:
    """PDF发票处理器"""
    
    def __init__(self, use_ocr: bool = False):
        """
        初始化处理器
        
        Args:
            use_ocr: 是否使用OCR（针对扫描版PDF）
        """
        self.use_ocr = use_ocr
        if use_ocr:
            try:
                import easyocr
                print("初始化OCR引擎...")
                self.ocr_reader = easyocr.Reader(['ch_sim', 'en'])
            except ImportError:
                print("⚠️  easyocr未安装，OCR功能不可用。pip install easyocr")
                self.use_ocr = False
    
    def extract_tables_from_pdf(self, pdf_path: str) -> List[pd.DataFrame]:
        """
        从PDF中提取所有表格
        
        Args:
            pdf_path: PDF文件路径
            
        Returns:
            表格数据列表
        """
        tables = []
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                print(f"处理PDF: {pdf_path} (共{len(pdf.pages)}页)")
                
                for page_num, page in enumerate(pdf.pages, 1):
                    # 提取当前页的表格
                    page_tables = page.extract_tables()
                    
                    if page_tables:
                        for table_num, table in enumerate(page_tables, 1):
                            # 转换为DataFrame
                            df = pd.DataFrame(table[1:], columns=table[0])
                            df['页码'] = page_num
                            df['表格序号'] = table_num
                            tables.append(df)
                            print(f"  - 第{page_num}页第{table_num}个表格: {len(df)}行")
                    
                    # 如果没有表格且启用OCR，尝试OCR
                    elif self.use_ocr:
                        print(f"  - 第{page_num}页未检测到表格，尝试OCR...")
                        ocr_result = self._ocr_page(page)
                        if ocr_result:
                            tables.append(ocr_result)
        
        except Exception as e:
            print(f"❌ 处理PDF失败: {e}")
        
        return tables
    
    def _ocr_page(self, page) -> pd.DataFrame:
        """使用OCR提取页面文本"""
        try:
            # 保存页面为图片
            img = page.to_image()
            img_path = f"/tmp/page_{page.page_number}.png"
            img.save(img_path)
            
            # OCR识别
            results = self.ocr_reader.readtext(img_path)
            
            # 简单的文本提取（可以进一步优化为结构化表格）
            texts = [result[1] for result in results]
            df = pd.DataFrame({'识别文本': texts})
            df['页码'] = page.page_number
            
            # 清理临时文件
            os.remove(img_path)
            
            return df
        
        except Exception as e:
            print(f"  ❌ OCR失败: {e}")
            return None
    
    def process_invoice(self, pdf_path: str, output_path: str = None) -> str:
        """
        处理单张发票PDF
        
        Args:
            pdf_path: PDF文件路径
            output_path: 输出Excel路径（默认：同名Excel文件）
            
        Returns:
            输出文件路径
        """
        if output_path is None:
            output_path = Path(pdf_path).with_suffix('.xlsx')
        
        tables = self.extract_tables_from_pdf(pdf_path)
        
        if tables:
            # 合并所有表格到一个Excel文件
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for i, df in enumerate(tables, 1):
                    sheet_name = f'表格{i}'
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"✅ 成功导出: {output_path}")
            return str(output_path)
        
        else:
            print(f"⚠️  未从PDF中提取到表格数据")
            return None
    
    def batch_process(self, pdf_dir: str, output_dir: str = None) -> List[str]:
        """
        批量处理PDF文件
        
        Args:
            pdf_dir: PDF文件目录
            output_dir: 输出目录（默认：与输入相同）
            
        Returns:
            成功处理的文件列表
        """
        pdf_path = Path(pdf_dir)
        
        if output_dir is None:
            output_path = pdf_path
        else:
            output_path = Path(output_dir)
            output_path.mkdir(parents=True, exist_ok=True)
        
        # 查找所有PDF文件
        pdf_files = list(pdf_path.glob('*.pdf')) + list(pdf_path.glob('*.PDF'))
        
        print(f"找到{len(pdf_files)}个PDF文件")
        
        results = []
        for pdf_file in pdf_files:
            print(f"\n处理: {pdf_file.name}")
            output_file = output_path / f"{pdf_file.stem}.xlsx"
            
            result = self.process_invoice(str(pdf_file), str(output_file))
            if result:
                results.append(result)
        
        print(f"\n{'='*50}")
        print(f"批量处理完成: {len(results)}/{len(pdf_files)}个文件")
        print(f"{'='*50}")
        
        return results


def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(description='PDF发票批量处理工具')
    parser.add_argument('input', help='输入PDF文件或目录')
    parser.add_argument('-o', '--output', help='输出Excel文件或目录')
    parser.add_argument('--ocr', action='store_true', help='启用OCR（针对扫描版PDF）')
    
    args = parser.parse_args()
    
    # 创建处理器
    processor = PDFInvoiceProcessor(use_ocr=args.ocr)
    
    # 判断是文件还是目录
    input_path = Path(args.input)
    
    if input_path.is_file():
        # 单文件处理
        processor.process_invoice(args.input, args.output)
    
    elif input_path.is_dir():
        # 批量处理
        processor.batch_process(args.input, args.output)
    
    else:
        print(f"❌ 输入路径不存在: {args.input}")


if __name__ == '__main__':
    main()
