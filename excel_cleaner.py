#!/usr/bin/env python3
"""
Excel数据清洗脚本
功能：去重、格式统一、空值处理、数据标准化
"""

import pandas as pd
import numpy as np
from pathlib import Path
from typing import List, Dict, Any, Optional
import re
from datetime import datetime


class ExcelCleaner:
    """Excel数据清洗器"""
    
    def __init__(self):
        self.cleaning_report = []
    
    def load_excel(self, file_path: str, sheet_name: str = None) -> pd.DataFrame:
        """
        加载Excel文件
        
        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称（默认：第一个）
            
        Returns:
            DataFrame
        """
        try:
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(file_path)
            
            print(f"✅ 加载成功: {file_path} ({len(df)}行 x {len(df.columns)}列)")
            return df
        
        except Exception as e:
            print(f"❌ 加载失败: {e}")
            return None
    
    def remove_duplicates(self, df: pd.DataFrame, 
                         subset: List[str] = None, 
                         keep: str = 'first') -> pd.DataFrame:
        """
        去除重复行
        
        Args:
            df: 输入DataFrame
            subset: 用于判断重复的列（默认：所有列）
            keep: 保留策略 ('first', 'last', False)
            
        Returns:
            去重后的DataFrame
        """
        before_count = len(df)
        df_cleaned = df.drop_duplicates(subset=subset, keep=keep)
        after_count = len(df_cleaned)
        
        removed = before_count - after_count
        if removed > 0:
            self.cleaning_report.append(f"去重: 删除了{removed}条重复行")
            print(f"✅ 去重: 删除了{removed}条重复行")
        
        return df_cleaned
    
    def remove_empty_rows(self, df: pd.DataFrame, 
                         threshold: float = 0.5) -> pd.DataFrame:
        """
        删除空值过多的行
        
        Args:
            df: 输入DataFrame
            threshold: 空值比例阈值（超过此比例的行将被删除）
            
        Returns:
            清洗后的DataFrame
        """
        before_count = len(df)
        
        # 计算每行的空值比例
        null_ratio = df.isnull().sum(axis=1) / len(df.columns)
        df_cleaned = df[null_ratio < threshold]
        
        after_count = len(df_cleaned)
        removed = before_count - after_count
        
        if removed > 0:
            self.cleaning_report.append(f"删除空行: 删除了{removed}条空值过多的行")
            print(f"✅ 删除空行: 删除了{removed}条空值过多的行")
        
        return df_cleaned
    
    def standardize_dates(self, df: pd.DataFrame, 
                         date_columns: List[str] = None,
                         date_format: str = '%Y-%m-%d') -> pd.DataFrame:
        """
        标准化日期格式
        
        Args:
            df: 输入DataFrame
            date_columns: 日期列名列表（自动检测或手动指定）
            date_format: 目标日期格式
            
        Returns:
            标准化后的DataFrame
        """
        df_cleaned = df.copy()
        
        # 自动检测日期列
        if date_columns is None:
            date_columns = []
            date_keywords = ['日期', '时间', 'date', 'time', '签订', '创建', '更新']
            for col in df_cleaned.columns:
                col_lower = str(col).lower()
                # 先检查列名是否包含日期关键词
                if any(kw in col_lower for kw in date_keywords):
                    sample = df_cleaned[col].dropna().head(100)
                    if len(sample) > 0:
                        try:
                            pd.to_datetime(sample, errors='raise', format='mixed')
                            date_columns.append(col)
                        except:
                            pass
                elif df_cleaned[col].dtype == 'object':
                    # 对非数值列，检查是否看起来像日期
                    sample = df_cleaned[col].dropna().head(20).astype(str)
                    date_like = sample.str.match(r'^\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}')
                    if date_like.sum() > len(sample) * 0.5:
                        date_columns.append(col)
        
        # 转换日期列
        for col in date_columns:
            try:
                df_cleaned[col] = pd.to_datetime(df_cleaned[col], errors='coerce')
                df_cleaned[col] = df_cleaned[col].dt.strftime(date_format)
                self.cleaning_report.append(f"日期格式化: {col} → {date_format}")
                print(f"✅ 日期格式化: {col} → {date_format}")
            except Exception as e:
                print(f"⚠️  日期转换失败 ({col}): {e}")
        
        return df_cleaned
    
    def standardize_phone(self, df: pd.DataFrame, 
                         phone_columns: List[str] = None) -> pd.DataFrame:
        """
        标准化电话号码格式
        
        Args:
            df: 输入DataFrame
            phone_columns: 电话列名列表
            
        Returns:
            标准化后的DataFrame
        """
        df_cleaned = df.copy()
        
        if phone_columns is None:
            # 自动检测电话列
            phone_columns = [col for col in df_cleaned.columns 
                           if any(keyword in str(col).lower() 
                                for keyword in ['电话', '手机', 'phone', 'tel', '联系'])]
        
        for col in phone_columns:
            if col in df_cleaned.columns:
                # 移除非数字字符
                df_cleaned[col] = df_cleaned[col].astype(str).apply(
                    lambda x: re.sub(r'[^\d]', '', x)
                )
                self.cleaning_report.append(f"电话格式化: {col}")
                print(f"✅ 电话格式化: {col}")
        
        return df_cleaned
    
    def clean_whitespace(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        清理字符串中的多余空格
        
        Args:
            df: 输入DataFrame
            
        Returns:
            清洗后的DataFrame
        """
        df_cleaned = df.copy()
        
        for col in df_cleaned.select_dtypes(include=['object']).columns:
            df_cleaned[col] = df_cleaned[col].astype(str).str.strip()
            df_cleaned[col] = df_cleaned[col].str.replace(r'\s+', ' ', regex=True)
        
        print(f"✅ 清理空格: {len(df_cleaned.select_dtypes(include=['object']).columns)}列")
        return df_cleaned
    
    def fill_missing_values(self, df: pd.DataFrame, 
                           fill_strategy: Dict[str, Any] = None) -> pd.DataFrame:
        """
        填充缺失值
        
        Args:
            df: 输入DataFrame
            fill_strategy: 填充策略字典 {'列名': '填充值或策略'}
                         策略: 'mean', 'median', 'mode', 'forward_fill', 'backward_fill', 或具体值
            
        Returns:
            填充后的DataFrame
        """
        df_cleaned = df.copy()
        
        if fill_strategy is None:
            fill_strategy = {}
        
        for col, strategy in fill_strategy.items():
            if col not in df_cleaned.columns:
                continue
            
            if strategy == 'mean':
                df_cleaned[col] = df_cleaned[col].fillna(df_cleaned[col].mean())
            elif strategy == 'median':
                df_cleaned[col] = df_cleaned[col].fillna(df_cleaned[col].median())
            elif strategy == 'mode':
                df_cleaned[col] = df_cleaned[col].fillna(df_cleaned[col].mode()[0])
            elif strategy == 'forward_fill':
                df_cleaned[col] = df_cleaned[col].ffill()
            elif strategy == 'backward_fill':
                df_cleaned[col] = df_cleaned[col].bfill()
            else:
                df_cleaned[col] = df_cleaned[col].fillna(strategy)
            
            print(f"✅ 填充缺失值: {col} → {strategy}")
        
        return df_cleaned
    
    def apply_all_cleaning(self, df: pd.DataFrame, 
                          remove_dup: bool = True,
                          remove_empty: bool = True,
                          clean_space: bool = True,
                          standardize_date: bool = True,
                          standardize_phone: bool = True) -> pd.DataFrame:
        """
        应用所有清洗步骤
        
        Args:
            df: 输入DataFrame
            remove_dup: 是否去重
            remove_empty: 是否删除空行
            clean_space: 是否清理空格
            standardize_date: 是否标准化日期
            standardize_phone: 是否标准化电话
            
        Returns:
            清洗后的DataFrame
        """
        print(f"\n开始数据清洗（原始数据: {len(df)}行）")
        print("=" * 50)
        
        df_cleaned = df.copy()
        
        # 1. 清理空格
        if clean_space:
            df_cleaned = self.clean_whitespace(df_cleaned)
        
        # 2. 去重
        if remove_dup:
            df_cleaned = self.remove_duplicates(df_cleaned)
        
        # 3. 删除空行
        if remove_empty:
            df_cleaned = self.remove_empty_rows(df_cleaned)
        
        # 4. 标准化日期
        if standardize_date:
            df_cleaned = self.standardize_dates(df_cleaned)
        
        # 5. 标准化电话
        if standardize_phone:
            df_cleaned = self.standardize_phone(df_cleaned)
        
        print("=" * 50)
        print(f"清洗完成: {len(df)}行 → {len(df_cleaned)}行")
        
        return df_cleaned
    
    def save_excel(self, df: pd.DataFrame, 
                   output_path: str,
                   sheet_name: str = '清洗后数据') -> None:
        """
        保存清洗后的数据
        
        Args:
            df: DataFrame
            output_path: 输出文件路径
            sheet_name: 工作表名称
        """
        try:
            # 创建写入器
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # 保存清洗后的数据
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # 如果有清洗报告，保存到另一个工作表
                if self.cleaning_report:
                    report_df = pd.DataFrame({'清洗步骤': self.cleaning_report})
                    report_df.to_excel(writer, sheet_name='清洗报告', index=False)
            
            print(f"✅ 保存成功: {output_path}")
        
        except Exception as e:
            print(f"❌ 保存失败: {e}")


def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Excel数据清洗工具')
    parser.add_argument('input', help='输入Excel文件')
    parser.add_argument('-o', '--output', help='输出Excel文件（默认：添加_cleaned后缀）')
    parser.add_argument('--sheet', help='指定工作表名称')
    parser.add_argument('--no-dup', action='store_false', help='不去重')
    parser.add_argument('--no-empty', action='store_false', help='不删除空行')
    parser.add_argument('--no-date', action='store_false', help='不标准化日期')
    parser.add_argument('--no-phone', action='store_false', help='不标准化电话')
    
    args = parser.parse_args()
    
    # 创建清洗器
    cleaner = ExcelCleaner()
    
    # 加载数据
    df = cleaner.load_excel(args.input, args.sheet)
    
    if df is None:
        return
    
    # 执行清洗
    df_cleaned = cleaner.apply_all_cleaning(
        df,
        remove_dup=args.no_dup,
        remove_empty=args.no_empty,
        standardize_date=args.no_date,
        standardize_phone=args.no_phone
    )
    
    # 保存结果
    if args.output:
        output_path = args.output
    else:
        input_path = Path(args.input)
        output_path = input_path.parent / f"{input_path.stem}_cleaned{input_path.suffix}"
    
    cleaner.save_excel(df_cleaned, str(output_path))


if __name__ == '__main__':
    main()
