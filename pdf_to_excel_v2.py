#!/usr/bin/env python3
"""
PDF发票批量处理 v2 - 增强版
支持：pdfplumber + pymupdf 双引擎，自动选择最优结果
"""

import pdfplumber
import pymupdf
import pandas as pd
import os
import time
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass, field


@dataclass
class ExtractionResult:
    """提取结果"""
    tables: List[pd.DataFrame] = field(default_factory=list)
    engine: str = ""
    page_count: int = 0
    total_rows: int = 0
    processing_time: float = 0.0
    errors: List[str] = field(default_factory=list)


class PDFProcessorV2:
    """增强版PDF处理器 - 双引擎"""

    def __init__(self):
        self.stats = {"processed": 0, "failed": 0, "total_pages": 0, "total_rows": 0}

    def _extract_with_pdfplumber(self, pdf_path: str) -> List[Tuple[pd.DataFrame, int]]:
        """pdfplumber引擎提取"""
        results = []
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                page_tables = page.extract_tables()
                if page_tables:
                    for table in page_tables:
                        if table and len(table) > 1:
                            # 清理None值
                            cleaned = [[c if c else "" for c in row] for row in table]
                            df = pd.DataFrame(cleaned[1:], columns=cleaned[0])
                            df = df.dropna(how="all")
                            if len(df) > 0:
                                results.append((df, page_num))
        return results

    def _extract_with_pymupdf(self, pdf_path: str) -> List[Tuple[pd.DataFrame, int]]:
        """pymupdf引擎提取"""
        results = []
        doc = pymupdf.open(pdf_path)
        for page_num, page in enumerate(doc, 1):
            tabs = page.find_tables()
            for tab in tabs:
                df = tab.to_pandas()
                # 清理空行空列
                df = df.dropna(how="all").dropna(axis=1, how="all")
                if len(df) > 0:
                    results.append((df, page_num))
        doc.close()
        return results

    def extract_tables(self, pdf_path: str) -> ExtractionResult:
        """
        双引擎提取，自动选择最优结果
        """
        result = ExtractionResult()
        start = time.time()

        try:
            doc = pymupdf.open(pdf_path)
            result.page_count = len(doc)
            doc.close()
        except Exception as e:
            result.errors.append(f"无法打开PDF: {e}")
            return result

        # 两个引擎都跑
        plumber_results = []
        mupdf_results = []

        try:
            plumber_results = self._extract_with_pdfplumber(pdf_path)
        except Exception as e:
            result.errors.append(f"pdfplumber失败: {e}")

        try:
            mupdf_results = self._extract_with_pymupdf(pdf_path)
        except Exception as e:
            result.errors.append(f"pymupdf失败: {e}")

        # 选择提取行数更多的引擎结果
        plumber_rows = sum(len(df) for df, _ in plumber_results)
        mupdf_rows = sum(len(df) for df, _ in mupdf_results)

        if plumber_rows >= mupdf_rows and plumber_results:
            chosen = plumber_results
            result.engine = "pdfplumber"
        elif mupdf_results:
            chosen = mupdf_results
            result.engine = "pymupdf"
        elif plumber_results:
            chosen = plumber_results
            result.engine = "pdfplumber"
        else:
            chosen = []
            result.engine = "none"

        for df, page_num in chosen:
            df["_页码"] = page_num
            result.tables.append(df)
            result.total_rows += len(df)

        result.processing_time = time.time() - start
        return result

    def process_single(self, pdf_path: str, output_path: str = None) -> Optional[str]:
        """处理单个PDF"""
        if output_path is None:
            output_path = str(Path(pdf_path).with_suffix(".xlsx"))

        result = self.extract_tables(pdf_path)

        if not result.tables:
            print(f"⚠️  {Path(pdf_path).name}: 未提取到表格")
            self.stats["failed"] += 1
            return None

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            # 合并所有表格到一个sheet
            all_data = pd.concat(result.tables, ignore_index=True)
            all_data.to_excel(writer, sheet_name="全部数据", index=False)

            # 每个表格单独一个sheet
            for i, df in enumerate(result.tables, 1):
                df.to_excel(writer, sheet_name=f"表格{i}", index=False)

        self.stats["processed"] += 1
        self.stats["total_pages"] += result.page_count
        self.stats["total_rows"] += result.total_rows

        print(f"✅ {Path(pdf_path).name}: {result.page_count}页, "
              f"{len(result.tables)}个表格, {result.total_rows}行, "
              f"引擎={result.engine}, {result.processing_time:.1f}s")
        return output_path

    def batch_process(self, input_dir: str, output_dir: str = None) -> Dict[str, Any]:
        """批量处理目录下所有PDF"""
        input_path = Path(input_dir)
        if output_dir:
            out_path = Path(output_dir)
            out_path.mkdir(parents=True, exist_ok=True)
        else:
            out_path = input_path

        pdf_files = sorted(input_path.glob("*.pdf")) + sorted(input_path.glob("*.PDF"))
        total = len(pdf_files)
        print(f"找到 {total} 个PDF文件\n")

        results = []
        for i, pdf_file in enumerate(pdf_files, 1):
            print(f"[{i}/{total}] ", end="")
            output_file = out_path / f"{pdf_file.stem}.xlsx"
            r = self.process_single(str(pdf_file), str(output_file))
            if r:
                results.append(r)

        print(f"\n{'='*50}")
        print(f"批量处理完成: {self.stats['processed']}/{total} 成功")
        print(f"总页数: {self.stats['total_pages']}, 总行数: {self.stats['total_rows']}")
        print(f"{'='*50}")

        return self.stats


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="PDF表格提取 v2（双引擎）")
    parser.add_argument("input", help="PDF文件或目录")
    parser.add_argument("-o", "--output", help="输出路径")
    args = parser.parse_args()

    processor = PDFProcessorV2()
    p = Path(args.input)

    if p.is_file():
        processor.process_single(args.input, args.output)
    elif p.is_dir():
        processor.batch_process(args.input, args.output)
    else:
        print(f"❌ 路径不存在: {args.input}")
