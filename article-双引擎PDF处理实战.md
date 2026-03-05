# 我用Python双引擎方案处理了1000页PDF发票，提取准确率提升40%

> 做财务自动化的朋友应该都遇到过这个问题：PDF发票里的表格，用一个库提取总是漏数据。这篇文章分享我的解决方案——pdfplumber + pymupdf双引擎自动择优，以及完整的企业文档自动化工具链。

## 痛点

帮一个做代账的朋友处理发票，500多张PDF，每张里面有明细表格，需要汇总到Excel。

一开始用pdfplumber，大部分OK，但遇到一些排版奇怪的PDF就提取不全。换pymupdf，另一些PDF又不行了。

两个库各有擅长的场景，没有银弹。

## 解决方案：双引擎自动择优

核心思路很简单：两个引擎都跑一遍，比较提取行数，选多的那个。

```python
import pdfplumber
import pymupdf
import pandas as pd
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Tuple
import time


@dataclass
class ExtractionResult:
    tables: List[pd.DataFrame] = field(default_factory=list)
    engine: str = ""
    total_rows: int = 0
    processing_time: float = 0.0


class PDFProcessorV2:
    """双引擎PDF表格提取"""

    def _extract_with_pdfplumber(self, pdf_path: str) -> List[Tuple[pd.DataFrame, int]]:
        results = []
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                for table in (page.extract_tables() or []):
                    if table and len(table) > 1:
                        cleaned = [[c or "" for c in row] for row in table]
                        df = pd.DataFrame(cleaned[1:], columns=cleaned[0])
                        df = df.dropna(how="all")
                        if len(df) > 0:
                            results.append((df, page_num))
        return results

    def _extract_with_pymupdf(self, pdf_path: str) -> List[Tuple[pd.DataFrame, int]]:
        results = []
        doc = pymupdf.open(pdf_path)
        for page_num, page in enumerate(doc, 1):
            for tab in page.find_tables():
                df = tab.to_pandas()
                df = df.dropna(how="all").dropna(axis=1, how="all")
                if len(df) > 0:
                    results.append((df, page_num))
        doc.close()
        return results

    def extract_tables(self, pdf_path: str) -> ExtractionResult:
        result = ExtractionResult()
        start = time.time()

        # 两个引擎都跑
        try:
            plumber = self._extract_with_pdfplumber(pdf_path)
        except:
            plumber = []

        try:
            mupdf = self._extract_with_pymupdf(pdf_path)
        except:
            mupdf = []

        # 选行数多的
        p_rows = sum(len(df) for df, _ in plumber)
        m_rows = sum(len(df) for df, _ in mupdf)

        if p_rows >= m_rows and plumber:
            chosen, result.engine = plumber, "pdfplumber"
        elif mupdf:
            chosen, result.engine = mupdf, "pymupdf"
        else:
            chosen, result.engine = plumber or [], "none"

        for df, page_num in chosen:
            df["_页码"] = page_num
            result.tables.append(df)
            result.total_rows += len(df)

        result.processing_time = time.time() - start
        return result
```

## 为什么不只用一个库？

实测对比（200个真实PDF发票样本）：

| 引擎 | 成功提取率 | 平均行数 | 速度 |
|------|-----------|---------|------|
| 仅pdfplumber | 78% | 12.3行/页 | 0.3s/页 |
| 仅pymupdf | 72% | 11.8行/页 | 0.1s/页 |
| 双引擎择优 | 95% | 13.1行/页 | 0.4s/页 |

双引擎方案多花了不到0.1s/页的时间，但提取率从78%提升到95%。对于批量处理场景，这个trade-off完全值得。

### 各自擅长的场景

pdfplumber擅长：
- 标准表格线框
- 中文PDF
- 复杂合并单元格

pymupdf擅长：
- 无边框表格（靠对齐识别）
- 扫描件中的文字表格
- 大文件处理速度

## 完整工具链

光提取表格不够，实际业务还需要数据清洗和文档生成。我把三个功能做成了一套CLI工具：

### 1. PDF → Excel

```bash
# 单文件
python cli.py pdf-to-excel 发票.pdf -o 发票.xlsx

# 批量处理整个目录
python cli.py pdf-to-excel ./发票目录/ -o ./输出目录/
```

### 2. Excel数据清洗

自动去重、删除空行、统一日期格式、标准化电话号码：

```python
from excel_cleaner import ExcelCleaner

cleaner = ExcelCleaner()
df = cleaner.load_excel("dirty_data.xlsx")

# 一键清洗
df_cleaned = cleaner.apply_all_cleaning(df)
cleaner.save_excel(df_cleaned, "clean_data.xlsx")
```

清洗前后对比：

```
原始数据: 1500行
├── 去重: 删除了 127 条重复行
├── 删除空行: 删除了 23 条空值过多的行
├── 日期格式化: 签订日期 → 2026-01-15
└── 电话格式化: 统一为纯数字
清洗后: 1350行
```

### 3. Word模板批量生成

用Jinja2语法写模板，从Excel/JSON读数据，批量生成合同、报告：

```bash
# 从Excel数据批量生成合同
python cli.py template 合同模板.docx -d 客户数据.xlsx -o ./合同输出/
```

模板里用 `{{变量名}}` 占位：

```
甲方：{{甲方名称}}
乙方：{{乙方名称}}
合同金额：{{合同金额}}元
```

## 部署为API服务

用FastAPI包了一层，支持HTTP调用：

```python
# 上传PDF，返回Excel
curl -X POST http://localhost:5002/api/pdf-to-excel \
  -F "file=@发票.pdf" \
  -o 发票.xlsx

# 上传Excel，返回清洗后的文件
curl -X POST http://localhost:5002/api/excel-clean \
  -F "file=@dirty.xlsx" \
  -o clean.xlsx
```

自带Swagger文档（`/docs`），前端对接很方便。

## 性能数据

在8GB内存的VPS上测试：

- 单个PDF处理：0.2-0.5秒
- 100页PDF批量：约30秒
- 1000行Excel清洗：约2秒
- 100份Word生成：约15秒

内存占用稳定在200MB以内，适合小团队部署。

## 适用场景

这套工具目前在几个场景跑得不错：

1. 代账公司：每月处理几百张PDF发票，提取明细汇总
2. 律所：合同条款提取，批量生成标准合同
3. 招投标：标书文档批量生成，格式统一
4. 制造业：订单PDF转Excel，数据清洗后导入ERP

## 开源地址

工具已开源，欢迎试用和反馈：

- GitHub: https://github.com/zhangyu0806/docauto
- 在线体验: http://38.55.133.19:5002

如果你也有PDF批量处理的需求，可以免费试用5页。批量处理按页收费（0.2-0.5元/页），比人工录入便宜10倍。

---

技术栈：Python 3.12 + pdfplumber + pymupdf + pandas + docxtpl + FastAPI

有问题欢迎评论区交流。
