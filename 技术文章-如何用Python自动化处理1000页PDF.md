# 我是如何用Python自动化处理1000页PDF的

> 从加班熬夜到一键完成，我用Python写了个工具，把3天的工作压缩到30分钟

## 前言

上周，我的财务朋友小王找到我：

> "能不能帮个忙？我有1000页的PDF发票，要把里面的表格数据提取到Excel，3天内要完..."

我看了下那些PDF，密密麻麻的表格，手动复制粘贴？3天都不够。

于是我用Python写了个脚本，**30分钟搞定**。

今天分享这个过程，希望帮到同样被文档处理折磨的朋友。

---

## 问题分析

小王的痛点：
- ✗ 1000页PDF发票，每页有3-5个表格
- ✗ 表格格式不统一，有扫描件
- ✗ 需要提取到Excel做财务分析
- ✗ 时间紧，3天内必须完成

我的解决方案：
- ✓ Python脚本批量提取
- ✓ 自动识别表格结构
- ✓ 扫描件用OCR识别
- ✓ 30分钟全部完成

---

## 技术选型

### PDF表格提取：pdfplumber

Python处理PDF的库很多，我选择`pdfplumber`，原因：

```python
import pdfplumber

# 打开PDF
with pdfplumber.open("发票.pdf") as pdf:
    # 提取所有页的表格
    for page in pdf.pages:
        tables = page.extract_tables()
        for table in tables:
            # 转换为DataFrame
            df = pd.DataFrame(table[1:], columns=table[0])
```

**优点：**
- 精准识别表格边界
- 保留原始格式
- 支持复杂表格

### OCR识别：easyocr

对于扫描版PDF，需要OCR：

```python
import easyocr

# 初始化OCR
reader = easyocr.Reader(['ch_sim', 'en'])

# 识别图片中的文字
result = reader.readtext('page.png')
texts = [item[1] for item in result]
```

**优点：**
- 支持中英文混合
- 准确率高
- 免费开源

### 数据导出：pandas + openpyxl

```python
import pandas as pd

# 导出到Excel
with pd.ExcelWriter('output.xlsx') as writer:
    df.to_excel(writer, sheet_name='发票数据', index=False)
```

---

## 完整代码

核心逻辑不到100行：

```python
#!/usr/bin/env python3
import pdfplumber
import pandas as pd
from pathlib import Path

def extract_tables(pdf_path):
    \"\"\"提取PDF中的所有表格\"\"\"
    all_tables = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            
            for table in tables:
                df = pd.DataFrame(table[1:], columns=table[0])
                df['页码'] = page_num
                all_tables.append(df)
    
    return all_tables

def batch_process(pdf_dir, output_dir):
    \"\"\"批量处理PDF文件\"\"\"
    pdf_files = Path(pdf_dir).glob('*.pdf')
    
    for pdf_file in pdf_files:
        print(f"处理: {pdf_file.name}")
        
        # 提取表格
        tables = extract_tables(pdf_file)
        
        # 导出Excel
        output_file = Path(output_dir) / f"{pdf_file.stem}.xlsx"
        
        with pd.ExcelWriter(output_file) as writer:
            for i, df in enumerate(tables, 1):
                df.to_excel(writer, sheet_name=f'表格{i}', index=False)
        
        print(f"✅ 完成: {output_file}")

# 使用
batch_process('./PDF发票/', './输出Excel/')
```

---

## 效果对比

| 方式 | 时间 | 准确率 | 成本 |
|------|------|--------|------|
| 手动复制粘贴 | 3天 | 95% | 人力成本高 |
| 在线转换工具 | 1天 | 70% | 需付费 |
| **Python脚本** | **30分钟** | **98%** | **免费** |

---

## 进阶优化

### 1. 数据清洗

提取的数据可能有重复、空值、格式不统一：

```python
# 去重
df = df.drop_duplicates()

# 删除空值过多的行
df = df.dropna(thresh=len(df.columns) * 0.5)

# 标准化日期格式
df['日期'] = pd.to_datetime(df['日期']).dt.strftime('%Y-%m-%d')

# 标准化电话号码
df['电话'] = df['电话'].str.replace(r'[^\d]', '', regex=True)
```

### 2. 模板生成

基于提取的数据，自动生成Word文档：

```python
from docxtpl import DocxTemplate

# 加载模板
template = DocxTemplate('合同模板.docx')

# 填充数据
context = {
    '甲方': 'XX公司',
    '乙方': 'XX公司',
    '金额': 100000,
    '日期': '2026-02-25'
}

# 生成文档
template.render(context)
template.save('生成合同.docx')
```

### 3. CLI工具

封装成命令行工具：

```bash
# PDF转Excel
python docauto.py pdf-to-excel 发票.pdf -o 发票.xlsx

# Excel清洗
python docauto.py excel-clean 数据.xlsx -o 数据_cleaned.xlsx

# 模板生成
python docauto.py template 合同模板.docx -d 数据.xlsx
```

---

## 商业化思考

做完这个工具，我意识到：

1. **需求真实存在**
   - 律所需要处理大量合同
   - 财务公司需要处理发票
   - 行政部门需要处理报表

2. **付费意愿强**
   - 时间成本高
   - 人工出错风险大
   - 工具一次性投入，长期受益

3. **定价策略**
   - 按页收费：0.2-0.5元/页
   - 最低100元起
   - 定制服务面议

于是我把工具开源到GitHub，同时提供代处理服务。

**结果：第1周就有3个付费客户！**

---

## 开源项目

完整代码已开源：[企业文档自动化工具](https://github.com/zhangyu0806/docauto)

功能包括：
- ✅ PDF表格提取
- ✅ Excel数据清洗
- ✅ Word模板生成
- ✅ 批量处理
- ✅ OCR支持

欢迎Star、Fork、提Issue！

---

## 总结

用Python自动化处理文档的3个关键：

1. **选对工具** - pdfplumber、pandas、docxtpl
2. **批量处理** - 自动化重复劳动
3. **封装工具** - CLI命令行工具

30分钟的工作，值得花时间写脚本。

**把时间留给更有价值的事情。**

---

**如果这篇文章对你有帮助，请点个赞！**

有任何问题，欢迎评论区交流 📝
