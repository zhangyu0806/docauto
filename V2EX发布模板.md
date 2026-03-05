# V2EX发布模板

## 版块：分享发现

## 标题
[开源] 写了个企业文档自动化工具，把3天的工作压缩到30分钟

## 正文

### 背景

上周财务朋友找我帮忙：
- 1000页PDF发票需要提取表格到Excel
- 3天内必须完成
- 手动复制粘贴要3天

我用Python写了个脚本，30分钟搞定。

### 功能

**1. PDF → Excel**
- 批量提取PDF表格
- 支持扫描版（OCR）
- 自动导出Excel

**2. Excel数据清洗**
- 去重
- 删除空行
- 标准化日期/电话格式

**3. Word模板生成**
- 批量生成合同
- 批量生成报告
- 自动填充变量

### 使用

```bash
# PDF转Excel
python cli.py pdf-to-excel 发票.pdf -o 发票.xlsx

# Excel清洗
python cli.py excel-clean 数据.xlsx -o 数据_cleaned.xlsx

# 模板生成
python cli.py template 合同模板.docx -d 数据.xlsx
```

### 效果

| 方式 | 时间 | 准确率 |
|------|------|--------|
| 手动 | 3天 | 95% |
| **Python工具** | **30分钟** | **98%** |

### 开源

GitHub: https://github.com/zhangyu0806/docauto

欢迎Star、Fork、提Issue！

### 适用场景

- **财务**：发票批量处理、账单整理
- **律所**：合同模板生成、案卷整理
- **行政**：报表自动生成、文档批量处理

### 定价服务

如果你不想自己折腾，我也提供代处理服务：
- PDF处理：0.2-0.5元/页
- 数据清洗：100元起
- 模板生成：200元起

### 技术栈

- pdfplumber（PDF表格提取）
- pandas（Excel处理）
- docxtpl（Word模板）
- easyocr（OCR）

---

希望对大家有帮助！

有任何问题欢迎交流。

---

## 标签
#Python #开源 #办公效率 #自动化 #PDF
