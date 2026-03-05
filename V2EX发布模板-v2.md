# V2EX发帖 - 分享帖

## 标题
分享一个PDF批量处理工具，双引擎自动择优，提取率95%+

## 正文

做了一个企业文档自动化工具，主要解决三个问题：

1. **PDF → Excel**：双引擎（pdfplumber + pymupdf）自动择优，比单引擎提取率高约20%
2. **Excel数据清洗**：自动去重、格式统一、日期标准化
3. **Word模板生成**：Jinja2语法，从Excel/JSON批量生成合同等文档

CLI + FastAPI双模式，可以命令行用也可以当API调。

```bash
# PDF转Excel
python cli.py pdf-to-excel 发票.pdf -o 发票.xlsx

# 批量处理
python cli.py pdf-to-excel ./发票目录/ -o ./输出/
```

双引擎的核心思路：两个库各有擅长场景，都跑一遍，比较提取行数，选多的。多花0.1s/页，提取率从78%到95%。

技术栈：Python 3.12 + pdfplumber + pymupdf + pandas + docxtpl + FastAPI

GitHub: https://github.com/zhangyu0806/docauto

如果你也有PDF批量处理的需求，可以免费试5页。

---

## 发布节点
- /t/python
- /t/share

## 注意事项
- V2EX不喜欢硬广，重点放在技术分享
- 定价信息不要放正文，有人问再说
- 保持简洁，V2EX用户不喜欢长文
