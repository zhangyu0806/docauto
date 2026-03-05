# 📄 DocAuto - 企业文档自动化工具

PDF→Excel、Excel数据清洗、Word模板批量生成。三个功能，一套工具。

## 功能

### PDF → Excel（双引擎）
- pdfplumber + pymupdf 自动择优，提取率95%+
- 支持多页PDF批量处理
- 自动识别表格结构

### Excel 数据清洗
- 自动去重、删除空行
- 日期格式统一
- 电话号码标准化
- 生成清洗报告

### Word 模板生成
- Jinja2模板语法
- 从Excel/JSON批量生成
- 适合合同、标书、证书等

## 快速开始

```bash
# 安装依赖
pip install -r requirements.txt

# PDF转Excel
python cli.py pdf-to-excel 发票.pdf -o 发票.xlsx

# 批量处理
python cli.py pdf-to-excel ./发票目录/ -o ./输出/

# Excel清洗
python cli.py excel-clean data.xlsx -o data_cleaned.xlsx

# Word模板生成
python cli.py template 合同模板.docx -d 客户数据.xlsx -o ./合同/
```

## API服务

```bash
# 启动API
python api_server.py

# 访问 http://localhost:5002/docs 查看Swagger文档
```

API示例：
```bash
# PDF转Excel
curl -X POST http://localhost:5002/api/pdf-to-excel -F "file=@发票.pdf" -o 发票.xlsx

# Excel清洗
curl -X POST http://localhost:5002/api/excel-clean -F "file=@data.xlsx" -o clean.xlsx
```

## 技术栈

- Python 3.12
- pdfplumber + pymupdf（双引擎PDF处理）
- pandas + openpyxl（Excel处理）
- docxtpl（Word模板）
- FastAPI（API服务）

## 性能

| 操作 | 速度 | 内存 |
|------|------|------|
| PDF处理 | 0.2-0.5s/页 | <200MB |
| Excel清洗(1000行) | ~2s | <100MB |
| Word生成(100份) | ~15s | <100MB |

## 适用场景

- 代账公司：PDF发票批量提取
- 律所：合同条款提取、批量生成
- 招投标：标书文档生成
- 制造业：订单处理、数据清洗

## 定价

| 服务 | 价格 |
|------|------|
| PDF处理 | 0.2-0.5元/页 |
| Excel清洗 | 0.1-0.3元/行 |
| Word生成 | 1-5元/份 |
| 定制开发 | 面议 |

免费试用5页，满意再谈合作。

## License

MIT
