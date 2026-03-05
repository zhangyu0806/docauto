#!/bin/bash
# 企业文档自动化 - 快速演示脚本
# 用于向客户展示工具能力

echo "================================"
echo "📄 企业文档自动化 - 演示"
echo "================================"
echo ""

# 进入项目目录
cd "$(dirname "$0")"

# 检查虚拟环境
if [ ! -d "../venv" ]; then
    echo "❌ 虚拟环境不存在，请先运行: python3 -m venv ../venv"
    exit 1
fi

PYTHON="../venv/bin/python3"

echo "🔍 检查依赖..."
$PYTHON -c "import pymupdf, pdfplumber, pandas" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "❌ 缺少依赖，请先安装: pip install -r requirements.txt"
    exit 1
fi
echo "✅ 依赖完整"
echo ""

echo "================================"
echo "演示1: PDF发票 → Excel"
echo "================================"
echo "📂 输入: test_data/test_invoice.pdf (1页发票)"
echo "📂 输出: test_results/demo_invoice.xlsx"
echo ""
$PYTHON pdf_to_excel_v2.py test_data/test_invoice.pdf -o test_results/demo_invoice.xlsx
echo ""

echo "================================"
echo "演示2: Excel数据清洗"
echo "================================"
echo "📂 输入: test_data/test_dirty_data.xlsx (15行，含重复/空行)"
echo "📂 输出: test_results/demo_cleaned.xlsx"
echo ""
$PYTHON cli.py excel-clean test_data/test_dirty_data.xlsx -o test_results/demo_cleaned.xlsx
echo ""

echo "================================"
echo "演示3: Word模板批量生成"
echo "================================"
echo "📂 模板: test_data/test_contract_template.docx"
echo "📂 数据: test_data/test_template_data.json"
echo "📂 输出: test_results/demo_contracts/"
echo ""
$PYTHON cli.py template test_data/test_contract_template.docx \
    -d test_data/test_template_data.json \
    -o test_results/demo_contracts/ \
    --json
echo ""

echo "================================"
echo "✅ 所有演示完成！"
echo "================================"
echo ""
echo "📊 查看结果:"
echo "  - test_results/ 目录"
echo ""
echo "🚀 开始使用:"
echo "  - 单个PDF: $PYTHON pdf_to_excel_v2.py your.pdf"
echo "  - 批量PDF: $PYTHON pdf_to_excel_v2.py ./pdf_dir/"
echo "  - 数据清洗: $PYTHON cli.py excel-clean data.xlsx"
echo "  - 模板生成: $PYTHON cli.py template template.docx -d data.xlsx"
echo ""
echo "💼 商务报价:"
echo "  - PDF处理: 0.2-0.5元/页 (最低100元起)"
echo "  - 数据清洗: 100元起"
echo "  - 模板生成: 200元起"
echo ""
