#!/usr/bin/env python3
"""
企业文档自动化工具 - 统一CLI
功能：PDF→Excel、数据清洗、模板生成、OFD发票处理
"""

import argparse
import sys
from pathlib import Path

# 添加当前目录到 Python 路径
sys.path.insert(0, str(Path(__file__).parent))

from pdf_to_excel_v2 import PDFProcessorV2
from excel_cleaner import ExcelCleaner
from template_generator import WordTemplateGenerator


def cmd_pdf_to_excel(args):
    """PDF转Excel命令"""
    print("=" * 60)
    print("PDF发票处理 → Excel（双引擎v2）")
    print("=" * 60)
    
    processor = PDFProcessorV2()
    
    input_path = Path(args.input)
    
    if input_path.is_file():
        processor.process_single(str(input_path), args.output)
    elif input_path.is_dir():
        processor.batch_process(str(input_path), args.output)
    else:
        print(f"❌ 输入路径不存在: {args.input}")
        sys.exit(1)


def cmd_excel_clean(args):
    """Excel清洗命令"""
    print("=" * 60)
    print("Excel数据清洗")
    print("=" * 60)
    
    cleaner = ExcelCleaner()
    df = cleaner.load_excel(args.input, args.sheet)
    
    if df is None:
        sys.exit(1)
    
    df_cleaned = cleaner.apply_all_cleaning(
        df,
        remove_dup=args.remove_dup,
        remove_empty=args.remove_empty,
        standardize_date=args.standardize_date,
        standardize_phone=args.standardize_phone
    )
    
    if args.output:
        output_path = args.output
    else:
        input_path = Path(args.input)
        output_path = input_path.parent / f"{input_path.stem}_cleaned{input_path.suffix}"
    
    cleaner.save_excel(df_cleaned, str(output_path))


def cmd_template_generate(args):
    """模板生成命令"""
    print("=" * 60)
    print("Word模板自动填充")
    print("=" * 60)
    
    if args.create_sample:
        from template_generator import create_sample_template
        create_sample_template(args.template)
        return
    
    try:
        generator = WordTemplateGenerator(args.template)
    except Exception as e:
        print(f"❌ {e}")
        sys.exit(1)
    
    variables = generator.preview_template_variables()
    print(f"\n模板变量: {', '.join(variables)}\n")
    
    if args.data:
        data_path = Path(args.data)
        
        if data_path.suffix == '.json':
            import json
            with open(data_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if isinstance(data, list):
                generator.batch_generate(data, args.output, args.pattern)
            else:
                generator.generate_document(data, args.output)
        
        elif data_path.suffix in ['.xlsx', '.xls']:
            generator.generate_from_excel(
                args.data, 
                args.sheet, 
                args.output, 
                args.pattern
            )
    else:
        print("⚠️  请提供数据文件（-d/--data）")


def main():
    """主入口"""
    parser = argparse.ArgumentParser(
        description='企业文档自动化工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
  python cli.py pdf-to-excel 发票.pdf -o 发票.xlsx
  python cli.py excel-clean data.xlsx -o data_cleaned.xlsx
  python cli.py template 合同模板.docx -d 数据.xlsx
        """
    )
    
    subparsers = parser.add_subparsers(dest='command', help='可用命令')
    
    # PDF转Excel
    pdf_parser = subparsers.add_parser('pdf-to-excel', help='PDF发票批量处理')
    pdf_parser.add_argument('input', help='输入PDF文件或目录')
    pdf_parser.add_argument('-o', '--output', help='输出Excel文件或目录')
    pdf_parser.add_argument('--ocr', action='store_true', help='启用OCR')
    pdf_parser.set_defaults(func=cmd_pdf_to_excel)
    
    # Excel清洗
    clean_parser = subparsers.add_parser('excel-clean', help='Excel数据清洗')
    clean_parser.add_argument('input', help='输入Excel文件')
    clean_parser.add_argument('-o', '--output', help='输出Excel文件')
    clean_parser.add_argument('--sheet', help='指定工作表名称')
    clean_parser.add_argument('--no-dup', action='store_false', dest='remove_dup')
    clean_parser.add_argument('--no-empty', action='store_false', dest='remove_empty')
    clean_parser.add_argument('--no-date', action='store_false', dest='standardize_date')
    clean_parser.add_argument('--no-phone', action='store_false', dest='standardize_phone')
    clean_parser.set_defaults(func=cmd_excel_clean)
    
    # 模板生成
    template_parser = subparsers.add_parser('template', help='Word模板自动填充')
    template_parser.add_argument('template', help='Word模板文件路径')
    template_parser.add_argument('-d', '--data', help='数据文件（JSON或Excel）')
    template_parser.add_argument('-o', '--output', help='输出目录或文件路径')
    template_parser.add_argument('--sheet', help='Excel工作表名称')
    template_parser.add_argument('--pattern', default='{序号}_{名称}.docx')
    template_parser.add_argument('--create-sample', action='store_true')
    template_parser.set_defaults(func=cmd_template_generate)
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        sys.exit(1)
    
    try:
        args.func(args)
    except KeyboardInterrupt:
        print("\n\n操作已取消")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
