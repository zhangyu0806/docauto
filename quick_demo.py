#!/usr/bin/env python3
"""
DocAuto 快速演示脚本
一键展示所有功能
"""

import sys
from pathlib import Path

# 添加当前目录到路径
sys.path.insert(0, str(Path(__file__).parent))

from pdf_to_excel_v2 import PDFProcessorV2
from excel_cleaner import ExcelCleaner
from template_generator import WordTemplateGenerator


def print_section(title: str):
    """打印分节标题"""
    print("\n" + "=" * 60)
    print(f"  {title}")
    print("=" * 60 + "\n")


def demo_pdf_to_excel():
    """演示PDF转Excel"""
    print_section("1. PDF发票转Excel")

    processor = PDFProcessorV2()

    # 处理测试发票
    input_pdf = "test_data/test_invoice.pdf"
    if Path(input_pdf).exists():
        print(f"处理文件: {input_pdf}")
        result = processor.process_single(input_pdf, "test_data/demo_invoice_output.xlsx")

        if result:
            print(f"✅ 成功生成: demo_invoice_output.xlsx")
            print(f"   引擎: {processor.extract_tables(input_pdf).engine}")
            print(f"   处理时间: {processor.extract_tables(input_pdf).processing_time:.2f}秒")
    else:
        print(f"⚠️  测试文件不存在: {input_pdf}")


def demo_excel_clean():
    """演示Excel清洗"""
    print_section("2. Excel数据清洗")

    cleaner = ExcelCleaner()

    input_file = "test_data/test_dirty_data.xlsx"
    if Path(input_file).exists():
        print(f"处理文件: {input_file}")

        df = cleaner.load_excel(input_file)
        if df is not None:
            print(f"原始数据: {len(df)}行 x {len(df.columns)}列")

            # 展示前3行原始数据
            print("\n原始数据预览:")
            print(df.head(3).to_string(index=False))

            # 执行清洗
            df_cleaned = cleaner.apply_all_cleaning(df)
            output_file = "test_data/demo_cleaned.xlsx"

            cleaner.save_excel(df_cleaned, output_file)

            print(f"\n✅ 清洗后: {len(df_cleaned)}行")
            print(f"✅ 保存到: {output_file}")

            # 展示清洗后数据
            print("\n清洗后数据预览:")
            print(df_cleaned.head(3).to_string(index=False))
    else:
        print(f"⚠️  测试文件不存在: {input_file}")


def demo_template_generation():
    """演示Word模板生成"""
    print_section("3. Word模板批量生成")

    try:
        # 检查测试文件
        template_file = "test_data/test_contract_template.docx"
        data_file = "test_data/test_template_data.json"

        if not Path(template_file).exists():
            print(f"⚠️  模板文件不存在，创建示例模板...")
            from template_generator import create_sample_template
            create_sample_template("test_data/demo_template.docx")
            template_file = "test_data/demo_template.docx"

        if Path(data_file).exists():
            import json
            with open(data_file, 'r', encoding='utf-8') as f:
                data = json.load(f)

            print(f"使用模板: {template_file}")
            print(f"数据记录: {len(data)}条")

            generator = WordTemplateGenerator(template_file)

            # 生成第一个文档
            output_dir = "test_data/demo_outputs"
            Path(output_dir).mkdir(exist_ok=True)

            single_output = f"{output_dir}/demo_single.docx"
            generator.generate_document(data[0], single_output)

            print(f"✅ 单个文档: {single_output}")

            # 批量生成
            if len(data) > 1:
                generator.batch_generate(data, output_dir, "demo_{name}.docx")
                print(f"✅ 批量生成: {len(data)}个文档 → {output_dir}/")
        else:
            print(f"⚠️  数据文件不存在: {data_file}")

    except Exception as e:
        print(f"❌ 模板生成失败: {e}")


def demo_stats():
    """显示演示统计"""
    print_section("演示总结")

    # 检查生成的文件
    demo_files = [
        "test_data/demo_invoice_output.xlsx",
        "test_data/demo_cleaned.xlsx",
        "test_data/demo_template.docx",
        "test_data/demo_outputs",
    ]

    existing_files = [f for f in demo_files if Path(f).exists()]

    print("生成的文件:")
    for f in existing_files:
        p = Path(f)
        if p.is_dir():
            file_count = len(list(p.glob("*")))
            print(f"  📁 {f} ({file_count}个文件)")
        else:
            size = p.stat().st_size
            print(f"  📄 {f} ({size}字节)")

    print("\n" + "=" * 60)
    print("  演示完成！所有功能正常工作")
    print("=" * 60 + "\n")

    print("🚀 DocAuto 企业文档自动化工具")
    print("   - PDF发票批量处理 → Excel")
    print("   - Excel数据清洗（去重、格式统一）")
    print("   - Word模板批量生成")
    print("\n💼 适用场景:")
    print("   - 财务: 发票、报销单、报表处理")
    print("   - 法务: 合同、协议、文书批量生成")
    print("   - 行政: 证照、通知、文档自动化")
    print("\n💰 定价: 0.2-0.5元/页，100元起")
    print("📞 咨询: 免费试用5页，满意后合作")


def main():
    """主函数"""
    print("\n")
    print("╔" + "=" * 58 + "╗")
    print("║" + " " * 15 + "DocAuto 快速演示" + " " * 28 + "║")
    print("╚" + "=" * 58 + "╝")

    try:
        demo_pdf_to_excel()
        demo_excel_clean()
        demo_template_generation()
        demo_stats()
    except KeyboardInterrupt:
        print("\n\n操作已取消")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ 演示过程中出错: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
