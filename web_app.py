#!/usr/bin/env python3
"""
企业文档自动化工具 - Flask Web应用
提供简单的Web界面用于PDF处理、Excel清洗、Word模板生成
"""

from flask import Flask, render_template, request, send_file, jsonify
import os
import uuid
from pathlib import Path
import traceback

# 导入核心模块
from pdf_to_excel import PDFInvoiceProcessor
from excel_cleaner import ExcelCleaner
from template_generator import WordTemplateGenerator

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
app.config['UPLOAD_FOLDER'] = Path(__file__).parent / 'temp_uploads'
app.config['OUTPUT_FOLDER'] = Path(__file__).parent / 'temp_outputs'
app.config['RESULT_FOLDER'] = Path(__file__).parent / 'test_results'

# 确保目录存在
for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER'], app.config['RESULT_FOLDER']]:
    folder.mkdir(exist_ok=True)


@app.route('/')
def index():
    """首页"""
    return render_template('index.html')


@app.route('/pdf-to-excel', methods=['GET', 'POST'])
def pdf_to_excel():
    """PDF转Excel页面"""
    if request.method == 'GET':
        return render_template('pdf_to_excel.html')
    
    try:
        if 'file' not in request.files:
            return jsonify({'error': '请上传PDF文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '请选择文件'}), 400
        
        if not file.filename.lower().endswith('.pdf'):
            return jsonify({'error': '只支持PDF文件'}), 400
        
        # 保存上传文件
        file_id = str(uuid.uuid4())[:8]
        upload_path = app.config['UPLOAD_FOLDER'] / f"{file_id}_{file.filename}"
        file.save(str(upload_path))
        
        # 处理PDF
        use_ocr = request.form.get('use_ocr') == 'true'
        processor = PDFInvoiceProcessor(use_ocr=use_ocr)
        
        output_path = app.config['OUTPUT_FOLDER'] / f"{file_id}_output.xlsx"
        result = processor.process_invoice(str(upload_path), str(output_path))
        
        if result:
            # 返回结果
            output_filename = f"{Path(file.filename).stem}.xlsx"
            return send_file(str(output_path), 
                           as_attachment=True,
                           download_name=output_filename,
                           mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            return jsonify({'error': 'PDF处理失败，请确认文件包含表格'}), 400
    
    except Exception as e:
        app.logger.error(f"PDF处理错误: {traceback.format_exc()}")
        return jsonify({'error': f'处理失败: {str(e)}'}), 500


@app.route('/excel-clean', methods=['GET', 'POST'])
def excel_clean():
    """Excel清洗页面"""
    if request.method == 'GET':
        return render_template('excel_clean.html')
    
    try:
        if 'file' not in request.files:
            return jsonify({'error': '请上传Excel文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '请选择文件'}), 400
        
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            return jsonify({'error': '只支持Excel文件'}), 400
        
        # 保存上传文件
        file_id = str(uuid.uuid4())[:8]
        upload_path = app.config['UPLOAD_FOLDER'] / f"{file_id}_{file.filename}"
        file.save(str(upload_path))
        
        # 获取清洗选项
        options = {
            'remove_dup': request.form.get('remove_dup') != 'false',
            'remove_empty': request.form.get('remove_empty') != 'false',
            'standardize_date': request.form.get('standardize_date') != 'false',
            'standardize_phone': request.form.get('standardize_phone') != 'false',
        }
        
        # 清洗数据
        cleaner = ExcelCleaner()
        df = cleaner.load_excel(str(upload_path))
        
        if df is None:
            return jsonify({'error': 'Excel文件读取失败'}), 400
        
        df_cleaned = cleaner.apply_all_cleaning(df, **options)
        
        # 保存结果
        output_path = app.config['OUTPUT_FOLDER'] / f"{file_id}_cleaned.xlsx"
        cleaner.save_excel(df_cleaned, str(output_path))
        
        # 返回结果
        output_filename = f"{Path(file.filename).stem}_cleaned.xlsx"
        return send_file(str(output_path),
                       as_attachment=True,
                       download_name=output_filename,
                       mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    
    except Exception as e:
        app.logger.error(f"Excel清洗错误: {traceback.format_exc()}")
        return jsonify({'error': f'处理失败: {str(e)}'}), 500


@app.route('/template-generate', methods=['GET', 'POST'])
def template_generate():
    """模板生成页面"""
    if request.method == 'GET':
        return render_template('template_generate.html')
    
    try:
        if 'template' not in request.files or 'data' not in request.files:
            return jsonify({'error': '请同时上传模板文件和数据文件'}), 400
        
        template_file = request.files['template']
        data_file = request.files['data']
        
        if template_file.filename == '' or data_file.filename == '':
            return jsonify({'error': '请选择文件'}), 400
        
        # 保存上传文件
        file_id = str(uuid.uuid4())[:8]
        template_path = app.config['UPLOAD_FOLDER'] / f"{file_id}_template_{template_file.filename}"
        data_path = app.config['UPLOAD_FOLDER'] / f"{file_id}_data_{data_file.filename}"
        
        template_file.save(str(template_path))
        data_file.save(str(data_path))
        
        # 获取输出目录
        output_dir = app.config['OUTPUT_FOLDER'] / f"{file_id}_batch"
        output_dir.mkdir(exist_ok=True)
        
        # 生成文档
        generator = WordTemplateGenerator(str(template_path))
        
        # 根据数据文件类型处理
        if data_file.filename.lower().endswith('.json'):
            import json
            with open(data_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if isinstance(data, list):
                results = generator.batch_generate(data, str(output_dir))
            else:
                output_file = output_dir / "generated.docx"
                results = [generator.generate_document(data, str(output_file))]
        
        elif data_file.filename.lower().endswith(('.xlsx', '.xls')):
            results = generator.generate_from_excel(str(data_path), output_dir=str(output_dir))
        
        else:
            return jsonify({'error': '数据文件格式不支持，请使用JSON或Excel'}), 400
        
        # 打包所有生成的文件
        import zipfile
        zip_path = app.config['OUTPUT_FOLDER'] / f"{file_id}_documents.zip"
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in output_dir.iterdir():
                if file_path.is_file():
                    zipf.write(file_path, file_path.name)
        
        output_filename = "generated_documents.zip"
        return send_file(str(zip_path),
                       as_attachment=True,
                       download_name=output_filename,
                       mimetype='application/zip')
    
    except Exception as e:
        app.logger.error(f"模板生成错误: {traceback.format_exc()}")
        return jsonify({'error': f'处理失败: {str(e)}'}), 500


@app.route('/api/health')
def health():
    """健康检查"""
    return jsonify({'status': 'ok', 'service': 'docauto'})


if __name__ == '__main__':
    print("=" * 50)
    print("企业文档自动化工具 - Web服务")
    print("=" * 50)
    print("访问地址:")
    print(f"  - 本地: http://127.0.0.1:5002")
    print(f"  - 公网: http://38.55.133.19:5002")
    print("=" * 50)
    
    app.run(host='0.0.0.0', port=5002, debug=True)
