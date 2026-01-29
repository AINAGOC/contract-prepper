import os
import io
import zipfile
import tempfile
import shutil
import traceback
import logging
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file
import uuid
from werkzeug.utils import secure_filename

# ログ設定
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

from processors.common import clean_formatting, extract_entity_info, cross_check_entities
from processors.contract import process_contract
from processors.estimate import process_estimate
from processors.oath import process_oath
from processors.checklist import process_checklist
from processors.confirmation import process_confirmation
from processors.pdf_converter import convert_to_pdf

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()

APPENDIX2_DIR = os.path.join(os.path.dirname(__file__), 'assets', 'appendix2')

FILE_TYPES = {
    'contract': {'label': '① 契約書', 'naming': '基本契約書_{company}'},
    'estimate': {'label': '② 見積書', 'naming': '別紙１_{company}'},
    'oath': {'label': '③ 誓約書', 'naming': '誓約書_{company}'},
    'checklist': {'label': '④ チェックシート', 'naming': '持続可能性の確保に向けた取組状況について（チェックシート）_{company}'},
    'confirmation': {'label': '⑤ 確認書', 'naming': '電子契約サービス利用確認書_{company}'},
}


@app.route('/')
def index():
    appendix2_files = []
    if os.path.isdir(APPENDIX2_DIR):
        appendix2_files = [f for f in os.listdir(APPENDIX2_DIR)
                           if f.endswith(('.docx', '.xlsx', '.pdf'))]
    return render_template('index.html', file_types=FILE_TYPES,
                           appendix2_files=appendix2_files)


@app.route('/process', methods=['POST'])
def process_files():
    company_name = request.form.get('company_name', '').strip()
    if not company_name:
        return jsonify({'error': '会社名を入力してください。'}), 400

    approval_type = request.form.get('approval_type', 'paper')  # paper or electronic
    appendix2_choice = request.form.get('appendix2_choice', '')

    work_dir = tempfile.mkdtemp()
    output_dir = os.path.join(work_dir, 'output')
    backup_dir = os.path.join(work_dir, 'backup')
    os.makedirs(output_dir)
    os.makedirs(backup_dir)

    results = {'processed': [], 'errors': [], 'warnings': []}
    uploaded_docs = {}

    # Save uploaded files and create backups
    for key in FILE_TYPES:
        file = request.files.get(key)
        if file and file.filename:
            # secure_filename strips Japanese chars, so preserve extension manually
            original_name = file.filename
            ext = os.path.splitext(original_name)[1].lower()
            safe_name = f'{key}_{uuid.uuid4().hex[:8]}{ext}'
            filepath = os.path.join(work_dir, safe_name)
            file.save(filepath)
            uploaded_docs[key] = filepath
            # Backup original with original filename (sanitised minimally)
            backup_name = original_name.replace('/', '_').replace('\\', '_')
            shutil.copy2(filepath, os.path.join(backup_dir, backup_name))
        else:
            results['warnings'].append(f'{FILE_TYPES[key]["label"]} がスキップされました（未アップロード）。')

    if not uploaded_docs:
        shutil.rmtree(work_dir)
        return jsonify({'error': '少なくとも1つのファイルをアップロードしてください。'}), 400

    # Process each document
    entity_infos = {}
    logger.info(f"Processing {len(uploaded_docs)} files for company: {company_name}")

    try:
        if 'contract' in uploaded_docs:
            logger.info("Processing contract...")
            res = process_contract(
                uploaded_docs['contract'], output_dir, company_name,
                approval_type, appendix2_choice, APPENDIX2_DIR
            )
            results['processed'].append(res['output_name'])
            results['errors'].extend(res.get('errors', []))
            results['warnings'].extend(res.get('warnings', []))
            if res.get('entity_info'):
                entity_infos['contract'] = res['entity_info']

        if 'estimate' in uploaded_docs:
            res = process_estimate(uploaded_docs['estimate'], output_dir, company_name)
            results['processed'].append(res['output_name'])
            results['errors'].extend(res.get('errors', []))
            if res.get('entity_info'):
                entity_infos['estimate'] = res['entity_info']

        if 'oath' in uploaded_docs:
            res = process_oath(uploaded_docs['oath'], output_dir, company_name)
            results['processed'].append(res['output_name'])
            results['errors'].extend(res.get('errors', []))
            results['warnings'].extend(res.get('warnings', []))
            if res.get('entity_info'):
                entity_infos['oath'] = res['entity_info']

        if 'checklist' in uploaded_docs:
            res = process_checklist(uploaded_docs['checklist'], output_dir, company_name)
            results['processed'].append(res['output_name'])
            results['errors'].extend(res.get('errors', []))
            if res.get('entity_info'):
                entity_infos['checklist'] = res['entity_info']

        if 'confirmation' in uploaded_docs:
            res = process_confirmation(uploaded_docs['confirmation'], output_dir, company_name)
            results['processed'].append(res['output_name'])
            results['errors'].extend(res.get('errors', []))
            if res.get('entity_info'):
                entity_infos['confirmation'] = res['entity_info']

        # Cross-check entity info
        if len(entity_infos) > 1:
            cross_errors = cross_check_entities(entity_infos)
            results['errors'].extend(cross_errors)

        # Convert all output docx/xlsx to PDF
        logger.info(f"Converting files to PDF in {output_dir}")
        for fname in os.listdir(output_dir):
            fpath = os.path.join(output_dir, fname)
            if fname.endswith(('.docx', '.xlsx')):
                try:
                    logger.info(f"Converting {fname} to PDF...")
                    pdf_path = convert_to_pdf(fpath)
                    logger.info(f"Converted: {pdf_path}")
                    os.remove(fpath)
                except Exception as e:
                    logger.error(f"PDF conversion failed for {fname}: {str(e)}\n{traceback.format_exc()}")
                    results['warnings'].append(f'{fname} のPDF変換に失敗: {str(e)}。元ファイルを同梱します。')

        # Create ZIP with output + backup
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            for fname in os.listdir(output_dir):
                zf.write(os.path.join(output_dir, fname), f'成果物/{fname}')
            for fname in os.listdir(backup_dir):
                zf.write(os.path.join(backup_dir, fname), f'バックアップ/{fname}')

        zip_buffer.seek(0)
        shutil.rmtree(work_dir)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'契約書_{company_name}_{timestamp}.zip'
        )

    except Exception as e:
        logger.error(f"Processing error: {str(e)}\n{traceback.format_exc()}")
        shutil.rmtree(work_dir, ignore_errors=True)
        return jsonify({'error': f'処理中にエラーが発生しました: {str(e)}'}), 500


@app.route('/validate', methods=['POST'])
def validate_only():
    """Run validation checks without producing output files."""
    company_name = request.form.get('company_name', '').strip()
    approval_type = request.form.get('approval_type', 'paper')
    results = {'errors': [], 'warnings': []}

    if not company_name:
        results['errors'].append('会社名を入力してください。')

    uploaded_count = sum(1 for key in FILE_TYPES if request.files.get(key) and request.files[key].filename)
    if uploaded_count == 0:
        results['errors'].append('少なくとも1つのファイルをアップロードしてください。')

    return jsonify(results)


if __name__ == '__main__':
    import webbrowser, threading
    threading.Timer(1.0, lambda: webbrowser.open('http://localhost:5000')).start()
    app.run(debug=False, port=5000)
