"""PDF変換ユーティリティ — LibreOffice使用"""
import os
import subprocess
import shutil
import tempfile


def convert_to_pdf(filepath):
    """docx/xlsxファイルをPDFに変換する。元ファイルと同じディレクトリにPDFを出力。"""
    output_dir = os.path.dirname(filepath)
    base_name = os.path.splitext(os.path.basename(filepath))[0]
    pdf_path = os.path.join(output_dir, base_name + '.pdf')

    lo_path = _find_libreoffice()
    if not lo_path:
        raise RuntimeError(
            'LibreOfficeが見つかりません。PDF変換にはLibreOfficeが必要です。\n'
            'インストール: sudo apt-get install libreoffice-writer libreoffice-calc fonts-noto-cjk'
        )

    # LibreOfficeは同時実行でロックファイル競合するため、
    # 一時的なユーザープロファイルを使って回避する
    with tempfile.TemporaryDirectory() as user_profile:
        proc = subprocess.run(
            [lo_path, '--headless', '--norestore',
             f'-env:UserInstallation=file://{user_profile}',
             '--convert-to', 'pdf',
             '--outdir', output_dir,
             filepath],
            capture_output=True, text=True, timeout=120
        )

    if os.path.exists(pdf_path):
        return pdf_path

    raise RuntimeError(
        f'PDF変換に失敗しました: {base_name}\n'
        f'stdout: {proc.stdout}\nstderr: {proc.stderr}'
    )


def _find_libreoffice():
    """LibreOfficeの実行パスを探す。"""
    for name in ('libreoffice', 'soffice'):
        path = shutil.which(name)
        if path:
            return path
    return None
