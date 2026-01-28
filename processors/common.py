"""共通処理: 書式クリーニング・エンティティ抽出・突合チェック"""
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn


def clean_formatting(doc: Document) -> Document:
    """網掛け・太字・コメント解除、黒字標準スタイルに統一する。"""
    for para in doc.paragraphs:
        for run in para.runs:
            # 太字解除
            run.bold = False
            # フォント色を黒に
            run.font.color.rgb = RGBColor(0, 0, 0)
            # 網掛け（ハイライト）解除
            run.font.highlight_color = None
            # シェーディング（背景色）解除
            rpr = run._element.find(qn('w:rPr'))
            if rpr is not None:
                shd = rpr.find(qn('w:shd'))
                if shd is not None:
                    rpr.remove(shd)

    # コメント削除
    _remove_comments(doc)
    return doc


def _remove_comments(doc: Document):
    """Word文書からコメントを削除する。"""
    body = doc.element.body
    # コメント参照を削除
    for tag in ('w:commentRangeStart', 'w:commentRangeEnd', 'w:commentReference'):
        for el in body.iter(qn(tag)):
            el.getparent().remove(el)
    # コメントパーツ自体を削除
    comments_part_name = '/word/comments.xml'
    if hasattr(doc.part, 'package'):
        try:
            parts = doc.part.package.parts
            to_remove = [p for p in parts if hasattr(p, 'partname')
                         and str(p.partname) == comments_part_name]
            for p in to_remove:
                parts.pop(p.partname)
        except Exception:
            pass


def extract_entity_info(doc: Document) -> dict:
    """文書から法人名・住所・役職者名・代表者名を抽出する。"""
    text = '\n'.join(p.text for p in doc.paragraphs)
    info = {
        'company': '',
        'address': '',
        'title': '',
        'representative': '',
    }

    # 法人名: 「乙」の後に続く法人名を探す
    m = re.search(r'乙\s*[：:]\s*(.+)', text)
    if m:
        info['company'] = m.group(1).strip()

    # 住所
    m = re.search(r'(?:住所|所在地)\s*[：:]\s*(.+)', text)
    if m:
        info['address'] = m.group(1).strip()

    # 代表者名
    m = re.search(r'代表(?:取締役|者)\s*(.+)', text)
    if m:
        info['representative'] = m.group(1).strip()

    return info


def cross_check_entities(entity_infos: dict) -> list:
    """複数書類間で法人情報が一致しているか突合する。"""
    errors = []
    keys_to_check = ['company', 'address', 'representative']
    labels = {'company': '法人名', 'address': '住所', 'representative': '代表者名'}

    docs = list(entity_infos.keys())
    if len(docs) < 2:
        return errors

    base_doc = docs[0]
    base_info = entity_infos[base_doc]

    for doc_name in docs[1:]:
        info = entity_infos[doc_name]
        for key in keys_to_check:
            base_val = base_info.get(key, '').strip()
            cur_val = info.get(key, '').strip()
            if base_val and cur_val and base_val != cur_val:
                errors.append(
                    f'【整合性エラー】{labels[key]}が不一致: '
                    f'{base_doc}「{base_val}」≠ {doc_name}「{cur_val}」'
                )

    return errors
