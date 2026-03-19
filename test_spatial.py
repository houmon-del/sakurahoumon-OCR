# -*- coding: utf-8 -*-
"""空間的フィールド抽出テスト - 既存PDFを直接OCRして検証"""
import sys, os, json, glob
sys.path.insert(0, '/app')
sys.stdout.reconfigure(encoding='utf-8')

import cv2
import numpy as np
from yomitoku.data.functions import load_pdf

import consultation_spatial

# 最新のアップロードPDFを探す
upload_dirs = sorted(glob.glob('/app/uploads/batch_*'))
pdf_path = None
for d in reversed(upload_dirs):
    pdfs = glob.glob(os.path.join(d, '*.pdf'))
    if pdfs:
        pdf_path = pdfs[0]
        break

if not pdf_path:
    print("ERROR: PDFが見つかりません")
    sys.exit(1)

print(f"テスト対象PDF: {pdf_path}")

# PDF→画像に変換（300dpi、アプリと同じ設定）
pages = load_pdf(pdf_path, dpi=300)
print(f"ページ数: {len(pages)}")

# OCR実行（GPUが必要なのでDocumentAnalyzerを使用）
from yomitoku import DocumentAnalyzer
analyzer = DocumentAnalyzer(
    device="cuda",
    configs={
        "ocr": {
            "text_detector": {"infer_onnx": False},
            "text_recognizer": {"infer_onnx": False},
        },
        "layout_analyzer": {
            "layout_parser": {"infer_onnx": False},
            "table_structure_recognizer": {"infer_onnx": False},
        },
    },
)

pages_ocr = []
pages_shape = []
for i, img in enumerate(pages[:1]):  # 1ページ目のみ（load_pdfはすでにnp.ndarray BGR）
    print(f"OCR実行中 (page {i+1}, size={img.shape[1]}x{img.shape[0]})...")
    result, _, _ = analyzer(img)
    pages_ocr.append(result.model_dump())
    pages_shape.append((img.shape[1], img.shape[0]))


# OCRで読み取れたwordsを相対座標で表示
img_w, img_h = pages_shape[0]
words = pages_ocr[0].get('words', [])
print(f"\n=== OCR words一覧 ({len(words)}件) ===")
import json as _json
with open('/app/static/templates/cell_map.json', encoding='utf-8') as f:
    cmap = _json.load(f)

for w in words:
    pts = w['points']  # [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]
    xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
    x1,y1,x2,y2 = min(xs),min(ys),max(xs),max(ys)
    rx1,ry1,rx2,ry2 = x1/img_w, y1/img_h, x2/img_w, y2/img_h
    # 最も重複するフィールドを探す
    best_field = "---"
    best_ov = 0.3
    for entry in cmap['fields']:
        r = entry['rel']
        ix = max(0, min(rx2,r[2]) - max(rx1,r[0]))
        iy = max(0, min(ry2,r[3]) - max(ry1,r[1]))
        ov = (ix*iy) / max(1e-9, (rx2-rx1)*(ry2-ry1))
        if ov > best_ov:
            best_ov = ov
            best_field = entry['field']
    print(f"  rel=({rx1:.3f},{ry1:.3f})-({rx2:.3f},{ry2:.3f}) [{best_field:30}] {w['content']}")

# 空間抽出実行
print("\n=== 空間抽出結果 ===")
structured = consultation_spatial.extract_by_position_pages(pages_ocr, pages_shape)

p = structured.get("patient", {})
print(f"氏名:      {p.get('sei','')} {p.get('mei','')}")
print(f"ふりがな:  {p.get('furigana_sei','')} {p.get('furigana_mei','')}")
print(f"性別:      {p.get('gender','')}")
print(f"生年月日:  {p.get('dob_era','')} {p.get('dob_year','')}年 {p.get('dob_month','')}月 {p.get('dob_day','')}日")
print(f"年齢:      {p.get('age','')}")
print(f"住所:      {p.get('address','')}")
c = structured.get("contact", {})
print(f"自宅TEL:   {c.get('home_phone','')}")
print(f"携帯:      {c.get('mobile_phone','')}")
ins = structured.get("insurance", {})
print(f"負担割合:  {ins.get('burden_ratio','')}")
print(f"要介護度:  {ins.get('care_level','')}")
print(f"既往歴:    {structured.get('medical_history',{}).get('conditions',[])}")
phy = structured.get("physician", {})
print(f"主治医:    {phy.get('hospital','')} / {phy.get('doctor','')}")
cm = structured.get("care_manager", {})
print(f"ケアマネ:  {cm.get('name','')}  TEL:{cm.get('phone','')}")
print(f"スケジュール: {json.dumps(structured.get('schedule',{}), ensure_ascii=False)}")
