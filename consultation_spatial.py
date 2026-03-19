# -*- coding: utf-8 -*-
"""
相談シートの空間的（位置ベース）フィールド抽出エンジン。

OCRが返すbounding boxと、テンプレートのセル座標マップを照合し、
各テキストがどのフィールドに属するかを特定する。

流れ:
    1. cell_map.json をロード（テンプレートのセル座標）
    2. OCRパラグラフのbboxをスキャン座標→テンプレート座標に変換
       - アンカーラベルが検出されれば精密アライメント
       - なければページサイズからスケール推定
    3. 各パラグラフをセルに割り当て → structured dict を返す
"""

import json
import os
import re

CELL_MAP_PATH = os.path.join(os.path.dirname(__file__), "static", "templates", "cell_map.json")

_cell_map = None


def _load_cell_map():
    global _cell_map
    if _cell_map is None:
        with open(CELL_MAP_PATH, encoding="utf-8") as f:
            _cell_map = json.load(f)
    return _cell_map


# ────────────────────────────────────────────────────────────────
# 座標変換
# ────────────────────────────────────────────────────────────────

def _para_center(bbox):
    """ParagraphSchema.box = [x1,y1,x2,y2] の中心を返す"""
    return (bbox[0] + bbox[2]) / 2, (bbox[1] + bbox[3]) / 2


def _estimate_transform(ocr_paragraphs, cell_map, img_w, img_h):
    """
    スキャン座標 → テンプレート座標 の (scale_x, scale_y, offset_x, offset_y) を推定。

    まずアンカーラベルで精密推定を試みる。
    失敗した場合はページサイズ比率から簡易推定。

    返り値: (sx, sy, ox, oy)
        template_x = scan_x * sx + ox
        template_y = scan_y * sy + oy
    """
    cmap = cell_map
    anchors = cmap["anchors"]

    # --- アンカーラベルマッチング ---
    scan_pts  = []   # (scan_cx, scan_cy)
    tmpl_pts  = []   # (tmpl_cx, tmpl_cy)

    for anchor in anchors:
        text = anchor["text"]
        tbbox = anchor["bbox"]
        tcx = (tbbox[0] + tbbox[2]) / 2
        tcy = (tbbox[1] + tbbox[3]) / 2

        # OCRパラグラフからラベルテキストを探す
        for para in ocr_paragraphs:
            if para.get("contents") and text in para["contents"]:
                b = para["box"]  # [x1,y1,x2,y2]
                scx = (b[0] + b[2]) / 2
                scy = (b[1] + b[3]) / 2
                scan_pts.append((scx, scy))
                tmpl_pts.append((tcx, tcy))
                break

    if len(scan_pts) >= 3:
        # 最小二乗でスケール・オフセットを推定（回転は無視）
        import numpy as np
        SP = np.array(scan_pts)
        TP = np.array(tmpl_pts)
        sx = float(np.mean(TP[:, 0] / (SP[:, 0] + 1e-9)))
        sy = float(np.mean(TP[:, 1] / (SP[:, 1] + 1e-9)))
        ox = float(np.mean(TP[:, 0] - SP[:, 0] * sx))
        oy = float(np.mean(TP[:, 1] - SP[:, 1] * sy))
        return sx, sy, ox, oy

    # --- フォールバック: ページサイズ比率 ---
    tw = cmap["template_w"]
    th = cmap["template_h"]
    sx = tw / max(img_w, 1)
    sy = th / max(img_h, 1)
    return sx, sy, 0.0, 0.0


def _to_template(bbox_scan, sx, sy, ox, oy):
    """スキャンbbox → テンプレート座標bbox"""
    x1, y1, x2, y2 = bbox_scan
    return [x1*sx+ox, y1*sy+oy, x2*sx+ox, y2*sy+oy]


def _iou_1d(a1, a2, b1, b2):
    """1次元のオーバーラップ比率 (intersection / union)"""
    inter = max(0, min(a2, b2) - max(a1, b1))
    union = max(a2, b2) - min(a1, b1)
    return inter / union if union > 0 else 0


def _overlap_ratio(para_bbox, cell_bbox):
    """
    パラグラフbboxとセルbboxの重複率を返す。
    重複面積 / パラグラフ面積 で正規化（セルが小さくても大きくてもパラグラフ基準）
    """
    px1, py1, px2, py2 = para_bbox
    cx1, cy1, cx2, cy2 = cell_bbox

    ix = max(0, min(px2, cx2) - max(px1, cx1))
    iy = max(0, min(py2, cy2) - max(py1, cy1))
    inter = ix * iy
    para_area = max(1, (px2-px1) * (py2-py1))
    return inter / para_area


# ────────────────────────────────────────────────────────────────
# フィールドへの割り当て
# ────────────────────────────────────────────────────────────────

OVERLAP_THRESHOLD = 0.4   # パラグラフ面積の40%以上がセル内なら割り当て


def _assign_paragraphs(ocr_paragraphs, cell_map, sx, sy, ox, oy):
    """
    各OCRパラグラフを最も重複するデータセルに割り当てる。

    返り値: dict { field_path: [text, ...] }
    """
    assignments = {}

    for para in ocr_paragraphs:
        text = para.get("contents", "") or ""
        text = text.strip()
        if not text:
            continue

        pbbox_tmpl = _to_template(para["box"], sx, sy, ox, oy)

        best_field = None
        best_overlap = OVERLAP_THRESHOLD

        # データフィールドとの重複確認
        for entry in cell_map["fields"]:
            ov = _overlap_ratio(pbbox_tmpl, entry["bbox"])
            if ov > best_overlap:
                best_overlap = ov
                best_field = entry["field"]

        # スケジュールセルとの重複確認
        if best_field is None:
            for entry in cell_map["schedule"]:
                ov = _overlap_ratio(pbbox_tmpl, entry["bbox"])
                if ov > best_overlap:
                    best_overlap = ov
                    best_field = f"schedule.{entry['slot']}.{entry['day']}"

        if best_field:
            assignments.setdefault(best_field, []).append(text)

    return assignments


# ────────────────────────────────────────────────────────────────
# structured dict への変換
# ────────────────────────────────────────────────────────────────

def _to_structured(assignments):
    """
    { field_path: [texts] } → consultation_xlsx.fill_template() が受け取るformat に変換。

    AI抽出と同じ構造にするため、不完全なフィールドは空文字で埋める。
    """
    def joined(key):
        return "　".join(assignments.get(key, []))

    def first(key):
        vals = assignments.get(key, [])
        return vals[0] if vals else ""

    # 訪問曜日
    schedule = {}
    for key, texts in assignments.items():
        if key.startswith("schedule."):
            _, slot, day = key.split(".", 2)
            schedule.setdefault(slot, {})[day] = texts[0] if texts else ""

    # 既往歴: カンマ/全角スペース区切りをリストに
    mh_text = joined("medical_history.text")
    conditions = [c.strip() for c in re.split(r"[,、　\s]+", mh_text) if c.strip()]

    # 感染症
    inf_text = joined("infection.text")

    # 依頼者: name_phone をスペースで分割（例: "田中太郎　090-xxxx-xxxx"）
    req_np = first("requester.name_phone")
    req_parts = req_np.split("　") if "　" in req_np else [req_np, ""]

    structured = {
        "patient": {
            "furigana_sei": "",
            "furigana_mei": "",
            "sei": "",
            "mei": "",
            "_furigana_raw": joined("patient.furigana"),
            "_name_raw":     joined("patient.name"),
            "gender":    first("patient.gender"),
            "dob_era":   first("patient.dob_era"),
            "dob_year":  first("patient.dob_year"),
            "dob_month": first("patient.dob_month"),
            "dob_day":   first("patient.dob_day"),
            "age":       first("patient.age"),
            "address":   joined("patient.address"),
            "room":      first("patient.room"),
        },
        "contact": {
            "home_phone":   first("contact.home_phone"),
            "mobile_phone": first("contact.mobile_phone"),
        },
        "insurance": {
            "burden_ratio":   first("insurance.burden_ratio"),
            "public_expense": first("insurance.public_expense"),
            "care_level":     first("insurance.care_level"),
        },
        "medical_history": {
            "conditions": conditions,
            "other": "",
        },
        "infection": {
            "status":  inf_text,
            "details": [],
        },
        "physician": {
            "hospital": first("physician.hospital"),
            "doctor":   first("physician.doctor"),
        },
        "communication": first("communication"),
        "diet": {
            "type": first("diet.type"),
        },
        "schedule": schedule,
        "requester": {
            "type":  first("requester.type"),
            "name":  req_parts[0].strip(),
            "phone": req_parts[1].strip() if len(req_parts) > 1 else "",
        },
        "key_person": {
            "furigana":     joined("key_person.furigana"),
            "name":         first("key_person.name"),
            "relationship": first("key_person.relationship"),
            "phone":        first("key_person.phone"),
            "address":      joined("key_person.address"),
        },
        "care_manager": {
            "furigana": joined("care_manager.furigana"),
            "name":     first("care_manager.name"),
            "facility": joined("care_manager.facility"),
            "phone":    first("care_manager.phone"),
            "fax":      first("care_manager.fax"),
        },
        "notes": joined("notes"),
        "referral_source": assignments.get("referral_source", []),
        "_source": "spatial",
    }

    # 氏名: "姓 名" の分割を試みる
    name_raw = structured["patient"]["_name_raw"]
    if name_raw:
        parts = name_raw.split()
        if len(parts) >= 2:
            structured["patient"]["sei"] = parts[0]
            structured["patient"]["mei"] = " ".join(parts[1:])
        else:
            structured["patient"]["sei"] = name_raw

    furi_raw = structured["patient"]["_furigana_raw"]
    if furi_raw:
        parts = furi_raw.split()
        if len(parts) >= 2:
            structured["patient"]["furigana_sei"] = parts[0]
            structured["patient"]["furigana_mei"] = " ".join(parts[1:])
        else:
            structured["patient"]["furigana_sei"] = furi_raw

    return structured


# ────────────────────────────────────────────────────────────────
# メインAPI
# ────────────────────────────────────────────────────────────────

def extract_by_position(ocr_result_dict, img_w, img_h):
    """
    OCR結果dictと画像サイズからフィールドを空間的に抽出する。

    Args:
        ocr_result_dict: result.model_dump() の出力
            { "paragraphs": [{box, contents, ...}], "words": [...], ... }
        img_w, img_h: スキャン画像の幅・高さ (pixels)

    Returns:
        structured dict (consultation_xlsx.fill_template() と同形式)
    """
    cmap = _load_cell_map()
    paragraphs = ocr_result_dict.get("paragraphs", [])

    sx, sy, ox, oy = _estimate_transform(paragraphs, cmap, img_w, img_h)
    assignments = _assign_paragraphs(paragraphs, cmap, sx, sy, ox, oy)
    return _to_structured(assignments)


def extract_by_position_pages(pages_ocr, pages_shape):
    """
    複数ページ対応（1ページ目のみ処理、相談シートは1枚想定）

    Args:
        pages_ocr:   list of ocr_result_dict (ページ順)
        pages_shape: list of (img_w, img_h)  (ページ順)
    """
    if not pages_ocr:
        return {}
    # 相談シートは1枚なので1ページ目のみ
    return extract_by_position(pages_ocr[0], *pages_shape[0])
