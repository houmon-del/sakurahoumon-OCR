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


def _para_to_rel(bbox, img_w, img_h):
    """スキャンbbox [x1,y1,x2,y2] → 相対座標 [0-1]"""
    return [bbox[0]/img_w, bbox[1]/img_h, bbox[2]/img_w, bbox[3]/img_h]


def _overlap_ratio_rel(para_rel, cell_rel):
    """
    相対座標同士の重複率。
    重複面積 / パラグラフ面積 で正規化。
    """
    px1, py1, px2, py2 = para_rel
    cx1, cy1, cx2, cy2 = cell_rel

    ix = max(0, min(px2, cx2) - max(px1, cx1))
    iy = max(0, min(py2, cy2) - max(py1, cy1))
    inter = ix * iy
    para_area = max(1e-9, (px2-px1) * (py2-py1))
    return inter / para_area


# ────────────────────────────────────────────────────────────────
# フィールドへの割り当て
# ────────────────────────────────────────────────────────────────

OVERLAP_THRESHOLD = 0.4   # パラグラフ面積の40%以上がセル内なら割り当て


def _word_to_bbox(word):
    """WordPrediction.points ([[x,y]×4]) → [x1,y1,x2,y2]"""
    pts = word["points"]
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return [min(xs), min(ys), max(xs), max(ys)]


def _assign_items(items, cell_map, img_w, img_h):
    """
    OCRアイテム（wordsまたはparagraphs）を最も重複するデータセルに割り当てる。

    返り値: dict { field_path: [(x_center, text), ...] }  ← X座標つきで保持
    """
    assignments = {}

    for item in items:
        # words と paragraphs で構造が異なる
        if "points" in item:
            text = item.get("content", "") or ""
            bbox = _word_to_bbox(item)
        else:
            text = item.get("contents", "") or ""
            bbox = item["box"]

        text = text.strip()
        if not text:
            continue

        item_rel = _para_to_rel(bbox, img_w, img_h)
        cx = (item_rel[0] + item_rel[2]) / 2

        best_field = None
        best_overlap = OVERLAP_THRESHOLD

        for entry in cell_map["fields"]:
            ov = _overlap_ratio_rel(item_rel, entry["rel"])
            if ov > best_overlap:
                best_overlap = ov
                best_field = entry["field"]

        if best_field is None:
            for entry in cell_map["schedule"]:
                ov = _overlap_ratio_rel(item_rel, entry["rel"])
                if ov > best_overlap:
                    best_overlap = ov
                    best_field = f"schedule.{entry['slot']}.{entry['day']}"

        if best_field:
            assignments.setdefault(best_field, []).append((cx, text))

    # X座標でソートして結合
    result = {}
    for field, items_list in assignments.items():
        items_list.sort(key=lambda t: t[0])
        result[field] = [t[1] for t in items_list]
    return result


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
    # wordsを優先（段落より細粒度）、なければparagraphsにフォールバック
    items = ocr_result_dict.get("words") or ocr_result_dict.get("paragraphs", [])

    assignments = _assign_items(items, cmap, img_w, img_h)
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
