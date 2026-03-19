# -*- coding: utf-8 -*-
"""
相談シートテンプレートのセル座標マップを生成する。

Excelの列幅・行高からピクセル座標を計算し、
各データフィールドのマージセル範囲をJSONとして保存する。

使い方:
    python consultation_cell_map.py  → cell_map.json を生成
"""

import json
import os
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "static", "templates", "consultation_template.xlsx")
OUTPUT_PATH   = os.path.join(os.path.dirname(__file__), "static", "templates", "cell_map.json")

# 基準DPI（この値でピクセル座標を計算。スキャン時のDPIに合わせてスケール変換する）
BASE_DPI = 150

# ──────────────────────────────────────────────
# データフィールド定義
# (field_path, write_cell)
# field_pathはstructured dictのパスに対応
# ──────────────────────────────────────────────
DATA_FIELDS = [
    ("patient.furigana",         "F5"),
    ("patient.name",             "F6"),
    ("patient.gender",           "O5"),
    ("patient.dob_era",          "U5"),
    ("patient.dob_year",         "AB5"),
    ("patient.dob_month",        "AE5"),
    ("patient.dob_day",          "AH5"),
    ("patient.age",              "Y7"),
    ("patient.address",          "G8"),
    ("patient.room",             "G9"),
    ("contact.home_phone",       "H11"),
    ("contact.mobile_phone",     "H13"),
    ("insurance.burden_ratio",   "S11"),
    ("insurance.public_expense", "W11"),
    ("insurance.care_level",     "S13"),
    ("medical_history.text",     "F15"),
    ("infection.text",           "F18"),
    ("physician.hospital",       "G19"),
    ("physician.doctor",         "U19"),
    ("communication",            "F21"),
    ("diet.type",                "H22"),
    # 訪問曜日 AM/PM × 日〜土 は schedule_cells で別定義
    ("requester.type",           "G32"),
    ("requester.name_phone",     "V32"),
    ("key_person.furigana",      "J33"),
    ("key_person.relationship",  "W33"),
    ("key_person.name",          "J34"),
    ("key_person.phone",         "G35"),
    ("key_person.address",       "B36"),
    ("care_manager.furigana",    "F40"),
    ("care_manager.phone",       "W40"),
    ("care_manager.name",        "F41"),
    ("care_manager.facility",    "C42"),
    ("care_manager.fax",         "W42"),
    ("notes",                    "C48"),
    ("referral_source",          "C52"),
]

# 訪問曜日セル (slot, day) → cell
SCHEDULE_CELLS = {
    ("am", "日"): "F26",  ("pm", "日"): "F28",
    ("am", "月"): "J26",  ("pm", "月"): "J28",
    ("am", "火"): "N26",  ("pm", "火"): "N28",
    ("am", "水"): "R26",  ("pm", "水"): "R28",
    ("am", "木"): "V26",  ("pm", "木"): "V28",
    ("am", "金"): "Z26",  ("pm", "金"): "Z28",
    ("am", "土"): "AD26", ("pm", "土"): "AD28",
}

# アンカーラベル（位置合わせ用の印刷済みテキスト）
# OCRが読み取るはずの固定テキストと、その代表セル
ANCHOR_LABELS = [
    ("ふりがな", "B5"),   # 患者ふりがなラベル
    ("名前",     "B6"),
    ("住所",     "B8"),
    ("電話番号", "B11"),
    ("既往歴",   "B15"),
    ("感染症",   "B18"),
    ("内科主治医","B19"),
    ("依頼者",   "B32"),
]


def col_width_to_px(w, dpi=BASE_DPI):
    if w is None or w == 0:
        w = 8.43  # Excel デフォルト幅
    return max(1, int((w * 7 + 5) * dpi / 96))


def row_height_to_px(h, dpi=BASE_DPI):
    if h is None or h == 0:
        h = 15.0  # Excel デフォルト高さ
    return max(1, int(h * dpi / 72))


def build_coord_tables(ws, dpi=BASE_DPI):
    """各列・行のピクセル開始座標テーブルを返す（0-indexed, 1-basedセルに対応）"""
    max_col = ws.max_column
    max_row = ws.max_row

    col_x = [0] * (max_col + 2)
    for c in range(1, max_col + 1):
        letter = get_column_letter(c)
        w = ws.column_dimensions[letter].width or 8.43
        col_x[c] = col_x[c - 1] + col_width_to_px(w, dpi)
    col_x[max_col + 1] = col_x[max_col]  # 右端

    row_y = [0] * (max_row + 2)
    for r in range(1, max_row + 1):
        h = ws.row_dimensions[r].height or 15.0
        row_y[r] = row_y[r - 1] + row_height_to_px(h, dpi)
    row_y[max_row + 1] = row_y[max_row]

    return col_x, row_y


def get_merge_extent(ws, addr):
    """セルアドレスが属するマージ範囲の (min_col, min_row, max_col, max_row) を返す"""
    for mr in ws.merged_cells.ranges:
        if addr in mr:
            return mr.min_col, mr.min_row, mr.max_col, mr.max_row
    col = column_index_from_string("".join(c for c in addr if c.isalpha()))
    row = int("".join(c for c in addr if c.isdigit()))
    return col, row, col, row


def cell_to_bbox(col_x, row_y, min_col, min_row, max_col, max_row):
    """マージ範囲 → ピクセルbbox [x1, y1, x2, y2]"""
    return [col_x[min_col - 1], row_y[min_row - 1],
            col_x[max_col],     row_y[max_row]]


def generate():
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active
    col_x, row_y = build_coord_tables(ws)

    total_w = col_x[ws.max_column]
    total_h = row_y[ws.max_row]

    # データフィールドのbbox
    fields = []
    for field_path, cell in DATA_FIELDS:
        extent = get_merge_extent(ws, cell)
        bbox   = cell_to_bbox(col_x, row_y, *extent)
        fields.append({
            "field": field_path,
            "cell":  cell,
            "bbox":  bbox,   # [x1, y1, x2, y2] in template pixels @ BASE_DPI
        })

    # スケジュールセルのbbox
    schedule = []
    for (slot, day), cell in SCHEDULE_CELLS.items():
        extent = get_merge_extent(ws, cell)
        bbox   = cell_to_bbox(col_x, row_y, *extent)
        schedule.append({
            "slot": slot,
            "day":  day,
            "cell": cell,
            "bbox": bbox,
        })

    # アンカーラベルのbbox（位置合わせ用）
    anchors = []
    for label_text, cell in ANCHOR_LABELS:
        extent = get_merge_extent(ws, cell)
        bbox   = cell_to_bbox(col_x, row_y, *extent)
        anchors.append({
            "text": label_text,
            "cell": cell,
            "bbox": bbox,
        })

    cell_map = {
        "base_dpi":   BASE_DPI,
        "template_w": total_w,
        "template_h": total_h,
        "fields":     fields,
        "schedule":   schedule,
        "anchors":    anchors,
    }

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(cell_map, f, ensure_ascii=False, indent=2)

    print(f"生成完了: {OUTPUT_PATH}")
    print(f"テンプレートサイズ: {total_w} x {total_h} px @ {BASE_DPI}dpi")
    print(f"データフィールド数: {len(fields)}")
    print(f"スケジュールセル数: {len(schedule)}")
    return cell_map


if __name__ == "__main__":
    generate()
