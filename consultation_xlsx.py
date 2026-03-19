# -*- coding: utf-8 -*-
"""相談シートExcelテンプレートへのデータ書き込み"""

import io
import os
import shutil
import openpyxl
from openpyxl.styles import Alignment

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "static", "templates", "consultation_template.xlsx")

# 来院理由マッピング (テンプレート内の選択肢テキスト → セル座標)
VISIT_REASON_CELLS = {
    "歯が痛い":           "B45",
    "グラグラ":           "J45",
    "歯ぐきが腫れた・痛い": "R45",
    "つめもの・かぶせものが取れた": "Z45",
    "口の中にできものがある": "B46",
    "入れ歯の調子が悪い":  "J46",
    "入れ歯を作りたい":    "R46",
    "歯が抜けた":          "Z46",
    "口腔ケアをして欲しい": "B47",
    "その他":             "J47",
}

# 知ったきっかけマッピング
REFERRAL_SOURCE_CELLS = {
    "HP":           "B52",
    "口コミ":        "B52",
    "院内パンフレット": "B52",
    "スタッフから":   "B52",
    "ポスターを見て": "B52",
    "他患者もさくら会に依頼しているため": "B52",
    "広告物":        "B53",
}

# 訪問曜日セルマッピング {('am'/'pm', '日'..'土'): cell}
SCHEDULE_CELLS = {
    ("am", "日"): "F26", ("pm", "日"): "F28",
    ("am", "月"): "J26", ("pm", "月"): "J28",
    ("am", "火"): "N26", ("pm", "火"): "N28",
    ("am", "水"): "R26", ("pm", "水"): "R28",
    ("am", "木"): "V26", ("pm", "木"): "V28",
    ("am", "金"): "Z26", ("pm", "金"): "Z28",
    ("am", "土"): "AD26", ("pm", "土"): "AD28",
}


def _top_left(ws, cell_addr):
    """マージ範囲の左上セルアドレスを返す。非マージなら cell_addr そのまま"""
    for mr in ws.merged_cells.ranges:
        if cell_addr in mr:
            col = openpyxl.utils.get_column_letter(mr.min_col)
            return f"{col}{mr.min_row}"
    return cell_addr


def _w(ws, cell_addr, value):
    """セルへの書き込みヘルパー（マージセルも対応）"""
    addr = _top_left(ws, cell_addr)
    ws[addr] = value
    ws[addr].alignment = Alignment(wrap_text=True, vertical="center")


def fill_template(structured: dict) -> bytes:
    """構造化データをテンプレートに埋め込みxlsxバイト列を返す"""
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active

    p   = structured.get("patient", {})
    c   = structured.get("contact", {})
    ins = structured.get("insurance", {})
    mh  = structured.get("medical_history", {})
    inf = structured.get("infection", {})
    phy = structured.get("physician", {})
    diet = structured.get("diet", {})
    sched = structured.get("schedule", {})
    req  = structured.get("requester", {})
    kp   = structured.get("key_person", {})
    cm   = structured.get("care_manager", {})

    # ── 本人情報 ──
    furigana = f"{p.get('furigana_sei','') or ''} {p.get('furigana_mei','') or ''}".strip()
    _w(ws, "F5", furigana)

    name = f"{p.get('sei','') or ''} {p.get('mei','') or ''}".strip()
    _w(ws, "F6", name)

    # 性別
    gender = p.get("gender", "")
    if gender:
        _w(ws, "P5", gender)

    # 生年月日
    dob_era  = p.get("dob_era", "")
    dob_year = p.get("dob_year", "")
    dob_mon  = p.get("dob_month", "")
    dob_day  = p.get("dob_day", "")
    if dob_era:  _w(ws, "V5",  dob_era)
    if dob_year: _w(ws, "AB5", dob_year)
    if dob_mon:  _w(ws, "AE5", dob_mon)
    if dob_day:  _w(ws, "AH5", dob_day)

    # 年齢
    if p.get("age"):
        _w(ws, "W7", p["age"])

    # 住所（〒 + 住所）
    addr_parts = []
    if p.get("postal_code"): addr_parts.append(p["postal_code"])
    if p.get("address"):     addr_parts.append(p["address"])
    if addr_parts: _w(ws, "G8", "  ".join(addr_parts))

    # 施設名・部屋番号
    if p.get("facility"): _w(ws, "G8", p["facility"] + ("　" + (p.get("address") or "")))
    if p.get("room"):     _w(ws, "G9", p["room"])

    # 駐車場
    if p.get("parking"): _w(ws, "AA8", p["parking"])

    # ── 電話番号 ──
    if c.get("home_phone"):   _w(ws, "H11", c["home_phone"])
    if c.get("mobile_phone"): _w(ws, "H13", c["mobile_phone"])

    # ── 保険情報 ──
    if ins.get("burden_ratio"):   _w(ws, "S11", f"{ins['burden_ratio']}割")
    if ins.get("public_expense"): _w(ws, "W11", ins["public_expense"])
    if ins.get("care_level"):     _w(ws, "S13", ins["care_level"])

    # ── 既往歴 ──
    conditions = mh.get("conditions") or []
    other_mh   = mh.get("other", "")
    mh_text = "　".join(conditions)
    if other_mh: mh_text += f"　その他: {other_mh}"
    if mh_text: _w(ws, "F15", mh_text)

    # ── 感染症 ──
    inf_status  = inf.get("status", "")
    inf_details = inf.get("details", "")
    if isinstance(inf_details, list): inf_details = ", ".join(inf_details)
    inf_text = inf_status
    if inf_details: inf_text += f"（{inf_details}）"
    if inf_text: _w(ws, "F18", inf_text)

    # ── 内科主治医 ──
    if phy.get("hospital"): _w(ws, "G19", phy["hospital"])
    if phy.get("doctor"):   _w(ws, "U19", phy["doctor"])

    # ── 意思疎通 ──
    if structured.get("communication"): _w(ws, "F21", structured["communication"])

    # ── 食事形態 ──
    if diet.get("type"):
        _w(ws, "H22", diet["type"])

    # ── 訪問可能曜日 ──
    days = ["日", "月", "火", "水", "木", "金", "土"]
    for slot in ("am", "pm"):
        slot_data = sched.get(slot) or {}
        for day in days:
            val = slot_data.get(day, "")
            cell_addr = SCHEDULE_CELLS.get((slot, day))
            if cell_addr and val:
                _w(ws, cell_addr, val)

    # ── 依頼者 ──
    if req.get("type"):  _w(ws, "H32", req["type"])
    if req.get("name"):  _w(ws, "V32", req["name"])
    if req.get("phone"): _w(ws, "AC32", req["phone"])

    # ── キーパーソン ──
    kp_furi = kp.get("furigana", "")
    kp_name = kp.get("name", "")
    kp_rel  = kp.get("relationship", "")
    kp_phone = kp.get("phone", "")
    kp_addr  = kp.get("address", "")
    if kp_furi:  _w(ws, "J33", kp_furi)
    if kp_rel:   _w(ws, "W33", kp_rel)
    if kp_name:  _w(ws, "J34", kp_name)
    if kp_phone: _w(ws, "H35", kp_phone)
    if kp_addr:  _w(ws, "C36", kp_addr)

    # ── ケアマネジャー ──
    if cm.get("furigana"): _w(ws, "F40", cm["furigana"])
    if cm.get("phone"):    _w(ws, "W40", cm["phone"])
    if cm.get("name"):     _w(ws, "F41", cm["name"])
    if cm.get("facility"): _w(ws, "C42", cm["facility"])
    if cm.get("fax"):      _w(ws, "W42", cm["fax"])

    # ── 来院理由 ──チェック記入
    visit_reasons = structured.get("visit_reason") or []
    for reason in visit_reasons:
        for key, cell_addr in VISIT_REASON_CELLS.items():
            if key in reason:
                current = ws[cell_addr].value or ""
                if "■" not in str(current):
                    ws[cell_addr] = str(current).replace("□", "■", 1)
                break

    # 備考（メモ欄）
    if structured.get("notes"):
        _w(ws, "C48", structured["notes"])

    # ── 知ったきっかけ ──
    referral_sources = structured.get("referral_source") or []
    if referral_sources:
        _w(ws, "C52", "　".join(referral_sources))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
