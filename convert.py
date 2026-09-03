#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
藥劑科小藥撥補單 - 主管格式轉換器
讀取系統匯出的 23 欄 Excel，自動過濾並產出主管要求的 14 欄標準格式（含「單位簽收人/日期」）。
"""

import sys
import os
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 主管要求的標準欄位對應表 (來源欄位名 -> 目標表頭文字)
TARGET_COLUMN_MAP = [
    ('申請單號', '申請單號'),
    ('申請時間', '申請時間'),
    ('申請單位', '申請\n單位'),
    ('費用部門代號', '費用部門代號'),
    ('申請人', '申請人'),
    ('送簽主管', '送簽\n主管'),
    ('實際簽核主管', '實際簽核主管'),
    ('發料日期', '發料日期'),
    ('領料品項名稱', '領料品項名稱'),
    ('材料編號', '材料編號'),
    ('單位', '單位'),
    ('申請數量', '申請\n數量'),
    ('發料數量', '發料\n數量'),
    (None, '單位簽收人/日期') # 新增欄位，初始為空
]

# 建議欄寬設定 (對照主管 Excel 範本比例)
COLUMN_WIDTHS = {
    1: 14,   # 申請單號
    2: 20,   # 申請時間
    3: 11,   # 申請單位
    4: 13,   # 費用部門代號
    5: 10,   # 申請人
    6: 10,   # 送簽主管
    7: 13,   # 實際簽核主管
    8: 12,   # 發料日期
    9: 28,   # 領料品項名稱
    10: 16,  # 材料編號
    11: 8,   # 單位
    12: 10,  # 申請數量
    13: 10,  # 發料數量
    14: 18   # 單位簽收人/日期
}

def convert_replenishment_excel(src_path, dest_path=None):
    if not os.path.exists(src_path):
        print(f"錯誤：找不到檔案 {src_path}")
        return None

    if dest_path is None:
        base, ext = os.path.splitext(src_path)
        dest_path = f"{base}_正確格式.xlsx"

    print(f"正在讀取：{src_path}")
    src_wb = openpyxl.load_workbook(src_path)
    src_ws = src_wb.active

    # 讀取來源表頭並建立索引對應
    src_headers = [cell.value for cell in src_ws[1]]
    col_mapping = []
    for src_col, target_header in TARGET_COLUMN_MAP:
        if src_col is not None:
            if src_col in src_headers:
                col_mapping.append((src_headers.index(src_col), target_header))
            else:
                print(f"警告：來源資料缺少欄位「{src_col}」，將填空值")
                col_mapping.append((None, target_header))
        else:
            # 額外新增的欄位
            col_mapping.append((None, target_header))

    # 建立目標活頁簿
    dest_wb = openpyxl.Workbook()
    dest_ws = dest_wb.active
    dest_ws.title = "藥品材料申請單"

    # 樣式定義
    font_header = Font(name="新細明體", size=11, bold=True)
    font_data = Font(name="新細明體", size=10, bold=False)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    thin_border = Border(
        left=Side(style='thin', color='A0A0A0'),
        right=Side(style='thin', color='A0A0A0'),
        top=Side(style='thin', color='A0A0A0'),
        bottom=Side(style='thin', color='A0A0A0')
    )

    # 寫入表頭 (第 1 列)
    dest_ws.row_dimensions[1].height = 28
    for col_idx, (_, target_header) in enumerate(col_mapping, start=1):
        cell = dest_ws.cell(row=1, column=col_idx, value=target_header)
        cell.font = font_header
        cell.alignment = align_center
        cell.border = thin_border

    # 寫入資料列 (第 2 列起)
    for row_idx, src_row in enumerate(src_ws.iter_rows(min_row=2, values_only=True), start=2):
        dest_ws.row_dimensions[row_idx].height = 24
        for col_idx, (src_idx, target_header) in enumerate(col_mapping, start=1):
            val = src_row[src_idx] if src_idx is not None else ""
            if val is None:
                val = ""
            cell = dest_ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_data
            cell.border = thin_border
            
            # 品項名稱偏左對齊，其餘欄位置中
            if '品項名稱' in target_header:
                cell.alignment = align_left
            else:
                cell.alignment = align_center

    # 設定欄寬
    for col_idx, width in COLUMN_WIDTHS.items():
        col_letter = get_column_letter(col_idx)
        dest_ws.column_dimensions[col_letter].width = width

    # 列印設定：A4 橫向、自動符合寬度
    dest_ws.page_setup.orientation = dest_ws.ORIENTATION_LANDSCAPE
    dest_ws.page_setup.paperSize = dest_ws.PAPERSIZE_A4
    dest_ws.page_setup.fitToPage = True
    dest_ws.page_setup.fitToWidth = 1
    dest_ws.page_setup.fitToHeight = 0
    dest_ws.sheet_properties.pageSetUpPr.fitToPage = True

    # 儲存
    dest_wb.save(dest_path)
    print(f"✅ 成功產出正確格式檔案：{dest_path}")
    return dest_path

if __name__ == '__main__':
    target_file = sys.argv[1] if len(sys.argv) > 1 else '/Users/jiangruiyi/Downloads/藥品材料申請單_2026-08-27_to_2026-09-02.xlsx'
    convert_replenishment_excel(target_file)
