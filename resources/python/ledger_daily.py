#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Daily ledger updater for the “融资及还款明细” sheet.

Rules are defined by docs/1融资及还款明细说明.txt and must:
- preserve every existing cell (values, formulas, styles);
- append new rows for 放款明细 + 保理/再保理融资还款明细 in order;
- merge AI cells with the same AR & AE (=目标日期) and sum AH.
"""

from __future__ import annotations

import argparse
import datetime as dt
from copy import copy
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string, get_column_letter

# === 常量 ===

TEMPLATE_ROW_INDEX = 10
TARGET_SHEET_KEYWORD = "融资及还款明细"

# 目标表列
COL_A = column_index_from_string("A")
COL_B = column_index_from_string("B")
COL_C = column_index_from_string("C")
COL_D = column_index_from_string("D")
COL_E = column_index_from_string("E")
COL_F = column_index_from_string("F")
COL_G = column_index_from_string("G")
COL_H = column_index_from_string("H")
COL_I = column_index_from_string("I")
COL_J = column_index_from_string("J")
COL_K = column_index_from_string("K")
COL_L = column_index_from_string("L")
COL_M = column_index_from_string("M")
COL_N = column_index_from_string("N")
COL_O = column_index_from_string("O")
COL_P = column_index_from_string("P")
COL_Q = column_index_from_string("Q")
COL_R = column_index_from_string("R")
COL_S = column_index_from_string("S")
COL_T = column_index_from_string("T")
COL_U = column_index_from_string("U")
COL_V = column_index_from_string("V")
COL_W = column_index_from_string("W")
COL_X = column_index_from_string("X")
COL_Y = column_index_from_string("Y")
COL_Z = column_index_from_string("Z")
COL_AA = column_index_from_string("AA")
COL_AB = column_index_from_string("AB")
COL_AC = column_index_from_string("AC")
COL_AD = column_index_from_string("AD")
COL_AE = column_index_from_string("AE")
COL_AF = column_index_from_string("AF")
COL_AG = column_index_from_string("AG")
COL_AH = column_index_from_string("AH")
COL_AI = column_index_from_string("AI")
COL_AJ = column_index_from_string("AJ")
COL_AK = column_index_from_string("AK")
COL_AR = column_index_from_string("AR")

# 放款明细列
LOAN_COL_B = column_index_from_string("B")
LOAN_COL_AE = column_index_from_string("AE")
LOAN_COL_AG = column_index_from_string("AG")
LOAN_COL_K = column_index_from_string("K")
LOAN_COL_C = column_index_from_string("C")
LOAN_COL_G = column_index_from_string("G")
LOAN_COL_AZ = column_index_from_string("AZ")
LOAN_COL_J = column_index_from_string("J")
LOAN_COL_AA = column_index_from_string("AA")
LOAN_COL_L = column_index_from_string("L")
LOAN_COL_AK = column_index_from_string("AK")
LOAN_COL_N = column_index_from_string("N")
LOAN_COL_M = column_index_from_string("M")
LOAN_COL_Y = column_index_from_string("Y")
LOAN_COL_AW = column_index_from_string("AW")
LOAN_COL_P = column_index_from_string("P")
LOAN_COL_BC = column_index_from_string("BC")
LOAN_COL_BF = column_index_from_string("BF")
LOAN_COL_Q = column_index_from_string("Q")
LOAN_COL_T = column_index_from_string("T")
LOAN_COL_U = column_index_from_string("U")

# 还款明细列（保理/再保理）
REPAY_COL_G = column_index_from_string("G")
REPAY_COL_H = column_index_from_string("H")
REPAY_COL_J = column_index_from_string("J")
REPAY_COL_M = column_index_from_string("M")
REPAY_COL_B = column_index_from_string("B")
REPAY_COL_C = column_index_from_string("C")
REPAY_COL_F = column_index_from_string("F")
REPAY_COL_AE = column_index_from_string("AE")
REPAY_COL_O = column_index_from_string("O")
REPAY_COL_AG = column_index_from_string("AG")
REPAY_COL_AH = column_index_from_string("AH")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Append ledger financing & repayment rows.")
    parser.add_argument("--ledger", required=True, help="现有台账文件路径")
    parser.add_argument("--loan", required=True, help="放款明细路径")
    parser.add_argument("--factoring-repay", required=True, help="保理融资还款明细路径")
    parser.add_argument("--refactoring-repay", required=True, help="再保理融资还款明细路径")
    parser.add_argument("--date", required=True, help="目标日期，格式 YYYYMMDD")
    parser.add_argument("--output", required=True, help="输出文件路径")
    return parser.parse_args()


def parse_input_date(date_str: str) -> dt.date:
    try:
        return dt.datetime.strptime(date_str, "%Y%m%d").date()
    except ValueError as exc:  # pragma: no cover - 防御性校验
        raise SystemExit(f"无效日期格式: {date_str}，需为 YYYYMMDD") from exc


def normalize_excel_date(value) -> Optional[dt.date]:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, str):
        text = value.strip()
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
            try:
                return dt.datetime.strptime(text, fmt).date()
            except ValueError:
                continue
    return None


def find_target_sheet(wb) -> "Worksheet":
    if TARGET_SHEET_KEYWORD in wb.sheetnames:
        return wb[TARGET_SHEET_KEYWORD]
    for name in wb.sheetnames:
        if "融资" in name and "还款" in name:
            return wb[name]
    raise SystemExit("未找到目标工作表：融资及还款明细")


def find_last_data_row(ws) -> int:
    # 使用 A 列（序号公式）向上查找
    for row_idx in range(ws.max_row, 0, -1):
        if ws.cell(row=row_idx, column=COL_A).value not in (None, ""):
            return row_idx
    return 0


def get_last_existing_date(ws) -> Optional[dt.date]:
    last_row = find_last_data_row(ws)
    if last_row == 0:
        return None
    last_w = normalize_excel_date(ws.cell(row=last_row, column=COL_W).value)
    last_ae = normalize_excel_date(ws.cell(row=last_row, column=COL_AE).value)
    return last_w or last_ae


def cache_template_row(ws, row_index: int) -> Dict[int, Dict[str, object]]:
    cached: Dict[int, Dict[str, object]] = {}
    for cell in ws[row_index]:
        cached[cell.column] = {
            "value": cell.value,
            "data_type": cell.data_type,
            "style": copy(cell._style),
        }
    return cached


def apply_template_row(ws, template_cache, target_row: int, template_row_index: int, template_height: Optional[float]):
    for col_idx, meta in template_cache.items():
        tpl_value = meta["value"]
        tpl_type = meta["data_type"]
        target_cell = ws.cell(row=target_row, column=col_idx)
        if meta.get("style"):
            target_cell._style = copy(meta["style"])

        if tpl_type == "f" and isinstance(tpl_value, str):
            origin = f"{get_column_letter(col_idx)}{template_row_index}"
            target = f"{get_column_letter(col_idx)}{target_row}"
            translator = Translator(tpl_value, origin=origin)
            # openpyxl 3.1 的接口是 translate_formula，而旧版本是 translate
            if hasattr(translator, "translate_formula"):
                target_cell.value = translator.translate_formula(target)
            else:  # pragma: no cover - 兼容旧接口
                target_cell.value = translator.translate(target)
        else:
            target_cell.value = tpl_value

    if template_height:
        ws.row_dimensions[target_row].height = template_height


def collect_loan_rows(path: Path, target_date: dt.date) -> List[Sequence]:
    wb = load_workbook(path, read_only=True, data_only=False)
    ws = wb.active
    matched: List[Sequence] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if normalize_excel_date(row[LOAN_COL_P - 1]) == target_date:
            matched.append(row)
    wb.close()
    return matched


def collect_repay_rows(path: Path, target_date: dt.date) -> List[Sequence]:
    wb = load_workbook(path, read_only=True, data_only=False)
    ws = wb.active
    matched: List[Sequence] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if normalize_excel_date(row[REPAY_COL_AE - 1]) != target_date:
            continue
        fee_type = row[column_index_from_string("AB") - 1]
        if (fee_type or "").strip() != "本金":
            continue
        matched.append(row)
    wb.close()
    # 按交易银行流水号（AH）升序
    return sorted(matched, key=lambda r: (r[REPAY_COL_AH - 1] is None, str(r[REPAY_COL_AH - 1])))


def set_cell(ws, row_idx: int, col_idx: int, value):
    ws.cell(row=row_idx, column=col_idx, value=value)


def append_loan_block(ws, template_cache, template_height, template_row_index, rows: Iterable[Sequence]) -> int:
    added = 0
    for record in rows:
        target_row = ws.max_row + 1
        apply_template_row(ws, template_cache, target_row, template_row_index, template_height)

        set_cell(ws, target_row, COL_D, record[LOAN_COL_B - 1])
        set_cell(ws, target_row, COL_E, record[LOAN_COL_AE - 1])
        set_cell(ws, target_row, COL_F, record[LOAN_COL_AG - 1])
        set_cell(ws, target_row, COL_G, record[LOAN_COL_K - 1])
        set_cell(ws, target_row, COL_H, record[LOAN_COL_C - 1])
        set_cell(ws, target_row, COL_I, record[LOAN_COL_G - 1])
        set_cell(ws, target_row, COL_J, record[LOAN_COL_AZ - 1])
        set_cell(ws, target_row, COL_K, record[LOAN_COL_J - 1])
        set_cell(ws, target_row, COL_L, record[LOAN_COL_AA - 1])
        set_cell(ws, target_row, COL_N, record[LOAN_COL_AK - 1])
        # O 列保留模板公式 (=GX&"-"&TX)
        set_cell(ws, target_row, COL_P, record[LOAN_COL_M - 1])
        set_cell(ws, target_row, COL_Q, record[LOAN_COL_L - 1])
        set_cell(ws, target_row, COL_R, record[LOAN_COL_AK - 1])
        biz_mode = record[LOAN_COL_Y - 1]
        set_cell(ws, target_row, COL_S, "保理" if isinstance(biz_mode, str) and biz_mode.strip() == "直接投放" else "再保理")
        set_cell(ws, target_row, COL_T, record[LOAN_COL_N - 1])
        set_cell(ws, target_row, COL_U, record[LOAN_COL_AW - 1])
        set_cell(ws, target_row, COL_V, record[LOAN_COL_AW - 1])
        set_cell(ws, target_row, COL_W, record[LOAN_COL_P - 1])
        set_cell(ws, target_row, COL_X, record[LOAN_COL_BC - 1])
        set_cell(ws, target_row, COL_Z, record[LOAN_COL_BF - 1])
        set_cell(ws, target_row, COL_AB, record[LOAN_COL_Q - 1])
        set_cell(ws, target_row, COL_AC, record[LOAN_COL_T - 1])
        set_cell(ws, target_row, COL_AD, record[LOAN_COL_U - 1])
        added += 1
    return added


def append_repay_block(ws, template_cache, template_height, template_row_index, rows: Iterable[Sequence], repay_type: str) -> int:
    added = 0
    for record in rows:
        target_row = ws.max_row + 1
        apply_template_row(ws, template_cache, target_row, template_row_index, template_height)

        set_cell(ws, target_row, COL_D, record[REPAY_COL_G - 1])
        set_cell(ws, target_row, COL_E, record[REPAY_COL_H - 1])
        set_cell(ws, target_row, COL_F, record[REPAY_COL_J - 1])
        set_cell(ws, target_row, COL_G, record[REPAY_COL_M - 1])
        set_cell(ws, target_row, COL_H, record[REPAY_COL_B - 1])
        set_cell(ws, target_row, COL_I, record[REPAY_COL_C - 1])
        set_cell(ws, target_row, COL_J, "/")
        set_cell(ws, target_row, COL_K, "/")
        set_cell(ws, target_row, COL_L, record[REPAY_COL_F - 1])

        for col_idx in (COL_N, COL_O, COL_P, COL_Q, COL_R, COL_S, COL_T, COL_U, COL_V, COL_W, COL_X, COL_Z, COL_AB, COL_AC, COL_AD):
            set_cell(ws, target_row, col_idx, None)

        set_cell(ws, target_row, COL_AE, record[REPAY_COL_AE - 1])
        set_cell(ws, target_row, COL_AG, record[REPAY_COL_O - 1])
        set_cell(ws, target_row, COL_AH, record[REPAY_COL_AG - 1])
        set_cell(ws, target_row, COL_AI, record[REPAY_COL_AG - 1])
        set_cell(ws, target_row, COL_AJ, record[REPAY_COL_AE - 1])
        set_cell(ws, target_row, COL_AK, repay_type)
        set_cell(ws, target_row, COL_AR, record[REPAY_COL_AH - 1])
        added += 1
    return added


def merge_ai_with_sum(ws, start_row: int, end_row: int, target_date: dt.date):
    if start_row > end_row:
        return

    def finalize(group: List[int], ar_value):
        if not group:
            return
        total = 0.0
        for r in group:
            val = ws.cell(row=r, column=COL_AH).value
            if isinstance(val, (int, float)):
                total += float(val)
            elif val not in (None, ""):
                try:
                    total += float(val)
                except Exception:
                    pass
        top = group[0]
        set_cell(ws, top, COL_AI, total)
        if len(group) > 1:
            ws.merge_cells(start_row=top, end_row=group[-1], start_column=COL_AI, end_column=COL_AI)

    current_group: List[int] = []
    current_ar = None

    for row_idx in range(start_row, end_row + 1):
        ae_date = normalize_excel_date(ws.cell(row=row_idx, column=COL_AE).value)
        ar_value = ws.cell(row=row_idx, column=COL_AR).value
        if ae_date == target_date and ar_value not in (None, ""):
            if current_ar is None:
                current_ar = ar_value
                current_group = [row_idx]
            elif ar_value == current_ar:
                current_group.append(row_idx)
            else:
                finalize(current_group, current_ar)
                current_ar = ar_value
                current_group = [row_idx]
        else:
            finalize(current_group, current_ar)
            current_group = []
            current_ar = None

    finalize(current_group, current_ar)


def main():
    args = parse_args()
    target_date = parse_input_date(args.date)

    ledger_path = Path(args.ledger).resolve()
    loan_path = Path(args.loan).resolve()
    factoring_path = Path(args.factoring_repay).resolve()
    refactoring_path = Path(args.refactoring_repay).resolve()
    output_path = Path(args.output).resolve()

    wb = load_workbook(ledger_path, data_only=False)
    ws = find_target_sheet(wb)
    template_cache = cache_template_row(ws, TEMPLATE_ROW_INDEX)
    template_height = ws.row_dimensions[TEMPLATE_ROW_INDEX].height

    last_date = get_last_existing_date(ws)
    if last_date and target_date <= last_date:
        # 无需更新，直接输出原文件
        wb.save(output_path)
        print(f"[ledger_daily] {output_path} 已生成（未追加，{target_date} <= {last_date}）")
        return

    append_start_row = find_last_data_row(ws) + 1

    loan_rows = collect_loan_rows(loan_path, target_date)
    factoring_repay_rows = collect_repay_rows(factoring_path, target_date)
    refactoring_repay_rows = collect_repay_rows(refactoring_path, target_date)

    total_added = 0
    total_added += append_loan_block(ws, template_cache, template_height, TEMPLATE_ROW_INDEX, loan_rows)
    total_added += append_repay_block(ws, template_cache, template_height, TEMPLATE_ROW_INDEX, factoring_repay_rows, "保理")
    total_added += append_repay_block(ws, template_cache, template_height, TEMPLATE_ROW_INDEX, refactoring_repay_rows, "再保理")

    if total_added:
        merge_ai_with_sum(ws, append_start_row, ws.max_row, target_date)

    wb.save(output_path)
    print(
        f"[ledger_daily] 完成写入 -> {output_path}，新增 {total_added} 行；"
        f" 放款 {len(loan_rows)} 行，保理还款 {len(factoring_repay_rows)} 行，再保理还款 {len(refactoring_repay_rows)} 行"
    )


if __name__ == "__main__":
    main()
