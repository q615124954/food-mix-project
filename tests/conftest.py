from __future__ import annotations

from openpyxl import Workbook

from foodmix.numbers import excel_serial

SAMPLE_DATE = "2026-08-01"


def make_material(path, *, egg_unit="个"):
    """最小矩阵表：品名在 D 列，日期列是数量，金额在数量列后两格。"""
    workbook = Workbook()
    daily = workbook.active
    daily.title = "日物料采购"
    daily["A1"] = "序号"
    daily["B1"] = "物料代码"
    daily["C1"] = "类别"
    daily["D1"] = "品名"
    daily["E1"] = "单位"
    daily["F1"] = "合计"
    serial = excel_serial(SAMPLE_DATE)
    daily.cell(1, 9, serial)
    daily.merge_cells(start_row=1, start_column=9, end_row=1, end_column=11)
    for col, label in ((6, "数量"), (7, "单价"), (8, "金额"), (9, "数量"), (10, "单价"), (11, "金额")):
        daily.cell(2, col, label)
    daily["B3"] = "V-葱"
    daily["D3"] = "小葱"
    daily["E3"] = "斤"
    daily["F3"] = "KEEP"
    daily["H3"] = "KEEPAMT"
    daily.cell(3, 9, 100)
    daily.cell(3, 11, 999)
    daily["B4"] = "V-蛋"
    daily["D4"] = "鸡蛋"
    daily["E4"] = egg_unit
    daily["B5"] = "V-袋"
    daily["D5"] = "煲汤袋"
    daily["E5"] = "包"
    daily.cell(5, 9, 77)

    meat = workbook.create_sheet("肉类采购")
    meat["A1"] = "肉类采购订货汇总"
    meat["A2"] = "序号"
    meat["B2"] = "物料代码"
    meat["D2"] = "品名"
    meat["E2"] = "单位"
    meat["F2"] = "合计"
    meat.cell(2, 9, serial)
    meat.merge_cells(start_row=2, start_column=9, end_row=2, end_column=11)
    for col, label in ((9, "数量"), (10, "单价"), (11, "金额")):
        meat.cell(3, col, label)
    meat["B4"] = "M-1"
    meat["D4"] = "五花肉"
    meat["E4"] = "斤"

    staff = workbook.create_sheet("员工餐")
    staff["A1"] = "员工餐"
    staff["B3"] = serial
    staff["D3"] = 12345
    staff["H3"] = 888
    workbook.save(path)
    return path
