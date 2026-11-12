# -*- coding: utf-8 -*-
"""一次性 fixture 生成器：读用户附件 xlsx（只读），产出 fixture.json。

fixture.json 是 500+ 用例的独立 oracle（与 GenOffice 代码无关，openpyxl 直读）：
- cells:   address -> {v?, f?}  全部非空格（含 F 列 =EVALUATE(...) 公式）
- cached:  address -> 缓存值（data_only，Excel 保存时的计算结果，导出测试的 oracle）
- sheet:   元信息
- colStats: 每列 nonEmpty / distinctTop / 数值文本数（按仓库 textLooksNumericOrDate 同规则粗算）
- sectionRows: unit=项 的行号（section 伪表头）
- functions: E 列文本公式里出现的全部函数名
"""
import json
import re
import sys

import openpyxl

SRC = "/home/rebootman/.cloudcli/assets/1794317649011-465612238-2000_______________.xlsx"
OUT = "/home/rebootman/claudeproject1-genoffice-fork/apps/sheets/tests/bug-hunt/fixture.json"


def looks_numeric_or_date(text: str) -> bool:
    """粗版规则：纯数字/日期样式的文本。测试里用仓库自己的 textLooksNumericOrDate 复算。"""
    t = text.strip().replace(",", "")
    if not t:
        return False
    try:
        float(t)
        return True
    except ValueError:
        pass
    return bool(re.match(r"^\d{1,4}[/.-]\d{1,2}[/.-]\d{1,4}$", t))


def main() -> int:
    wb = openpyxl.load_workbook(SRC, data_only=False)
    wbv = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb.active
    wsv = wbv.active
    rows, cols = ws.max_row, ws.max_column
    sheet = {"name": ws.title, "rows": rows, "cols": cols}

    cells = {}
    cached = {}
    for row in ws.iter_rows(min_row=1, max_row=rows, max_col=cols):
        for c in row:
            v = c.value
            if v is None:
                continue
            addr = f"{c.column_letter}{c.row}"
            entry = {}
            if c.data_type == "f":
                # 公式格：value 即公式文本（openpyxl 带 '='），不另记 v——
                # GenOffice 模型里公式格 value=null，显示值来自 cached
                formula = str(v)
                entry["f"] = formula if formula.startswith("=") else "=" + formula
            else:
                # E 列"文本公式" data_type='s'（字符串 =...），记 v——
                # GenOffice 读出来就是纯文本
                entry["v"] = v if isinstance(v, (int, float)) else str(v)
            cells[addr] = entry
            cv = wsv.cell(c.row, c.column).value
            if cv is not None:
                cached[addr] = cv if isinstance(cv, (int, float)) else str(cv)

    # 每列统计
    col_stats = {}
    for ci in range(1, cols + 1):
        letter = openpyxl.utils.get_column_letter(ci)
        values = []
        numeric_text = 0
        for ri in range(1, rows + 1):
            v = ws.cell(ri, ci).value
            if v is None:
                continue
            if isinstance(v, str) and v.startswith("="):
                continue  # 公式文本列（E 列存的是公式文本）不算显示值
            s = v if isinstance(v, str) else str(v)
            if isinstance(v, str):
                if looks_numeric_or_date(s):
                    numeric_text += 1
            values.append(s)
        top = {}
        for s in values:
            top[s] = top.get(s, 0) + 1
        top_items = sorted(top.items(), key=lambda kv: (-kv[1], kv[0]))[:50]
        col_stats[letter] = {
            "nonEmpty": len(values),
            "distinct": len(top),
            "distinctTop": top_items,
            "numericText": numeric_text,
        }

    section_rows = [
        ri for ri in range(1, rows + 1) if ws.cell(ri, 2).value == "项"
    ]

    funcs = set()
    for ri in range(1, rows + 1):
        v = ws.cell(ri, 5).value
        if isinstance(v, str):
            funcs.update(re.findall(r"([A-Z][A-Z0-9_.]{1,20})\s*\(", v.upper()))

    fixture = {
        "sheet": sheet,
        "cells": cells,
        "cached": cached,
        "colStats": col_stats,
        "sectionRows": section_rows,
        "functions": sorted(funcs),
    }
    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump(fixture, fh, ensure_ascii=False)
    print(
        f"rows={rows} cols={cols} cells={len(cells)} cached={len(cached)} "
        f"sectionRows={len(section_rows)} functions={sorted(funcs)}"
    )
    for k, s in col_stats.items():
        print(
            f"col {k}: nonEmpty={s['nonEmpty']} distinct={s['distinct']} "
            f"numericText={s['numericText']} top={[x[0][:18] for x in s['distinctTop'][:6]]}"
        )
    return 0


if __name__ == "__main__":
    sys.exit(main())
