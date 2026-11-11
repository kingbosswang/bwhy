# -*- coding: utf-8 -*-
"""把「不玩黄油游戏目录.xlsx」转成网页用的 games.js。

输出格式：window.GAME_DATA = {updated:'YYYY-MM-DD', items:[[序号, 名字, 备注], ...]}
分类与批次都能从序号推导：1=3D 2=SLG 4=galgame 5=RPG 6=安卓；纯数字=老资源。
"""
import datetime
import json
import re

import openpyxl
from openpyxl.utils import get_column_letter

SRC = r"C:\Users\PC\Desktop\目录\不玩黄油游戏目录.xlsx"
OUT = r"C:\Users\PC\Desktop\bwhy-web\games.js"

SHEETS = [
    "❤️电脑  3D  1xxx",
    "❤️电脑  互动SLG  2xxx",
    "❤️电脑  galgame 4xxx",
    "❤️电脑  RPG 5xxx",
    "❤️安卓6xxx",
]


def clean(v):
    if v is None:
        return ""
    s = str(v)
    s = s.replace("\xa0", " ").replace("\u3000", " ")
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)
    return s.strip()


def section_of(ws, row, col):
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return clean(ws.cell(rng.min_row, rng.min_col).value)
    return clean(ws.cell(row, col).value)


def is_placeholder(name):
    return bool(re.fullmatch(r"[xX×\s\.\-_?？]+", name))


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    rows = []
    skipped = 0
    for sheet in SHEETS:
        ws = wb[sheet]
        starts = [c for c in range(1, ws.max_column + 1)
                  if clean(ws.cell(3, c).value) == "序号"]
        for c in starts:
            for r in range(4, ws.max_row + 1):
                code = clean(ws.cell(r, c).value)
                name = clean(ws.cell(r, c + 1).value)
                note = clean(ws.cell(r, c + 2).value)
                if not name:
                    continue
                if is_placeholder(name):
                    skipped += 1
                    continue
                rows.append([code, name, note])

    # 去重
    seen, items = set(), []
    for it in rows:
        key = (it[0], it[1])
        if key in seen:
            continue
        seen.add(key)
        items.append(it)

    data = {
        "updated": datetime.date.today().isoformat(),
        "items": items,
    }
    text = "window.GAME_DATA=" + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";\n"
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)

    print("entries:", len(items), "skipped placeholders:", skipped)
    print("file bytes:", len(text.encode("utf-8")))


if __name__ == "__main__":
    main()
