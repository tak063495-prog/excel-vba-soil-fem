"""Read-only XLSM settings/binding inventory; never rewrites the VBA package."""
import argparse
import hashlib
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile

PAGES = ("メッシュ設定", "解析設定", "高速化設定", "出力・表示", "詳細設定")
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def excel_text(value):
    return re.sub(r"_x([0-9a-fA-F]{4})_", lambda m: chr(int(m[1], 16)), value)


def read_book(path):
    with ZipFile(path) as archive:
        assert "xl/vbaProject.bin" in archive.namelist(), "VBA project missing"
        strings = [excel_text("".join(t.text or "" for t in si.findall(".//m:t", NS)))
                   for si in ET.fromstring(archive.read("xl/sharedStrings.xml")).findall("m:si", NS)]
        relations = {r.attrib["Id"]: r.attrib["Target"] for r in
                     ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))}
        sheets = {}
        for sheet in ET.fromstring(archive.read("xl/workbook.xml")).findall("m:sheets/m:sheet", NS):
            target = relations[sheet.attrib[REL_ID]]
            target = target.lstrip("/") if target.startswith("/") else "xl/" + target
            cells, formulas = {}, {}
            for cell in ET.fromstring(archive.read(target)).findall("m:sheetData/m:row/m:c", NS):
                address, kind = cell.attrib["r"], cell.attrib.get("t")
                raw = cell.find("m:v", NS)
                value = raw.text if raw is not None else None
                if kind == "s":
                    value = strings[int(value)]
                elif kind == "inlineStr":
                    value = excel_text("".join(t.text or "" for t in cell.findall(".//m:t", NS)))
                elif kind == "b":
                    value = value == "1"
                elif kind == "str":
                    value = excel_text(value or "")
                elif kind != "e" and value is not None:
                    value = float(value)
                cells[address] = value
                formula = cell.find("m:f", NS)
                if formula is not None:
                    formulas[address] = formula.text or ""
            sheets[sheet.attrib["name"]] = (cells, formulas)
        return sheets


def inventory(path):
    sheets = read_book(path)
    legacy, formulas = sheets["設定"]
    registry = {str(value): int(address[1:]) for address, value in legacy.items()
                if address.startswith("E") and int(address[1:]) >= 5 and value}
    records, seen = [], set()
    for page in PAGES:
        cells, _ = sheets[page]
        rows = sorted(int(a[1:]) for a, v in cells.items() if a.startswith("F") and v)
        for row in rows:
            key = cells[f"F{row}"]
            assert key not in seen, f"duplicate canonical key: {key}"
            seen.add(key)
            assert key in registry, f"missing legacy key: {key}"
            old_row = registry[key]
            normalized = formulas[f"C{old_row}"].replace("'", "")
            expected = f'IF({page}!C{row}="","",{page}!C{row})'
            assert normalized == expected, f"wrong binding: {key}: {normalized}"
            assert cells.get(f"G{row}") is not None, f"missing recommendation: {key}"
            records.append(dict(page=page, cell=f"C{row}", key=key,
                                value=cells.get(f"C{row}"), label=cells.get(f"B{row}"),
                                note=cells.get(f"E{row}"), recommendation=cells.get(f"G{row}"),
                                legacy_row=old_row))
    return dict(status="PASS", book=path.name, sha256=hashlib.sha256(path.read_bytes()).hexdigest(),
                count=len(records), settings=records,
                internal_keys=sorted(set(registry) - seen))


def markdown(data):
    def cell(value):
        if value is None or value == "":
            return "（空欄）"
        if isinstance(value, float):
            return format(value, ".15g")
        return str(value).replace("|", " / ").replace("\r", "").replace("\n", "<br>")
    lines = ["# UI版の設定一覧（2026-10-20）", "",
             "配布XLSMから読み取った初期入力値とG列の推奨・目安です。材料・荷重・支持・Fs等の条件を一括変更するものではありません。",
             "設定値の変更はC列だけで行います。再表示・再読込はユーザーの設定値を保持します。",
             "[操作・対応機能・検証記録](WORKBOOK_UI_20261020.md)。自動出力は編集しません。", ""]
    for page in PAGES:
        lines += ["## " + page, "", "| 入力セル | 項目 / KEY | 配布保存値 | 推奨・目安 |",
                  "| --- | --- | --- | --- |"]
        for row in data["settings"]:
            if row["page"] == page:
                value = "（自動出力）" if row["key"].startswith("MESH_GEN_LAST_") else cell(row["value"])
                lines.append(f'| {row["cell"]} | {cell(row["label"])}<br>`{row["key"]}` | {value} | {cell(row["recommendation"])} |')
        lines += [""]
    return "\n".join(lines)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("book", type=Path)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--markdown", type=Path)
    args = parser.parse_args()
    data = inventory(args.book)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf8")
    if args.markdown:
        args.markdown.write_text(markdown(data), encoding="utf8")
    print(f'PASS {data["count"]} canonical settings, exact legacy bindings, VBA package present')
