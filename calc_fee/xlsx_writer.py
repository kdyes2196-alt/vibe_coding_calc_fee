"""계산파일(.xlsx)의 셀에 값을 직접 써 넣는다.

계산파일은 한컴오피스 한셀로 저장되어 openpyxl 로는 열리지 않는다(스타일 정의가 표준과 다름).
그래서 xlsx 압축 안의 시트 XML 만 직접 고친다. 서식, 메모, 숨김 시트 등 나머지는 그대로 유지된다.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import PurePosixPath

from lxml import etree

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


@dataclass(frozen=True)
class Formula:
    """셀에 넣을 수식 ('=' 없이, 예: 'SUM(C4:E4)')."""

    text: str


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def split_ref(ref: str) -> tuple[str, int]:
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    if not m:
        raise ValueError(f"잘못된 셀 주소: {ref}")
    return m.group(1), int(m.group(2))


def _q(tag: str) -> str:
    return f"{{{NS_MAIN}}}{tag}"


class XlsxPatcher:
    """xlsx 파일 안의 시트 XML 을 읽고 고쳐서 새 파일로 저장한다."""

    def __init__(self, path: str):
        self.path = path
        self._zip = zipfile.ZipFile(path)
        self._modified: dict[str, etree._Element] = {}
        self._sheet_paths = self._read_sheet_paths()

    # ---- 시트 찾기 -------------------------------------------------------
    def _read_sheet_paths(self) -> dict[str, str]:
        wb = etree.fromstring(self._zip.read("xl/workbook.xml"))
        rels = etree.fromstring(self._zip.read("xl/_rels/workbook.xml.rels"))
        targets = {r.get("Id"): r.get("Target") for r in rels.iter(f"{{{NS_PKG_REL}}}Relationship")}
        paths = {}
        for sheet in wb.iter(_q("sheet")):
            target = targets[sheet.get(f"{{{NS_REL}}}id")]
            target = target.lstrip("/") if target.startswith("/") else str(PurePosixPath("xl") / target)
            paths[sheet.get("name")] = target
        return paths

    @property
    def sheet_names(self) -> list[str]:
        return list(self._sheet_paths)

    def _tree(self, part: str) -> etree._Element:
        if part not in self._modified:
            self._modified[part] = etree.fromstring(self._zip.read(part))
        return self._modified[part]

    def sheet(self, name: str) -> etree._Element:
        if name not in self._sheet_paths:
            raise KeyError(f"계산파일에 '{name}' 시트가 없습니다. 있는 시트: {self.sheet_names}")
        return self._tree(self._sheet_paths[name])

    # ---- 셀 읽기/쓰기 ----------------------------------------------------
    def _shared_strings(self) -> list[str]:
        if not hasattr(self, "_sst"):
            try:
                root = etree.fromstring(self._zip.read("xl/sharedStrings.xml"))
            except KeyError:
                self._sst = []
            else:
                self._sst = ["".join(si.itertext()) for si in root.iter(_q("si"))]
        return self._sst

    def read_column(self, sheet_name: str, col: str, rows: range) -> dict[int, str]:
        """열의 값(문자열)을 행 번호별로 읽는다. 기관코드 열을 읽을 때 쓴다."""
        data = self.sheet(sheet_name).find(_q("sheetData"))
        wanted = {f"{col}{r}" for r in rows}
        out = {}
        for c in data.iter(_q("c")):
            ref = c.get("r")
            if ref not in wanted:
                continue
            v = c.find(_q("v"))
            is_ = c.find(_q("is"))
            if c.get("t") == "s" and v is not None:
                text = self._shared_strings()[int(v.text)]
            elif is_ is not None:
                text = "".join(is_.itertext())
            elif v is not None:
                text = v.text
            else:
                continue
            out[split_ref(ref)[1]] = text.strip()
        return out

    def _get_cell(self, sheet_data: etree._Element, ref: str) -> etree._Element:
        col, row_num = split_ref(ref)
        row = None
        for r in sheet_data.iterchildren(_q("row")):
            n = int(r.get("r"))
            if n == row_num:
                row = r
                break
            if n > row_num:
                row = etree.Element(_q("row"), r=str(row_num))
                r.addprevious(row)
                break
        if row is None:
            row = etree.SubElement(sheet_data, _q("row"), r=str(row_num))
        for c in row.iterchildren(_q("c")):
            if c.get("r") == ref:
                return c
            if col_to_num(split_ref(c.get("r"))[0]) > col_to_num(col):
                new = etree.Element(_q("c"), r=ref)
                c.addprevious(new)
                return new
        return etree.SubElement(row, _q("c"), r=ref)

    def set_cells(self, sheet_name: str, values: dict[str, int | float | Formula]) -> None:
        """셀 서식(s 속성)은 그대로 두고 값 또는 수식만 바꾼다."""
        sheet_data = self.sheet(sheet_name).find(_q("sheetData"))
        for ref, value in values.items():
            c = self._get_cell(sheet_data, ref)
            for child in list(c):
                c.remove(child)
            c.attrib.pop("t", None)
            if isinstance(value, Formula):
                etree.SubElement(c, _q("f")).text = value.text
            else:
                etree.SubElement(c, _q("v")).text = repr(value) if isinstance(value, float) else str(int(value))

    # ---- 재계산 표시 -----------------------------------------------------
    def request_full_recalc(self) -> None:
        """파일을 열 때 모든 수식을 다시 계산하도록 표시하고, 예전 계산 결과는 지운다."""
        wb = self._tree("xl/workbook.xml")
        calc = wb.find(_q("calcPr"))
        if calc is None:
            calc = etree.SubElement(wb, _q("calcPr"))
        calc.set("fullCalcOnLoad", "1")
        for part in self._sheet_paths.values():
            for c in self._tree(part).iter(_q("c")):
                if c.find(_q("f")) is not None:
                    v = c.find(_q("v"))
                    if v is not None:
                        c.remove(v)
                    c.attrib.pop("t", None)

    # ---- 저장 ------------------------------------------------------------
    def save(self, out_path: str) -> None:
        if out_path == self.path:
            raise ValueError("원본 계산파일을 덮어쓸 수 없습니다. 다른 출력 경로를 지정하세요.")
        tmp = out_path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for info in self._zip.infolist():
                if info.filename in self._modified:
                    data = etree.tostring(
                        self._modified[info.filename], xml_declaration=True, encoding="UTF-8", standalone=True
                    )
                else:
                    data = self._zip.read(info.filename)
                out.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)
        shutil.move(tmp, out_path)

    def close(self) -> None:
        self._zip.close()
