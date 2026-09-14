from __future__ import annotations

import io
import math
import re
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal
from xml.sax.saxutils import escape, quoteattr


STYLE = {
    "base": 0,
    "title": 1,
    "subtitle": 2,
    "section": 3,
    "header_yellow": 4,
    "data_yellow": 5,
    "header_blue": 6,
    "data_blue": 7,
    "days_blue": 8,
    "time_blue": 9,
    "number_blue": 10,
    "money_blue": 11,
    "percent_blue": 12,
    "header_orange": 13,
    "data_orange": 14,
    "table_header": 15,
    "table_text": 16,
    "table_center": 17,
    "table_days": 18,
    "table_time": 19,
    "table_number": 20,
    "table_money": 21,
    "table_percent": 22,
    "total_text": 23,
    "total_days": 24,
    "total_time": 25,
    "total_number": 26,
    "total_money": 27,
    "total_percent": 28,
    "note": 29,
    "table_id": 30,
    "table_date": 31,
}


def _column_name(index: int) -> str:
    result = ""
    value = index + 1
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _safe_text(value: object) -> str:
    text = str(value)
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", text)


@dataclass(frozen=True)
class Formula:
    expression: str
    cached_value: int | float | Decimal


@dataclass(frozen=True)
class Cell:
    value: object = None
    style: int = STYLE["base"]


@dataclass
class Sheet:
    name: str
    rows: list[list[Cell]] = field(default_factory=list)
    widths: dict[int, float] = field(default_factory=dict)
    row_heights: dict[int, float] = field(default_factory=dict)
    merges: list[str] = field(default_factory=list)
    freeze_rows: int = 0
    auto_filter: str | None = None

    def append(self, values: list[object], styles: list[int] | int = STYLE["base"]) -> int:
        if isinstance(styles, int):
            style_values = [styles] * len(values)
        else:
            style_values = styles
        if len(style_values) != len(values):
            raise ValueError("Each value needs a matching style")
        self.rows.append([Cell(value, style) for value, style in zip(values, style_values)])
        return len(self.rows)


class Workbook:
    def __init__(self, title: str) -> None:
        self.title = title
        self.sheets: list[Sheet] = []

    def add_sheet(self, name: str) -> Sheet:
        safe = re.sub(r"[\\/*?:\[\]]", " ", name).strip()[:31] or "Sheet"
        existing = {sheet.name for sheet in self.sheets}
        base = safe
        suffix = 2
        while safe in existing:
            marker = f" {suffix}"
            safe = base[: 31 - len(marker)] + marker
            suffix += 1
        sheet = Sheet(safe)
        self.sheets.append(sheet)
        return sheet

    def to_bytes(self) -> bytes:
        if not self.sheets:
            raise ValueError("Workbook needs at least one worksheet")
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml", self._content_types())
            archive.writestr("_rels/.rels", self._root_relationships())
            archive.writestr("docProps/core.xml", self._core_properties())
            archive.writestr("docProps/app.xml", self._app_properties())
            archive.writestr("xl/workbook.xml", self._workbook_xml())
            archive.writestr("xl/_rels/workbook.xml.rels", self._workbook_relationships())
            archive.writestr("xl/styles.xml", _styles_xml())
            for index, sheet in enumerate(self.sheets, start=1):
                archive.writestr(f"xl/worksheets/sheet{index}.xml", _sheet_xml(sheet))
        return output.getvalue()

    def _content_types(self) -> str:
        sheets = "".join(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for index in range(1, len(self.sheets) + 1)
        )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            f"{sheets}</Types>"
        )

    @staticmethod
    def _root_relationships() -> str:
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            '</Relationships>'
        )

    def _core_properties(self) -> str:
        created = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:title>{escape(self.title)}</dc:title><dc:creator>Telegram Reporting Bot</dc:creator>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>'
            '</cp:coreProperties>'
        )

    def _app_properties(self) -> str:
        titles = "".join(f"<vt:lpstr>{escape(sheet.name)}</vt:lpstr>" for sheet in self.sheets)
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>Telegram Reporting Bot</Application>'
            f'<TitlesOfParts><vt:vector size="{len(self.sheets)}" baseType="lpstr">{titles}</vt:vector></TitlesOfParts>'
            '</Properties>'
        )

    def _workbook_xml(self) -> str:
        sheets = "".join(
            f'<sheet name={quoteattr(sheet.name)} sheetId="{index}" r:id="rId{index}"/>'
            for index, sheet in enumerate(self.sheets, start=1)
        )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<sheets>{sheets}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>'
        )

    def _workbook_relationships(self) -> str:
        sheets = "".join(
            f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>'
            for index in range(1, len(self.sheets) + 1)
        )
        style_id = len(self.sheets) + 1
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'{sheets}<Relationship Id="rId{style_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>'
        )


def _cell_xml(row: int, column: int, cell: Cell) -> str:
    reference = f"{_column_name(column)}{row}"
    style = f' s="{cell.style}"' if cell.style else ""
    value = cell.value
    if value is None:
        return f'<c r="{reference}"{style}/>' if cell.style else ""
    if isinstance(value, Formula):
        return f'<c r="{reference}"{style}><f>{escape(value.expression)}</f><v>{value.cached_value}</v></c>'
    if isinstance(value, date) and cell.style == STYLE["table_date"]:
        day = value.date() if isinstance(value, datetime) else value
        value = (day - date(1899, 12, 30)).days
    if isinstance(value, bool):
        return f'<c r="{reference}"{style} t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, timedelta):
        value = value.total_seconds() / 86400
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        number = float(value) if isinstance(value, Decimal) else value
        if isinstance(number, float) and not math.isfinite(number):
            value = "—"
        else:
            return f'<c r="{reference}"{style}><v>{number}</v></c>'
    if isinstance(value, (date, datetime)):
        value = value.strftime("%d.%m.%Y %H:%M") if isinstance(value, datetime) else value.strftime("%d.%m.%Y")
    text = escape(_safe_text(value))
    preserve = ' xml:space="preserve"' if text != text.strip() or "\n" in text else ""
    return f'<c r="{reference}"{style} t="inlineStr"><is><t{preserve}>{text}</t></is></c>'


def _sheet_xml(sheet: Sheet) -> str:
    max_columns = max((len(row) for row in sheet.rows), default=1)
    max_rows = max(len(sheet.rows), 1)
    dimension = f"A1:{_column_name(max_columns - 1)}{max_rows}"
    cols = "".join(
        f'<col min="{index + 1}" max="{index + 1}" width="{width}" customWidth="1"/>'
        for index, width in sorted(sheet.widths.items())
    )
    row_parts = []
    for row_index, row in enumerate(sheet.rows, start=1):
        height = sheet.row_heights.get(row_index)
        height_attr = f' ht="{height}" customHeight="1"' if height else ""
        cells = "".join(_cell_xml(row_index, column, cell) for column, cell in enumerate(row))
        row_parts.append(f'<row r="{row_index}"{height_attr}>{cells}</row>')
    pane = ""
    if sheet.freeze_rows:
        top_left = f"A{sheet.freeze_rows + 1}"
        pane = (
            f'<pane ySplit="{sheet.freeze_rows}" topLeftCell="{top_left}" '
            'activePane="bottomLeft" state="frozen"/>'
        )
    auto_filter = f'<autoFilter ref={quoteattr(sheet.auto_filter)}/>' if sheet.auto_filter else ""
    merges = ""
    if sheet.merges:
        items = "".join(f'<mergeCell ref={quoteattr(value)}/>' for value in sheet.merges)
        merges = f'<mergeCells count="{len(sheet.merges)}">{items}</mergeCells>'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="{dimension}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0">{pane}</sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f'<cols>{cols}</cols><sheetData>{"".join(row_parts)}</sheetData>{auto_filter}{merges}'
        '<pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>'
        '<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>'
        '</worksheet>'
    )


def _styles_xml() -> str:
    fonts = (
        '<fonts count="4">'
        '<font><sz val="10"/><name val="Arial"/><family val="2"/></font>'
        '<font><b/><sz val="10"/><name val="Arial"/><family val="2"/></font>'
        '<font><b/><color rgb="FFFFFFFF"/><sz val="14"/><name val="Arial"/><family val="2"/></font>'
        '<font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Arial"/><family val="2"/></font>'
        '</fonts>'
    )
    fills = (
        '<fills count="7">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD0E0E3"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFCE5CD"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFE7E6E6"/><bgColor indexed="64"/></patternFill></fill>'
        '</fills>'
    )
    borders = (
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
        '<border><left style="thin"><color rgb="FFD9E1F2"/></left><right style="thin"><color rgb="FFD9E1F2"/></right>'
        '<top style="thin"><color rgb="FFD9E1F2"/></top><bottom style="thin"><color rgb="FFD9E1F2"/></bottom><diagonal/></border></borders>'
    )
    definitions = [
        (0, 0, 0, 0, "left", False),
        (0, 2, 2, 0, "center", False),
        (0, 1, 0, 0, "center", False),
        (0, 1, 0, 0, "left", False),
        (0, 1, 3, 1, "center", True),
        (0, 0, 3, 1, "center", False),
        (0, 1, 4, 1, "center", True),
        (0, 0, 4, 1, "center", False),
        (164, 0, 4, 1, "right", False),
        (166, 0, 4, 1, "right", False),
        (1, 0, 4, 1, "right", False),
        (167, 0, 4, 1, "right", False),
        (165, 0, 4, 1, "right", False),
        (0, 1, 5, 1, "center", True),
        (0, 0, 5, 1, "center", True),
        (0, 3, 2, 1, "center", True),
        (0, 0, 0, 1, "left", True),
        (0, 0, 0, 1, "center", True),
        (164, 0, 0, 1, "right", False),
        (166, 0, 0, 1, "right", False),
        (1, 0, 0, 1, "right", False),
        (167, 0, 0, 1, "right", False),
        (165, 0, 0, 1, "right", False),
        (0, 1, 6, 1, "left", True),
        (164, 1, 6, 1, "right", False),
        (166, 1, 6, 1, "right", False),
        (1, 1, 6, 1, "right", False),
        (167, 1, 6, 1, "right", False),
        (165, 1, 6, 1, "right", False),
        (0, 0, 0, 0, "left", True),
        (49, 0, 0, 1, "left", True),
        (168, 0, 0, 1, "center", False),
    ]
    xfs = []
    for num_fmt, font, fill, border, alignment, wrap in definitions:
        alignment_xml = f'<alignment horizontal="{alignment}" vertical="center" wrapText="{1 if wrap else 0}"/>'
        xfs.append(
            f'<xf numFmtId="{num_fmt}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0" '
            f'applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1">{alignment_xml}</xf>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<numFmts count="5"><numFmt numFmtId="164" formatCode="0.00"/>'
        '<numFmt numFmtId="165" formatCode="0.00%"/><numFmt numFmtId="166" formatCode="[h]:mm:ss"/>'
        '<numFmt numFmtId="167" formatCode="#,##0.00"/><numFmt numFmtId="168" formatCode="dd.mm.yyyy"/></numFmts>'
        f'{fonts}{fills}{borders}<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">{"".join(xfs)}</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )
