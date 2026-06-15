from __future__ import annotations

import csv
import io
import json
import os
import re
import sqlite3
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from datetime import datetime
from html import escape as html_escape
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape as xml_escape


ROOT = Path(__file__).resolve().parents[1]
DB_PATH = Path(os.getenv("DEMO_DB_PATH") or ("/tmp/demo.sqlite3" if os.getenv("VERCEL") else ROOT / "storage" / "demo.sqlite3"))
SAMPLES_DIR = ROOT / "samples"


def load_env_file() -> None:
    env_path = ROOT / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8", errors="ignore").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        key, value = stripped.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def init_db() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            create table if not exists projects (
                id integer primary key autoincrement,
                name text not null,
                created_at text not null,
                total real not null default 0,
                payload_json text not null
            )
            """
        )
        conn.commit()


def read_json(handler: BaseHTTPRequestHandler) -> dict:
    length = int(handler.headers.get("Content-Length", "0"))
    raw = handler.rfile.read(length) if length else b"{}"
    if not raw:
        return {}
    return json.loads(raw.decode("utf-8"))


def send_json(handler: BaseHTTPRequestHandler, data: dict | list, status: int = 200) -> None:
    body = json.dumps(data, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def safe_sample_path(url_path: str) -> Path | None:
    name = Path(urllib.parse.unquote(url_path).replace("/samples/", "")).name
    candidate = SAMPLES_DIR / name
    if candidate.exists() and candidate.is_file():
        return candidate
    return None


def normalize_header(value: str) -> str:
    text = str(value or "").strip().lower().replace("ё", "е")
    text = re.sub(r"[\s._/-]+", "", text)
    return text


HEADER_ALIASES = {
    "spec": {
        "section": {"section", "раздел", "система", "группа"},
        "item": {"item", "поз", "пп", "номер", "n", "no"},
        "description": {"description", "name", "наименование", "материал", "оборудование", "номенклатура", "описание"},
        "unit": {"unit", "ед", "едизм", "единица", "единицаизмерения"},
        "quantity": {"quantity", "qty", "кол", "колво", "количество", "объем"},
        "comment": {"comment", "комментарий", "примечание"},
    },
    "price": {
        "sku": {"sku", "артикул", "код", "кодтовара", "номеркод"},
        "name": {"name", "description", "наименование", "товар", "номенклатура", "описание"},
        "category": {"category", "категория", "группа", "раздел"},
        "unit": {"unit", "ед", "едизм", "единица", "единицаизмерения"},
        "price_rub": {"price", "price_rub", "ценаруб", "цена", "стоимость", "розница", "розничнаяцена"},
        "manufacturer": {"manufacturer", "производитель", "бренд", "завод"},
        "lead_time_days": {"leadtime", "срок", "срокпоставки", "поставка"},
    },
}


def canonical_field(header: str, kind: str) -> str | None:
    normalized = normalize_header(header)
    for field, aliases in HEADER_ALIASES[kind].items():
        if normalized in aliases:
            return field
    return None


def clean_number(value) -> str:
    text = str(value or "").strip()
    text = re.sub(r"[^\d,.\-]", "", text)
    if "," in text and "." not in text:
        text = text.replace(",", ".")
    return text


def canonicalize_rows(raw_rows: list[dict], kind: str) -> tuple[list[dict], str]:
    if not raw_rows:
        return [], "Файл пустой или не содержит строк."
    field_map = {}
    for header in raw_rows[0].keys():
        field = canonical_field(header, kind)
        if field and field not in field_map:
            field_map[field] = header
    required = ["description", "unit", "quantity"] if kind == "spec" else ["name", "unit", "price_rub"]
    missing = [field for field in required if field not in field_map]
    if missing:
        return [], "Не удалось найти колонки: " + ", ".join(missing)

    rows = []
    for index, raw in enumerate(raw_rows, start=1):
        row = {field: str(raw.get(source, "")).strip() for field, source in field_map.items()}
        if kind == "spec":
            row.setdefault("section", "")
            row.setdefault("item", str(index))
            row.setdefault("comment", "")
            row["quantity"] = clean_number(row.get("quantity"))
            if row.get("description"):
                rows.append(row)
        else:
            row.setdefault("sku", "")
            row.setdefault("category", "")
            row.setdefault("manufacturer", "")
            row.setdefault("lead_time_days", "")
            row["price_rub"] = clean_number(row.get("price_rub"))
            if row.get("name"):
                rows.append(row)
    if not rows:
        return [], "После чтения файла не осталось строк с наименованиями."
    return rows, ""


def parse_csv_bytes(content: bytes) -> list[dict]:
    text = content.decode("utf-8-sig", errors="replace")
    first_line = next((line for line in text.splitlines() if line.strip()), "")
    delimiter = ";" if first_line.count(";") > first_line.count(",") else ","
    reader = csv.DictReader(io.StringIO(text), delimiter=delimiter)
    return [{key or "": value or "" for key, value in row.items()} for row in reader]


def xlsx_col_index(cell_ref: str) -> int:
    letters = re.match(r"[A-Z]+", cell_ref or "")
    if not letters:
        return 0
    result = 0
    for char in letters.group(0):
        result = result * 26 + ord(char) - 64
    return result - 1


def parse_xlsx_bytes(content: bytes) -> list[dict]:
    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        shared = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", ns):
                text = "".join(node.text or "" for node in item.findall(".//a:t", ns))
                shared.append(text)
        sheet_name = "xl/worksheets/sheet1.xml"
        root = ET.fromstring(archive.read(sheet_name))
        rows = []
        for row_node in root.findall(".//a:sheetData/a:row", ns):
            values = []
            for cell in row_node.findall("a:c", ns):
                idx = xlsx_col_index(cell.attrib.get("r", ""))
                while len(values) <= idx:
                    values.append("")
                cell_type = cell.attrib.get("t")
                if cell_type == "s":
                    raw = cell.findtext("a:v", default="", namespaces=ns)
                    values[idx] = shared[int(raw)] if raw.isdigit() and int(raw) < len(shared) else ""
                elif cell_type == "inlineStr":
                    values[idx] = "".join(node.text or "" for node in cell.findall(".//a:t", ns))
                else:
                    values[idx] = cell.findtext("a:v", default="", namespaces=ns)
            if any(str(value).strip() for value in values):
                rows.append(values)
    if not rows:
        return []
    headers = [str(value).strip() for value in rows[0]]
    result = []
    for values in rows[1:]:
        result.append({headers[index] if index < len(headers) else f"column_{index}": value for index, value in enumerate(values)})
    return result


def parse_multipart_file(handler: BaseHTTPRequestHandler) -> tuple[str, bytes]:
    content_type = handler.headers.get("Content-Type", "")
    match = re.search(r"boundary=(.+)", content_type)
    if not match:
        raise ValueError("Не найден boundary multipart-запроса.")
    boundary = ("--" + match.group(1).strip().strip('"')).encode()
    length = int(handler.headers.get("Content-Length", "0"))
    body = handler.rfile.read(length)
    for part in body.split(boundary):
        if b"Content-Disposition" not in part or b"filename=" not in part:
            continue
        head, _, file_body = part.partition(b"\r\n\r\n")
        filename_match = re.search(rb'filename="([^"]*)"', head)
        filename = filename_match.group(1).decode("utf-8", errors="replace") if filename_match else "upload"
        return filename, file_body.rstrip(b"\r\n-")
    raise ValueError("Файл не найден в запросе.")


def col_name(index: int) -> str:
    result = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def xlsx_cell(value, row_idx: int, col_idx: int) -> str:
    ref = f"{col_name(col_idx)}{row_idx}"
    if value is None:
        return f'<c r="{ref}"/>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"><v>{value}</v></c>'
    text = xml_escape(str(value))
    return f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>'


def build_xlsx(rows: list[list]) -> bytes:
    sheet_rows = []
    for row_idx, row in enumerate(rows, start=1):
        cells = "".join(xlsx_cell(value, row_idx, col_idx) for col_idx, value in enumerate(row))
        sheet_rows.append(f'<row r="{row_idx}">{cells}</row>')

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="18"/>'
        '<cols><col min="1" max="12" width="20" customWidth="1"/></cols>'
        f'<sheetData>{"".join(sheet_rows)}</sheetData>'
        '</worksheet>'
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="КП" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>'
    )
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    )
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '</styleSheet>'
    )

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml)
        archive.writestr("_rels/.rels", rels_xml)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        archive.writestr("xl/styles.xml", styles_xml)
    return buffer.getvalue()


def export_rows(payload: dict) -> list[list]:
    rows = [
        [
            "Раздел",
            "Позиция",
            "Исходное наименование",
            "Ед.",
            "Кол-во",
            "Артикул",
            "Позиция поставщика",
            "Производитель",
            "Цена",
            "Сумма",
            "Уверенность",
            "Статус",
        ]
    ]
    for item in payload.get("matches", []):
        quantity = float(item.get("quantity") or 0)
        price = float(item.get("price") or 0)
        rows.append(
            [
                item.get("section", ""),
                item.get("item", ""),
                item.get("description", ""),
                item.get("unit", ""),
                quantity,
                item.get("sku", ""),
                item.get("matchedName", ""),
                item.get("manufacturer", ""),
                price,
                round(quantity * price, 2),
                item.get("confidence", ""),
                item.get("status", ""),
            ]
        )
    rows.append(["", "", "", "", "", "", "", "Итого", "", round(float(payload.get("total") or 0), 2), "", ""])
    return rows


def build_csv(rows: list[list]) -> bytes:
    buffer = io.StringIO()
    writer = csv.writer(buffer)
    writer.writerows(rows)
    return ("\ufeff" + buffer.getvalue()).encode("utf-8")


def rub(value: float) -> str:
    return f"{value:,.2f}".replace(",", " ").replace(".", ",")


def docx_text(value: str) -> str:
    return xml_escape(str(value)).replace("\n", " ")


def docx_run(text: str, bold: bool = False, size: int = 22, color: str | None = None) -> str:
    props = [f'<w:sz w:val="{size}"/>', f'<w:szCs w:val="{size}"/>']
    if bold:
        props.append("<w:b/><w:bCs/>")
    if color:
        props.append(f'<w:color w:val="{color}"/>')
    return f"<w:r><w:rPr>{''.join(props)}</w:rPr><w:t>{docx_text(text)}</w:t></w:r>"


def docx_paragraph(text: str, bold: bool = False, size: int = 22, color: str | None = None, align: str | None = None) -> str:
    ppr = f'<w:pPr><w:jc w:val="{align}"/></w:pPr>' if align else ""
    return f"<w:p>{ppr}{docx_run(text, bold=bold, size=size, color=color)}</w:p>"


def docx_cell(text: str, width: int, bold: bool = False, fill: str | None = None, align: str | None = None) -> str:
    shading = f'<w:shd w:fill="{fill}"/>' if fill else ""
    return (
        "<w:tc>"
        f'<w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{shading}</w:tcPr>'
        f'{docx_paragraph(text, bold=bold, size=18, align=align)}'
        "</w:tc>"
    )


def docx_table(headers: list[str], rows: list[list[str]], widths: list[int]) -> str:
    grid = "".join(f'<w:gridCol w:w="{width}"/>' for width in widths)
    header_row = "<w:tr>" + "".join(
        docx_cell(header, widths[index], bold=True, fill="EAF3EF", align="center")
        for index, header in enumerate(headers)
    ) + "</w:tr>"
    body_rows = []
    for row in rows:
        body_rows.append(
            "<w:tr>"
            + "".join(docx_cell(str(value), widths[index], align="center" if index in {0, 2, 3, 4, 5} else None) for index, value in enumerate(row))
            + "</w:tr>"
        )
    borders = (
        '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="B8C7C0"/>'
        '<w:left w:val="single" w:sz="4" w:color="B8C7C0"/>'
        '<w:bottom w:val="single" w:sz="4" w:color="B8C7C0"/>'
        '<w:right w:val="single" w:sz="4" w:color="B8C7C0"/>'
        '<w:insideH w:val="single" w:sz="4" w:color="B8C7C0"/>'
        '<w:insideV w:val="single" w:sz="4" w:color="B8C7C0"/></w:tblBorders>'
    )
    return f'<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>{borders}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{header_row}{"".join(body_rows)}</w:tbl>'


def build_proposal_docx(payload: dict) -> bytes:
    report_name = str(payload.get("name") or "Расчет КП")
    total = round(float(payload.get("total") or 0), 2)
    vat = round(total * 0.2, 2)
    total_with_vat = round(total + vat, 2)
    created = datetime.now().strftime("%d.%m.%Y")

    item_rows = []
    for index, item in enumerate(payload.get("matches", []), start=1):
        quantity = float(item.get("quantity") or 0)
        price = float(item.get("price") or 0)
        item_rows.append(
            [
                str(index),
                str(item.get("matchedName") or item.get("description") or ""),
                str(item.get("unit") or ""),
                f"{quantity:g}",
                rub(price),
                rub(quantity * price),
            ]
        )
    if not item_rows:
        item_rows.append(["1", "Позиции будут уточнены после проверки спецификации", "компл", "1", rub(0), rub(0)])

    body = [
        docx_paragraph("КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ", bold=True, size=32, color="0F7C66", align="center"),
        docx_paragraph(f"№ КП-042/26 от {created}", size=22, align="center"),
        docx_paragraph(""),
        docx_paragraph("Продавец", bold=True, size=24, color="1F2722"),
        docx_paragraph('ООО "ИнжСервис-Монтаж"', bold=True),
        docx_paragraph("ИНН 7704826140, КПП 770401001"),
        docx_paragraph("Адрес: 115114, г. Москва, Дербеневская наб., д. 7, стр. 2"),
        docx_paragraph("Тел.: +7 (495) 222-18-40, email: tender@ism-demo.ru"),
        docx_paragraph(""),
        docx_paragraph("Покупатель", bold=True, size=24, color="1F2722"),
        docx_paragraph('АО "Северный машиностроительный завод"', bold=True),
        docx_paragraph("ИНН 2901182045, КПП 290101001"),
        docx_paragraph("Объект: модернизация узла теплоснабжения производственного корпуса №3"),
        docx_paragraph(""),
        docx_paragraph("Основание расчета", bold=True, size=24, color="1F2722"),
        docx_paragraph(f"Расчет подготовлен на основании сохраненного отчета: {report_name}."),
        docx_paragraph("Позиции с низкой уверенностью совпадения требуют финальной проверки перед отправкой заказчику."),
        docx_paragraph(""),
        docx_paragraph("Спецификация материалов", bold=True, size=24, color="1F2722"),
        docx_table(
            ["№", "Наименование", "Ед.", "Кол-во", "Цена, руб.", "Сумма, руб."],
            item_rows,
            [500, 4300, 700, 900, 1400, 1560],
        ),
        docx_paragraph(""),
        docx_paragraph(f"Итого без НДС: {rub(total)} руб.", bold=True, size=24, align="right"),
        docx_paragraph(f"НДС 20%: {rub(vat)} руб.", size=22, align="right"),
        docx_paragraph(f"Итого с НДС: {rub(total_with_vat)} руб.", bold=True, size=26, color="0F7C66", align="right"),
        docx_paragraph(""),
        docx_paragraph("Условия поставки", bold=True, size=24, color="1F2722"),
        docx_paragraph("Срок поставки: 7-10 рабочих дней после оплаты и подтверждения спецификации."),
        docx_paragraph("Условия оплаты: 70% предоплата, 30% по готовности к отгрузке."),
        docx_paragraph("Срок действия предложения: 10 рабочих дней."),
        docx_paragraph("Доставка: до склада покупателя, стоимость доставки уточняется после согласования объема поставки."),
        docx_paragraph(""),
        docx_paragraph("Контактное лицо", bold=True, size=24, color="1F2722"),
        docx_paragraph("Иванов Андрей Сергеевич, руководитель проектного отдела"),
        docx_paragraph("Тел.: +7 (495) 222-18-40 доб. 124, email: a.ivanov@ism-demo.ru"),
        docx_paragraph(""),
        docx_paragraph("Подпись продавца: __________________ / Иванов А.С. /"),
    ]

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body>'
        + "".join(body)
        + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr>'
        '</w:body></w:document>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    )
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", rels)
        archive.writestr("word/document.xml", document_xml)
    return buffer.getvalue()


def get_project_payload(project_id: int) -> dict | None:
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute("select payload_json from projects where id = ?", (project_id,)).fetchone()
    if not row:
        return None
    return json.loads(row[0])


def simple_normalize_name(text: str) -> str:
    value = str(text or "").strip()
    value = re.sub(r"\s+", " ", value)
    value = value.replace("Ду", "DN").replace("ду", "DN")
    value = value.replace("Ру", "PN").replace("ру", "PN")
    return value


def normalize_items(payload: dict) -> dict:
    items = payload.get("items", [])
    normalized = [
        {"index": item.get("index", index), "normalized": simple_normalize_name(item.get("description", ""))}
        for index, item in enumerate(items)
    ]
    return {"normalized": normalized}


INDEX_HTML = r"""<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Инженерные КП - демо</title>
  <style>
    :root {
      --bg: #f6f7f4;
      --ink: #1f2722;
      --muted: #66716a;
      --line: #d7ddd5;
      --panel: #ffffff;
      --accent: #0f7c66;
      --accent-2: #e56b3f;
      --soft: #e8f3ef;
      --warn: #fff5d7;
      --danger: #fff0ec;
      --shadow: 0 12px 30px rgba(31, 39, 34, 0.08);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      color: var(--ink);
      background: var(--bg);
    }
    header {
      border-bottom: 1px solid var(--line);
      background: rgba(255,255,255,.86);
      position: sticky;
      top: 0;
      z-index: 20;
      backdrop-filter: blur(12px);
    }
    .shell { max-width: 1440px; margin: 0 auto; padding: 0 24px; }
    .topbar { min-height: 68px; display: flex; align-items: center; justify-content: space-between; gap: 20px; }
    .brand { display: flex; align-items: center; gap: 12px; min-width: 260px; }
    .mark {
      width: 38px; height: 38px; border-radius: 8px; background: var(--ink); color: white;
      display: grid; place-items: center; font-weight: 800; letter-spacing: 0;
    }
    h1 { margin: 0; font-size: 20px; line-height: 1.2; }
    .subtitle { color: var(--muted); font-size: 13px; margin-top: 3px; }
    .status-strip { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; justify-content: flex-end; }
    .pill {
      border: 1px solid var(--line); background: #fff; border-radius: 999px; padding: 8px 11px;
      color: var(--muted); font-size: 13px; white-space: nowrap;
    }
    .pill strong { color: var(--ink); font-weight: 700; }
    main { padding: 24px 0 42px; }
    .layout { display: grid; grid-template-columns: 320px minmax(0, 1fr); gap: 20px; align-items: start; }
    .flow {
      display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; margin-bottom: 18px;
    }
    .flow-card {
      background: #fff; border: 1px solid var(--line); border-radius: 8px; padding: 14px;
      min-height: 112px; box-shadow: var(--shadow);
    }
    .flow-card b {
      display: inline-grid; place-items: center; width: 26px; height: 26px; border-radius: 50%;
      background: var(--ink); color: #fff; font-size: 13px; margin-bottom: 10px;
    }
    .flow-card strong { display: block; font-size: 14px; margin-bottom: 5px; }
    .flow-card span { color: var(--muted); font-size: 12px; line-height: 1.4; display: block; }
    .notice {
      margin-bottom: 18px; border: 1px solid #c9ddd5; background: #eef8f4; color: #17483d;
      border-radius: 8px; padding: 14px 16px; display: flex; align-items: center; justify-content: space-between; gap: 12px;
      box-shadow: var(--shadow);
    }
    .notice strong { font-size: 14px; }
    .notice span { color: #3a675d; font-size: 13px; }
    aside {
      position: sticky; top: 92px; background: var(--panel); border: 1px solid var(--line);
      border-radius: 8px; box-shadow: var(--shadow); overflow: hidden;
    }
    .step { padding: 16px; border-bottom: 1px solid var(--line); }
    .step:last-child { border-bottom: 0; }
    .step h2 { font-size: 14px; margin: 0 0 10px; }
    .step p { margin: 0 0 12px; color: var(--muted); font-size: 13px; line-height: 1.45; }
    .row { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
    .file-links { display: grid; gap: 8px; margin-top: 12px; }
    .file-links a {
      color: var(--accent); text-decoration: none; font-size: 13px; font-weight: 700;
      border-bottom: 1px solid rgba(15,124,102,.25); width: fit-content;
    }
    .report-actions { display: flex; gap: 6px; flex-wrap: wrap; }
    .mini-btn {
      height: 30px; padding: 0 9px; border-radius: 6px; font-size: 12px; font-weight: 800;
      border: 1px solid var(--line); background: #fff; color: var(--ink);
    }
    .mini-btn.primary { background: var(--accent); border-color: var(--accent); color: #fff; }
    .mini-btn.warm { background: #fff5d7; border-color: #ead78e; color: #6b4c0f; }
    .roadmap {
      display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; padding: 16px;
    }
    .roadmap-item {
      border: 1px solid var(--line); border-radius: 8px; padding: 13px; background: #fbfcfb;
    }
    .roadmap-item strong { display: block; font-size: 13px; margin-bottom: 5px; }
    .roadmap-item span { color: var(--muted); font-size: 12px; line-height: 1.4; display: block; }
    button, .file-label, select, input[type="text"] {
      height: 38px; border-radius: 7px; border: 1px solid var(--line); background: #fff;
      color: var(--ink); padding: 0 12px; font: inherit; font-size: 13px;
    }
    button, .file-label { cursor: pointer; display: inline-flex; align-items: center; justify-content: center; gap: 8px; font-weight: 700; }
    button.primary { background: var(--accent); border-color: var(--accent); color: #fff; }
    button.dark { background: var(--ink); border-color: var(--ink); color: #fff; }
    button.ghost { background: transparent; }
    button:disabled { opacity: .45; cursor: not-allowed; }
    input[type="file"] { display: none; }
    select, input[type="text"] { width: 100%; }
    .workspace { display: grid; gap: 18px; }
    .summary {
      display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px;
    }
    .metric {
      background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 16px;
      min-height: 92px;
    }
    .metric span { color: var(--muted); font-size: 13px; }
    .metric strong { display: block; margin-top: 10px; font-size: 26px; line-height: 1; }
    section.panel {
      background: var(--panel); border: 1px solid var(--line); border-radius: 8px; box-shadow: var(--shadow);
      overflow: hidden;
    }
    .panel-head {
      padding: 15px 16px; display: flex; align-items: center; justify-content: space-between; gap: 14px;
      border-bottom: 1px solid var(--line);
    }
    .panel-head h2 { margin: 0; font-size: 16px; }
    .panel-head p { margin: 3px 0 0; color: var(--muted); font-size: 13px; }
    .table-wrap { overflow: auto; max-height: 520px; }
    table { border-collapse: collapse; width: 100%; min-width: 1080px; font-size: 13px; }
    th, td { border-bottom: 1px solid var(--line); padding: 10px 11px; text-align: left; vertical-align: top; }
    th { background: #f0f3ef; color: #405048; position: sticky; top: 0; z-index: 3; font-size: 12px; }
    td[contenteditable="true"] { background: #fffdfa; outline: none; }
    td[contenteditable="true"]:focus { box-shadow: inset 0 0 0 2px rgba(15,124,102,.32); }
    .confidence { border-radius: 999px; padding: 5px 8px; font-weight: 800; font-size: 12px; display: inline-block; }
    .ok { background: var(--soft); color: #09614f; }
    .mid { background: var(--warn); color: #7a5610; }
    .bad { background: var(--danger); color: #a13e21; }
    .empty {
      padding: 30px 20px; color: var(--muted); text-align: center;
      border: 1px dashed var(--line); margin: 16px; border-radius: 8px; background: #fbfcfb;
    }
    .log { color: var(--muted); font-size: 12px; line-height: 1.45; margin-top: 8px; }
    .toast {
      position: fixed; right: 18px; bottom: 18px; max-width: 420px; padding: 13px 14px;
      border-radius: 8px; color: #fff; background: var(--ink); box-shadow: var(--shadow); display: none;
      font-size: 13px; z-index: 50;
    }
    .toast.show { display: block; }
    @media (max-width: 980px) {
      .layout { grid-template-columns: 1fr; }
      aside { position: static; }
      .summary { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .flow { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .roadmap { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .notice { align-items: flex-start; flex-direction: column; }
      .topbar { align-items: flex-start; flex-direction: column; padding: 14px 0; }
      .status-strip { justify-content: flex-start; }
    }
    @media (max-width: 640px) {
      .shell { padding: 0 14px; }
      .summary { grid-template-columns: 1fr; }
      .flow { grid-template-columns: 1fr; }
      .roadmap { grid-template-columns: 1fr; }
      .brand { min-width: 0; }
    }
  </style>
</head>
<body>
  <header>
    <div class="shell topbar">
      <div class="brand">
        <div class="mark">КП</div>
        <div>
          <h1>Расчет КП по спецификации</h1>
          <div class="subtitle">Демо: спецификация, прайс, сопоставление, выгрузки</div>
        </div>
      </div>
      <div class="status-strip">
        <div class="pill">Демо: <strong>моковые данные</strong></div>
        <div class="pill">Контроль: <strong>проверка инженером</strong></div>
        <div class="pill">Выход: <strong>Excel / CSV / КП</strong></div>
      </div>
    </div>
  </header>

  <main class="shell">
    <div class="notice">
      <div>
        <strong id="nextAction">Начните с загрузки спецификации.</strong>
        <span>Можно взять готовые тестовые файлы ниже или нажать кнопки демо.</span>
      </div>
      <button class="primary" id="quickDemo">Заполнить демо</button>
    </div>

    <div class="flow">
      <div class="flow-card"><b>1</b><strong>Спецификация</strong><span id="checkSpec">Не загружена</span></div>
      <div class="flow-card"><b>2</b><strong>Прайс</strong><span id="checkPrice">Не загружен</span></div>
      <div class="flow-card"><b>3</b><strong>Сопоставление</strong><span id="checkMatch">Ожидает данных</span></div>
      <div class="flow-card"><b>4</b><strong>Выгрузки</strong><span id="checkExport">Появятся после сохранения</span></div>
    </div>

    <div class="layout">
      <aside>
        <div class="step">
          <h2>1. Спецификация</h2>
          <p>Загрузите CSV со строками спецификации или нажмите кнопку с тестовым примером.</p>
          <div class="row">
            <button class="primary" id="loadSpec">Загрузить тест</button>
            <button class="ghost" id="loadSpecFull">Большой тест</button>
            <label class="file-label">Выбрать CSV/XLSX<input id="specFile" type="file" accept=".csv,.txt,.xlsx" /></label>
          </div>
          <div class="file-links">
            <a href="/samples/test_spec_small.csv" download>Маленькая спецификация CSV</a>
            <a href="/samples/test_spec_full.csv" download>Расширенная спецификация CSV</a>
            <a href="/samples/test_spec_full.xlsx" download>Расширенная спецификация XLSX</a>
          </div>
          <div class="log" id="specLog">Файл не загружен.</div>
        </div>
        <div class="step">
          <h2>2. Прайс поставщика</h2>
          <p>Загрузите CSV с позициями поставщика: артикул, название, единица, цена.</p>
          <div class="row">
            <button class="primary" id="loadPrices">Загрузить тест</button>
            <button class="ghost" id="loadPricesFull">Большой тест</button>
            <label class="file-label">Выбрать CSV/XLSX<input id="priceFile" type="file" accept=".csv,.txt,.xlsx" /></label>
          </div>
          <div class="file-links">
            <a href="/samples/test_price_small.csv" download>Маленький прайс CSV</a>
            <a href="/samples/test_price_full.csv" download>Расширенный прайс CSV</a>
            <a href="/samples/test_price_full.xlsx" download>Расширенный прайс XLSX</a>
          </div>
          <div class="log" id="priceLog">Прайс не загружен.</div>
        </div>
        <div class="step">
          <h2>3. Подготовка данных</h2>
          <p>После загрузки спецификации можно привести обозначения к единому виду.</p>
          <div class="row" style="margin-top:8px">
            <button class="ghost" id="normalizeBtn">Нормализовать позиции</button>
          </div>
          <div class="log" id="modelLog">Готово к обработке.</div>
        </div>
        <div class="step">
          <h2>4. Результат</h2>
          <p>Когда оба файла загружены, нажмите сопоставление. После проверки сохраните расчет, а выгрузки появятся в сохраненных отчетах.</p>
          <div class="row">
            <button class="dark" id="matchBtn">Сопоставить</button>
            <button id="saveBtn">Сохранить</button>
          </div>
          <div class="log">Строки с низкой уверенностью нужно проверить вручную перед отправкой КП.</div>
        </div>
      </aside>

      <div class="workspace">
        <div class="summary">
          <div class="metric"><span>Позиций спецификации</span><strong id="specCount">0</strong></div>
          <div class="metric"><span>Позиций в прайсе</span><strong id="priceCount">0</strong></div>
          <div class="metric"><span>Требуют проверки</span><strong id="reviewCount">0</strong></div>
          <div class="metric"><span>Итого, ₽</span><strong id="totalSum">0</strong></div>
        </div>

        <section class="panel">
          <div class="panel-head">
            <div>
              <h2>Сопоставленная ведомость</h2>
              <p>Редактируемые поля подсвечены светлым фоном. Финальное решение остается за человеком.</p>
            </div>
            <input id="projectName" type="text" value="Демо КП - инженерные сети" />
          </div>
          <div id="matchEmpty" class="empty">Загрузите спецификацию и прайс, затем нажмите “Сопоставить”.</div>
          <div class="table-wrap" id="matchWrap" style="display:none">
            <table>
              <thead>
                <tr>
                  <th>Раздел</th><th>Поз.</th><th>Исходное наименование</th><th>Ед.</th><th>Кол-во</th>
                  <th>Артикул</th><th>Позиция поставщика</th><th>Производитель</th><th>Цена</th><th>Сумма</th><th>Уверенность</th><th>Статус</th>
                </tr>
              </thead>
              <tbody id="matchBody"></tbody>
            </table>
          </div>
        </section>

        <section class="panel">
          <div class="panel-head">
            <div>
              <h2>Сохраненные отчеты</h2>
              <p>Из каждого отчета можно скачать таблицу или черновик коммерческого предложения.</p>
            </div>
            <button class="ghost" id="reloadProjects">Обновить</button>
          </div>
          <div class="table-wrap">
            <table style="min-width:720px">
              <thead><tr><th>ID</th><th>Проект</th><th>Дата</th><th>Итого</th><th>Скачать</th></tr></thead>
              <tbody id="projectsBody"></tbody>
            </table>
          </div>
        </section>

        <section class="panel">
          <div class="panel-head">
            <div>
              <h2>Что подключается следующим этапом</h2>
              <p>Это не мешает демо, но показывает развитие продукта под реальные процессы заказчика.</p>
            </div>
          </div>
          <div class="roadmap">
            <div class="roadmap-item"><strong>Исходные документы</strong><span>Разбор спецификаций из проектных файлов и перенос в редактируемую таблицу.</span></div>
            <div class="roadmap-item"><strong>Прайсы поставщиков</strong><span>Регулярное обновление цен, хранение даты прайса и истории изменений.</span></div>
            <div class="roadmap-item"><strong>Проверка совпадений</strong><span>Подсветка спорных строк, ручное подтверждение и обучение на исправлениях.</span></div>
            <div class="roadmap-item"><strong>Документы</strong><span>Шаблоны КП, актов, приказов и исполнительной документации по данным проекта.</span></div>
          </div>
        </section>
      </div>
    </div>
  </main>
  <div class="toast" id="toast"></div>

  <script>
    const state = { spec: [], prices: [], matches: [] };

    const el = (id) => document.getElementById(id);
    const money = (n) => new Intl.NumberFormat('ru-RU').format(Math.round(Number(n) || 0));
    const toast = (text) => {
      el('toast').textContent = text;
      el('toast').classList.add('show');
      setTimeout(() => el('toast').classList.remove('show'), 3200);
    };

    function parseCSV(text) {
      const firstLine = String(text || '').split(/\r?\n/).find((line) => line.trim()) || '';
      const delimiter = firstLine.includes(';') && !firstLine.includes(',') ? ';' : ',';
      const rows = [];
      let row = [], cell = '', quoted = false;
      for (let i = 0; i < text.length; i++) {
        const ch = text[i], next = text[i + 1];
        if (ch === '"' && quoted && next === '"') { cell += '"'; i++; continue; }
        if (ch === '"') { quoted = !quoted; continue; }
        if (ch === delimiter && !quoted) { row.push(cell.trim()); cell = ''; continue; }
        if ((ch === '\n' || ch === '\r') && !quoted) {
          if (ch === '\r' && next === '\n') i++;
          row.push(cell.trim());
          if (row.some(Boolean)) rows.push(row);
          row = []; cell = '';
          continue;
        }
        cell += ch;
      }
      row.push(cell.trim());
      if (row.some(Boolean)) rows.push(row);
      if (!rows.length) return [];
      const headers = rows.shift().map((h) => h.trim());
      return rows.map((values) => Object.fromEntries(headers.map((h, i) => [h, values[i] ?? ''])));
    }

    function validateRows(rows, kind) {
      if (!rows.length) return 'Файл пустой или не содержит строк.';
      const required = kind === 'spec'
        ? ['description', 'unit', 'quantity']
        : ['sku', 'name', 'unit', 'price_rub'];
      const columns = new Set(Object.keys(rows[0] || {}));
      const missing = required.filter((column) => !columns.has(column));
      if (missing.length) return `Не хватает колонок: ${missing.join(', ')}.`;
      return '';
    }

    function normalize(text) {
      return String(text || '')
        .toLowerCase()
        .replaceAll('ё', 'е')
        .replace(/dn\s*/g, 'dn')
        .replace(/pn\s*/g, 'pn')
        .replace(/[^\p{L}\p{N},.xх]+/gu, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    }

    function tokens(text) {
      return normalize(text).split(' ').filter((t) => t.length > 1);
    }

    function score(spec, price) {
      const source = normalize(spec.description);
      const target = normalize(price.name);
      if (!source || !target) return 0;
      if (source === target) return 0.99;
      const a = new Set(tokens(source));
      const b = new Set(tokens(target));
      let overlap = 0;
      for (const token of a) if (b.has(token) || target.includes(token)) overlap++;
      const base = overlap / Math.max(a.size, 1);
      const unitBonus = normalize(spec.unit) && normalize(spec.unit) === normalize(price.unit) ? 0.12 : -0.08;
      const dnSource = source.match(/dn\s?(\d+)/i)?.[1] || source.match(/\b(\d{2,3})x/i)?.[1];
      const dnTarget = target.match(/dn\s?(\d+)/i)?.[1] || target.match(/\b(\d{2,3})x/i)?.[1];
      const sizeBonus = dnSource && dnTarget && dnSource === dnTarget ? 0.18 : 0;
      return Math.max(0, Math.min(0.98, base + unitBonus + sizeBonus));
    }

    function bestMatch(spec) {
      let best = null;
      for (const price of state.prices) {
        const current = score(spec, price);
        if (!best || current > best.score) best = { price, score: current };
      }
      if (!best || best.score < 0.38) return { price: {}, score: 0, status: 'не найдено' };
      const status = best.score >= 0.82 ? 'точное совпадение' : best.score >= 0.62 ? 'похожее совпадение' : 'нужно проверить';
      return { ...best, status };
    }

    function runMatching() {
      if (!state.spec.length || !state.prices.length) {
        toast('Нужны спецификация и прайс.');
        return false;
      }
      state.matches = state.spec.map((spec, index) => {
        const found = bestMatch(spec);
        const price = found.price || {};
        const quantity = Number(String(spec.quantity || '0').replace(',', '.')) || 0;
        const priceRub = Number(String(price.price_rub || '0').replace(',', '.')) || 0;
        return {
          index,
          section: spec.section || '',
          item: spec.item || index + 1,
          description: spec.description || spec.name || '',
          unit: spec.unit || '',
          quantity,
          comment: spec.comment || '',
          sku: price.sku || '',
          matchedName: price.name || '',
          manufacturer: price.manufacturer || '',
          price: priceRub,
          sum: Math.round(quantity * priceRub * 100) / 100,
          confidence: Math.round((found.score || 0) * 100),
          status: found.status
        };
      });
      renderMatches();
      toast('Сопоставление выполнено. Проверьте желтые и красные строки.');
      return true;
    }

    function confidenceClass(row) {
      if (row.confidence >= 82) return 'confidence ok';
      if (row.confidence >= 62) return 'confidence mid';
      return 'confidence bad';
    }

    function updateSummary() {
      const total = state.matches.reduce((sum, row) => sum + (Number(row.sum) || 0), 0);
      const review = state.matches.filter((row) => row.confidence < 82).length;
      el('specCount').textContent = state.spec.length;
      el('priceCount').textContent = state.prices.length;
      el('reviewCount').textContent = review;
      el('totalSum').textContent = money(total);
      updateWorkflow();
    }

    function updateWorkflow() {
      const hasSpec = state.spec.length > 0;
      const hasPrices = state.prices.length > 0;
      const hasMatches = state.matches.length > 0;
      el('checkSpec').textContent = hasSpec ? `Загружено ${state.spec.length} строк` : 'Не загружена';
      el('checkPrice').textContent = hasPrices ? `Загружено ${state.prices.length} строк` : 'Не загружен';
      el('checkMatch').textContent = hasMatches ? `Готово ${state.matches.length} строк` : hasSpec && hasPrices ? 'Можно запускать' : 'Ожидает данных';
      el('checkExport').textContent = hasMatches ? 'Сохраните отчет' : 'Появится после расчета';
      el('normalizeBtn').disabled = !hasSpec;
      el('matchBtn').disabled = !(hasSpec && hasPrices);
      el('saveBtn').disabled = !hasMatches;
      if (!hasSpec) {
        el('nextAction').textContent = 'Шаг 1: загрузите спецификацию.';
      } else if (!hasPrices) {
        el('nextAction').textContent = 'Шаг 2: загрузите прайс поставщика.';
      } else if (!hasMatches) {
        el('nextAction').textContent = 'Шаг 3: нажмите “Сопоставить”.';
      } else {
        el('nextAction').textContent = 'Расчет готов: проверьте строки и нажмите “Сохранить”.';
      }
    }

    function renderMatches() {
      updateSummary();
      el('matchEmpty').style.display = state.matches.length ? 'none' : 'block';
      el('matchWrap').style.display = state.matches.length ? 'block' : 'none';
      el('matchBody').innerHTML = state.matches.map((row, i) => `
        <tr data-index="${i}">
          <td>${html(row.section)}</td>
          <td>${html(row.item)}</td>
          <td contenteditable="true" data-field="description">${html(row.description)}</td>
          <td contenteditable="true" data-field="unit">${html(row.unit)}</td>
          <td contenteditable="true" data-field="quantity">${html(row.quantity)}</td>
          <td contenteditable="true" data-field="sku">${html(row.sku)}</td>
          <td contenteditable="true" data-field="matchedName">${html(row.matchedName)}</td>
          <td contenteditable="true" data-field="manufacturer">${html(row.manufacturer)}</td>
          <td contenteditable="true" data-field="price">${html(row.price)}</td>
          <td>${money(row.sum)}</td>
          <td><span class="${confidenceClass(row)}">${row.confidence}%</span></td>
          <td>${html(row.status)}</td>
        </tr>
      `).join('');
    }

    function html(value) {
      return String(value ?? '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');
    }

    el('matchBody').addEventListener('input', (event) => {
      const td = event.target.closest('td[data-field]');
      if (!td) return;
      const tr = td.closest('tr');
      const row = state.matches[Number(tr.dataset.index)];
      const field = td.dataset.field;
      row[field] = td.textContent.trim();
      if (field === 'quantity' || field === 'price') {
        row.quantity = Number(String(row.quantity).replace(',', '.')) || 0;
        row.price = Number(String(row.price).replace(',', '.')) || 0;
        row.sum = Math.round(row.quantity * row.price * 100) / 100;
        renderMatches();
      }
    });

    async function loadSample(kind, size = 'small') {
      const path = kind === 'spec'
        ? (size === 'full' ? '/samples/test_spec_full.csv' : '/samples/test_spec_small.csv')
        : (size === 'full' ? '/samples/test_price_full.csv' : '/samples/test_price_small.csv');
      const text = await fetch(path).then((r) => r.text());
      if (kind === 'spec') {
        const rows = mapRows(parseCSV(text), 'spec');
        const error = validateRows(rows, 'spec');
        if (error) {
          toast(`Ошибка спецификации: ${error}`);
          return;
        }
        state.spec = rows;
        state.matches = [];
        el('specLog').textContent = `Загружено ${state.spec.length} строк из ${size === 'full' ? 'расширенной' : 'маленькой'} спецификации.`;
      } else {
        const rows = mapRows(parseCSV(text), 'price');
        const error = validateRows(rows, 'price');
        if (error) {
          toast(`Ошибка прайса: ${error}`);
          return;
        }
        state.prices = rows;
        state.matches = [];
        el('priceLog').textContent = `Загружено ${state.prices.length} строк из ${size === 'full' ? 'расширенного' : 'маленького'} прайса.`;
      }
      updateSummary();
    }

    function canonicalHeader(value) {
      return String(value || '').trim().toLowerCase().replaceAll('ё', 'е').replace(/[\s._/-]+/g, '');
    }

    function fieldForHeader(header, kind) {
      const aliases = kind === 'spec'
        ? {
            section: ['section','раздел','система','группа'],
            item: ['item','поз','пп','номер','n','no'],
            description: ['description','name','наименование','материал','оборудование','номенклатура','описание'],
            unit: ['unit','ед','едизм','единица','единицаизмерения'],
            quantity: ['quantity','qty','кол','колво','количество','объем'],
            comment: ['comment','комментарий','примечание']
          }
        : {
            sku: ['sku','артикул','код','кодтовара','номеркод'],
            name: ['name','description','наименование','товар','номенклатура','описание'],
            category: ['category','категория','группа','раздел'],
            unit: ['unit','ед','едизм','единица','единицаизмерения'],
            price_rub: ['price','price_rub','ценаруб','цена','стоимость','розница','розничнаяцена'],
            manufacturer: ['manufacturer','производитель','бренд','завод'],
            lead_time_days: ['leadtime','срок','срокпоставки','поставка']
          };
      const normalized = canonicalHeader(header);
      for (const [field, values] of Object.entries(aliases)) if (values.includes(normalized)) return field;
      return '';
    }

    function cleanNumber(value) {
      const text = String(value || '').replace(/[^\d,.-]/g, '');
      return text.includes(',') && !text.includes('.') ? text.replace(',', '.') : text;
    }

    function mapRows(rawRows, kind) {
      if (!rawRows.length) return [];
      const headerMap = {};
      Object.keys(rawRows[0]).forEach((header) => {
        const field = fieldForHeader(header, kind);
        if (field && !headerMap[field]) headerMap[field] = header;
      });
      return rawRows.map((raw, index) => {
        const row = {};
        Object.entries(headerMap).forEach(([field, header]) => row[field] = String(raw[header] ?? '').trim());
        if (kind === 'spec') {
          row.section ??= '';
          row.item ||= String(index + 1);
          row.comment ??= '';
          row.quantity = cleanNumber(row.quantity);
        } else {
          row.sku ??= '';
          row.category ??= '';
          row.manufacturer ??= '';
          row.lead_time_days ??= '';
          row.price_rub = cleanNumber(row.price_rub);
        }
        return row;
      }).filter((row) => kind === 'spec' ? row.description : row.name);
    }

    async function readFile(input, kind) {
      const file = input.files?.[0];
      if (!file) return;
      if (!file.name.toLowerCase().match(/\.(csv|txt|xlsx)$/)) {
        toast('Загрузите CSV или XLSX.');
        return;
      }
      const formData = new FormData();
      formData.append('file', file);
      const response = await fetch(`/api/parse-file?kind=${kind}`, { method: 'POST', body: formData });
      const data = await response.json();
      if (!response.ok) {
        toast(data.error || 'Не удалось прочитать файл.');
        return;
      }
      if (kind === 'spec') {
        state.spec = data.rows;
        state.matches = [];
        el('specLog').textContent = `Загружено ${data.rows.length} строк из ${file.name}.`;
      } else {
        state.prices = data.rows;
        state.matches = [];
        el('priceLog').textContent = `Загружено ${data.rows.length} строк из ${file.name}.`;
      }
        updateSummary();
    }

    function payload() {
      const total = state.matches.reduce((sum, row) => sum + (Number(row.sum) || 0), 0);
      return {
        name: el('projectName').value || 'Демо КП',
        spec: state.spec,
        prices: state.prices,
        matches: state.matches,
        total,
        processing: {
          mode: 'demo-normalization',
          humanReview: true
        }
      };
    }

    async function saveProject() {
      if (!state.matches.length && !runMatching()) return;
      if (!state.matches.length) {
        toast('Сначала выполните сопоставление.');
        return;
      }
      const hasUsefulRows = state.matches.some((row) => row.matchedName && Number(row.quantity) > 0);
      if (!hasUsefulRows) {
        toast('Не найдено ни одной позиции для отчета. Проверьте прайс или спецификацию.');
        return;
      }
      const response = await fetch('/api/projects', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload())
      });
      const data = await response.json();
      toast(`Расчет сохранен, номер ${data.id}.`);
      loadProjects();
    }

    async function loadProjects() {
      const data = await fetch('/api/projects').then((r) => r.json());
      el('projectsBody').innerHTML = data.projects.map((project) => `
        <tr>
          <td>${project.id}</td>
          <td>${html(project.name)}</td>
          <td>${html(project.created_at)}</td>
          <td>${money(project.total)}</td>
          <td>
            <div class="report-actions">
              <button class="mini-btn primary" data-report="${project.id}" data-format="xlsx">Excel</button>
              <button class="mini-btn" data-report="${project.id}" data-format="csv">CSV</button>
              <button class="mini-btn warm" data-report="${project.id}" data-format="proposal">КП DOCX</button>
            </div>
          </td>
        </tr>
      `).join('') || '<tr><td colspan="5">Пока нет сохраненных отчетов.</td></tr>';
    }

    function downloadReport(id, format) {
      const url = `/api/projects/${id}/download?format=${encodeURIComponent(format)}`;
      const a = document.createElement('a');
      a.href = url;
      a.click();
      const label = format === 'proposal' ? 'черновик КП' : format.toUpperCase();
      toast(`Скачиваем ${label} из сохраненного отчета.`);
    }

    async function normalizeRows() {
      if (!state.spec.length) {
        toast('Сначала загрузите спецификацию.');
        return;
      }
      const items = state.spec.map((row, index) => ({ index, description: row.description || row.name || '' }));
      const response = await fetch('/api/normalize', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          items
        })
      });
      const data = await response.json();
      for (const item of data.normalized || []) {
        if (state.spec[item.index] && item.normalized) state.spec[item.index].description = item.normalized;
      }
      toast('Позиции нормализованы.');
      if (state.matches.length) runMatching();
    }

    el('loadSpec').addEventListener('click', () => loadSample('spec'));
    el('loadSpecFull').addEventListener('click', () => loadSample('spec', 'full'));
    el('loadPrices').addEventListener('click', () => loadSample('prices'));
    el('loadPricesFull').addEventListener('click', () => loadSample('prices', 'full'));
    el('quickDemo').addEventListener('click', async () => {
      await loadSample('spec');
      await loadSample('prices');
      runMatching();
    });
    el('specFile').addEventListener('change', (e) => readFile(e.target, 'spec'));
    el('priceFile').addEventListener('change', (e) => readFile(e.target, 'prices'));
    el('matchBtn').addEventListener('click', runMatching);
    el('saveBtn').addEventListener('click', saveProject);
    el('reloadProjects').addEventListener('click', loadProjects);
    el('normalizeBtn').addEventListener('click', normalizeRows);
    el('projectsBody').addEventListener('click', (event) => {
      const button = event.target.closest('button[data-report]');
      if (!button) return;
      downloadReport(button.dataset.report, button.dataset.format);
    });

    loadProjects();
    updateSummary();
  </script>
</body>
</html>"""


class DemoHandler(BaseHTTPRequestHandler):
    server_version = "KPDemo/0.1"

    def log_message(self, format: str, *args) -> None:
        print(f"{self.address_string()} - {format % args}")

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/":
            body = INDEX_HTML.encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        if parsed.path == "/api/projects":
            with sqlite3.connect(DB_PATH) as conn:
                conn.row_factory = sqlite3.Row
                rows = conn.execute(
                    "select id, name, created_at, total from projects order by id desc limit 20"
                ).fetchall()
            send_json(self, {"projects": [dict(row) for row in rows]})
            return

        download_match = re.fullmatch(r"/api/projects/(\d+)/download", parsed.path)
        if download_match:
            project_id = int(download_match.group(1))
            params = urllib.parse.parse_qs(parsed.query)
            file_format = (params.get("format") or ["xlsx"])[0]
            payload = get_project_payload(project_id)
            if not payload:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            if not payload.get("matches"):
                send_json(self, {"error": "В отчете нет позиций для выгрузки."}, status=400)
                return

            rows = export_rows(payload)
            if file_format == "csv":
                body = build_csv(rows)
                content_type = "text/csv; charset=utf-8"
                filename = f"report-{project_id}.csv"
            elif file_format == "proposal":
                body = build_proposal_docx(payload)
                content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                filename = f"commercial-proposal-{project_id}.docx"
            else:
                body = build_xlsx(rows)
                content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                filename = f"report-{project_id}.xlsx"

            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        if parsed.path.startswith("/samples/"):
            sample = safe_sample_path(parsed.path)
            if not sample:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            body = sample.read_bytes()
            if sample.suffix.lower() == ".xlsx":
                content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            else:
                content_type = "text/csv; charset=utf-8"
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/projects":
            payload = read_json(self)
            matches = payload.get("matches") or []
            has_useful_rows = any(item.get("matchedName") and float(item.get("quantity") or 0) > 0 for item in matches)
            if not has_useful_rows:
                send_json(self, {"error": "Сначала выполните сопоставление и проверьте найденные позиции."}, status=400)
                return
            name = str(payload.get("name") or "Демо КП")
            total = float(payload.get("total") or 0)
            created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with sqlite3.connect(DB_PATH) as conn:
                cursor = conn.execute(
                    "insert into projects(name, created_at, total, payload_json) values (?, ?, ?, ?)",
                    (name, created_at, total, json.dumps(payload, ensure_ascii=False)),
                )
                conn.commit()
            send_json(self, {"id": cursor.lastrowid, "created_at": created_at})
            return

        if parsed.path == "/api/normalize":
            payload = read_json(self)
            send_json(self, normalize_items(payload))
            return

        if parsed.path == "/api/parse-file":
            try:
                params = urllib.parse.parse_qs(parsed.query)
                kind = (params.get("kind") or ["spec"])[0]
                filename, content = parse_multipart_file(self)
                lower_name = filename.lower()
                if lower_name.endswith(".xlsx"):
                    raw_rows = parse_xlsx_bytes(content)
                elif lower_name.endswith(".csv") or lower_name.endswith(".txt"):
                    raw_rows = parse_csv_bytes(content)
                else:
                    send_json(self, {"error": "Поддерживаются CSV и XLSX."}, status=400)
                    return
                rows, error = canonicalize_rows(raw_rows, "price" if kind == "price" else "spec")
                if error:
                    send_json(self, {"error": error}, status=400)
                    return
                send_json(self, {"filename": filename, "rows": rows})
            except (ValueError, KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
                send_json(self, {"error": str(exc)}, status=400)
            return

        self.send_error(HTTPStatus.NOT_FOUND)


class FastLocalServer(ThreadingHTTPServer):
    def server_bind(self) -> None:
        # Avoid slow reverse DNS lookup in http.server.HTTPServer.server_bind.
        self.socket.bind(self.server_address)
        self.server_name = self.server_address[0]
        self.server_port = self.server_address[1]


def main() -> None:
    load_env_file()
    init_db()
    host = os.getenv("HOST", "127.0.0.1")
    port = int(os.getenv("PORT", "8000"))
    print(f"Demo app: http://{host}:{port}")
    FastLocalServer((host, port), DemoHandler).serve_forever()


if __name__ == "__main__":
    main()
