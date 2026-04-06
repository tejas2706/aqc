import json
import re
import zipfile
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import Path
from typing import Any, Optional
import xml.etree.ElementTree as ET


WORKBOOK_PATH = Path(__file__).resolve().parent / "AQC Attendee List.xlsx"
DATA_PATH = Path(__file__).resolve().parent / "attendees.json"
SHEET_NAME = "Attendee List"
XML_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
MAIN_NS = XML_NS["main"]
REL_NS = XML_NS["rel"]
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)


@dataclass
class Attendee:
    name: str
    tag: str
    institute: str = ""
    hub: str = ""
    emailId: str = ""
    tagColor: str = ""
    tagBorderColor: str = ""
    tagTextColor: str = ""
    source: str = "excel"


def normalize_text(value: str) -> str:
    cleaned = (value or "").replace("\xa0", " ").strip().casefold()
    return re.sub(r"\s+", " ", cleaned)


def _column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    index = 0
    for letter in letters:
        index = index * 26 + (ord(letter.upper()) - ord("A") + 1)
    return index - 1


def _read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    tree = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    shared: list[str] = []
    for item in tree.findall("main:si", XML_NS):
        value = "".join(text.text or "" for text in item.iterfind(".//main:t", XML_NS))
        shared.append(value)
    return shared


def _theme_color_map(archive: zipfile.ZipFile) -> dict[int, str]:
    theme_root = ET.fromstring(archive.read("xl/theme/theme1.xml"))
    ns = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }
    color_scheme = theme_root.find(".//a:clrScheme", ns)
    if color_scheme is None:
        return {}

    ordered_names = [
        "lt1",
        "dk1",
        "lt2",
        "dk2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hlink",
        "folHlink",
    ]
    mapping: dict[int, str] = {}
    for index, name in enumerate(ordered_names):
        node = color_scheme.find(f"a:{name}", ns)
        if node is None or not list(node):
            continue
        child = list(node)[0]
        color = child.attrib.get("lastClr") or child.attrib.get("val")
        if color:
            mapping[index] = f"#{color[-6:]}"
    return mapping


def _apply_tint(channel: int, tint: float) -> int:
    if tint < 0:
        return max(0, min(255, round(channel * (1.0 + tint))))
    return max(0, min(255, round(channel * (1.0 - tint) + (255 - 255 * (1.0 - tint)))))


def _apply_tint_to_hex(hex_color: str, tint_value: str) -> str:
    tint = float(tint_value)
    red = int(hex_color[1:3], 16)
    green = int(hex_color[3:5], 16)
    blue = int(hex_color[5:7], 16)
    return "#{:02X}{:02X}{:02X}".format(
        _apply_tint(red, tint),
        _apply_tint(green, tint),
        _apply_tint(blue, tint),
    )


def _hex_from_excel_color(color_attrib: Optional[dict[str, str]]) -> str:
    if not color_attrib:
        return ""
    if "theme" in color_attrib:
        return ""
    rgb = color_attrib.get("rgb", "")
    if len(rgb) == 8:
        return f"#{rgb[2:]}"
    if len(rgb) == 6:
        return f"#{rgb}"
    return ""


def _resolve_excel_color(
    color_attrib: Optional[dict[str, str]],
    theme_colors: dict[int, str],
) -> str:
    if not color_attrib:
        return ""

    if "rgb" in color_attrib:
        return _hex_from_excel_color(color_attrib)

    if "theme" in color_attrib:
        base = theme_colors.get(int(color_attrib["theme"]), "")
        if not base:
            return ""
        if "tint" in color_attrib:
            return _apply_tint_to_hex(base, color_attrib["tint"])
        return base

    return ""


def _tag_style_map(
    archive: zipfile.ZipFile,
    sheet_root: ET.Element,
    shared_strings: list[str],
) -> dict[str, dict[str, str]]:
    styles = ET.fromstring(archive.read("xl/styles.xml"))
    fills = styles.find("main:fills", XML_NS)
    fonts = styles.find("main:fonts", XML_NS)
    cell_xfs = styles.find("main:cellXfs", XML_NS)
    theme_colors = _theme_color_map(archive)
    style_map: dict[str, dict[str, str]] = {}

    rows = sheet_root.find("main:sheetData", XML_NS)
    for row in rows.findall("main:row", XML_NS)[1:]:
        for cell in row.findall("main:c", XML_NS):
            if not cell.attrib.get("r", "").startswith("G"):
                continue
            value_node = cell.find("main:v", XML_NS)
            if value_node is None:
                continue

            value = value_node.text or ""
            if cell.attrib.get("t") == "s":
                value = shared_strings[int(value)]

            style_index = int(cell.attrib.get("s", "0"))
            cell_xf = cell_xfs[style_index]
            fill_id = int(cell_xf.attrib.get("fillId", "0"))
            font_id = int(cell_xf.attrib.get("fontId", "0"))
            fill = fills[fill_id]
            font = fonts[font_id]
            pattern_fill = fill.find("main:patternFill", XML_NS)
            fg_color = (
                pattern_fill.find("main:fgColor", XML_NS).attrib
                if pattern_fill is not None and pattern_fill.find("main:fgColor", XML_NS) is not None
                else None
            )
            font_color_node = font.find("main:color", XML_NS)
            background = _resolve_excel_color(fg_color, theme_colors)
            text_color = _resolve_excel_color(
                font_color_node.attrib if font_color_node is not None else None,
                theme_colors,
            )
            if value and background and value not in style_map:
                style_map[value] = {
                    "tagColor": background,
                    "tagBorderColor": background,
                    "tagTextColor": text_color,
                }

    return style_map


def _resolve_sheet_path(archive: zipfile.ZipFile, sheet_name: str) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

    for sheet in workbook.find("main:sheets", XML_NS):
        if sheet.attrib["name"] == sheet_name:
            rel_id = sheet.attrib[
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            ]
            return "xl/" + rel_map[rel_id]

    raise ValueError(f"Sheet '{sheet_name}' not found in workbook")


def _sheet_names(archive: zipfile.ZipFile) -> list[str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    return [sheet.attrib["name"] for sheet in workbook.find("main:sheets", XML_NS)]


def _parse_attendees_from_archive(
    archive: zipfile.ZipFile,
    preferred_sheet_name: str = SHEET_NAME,
) -> list[dict[str, Any]]:
    shared_strings = _read_shared_strings(archive)
    sheet_names = _sheet_names(archive)
    sheet_name = preferred_sheet_name if preferred_sheet_name in sheet_names else sheet_names[0]
    sheet_path = _resolve_sheet_path(archive, sheet_name)
    sheet = ET.fromstring(archive.read(sheet_path))
    tag_style_map = _tag_style_map(archive, sheet, shared_strings)

    rows = sheet.find("main:sheetData", XML_NS)
    extracted_rows: list[list[str]] = []

    for row in rows.findall("main:row", XML_NS):
        values: list[str] = []
        current_index = 0

        for cell in row.findall("main:c", XML_NS):
            cell_index = _column_index(cell.attrib.get("r", "A1"))
            while current_index < cell_index:
                values.append("")
                current_index += 1

            cell_type = cell.attrib.get("t")
            value_node = cell.find("main:v", XML_NS)
            value = ""
            if value_node is not None:
                if cell_type == "s":
                    value = shared_strings[int(value_node.text)]
                else:
                    value = value_node.text or ""

            values.append(value)
            current_index += 1

        extracted_rows.append(values)

    if not extracted_rows:
        return []

    header = [normalize_text(item) for item in extracted_rows[0]]
    records: list[dict[str, Any]] = []

    for raw_row in extracted_rows[1:]:
        padded = raw_row + [""] * (len(header) - len(raw_row))
        row = dict(zip(header, padded))
        tag_value = (row.get("tag", "") or "").replace("\xa0", " ").strip()
        tag_style = tag_style_map.get(tag_value, {})

        attendee = Attendee(
            name=(row.get("name", "") or "").replace("\xa0", " ").strip(),
            institute=(row.get("institute", "") or "").replace("\xa0", " ").strip(),
            hub=(row.get("hub", "") or "").replace("\xa0", " ").strip(),
            emailId=(row.get("email", "") or row.get("emailid", "") or "")
            .replace("\xa0", " ")
            .strip(),
            tag=tag_value,
            tagColor=tag_style.get("tagColor", ""),
            tagBorderColor=tag_style.get("tagBorderColor", ""),
            tagTextColor=tag_style.get("tagTextColor", ""),
            source="excel",
        )

        if attendee.name and attendee.tag:
            records.append(asdict(attendee))

    return records


def _resolve_sheet_archive_path(sheet_target: str) -> str:
    return sheet_target[3:] if sheet_target.startswith("xl/") else sheet_target


def _read_shared_strings_with_tree(
    archive: zipfile.ZipFile,
) -> tuple[ET.Element, list[str], dict[str, int]]:
    tree = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    values: list[str] = []
    lookup: dict[str, int] = {}
    for index, item in enumerate(tree.findall("main:si", XML_NS)):
        value = "".join(text.text or "" for text in item.iterfind(".//main:t", XML_NS))
        values.append(value)
        lookup.setdefault(value, index)
    return tree, values, lookup


def _last_row_number(sheet_root: ET.Element) -> int:
    rows = sheet_root.find("main:sheetData", XML_NS)
    if rows is None:
        return 2
    return max((int(row.attrib["r"]) for row in rows.findall("main:row", XML_NS)), default=2)


def _next_shared_string_index(
    shared_root: ET.Element,
    shared_values: list[str],
    shared_lookup: dict[str, int],
    value: str,
) -> int:
    if value in shared_lookup:
        return shared_lookup[value]

    item = ET.Element(f"{{{MAIN_NS}}}si")
    text_node = ET.SubElement(item, f"{{{MAIN_NS}}}t")
    text_node.text = value
    shared_root.append(item)

    index = len(shared_values)
    shared_values.append(value)
    shared_lookup[value] = index
    shared_root.attrib["uniqueCount"] = str(len(shared_values))
    return index


def _update_shared_string_count(shared_root: ET.Element, increment: int) -> None:
    current = int(shared_root.attrib.get("count", "0"))
    shared_root.attrib["count"] = str(current + increment)


def append_attendee_to_workbook(attendee: Attendee) -> None:
    with zipfile.ZipFile(WORKBOOK_PATH, "r") as source_archive:
        workbook = ET.fromstring(source_archive.read("xl/workbook.xml"))
        rels = ET.fromstring(source_archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

        sheet_target = None
        for sheet in workbook.find("main:sheets", XML_NS):
            if sheet.attrib["name"] == SHEET_NAME:
                rel_id = sheet.attrib[
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                ]
                sheet_target = rel_map[rel_id]
                break

        if sheet_target is None:
            raise ValueError(f"Sheet '{SHEET_NAME}' not found in workbook")

        sheet_archive_path = _resolve_sheet_archive_path(sheet_target)
        sheet_root = ET.fromstring(source_archive.read(f"xl/{sheet_archive_path}"))
        shared_root, shared_values, shared_lookup = _read_shared_strings_with_tree(source_archive)

        sheet_data = sheet_root.find("main:sheetData", XML_NS)
        existing_rows = sheet_data.findall("main:row", XML_NS)
        template_row = existing_rows[-1]
        template_cells = {
            cell.attrib["r"][0]: cell for cell in template_row.findall("main:c", XML_NS)
        }

        row_number = _last_row_number(sheet_root) + 1
        new_row = ET.Element(
            f"{{{MAIN_NS}}}row",
            {
                "r": str(row_number),
                "spans": template_row.attrib.get("spans", "2:7"),
                "ht": template_row.attrib.get("ht", "13"),
            },
        )
        if "customHeight" in template_row.attrib:
            new_row.attrib["customHeight"] = template_row.attrib["customHeight"]

        serial_cell = ET.Element(
            f"{{{MAIN_NS}}}c",
            {"r": f"B{row_number}", "s": template_cells["B"].attrib.get("s", "189")},
        )
        serial_value = ET.SubElement(serial_cell, f"{{{MAIN_NS}}}v")
        serial_value.text = str(row_number - 2)
        new_row.append(serial_cell)

        column_values = {
            "C": attendee.name,
            "D": attendee.institute,
            "E": attendee.hub,
            "F": attendee.emailId,
            "G": attendee.tag,
        }
        for column, value in column_values.items():
            shared_index = _next_shared_string_index(
                shared_root, shared_values, shared_lookup, value
            )
            cell = ET.Element(
                f"{{{MAIN_NS}}}c",
                {
                    "r": f"{column}{row_number}",
                    "s": template_cells[column].attrib.get("s", "189"),
                    "t": "s",
                },
            )
            cell_value = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
            cell_value.text = str(shared_index)
            new_row.append(cell)

        sheet_data.append(new_row)

        dimension = sheet_root.find("main:dimension", XML_NS)
        if dimension is not None:
            start_ref, _sep, end_ref = dimension.attrib.get("ref", "B2:G2").partition(":")
            end_column = "".join(ch for ch in end_ref if ch.isalpha()) or "G"
            dimension.attrib["ref"] = f"{start_ref}:{end_column}{row_number}"

        _update_shared_string_count(shared_root, len(column_values))

        temp_path = WORKBOOK_PATH.with_suffix(".tmp.xlsx")
        updated_files = {
            f"xl/{sheet_archive_path}": ET.tostring(
                sheet_root, encoding="utf-8", xml_declaration=True
            ),
            "xl/sharedStrings.xml": ET.tostring(
                shared_root, encoding="utf-8", xml_declaration=True
            ),
        }
        with zipfile.ZipFile(temp_path, "w") as target_archive:
            for item in source_archive.infolist():
                data = updated_files.get(item.filename, source_archive.read(item.filename))
                target_archive.writestr(item, data)

    temp_path.replace(WORKBOOK_PATH)


def workbook_bytes() -> bytes:
    return WORKBOOK_PATH.read_bytes()


def import_attendees_from_workbook(file_bytes: bytes) -> dict[str, int]:
    uploaded_records = parse_attendees_from_uploaded_workbook(file_bytes)
    existing_records = parse_attendees_from_workbook()
    existing_keys = {attendee_identity_key(record) for record in existing_records}

    imported_count = 0
    skipped_count = 0
    for record in uploaded_records:
        identity = attendee_identity_key(record)
        if identity in existing_keys:
            skipped_count += 1
            continue

        append_attendee_to_workbook(
            Attendee(
                name=record.get("name", ""),
                tag=record.get("tag", ""),
                institute=record.get("institute", ""),
                hub=record.get("hub", ""),
                emailId=record.get("emailId", ""),
                tagColor=record.get("tagColor", ""),
                tagBorderColor=record.get("tagBorderColor", ""),
                tagTextColor=record.get("tagTextColor", ""),
                source="excel",
            )
        )
        existing_keys.add(identity)
        imported_count += 1

    dataset = build_dataset()
    save_dataset(dataset)
    return {
        "uploaded": len(uploaded_records),
        "imported": imported_count,
        "skipped": skipped_count,
    }


def parse_attendees_from_workbook() -> list[dict[str, Any]]:
    with zipfile.ZipFile(WORKBOOK_PATH) as archive:
        return _parse_attendees_from_archive(archive)


def parse_attendees_from_uploaded_workbook(file_bytes: bytes) -> list[dict[str, Any]]:
    with zipfile.ZipFile(BytesIO(file_bytes)) as archive:
        return _parse_attendees_from_archive(archive)


def attendee_identity_key(record: dict[str, Any]) -> str:
    email_key = normalize_text(record.get("emailId", ""))
    if email_key:
        return f"email:{email_key}"
    return "|".join(
        [
            normalize_text(record.get("name", "")),
            normalize_text(record.get("tag", "")),
            normalize_text(record.get("institute", "")),
            normalize_text(record.get("hub", "")),
        ]
    )


def build_dataset() -> dict[str, Any]:
    excel_records = parse_attendees_from_workbook()
    existing_additions: list[dict[str, Any]] = []

    if DATA_PATH.exists():
        with DATA_PATH.open("r", encoding="utf-8") as file:
            existing = json.load(file)
        existing_additions = existing.get("added_records", [])

    return {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "source_file": WORKBOOK_PATH.name,
        "sheet_name": SHEET_NAME,
        "records": excel_records + existing_additions,
        "added_records": existing_additions,
    }


def save_dataset(dataset: dict[str, Any]) -> None:
    with DATA_PATH.open("w", encoding="utf-8") as file:
        json.dump(dataset, file, indent=2, ensure_ascii=False)


def load_or_create_dataset(force_refresh: bool = False) -> dict[str, Any]:
    if force_refresh or not DATA_PATH.exists():
        dataset = build_dataset()
        save_dataset(dataset)
        return dataset

    with DATA_PATH.open("r", encoding="utf-8") as file:
        dataset = json.load(file)

    if any(
        record.get("source") == "excel"
        and (
            "tagColor" not in record
            or "tagBorderColor" not in record
            or "tagTextColor" not in record
        )
        for record in dataset.get("records", [])
    ):
        dataset = build_dataset()
        save_dataset(dataset)

    return dataset


def add_attendee(
    name: str,
    tag: str,
    institute: str = "",
    hub: str = "",
    email: str = "",
    update_excel: bool = False,
) -> dict[str, Any]:
    attendee_model = Attendee(
        name=name.strip(),
        tag=tag.strip(),
        institute=institute.strip(),
        hub=hub.strip(),
        emailId=email.strip(),
        tagColor="",
        tagBorderColor="",
        tagTextColor="",
        source="excel" if update_excel else "manual",
    )
    attendee = asdict(attendee_model)

    if update_excel:
        append_attendee_to_workbook(attendee_model)
        dataset = build_dataset()
    else:
        dataset = load_or_create_dataset()
        dataset.setdefault("added_records", []).append(attendee)
        dataset["records"] = parse_attendees_from_workbook() + dataset["added_records"]
        dataset["generated_at"] = datetime.now(timezone.utc).isoformat()

    save_dataset(dataset)
    return attendee


def search_attendees(query: str, records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    query = query.strip()
    if not query:
        return []

    normalized_query = normalize_text(query)
    regex = None
    try:
        regex = re.compile(query, re.IGNORECASE)
    except re.error:
        regex = None

    scored_matches: list[tuple[float, dict[str, Any]]] = []
    for record in records:
        name = record.get("name", "")
        email = record.get("emailId", "")
        haystacks = [name, email]
        normalized_haystacks = [normalize_text(value) for value in haystacks]

        score = 0.0

        if any(normalized_query == text for text in normalized_haystacks if text):
            score = max(score, 100.0)
        if any(normalized_query in text for text in normalized_haystacks if text):
            score = max(score, 85.0)
        if regex and any(regex.search(value) for value in haystacks if value):
            score = max(score, 75.0)

        fuzzy_score = max(
            (SequenceMatcher(None, normalized_query, text).ratio() for text in normalized_haystacks if text),
            default=0.0,
        )
        score = max(score, fuzzy_score * 70)

        if score >= 45:
            scored_matches.append((score, record))

    scored_matches.sort(
        key=lambda item: (
            -item[0],
            normalize_text(item[1].get("name", "")),
            normalize_text(item[1].get("emailId", "")),
        )
    )
    return [record for _, record in scored_matches[:10]]
