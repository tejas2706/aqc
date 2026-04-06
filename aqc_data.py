import json
import re
import zipfile
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET


WORKBOOK_PATH = Path(__file__).resolve().parent / "AQC Planning.xlsx"
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


def parse_attendees_from_workbook() -> list[dict[str, Any]]:
    with zipfile.ZipFile(WORKBOOK_PATH) as archive:
        shared_strings = _read_shared_strings(archive)
        sheet_path = _resolve_sheet_path(archive, SHEET_NAME)
        sheet = ET.fromstring(archive.read(sheet_path))

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

        attendee = Attendee(
            name=(row.get("name", "") or "").replace("\xa0", " ").strip(),
            institute=(row.get("institute", "") or "").replace("\xa0", " ").strip(),
            hub=(row.get("hub", "") or "").replace("\xa0", " ").strip(),
            emailId=(row.get("email", "") or row.get("emailid", "") or "")
            .replace("\xa0", " ")
            .strip(),
            tag=(row.get("tag", "") or "").replace("\xa0", " ").strip(),
            source="excel",
        )

        if attendee.name and attendee.tag:
            records.append(asdict(attendee))

    return records


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
        return json.load(file)


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
