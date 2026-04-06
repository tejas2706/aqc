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
) -> dict[str, Any]:
    dataset = load_or_create_dataset()
    attendee = asdict(
        Attendee(
            name=name.strip(),
            tag=tag.strip(),
            institute=institute.strip(),
            hub=hub.strip(),
            emailId=email.strip(),
            source="manual",
        )
    )
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
