from __future__ import annotations

import argparse
import json
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment

try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover
    PdfReader = None


PARTY_MARKERS = {
    "shipper": ["SHIPPER:", "SHIPPER"],
    "consignee": ["CONSIGNEE:", "CONSIGNEE"],
    "notify": ["NOTIFY PARTY:", "NOTIFY PARTY", "NOTIFY:"],
}

STOP_LABELS = {
    "B/L NO",
    "MBL/HBL",
    "NOTE:",
    "VESSEL VOYAGE:",
    "PORT OF LOADING:",
    "PORT OF DISCHARGE:",
    "FINAL DESTINATION:",
    "PLACE OF RECEIPT:",
    "PLACE OF DELIVERY:",
    "MARKS & NOS",
    "DESCRIPTION OF GOODS",
    "NO. OF PKGS",
    "FREIGHT & CHARGES",
    "CONTAINER NO",
    "SEAL NO",
    "GROSS WEIGHT",
    "MEASUREMENT",
    "SAILING ON ABOUT",
    "ON BOARD DATE",
    "PRE-CARRIAGE BY",
    "EXPORT REFERENCES",
}

PARTY_FIELD_MAP = {
    "shipper": "发货人",
    "consignee": "收货人",
    "notify": "通知人",
}

HEADERS = [
    "提单号",
    "发货人",
    "收货人",
    "通知人",
    "文件源",
    "文件路径",
    "来源文件夹",
    "发货人备注",
    "收货人备注",
    "通知人备注",
]


def normalize_text(text: str) -> str:
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\u3000", " ").replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()


def canonical_marker(text: str) -> str:
    text = normalize_text(text).upper()
    return re.sub(r"^\d+\s*[.)]?\s*", "", text)


def has_company_keyword(text: str) -> bool:
    upper = normalize_text(text).upper()
    return bool(re.search(r"\b(LLC|LTD|CO|CORP|INC|OOO|LIMITED|COMPANY|PTE)\b", upper))


def is_junk_line(text: str) -> bool:
    text = normalize_text(text)
    if not text:
        return True
    upper = text.upper()
    if upper in STOP_LABELS:
        return True
    if any(marker in upper for names in PARTY_MARKERS.values() for marker in names):
        return True
    if upper.startswith(("TEL", "PHONE", "MOB", "FAX", "INN", "TIN", "VAT", "ZIP", "EMAIL", "E-MAIL", "ATTN")):
        return True
    if re.fullmatch(r"[0-9A-Z+/()\-]{8,}", upper) and not re.search(r"[A-Z]{3,}.*\s", upper):
        return True
    return False


def looks_like_company_name(text: str) -> bool:
    text = normalize_text(text)
    text = re.sub(r"^(NAME|COMPANY|CONSIGNEE NAME|SHIPPER NAME|NOTIFY NAME)\s*:\s*", "", text, flags=re.I)
    upper = text.upper()
    if is_junk_line(text):
        return False
    if len(re.findall(r"[A-Za-z]", text)) < 3:
        return False
    if "@" in text:
        return False
    if re.search(r"\bXHLTC\d+\b", upper):
        return False
    if any(token in upper for token in ["XIN HE LU", "IN CHINESE", "PRE-CARRIAGE", "VOY. NO."]):
        return False
    if any(token in upper for token in ["MARK & NUMBERS", "MARKS & NUMBERS", "PKGS", "CONTAINER NO./SEAL NO."]):
        return False
    if re.match(r"^\d+\.\s", text):
        return False
    if text[:1].isdigit() and "," in text:
        return False
    if any(upper.startswith(prefix) for prefix in ["NO.", "TEL", "PHONE", "FAX", "ADD:", "ADDRESS:"]):
        return False
    if " " not in text and not has_company_keyword(text):
        return False
    return True


def extract_binary_strings(path: Path) -> list[str]:
    data = path.read_bytes()
    items: list[tuple[int, str]] = []
    for match in re.finditer(rb"[ -~]{4,}", data):
        items.append((match.start(), match.group().decode("latin1", "ignore")))
    for match in re.finditer(rb"(?:[ -~]\x00){4,}", data):
        items.append((match.start(), match.group().decode("utf-16le", "ignore")))
    items.sort(key=lambda item: item[0])
    return [normalized for _, text in items if (normalized := normalize_text(text))]


def extract_pdf_lines(path: Path) -> list[str]:
    if PdfReader is None:
        raise RuntimeError("pypdf is required for PDF extraction")
    reader = PdfReader(str(path))
    lines: list[str] = []
    for page in reader.pages:
        for line in (page.extract_text() or "").splitlines():
            if normalized := normalize_text(line):
                lines.append(normalized)
    return lines


def extract_bl_no(lines: list[str], fallback_name: str) -> str:
    joined = "\n".join(lines)
    match = re.search(r"\bXHLTC\d+\b", joined, flags=re.I)
    if match:
        return match.group(0).upper()
    match = re.search(r"XHLTC\d+", fallback_name, flags=re.I)
    return match.group(0).upper() if match else Path(fallback_name).stem


def extract_from_lines(lines: list[str], filename: str) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    markers_flat = [marker for names in PARTY_MARKERS.values() for marker in names]
    canonical_lines = [canonical_marker(line) for line in lines]
    bl_no = extract_bl_no(lines, filename)

    for party_type, markers in PARTY_MARKERS.items():
        for idx, current in enumerate(canonical_lines):
            if current not in markers:
                continue
            block: list[str] = []
            for nxt in range(idx + 1, min(len(lines), idx + 15)):
                candidate = normalize_text(lines[nxt])
                if not candidate:
                    continue
                if canonical_marker(candidate) in markers_flat:
                    break
                block.append(candidate)

            company = ""
            for candidate in block:
                cleaned = re.sub(
                    r"^(NAME|COMPANY|CONSIGNEE NAME|SHIPPER NAME|NOTIFY NAME)\s*:\s*",
                    "",
                    normalize_text(candidate),
                    flags=re.I,
                )
                if looks_like_company_name(cleaned):
                    company = cleaned
                    break

            results.append(
                {
                    "file": filename,
                    "bl_no": bl_no,
                    "party_type": party_type,
                    "company": company,
                    "raw_block": block,
                }
            )
    return results


def infer_missing_notify(entries: list[dict[str, Any]], all_lines_by_source: dict[str, list[str]]) -> list[dict[str, Any]]:
    by_source: dict[str, list[dict[str, Any]]] = {}
    for entry in entries:
        by_source.setdefault(entry["source_path"], []).append(entry)

    for source_path, file_entries in by_source.items():
        consignee = next((entry for entry in file_entries if entry["party_type"] == "consignee"), None)
        notify = next((entry for entry in file_entries if entry["party_type"] == "notify"), None)
        if not notify or notify.get("company"):
            continue

        source_lines = all_lines_by_source.get(source_path, [])
        lower_source = source_path.lower()
        if lower_source.endswith(".pdf"):
            for index, line in enumerate(source_lines):
                if "NOTIFY@" not in line.upper():
                    continue
                preferred = ""
                fallback = ""
                for prev in range(index - 1, max(-1, index - 8), -1):
                    candidate = normalize_text(source_lines[prev])
                    if has_company_keyword(candidate) and looks_like_company_name(candidate):
                        preferred = candidate
                        break
                    if looks_like_company_name(candidate) and not candidate.upper().startswith(("REG ", "VAT ", "E-MAIL", "TEL", "INN")):
                        fallback = fallback or candidate
                if preferred or fallback:
                    notify["company"] = preferred or fallback
                    notify["inferred"] = "from notify email block"
                    break
        elif consignee and consignee.get("company"):
            notify["company"] = consignee["company"]
            notify["inferred"] = "same as consignee"

    return entries


def iter_bill_files(root: Path) -> list[Path]:
    files: list[Path] = []
    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue
        if path.name.startswith("~$"):
            continue
        if path.suffix.lower() not in {".xls", ".xlsx", ".pdf"}:
            continue
        files.append(path)
    return files


def collect_entries(root: Path) -> list[dict[str, Any]]:
    entries: list[dict[str, Any]] = []
    all_lines_by_source: dict[str, list[str]] = {}
    for path in iter_bill_files(root):
        if path.suffix.lower() == ".pdf":
            lines = extract_pdf_lines(path)
        else:
            lines = extract_binary_strings(path)
        all_lines_by_source[str(path)] = lines
        file_entries = extract_from_lines(lines, path.name)
        for entry in file_entries:
            entry["source_path"] = str(path)
            entry["source_folder"] = str(path.parent)
        entries.extend(file_entries)
    return infer_missing_notify(entries, all_lines_by_source)


def build_rows(entries: list[dict[str, Any]]) -> list[dict[str, str]]:
    grouped: dict[tuple[str, str], dict[str, str]] = {}
    for entry in entries:
        key = (entry["bl_no"], entry["source_path"])
        row = grouped.setdefault(
            key,
            {
                "提单号": entry["bl_no"],
                "发货人": "",
                "收货人": "",
                "通知人": "",
                "文件源": Path(entry["source_path"]).name,
                "文件路径": entry["source_path"],
                "来源文件夹": entry["source_folder"],
                "发货人备注": "",
                "收货人备注": "",
                "通知人备注": "",
            },
        )
        field = PARTY_FIELD_MAP.get(entry["party_type"])
        if not field:
            continue
        row[field] = normalize_text(entry.get("company") or "")
        inferred = normalize_text(str(entry.get("inferred", "")))
        if inferred:
            row[f"{field}备注"] = inferred

    return sorted(grouped.values(), key=lambda item: (item["提单号"], item["文件源"]))


def write_json(rows: list[dict[str, str]], path: Path) -> None:
    path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")


def write_excel(rows: list[dict[str, str]], path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "收发通"
    sheet.append(HEADERS)
    for row in rows:
        sheet.append([row.get(header, "") for header in HEADERS])

    for cell in sheet[1]:
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    widths = {
        "A": 16,
        "B": 34,
        "C": 34,
        "D": 34,
        "E": 28,
        "F": 78,
        "G": 32,
        "H": 18,
        "I": 18,
        "J": 18,
    }
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width

    workbook.save(path)


def output_dir_for(input_path: Path, override: Path | None) -> Path:
    if override:
        return override
    return input_path if input_path.is_dir() else input_path.parent


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--excel", type=Path)
    parser.add_argument("--json", dest="json_path", type=Path)
    args = parser.parse_args()

    input_path = args.input
    root = input_path if input_path.is_dir() else input_path.parent
    rows = build_rows(collect_entries(root))
    timestamp = datetime.now().strftime("%Y%m%d-%H%M")
    output_dir = output_dir_for(input_path, args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    excel_path = args.excel or output_dir / f"收发通_提取信息_{timestamp}.xlsx"
    json_path = args.json_path or output_dir / f"收发通_提取结果_{timestamp}.json"

    write_excel(rows, excel_path)
    write_json(rows, json_path)

    print(excel_path)
    print(json_path)
    print(f"records={len(rows)}")


if __name__ == "__main__":
    main()
