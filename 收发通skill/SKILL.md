---
name: 收发通skill
description: Extract bill-of-lading shipper, consignee, and notify-party headers from Excel or PDF files and export them to Excel. Use when the user asks to抓取收发通、提取提单抬头、输出收发通Excel、整理发货人/收货人/通知人, or provides a folder of bill files for header extraction only.
---

# 收发通skill

## Purpose

Use this skill to scan bill files, extract `发货人`、`收货人`、`通知人`, and export a clean Excel workbook without doing any website lookup.

## Quick Start

Resolve script paths relative to this skill directory. Do not use a hard-coded
user home path, because Codex may be running on macOS or Windows.

macOS/Linux:

```bash
python3 scripts/export_receipt_parties.py --input /path/to/folder
```

Windows PowerShell:

```powershell
python .\scripts\export_receipt_parties.py --input "C:\path\to\folder"
```

If `python` is not available on Windows, use `py` instead.

Default outputs:

- `收发通_提取信息_YYYYMMDD-HHMM.xlsx`
- `收发通_提取结果_YYYYMMDD-HHMM.json`

## Workflow

1. Scan the target folder recursively for `.xls`, `.xlsx`, and `.pdf` bill files.
2. Extract the B/L number plus `发货人`、`收货人`、`通知人`.
3. Normalize values with Unicode `NFKC`, ASCII spaces, and collapsed whitespace.
4. Infer `通知人` from nearby context when the bill visually implies `same as consignee` or a PDF notify-email block.
5. Export the result workbook and return its clickable path.

## Extraction Rules

- Prefer company-like lines containing keywords such as `LLC`, `LTD`, `CO`, `CORP`, `INC`, `LIMITED`, `COMPANY`, `OOO`, or `PTE`.
- Ignore phone, email, tax, registration, port, vessel, cargo, and container lines as company names.
- Never leave `通知人` blank after the first pass; retry the local context and infer when justified.
- Keep source provenance for each row so later query workflows can trace the original bill file.

## Output Layout

The Excel workbook should start with these business columns first:

1. `提单号`
2. `发货人`
3. `收货人`
4. `通知人`
5. `文件源`

Then include supporting provenance fields such as `文件路径`, `来源文件夹`, and any inference notes.

Final response after using this skill should return the workbook path and a short count summary.
