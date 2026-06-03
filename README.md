# Codex Skills Portable Bundle

This repository contains three Codex skills:

- `收发通skill`: extract bill-of-lading shipper, consignee, and notify-party headers from Excel or PDF files.
- `查询比对skill`: query EB customer records and compare them against extracted bill parties.
- `vgm`: integrate received VGM container data into the correct terminal upload workbook.

## Install

Copy the skill folders into the Codex skills directory for the target machine.

macOS:

```bash
cp -R vgm "$HOME/.codex/skills/"
cp -R "收发通skill" "$HOME/.codex/skills/"
cp -R "查询比对skill" "$HOME/.codex/skills/"
```

Windows PowerShell:

```powershell
Copy-Item -Recurse .\vgm "$env:USERPROFILE\.codex\skills\"
Copy-Item -Recurse ".\收发通skill" "$env:USERPROFILE\.codex\skills\"
Copy-Item -Recurse ".\查询比对skill" "$env:USERPROFILE\.codex\skills\"
```

## Python Dependencies

`收发通skill` and `查询比对skill` use local Python helper scripts. Install the
dependencies into the Python environment Codex will use:

```bash
python -m pip install -r requirements.txt
```

On Windows, use `py -m pip install -r requirements.txt` if `python` is not on
`PATH`.

The scripts use cross-platform Python APIs (`pathlib`, `openpyxl`, `pypdf`,
`urllib`, and optional `curl`). No macOS-only shell commands or absolute user
paths are required.

## Cross-Platform Notes

- Keep paths relative to each skill folder when invoking bundled scripts or
  resources.
- The VGM templates are binary `.xlsx` assets and should be committed as-is.
- Unicode folder names and workbook content are expected; use UTF-8 capable Git
  clients and terminals.
