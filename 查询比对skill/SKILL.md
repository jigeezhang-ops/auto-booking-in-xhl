---
name: 查询比对skill
description: Query EB customer records from the website backend and compare them against extracted bill parties in an Excel workbook. Use when the user asks to登录网址查询、EB客户信息比对、查询公司抬头是否存在、批量回写查询结果, or already has a 收发通提取Excel that needs website validation.
---

# 查询比对skill

## Purpose

Use this skill only for the website lookup and comparison stage after a party-extraction workbook already exists.

## Quick Start

Resolve script paths relative to this skill directory. Do not hard-code a macOS
home path; the skill may run on Windows.

macOS/Linux:

```bash
python3 scripts/query_eb_customers.py --workbook /path/to/收发通_提取信息.xlsx --account /path/to/账号测试.txt
```

Windows PowerShell:

```powershell
python .\scripts\query_eb_customers.py --workbook "C:\path\to\收发通_提取信息.xlsx" --account "C:\path\to\账号测试.txt"
```

If the user asks for account 2, pass `--account-index 2`. The script accepts
both `账号/密码` and numbered account files such as `账号2/密码2`.

Default output:

- `查询比对_结果_YYYYMMDD-HHMM.xlsx`

## Workflow

1. Load the extraction workbook produced by `收发通skill`, or a compatible workbook with `提单号`、`发货人`、`收货人`、`通知人`.
2. Read the account text file and log in to the site backend. Use `--account-index N` when the user specifies 账号N.
3. Query `EB客户信息` by full company name first, then by strong tokens if needed.
4. For names containing punctuation or brackets, query normalized variants too. Example: `JSC <<ELIS FASHION RUS>>` may match `JSC "ELIS FASHION RUS"` or `JSC ELIS FASHION RUS`.
5. Match by normalized company text, ignoring punctuation and `(EB00123400)` suffixes.
6. Write the query result back to the workbook and generate `可查询到` and `查询不到` sheets.

## Access Rules

- Use backend requests or browser automation first. Do not default to `computer-use` or manual front-end clicking.
- If Python `urllib` has DNS or timeout failures, retry with `--curl-fallback` before switching to front-end control. `curl` is available on modern Windows as `curl.exe`; if it is missing, the script reports the fallback failure and continues to mark rows as `查询异常`.
- If DNS is flaky but an IP was already verified, pass `--resolve-ip <ip>` with `--curl-fallback` so curl uses `--resolve host:port:ip`.
- Screenshots are feedback artifacts only; they do not justify switching back to front-end control.
- If individual company queries time out or fail, mark that row as `查询异常` and continue the batch instead of aborting the whole run.

## Matching Rules

- Normalize with Unicode `NFKC`, ASCII spaces, and collapsed whitespace before every query.
- Query the full company first.
- Skip weak fallback tokens such as `LLC`, `LTD`, `CO`, `CORP`, `OOO`, `THE`, `AND`, `JSC`, and `IP`.
- Do not treat broad location or business words such as `QINGDAO`, `NINGBO`, `CIXI`, `CHINA`, `INTERNATIONAL`, or `TRADE` as strong evidence by themselves.
- Aggregate roles from `发货人`、`收货人`、`通知人` when multiple equivalent matches exist.

## Output

Keep the extraction columns at the front, then append query columns such as:

- `查询状态`
- `查询说明`
- `发货人匹配名称` / `匹配编码` / `匹配角色` / `查询词`
- `收货人...`
- `通知人...`

Final response after using this skill should return the output workbook path plus hit/miss counts.
