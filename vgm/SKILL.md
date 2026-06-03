---
name: vgm
description: Integrate received VGM container data into the correct terminal-upload Excel template for transit cargo (中转模板) or local-port cargo (本港模板), preserving blanks, trimming boundary whitespace, validating verification methods, red-marking invalid cells, and formatting all transferred values as text with ASCII/half-width punctuation. Use when the user supplies a VGM file, requests VGM整理/整合/上传模板/码头比对, or needs an `.xlsx`, `.xls`, or `.csv` VGM source converted into the terminal template.
---

# VGM

## Purpose

Convert a received VGM source file into the correct upload workbook for terminal container-information comparison. Preserve source facts exactly except for required half-width symbol normalization, boundary-whitespace cleanup, and explicit validation marking.

## Workflow

1. Identify every supplied VGM source file and determine whether the upload requires `中转模板` or `本港模板`.
2. Use the user's current template if one is supplied. Otherwise copy the matching bundled asset:
   - `assets/中转模板.xlsx`
   - `assets/本港模板.xlsx`
3. Read [references/template-schema.md](references/template-schema.md) before mapping fields or normalizing cell values.
4. Use the spreadsheet artifact workflow to edit a copy of the selected template; retain sheet name `VGM申报填写`, its header row, layout, and formatting.
5. Write one output data row for each source container record, beginning below the existing header row. Map only fields represented in the selected template.
6. Validate every output row after writing. If a populated value violates a hard rule, keep the source value visible where possible and mark the invalid output cell with red font/fill so the user can audit it.
7. Save a completed `.xlsx` in the user's source folder, preferably named `VGM_中转_整合结果_YYYYMMDD-HHMM.xlsx` or `VGM_本港_整合结果_YYYYMMDD-HHMM.xlsx`.
8. Return the output workbook link with the selected template type, input/output row counts, validation issues, and any fields that could not be mapped confidently.

## Portability

- This skill must work in Codex on both macOS and Windows. Resolve bundled resources relative to this skill directory; do not hard-code `/Users/...`, drive letters, or OS-specific install paths.
- Use `assets/中转模板.xlsx`, `assets/本港模板.xlsx`, and `references/template-schema.md` by joining them to the directory containing this `SKILL.md`.
- Treat GUI/mail-collection steps as environment-specific helpers. The NetEase MailMaster (`网易邮箱大师`) database and macOS Accessibility/Screen Recording notes below apply only when that app and OS are present; on Windows, use the available mail client/filesystem equivalents and keep the same B/L inventory and skip-duplicate logic.

## Mail Attachment Collection

When VGM files must be collected from mail folders before integration:

- Do not rely on the currently visible message list. First inventory the local destination folder and the mail source, then compare them by B/L number.
- For NetEase MailMaster (`网易邮箱大师`), inspect the local mail database when available before using the GUI. The useful tables are commonly `Mailbox`, `MailMeta`, and `MailAttachment`; match only the requested mail folders such as `宁波截单` and `大连截单`.
- Treat a MailMaster attachment as locally available only when `MailAttachment.LocalPath` is nonblank and `TransferredSize` is greater than zero. Cached files are normally under the account data directory's `parts/<LocalPath>` path and can be copied into the corresponding `截单/<提单号>/` folder while preserving the original attachment name.
- If `LocalPath` is blank or `TransferredSize` is zero, the mail database has metadata only; use the MailMaster GUI to trigger the attachment download.
- In MailMaster, use the small search field in the upper-left mail-list toolbar for B/L searches. Avoid macOS/large top overlay search or advanced search unless specifically needed; those search surfaces can focus the wrong UI and waste time.
- Before GUI automation, verify macOS Accessibility and Screen Recording permissions for the automation host. If screenshots are black or clicks do not land, stop and request the missing permission instead of continuing blindly.
- Skip downloading a B/L only when the destination folder already contains a matching VGM source file, not merely because a B/L folder exists.
- After collection, rerun the source inventory and compare mail VGM attachment B/L numbers against local VGM source B/L numbers. Explicitly report any remaining uncached or unavailable attachments.

## Template Choice

- Use `中转模板` only when the user identifies the shipment as 中转/transit or supplies that template for the task.
- Use `本港模板` only when the user identifies the shipment as 本港/local port or supplies that template for the task.
- If the template type is not stated and cannot be established from supplied material, ask the user before producing an upload file. Do not infer from extra columns alone.
- Do not combine records that require different templates into one workbook.

## Data Rules

- Treat all output data cells as text, including VGM weight, voyage, B/L number, container number, size, and codes. Set the filled data region's Excel number format to text (`@`) and write string values so leading zeros remain intact.
- Normalize transferred nonblank values to English half-width form as described in the schema reference. Preserve Chinese wording; normalize symbols and full-width ASCII variants only.
- Remove leading and trailing whitespace from populated transferred values after half-width normalization. If a cell becomes empty after this cleanup, write a real blank cell rather than a string containing spaces.
- For `vgm总量`, remove any unit suffix from the source value before writing the output. If the source is entered as `23486KGS`, `23486 KGS`, `23486kg`, or similar, keep only the weight value text and do not carry the unit into the template.
- For `验证方式（不填写默认累加计算）`, write only a blank, `累加计算`, or `整体称重`. If the source contains another value such as `累计计算`, keep the cell visible for audit but mark it red and report it as a validation issue.
- Preserve all source blanks. If a mapped source field is absent, empty, whitespace-only, or visually blank, write an empty cell in the output.
- Never populate a blank with `0`, `N/A`, `-`, a calculated value, a portal default, or an inferred value.
- Never calculate, round, convert, sum, or otherwise alter `vgm总量` beyond removing a unit suffix when one is present in the source text.
- Do not populate template fields absent from the source even when a header note says an omitted value will default automatically on upload.
- Do not modify the source VGM file.

## Verification

- Verify the selected template and confirm the expected headers before filling.
- Reconcile output records against the source using `提单号` and `箱号` where available; the number of container rows must match the mapped source records.
- Check every mapped blank remains blank and every populated output data cell is text-formatted.
- Compare nonblank values against source values after only the permitted half-width normalization, boundary-whitespace cleanup, and `vgm总量` unit stripping.
- Confirm no populated output value starts or ends with whitespace. Confirm whitespace-only values are real blank cells.
- Confirm `验证方式（不填写默认累加计算）` contains only blanks, `累加计算`, or `整体称重`; any other populated value must be marked red in the workbook and listed in the response.
- Inspect the populated range for formula errors and render the final worksheet once for visual confirmation before delivery.

## Resources

Resolve bundled resources relative to this skill directory so the same skill works from both macOS and Windows Codex installations.

- `assets/中转模板.xlsx`: transit upload workbook template.
- `assets/本港模板.xlsx`: local-port upload workbook template.
- [references/template-schema.md](references/template-schema.md): column definitions, field matching, and normalization requirements.
