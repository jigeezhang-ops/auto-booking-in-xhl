# Template Schema And Normalization

## Bundled Templates

Both templates contain the worksheet `VGM申报填写`. Keep row 1 unchanged and write records from row 2 onward.

### 中转模板

| Column | Header | Common Source Match |
| --- | --- | --- |
| A | 箱类型（不填写默认中转箱） | 箱类型, container category |
| B | 船名 | 船名, vessel, vessel name |
| C | 航次 | 航次, voyage, voyage no. |
| D | 船公司（自行填写或不填写默认自动匹配） | 船公司, carrier, shipping line |
| E | 提单号 | 提单号, B/L no., BL no., bill of lading no. |
| F | 箱号 | 箱号, container no., container number |
| G | 箱尺寸 | 箱尺寸, size, container size |
| H | 箱型 | 箱型, type, container type |
| I | vgm总量 | VGM, VGM总量, verified gross mass |
| J | 备注 | 备注, remark, remarks |
| K | 验证方式（不填写默认累加计算） | 验证方式, verification method |
| L | 授权方 | 授权方, authorized party |
| M | 责任方 | 责任方, responsible party |

### 本港模板

| Column | Header | Common Source Match |
| --- | --- | --- |
| A | 提单号 | 提单号, B/L no., BL no., bill of lading no. |
| B | 箱号 | 箱号, container no., container number |
| C | vgm总量 | VGM, VGM总量, verified gross mass |
| D | 备注 | 备注, remark, remarks |
| E | 验证方式（不填写默认累加计算） | 验证方式, verification method |
| F | 授权方 | 授权方, authorized party |
| G | 责任方 | 责任方, responsible party |

Use matching labels as guidance, not permission to guess. If two source fields could map to one output field or a required field is unclear, flag it for the user rather than silently choosing.

## Text And Blank Policy

- Store every transferred field as text and set output data cells to text number format (`@`).
- Do not parse numeric-looking strings into numbers. This includes `vgm总量`, dates, voyage numbers, codes, container numbers, and B/L numbers.
- For `vgm总量`, strip any trailing unit text before writing the output cell. Remove units such as `KGS`, `KG`, `kgs`, `kg`, and the same values after half-width normalization, while preserving the weight text itself exactly.
- Map an input blank to a blank output cell. Treat `null`, empty strings, whitespace-only cells, and displayed empty formula results as blank.
- Preserve an unprovided output field as blank; do not use text such as `N/A`, `NULL`, `-`, or `0`.

## Half-Width Normalization

Apply normalization only to nonblank transferred cell text:

1. Convert full-width ASCII letters, digits, spaces, and punctuation to their ASCII half-width equivalents, for example `ＡＢＣ１２３：／（ ）` to `ABC123:/()`.
2. Replace Chinese punctuation used as separators with ASCII equivalents when present in field values:

| Source | Output |
| --- | --- |
| `，` | `,` |
| `。` | `.` |
| `：` | `:` |
| `；` | `;` |
| `（` / `）` | `(` / `)` |
| `【` / `】` | `[` / `]` |
| `／` | `/` |
| `－` | `-` |
| `＃` | `#` |

3. Convert non-breaking or full-width spaces to ASCII spaces. Do not add content to a blank cell.
4. Keep Chinese words and names as written; this policy concerns symbol width, not translation or rewriting.

Record any normalization that materially changes a value the user may need to audit.
