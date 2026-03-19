# xlsx-writer

A command-line tool that generates formatted Excel (.xlsx) spreadsheets from
simple text commands piped via stdin.

## Installation

```bash
cargo build --release
cp target/release/xlsx-writer /usr/local/bin/
```

Or run directly from the project:

```bash
cargo run --release -- output.xlsx < commands.txt
```

## How It Works

`xlsx-writer` reads one command per line from stdin. Each line starts with a
command name followed by space-separated arguments. The tool echoes every line
to stdout as it processes it, making it easy to debug pipelines.

```
echo "ADD_WORKSHEET Sheet1
TEXT A1 Hello World" | xlsx-writer output.xlsx
```

Unknown commands are silently ignored, so you can use comment-like lines
(e.g. starting with `#`) without errors.

## Commands

### ADD_WORKSHEET [title ...]

Add a new worksheet. All subsequent cell operations apply to this worksheet.

```
ADD_WORKSHEET Sales Report Q4
```

### TEXT [.styles] cell text ...

Write a text string to a cell (A1 notation).

```
TEXT A1 Hello World
TEXT .bold:1 .color:red B2 Important note
```

### NUM [.styles] cell number

Write a number to a cell (A1 notation).

```
NUM A1 42
NUM .num_format:$#,##0.00 B1 1234.56
```

### FORMULA [.styles] cell expression

Write an Excel formula to a cell. The formula is passed as-is to Excel.

```
FORMULA A5 =SUM(A1:A4)
FORMULA .bold:1 .num_format:$#,##0 B5 =B1+B2+B3+B4
FORMULA .num_format:0.0% C5 =(C4-C1)/C1
FORMULA D1 =IF(A1>100,"High","Low")
FORMULA E1 =VLOOKUP(A1,Sheet2!A:B,2,FALSE)
```

### ROW [.styles] start_cell values...

Write an entire row of values starting at the given cell. Values are
auto-detected as numbers (if parseable as float), formulas (if starting with
`=`), or strings.

When the input contains **tab characters**, fields are split on tabs —
preserving spaces within values. Without tabs, fields are split on whitespace.
This makes ROW ideal for piping from SQL, CSV tools, or `awk`.

```
ROW A1 Name Age City Score
ROW .bold:1 .bg_color:navy .color:white A1 Name Age City Score
```

With tab-delimited data from a pipeline:

```bash
psql -t -A -F $'\t' -c "SELECT name, age, city FROM users" mydb \
  | awk -F'\t' '{print "ROW A" NR+1 "\t" $0}' \
  | xlsx-writer users.xlsx
```

### FAST row col style_name text ...

High-performance write using 0-based row/col and a **named** style (defined
with `STYLE`). Automatically writes as a number if the text starts with a
digit, otherwise as a string. Designed for bulk data output.

```
STYLE data num_format 0.00
FAST 0 0 data 100.5
FAST 0 1 data Hello
```

### URL [.styles] cell url [title ...]

Write a clickable hyperlink.

```
URL A1 https://example.com
URL .color:blue .underline:1 A2 https://example.com Click here
```

### DATE [.styles] cell date

Write a date/datetime value. Dates should be in ISO 8601 format
(`YYYY-MM-DD` or `YYYY-MM-DDTHH:MM:SS`).

```
DATE .num_format:yyyy-mm-dd A1 2024-06-15
DATE .num_format:yyyy-mm-dd\ hh:mm A2 2024-06-15T14:30:00
```

### MERGE [.styles] range text ...

Merge a range of cells and write centered text.

```
MERGE .bold:1 .align:center .border:1 A1:D1 Quarterly Report
```

### BLANK [.styles] cell

Write a formatted empty cell. Useful for applying background colors or borders
to cells without content.

```
BLANK .bg_color:yellow .border:1 A5
BLANK .bg_color:#E8E8E8 .border:2 B10
```

### COMMENT cell text...

Add a note/comment to a cell. The note appears as a popup when hovering over
the cell in Excel.

```
COMMENT A1 This value is estimated
COMMENT B5 Includes $8K for Q4 campaign — approved by VP Marketing
```

### FREEZE cell

Freeze panes at the given cell reference. Everything above and to the left of
this cell will be frozen (stay visible when scrolling).

```
FREEZE A2       # Freeze the top row
FREEZE B3       # Freeze top 2 rows and first column
FREEZE A1       # No freeze (no rows/cols above A1)
```

### AUTOFILTER range

Add dropdown filter buttons to the header row of a range.

```
AUTOFILTER A1:D100
AUTOFILTER A1:F1
```

### TABLE range [name] [style]

Convert a range into an Excel Table object with banded rows, auto-headers,
and built-in filter dropdowns. Data must already be written to the range before
applying the table.

```
TABLE A1:D10
TABLE A1:D10 SalesData
TABLE A1:D10 SalesData Medium10
TABLE A1:B4 _ Light6        # _ = auto-name
```

**Available styles:** `None`, `Light1`–`Light21`, `Medium1`–`Medium28`,
`Dark1`–`Dark11`

### CONDITIONAL [.styles] range type [criteria] [values...]

Apply conditional formatting to a range. The inline styles (`.bg_color:red`,
etc.) define the format applied when the condition is met.

**Cell value rules:**

```
CONDITIONAL .bg_color:green B2:B10 cell greater_than 90
CONDITIONAL .bg_color:red B2:B10 cell less_than 60
CONDITIONAL .bg_color:yellow B2:B10 cell between 60 75
CONDITIONAL .bold:1 B2:B10 cell equal_to 100
CONDITIONAL .bg_color:pink B2:B10 cell not_between 25 75
```

Supported criteria: `equal_to`, `not_equal_to`, `greater_than`,
`greater_than_or_equal_to`, `less_than`, `less_than_or_equal_to`, `between`,
`not_between`

**Duplicate/unique detection:**

```
CONDITIONAL .bg_color:#FFC7CE A2:A100 duplicate
CONDITIONAL .bg_color:#C6EFCE A2:A100 unique
```

**Blank/non-blank cells:**

```
CONDITIONAL .bg_color:gray A2:A100 blank
CONDITIONAL .bg_color:white A2:A100 not_blank
```

**Formula-based rules:**

```
CONDITIONAL .bg_color:red A2:A100 formula =A2>B2
```

**Top/bottom N:**

```
CONDITIONAL .bg_color:green B2:B11 top 3
CONDITIONAL .bg_color:red B2:B11 bottom 5
CONDITIONAL .bg_color:green B2:B11 top_percent 10
CONDITIONAL .bg_color:red B2:B11 bottom_percent 25
```

**Color scales (no inline styles needed):**

```
CONDITIONAL A2:A100 2_color_scale
CONDITIONAL A2:A100 2_color_scale red green
CONDITIONAL A2:A100 3_color_scale
CONDITIONAL A2:A100 3_color_scale red white blue
```

**Data bars (no inline styles needed):**

```
CONDITIONAL A2:A100 data_bar
CONDITIONAL A2:A100 data_bar blue
CONDITIONAL A2:A100 data_bar #FF6600
```

### DATA_VALIDATION range type [values/criteria...]

Add data validation constraints to cells.

**Dropdown list:**

```
DATA_VALIDATION B2:B100 list Yes,No,Maybe
DATA_VALIDATION C2:C100 list High Priority,Medium Priority,Low Priority
```

**Whole number (integer) validation:**

```
DATA_VALIDATION D2:D100 whole_number between 1 100
DATA_VALIDATION D2:D100 integer greater_than 0
DATA_VALIDATION D2:D100 whole_number less_than_or_equal 999
```

**Decimal number validation:**

```
DATA_VALIDATION E2:E100 decimal between 0.0 100.0
DATA_VALIDATION E2:E100 float greater_than 0.0
```

**Text length validation:**

```
DATA_VALIDATION F2:F100 text_length less_than_or_equal 200
DATA_VALIDATION F2:F100 text_length between 5 50
```

Supported criteria for number/text validations: `equal_to`, `not_equal_to`,
`greater_than`, `greater_than_or_equal_to`, `less_than`,
`less_than_or_equal_to`, `between`, `not_between`

### RICH_TEXT [.styles] cell format1|text1||format2|text2||...

Write mixed-format text within a single cell. Segments are separated by `||`.
Within each segment, the format and text are separated by `|`. Format
properties use comma-separated `key:value` pairs. Use `_` for default
formatting.

Optional inline `.styles` set cell-level formatting (background, borders).

```
RICH_TEXT A1 bold:1|Important: ||_|This is normal text.
RICH_TEXT A2 color:red|Red ||color:green|Green ||color:blue|Blue
RICH_TEXT A3 bold:1|Name: ||_|John Smith  ||bold:1|Role: ||italic:1|Engineer
RICH_TEXT .bg_color:yellow .border:1 A4 bold:1|Status: ||color:green|APPROVED
RICH_TEXT A5 _|Email: ||color:blue,underline:1|support@example.com
```

### SETCOL [.styles] col_range width

Set column width (and optionally a default format) for a column range.

```
SETCOL A:A 30
SETCOL .bold:1 B:D 15
```

### SETROW [.styles] row height

Set row height (and optionally a default format). **Row is 1-indexed** (row 1
is the first row).

```
SETROW 1 25
SETROW .bold:1 .bg_color:yellow 1 30
```

### STYLE name key value [key value ...]

Define a **named style** for use with `FAST`. Properties are specified as
space-separated key/value pairs (not dot-notation).

```
STYLE header bold 1 color white bg_color navy font_size 14
STYLE currency num_format $#,##0.00
STYLE percent num_format 0.00%
```

### SET_PROPERTY key value ...

Set a workbook-level document property.

```
SET_PROPERTY title Monthly Sales Report
SET_PROPERTY author Finance Team
SET_PROPERTY company Acme Corp
```

Supported keys: `title`, `author`, `subject`, `manager`, `company`,
`category`, `keywords`, `comment`.

### IMAGE row col filename [x_off y_off [x_scale y_scale]]

Insert an image at a cell position. Row and col are 0-based integers.

```
IMAGE 0 0 logo.png
IMAGE 2 3 chart.png 10 10 0.5 0.5
```

### TAB_COLOR color

Set the worksheet tab color for the current worksheet.

```
TAB_COLOR red
TAB_COLOR #FF6600
TAB_COLOR navy
```

### HIDE_SHEET

Hide the current worksheet. The sheet remains in the workbook but is not
visible in the tab bar. At least one sheet must remain visible.

```
ADD_WORKSHEET Config
HIDE_SHEET
```

### PROTECT_SHEET [password]

Protect the current worksheet from editing. Optionally set a password.

```
PROTECT_SHEET
PROTECT_SHEET mypassword123
```

### DEFINE_NAME name formula

Define a named range at the workbook level. The formula should use absolute
cell references.

```
DEFINE_NAME TaxRate Config!$B$1
DEFINE_NAME SalesData Sheet1!$A$1:$D$100
DEFINE_NAME Threshold 'Sheet With Spaces'!$A$1
```

### PAGE_SETUP key value...

Configure page setup for printing.

```
PAGE_SETUP orientation landscape
PAGE_SETUP paper_size 1
PAGE_SETUP margins 0.5 0.5 0.75 0.75 0.3 0.3
PAGE_SETUP header &CMonthly Report&R&D
PAGE_SETUP footer &LPage &P of &N&R&F
PAGE_SETUP print_scale 75
PAGE_SETUP print_gridlines 1
PAGE_SETUP print_headings 1
PAGE_SETUP print_area A1:G50
PAGE_SETUP repeat_rows 1 1
PAGE_SETUP repeat_columns A:A
PAGE_SETUP center_horizontally 1
PAGE_SETUP center_vertically 1
PAGE_SETUP print_first_page_number 1
```

**Supported keys:**

| Key | Description | Values |
|-----|-------------|--------|
| `orientation` | Page orientation | `landscape`, `portrait` |
| `paper_size` | Paper size code | `1`=Letter, `9`=A4, etc. |
| `margins` | Page margins (inches) | `left right top bottom [header footer]` |
| `header` | Page header text | Excel header codes (`&L`, `&C`, `&R`, `&D`, `&P`, `&N`, `&F`) |
| `footer` | Page footer text | Same codes as header |
| `print_scale` | Print scale percentage | `10`–`400` |
| `print_gridlines` | Print cell gridlines | `1`/`0` |
| `print_headings` | Print row/column headings | `1`/`0` |
| `print_area` | Restrict printed area | Range like `A1:G50` |
| `repeat_rows` | Repeat rows on each page | `first_row last_row` (1-indexed) |
| `repeat_columns` | Repeat columns on each page | Column range like `A:B` |
| `center_horizontally` | Center content horizontally | `1`/`0` |
| `center_vertically` | Center content vertically | `1`/`0` |
| `print_first_page_number` | Starting page number | Integer |

## Inline Styles

Most commands support **inline styles** mixed freely with their arguments.
An inline style token has the form `.property:value` (leading dot, lowercase
property name, colon, value). They can appear anywhere in the argument list.

```
TEXT .bold:1 .font_size:14 .color:white .bg_color:#003366 A1 Header
```

### Supported Style Properties

| Property       | Values / Examples                          |
|----------------|--------------------------------------------|
| `bold`         | `1`                                        |
| `italic`       | `1`                                        |
| `underline`    | `1` (single), `2` (double), `33`, `34`     |
| `strikeout`    | `1`                                        |
| `font_size`    | `12`, `14.5`                               |
| `font`         | `Arial`, `Courier New`                     |
| `color`        | Named: `red`, `blue`, `navy` ... or hex: `#FF0000`, `0xFF0000` |
| `bg_color`     | Same as `color`                            |
| `fg_color`     | Same as `color` (for patterns)             |
| `num_format`   | `0.00`, `$#,##0.00`, `yyyy-mm-dd`, `0.00%` |
| `align`        | `left`, `center`, `right`, `fill`, `justify`, `center_across` |
| `valign`       | `top`, `vcenter`, `bottom`, `vjustify`     |
| `text_wrap`    | `1`                                        |
| `border`       | `0`-`13` (0=none, 1=thin, 2=medium, 5=thick, 6=double, ...) |
| `border_color` | Same as `color`                            |
| `top`          | Border style for top edge                  |
| `bottom`       | Border style for bottom edge               |
| `left`         | Border style for left edge                 |
| `right`        | Border style for right edge                |
| `top_color`    | Color for top border                       |
| `bottom_color` | Color for bottom border                    |
| `left_color`   | Color for left border                      |
| `right_color`  | Color for right border                     |
| `indent`       | Indentation level (integer)                |
| `rotation`     | Text rotation in degrees                   |
| `shrink`       | `1` — shrink text to fit                   |
| `pattern`      | Fill pattern 0-18                          |

### Named Colors

`black`, `blue`, `brown`, `cyan`, `gray`, `green`, `lime`, `magenta`,
`navy`, `orange`, `pink`, `purple`, `red`, `silver`, `white`, `yellow`

Hex colors: `#RRGGBB` or `0xRRGGBB`

## Examples

The `examples/` directory contains ready-to-use input files demonstrating
every feature. Run them individually:

```bash
xlsx-writer output.xlsx < examples/01_simple_table.txt
```

Or generate all examples at once:

```bash
bash examples/run_all.sh
```

### Available Examples

| File | Features Demonstrated |
|------|----------------------|
| `01_simple_table.txt` | TEXT, NUM, SETCOL, inline styles |
| `02_styled_report.txt` | MERGE, SETROW, BLANK, borders, colors, document properties |
| `03_formulas.txt` | FORMULA with SUM, AVERAGE, MIN, MAX, IF |
| `04_frozen_panes.txt` | FREEZE, AUTOFILTER, TAB_COLOR |
| `05_bulk_data_row.txt` | ROW for bulk data entry |
| `06_table_object.txt` | TABLE with named styles (Medium10, Light6) |
| `07_conditional_formatting.txt` | CONDITIONAL: cell rules, color scales, data bars, duplicates, top/bottom |
| `08_data_validation.txt` | DATA_VALIDATION: dropdowns, number ranges, text length |
| `09_page_setup.txt` | PAGE_SETUP: landscape, margins, headers/footers, print area |
| `10_rich_text.txt` | RICH_TEXT: mixed formatting in single cells |
| `11_sheet_management.txt` | TAB_COLOR, HIDE_SHEET, PROTECT_SHEET, DEFINE_NAME |
| `12_comments.txt` | COMMENT (cell notes/annotations) |
| `13_comprehensive.txt` | Full workbook combining most features |

### Quick Start: Simple Table

```bash
echo 'ADD_WORKSHEET Data
SETCOL A:A 20
SETCOL B:B 12
TEXT .bold:1 A1 Name
TEXT .bold:1 B1 Score
TEXT A2 Alice
NUM B2 95
TEXT A3 Bob
NUM B3 87
TEXT A4 Charlie
NUM B4 92' | xlsx-writer scores.xlsx
```

### Styled Report with Merged Header

```bash
cat <<'EOF' | xlsx-writer report.xlsx
SET_PROPERTY title Q4 Sales Report
SET_PROPERTY author Finance Team
ADD_WORKSHEET Q4 Sales
SETCOL A:A 25
SETCOL B:C 15
SETROW .bold:1 1 30
MERGE .bold:1 .font_size:16 .align:center .color:white .bg_color:navy A1:C1 Q4 Sales Report
TEXT .bold:1 .bg_color:silver .border:1 A2 Product
TEXT .bold:1 .bg_color:silver .border:1 B2 Units
TEXT .bold:1 .bg_color:silver .border:1 C2 Revenue
TEXT .border:1 A3 Widget Pro
NUM .border:1 B3 1250
NUM .border:1 .num_format:$#,##0 C3 62500
TEXT .bold:1 .border:2 A4 Total
NUM .bold:1 .border:2 B4 1250
NUM .bold:1 .border:2 .num_format:$#,##0 C4 62500
EOF
```

### Formulas and Frozen Panes

```bash
cat <<'EOF' | xlsx-writer budget.xlsx
ADD_WORKSHEET Budget
SETCOL A:A 15
SETCOL B:D 12
FREEZE A2
TEXT .bold:1 A1 Item
TEXT .bold:1 B1 Q1
TEXT .bold:1 C1 Q2
TEXT .bold:1 D1 Total
TEXT A2 Sales
NUM B2 50000
NUM C2 55000
FORMULA .bold:1 D2 =B2+C2
TEXT A3 Costs
NUM B3 30000
NUM C3 32000
FORMULA .bold:1 D3 =B3+C3
TEXT .bold:1 .border:2 A5 Net
FORMULA .bold:1 .border:2 .num_format:$#,##0 B5 =B2-B3
FORMULA .bold:1 .border:2 .num_format:$#,##0 C5 =C2-C3
FORMULA .bold:1 .border:2 .num_format:$#,##0 D5 =D2-D3
EOF
```

### Pipeline from SQL

```bash
psql -t -A -F $'\t' -c "SELECT name, revenue FROM sales" mydb \
  | awk -F'\t' 'BEGIN {
      print "ADD_WORKSHEET SQL Export"
      print "SETCOL A:A 30"
      print "SETCOL B:B 15"
      print "TEXT .bold:1 A1 Name"
      print "TEXT .bold:1 B1 Revenue"
    }
    {
      print "TEXT A" NR+1 " " $1
      print "NUM .num_format:$#,##0.00 B" NR+1 " " $2
    }' \
  | xlsx-writer sales_export.xlsx
```

### Bulk Data with ROW

```bash
{
echo "ADD_WORKSHEET Data"
echo "FREEZE A2"
printf "ROW .bold:1 A1\tID\tName\tAmount\n"
seq 1 100 | while read i; do
  printf "ROW A$((i+1))\t$i\tItem_$i\t$((RANDOM % 10000)).$((RANDOM % 100))\n"
done
} | xlsx-writer bulk.xlsx
```

### Data Validation Form

```bash
cat <<'EOF' | xlsx-writer form.xlsx
ADD_WORKSHEET Entry Form
SETCOL A:A 15
SETCOL B:B 25
TEXT .bold:1 A1 Status
TEXT .bold:1 B1 Comment
DATA_VALIDATION A2:A100 list Open,In Progress,Resolved,Closed
DATA_VALIDATION B2:B100 text_length less_than_or_equal 500
EOF
```

## Design Notes

- **Colon-containing values**: Inline style values like `.num_format:hh:mm:ss`
  are handled correctly (splits on first `:` only).
- **Error handling**: Each command validates its arguments and reports clear
  errors to stderr while continuing to process remaining input.
- **Performance**: Fast for large datasets due to Rust's compiled nature and
  zero-cost abstractions. See [Performance](#performance) below.

## Command Reference (Quick)

| Command | Syntax | Description |
|---------|--------|-------------|
| `ADD_WORKSHEET` | `[title]` | Add a new worksheet |
| `TEXT` | `[.styles] cell text...` | Write string |
| `NUM` | `[.styles] cell number` | Write number |
| `FORMULA` | `[.styles] cell expr` | Write formula |
| `ROW` | `[.styles] cell values...` | Write entire row (tab-delimited) |
| `FAST` | `row col style text...` | Bulk write (0-indexed, named style) |
| `URL` | `[.styles] cell url [title]` | Write hyperlink |
| `DATE` | `[.styles] cell date` | Write date/datetime |
| `BLANK` | `[.styles] cell` | Write formatted empty cell |
| `MERGE` | `[.styles] range text...` | Merge cells with text |
| `COMMENT` | `cell text...` | Add cell note |
| `RICH_TEXT` | `[.styles] cell segs...` | Mixed-format text |
| `SETCOL` | `[.styles] range width` | Set column width |
| `SETROW` | `[.styles] row height` | Set row height |
| `STYLE` | `name key val...` | Define named style |
| `FREEZE` | `cell` | Freeze panes |
| `AUTOFILTER` | `range` | Add dropdown filters |
| `TABLE` | `range [name] [style]` | Create Table object |
| `CONDITIONAL` | `[.styles] range type...` | Conditional formatting |
| `DATA_VALIDATION` | `range type [vals]` | Input validation |
| `TAB_COLOR` | `color` | Set tab color |
| `HIDE_SHEET` | | Hide current sheet |
| `PROTECT_SHEET` | `[password]` | Protect current sheet |
| `DEFINE_NAME` | `name formula` | Define named range |
| `PAGE_SETUP` | `key value...` | Print configuration |
| `SET_PROPERTY` | `key value...` | Document properties |
| `IMAGE` | `row col file [offsets]` | Insert image |

## Performance

Run `bash bench.sh` to benchmark with 100,000 rows. Results on an Apple M1
(median of 5 runs):

| Scenario | Input Lines | Time | Output | Throughput |
|----------|-------------|------|--------|------------|
| FAST, no formatting | 300K | 0.50s | 2.2 MB | 200K rows/s |
| FAST + named styles | 300K | 0.64s | 2.3 MB | 156K rows/s |
| ROW (1 line per row) | 100K | 0.40s | 2.2 MB | 250K rows/s |
| Inline styles every cell | 300K | 0.82s | 2.2 MB | 122K rows/s |
| Mixed (formulas+freeze+filter) | 400K | 0.77s | 2.9 MB | 130K rows/s |

- **ROW is the fastest path** — one stdin line per row means less I/O and parsing
- **FAST** uses 0-based numeric
  row/col indices and pre-defined named styles, avoiding cell-reference parsing
  and inline-style parsing on each line
- Inline styles on every cell add ~60% overhead vs plain FAST, mitigated by
  format caching (identical style combinations are only built once)
- 100K rows with formulas + autofilter completes in under a second

## License

MIT
