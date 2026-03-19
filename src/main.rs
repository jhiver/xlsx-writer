use std::collections::HashMap;
use std::env;
use std::io::{self, BufRead, Write as IoWrite};

use anyhow::{bail, Context, Result};
use rust_xlsxwriter::*;

// ---------------------------------------------------------------------------
// Cell reference parsing
// ---------------------------------------------------------------------------

/// Parse "B2" → (row=1, col=1)  (0-indexed)
fn parse_cell_ref(s: &str) -> Result<(u32, u16)> {
    let mut col: u16 = 0;
    let mut row_str = String::new();
    let mut in_row = false;

    for ch in s.chars() {
        if !in_row && ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as u16 - b'A' as u16 + 1);
        } else {
            in_row = true;
            row_str.push(ch);
        }
    }

    let row: u32 = row_str
        .parse::<u32>()
        .with_context(|| format!("Invalid cell reference: '{s}'"))?
        .checked_sub(1)
        .with_context(|| format!("Row must be >= 1 in cell reference: '{s}'"))?;
    let col = col
        .checked_sub(1)
        .with_context(|| format!("Missing column letter in cell reference: '{s}'"))?;

    Ok((row, col))
}

/// Parse "A1:C3" → (0, 0, 2, 2)
fn parse_range(s: &str) -> Result<(u32, u16, u32, u16)> {
    let (start, end) = s
        .split_once(':')
        .with_context(|| format!("Invalid range (expected ':'): '{s}'"))?;
    let (r1, c1) = parse_cell_ref(start)?;
    let (r2, c2) = parse_cell_ref(end)?;
    Ok((r1, c1, r2, c2))
}

/// Parse a column letter like "A" → 0, "B" → 1, "AA" → 26
fn parse_col_letter(s: &str) -> Result<u16> {
    let mut col: u16 = 0;
    for ch in s.chars() {
        if !ch.is_ascii_alphabetic() {
            bail!("Invalid column letter: '{s}'");
        }
        col = col * 26 + (ch.to_ascii_uppercase() as u16 - b'A' as u16 + 1);
    }
    col.checked_sub(1)
        .with_context(|| format!("Empty column: '{s}'"))
}

/// Parse "A:B" → (0, 1)
fn parse_col_range(s: &str) -> Result<(u16, u16)> {
    let (a, b) = s
        .split_once(':')
        .with_context(|| format!("Invalid column range (expected ':'): '{s}'"))?;
    Ok((parse_col_letter(a)?, parse_col_letter(b)?))
}

// ---------------------------------------------------------------------------
// Color / alignment / border / pattern parsing
// ---------------------------------------------------------------------------

fn parse_color(s: &str) -> Color {
    match s.to_lowercase().as_str() {
        "black" => Color::Black,
        "blue" => Color::Blue,
        "brown" => Color::Brown,
        "cyan" => Color::Cyan,
        "gray" | "grey" => Color::Gray,
        "green" => Color::Green,
        "lime" => Color::Lime,
        "magenta" => Color::Magenta,
        "navy" => Color::Navy,
        "orange" => Color::Orange,
        "pink" => Color::Pink,
        "purple" => Color::Purple,
        "red" => Color::Red,
        "silver" => Color::Silver,
        "white" => Color::White,
        "yellow" => Color::Yellow,
        other => {
            let hex = other.trim_start_matches('#').trim_start_matches("0x");
            u32::from_str_radix(hex, 16)
                .map(Color::RGB)
                .unwrap_or_else(|_| {
                    eprintln!("Warning: unknown color '{s}', defaulting to black");
                    Color::Black
                })
        }
    }
}

fn parse_align(s: &str) -> Option<FormatAlign> {
    Some(match s.to_lowercase().as_str() {
        "left" => FormatAlign::Left,
        "center" | "centre" => FormatAlign::Center,
        "right" => FormatAlign::Right,
        "fill" => FormatAlign::Fill,
        "justify" => FormatAlign::Justify,
        "center_across" | "centre_across" => FormatAlign::CenterAcross,
        _ => {
            eprintln!("Warning: unknown alignment '{s}'");
            return None;
        }
    })
}

fn parse_valign(s: &str) -> Option<FormatAlign> {
    Some(match s.to_lowercase().as_str() {
        "top" => FormatAlign::Top,
        "vcenter" | "vcentre" => FormatAlign::VerticalCenter,
        "bottom" => FormatAlign::Bottom,
        "vjustify" => FormatAlign::VerticalJustify,
        _ => {
            eprintln!("Warning: unknown vertical alignment '{s}'");
            return None;
        }
    })
}

fn parse_border(s: &str) -> FormatBorder {
    match s {
        "0" => FormatBorder::None,
        "1" => FormatBorder::Thin,
        "2" => FormatBorder::Medium,
        "3" => FormatBorder::Dashed,
        "4" => FormatBorder::Dotted,
        "5" => FormatBorder::Thick,
        "6" => FormatBorder::Double,
        "7" => FormatBorder::Hair,
        "8" => FormatBorder::MediumDashed,
        "9" => FormatBorder::DashDot,
        "10" => FormatBorder::MediumDashDot,
        "11" => FormatBorder::DashDotDot,
        "12" => FormatBorder::MediumDashDotDot,
        "13" => FormatBorder::SlantDashDot,
        _ => {
            eprintln!("Warning: unknown border type '{s}', using Thin");
            FormatBorder::Thin
        }
    }
}

fn parse_pattern(n: u8) -> FormatPattern {
    match n {
        0 => FormatPattern::None,
        1 => FormatPattern::Solid,
        2 => FormatPattern::MediumGray,
        3 => FormatPattern::DarkGray,
        4 => FormatPattern::LightGray,
        5 => FormatPattern::DarkHorizontal,
        6 => FormatPattern::DarkVertical,
        7 => FormatPattern::DarkDown,
        8 => FormatPattern::DarkUp,
        9 => FormatPattern::DarkGrid,
        10 => FormatPattern::DarkTrellis,
        11 => FormatPattern::LightHorizontal,
        12 => FormatPattern::LightVertical,
        13 => FormatPattern::LightDown,
        14 => FormatPattern::LightUp,
        15 => FormatPattern::LightGrid,
        16 => FormatPattern::LightTrellis,
        17 => FormatPattern::Gray125,
        18 => FormatPattern::Gray0625,
        _ => FormatPattern::None,
    }
}

// ---------------------------------------------------------------------------
// Table style parsing
// ---------------------------------------------------------------------------

fn parse_table_style(s: &str) -> Result<TableStyle> {
    let (prefix, num_str) = if s.eq_ignore_ascii_case("none") {
        return Ok(TableStyle::None);
    } else if let Some(n) = s.strip_prefix("Light") {
        ("light", n)
    } else if let Some(n) = s.strip_prefix("Medium") {
        ("medium", n)
    } else if let Some(n) = s.strip_prefix("Dark") {
        ("dark", n)
    } else {
        bail!("Unknown table style: '{s}' (expected None/Light1-21/Medium1-28/Dark1-11)");
    };

    let num: u8 = num_str
        .parse()
        .with_context(|| format!("Invalid table style number: '{s}'"))?;

    Ok(match (prefix, num) {
        ("light", 1) => TableStyle::Light1,
        ("light", 2) => TableStyle::Light2,
        ("light", 3) => TableStyle::Light3,
        ("light", 4) => TableStyle::Light4,
        ("light", 5) => TableStyle::Light5,
        ("light", 6) => TableStyle::Light6,
        ("light", 7) => TableStyle::Light7,
        ("light", 8) => TableStyle::Light8,
        ("light", 9) => TableStyle::Light9,
        ("light", 10) => TableStyle::Light10,
        ("light", 11) => TableStyle::Light11,
        ("light", 12) => TableStyle::Light12,
        ("light", 13) => TableStyle::Light13,
        ("light", 14) => TableStyle::Light14,
        ("light", 15) => TableStyle::Light15,
        ("light", 16) => TableStyle::Light16,
        ("light", 17) => TableStyle::Light17,
        ("light", 18) => TableStyle::Light18,
        ("light", 19) => TableStyle::Light19,
        ("light", 20) => TableStyle::Light20,
        ("light", 21) => TableStyle::Light21,
        ("medium", 1) => TableStyle::Medium1,
        ("medium", 2) => TableStyle::Medium2,
        ("medium", 3) => TableStyle::Medium3,
        ("medium", 4) => TableStyle::Medium4,
        ("medium", 5) => TableStyle::Medium5,
        ("medium", 6) => TableStyle::Medium6,
        ("medium", 7) => TableStyle::Medium7,
        ("medium", 8) => TableStyle::Medium8,
        ("medium", 9) => TableStyle::Medium9,
        ("medium", 10) => TableStyle::Medium10,
        ("medium", 11) => TableStyle::Medium11,
        ("medium", 12) => TableStyle::Medium12,
        ("medium", 13) => TableStyle::Medium13,
        ("medium", 14) => TableStyle::Medium14,
        ("medium", 15) => TableStyle::Medium15,
        ("medium", 16) => TableStyle::Medium16,
        ("medium", 17) => TableStyle::Medium17,
        ("medium", 18) => TableStyle::Medium18,
        ("medium", 19) => TableStyle::Medium19,
        ("medium", 20) => TableStyle::Medium20,
        ("medium", 21) => TableStyle::Medium21,
        ("medium", 22) => TableStyle::Medium22,
        ("medium", 23) => TableStyle::Medium23,
        ("medium", 24) => TableStyle::Medium24,
        ("medium", 25) => TableStyle::Medium25,
        ("medium", 26) => TableStyle::Medium26,
        ("medium", 27) => TableStyle::Medium27,
        ("medium", 28) => TableStyle::Medium28,
        ("dark", 1) => TableStyle::Dark1,
        ("dark", 2) => TableStyle::Dark2,
        ("dark", 3) => TableStyle::Dark3,
        ("dark", 4) => TableStyle::Dark4,
        ("dark", 5) => TableStyle::Dark5,
        ("dark", 6) => TableStyle::Dark6,
        ("dark", 7) => TableStyle::Dark7,
        ("dark", 8) => TableStyle::Dark8,
        ("dark", 9) => TableStyle::Dark9,
        ("dark", 10) => TableStyle::Dark10,
        ("dark", 11) => TableStyle::Dark11,
        _ => bail!("Table style out of range: '{s}'"),
    })
}

// ---------------------------------------------------------------------------
// Boolean parsing helper
// ---------------------------------------------------------------------------

fn parse_bool(s: &str) -> bool {
    matches!(s, "1" | "true" | "yes" | "on")
}

// ---------------------------------------------------------------------------
// Format building
// ---------------------------------------------------------------------------

fn apply_property(format: Format, key: &str, value: &str) -> Format {
    match key {
        "bold" => format.set_bold(),
        "italic" => format.set_italic(),
        "underline" => match value {
            "2" => format.set_underline(FormatUnderline::Double),
            "33" => format.set_underline(FormatUnderline::SingleAccounting),
            "34" => format.set_underline(FormatUnderline::DoubleAccounting),
            _ => format.set_underline(FormatUnderline::Single),
        },
        "strikeout" | "font_strikeout" => format.set_font_strikethrough(),
        "font_size" | "size" => {
            if let Ok(v) = value.parse::<f64>() {
                format.set_font_size(v)
            } else {
                format
            }
        }
        "font" | "font_name" => format.set_font_name(value),
        "color" | "font_color" => format.set_font_color(parse_color(value)),
        "bg_color" => format.set_background_color(parse_color(value)),
        "fg_color" => format.set_foreground_color(parse_color(value)),
        "num_format" => format.set_num_format(value),
        "align" => {
            if let Some(a) = parse_align(value) {
                format.set_align(a)
            } else {
                format
            }
        }
        "valign" => {
            if let Some(a) = parse_valign(value) {
                format.set_align(a)
            } else {
                format
            }
        }
        "text_wrap" => format.set_text_wrap(),
        "border" => format.set_border(parse_border(value)),
        "border_color" => format.set_border_color(parse_color(value)),
        "bottom" => format.set_border_bottom(parse_border(value)),
        "top" => format.set_border_top(parse_border(value)),
        "left" => format.set_border_left(parse_border(value)),
        "right" => format.set_border_right(parse_border(value)),
        "bottom_color" => format.set_border_bottom_color(parse_color(value)),
        "top_color" => format.set_border_top_color(parse_color(value)),
        "left_color" => format.set_border_left_color(parse_color(value)),
        "right_color" => format.set_border_right_color(parse_color(value)),
        "indent" => {
            if let Ok(v) = value.parse::<u8>() {
                format.set_indent(v)
            } else {
                format
            }
        }
        "rotation" => {
            if let Ok(v) = value.parse::<i16>() {
                format.set_rotation(v)
            } else {
                format
            }
        }
        "shrink" => format.set_shrink(),
        "pattern" => {
            if let Ok(n) = value.parse::<u8>() {
                format.set_pattern(parse_pattern(n))
            } else {
                format
            }
        }
        _ => {
            eprintln!("Warning: unknown format property '{key}'");
            format
        }
    }
}

// ---------------------------------------------------------------------------
// Inline style extraction & caching
// ---------------------------------------------------------------------------

/// Returns true for tokens like `.bold:1`, `.font_size:12`, `.bg_color:#FF0000`
fn is_style_token(s: &str) -> bool {
    if !s.starts_with('.') || s.len() < 3 {
        return false;
    }
    let rest = &s[1..];
    match rest.find(':') {
        Some(pos) if pos > 0 => rest[..pos]
            .chars()
            .all(|c| c.is_ascii_lowercase() || c == '_' || c == '-'),
        _ => false,
    }
}

/// Split arguments into (style properties, remaining args).
fn extract_inline_styles<'a>(args: &[&'a str]) -> (Vec<(&'a str, &'a str)>, Vec<&'a str>) {
    let mut styles = Vec::new();
    let mut rest = Vec::new();

    for &arg in args {
        if is_style_token(arg) {
            // split_once on first ':' to preserve colons in values (e.g. hh:mm:ss)
            if let Some((key, value)) = arg[1..].split_once(':') {
                styles.push((key, value));
            }
        } else {
            rest.push(arg);
        }
    }

    (styles, rest)
}

fn styles_to_key(styles: &[(&str, &str)]) -> String {
    styles.iter().map(|(k, v)| format!(".{k}:{v}")).collect()
}

fn build_format(styles: &[(&str, &str)]) -> Format {
    let mut format = Format::new();
    for &(key, value) in styles {
        format = apply_property(format, key, value);
    }
    format
}

fn get_or_create_format(cache: &mut HashMap<String, Format>, styles: &[(&str, &str)]) -> Format {
    let key = styles_to_key(styles);
    if !cache.contains_key(&key) {
        cache.insert(key.clone(), build_format(styles));
    }
    cache[&key].clone()
}

// ---------------------------------------------------------------------------
// XlsxWriter – the main engine
// ---------------------------------------------------------------------------

struct XlsxWriter {
    workbook: Workbook,
    current_ws_index: Option<usize>,
    ws_count: usize,
    named_formats: HashMap<String, Format>,
    inline_cache: HashMap<String, Format>,
    merged_cache: HashMap<String, Format>,
    doc_properties: Vec<(String, String)>,
}

impl XlsxWriter {
    fn new() -> Self {
        Self {
            workbook: Workbook::new(),
            current_ws_index: None,
            ws_count: 0,
            named_formats: HashMap::new(),
            inline_cache: HashMap::new(),
            merged_cache: HashMap::new(),
            doc_properties: Vec::new(),
        }
    }

    fn execute(&mut self, command: &str, args: &[&str]) -> Result<()> {
        match command {
            // Original commands
            "STYLE" => self.cmd_style(args),
            "SET_PROPERTY" => self.cmd_set_property(args),
            "ADD_WORKSHEET" => self.cmd_add_worksheet(args),
            "IMAGE" => self.cmd_image(args),
            "URL" => self.cmd_url(args),
            "TEXT" => self.cmd_text(args),
            "FAST" => self.cmd_fast(args),
            "DATE" => self.cmd_date(args),
            "MERGE" => self.cmd_merge(args),
            "SETCOL" => self.cmd_setcol(args),
            "SETROW" => self.cmd_setrow(args),
            "NUM" => self.cmd_num(args),
            // New commands
            "FORMULA" => self.cmd_formula(args),
            "FREEZE" => self.cmd_freeze(args),
            "AUTOFILTER" => self.cmd_autofilter(args),
            "BLANK" => self.cmd_blank(args),
            "TABLE" => self.cmd_table(args),
            "CONDITIONAL" => self.cmd_conditional(args),
            "COMMENT" => self.cmd_comment(args),
            "TAB_COLOR" => self.cmd_tab_color(args),
            "DATA_VALIDATION" => self.cmd_data_validation(args),
            "PAGE_SETUP" => self.cmd_page_setup(args),
            "HIDE_SHEET" => self.cmd_hide_sheet(args),
            "PROTECT_SHEET" => self.cmd_protect_sheet(args),
            "DEFINE_NAME" => self.cmd_define_name(args),
            _ => Ok(()), // unknown commands silently ignored
        }
    }

    /// Handle commands that need raw line access (tab preservation, special delimiters)
    fn execute_raw(&mut self, command: &str, line: &str) -> Result<()> {
        match command {
            "ROW" => self.cmd_row_raw(line),
            "RICH_TEXT" => self.cmd_rich_text_raw(line),
            _ => Ok(()),
        }
    }

    // -----------------------------------------------------------------------
    // Original commands
    // -----------------------------------------------------------------------

    // -- STYLE name key value key value ... --------------------------------
    fn cmd_style(&mut self, args: &[&str]) -> Result<()> {
        if args.is_empty() {
            bail!("STYLE requires a name");
        }
        let name = args[0];
        let mut format = Format::new();
        let mut i = 1;
        while i + 1 < args.len() {
            format = apply_property(format, args[i], args[i + 1]);
            i += 2;
        }
        self.named_formats.insert(name.to_string(), format);
        Ok(())
    }

    // -- SET_PROPERTY key value ... ----------------------------------------
    fn cmd_set_property(&mut self, args: &[&str]) -> Result<()> {
        if args.is_empty() {
            bail!("SET_PROPERTY requires a key");
        }
        let key = args[0];
        let value = args[1..].join(" ");
        self.doc_properties.push((key.to_string(), value));
        Ok(())
    }

    // -- ADD_WORKSHEET [title ...] -----------------------------------------
    fn cmd_add_worksheet(&mut self, args: &[&str]) -> Result<()> {
        let title = args.join(" ");
        let ws = self.workbook.add_worksheet();
        if !title.is_empty() {
            ws.set_name(&title)?;
        }
        self.current_ws_index = Some(self.ws_count);
        self.ws_count += 1;
        Ok(())
    }

    // -- IMAGE row col filename [x_off y_off [x_scale y_scale]] ------------
    fn cmd_image(&mut self, args: &[&str]) -> Result<()> {
        if args.len() < 3 {
            bail!("IMAGE requires: row col filename");
        }
        let row: u32 = args[0].parse().context("IMAGE: invalid row")?;
        let col: u16 = args[1].parse().context("IMAGE: invalid col")?;
        let filename = args[2];

        let mut image = Image::new(filename)?;

        if args.len() > 5 {
            if let Ok(xs) = args[5].parse::<f64>() {
                image = image.set_scale_width(xs);
            }
        }
        if args.len() > 6 {
            if let Ok(ys) = args[6].parse::<f64>() {
                image = image.set_scale_height(ys);
            }
        }

        let idx = self.current_ws_index.context("No worksheet")?;
        let ws = self.workbook.worksheet_from_index(idx)?;

        if args.len() > 4 {
            let xo: u32 = args.get(3).and_then(|s| s.parse().ok()).unwrap_or(0);
            let yo: u32 = args.get(4).and_then(|s| s.parse().ok()).unwrap_or(0);
            ws.insert_image_with_offset(row, col, &image, xo, yo)?;
        } else {
            ws.insert_image(row, col, &image)?;
        }
        Ok(())
    }

    // -- URL [.styles] cell url [title ...] --------------------------------
    fn cmd_url(&mut self, args: &[&str]) -> Result<()> {
        let (styles, rest) = extract_inline_styles(args);
        if rest.len() < 2 {
            bail!("URL requires: cell url [title...]");
        }
        let (row, col) = parse_cell_ref(rest[0])?;
        let url_str = rest[1];
        let title = rest[2..].join(" ");

        let url = if title.is_empty() {
            Url::new(url_str)
        } else {
            Url::new(url_str).set_text(&title)
        };

        let idx = self.current_ws_index.context("No worksheet")?;
        if styles.is_empty() {
            self.workbook
                .worksheet_from_index(idx)?
                .write_url(row, col, url)?;
        } else {
            let format = get_or_create_format(&mut self.inline_cache, &styles);
            self.workbook
                .worksheet_from_index(idx)?
                .write_url_with_format(row, col, url, &format)?;
        }
        Ok(())
    }

    // -- TEXT [.styles] cell text ... ---------------------------------------
    fn cmd_text(&mut self, args: &[&str]) -> Result<()> {
        let (styles, rest) = extract_inline_styles(args);
        if rest.is_empty() {
            bail!("TEXT requires: cell [text...]");
        }
        let (row, col) = parse_cell_ref(rest[0])?;
        let text = rest[1..].join(" ");

        let idx = self.current_ws_index.context("No worksheet")?;
        if styles.is_empty() {
            self.workbook
                .worksheet_from_index(idx)?
                .write_string(row, col, &text)?;
        } else {
            let format = get_or_create_format(&mut self.inline_cache, &styles);
            self.workbook
                .worksheet_from_index(idx)?
                .write_string_with_format(row, col, &text, &format)?;
        }
        Ok(())
    }

    // -- NUM [.styles] cell number -----------------------------------------
    fn cmd_num(&mut self, args: &[&str]) -> Result<()> {
        let (styles, rest) = extract_inline_styles(args);
        if rest.len() < 2 {
            bail!("NUM requires: cell number");
        }
        let (row, col) = parse_cell_ref(rest[0])?;
        let num: f64 = rest[1..].join(" ").parse().context("NUM: invalid number")?;

        let idx = self.current_ws_index.context("No worksheet")?;
        if styles.is_empty() {
            self.workbook
                .worksheet_from_index(idx)?
                .write_number(row, col, num)?;
        } else {
            let format = get_or_create_format(&mut self.inline_cache, &styles);
            self.workbook
                .worksheet_from_index(idx)?
                .write_number_with_format(row, col, num, &format)?;
        }
        Ok(())
    }

    // -- FAST row col style_name text ... ----------------------------------
    // Uses numeric (0-based) row/col and a named style.
    // Writes as number if text starts with a digit, else as string.
    fn cmd_fast(&mut self, args: &[&str]) -> Result<()> {
        if args.len() < 3 {
            bail!("FAST requires: row col style_name [text...]");
        }
        let row: u32 = args[0].parse().context("FAST: invalid row")?;
        let col: u16 = args[1].parse().context("FAST: invalid col")?;
        let style_name = args[2];
        let text = args[3..].join(" ");

        let format = self.named_formats.get(style_name).cloned();
        let idx = self.current_ws_index.context("No worksheet")?;
        let ws = self.workbook.worksheet_from_index(idx)?;

        // Starts with digit → try number, else → string
        let starts_with_digit = text.as_bytes().first().is_some_and(|b| b.is_ascii_digit());

        match (starts_with_digit, &format) {
            (true, Some(fmt)) => {
                if let Ok(n) = text.parse::<f64>() {
                    ws.write_number_with_format(row, col, n, fmt)?;
                } else {
                    ws.write_string_with_format(row, col, &text, fmt)?;
                }
            }
            (true, None) => {
                if let Ok(n) = text.parse::<f64>() {
                    ws.write_number(row, col, n)?;
                } else {
                    ws.write_string(row, col, &text)?;
                }
            }
            (false, Some(fmt)) => {
                ws.write_string_with_format(row, col, &text, fmt)?;
            }
            (false, None) => {
                ws.write_string(row, col, &text)?;
            }
        }
        Ok(())
    }

    // -- DATE [.styles] cell date ------------------------------------------
    fn cmd_date(&mut self, args: &[&str]) -> Result<()> {
        let (styles, rest) = extract_inline_styles(args);
        if rest.len() < 2 {
            bail!("DATE requires: cell date");
        }
        let (row, col) = parse_cell_ref(rest[0])?;
        let date_str = rest[1];

        let datetime = ExcelDateTime::parse_from_str(date_str)?;

        let idx = self.current_ws_index.context("No worksheet")?;
        if styles.is_empty() {
            self.workbook
                .worksheet_from_index(idx)?
                .write(row, col, &datetime)?;
        } else {
            let format = get_or_create_format(&mut self.inline_cache, &styles);
            self.workbook
                .worksheet_from_index(idx)?
                .write_with_format(row, col, &datetime, &format)?;
        }
        Ok(())
    }

    // -- MERGE [.styles] range text ... ------------------------------------
    fn cmd_merge(&mut self, args: &[&str]) -> Result<()> {
        let (styles, rest) = extract_inline_styles(args);
        if rest.is_empty() {
            bail!("MERGE requires: range [text...]");
        }
        let (r1, c1, r2, c2) = parse_range(rest[0])?;
        let text = rest[1..].join(" ");

        let format = if styles.is_empty() {
            Format::new()
        } else {
            get_or_create_format(&mut self.merged_cache, &styles)
        };

        let idx = self.current_ws_index.context("No worksheet")?;
        self.workbook
            .worksheet_from_index(idx)?
            .merge_range(r1, c1, r2, c2, &text, &format)?;
        Ok(())
    }

    // -- SETCOL [.styles] col_range width ----------------------------------
    fn cmd_setcol(&mut self, args: &[&str]) -> Result<()> {
        let (styles, rest) = extract_inline_styles(args);
        if rest.len() < 2 {
            bail!("SETCOL requires: col_range width");
        }
        let (first_col, last_col) = parse_col_range(rest[0])?;
        let width: f64 = rest[1].parse().context("SETCOL: invalid width")?;

        let format = if !styles.is_empty() {
            Some(get_or_create_format(&mut self.inline_cache, &styles))
        } else {
            None
        };

        let idx = self.current_ws_index.context("No worksheet")?;
        let ws = self.workbook.worksheet_from_index(idx)?;

        for col in first_col..=last_col {
            ws.set_column_width(col, width)?;
            if let Some(ref fmt) = format {
                ws.set_column_format(col, fmt)?;
            }
        }
        Ok(())
    }

    // -- SETROW [.styles] row height (row is 1-indexed) --------------------
    fn cmd_setrow(&mut self, args: &[&str]) -> Result<()> {
        let (styles, rest) = extract_inline_styles(args);
        if rest.len() < 2 {
            bail!("SETROW requires: row height");
        }
        let row: u32 = rest[0]
            .parse::<u32>()
            .context("SETROW: invalid row")?
            .checked_sub(1)
            .context("SETROW: row must be >= 1")?;
        let height: f64 = rest[1].parse().context("SETROW: invalid height")?;

        let format = if !styles.is_empty() {
            Some(get_or_create_format(&mut self.inline_cache, &styles))
        } else {
            None
        };

        let idx = self.current_ws_index.context("No worksheet")?;
        let ws = self.workbook.worksheet_from_index(idx)?;

        ws.set_row_height(row, height)?;
        if let Some(ref fmt) = format {
            ws.set_row_format(row, fmt)?;
        }
        Ok(())
    }

    // -----------------------------------------------------------------------
    // New commands
    // -----------------------------------------------------------------------

    // -- FORMULA [.styles] cell expression ---------------------------------
    fn cmd_formula(&mut self, args: &[&str]) -> Result<()> {
        let (styles, rest) = extract_inline_styles(args);
        if rest.len() < 2 {
            bail!("FORMULA requires: cell expression");
        }
        let (row, col) = parse_cell_ref(rest[0])?;
        let formula = rest[1..].join(" ");

        let idx = self.current_ws_index.context("No worksheet")?;
        if styles.is_empty() {
            self.workbook
                .worksheet_from_index(idx)?
                .write_formula(row, col, formula.as_str())?;
        } else {
            let format = get_or_create_format(&mut self.inline_cache, &styles);
            self.workbook
                .worksheet_from_index(idx)?
                .write_formula_with_format(row, col, formula.as_str(), &format)?;
        }
        Ok(())
    }

    // -- FREEZE cell -------------------------------------------------------
    // Freeze panes at the given cell. FREEZE A2 freezes the top row.
    fn cmd_freeze(&mut self, args: &[&str]) -> Result<()> {
        if args.is_empty() {
            bail!("FREEZE requires: cell (e.g. A2 to freeze top row)");
        }
        let (row, col) = parse_cell_ref(args[0])?;
        let idx = self.current_ws_index.context("No worksheet")?;
        self.workbook
            .worksheet_from_index(idx)?
            .set_freeze_panes(row, col)?;
        Ok(())
    }

    // -- AUTOFILTER range --------------------------------------------------
    fn cmd_autofilter(&mut self, args: &[&str]) -> Result<()> {
        if args.is_empty() {
            bail!("AUTOFILTER requires: range (e.g. A1:D100)");
        }
        let (r1, c1, r2, c2) = parse_range(args[0])?;
        let idx = self.current_ws_index.context("No worksheet")?;
        self.workbook
            .worksheet_from_index(idx)?
            .autofilter(r1, c1, r2, c2)?;
        Ok(())
    }

    // -- ROW [.styles] start_cell tab-delimited-values ---------------------
    // Handled via execute_raw for tab preservation.
    fn cmd_row_raw(&mut self, line: &str) -> Result<()> {
        let rest = line
            .strip_prefix("ROW")
            .context("expected ROW")?
            .trim_start();

        // Scan for inline style tokens and cell ref
        let mut styles: Vec<(&str, &str)> = Vec::new();
        let mut remaining = rest;
        let mut cell_ref: Option<&str> = None;

        loop {
            remaining = remaining.trim_start();
            if remaining.is_empty() {
                break;
            }
            let token_end = remaining
                .find(|c: char| c.is_whitespace())
                .unwrap_or(remaining.len());
            let token = &remaining[..token_end];

            if is_style_token(token) {
                if let Some((key, value)) = token[1..].split_once(':') {
                    styles.push((key, value));
                }
                remaining = &remaining[token_end..];
            } else if cell_ref.is_none() {
                cell_ref = Some(token);
                remaining = &remaining[token_end..];
                // Skip exactly one whitespace separator to get to data portion
                if !remaining.is_empty() {
                    remaining = &remaining[1..];
                }
                break;
            }
        }

        let cell_ref = cell_ref.context("ROW requires: cell [data...]")?;
        let (row, start_col) = parse_cell_ref(cell_ref)?;

        // Split data on tabs if tabs are present, otherwise fall back to whitespace
        let values: Vec<&str> = if remaining.contains('\t') {
            remaining.split('\t').collect()
        } else {
            remaining.split_whitespace().collect()
        };

        let format = if !styles.is_empty() {
            Some(get_or_create_format(&mut self.inline_cache, &styles))
        } else {
            None
        };

        let idx = self.current_ws_index.context("No worksheet")?;
        let ws = self.workbook.worksheet_from_index(idx)?;

        for (i, val) in values.iter().enumerate() {
            let col = start_col + i as u16;
            let val = val.trim();
            if val.is_empty() {
                continue;
            }

            if val.starts_with('=') {
                // Formula
                match &format {
                    Some(fmt) => {
                        ws.write_formula_with_format(row, col, val, fmt)?;
                    }
                    None => {
                        ws.write_formula(row, col, val)?;
                    }
                };
            } else if let Ok(n) = val.parse::<f64>() {
                // Number
                match &format {
                    Some(fmt) => {
                        ws.write_number_with_format(row, col, n, fmt)?;
                    }
                    None => {
                        ws.write_number(row, col, n)?;
                    }
                };
            } else {
                // String
                match &format {
                    Some(fmt) => {
                        ws.write_string_with_format(row, col, val, fmt)?;
                    }
                    None => {
                        ws.write_string(row, col, val)?;
                    }
                };
            }
        }
        Ok(())
    }

    // -- BLANK [.styles] cell ----------------------------------------------
    fn cmd_blank(&mut self, args: &[&str]) -> Result<()> {
        let (styles, rest) = extract_inline_styles(args);
        if rest.is_empty() {
            bail!("BLANK requires: cell");
        }
        let (row, col) = parse_cell_ref(rest[0])?;

        let format = if styles.is_empty() {
            Format::new()
        } else {
            get_or_create_format(&mut self.inline_cache, &styles)
        };

        let idx = self.current_ws_index.context("No worksheet")?;
        self.workbook
            .worksheet_from_index(idx)?
            .write_blank(row, col, &format)?;
        Ok(())
    }

    // -- TABLE range [name] [style] ----------------------------------------
    fn cmd_table(&mut self, args: &[&str]) -> Result<()> {
        if args.is_empty() {
            bail!("TABLE requires: range [name] [style]");
        }
        let (r1, c1, r2, c2) = parse_range(args[0])?;

        let mut table = Table::new();

        if args.len() > 1 && args[1] != "_" {
            table = table.set_name(args[1]);
        }

        if args.len() > 2 {
            table = table.set_style(parse_table_style(args[2])?);
        }

        let idx = self.current_ws_index.context("No worksheet")?;
        self.workbook
            .worksheet_from_index(idx)?
            .add_table(r1, c1, r2, c2, &table)?;
        Ok(())
    }

    // -- CONDITIONAL [.styles] range type [criteria] [values...] -----------
    fn cmd_conditional(&mut self, args: &[&str]) -> Result<()> {
        let (styles, rest) = extract_inline_styles(args);
        if rest.len() < 2 {
            bail!("CONDITIONAL requires: range type [criteria] [values...]");
        }
        let (r1, c1, r2, c2) = parse_range(rest[0])?;
        let cf_type = rest[1];
        let cf_args = &rest[2..];

        let idx = self.current_ws_index.context("No worksheet")?;

        match cf_type {
            "cell" => {
                if cf_args.is_empty() {
                    bail!("CONDITIONAL cell requires: criteria value [value2]");
                }
                let format = build_format(&styles);
                let criteria = cf_args[0];
                let val1: f64 = cf_args
                    .get(1)
                    .context("CONDITIONAL cell requires a value")?
                    .parse()
                    .context("CONDITIONAL cell: invalid number")?;

                let rule = match criteria {
                    "equal_to" => ConditionalFormatCellRule::EqualTo(val1),
                    "not_equal_to" => ConditionalFormatCellRule::NotEqualTo(val1),
                    "greater_than" => ConditionalFormatCellRule::GreaterThan(val1),
                    "greater_than_or_equal_to" | "greater_than_or_equal" => {
                        ConditionalFormatCellRule::GreaterThanOrEqualTo(val1)
                    }
                    "less_than" => ConditionalFormatCellRule::LessThan(val1),
                    "less_than_or_equal_to" | "less_than_or_equal" => {
                        ConditionalFormatCellRule::LessThanOrEqualTo(val1)
                    }
                    "between" => {
                        let val2: f64 = cf_args
                            .get(2)
                            .context("CONDITIONAL cell between requires two values")?
                            .parse()
                            .context("CONDITIONAL cell: invalid second number")?;
                        ConditionalFormatCellRule::Between(val1, val2)
                    }
                    "not_between" => {
                        let val2: f64 = cf_args
                            .get(2)
                            .context("CONDITIONAL cell not_between requires two values")?
                            .parse()
                            .context("CONDITIONAL cell: invalid second number")?;
                        ConditionalFormatCellRule::NotBetween(val1, val2)
                    }
                    _ => bail!("Unknown CONDITIONAL cell criteria: '{criteria}'"),
                };

                let cf = ConditionalFormatCell::new()
                    .set_rule(rule)
                    .set_format(format);
                self.workbook
                    .worksheet_from_index(idx)?
                    .add_conditional_format(r1, c1, r2, c2, &cf)?;
            }

            "duplicate" => {
                let format = build_format(&styles);
                let cf = ConditionalFormatDuplicate::new().set_format(format);
                self.workbook
                    .worksheet_from_index(idx)?
                    .add_conditional_format(r1, c1, r2, c2, &cf)?;
            }

            "unique" => {
                let format = build_format(&styles);
                let cf = ConditionalFormatDuplicate::new()
                    .invert()
                    .set_format(format);
                self.workbook
                    .worksheet_from_index(idx)?
                    .add_conditional_format(r1, c1, r2, c2, &cf)?;
            }

            "blank" => {
                let format = build_format(&styles);
                let cf = ConditionalFormatBlank::new().set_format(format);
                self.workbook
                    .worksheet_from_index(idx)?
                    .add_conditional_format(r1, c1, r2, c2, &cf)?;
            }

            "not_blank" => {
                let format = build_format(&styles);
                let cf = ConditionalFormatBlank::new().invert().set_format(format);
                self.workbook
                    .worksheet_from_index(idx)?
                    .add_conditional_format(r1, c1, r2, c2, &cf)?;
            }

            "formula" => {
                if cf_args.is_empty() {
                    bail!("CONDITIONAL formula requires: formula_expression");
                }
                let format = build_format(&styles);
                let formula_str = cf_args.join(" ");
                let cf = ConditionalFormatFormula::new()
                    .set_rule(formula_str.as_str())
                    .set_format(format);
                self.workbook
                    .worksheet_from_index(idx)?
                    .add_conditional_format(r1, c1, r2, c2, &cf)?;
            }

            "top" => {
                let n: u16 = cf_args
                    .first()
                    .unwrap_or(&"10")
                    .parse()
                    .context("CONDITIONAL top: invalid count")?;
                let format = build_format(&styles);
                let cf = ConditionalFormatTop::new()
                    .set_rule(ConditionalFormatTopRule::Top(n))
                    .set_format(format);
                self.workbook
                    .worksheet_from_index(idx)?
                    .add_conditional_format(r1, c1, r2, c2, &cf)?;
            }

            "bottom" => {
                let n: u16 = cf_args
                    .first()
                    .unwrap_or(&"10")
                    .parse()
                    .context("CONDITIONAL bottom: invalid count")?;
                let format = build_format(&styles);
                let cf = ConditionalFormatTop::new()
                    .set_rule(ConditionalFormatTopRule::Bottom(n))
                    .set_format(format);
                self.workbook
                    .worksheet_from_index(idx)?
                    .add_conditional_format(r1, c1, r2, c2, &cf)?;
            }

            "top_percent" => {
                let n: u16 = cf_args
                    .first()
                    .unwrap_or(&"10")
                    .parse()
                    .context("CONDITIONAL top_percent: invalid count")?;
                let format = build_format(&styles);
                let cf = ConditionalFormatTop::new()
                    .set_rule(ConditionalFormatTopRule::TopPercent(n))
                    .set_format(format);
                self.workbook
                    .worksheet_from_index(idx)?
                    .add_conditional_format(r1, c1, r2, c2, &cf)?;
            }

            "bottom_percent" => {
                let n: u16 = cf_args
                    .first()
                    .unwrap_or(&"10")
                    .parse()
                    .context("CONDITIONAL bottom_percent: invalid count")?;
                let format = build_format(&styles);
                let cf = ConditionalFormatTop::new()
                    .set_rule(ConditionalFormatTopRule::BottomPercent(n))
                    .set_format(format);
                self.workbook
                    .worksheet_from_index(idx)?
                    .add_conditional_format(r1, c1, r2, c2, &cf)?;
            }

            "2_color_scale" => {
                let mut cf = ConditionalFormat2ColorScale::new();
                if let Some(min_color) = cf_args.first() {
                    cf = cf.set_minimum_color(parse_color(min_color));
                }
                if let Some(max_color) = cf_args.get(1) {
                    cf = cf.set_maximum_color(parse_color(max_color));
                }
                self.workbook
                    .worksheet_from_index(idx)?
                    .add_conditional_format(r1, c1, r2, c2, &cf)?;
            }

            "3_color_scale" => {
                let mut cf = ConditionalFormat3ColorScale::new();
                if let Some(min_color) = cf_args.first() {
                    cf = cf.set_minimum_color(parse_color(min_color));
                }
                if let Some(mid_color) = cf_args.get(1) {
                    cf = cf.set_midpoint_color(parse_color(mid_color));
                }
                if let Some(max_color) = cf_args.get(2) {
                    cf = cf.set_maximum_color(parse_color(max_color));
                }
                self.workbook
                    .worksheet_from_index(idx)?
                    .add_conditional_format(r1, c1, r2, c2, &cf)?;
            }

            "data_bar" => {
                let mut cf = ConditionalFormatDataBar::new();
                if let Some(color) = cf_args.first() {
                    cf = cf.set_fill_color(parse_color(color));
                }
                self.workbook
                    .worksheet_from_index(idx)?
                    .add_conditional_format(r1, c1, r2, c2, &cf)?;
            }

            _ => bail!("Unknown CONDITIONAL type: '{cf_type}'"),
        }
        Ok(())
    }

    // -- COMMENT cell text... ----------------------------------------------
    fn cmd_comment(&mut self, args: &[&str]) -> Result<()> {
        if args.is_empty() {
            bail!("COMMENT requires: cell text...");
        }
        let (row, col) = parse_cell_ref(args[0])?;
        let text = args[1..].join(" ");
        let note = Note::new(&text);

        let idx = self.current_ws_index.context("No worksheet")?;
        self.workbook
            .worksheet_from_index(idx)?
            .insert_note(row, col, &note)?;
        Ok(())
    }

    // -- TAB_COLOR color ---------------------------------------------------
    fn cmd_tab_color(&mut self, args: &[&str]) -> Result<()> {
        if args.is_empty() {
            bail!("TAB_COLOR requires: color");
        }
        let color = parse_color(args[0]);
        let idx = self.current_ws_index.context("No worksheet")?;
        self.workbook
            .worksheet_from_index(idx)?
            .set_tab_color(color);
        Ok(())
    }

    // -- DATA_VALIDATION range type [values/criteria...] -------------------
    fn cmd_data_validation(&mut self, args: &[&str]) -> Result<()> {
        if args.len() < 2 {
            bail!("DATA_VALIDATION requires: range type [values...]");
        }
        let (r1, c1, r2, c2) = parse_range(args[0])?;
        let dv_type = args[1];
        let dv_args = &args[2..];

        let dv = match dv_type {
            "list" => {
                let list_str = dv_args.join(" ");
                let values: Vec<&str> = list_str.split(',').map(|s| s.trim()).collect();
                DataValidation::new().allow_list_strings(&values)?
            }

            "whole_number" | "integer" => {
                if dv_args.is_empty() {
                    bail!("DATA_VALIDATION whole_number requires: criteria value [value2]");
                }
                let criteria = dv_args[0];
                let val1: i32 = dv_args
                    .get(1)
                    .context("DATA_VALIDATION whole_number requires a value")?
                    .parse()
                    .context("DATA_VALIDATION whole_number: invalid number")?;

                let rule = match criteria {
                    "equal_to" => DataValidationRule::EqualTo(val1),
                    "not_equal_to" => DataValidationRule::NotEqualTo(val1),
                    "greater_than" => DataValidationRule::GreaterThan(val1),
                    "greater_than_or_equal_to" | "greater_than_or_equal" => {
                        DataValidationRule::GreaterThanOrEqualTo(val1)
                    }
                    "less_than" => DataValidationRule::LessThan(val1),
                    "less_than_or_equal_to" | "less_than_or_equal" => {
                        DataValidationRule::LessThanOrEqualTo(val1)
                    }
                    "between" => {
                        let val2: i32 = dv_args
                            .get(2)
                            .context("DATA_VALIDATION between requires two values")?
                            .parse()
                            .context("DATA_VALIDATION: invalid second number")?;
                        DataValidationRule::Between(val1, val2)
                    }
                    "not_between" => {
                        let val2: i32 = dv_args
                            .get(2)
                            .context("DATA_VALIDATION not_between requires two values")?
                            .parse()
                            .context("DATA_VALIDATION: invalid second number")?;
                        DataValidationRule::NotBetween(val1, val2)
                    }
                    _ => bail!("Unknown DATA_VALIDATION criteria: '{criteria}'"),
                };

                DataValidation::new().allow_whole_number(rule)
            }

            "decimal" | "float" => {
                if dv_args.is_empty() {
                    bail!("DATA_VALIDATION decimal requires: criteria value [value2]");
                }
                let criteria = dv_args[0];
                let val1: f64 = dv_args
                    .get(1)
                    .context("DATA_VALIDATION decimal requires a value")?
                    .parse()
                    .context("DATA_VALIDATION decimal: invalid number")?;

                let rule = match criteria {
                    "equal_to" => DataValidationRule::EqualTo(val1),
                    "not_equal_to" => DataValidationRule::NotEqualTo(val1),
                    "greater_than" => DataValidationRule::GreaterThan(val1),
                    "greater_than_or_equal_to" | "greater_than_or_equal" => {
                        DataValidationRule::GreaterThanOrEqualTo(val1)
                    }
                    "less_than" => DataValidationRule::LessThan(val1),
                    "less_than_or_equal_to" | "less_than_or_equal" => {
                        DataValidationRule::LessThanOrEqualTo(val1)
                    }
                    "between" => {
                        let val2: f64 = dv_args
                            .get(2)
                            .context("DATA_VALIDATION between requires two values")?
                            .parse()
                            .context("DATA_VALIDATION: invalid second number")?;
                        DataValidationRule::Between(val1, val2)
                    }
                    "not_between" => {
                        let val2: f64 = dv_args
                            .get(2)
                            .context("DATA_VALIDATION not_between requires two values")?
                            .parse()
                            .context("DATA_VALIDATION: invalid second number")?;
                        DataValidationRule::NotBetween(val1, val2)
                    }
                    _ => bail!("Unknown DATA_VALIDATION criteria: '{criteria}'"),
                };

                DataValidation::new().allow_decimal_number(rule)
            }

            "text_length" => {
                if dv_args.is_empty() {
                    bail!("DATA_VALIDATION text_length requires: criteria value [value2]");
                }
                let criteria = dv_args[0];
                let val1: u32 = dv_args
                    .get(1)
                    .context("DATA_VALIDATION text_length requires a value")?
                    .parse()
                    .context("DATA_VALIDATION text_length: invalid number")?;

                let rule = match criteria {
                    "equal_to" => DataValidationRule::EqualTo(val1),
                    "not_equal_to" => DataValidationRule::NotEqualTo(val1),
                    "greater_than" => DataValidationRule::GreaterThan(val1),
                    "greater_than_or_equal_to" | "greater_than_or_equal" => {
                        DataValidationRule::GreaterThanOrEqualTo(val1)
                    }
                    "less_than" => DataValidationRule::LessThan(val1),
                    "less_than_or_equal_to" | "less_than_or_equal" => {
                        DataValidationRule::LessThanOrEqualTo(val1)
                    }
                    "between" => {
                        let val2: u32 = dv_args
                            .get(2)
                            .context("DATA_VALIDATION between requires two values")?
                            .parse()
                            .context("DATA_VALIDATION: invalid second number")?;
                        DataValidationRule::Between(val1, val2)
                    }
                    "not_between" => {
                        let val2: u32 = dv_args
                            .get(2)
                            .context("DATA_VALIDATION not_between requires two values")?
                            .parse()
                            .context("DATA_VALIDATION: invalid second number")?;
                        DataValidationRule::NotBetween(val1, val2)
                    }
                    _ => bail!("Unknown DATA_VALIDATION criteria: '{criteria}'"),
                };

                DataValidation::new().allow_text_length(rule)
            }

            _ => bail!("Unknown DATA_VALIDATION type: '{dv_type}'"),
        };

        let idx = self.current_ws_index.context("No worksheet")?;
        self.workbook
            .worksheet_from_index(idx)?
            .add_data_validation(r1, c1, r2, c2, &dv)?;
        Ok(())
    }

    // -- PAGE_SETUP key value ... ------------------------------------------
    fn cmd_page_setup(&mut self, args: &[&str]) -> Result<()> {
        if args.is_empty() {
            bail!("PAGE_SETUP requires: key value...");
        }
        let key = args[0];
        let idx = self.current_ws_index.context("No worksheet")?;
        let ws = self.workbook.worksheet_from_index(idx)?;

        match key {
            "orientation" => {
                let val = args.get(1).unwrap_or(&"portrait");
                match *val {
                    "landscape" => {
                        ws.set_landscape();
                    }
                    _ => {
                        ws.set_portrait();
                    }
                }
            }
            "paper_size" => {
                let size: u8 = args
                    .get(1)
                    .context("PAGE_SETUP paper_size requires a value")?
                    .parse()
                    .context("PAGE_SETUP paper_size: invalid number")?;
                ws.set_paper_size(size);
            }
            "margins" => {
                let left: f64 = args
                    .get(1)
                    .unwrap_or(&"0.7")
                    .parse()
                    .context("PAGE_SETUP margins: invalid left")?;
                let right: f64 = args
                    .get(2)
                    .unwrap_or(&"0.7")
                    .parse()
                    .context("PAGE_SETUP margins: invalid right")?;
                let top: f64 = args
                    .get(3)
                    .unwrap_or(&"0.75")
                    .parse()
                    .context("PAGE_SETUP margins: invalid top")?;
                let bottom: f64 = args
                    .get(4)
                    .unwrap_or(&"0.75")
                    .parse()
                    .context("PAGE_SETUP margins: invalid bottom")?;
                let header: f64 = args
                    .get(5)
                    .unwrap_or(&"0.3")
                    .parse()
                    .context("PAGE_SETUP margins: invalid header")?;
                let footer: f64 = args
                    .get(6)
                    .unwrap_or(&"0.3")
                    .parse()
                    .context("PAGE_SETUP margins: invalid footer")?;
                ws.set_margins(left, right, top, bottom, header, footer);
            }
            "header" => {
                let val = args[1..].join(" ");
                ws.set_header(&val);
            }
            "footer" => {
                let val = args[1..].join(" ");
                ws.set_footer(&val);
            }
            "print_scale" => {
                let scale: u16 = args
                    .get(1)
                    .context("PAGE_SETUP print_scale requires a value")?
                    .parse()
                    .context("PAGE_SETUP print_scale: invalid number")?;
                ws.set_print_scale(scale);
            }
            "print_gridlines" => {
                let val = args.get(1).unwrap_or(&"1");
                ws.set_print_gridlines(parse_bool(val));
            }
            "print_headings" => {
                let val = args.get(1).unwrap_or(&"1");
                ws.set_print_headings(parse_bool(val));
            }
            "print_area" => {
                let range_str = args
                    .get(1)
                    .context("PAGE_SETUP print_area requires a range")?;
                let (r1, c1, r2, c2) = parse_range(range_str)?;
                ws.set_print_area(r1, c1, r2, c2)?;
            }
            "repeat_rows" => {
                let first: u32 = args
                    .get(1)
                    .context("PAGE_SETUP repeat_rows requires first row")?
                    .parse::<u32>()
                    .context("PAGE_SETUP repeat_rows: invalid first row")?
                    .checked_sub(1)
                    .context("PAGE_SETUP repeat_rows: row must be >= 1")?;
                let last: u32 = args
                    .get(2)
                    .context("PAGE_SETUP repeat_rows requires last row")?
                    .parse::<u32>()
                    .context("PAGE_SETUP repeat_rows: invalid last row")?
                    .checked_sub(1)
                    .context("PAGE_SETUP repeat_rows: row must be >= 1")?;
                ws.set_repeat_rows(first, last)?;
            }
            "repeat_columns" => {
                let (first, last) = parse_col_range(
                    args.get(1)
                        .context("PAGE_SETUP repeat_columns requires col range")?,
                )?;
                ws.set_repeat_columns(first, last)?;
            }
            "center_horizontally" => {
                ws.set_print_center_horizontally(parse_bool(args.get(1).unwrap_or(&"1")));
            }
            "center_vertically" => {
                ws.set_print_center_vertically(parse_bool(args.get(1).unwrap_or(&"1")));
            }
            "print_first_page_number" => {
                if let Some(val) = args.get(1) {
                    if let Ok(n) = val.parse::<u16>() {
                        ws.set_print_first_page_number(n);
                    }
                }
            }
            _ => {
                eprintln!("Warning: unknown PAGE_SETUP property '{key}'");
            }
        }
        Ok(())
    }

    // -- HIDE_SHEET --------------------------------------------------------
    fn cmd_hide_sheet(&mut self, _args: &[&str]) -> Result<()> {
        let idx = self.current_ws_index.context("No worksheet")?;
        self.workbook.worksheet_from_index(idx)?.set_hidden(true);
        Ok(())
    }

    // -- PROTECT_SHEET [password] ------------------------------------------
    fn cmd_protect_sheet(&mut self, args: &[&str]) -> Result<()> {
        let idx = self.current_ws_index.context("No worksheet")?;
        let ws = self.workbook.worksheet_from_index(idx)?;
        if let Some(password) = args.first() {
            ws.protect_with_password(password);
        } else {
            ws.protect();
        }
        Ok(())
    }

    // -- DEFINE_NAME name formula ------------------------------------------
    fn cmd_define_name(&mut self, args: &[&str]) -> Result<()> {
        if args.len() < 2 {
            bail!("DEFINE_NAME requires: name formula");
        }
        let name = args[0];
        let formula = args[1..].join(" ");
        self.workbook.define_name(name, &formula)?;
        Ok(())
    }

    // -- RICH_TEXT [.styles] cell format1|text1||format2|text2||... ---------
    // Handled via execute_raw for whitespace preservation.
    fn cmd_rich_text_raw(&mut self, line: &str) -> Result<()> {
        let rest = line
            .strip_prefix("RICH_TEXT")
            .context("expected RICH_TEXT")?
            .trim_start();

        // Scan for inline cell-level style tokens and cell ref
        let mut cell_styles: Vec<(&str, &str)> = Vec::new();
        let mut remaining = rest;
        let mut cell_ref: Option<&str> = None;

        loop {
            remaining = remaining.trim_start();
            if remaining.is_empty() {
                break;
            }
            let token_end = remaining
                .find(|c: char| c.is_whitespace())
                .unwrap_or(remaining.len());
            let token = &remaining[..token_end];

            if is_style_token(token) {
                if let Some((key, value)) = token[1..].split_once(':') {
                    cell_styles.push((key, value));
                }
                remaining = &remaining[token_end..];
            } else if cell_ref.is_none() {
                cell_ref = Some(token);
                remaining = &remaining[token_end..];
                if !remaining.is_empty() {
                    remaining = &remaining[1..]; // skip separator
                }
                break;
            }
        }

        let cell_ref = cell_ref.context("RICH_TEXT requires: cell segments...")?;
        let (row, col) = parse_cell_ref(cell_ref)?;

        // Parse segments separated by ||
        let segments: Vec<&str> = remaining.split("||").collect();

        let mut rich_segments: Vec<(Format, String)> = Vec::new();

        for seg in segments {
            let seg = seg.trim();
            if seg.is_empty() {
                continue;
            }

            if let Some((fmt_str, text)) = seg.split_once('|') {
                let format = if fmt_str.trim() == "_" || fmt_str.trim().is_empty() {
                    Format::new()
                } else {
                    let mut format = Format::new();
                    for prop in fmt_str.split(',') {
                        if let Some((key, value)) = prop.trim().split_once(':') {
                            format = apply_property(format, key.trim(), value.trim());
                        }
                    }
                    format
                };
                rich_segments.push((format, text.to_string()));
            } else {
                // No pipe — treat entire segment as default-formatted text
                rich_segments.push((Format::new(), seg.to_string()));
            }
        }

        if rich_segments.is_empty() {
            bail!("RICH_TEXT requires at least one segment");
        }

        let refs: Vec<(&Format, &str)> =
            rich_segments.iter().map(|(f, t)| (f, t.as_str())).collect();

        let idx = self.current_ws_index.context("No worksheet")?;

        if cell_styles.is_empty() {
            self.workbook
                .worksheet_from_index(idx)?
                .write_rich_string(row, col, &refs)?;
        } else {
            let format = get_or_create_format(&mut self.inline_cache, &cell_styles);
            self.workbook
                .worksheet_from_index(idx)?
                .write_rich_string_with_format(row, col, &refs, &format)?;
        }
        Ok(())
    }

    // -- Save the workbook -------------------------------------------------
    fn save(mut self, filename: &str) -> Result<()> {
        if !self.doc_properties.is_empty() {
            let mut props = DocProperties::new();
            for (key, value) in &self.doc_properties {
                props = match key.as_str() {
                    "title" => props.set_title(value),
                    "author" => props.set_author(value),
                    "subject" => props.set_subject(value),
                    "manager" => props.set_manager(value),
                    "company" => props.set_company(value),
                    "category" => props.set_category(value),
                    "keywords" => props.set_keywords(value),
                    "comment" | "comments" => props.set_comment(value),
                    other => {
                        eprintln!("Warning: unknown document property '{other}'");
                        props
                    }
                };
            }
            self.workbook.set_properties(&props);
        }
        self.workbook.save(filename)?;
        Ok(())
    }
}

// ---------------------------------------------------------------------------
// main
// ---------------------------------------------------------------------------

fn main() -> Result<()> {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <file.xlsx>", args[0]);
        std::process::exit(1);
    }

    let filename = &args[1];
    let mut writer = XlsxWriter::new();
    let stdin = io::stdin();
    let mut stdout = io::stdout();

    for line in stdin.lock().lines() {
        let line = line.context("Failed to read stdin")?;

        // Echo every line to stdout for pipeline debugging
        writeln!(stdout, "{}", line)?;
        stdout.flush()?;

        let parts: Vec<&str> = line.split_whitespace().collect();
        if parts.is_empty() {
            continue;
        }

        let command = parts[0];

        // Commands that need raw line access (tab preservation / special delimiters)
        match command {
            "ROW" | "RICH_TEXT" => {
                if let Err(e) = writer.execute_raw(command, &line) {
                    eprintln!("Error in '{}': {}", line.trim(), e);
                }
                continue;
            }
            _ => {}
        }

        let cmd_args = &parts[1..];

        if let Err(e) = writer.execute(command, cmd_args) {
            eprintln!("Error in '{}': {}", line.trim(), e);
        }
    }

    writer.save(filename)?;
    Ok(())
}
