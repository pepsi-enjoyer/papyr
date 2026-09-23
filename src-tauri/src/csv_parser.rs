use crate::xlsx_model::{
    XlsxCell, XlsxCellValueType, XlsxPreviewLimits, XlsxRow, XlsxSheet, XlsxSheetSummary,
    XlsxWorkbook,
};
use std::fs;
use std::path::Path;

const MAX_CSV_FILE_SIZE: u64 = 100 * 1024 * 1024;

pub fn parse_workbook(path: &str) -> Result<XlsxWorkbook, String> {
    let summary = sheet_summary(path);
    let active_sheet = Some(read_sheet(path, &summary)?);

    Ok(XlsxWorkbook {
        sheets: vec![summary],
        active_sheet_index: 0,
        active_sheet,
        limits: XlsxPreviewLimits::default(),
    })
}

pub fn parse_sheet(path: &str, sheet_index: usize) -> Result<XlsxSheet, String> {
    if sheet_index != 0 {
        return Err(format!("Sheet index out of range: {sheet_index}"));
    }
    read_sheet(path, &sheet_summary(path))
}

fn sheet_summary(path: &str) -> XlsxSheetSummary {
    let name = Path::new(path)
        .file_stem()
        .and_then(|stem| stem.to_str())
        .filter(|stem| !stem.is_empty())
        .map(|stem| stem.to_string())
        .unwrap_or_else(|| "CSV".to_string());

    XlsxSheetSummary {
        index: 0,
        sheet_id: "1".to_string(),
        name,
        visible: true,
        state: None,
    }
}

fn read_sheet(path: &str, summary: &XlsxSheetSummary) -> Result<XlsxSheet, String> {
    let contents = read_csv_text(path)?;
    let delimiter = detect_delimiter(&contents);

    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(false)
        .flexible(true)
        .from_reader(contents.as_bytes());

    let limits = XlsxPreviewLimits::default();
    let mut truncated_reasons = Vec::new();
    let mut rows = Vec::new();
    let mut stored_cell_count = 0usize;
    let mut total_rows = 0u32;
    let mut max_column = 0u32;
    let mut row_limit_hit = false;
    let mut column_limit_hit = false;
    let mut cell_limit_hit = false;

    for record in reader.records() {
        let record = record.map_err(|e| format!("Failed to parse CSV file: {e}"))?;
        total_rows += 1;

        if record.len() as u32 > max_column {
            max_column = record.len() as u32;
        }
        if record.len() as u32 > limits.max_columns {
            column_limit_hit = true;
        }

        if total_rows > limits.max_rows {
            row_limit_hit = true;
            continue;
        }

        let row_index = total_rows;
        let mut cells = Vec::new();

        for (column_offset, field) in record.iter().enumerate().take(limits.max_columns as usize) {
            if stored_cell_count >= limits.max_cells {
                cell_limit_hit = true;
                break;
            }

            if let Some(cell) = build_cell(field, row_index, column_offset as u32 + 1) {
                cells.push(cell);
                stored_cell_count += 1;
            }
        }

        if !cells.is_empty() {
            rows.push(XlsxRow {
                index: row_index,
                cells,
                height: None,
            });
        }

        if stored_cell_count >= limits.max_cells {
            cell_limit_hit = true;
        }
    }

    if row_limit_hit {
        add_truncation_reason(
            &mut truncated_reasons,
            format!("Only the first {} rows are loaded.", limits.max_rows),
        );
    }
    if column_limit_hit {
        add_truncation_reason(
            &mut truncated_reasons,
            format!("Only the first {} columns are loaded.", limits.max_columns),
        );
    }
    if cell_limit_hit {
        add_truncation_reason(
            &mut truncated_reasons,
            format!(
                "Only the first {} non-empty cells are loaded.",
                limits.max_cells
            ),
        );
    }

    Ok(XlsxSheet {
        index: summary.index,
        name: summary.name.clone(),
        rows,
        row_count: total_rows,
        column_count: max_column,
        max_row: total_rows,
        max_column,
        truncated: !truncated_reasons.is_empty(),
        truncated_reasons,
        images: Vec::new(),
        default_col_width: None,
        default_row_height: None,
        columns: Vec::new(),
    })
}

fn read_csv_text(path: &str) -> Result<String, String> {
    let metadata = fs::metadata(path).map_err(|e| format!("Failed to open CSV file: {e}"))?;
    let file_size = metadata.len();
    if file_size == 0 {
        return Err("Invalid CSV file: empty file".to_string());
    }
    if file_size > MAX_CSV_FILE_SIZE {
        return Err("CSV file too large: exceeds 100 MB".to_string());
    }

    let bytes = fs::read(path).map_err(|e| format!("Failed to read CSV file: {e}"))?;
    let text = match String::from_utf8(bytes) {
        Ok(text) => text,
        Err(err) => err
            .into_bytes()
            .iter()
            .map(|&byte| byte as char)
            .collect(),
    };

    Ok(strip_bom(&text).to_string())
}

fn strip_bom(text: &str) -> &str {
    text.strip_prefix('\u{feff}').unwrap_or(text)
}

fn detect_delimiter(contents: &str) -> u8 {
    let first_line = contents.lines().next().unwrap_or("");
    let mut in_quotes = false;
    let mut counts = [0usize; 3];
    const CANDIDATES: [char; 3] = [',', ';', '\t'];

    for ch in first_line.chars() {
        if ch == '"' {
            in_quotes = !in_quotes;
            continue;
        }
        if in_quotes {
            continue;
        }
        for (idx, candidate) in CANDIDATES.iter().enumerate() {
            if ch == *candidate {
                counts[idx] += 1;
            }
        }
    }

    let best = counts
        .iter()
        .enumerate()
        .max_by_key(|(_, count)| **count)
        .filter(|(_, count)| **count > 0)
        .map(|(idx, _)| idx)
        .unwrap_or(0);

    CANDIDATES[best] as u8
}

fn build_cell(field: &str, row: u32, column: u32) -> Option<XlsxCell> {
    if field.is_empty() {
        return None;
    }

    let value_type = if is_number(field) {
        XlsxCellValueType::Number
    } else {
        XlsxCellValueType::String
    };

    Some(XlsxCell {
        reference: format!("{}{}", column_name(column), row),
        row,
        column,
        value: field.to_string(),
        raw_value: Some(field.to_string()),
        value_type,
        formula: None,
        style_index: None,
        number_format: None,
    })
}

fn is_number(field: &str) -> bool {
    let trimmed = field.trim();
    if trimmed.is_empty() {
        return false;
    }
    trimmed.parse::<f64>().is_ok()
}

fn add_truncation_reason(reasons: &mut Vec<String>, reason: String) {
    if !reasons.iter().any(|existing| existing == &reason) {
        reasons.push(reason);
    }
}

fn column_name(mut index: u32) -> String {
    if index == 0 {
        return String::new();
    }

    let mut chars = Vec::new();
    while index > 0 {
        index -= 1;
        chars.push((b'A' + (index % 26) as u8) as char);
        index /= 26;
    }
    chars.iter().rev().collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn write_temp(name: &str, contents: &[u8]) -> String {
        let mut path = std::env::temp_dir();
        path.push(name);
        let mut file = fs::File::create(&path).unwrap();
        file.write_all(contents).unwrap();
        path.to_string_lossy().into_owned()
    }

    #[test]
    fn parses_basic_comma_csv() {
        let path = write_temp("papyr_basic.csv", b"name,age\nAlice,30\nBob,25\n");
        let workbook = parse_workbook(&path).unwrap();
        let sheet = workbook.active_sheet.unwrap();

        assert_eq!(sheet.row_count, 3);
        assert_eq!(sheet.column_count, 2);
        assert_eq!(sheet.rows.len(), 3);

        let age = &sheet.rows[1].cells[1];
        assert_eq!(age.value, "30");
        assert_eq!(age.value_type, XlsxCellValueType::Number);
        assert_eq!(age.reference, "B2");

        let name = &sheet.rows[1].cells[0];
        assert_eq!(name.value, "Alice");
        assert_eq!(name.value_type, XlsxCellValueType::String);
    }

    #[test]
    fn detects_semicolon_and_strips_bom() {
        let path = write_temp(
            "papyr_semicolon.csv",
            "\u{feff}a;b;c\n1;2;3\n".as_bytes(),
        );
        let sheet = parse_sheet(&path, 0).unwrap();
        assert_eq!(sheet.column_count, 3);
        assert_eq!(sheet.rows[0].cells[0].value, "a");
    }

    #[test]
    fn handles_quoted_fields_with_delimiters() {
        let path = write_temp("papyr_quoted.csv", b"\"a,b\",c\n");
        let sheet = parse_sheet(&path, 0).unwrap();
        assert_eq!(sheet.rows[0].cells.len(), 2);
        assert_eq!(sheet.rows[0].cells[0].value, "a,b");
    }
}
