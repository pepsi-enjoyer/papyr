use crate::xlsx_model::{
    XlsxCell, XlsxCellValueType, XlsxColumn, XlsxImage, XlsxPreviewLimits, XlsxRow, XlsxSheet,
    XlsxSheetSummary, XlsxWorkbook,
};
use base64::Engine;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{BufReader, Read};
use zip::ZipArchive;

const MAX_XLSX_FILE_SIZE: u64 = 100 * 1024 * 1024;
const MAX_XML_PART_SIZE: u64 = 128 * 1024 * 1024;

#[derive(Debug, thiserror::Error)]
pub enum XlsxParseError {
    #[error("File not found or cannot be accessed: {0}")]
    FileNotFound(std::io::Error),
    #[error("Not a valid XLSX workbook: {0}")]
    NotAnXlsxFile(zip::result::ZipError),
    #[error("ZIP error: {0}")]
    Zip(zip::result::ZipError),
    #[error("Corrupted XML: {0}")]
    Xml(quick_xml::Error),
    #[error("Unsupported XLSX format: {0}")]
    UnsupportedFormat(String),
    #[error("Invalid XLSX format: {0}")]
    InvalidFormat(String),
    #[error("Missing file: {0}")]
    MissingFile(String),
    #[error("XLSX file too large")]
    FileTooLarge(String),
    #[error("XLSX part too large: {0}")]
    PartTooLarge(String),
    #[error("Sheet index out of range: {0}")]
    SheetOutOfRange(usize),
}

impl From<std::io::Error> for XlsxParseError {
    fn from(e: std::io::Error) -> Self {
        XlsxParseError::FileNotFound(e)
    }
}

impl From<quick_xml::Error> for XlsxParseError {
    fn from(e: quick_xml::Error) -> Self {
        XlsxParseError::Xml(e)
    }
}

impl From<zip::result::ZipError> for XlsxParseError {
    fn from(e: zip::result::ZipError) -> Self {
        XlsxParseError::Zip(e)
    }
}

pub type Result<T> = std::result::Result<T, XlsxParseError>;

pub struct XlsxParser {
    archive: ZipArchive<File>,
}

impl XlsxParser {
    pub fn from_path(path: &str) -> Result<Self> {
        let file = File::open(path).map_err(XlsxParseError::FileNotFound)?;
        let file_size = file.metadata().map_err(XlsxParseError::FileNotFound)?.len();

        if file_size == 0 {
            return Err(XlsxParseError::InvalidFormat("Empty workbook".into()));
        }
        if file_size > MAX_XLSX_FILE_SIZE {
            return Err(XlsxParseError::FileTooLarge("Exceeds 100 MB".into()));
        }

        let archive = ZipArchive::new(file).map_err(XlsxParseError::NotAnXlsxFile)?;
        Ok(Self { archive })
    }

    pub fn parse(&mut self) -> Result<XlsxWorkbook> {
        let workbook = self.read_workbook_info()?;
        let active_sheet_index = normalize_active_sheet_index(&workbook);
        let styles = self.read_styles().unwrap_or_default();
        let active_sheet = workbook
            .sheets
            .get(active_sheet_index)
            .map(|sheet| self.parse_sheet_data(sheet, &styles, workbook.date1904))
            .transpose()?;

        Ok(XlsxWorkbook {
            sheets: workbook.sheets.iter().map(SheetInfo::to_summary).collect(),
            active_sheet_index,
            active_sheet,
            limits: XlsxPreviewLimits::default(),
        })
    }

    pub fn parse_sheet(&mut self, sheet_index: usize) -> Result<XlsxSheet> {
        let workbook = self.read_workbook_info()?;
        let sheet = workbook
            .sheets
            .get(sheet_index)
            .ok_or(XlsxParseError::SheetOutOfRange(sheet_index))?
            .clone();
        let styles = self.read_styles().unwrap_or_default();
        self.parse_sheet_data(&sheet, &styles, workbook.date1904)
    }

    fn read_workbook_info(&mut self) -> Result<WorkbookInfo> {
        let content = self.read_archive_file("xl/workbook.xml", MAX_XML_PART_SIZE)?;
        let relationships = self.read_workbook_relationships()?;
        let mut reader = Reader::from_str(&content);
        let mut buf = Vec::new();
        let mut sheets = Vec::new();
        let mut active_sheet_index = 0usize;
        let mut date1904 = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    match local_name(e.name().as_ref()) {
                        b"workbookPr" => {
                            date1904 = get_attr(e, b"date1904")
                                .map(|value| is_truthy(&value))
                                .unwrap_or(false);
                        }
                        b"workbookView" => {
                            if let Some(value) =
                                get_attr(e, b"activeTab").and_then(|v| v.parse().ok())
                            {
                                active_sheet_index = value;
                            }
                        }
                        b"sheet" => {
                            let relationship_id = get_attr(e, b"r:id").ok_or_else(|| {
                                XlsxParseError::InvalidFormat(
                                    "Sheet is missing relationship id".into(),
                                )
                            })?;
                            let relationship =
                                relationships.get(&relationship_id).ok_or_else(|| {
                                    XlsxParseError::InvalidFormat(format!(
                                        "Sheet relationship '{}' is missing",
                                        relationship_id
                                    ))
                                })?;
                            let index = sheets.len();
                            let state = get_attr(e, b"state");
                            sheets.push(SheetInfo {
                                index,
                                sheet_id: get_attr(e, b"sheetId")
                                    .unwrap_or_else(|| (index + 1).to_string()),
                                name: get_attr(e, b"name")
                                    .unwrap_or_else(|| format!("Sheet {}", index + 1)),
                                state: state.clone(),
                                visible: !matches!(
                                    state.as_deref(),
                                    Some("hidden") | Some("veryHidden")
                                ),
                                path: relationship.target.clone(),
                            });
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxParseError::Xml(e)),
                _ => {}
            }
            buf.clear();
        }

        if sheets.is_empty() {
            return Err(XlsxParseError::UnsupportedFormat(
                "Workbook does not contain sheets".into(),
            ));
        }

        Ok(WorkbookInfo {
            sheets,
            active_sheet_index,
            date1904,
        })
    }

    fn read_workbook_relationships(&mut self) -> Result<HashMap<String, WorkbookRelationship>> {
        let content = self.read_archive_file("xl/_rels/workbook.xml.rels", MAX_XML_PART_SIZE)?;
        let mut reader = Reader::from_str(&content);
        let mut buf = Vec::new();
        let mut relationships = HashMap::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                    if local_name(e.name().as_ref()) == b"Relationship" =>
                {
                    let id = get_attr(e, b"Id");
                    let target = get_attr(e, b"Target");
                    let rel_type = get_attr(e, b"Type");
                    if let (Some(id), Some(target), Some(rel_type)) = (id, target, rel_type) {
                        if rel_type.ends_with("/worksheet") {
                            relationships.insert(
                                id,
                                WorkbookRelationship {
                                    target: resolve_package_path("xl", &target),
                                },
                            );
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxParseError::Xml(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(relationships)
    }

    fn read_styles(&mut self) -> Result<WorkbookStyles> {
        if self.archive.by_name("xl/styles.xml").is_err() {
            return Ok(WorkbookStyles::default());
        }

        let content = self.read_archive_file("xl/styles.xml", MAX_XML_PART_SIZE)?;
        let mut reader = Reader::from_str(&content);
        let mut buf = Vec::new();
        let mut custom_formats: HashMap<u32, String> = HashMap::new();
        let mut cell_formats = Vec::new();
        let mut in_cell_xfs = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    match local_name(e.name().as_ref()) {
                        b"numFmt" => {
                            if let (Some(id), Some(format_code)) = (
                                get_attr(e, b"numFmtId").and_then(|v| v.parse().ok()),
                                get_attr(e, b"formatCode"),
                            ) {
                                custom_formats.insert(id, format_code);
                            }
                        }
                        b"cellXfs" => in_cell_xfs = true,
                        b"xf" if in_cell_xfs => {
                            let num_fmt_id = get_attr(e, b"numFmtId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0);
                            cell_formats.push(CellFormat::new(
                                num_fmt_id,
                                custom_formats.get(&num_fmt_id).cloned(),
                            ));
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(ref e)) if local_name(e.name().as_ref()) == b"cellXfs" => {
                    in_cell_xfs = false;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxParseError::Xml(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(WorkbookStyles { cell_formats })
    }

    fn parse_sheet_data(
        &mut self,
        sheet: &SheetInfo,
        styles: &WorkbookStyles,
        date1904: bool,
    ) -> Result<XlsxSheet> {
        let parsed = self.parse_raw_sheet(sheet, styles, date1904)?;
        let shared_strings = self.read_needed_shared_strings(&parsed.shared_string_indexes)?;
        let mut result = parsed.into_sheet(&shared_strings);
        result.images = self.parse_sheet_images(&sheet.path);
        Ok(result)
    }

    fn parse_raw_sheet(
        &mut self,
        sheet: &SheetInfo,
        styles: &WorkbookStyles,
        date1904: bool,
    ) -> Result<ParsedSheet> {
        let file = self
            .archive
            .by_name(&sheet.path)
            .map_err(|_| XlsxParseError::MissingFile(sheet.path.clone()))?;
        if file.size() > MAX_XML_PART_SIZE {
            return Err(XlsxParseError::PartTooLarge(sheet.path.clone()));
        }

        let limits = XlsxPreviewLimits::default();
        let mut reader = Reader::from_reader(BufReader::new(file));
        let mut buf = Vec::new();
        let mut rows = Vec::new();
        let mut current_row: Option<ParsedRow> = None;
        let mut current_cell: Option<CellCtx> = None;
        let mut current_row_index: u32 = 0;
        let mut fallback_row_index: u32 = 1;
        let mut last_column_index: u32 = 0;
        let mut stored_cell_count = 0usize;
        let mut shared_string_indexes = HashSet::new();
        let mut dimension_row_count = None;
        let mut dimension_column_count = None;
        let mut max_observed_row = 0;
        let mut max_observed_column = 0;
        let mut truncated_reasons = Vec::new();
        let mut collecting_value = false;
        let mut collecting_formula = false;
        let mut collecting_inline_text = false;
        let mut in_inline_string = false;
        let mut stop = false;
        let mut default_col_width = None;
        let mut default_row_height = None;
        let mut columns: Vec<XlsxColumn> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => match local_name(e.name().as_ref()) {
                    b"sheetFormatPr" => {
                        let (col_width, row_height) = parse_sheet_format_pr(e);
                        default_col_width = default_col_width.or(col_width);
                        default_row_height = default_row_height.or(row_height);
                    }
                    b"col" => {
                        if let Some(column) = parse_col(e) {
                            columns.push(column);
                        }
                    }
                    b"dimension" => {
                        if let Some((row_count, column_count)) =
                            get_attr(e, b"ref").and_then(|value| parse_dimension_ref(&value))
                        {
                            dimension_row_count = Some(row_count);
                            dimension_column_count = Some(column_count);
                            add_dimension_truncation_reasons(
                                row_count,
                                column_count,
                                &limits,
                                &mut truncated_reasons,
                            );
                        }
                    }
                    b"row" => {
                        let row_index = get_attr(e, b"r")
                            .and_then(|value| value.parse().ok())
                            .unwrap_or(fallback_row_index);
                        fallback_row_index = row_index.saturating_add(1);
                        current_row_index = row_index;
                        last_column_index = 0;
                        if row_index > limits.max_rows {
                            add_truncation_reason(
                                &mut truncated_reasons,
                                format!("Only the first {} rows are loaded.", limits.max_rows),
                            );
                            stop = true;
                        } else {
                            current_row = Some(ParsedRow {
                                index: row_index,
                                cells: Vec::new(),
                                height: get_attr(e, b"ht").and_then(|value| value.parse().ok()),
                            });
                        }
                    }
                    b"c" => {
                        current_cell = Some(begin_cell(e, current_row_index, last_column_index));
                        if let Some(cell) = &current_cell {
                            last_column_index = cell.column;
                        }
                    }
                    b"v" => collecting_value = true,
                    b"f" => collecting_formula = true,
                    b"is" => in_inline_string = true,
                    b"t" if in_inline_string => collecting_inline_text = true,
                    _ => {}
                },
                Ok(Event::Empty(ref e)) => match local_name(e.name().as_ref()) {
                    b"sheetFormatPr" => {
                        let (col_width, row_height) = parse_sheet_format_pr(e);
                        default_col_width = default_col_width.or(col_width);
                        default_row_height = default_row_height.or(row_height);
                    }
                    b"col" => {
                        if let Some(column) = parse_col(e) {
                            columns.push(column);
                        }
                    }
                    b"dimension" => {
                        if let Some((row_count, column_count)) =
                            get_attr(e, b"ref").and_then(|value| parse_dimension_ref(&value))
                        {
                            dimension_row_count = Some(row_count);
                            dimension_column_count = Some(column_count);
                            add_dimension_truncation_reasons(
                                row_count,
                                column_count,
                                &limits,
                                &mut truncated_reasons,
                            );
                        }
                    }
                    b"row" => {
                        let row_index = get_attr(e, b"r")
                            .and_then(|value| value.parse().ok())
                            .unwrap_or(fallback_row_index);
                        fallback_row_index = row_index.saturating_add(1);
                        if row_index > limits.max_rows {
                            add_truncation_reason(
                                &mut truncated_reasons,
                                format!("Only the first {} rows are loaded.", limits.max_rows),
                            );
                            stop = true;
                        }
                    }
                    b"c" => {
                        let cell = begin_cell(e, current_row_index, last_column_index);
                        last_column_index = cell.column;
                        finish_cell(
                            cell,
                            &mut current_row,
                            styles,
                            date1904,
                            &limits,
                            &mut stored_cell_count,
                            &mut shared_string_indexes,
                            &mut max_observed_row,
                            &mut max_observed_column,
                            &mut truncated_reasons,
                        );
                    }
                    b"v" => collecting_value = false,
                    b"f" => collecting_formula = false,
                    b"is" => in_inline_string = false,
                    b"t" if in_inline_string => collecting_inline_text = false,
                    _ => {}
                },
                Ok(Event::Text(ref e)) => {
                    if let Some(cell) = current_cell.as_mut() {
                        if let Ok(text) = e.unescape() {
                            if collecting_value {
                                cell.raw_value.push_str(&text);
                            } else if collecting_formula {
                                cell.formula.get_or_insert_with(String::new).push_str(&text);
                            } else if collecting_inline_text {
                                cell.inline_text.push_str(&text);
                            }
                        }
                    }
                }
                Ok(Event::End(ref e)) => match local_name(e.name().as_ref()) {
                    b"v" => collecting_value = false,
                    b"f" => collecting_formula = false,
                    b"is" => in_inline_string = false,
                    b"t" if in_inline_string => collecting_inline_text = false,
                    b"c" => {
                        if let Some(cell) = current_cell.take() {
                            finish_cell(
                                cell,
                                &mut current_row,
                                styles,
                                date1904,
                                &limits,
                                &mut stored_cell_count,
                                &mut shared_string_indexes,
                                &mut max_observed_row,
                                &mut max_observed_column,
                                &mut truncated_reasons,
                            );
                            if stored_cell_count >= limits.max_cells {
                                add_truncation_reason(
                                    &mut truncated_reasons,
                                    format!(
                                        "Only the first {} non-empty cells are loaded.",
                                        limits.max_cells
                                    ),
                                );
                                stop = true;
                            }
                        }
                    }
                    b"row" => {
                        push_current_row(&mut rows, &mut current_row);
                    }
                    b"sheetData" => break,
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxParseError::Xml(e)),
                _ => {}
            }

            if stop {
                break;
            }
            buf.clear();
        }

        push_current_row(&mut rows, &mut current_row);

        let row_count = dimension_row_count.unwrap_or(max_observed_row);
        let column_count = dimension_column_count.unwrap_or(max_observed_column);

        Ok(ParsedSheet {
            index: sheet.index,
            name: sheet.name.clone(),
            rows,
            row_count,
            column_count,
            max_row: max_observed_row,
            max_column: max_observed_column,
            truncated: !truncated_reasons.is_empty(),
            truncated_reasons,
            shared_string_indexes,
            default_col_width,
            default_row_height,
            columns,
        })
    }

    fn read_needed_shared_strings(
        &mut self,
        needed: &HashSet<usize>,
    ) -> Result<HashMap<usize, String>> {
        if needed.is_empty() {
            return Ok(HashMap::new());
        }

        let file = match self.archive.by_name("xl/sharedStrings.xml") {
            Ok(file) => file,
            Err(_) => return Ok(HashMap::new()),
        };
        if file.size() > MAX_XML_PART_SIZE {
            return Err(XlsxParseError::PartTooLarge("xl/sharedStrings.xml".into()));
        }

        let mut reader = Reader::from_reader(BufReader::new(file));
        let mut buf = Vec::new();
        let mut values = HashMap::new();
        let mut current_index = 0usize;
        let mut in_si = false;
        let mut collecting_text = false;
        let mut text_buf = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => match local_name(e.name().as_ref()) {
                    b"si" => {
                        in_si = true;
                        text_buf.clear();
                    }
                    b"t" if in_si => collecting_text = true,
                    _ => {}
                },
                Ok(Event::Empty(ref e)) => match local_name(e.name().as_ref()) {
                    b"si" => {
                        if needed.contains(&current_index) {
                            values.insert(current_index, String::new());
                        }
                        current_index += 1;
                    }
                    b"t" if in_si => {}
                    _ => {}
                },
                Ok(Event::Text(ref e)) => {
                    if collecting_text {
                        if let Ok(text) = e.unescape() {
                            text_buf.push_str(&text);
                        }
                    }
                }
                Ok(Event::End(ref e)) => match local_name(e.name().as_ref()) {
                    b"t" => collecting_text = false,
                    b"si" => {
                        if needed.contains(&current_index) {
                            values.insert(current_index, text_buf.clone());
                            if values.len() == needed.len() {
                                break;
                            }
                        }
                        current_index += 1;
                        in_si = false;
                        text_buf.clear();
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxParseError::Xml(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(values)
    }

    fn read_archive_file(&mut self, path: &str, max_size: u64) -> Result<String> {
        let mut file = self
            .archive
            .by_name(path)
            .map_err(|_| XlsxParseError::MissingFile(path.to_string()))?;
        if file.size() > max_size {
            return Err(XlsxParseError::PartTooLarge(path.to_string()));
        }

        let mut content = String::new();
        file.read_to_string(&mut content)?;
        Ok(content)
    }

    fn read_archive_bytes(&mut self, path: &str, max_size: u64) -> Option<Vec<u8>> {
        let mut file = self.archive.by_name(path).ok()?;
        if file.size() > max_size {
            return None;
        }
        let mut buffer = Vec::with_capacity(file.size() as usize);
        file.read_to_end(&mut buffer).ok()?;
        Some(buffer)
    }

    /// Collects every picture anchored onto a worksheet, resolving the chain
    /// worksheet -> drawing -> media so the front end can place each image where
    /// it sits in the workbook. Any missing or malformed part is skipped rather
    /// than failing the whole sheet.
    fn parse_sheet_images(&mut self, sheet_path: &str) -> Vec<XlsxImage> {
        let sheet_rels = rels_path_for(sheet_path);
        let sheet_dir = parent_dir(sheet_path);
        let drawing_paths = self.read_relationship_targets(&sheet_rels, &sheet_dir, "/drawing");

        let mut images = Vec::new();
        for drawing_path in drawing_paths {
            images.extend(self.parse_drawing(&drawing_path));
        }
        images
    }

    fn parse_drawing(&mut self, drawing_path: &str) -> Vec<XlsxImage> {
        let Some(content) = self
            .read_archive_bytes(drawing_path, MAX_XML_PART_SIZE)
            .and_then(|bytes| String::from_utf8(bytes).ok())
        else {
            return Vec::new();
        };

        let drawing_rels = rels_path_for(drawing_path);
        let drawing_dir = parent_dir(drawing_path);
        let media_by_id = self.read_relationship_map(&drawing_rels, &drawing_dir);

        let anchors = parse_drawing_anchors(&content);
        let mut images = Vec::new();
        for anchor in anchors {
            let Some(embed_id) = anchor.embed.as_ref() else {
                continue;
            };
            let Some(media_path) = media_by_id.get(embed_id) else {
                continue;
            };
            let Some(bytes) = self.read_archive_bytes(media_path, MAX_XML_PART_SIZE) else {
                continue;
            };
            let mime = mime_for_path(media_path);
            let encoded = base64::engine::general_purpose::STANDARD.encode(&bytes);
            images.push(anchor.into_image(format!("data:{};base64,{}", mime, encoded)));
        }
        images
    }

    /// Resolves the targets of every relationship whose type ends with
    /// `type_suffix` (e.g. "/drawing"), returning package paths.
    fn read_relationship_targets(
        &mut self,
        rels_path: &str,
        base_dir: &str,
        type_suffix: &str,
    ) -> Vec<String> {
        let Ok(content) = self.read_archive_file(rels_path, MAX_XML_PART_SIZE) else {
            return Vec::new();
        };

        let mut reader = Reader::from_str(&content);
        let mut buf = Vec::new();
        let mut targets = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                    if local_name(e.name().as_ref()) == b"Relationship" =>
                {
                    let rel_type = get_attr(e, b"Type").unwrap_or_default();
                    if rel_type.ends_with(type_suffix) {
                        if let Some(target) = get_attr(e, b"Target") {
                            targets.push(resolve_package_path(base_dir, &target));
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        targets
    }

    /// Maps every relationship id to its resolved package path.
    fn read_relationship_map(
        &mut self,
        rels_path: &str,
        base_dir: &str,
    ) -> HashMap<String, String> {
        let Ok(content) = self.read_archive_file(rels_path, MAX_XML_PART_SIZE) else {
            return HashMap::new();
        };

        let mut reader = Reader::from_str(&content);
        let mut buf = Vec::new();
        let mut map = HashMap::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                    if local_name(e.name().as_ref()) == b"Relationship" =>
                {
                    if let (Some(id), Some(target)) = (get_attr(e, b"Id"), get_attr(e, b"Target")) {
                        map.insert(id, resolve_package_path(base_dir, &target));
                    }
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        map
    }
}

#[derive(Debug, Clone)]
struct WorkbookInfo {
    sheets: Vec<SheetInfo>,
    active_sheet_index: usize,
    date1904: bool,
}

#[derive(Debug, Clone)]
struct SheetInfo {
    index: usize,
    sheet_id: String,
    name: String,
    state: Option<String>,
    visible: bool,
    path: String,
}

impl SheetInfo {
    fn to_summary(&self) -> XlsxSheetSummary {
        XlsxSheetSummary {
            index: self.index,
            sheet_id: self.sheet_id.clone(),
            name: self.name.clone(),
            visible: self.visible,
            state: self.state.clone(),
        }
    }
}

struct WorkbookRelationship {
    target: String,
}

#[derive(Debug, Default, Clone)]
struct WorkbookStyles {
    cell_formats: Vec<CellFormat>,
}

#[derive(Debug, Clone)]
struct CellFormat {
    format_code: Option<String>,
    is_date: bool,
    is_text: bool,
    is_percent: bool,
    percent_decimal_places: usize,
}

impl CellFormat {
    fn new(num_fmt_id: u32, custom_format_code: Option<String>) -> Self {
        let format_code =
            custom_format_code.or_else(|| builtin_number_format(num_fmt_id).map(str::to_string));
        let normalized = format_code
            .as_deref()
            .map(normalize_format_code)
            .unwrap_or_default();
        let is_date = is_date_number_format(num_fmt_id, &normalized);
        let is_text = num_fmt_id == 49 || normalized.trim() == "@";
        let is_percent = normalized.contains('%');
        let percent_decimal_places = percent_decimal_places(&normalized);

        Self {
            format_code,
            is_date,
            is_text,
            is_percent,
            percent_decimal_places,
        }
    }
}

struct ParsedSheet {
    index: usize,
    name: String,
    rows: Vec<ParsedRow>,
    row_count: u32,
    column_count: u32,
    max_row: u32,
    max_column: u32,
    truncated: bool,
    truncated_reasons: Vec<String>,
    shared_string_indexes: HashSet<usize>,
    default_col_width: Option<f64>,
    default_row_height: Option<f64>,
    columns: Vec<XlsxColumn>,
}

impl ParsedSheet {
    fn into_sheet(self, shared_strings: &HashMap<usize, String>) -> XlsxSheet {
        XlsxSheet {
            index: self.index,
            name: self.name,
            rows: self
                .rows
                .into_iter()
                .map(|row| XlsxRow {
                    index: row.index,
                    cells: row
                        .cells
                        .into_iter()
                        .map(|cell| cell.into_cell(shared_strings))
                        .collect(),
                    height: row.height,
                })
                .collect(),
            row_count: self.row_count,
            column_count: self.column_count,
            max_row: self.max_row,
            max_column: self.max_column,
            truncated: self.truncated,
            truncated_reasons: self.truncated_reasons,
            images: Vec::new(),
            default_col_width: self.default_col_width,
            default_row_height: self.default_row_height,
            columns: self.columns,
        }
    }
}

struct ParsedRow {
    index: u32,
    cells: Vec<ParsedCell>,
    height: Option<f64>,
}

struct ParsedCell {
    cell: XlsxCell,
    shared_string_index: Option<usize>,
}

impl ParsedCell {
    fn into_cell(mut self, shared_strings: &HashMap<usize, String>) -> XlsxCell {
        if let Some(index) = self.shared_string_index {
            self.cell.value = shared_strings.get(&index).cloned().unwrap_or_default();
        }
        self.cell
    }
}

#[derive(Debug)]
struct CellCtx {
    reference: String,
    row: u32,
    column: u32,
    cell_type: Option<String>,
    style_index: Option<u32>,
    raw_value: String,
    formula: Option<String>,
    inline_text: String,
}

fn normalize_active_sheet_index(workbook: &WorkbookInfo) -> usize {
    if workbook.active_sheet_index < workbook.sheets.len() {
        return workbook.active_sheet_index;
    }

    workbook
        .sheets
        .iter()
        .find(|sheet| sheet.visible)
        .map(|sheet| sheet.index)
        .unwrap_or(0)
}

fn begin_cell(e: &BytesStart, current_row_index: u32, previous_column: u32) -> CellCtx {
    let fallback_column = previous_column.saturating_add(1).max(1);
    let fallback_row = current_row_index.max(1);
    let reference = get_attr(e, b"r")
        .unwrap_or_else(|| format!("{}{}", column_name(fallback_column), fallback_row));
    let (column, row) = parse_cell_reference(&reference).unwrap_or((fallback_column, fallback_row));

    CellCtx {
        reference,
        row,
        column,
        cell_type: get_attr(e, b"t"),
        style_index: get_attr(e, b"s").and_then(|value| value.parse().ok()),
        raw_value: String::new(),
        formula: None,
        inline_text: String::new(),
    }
}

fn finish_cell(
    cell: CellCtx,
    current_row: &mut Option<ParsedRow>,
    styles: &WorkbookStyles,
    date1904: bool,
    limits: &XlsxPreviewLimits,
    stored_cell_count: &mut usize,
    shared_string_indexes: &mut HashSet<usize>,
    max_observed_row: &mut u32,
    max_observed_column: &mut u32,
    truncated_reasons: &mut Vec<String>,
) {
    if cell.row > limits.max_rows {
        add_truncation_reason(
            truncated_reasons,
            format!("Only the first {} rows are loaded.", limits.max_rows),
        );
        return;
    }
    if cell.column > limits.max_columns {
        add_truncation_reason(
            truncated_reasons,
            format!("Only the first {} columns are loaded.", limits.max_columns),
        );
        return;
    }
    if *stored_cell_count >= limits.max_cells {
        add_truncation_reason(
            truncated_reasons,
            format!(
                "Only the first {} non-empty cells are loaded.",
                limits.max_cells
            ),
        );
        return;
    }

    let parsed = build_cell(cell, styles, date1904);
    if let Some(parsed) = parsed {
        if let Some(index) = parsed.shared_string_index {
            shared_string_indexes.insert(index);
        }
        *max_observed_row = (*max_observed_row).max(parsed.cell.row);
        *max_observed_column = (*max_observed_column).max(parsed.cell.column);
        if let Some(row) = current_row.as_mut() {
            row.cells.push(parsed);
            *stored_cell_count += 1;
        }
    }
}

fn build_cell(cell: CellCtx, styles: &WorkbookStyles, date1904: bool) -> Option<ParsedCell> {
    let style = cell
        .style_index
        .and_then(|index| styles.cell_formats.get(index as usize));
    let raw_value = cell.raw_value.trim().to_string();
    let formula = cell.formula.filter(|value| !value.trim().is_empty());
    let number_format = style.and_then(|format| format.format_code.clone());
    let style_index = cell.style_index;

    let (value, value_type, shared_string_index) = match cell.cell_type.as_deref() {
        Some("s") => {
            let index = raw_value.parse::<usize>().ok();
            (String::new(), XlsxCellValueType::String, index)
        }
        Some("inlineStr") => (cell.inline_text, XlsxCellValueType::String, None),
        Some("b") => (
            if raw_value == "1" {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            },
            XlsxCellValueType::Boolean,
            None,
        ),
        Some("e") => (raw_value.clone(), XlsxCellValueType::Error, None),
        Some("str") => (raw_value.clone(), XlsxCellValueType::String, None),
        _ if style.map(|format| format.is_text).unwrap_or(false) => {
            (raw_value.clone(), XlsxCellValueType::String, None)
        }
        _ => format_numeric_value(&raw_value, style, date1904),
    };

    if value.is_empty()
        && raw_value.is_empty()
        && formula.is_none()
        && shared_string_index.is_none()
    {
        return None;
    }

    Some(ParsedCell {
        cell: XlsxCell {
            reference: cell.reference,
            row: cell.row,
            column: cell.column,
            value,
            raw_value: if raw_value.is_empty() {
                None
            } else {
                Some(raw_value)
            },
            value_type,
            formula,
            style_index,
            number_format,
        },
        shared_string_index,
    })
}

fn format_numeric_value(
    raw_value: &str,
    style: Option<&CellFormat>,
    date1904: bool,
) -> (String, XlsxCellValueType, Option<usize>) {
    if raw_value.is_empty() {
        return (String::new(), XlsxCellValueType::Blank, None);
    }

    if let Some(style) = style {
        if style.is_date {
            if let Some(value) = excel_serial_to_display(raw_value, date1904) {
                return (value, XlsxCellValueType::Date, None);
            }
        }
        if style.is_percent {
            if let Ok(value) = raw_value.parse::<f64>() {
                return (
                    format!("{:.*}%", style.percent_decimal_places, value * 100.0),
                    XlsxCellValueType::Number,
                    None,
                );
            }
        }
    }

    (trim_number(raw_value), XlsxCellValueType::Number, None)
}

fn push_current_row(rows: &mut Vec<ParsedRow>, current_row: &mut Option<ParsedRow>) {
    if let Some(row) = current_row.take() {
        if !row.cells.is_empty() {
            rows.push(row);
        }
    }
}

fn add_dimension_truncation_reasons(
    row_count: u32,
    column_count: u32,
    limits: &XlsxPreviewLimits,
    truncated_reasons: &mut Vec<String>,
) {
    if row_count > limits.max_rows {
        add_truncation_reason(
            truncated_reasons,
            format!("Only the first {} rows are loaded.", limits.max_rows),
        );
    }
    if column_count > limits.max_columns {
        add_truncation_reason(
            truncated_reasons,
            format!("Only the first {} columns are loaded.", limits.max_columns),
        );
    }
}

fn add_truncation_reason(reasons: &mut Vec<String>, reason: String) {
    if !reasons.iter().any(|existing| existing == &reason) {
        reasons.push(reason);
    }
}

fn parse_dimension_ref(reference: &str) -> Option<(u32, u32)> {
    let last = reference.split(':').next_back().unwrap_or(reference);
    parse_cell_reference(last).map(|(column, row)| (row, column))
}

fn parse_cell_reference(reference: &str) -> Option<(u32, u32)> {
    let mut column = 0u32;
    let mut row = 0u32;

    for ch in reference.chars() {
        if ch == '$' {
            continue;
        }
        if ch.is_ascii_alphabetic() {
            column = column
                .saturating_mul(26)
                .saturating_add((ch.to_ascii_uppercase() as u8 - b'A' + 1) as u32);
        } else if ch.is_ascii_digit() {
            row = row
                .saturating_mul(10)
                .saturating_add(ch.to_digit(10).unwrap_or(0));
        }
    }

    if column == 0 || row == 0 {
        None
    } else {
        Some((column, row))
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

fn resolve_package_path(base_dir: &str, target: &str) -> String {
    if target.starts_with('/') {
        return normalize_package_path(target.trim_start_matches('/'));
    }

    normalize_package_path(&format!("{}/{}", base_dir.trim_end_matches('/'), target))
}

fn normalize_package_path(path: &str) -> String {
    let normalized_path = path.replace('\\', "/");
    let mut parts = Vec::new();
    for part in normalized_path.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(part),
        }
    }
    parts.join("/")
}

fn parse_sheet_format_pr(e: &BytesStart) -> (Option<f64>, Option<f64>) {
    let col_width = get_attr(e, b"defaultColWidth").and_then(|value| value.parse().ok());
    let row_height = get_attr(e, b"defaultRowHeight").and_then(|value| value.parse().ok());
    (col_width, row_height)
}

fn parse_col(e: &BytesStart) -> Option<XlsxColumn> {
    let min = get_attr(e, b"min").and_then(|value| value.parse().ok())?;
    let max = get_attr(e, b"max").and_then(|value| value.parse().ok())?;
    let width = get_attr(e, b"width").and_then(|value| value.parse().ok())?;
    let hidden = get_attr(e, b"hidden")
        .map(|value| is_truthy(&value))
        .unwrap_or(false);
    Some(XlsxColumn {
        min,
        max,
        width,
        hidden,
    })
}

fn rels_path_for(path: &str) -> String {
    match path.rfind('/') {
        Some(index) => format!("{}/_rels/{}.rels", &path[..index], &path[index + 1..]),
        None => format!("_rels/{}.rels", path),
    }
}

fn parent_dir(path: &str) -> String {
    match path.rfind('/') {
        Some(index) => path[..index].to_string(),
        None => String::new(),
    }
}

fn mime_for_path(path: &str) -> &'static str {
    let lower = path.to_lowercase();
    if lower.ends_with(".png") {
        "image/png"
    } else if lower.ends_with(".jpg") || lower.ends_with(".jpeg") {
        "image/jpeg"
    } else if lower.ends_with(".gif") {
        "image/gif"
    } else if lower.ends_with(".bmp") {
        "image/bmp"
    } else if lower.ends_with(".svg") {
        "image/svg+xml"
    } else if lower.ends_with(".webp") {
        "image/webp"
    } else if lower.ends_with(".tif") || lower.ends_with(".tiff") {
        "image/tiff"
    } else {
        "application/octet-stream"
    }
}

#[derive(Default)]
struct DrawingAnchor {
    anchor_type: String,
    from_col: u32,
    from_col_off: i64,
    from_row: u32,
    from_row_off: i64,
    to_col: Option<u32>,
    to_col_off: Option<i64>,
    to_row: Option<u32>,
    to_row_off: Option<i64>,
    ext_cx: Option<i64>,
    ext_cy: Option<i64>,
    embed: Option<String>,
}

impl DrawingAnchor {
    fn into_image(self, data_uri: String) -> XlsxImage {
        XlsxImage {
            anchor_type: self.anchor_type,
            from_col: self.from_col,
            from_col_off: self.from_col_off,
            from_row: self.from_row,
            from_row_off: self.from_row_off,
            to_col: self.to_col,
            to_col_off: self.to_col_off,
            to_row: self.to_row,
            to_row_off: self.to_row_off,
            ext_cx: self.ext_cx,
            ext_cy: self.ext_cy,
            data_uri,
        }
    }
}

#[derive(Clone, Copy, PartialEq)]
enum AnchorSide {
    None,
    From,
    To,
}

#[derive(Clone, Copy, PartialEq)]
enum AnchorField {
    None,
    Col,
    ColOff,
    Row,
    RowOff,
}

/// Parses the anchors out of a SpreadsheetDrawingML part. Each `twoCellAnchor`,
/// `oneCellAnchor` or `absoluteAnchor` that embeds a picture becomes one anchor.
fn parse_drawing_anchors(content: &str) -> Vec<DrawingAnchor> {
    let mut reader = Reader::from_str(content);
    let mut buf = Vec::new();
    let mut anchors = Vec::new();
    let mut current: Option<DrawingAnchor> = None;
    let mut side = AnchorSide::None;
    let mut field = AnchorField::None;
    let mut text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match local_name(e.name().as_ref()) {
                b"twoCellAnchor" => current = Some(new_anchor("two")),
                b"oneCellAnchor" => current = Some(new_anchor("one")),
                b"absoluteAnchor" => current = Some(new_anchor("absolute")),
                b"from" => side = AnchorSide::From,
                b"to" => side = AnchorSide::To,
                b"col" => {
                    field = AnchorField::Col;
                    text.clear();
                }
                b"colOff" => {
                    field = AnchorField::ColOff;
                    text.clear();
                }
                b"row" => {
                    field = AnchorField::Row;
                    text.clear();
                }
                b"rowOff" => {
                    field = AnchorField::RowOff;
                    text.clear();
                }
                b"blip" => capture_embed(e, &mut current),
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match local_name(e.name().as_ref()) {
                b"ext" => capture_ext(e, &mut current),
                b"blip" => capture_embed(e, &mut current),
                _ => {}
            },
            Ok(Event::Text(ref e)) => {
                if field != AnchorField::None {
                    if let Ok(value) = e.unescape() {
                        text.push_str(&value);
                    }
                }
            }
            Ok(Event::End(ref e)) => match local_name(e.name().as_ref()) {
                b"col" | b"colOff" | b"row" | b"rowOff" => {
                    commit_field(current.as_mut(), side, field, text.trim());
                    field = AnchorField::None;
                }
                b"from" | b"to" => side = AnchorSide::None,
                b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => {
                    if let Some(anchor) = current.take() {
                        anchors.push(anchor);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    anchors
}

fn new_anchor(anchor_type: &str) -> DrawingAnchor {
    DrawingAnchor {
        anchor_type: anchor_type.to_string(),
        ..DrawingAnchor::default()
    }
}

fn capture_embed(e: &BytesStart, current: &mut Option<DrawingAnchor>) {
    if let Some(anchor) = current.as_mut() {
        if anchor.embed.is_none() {
            anchor.embed = get_attr(e, b"r:embed").or_else(|| get_attr(e, b"embed"));
        }
    }
}

fn capture_ext(e: &BytesStart, current: &mut Option<DrawingAnchor>) {
    if let Some(anchor) = current.as_mut() {
        // Only the anchor's own <xdr:ext> matters; ignore the <a:ext> shape
        // extents nested inside the picture (the first ext seen is the real one).
        if anchor.ext_cx.is_none() && anchor.ext_cy.is_none() {
            anchor.ext_cx = get_attr(e, b"cx").and_then(|value| value.parse().ok());
            anchor.ext_cy = get_attr(e, b"cy").and_then(|value| value.parse().ok());
        }
    }
}

fn commit_field(
    anchor: Option<&mut DrawingAnchor>,
    side: AnchorSide,
    field: AnchorField,
    text: &str,
) {
    let Some(anchor) = anchor else {
        return;
    };
    match (side, field) {
        (AnchorSide::From, AnchorField::Col) => {
            anchor.from_col = text.parse().unwrap_or(0);
        }
        (AnchorSide::From, AnchorField::ColOff) => {
            anchor.from_col_off = text.parse().unwrap_or(0);
        }
        (AnchorSide::From, AnchorField::Row) => {
            anchor.from_row = text.parse().unwrap_or(0);
        }
        (AnchorSide::From, AnchorField::RowOff) => {
            anchor.from_row_off = text.parse().unwrap_or(0);
        }
        (AnchorSide::To, AnchorField::Col) => {
            anchor.to_col = text.parse().ok();
        }
        (AnchorSide::To, AnchorField::ColOff) => {
            anchor.to_col_off = text.parse().ok();
        }
        (AnchorSide::To, AnchorField::Row) => {
            anchor.to_row = text.parse().ok();
        }
        (AnchorSide::To, AnchorField::RowOff) => {
            anchor.to_row_off = text.parse().ok();
        }
        _ => {}
    }
}

fn get_attr(e: &BytesStart, key: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == key {
            return std::str::from_utf8(attr.value.as_ref())
                .ok()
                .map(unescape_xml_entities);
        }
    }
    None
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn unescape_xml_entities(value: &str) -> String {
    value
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
}

fn is_truthy(value: &str) -> bool {
    matches!(value, "1" | "true" | "TRUE" | "True")
}

fn builtin_number_format(id: u32) -> Option<&'static str> {
    match id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        14 => Some("mm-dd-yy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        37 => Some("#,##0 ;(#,##0)"),
        38 => Some("#,##0 ;[Red](#,##0)"),
        39 => Some("#,##0.00;(#,##0.00)"),
        40 => Some("#,##0.00;[Red](#,##0.00)"),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mmss.0"),
        49 => Some("@"),
        _ => None,
    }
}

fn normalize_format_code(format_code: &str) -> String {
    let mut normalized = String::new();
    let mut chars = format_code.chars().peekable();
    let mut in_quote = false;
    let mut in_bracket = false;

    while let Some(ch) = chars.next() {
        match ch {
            '"' => in_quote = !in_quote,
            '[' if !in_quote => in_bracket = true,
            ']' if !in_quote => in_bracket = false,
            '\\' if !in_quote => {
                chars.next();
            }
            '_' | '*' if !in_quote => {
                chars.next();
            }
            _ if !in_quote && !in_bracket => normalized.push(ch.to_ascii_lowercase()),
            _ => {}
        }
    }

    normalized
}

fn is_date_number_format(num_fmt_id: u32, normalized_format: &str) -> bool {
    if matches!(num_fmt_id, 14..=22 | 45..=47) {
        return true;
    }
    if normalized_format.contains('y') || normalized_format.contains('d') {
        return true;
    }
    normalized_format.contains('h')
        && (normalized_format.contains('m') || normalized_format.contains('s'))
}

fn percent_decimal_places(normalized_format: &str) -> usize {
    let Some(percent_index) = normalized_format.find('%') else {
        return 0;
    };
    let before_percent = &normalized_format[..percent_index];
    let Some(decimal_index) = before_percent.rfind('.') else {
        return 0;
    };

    before_percent[decimal_index + 1..]
        .chars()
        .filter(|ch| matches!(ch, '0' | '#'))
        .count()
}

fn trim_number(raw_value: &str) -> String {
    let Ok(value) = raw_value.parse::<f64>() else {
        return raw_value.to_string();
    };
    if !value.is_finite() {
        return raw_value.to_string();
    }
    if value.fract() == 0.0 {
        format!("{value:.0}")
    } else {
        value.to_string()
    }
}

fn excel_serial_to_display(raw_value: &str, date1904: bool) -> Option<String> {
    let serial = raw_value.parse::<f64>().ok()?;
    if !serial.is_finite() {
        return None;
    }

    let whole_days = serial.floor() as i64;
    let fraction = serial - whole_days as f64;
    let offset = if date1904 { 24_107 } else { 25_569 };
    let mut total_seconds = (fraction * 86_400.0).round() as i64;
    let day_adjustment = total_seconds.div_euclid(86_400);
    total_seconds = total_seconds.rem_euclid(86_400);

    let (year, month, day) = civil_from_days(whole_days - offset + day_adjustment);
    if total_seconds == 0 {
        return Some(format!("{year:04}-{month:02}-{day:02}"));
    }

    let hour = total_seconds / 3_600;
    let minute = (total_seconds % 3_600) / 60;
    let second = total_seconds % 60;
    if second == 0 {
        Some(format!(
            "{year:04}-{month:02}-{day:02} {hour:02}:{minute:02}"
        ))
    } else {
        Some(format!(
            "{year:04}-{month:02}-{day:02} {hour:02}:{minute:02}:{second:02}"
        ))
    }
}

fn civil_from_days(days_since_unix_epoch: i64) -> (i32, u32, u32) {
    let days = days_since_unix_epoch + 719_468;
    let era = if days >= 0 { days } else { days - 146_096 } / 146_097;
    let day_of_era = days - era * 146_097;
    let year_of_era =
        (day_of_era - day_of_era / 1_460 + day_of_era / 36_524 - day_of_era / 146_096) / 365;
    let year = year_of_era + era * 400;
    let day_of_year = day_of_era - (365 * year_of_era + year_of_era / 4 - year_of_era / 100);
    let month_prime = (5 * day_of_year + 2) / 153;
    let day = day_of_year - (153 * month_prime + 2) / 5 + 1;
    let month = month_prime + if month_prime < 10 { 3 } else { -9 };
    let year = year + if month <= 2 { 1 } else { 0 };

    (year as i32, month as u32, day as u32)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs::{self, File};
    use std::io::Write;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};
    use zip::write::FileOptions;

    #[test]
    fn parses_workbook_preview_with_common_cell_types() {
        let xlsx_path = write_test_xlsx(&[
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
            ),
            (
                "xl/sharedStrings.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><t>Hello</t></si>
</sst>"#,
            ),
            (
                "xl/styles.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cellXfs count="2">
    <xf numFmtId="0"/>
    <xf numFmtId="14"/>
  </cellXfs>
</styleSheet>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:D2"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1"><v>42</v></c>
      <c r="C1" t="b"><v>1</v></c>
      <c r="D1" s="1"><v>45292</v></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Inline</t></is></c>
      <c r="B2"><f>SUM(B1,8)</f><v>50</v></c>
    </row>
  </sheetData>
</worksheet>"#,
            ),
        ]);

        let mut parser = XlsxParser::from_path(xlsx_path.to_str().unwrap()).expect("parser init");
        let workbook = parser.parse().expect("parse workbook");
        let sheet = workbook.active_sheet.expect("active sheet");

        assert_eq!(workbook.sheets[0].name, "Summary");
        assert_eq!(sheet.rows[0].cells[0].value, "Hello");
        assert_eq!(sheet.rows[0].cells[1].value, "42");
        assert_eq!(sheet.rows[0].cells[2].value, "TRUE");
        assert_eq!(sheet.rows[0].cells[3].value, "2024-01-01");
        assert_eq!(sheet.rows[1].cells[0].value, "Inline");
        assert_eq!(sheet.rows[1].cells[1].formula.as_deref(), Some("SUM(B1,8)"));
        assert_eq!(sheet.rows[1].cells[1].value, "50");

        let _ = fs::remove_file(xlsx_path);
    }

    #[test]
    fn parses_sheet_by_index() {
        let xlsx_path = write_test_xlsx(&[
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="First" sheetId="1" r:id="rId1"/>
    <sheet name="Second" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#,
            ),
            (
                "xl/worksheets/sheet2.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Loaded</t></is></c></row></sheetData></worksheet>"#,
            ),
        ]);

        let mut parser = XlsxParser::from_path(xlsx_path.to_str().unwrap()).expect("parser init");
        let sheet = parser.parse_sheet(1).expect("parse second sheet");

        assert_eq!(sheet.name, "Second");
        assert_eq!(sheet.rows[0].cells[0].value, "Loaded");

        let _ = fs::remove_file(xlsx_path);
    }

    #[test]
    fn extracts_anchored_worksheet_image() {
        let xlsx_path = write_test_xlsx(&[
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Pictures" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
  <drawing r:id="rId1"/>
</worksheet>"#,
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>"#,
            ),
            (
                "xl/drawings/drawing1.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>2</xdr:col><xdr:colOff>10</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>20</xdr:rowOff></xdr:from>
    <xdr:ext cx="100" cy="200"/>
    <xdr:pic><xdr:blipFill><a:blip r:embed="rIdImg1"/></xdr:blipFill></xdr:pic>
  </xdr:oneCellAnchor>
</xdr:wsDr>"#,
            ),
            (
                "xl/drawings/_rels/drawing1.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#,
            ),
            ("xl/media/image1.png", "fake-png-bytes"),
        ]);

        let mut parser = XlsxParser::from_path(xlsx_path.to_str().unwrap()).expect("parser init");
        let workbook = parser.parse().expect("parse workbook");
        let sheet = workbook.active_sheet.expect("active sheet");

        assert_eq!(sheet.images.len(), 1);
        let image = &sheet.images[0];
        assert_eq!(image.anchor_type, "one");
        assert_eq!(image.from_col, 2);
        assert_eq!(image.from_col_off, 10);
        assert_eq!(image.from_row, 3);
        assert_eq!(image.from_row_off, 20);
        assert_eq!(image.ext_cx, Some(100));
        assert_eq!(image.ext_cy, Some(200));
        assert!(image.data_uri.starts_with("data:image/png;base64,"));

        let _ = fs::remove_file(xlsx_path);
    }

    #[test]
    fn rejects_missing_workbook_part() {
        let xlsx_path = write_test_xlsx(&[(
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
        )]);

        let mut parser = XlsxParser::from_path(xlsx_path.to_str().unwrap()).expect("parser init");
        assert!(parser.parse().is_err());

        let _ = fs::remove_file(xlsx_path);
    }

    fn write_test_xlsx(files: &[(&str, &str)]) -> PathBuf {
        let path = std::env::temp_dir().join(format!(
            "papyr-xlsx-parser-test-{}-{}.xlsx",
            std::process::id(),
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .expect("clock drift")
                .as_nanos()
        ));

        let file = File::create(&path).expect("create xlsx file");
        let mut zip = zip::ZipWriter::new(file);
        let options: FileOptions<'_, ()> = FileOptions::default();

        for &(file_name, content) in files {
            zip.start_file(file_name, options).expect("start zip file");
            zip.write_all(content.as_bytes()).expect("write zip file");
        }
        zip.finish().expect("finish xlsx");

        path
    }
}
