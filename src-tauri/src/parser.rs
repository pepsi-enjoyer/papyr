use crate::model::{
    BlockElement, Comment, Document, Footnote, HeaderFooter, Run, Style, TableCell, TableRow,
};
use base64::Engine;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::Read;
use zip::ZipArchive;

#[derive(Debug, Clone)]
struct Relationship {
    id: String,
    rel_type: String,
    target: String,
}

#[derive(Debug, thiserror::Error)]
pub enum ParseError {
    #[error("File not found or cannot be accessed: {0}")]
    FileNotFound(std::io::Error),
    #[error("Not a valid DOCX document: {0}")]
    NotADocxFile(zip::result::ZipError),
    #[error("Corrupted XML: {0}")]
    Xml(quick_xml::Error),
    #[error("ZIP error: {0}")]
    Zip(zip::result::ZipError),
    #[error("Unsupported format: {0}")]
    UnsupportedFormat(String),
    #[error("Invalid format: {0}")]
    InvalidFormat(String),
    #[error("Missing file: {0}")]
    MissingFile(String),
    #[error("Document too large")]
    DocumentTooLarge(String),
    #[error("Empty document")]
    EmptyDocument,
}

impl From<std::io::Error> for ParseError {
    fn from(e: std::io::Error) -> Self {
        ParseError::FileNotFound(e)
    }
}

impl From<quick_xml::Error> for ParseError {
    fn from(e: quick_xml::Error) -> Self {
        ParseError::Xml(e)
    }
}

pub type Result<T> = std::result::Result<T, ParseError>;

#[derive(Clone)]
struct NumberingLevelInfo {
    num_fmt: String,
    #[allow(dead_code)]
    lvl_text: String,
    start: u32,
}

pub struct DocxParser {
    archive: ZipArchive<File>,
    relationships: Vec<Relationship>,
    referenced_image_ids: HashSet<String>,
    numbering_levels: HashMap<(u32, u8), NumberingLevelInfo>,
    /// Maps a paragraph style id to the (numId, ilvl) it carries, so paragraphs
    /// that join a list purely through their style still render as list items.
    style_numbering: HashMap<String, (u32, u8)>,
}

impl DocxParser {
    pub fn from_path(path: &str) -> Result<Self> {
        let file = File::open(path).map_err(ParseError::FileNotFound)?;
        let file_size = file.metadata().map_err(ParseError::FileNotFound)?.len();

        if file_size == 0 {
            return Err(ParseError::EmptyDocument);
        }
        if file_size > 100 * 1024 * 1024 {
            return Err(ParseError::DocumentTooLarge("Exceeds 100 MB".into()));
        }

        let archive = ZipArchive::new(file).map_err(ParseError::NotADocxFile)?;
        let mut parser = Self {
            archive,
            relationships: Vec::new(),
            referenced_image_ids: HashSet::new(),
            numbering_levels: HashMap::new(),
            style_numbering: HashMap::new(),
        };

        // Validate required files
        if parser.archive.by_name("word/document.xml").is_err() {
            return Err(ParseError::UnsupportedFormat(
                "Missing word/document.xml".into(),
            ));
        }

        // Pre-parse relationships and numbering
        parser.relationships = parser.read_relationships();
        parser.numbering_levels = parser.parse_numbering();

        Ok(parser)
    }

    pub fn parse(&mut self) -> Result<Document> {
        let mut document = Document::new();

        // Styles must be resolved before the body so paragraphs can inherit list
        // membership from their paragraph style.
        document.styles = self.parse_styles().unwrap_or_default();
        self.style_numbering = document
            .styles
            .iter()
            .filter_map(|(id, style)| {
                style
                    .num_id
                    .map(|num_id| (id.clone(), (num_id, style.ilvl.unwrap_or(0))))
            })
            .collect();
        document.body = self.parse_document_body()?;
        let comment_threads = self.parse_comment_threads().unwrap_or_default();
        document.comments = self.parse_comments().unwrap_or_default();
        self.attach_comment_threads(&mut document.comments, comment_threads);
        let (headers, footers) = self.parse_headers_footers().unwrap_or_default();
        document.headers = headers;
        document.footers = footers;
        document.footnotes = self.parse_footnotes().unwrap_or_default();
        document.images = self.parse_images().unwrap_or_default();

        Ok(document)
    }

    // --- Document body parsing ---

    fn parse_document_body(&mut self) -> Result<Vec<BlockElement>> {
        let content = self.read_archive_file("word/document.xml")?;
        let mut reader = Reader::from_str(&content);

        let mut body = Vec::new();
        let mut buf = Vec::new();

        // Parsing context stacks
        let mut in_body = false;
        let mut para_stack: Vec<ParagraphCtx> = Vec::new();
        let mut table_stack: Vec<TableCtx> = Vec::new();
        let mut run_ctx: Option<RunCtx> = None;
        let mut collecting_text = false;
        let mut text_buf = String::new();
        let mut comment_ranges: Vec<u32> = Vec::new();
        let mut hyperlink_url: Option<String> = None;
        let mut in_num_pr = false;
        let mut list_counters: HashMap<(u32, u8), u32> = HashMap::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    let tag = e.name();
                    match tag.as_ref() {
                        b"w:body" => in_body = true,

                        b"w:p" if in_body || !table_stack.is_empty() => {
                            para_stack.push(ParagraphCtx::new());
                        }

                        b"w:pPr" => {}

                        b"w:pStyle" => {
                            if let Some(ctx) = para_stack.last_mut() {
                                ctx.style = get_attr(e, b"w:val");
                            }
                        }

                        b"w:jc" => {
                            if let Some(ctx) = para_stack.last_mut() {
                                ctx.alignment = get_attr(e, b"w:val");
                            }
                        }

                        b"w:hyperlink" => {
                            hyperlink_url = get_attr(e, b"r:id").and_then(|id| {
                                self.relationships
                                    .iter()
                                    .find(|r| r.id == id)
                                    .map(|r| r.target.clone())
                            });
                        }

                        b"w:numPr" => {
                            in_num_pr = true;
                        }
                        b"w:ilvl" if in_num_pr => {
                            if let Some(ctx) = para_stack.last_mut() {
                                ctx.ilvl = get_attr(e, b"w:val").and_then(|v| v.parse().ok());
                            }
                        }
                        b"w:numId" if in_num_pr => {
                            if let Some(ctx) = para_stack.last_mut() {
                                ctx.num_id = get_attr(e, b"w:val").and_then(|v| v.parse().ok());
                            }
                        }

                        b"w:r" => {
                            if !para_stack.is_empty() {
                                let mut rctx = RunCtx::new();
                                rctx.link_url = hyperlink_url.clone();
                                run_ctx = Some(rctx);
                            }
                        }

                        b"w:rPr" => {}

                        b"w:b" => {
                            if let Some(ref mut r) = run_ctx {
                                r.bold = !is_val_false(e);
                            }
                        }
                        b"w:i" => {
                            if let Some(ref mut r) = run_ctx {
                                r.italic = !is_val_false(e);
                            }
                        }
                        b"w:u" => {
                            if let Some(ref mut r) = run_ctx {
                                let val = get_attr(e, b"w:val");
                                r.underline = val.as_deref() != Some("none");
                            }
                        }
                        b"w:strike" => {
                            if let Some(ref mut r) = run_ctx {
                                r.strikethrough = !is_val_false(e);
                            }
                        }
                        b"w:sz" => {
                            if let Some(ref mut r) = run_ctx {
                                if let Some(val) = get_attr(e, b"w:val") {
                                    if let Ok(half_pt) = val.parse::<f32>() {
                                        r.font_size = Some(half_pt / 2.0);
                                    }
                                }
                            }
                        }
                        b"w:rFonts" => {
                            if let Some(ref mut r) = run_ctx {
                                r.font_family = extract_font_family(e);
                            }
                        }
                        b"w:color" => {
                            if let Some(ref mut r) = run_ctx {
                                r.color = get_attr(e, b"w:val");
                            }
                        }
                        b"w:highlight" => {
                            if let Some(ref mut r) = run_ctx {
                                r.highlight = get_attr(e, b"w:val");
                            }
                        }

                        b"w:t" => {
                            collecting_text = true;
                            text_buf.clear();
                        }

                        b"w:tab" => {
                            if let Some(ref mut r) = run_ctx {
                                r.text.push('\t');
                            }
                        }

                        b"w:br" => {
                            let br_type = get_attr(e, b"w:type");
                            if br_type.as_deref() == Some("page") {
                                // Flush current paragraph context if any, then add page break
                                if let Some(pctx) = para_stack.last_mut() {
                                    // Add the run so far
                                    if let Some(rctx) = run_ctx.take() {
                                        pctx.runs.push(rctx.into_run(&comment_ranges));
                                    }
                                }
                                // The page break will be emitted after the current paragraph ends
                                if let Some(pctx) = para_stack.last_mut() {
                                    pctx.has_page_break_after = true;
                                }
                            } else if let Some(ref mut r) = run_ctx {
                                r.text.push('\n');
                            }
                        }

                        // Comment range markers
                        b"w:commentRangeStart" => {
                            if let Some(id_str) = get_attr(e, b"w:id") {
                                if let Ok(id) = id_str.parse::<u32>() {
                                    comment_ranges.push(id);
                                }
                            }
                        }
                        b"w:commentRangeEnd" => {
                            if let Some(id_str) = get_attr(e, b"w:id") {
                                if let Ok(id) = id_str.parse::<u32>() {
                                    if let Some(pos) =
                                        comment_ranges.iter().position(|&cid| cid == id)
                                    {
                                        comment_ranges.remove(pos);
                                    }
                                }
                            }
                        }

                        // Footnote references
                        b"w:footnoteReference" => {
                            if let Some(ref mut r) = run_ctx {
                                if let Some(id_str) = get_attr(e, b"w:id") {
                                    r.footnote_ref = id_str.parse().ok();
                                }
                            }
                        }

                        // Image references via w:drawing -> ... -> a:blip
                        b"a:blip" => {
                            if let Some(ref mut r) = run_ctx {
                                if let Some(image_id) = get_attr(e, b"r:embed") {
                                    self.referenced_image_ids.insert(image_id.clone());
                                    r.image_id = Some(image_id);
                                }
                            }
                        }
                        b"v:imagedata" => {
                            if let Some(ref mut r) = run_ctx {
                                if let Some(image_id) = get_attr(e, b"r:id") {
                                    self.referenced_image_ids.insert(image_id.clone());
                                    r.image_id = Some(image_id);
                                }
                            }
                        }

                        // Tables
                        b"w:tbl" => {
                            table_stack.push(TableCtx::new());
                        }
                        b"w:tr" => {
                            if let Some(tctx) = table_stack.last_mut() {
                                tctx.current_row = Some(Vec::new());
                            }
                        }
                        b"w:tc" => {
                            if table_stack
                                .last()
                                .and_then(|t| t.current_row.as_ref())
                                .is_some()
                            {
                                // Push a new cell context -- paragraphs will go into it
                            }
                        }
                        b"w:gridSpan" => {
                            // Will be handled at tc level
                        }
                        b"w:vMerge" => {}
                        b"w:shd" => {}

                        _ => {}
                    }
                }

                Ok(Event::Text(ref e)) => {
                    if collecting_text {
                        if let Ok(text) = e.unescape() {
                            text_buf.push_str(&text);
                        }
                    }
                }

                Ok(Event::End(ref e)) => {
                    match e.name().as_ref() {
                        b"w:body" => in_body = false,

                        b"w:t" => {
                            collecting_text = false;
                            if let Some(ref mut r) = run_ctx {
                                r.text.push_str(&text_buf);
                            }
                        }

                        b"w:r" => {
                            if let Some(rctx) = run_ctx.take() {
                                if let Some(pctx) = para_stack.last_mut() {
                                    pctx.runs.push(rctx.into_run(&comment_ranges));
                                }
                            }
                        }

                        b"w:hyperlink" => {
                            hyperlink_url = None;
                        }
                        b"w:numPr" => {
                            in_num_pr = false;
                        }

                        b"w:p" => {
                            if let Some(pctx) = para_stack.pop() {
                                let page_break = pctx.has_page_break_after;

                                // Compute list info. A paragraph joins a list either
                                // through an inline numPr or through its paragraph
                                // style; ilvl defaults to 0 when omitted.
                                let style_list = pctx
                                    .style
                                    .as_ref()
                                    .and_then(|id| self.style_numbering.get(id))
                                    .copied();
                                let effective_num_id =
                                    pctx.num_id.or(style_list.map(|(num_id, _)| num_id));
                                let effective_ilvl = pctx
                                    .ilvl
                                    .or(style_list.map(|(_, ilvl)| ilvl))
                                    .unwrap_or(0);
                                let (list_level, list_format) = match effective_num_id {
                                    Some(num_id) if num_id > 0 => {
                                        if let Some(info) =
                                            self.numbering_levels.get(&(num_id, effective_ilvl))
                                        {
                                            let counter = list_counters
                                                .entry((num_id, effective_ilvl))
                                                .or_insert(info.start.saturating_sub(1));
                                            *counter += 1;
                                            for higher in (effective_ilvl + 1)..=8 {
                                                list_counters.remove(&(num_id, higher));
                                            }
                                            (Some(effective_ilvl), Some(info.num_fmt.clone()))
                                        } else {
                                            (Some(effective_ilvl), Some("bullet".to_string()))
                                        }
                                    }
                                    _ => (None, None),
                                };

                                let element = pctx.into_block_element(list_level, list_format);

                                if let Some(tctx) = table_stack.last_mut() {
                                    tctx.push_cell_content(element);
                                } else {
                                    body.push(element);
                                    if page_break {
                                        body.push(BlockElement::PageBreak);
                                    }
                                }
                            }
                        }

                        b"w:tc" => {
                            if let Some(tctx) = table_stack.last_mut() {
                                tctx.finish_cell();
                            }
                        }

                        b"w:tr" => {
                            if let Some(tctx) = table_stack.last_mut() {
                                tctx.finish_row();
                            }
                        }

                        b"w:tbl" => {
                            if let Some(tctx) = table_stack.pop() {
                                let table = BlockElement::Table { rows: tctx.rows };
                                if let Some(parent_tctx) = table_stack.last_mut() {
                                    parent_tctx.push_cell_content(table);
                                } else {
                                    body.push(table);
                                }
                            }
                        }

                        _ => {}
                    }
                }

                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(body)
    }

    // --- Styles ---

    fn parse_styles(&mut self) -> Result<HashMap<String, Style>> {
        let content = match self.read_archive_file("word/styles.xml") {
            Ok(c) => c,
            Err(_) => return Ok(HashMap::new()),
        };

        let mut styles = HashMap::new();
        let mut reader = Reader::from_str(&content);
        let mut buf = Vec::new();

        let mut current_style_id: Option<String> = None;
        let mut current_style = Style::new();
        let mut in_rpr = false;
        let mut in_ppr = false;
        let mut in_style_num_pr = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    match e.name().as_ref() {
                        b"w:style" => {
                            current_style_id = get_attr(e, b"w:styleId");
                            current_style = Style::new();
                            // Check if it's a heading style
                            if let Some(ref id) = current_style_id {
                                if id.starts_with("Heading") || id.starts_with("heading") {
                                    if let Some(level) =
                                        id.chars().last().and_then(|c| c.to_digit(10))
                                    {
                                        if level >= 1 && level <= 6 {
                                            current_style.heading_level = Some(level as u8);
                                        }
                                    }
                                }
                            }
                        }
                        b"w:name" => {
                            if let Some(val) = get_attr(e, b"w:val") {
                                let lower = val.to_lowercase();
                                if lower.starts_with("heading ") {
                                    if let Some(level) = lower
                                        .strip_prefix("heading ")
                                        .and_then(|s| s.parse::<u8>().ok())
                                    {
                                        if level >= 1 && level <= 6 {
                                            current_style.heading_level = Some(level);
                                        }
                                    }
                                }
                            }
                        }
                        b"w:basedOn" => {
                            current_style.based_on = get_attr(e, b"w:val");
                        }
                        b"w:rPr" => in_rpr = true,
                        b"w:pPr" => in_ppr = true,
                        b"w:numPr" if in_ppr => in_style_num_pr = true,
                        b"w:numId" if in_style_num_pr => {
                            current_style.num_id = get_attr(e, b"w:val").and_then(|v| v.parse().ok());
                        }
                        b"w:ilvl" if in_style_num_pr => {
                            current_style.ilvl = get_attr(e, b"w:val").and_then(|v| v.parse().ok());
                        }
                        b"w:jc" if in_ppr => {
                            current_style.alignment = get_attr(e, b"w:val");
                        }
                        b"w:b" if in_rpr => {
                            current_style.bold = Some(!is_val_false(e));
                        }
                        b"w:i" if in_rpr => {
                            current_style.italic = Some(!is_val_false(e));
                        }
                        b"w:sz" if in_rpr => {
                            if let Some(val) = get_attr(e, b"w:val") {
                                if let Ok(half_pt) = val.parse::<f32>() {
                                    current_style.font_size = Some(half_pt / 2.0);
                                }
                            }
                        }
                        b"w:rFonts" if in_rpr => {
                            current_style.font_family = extract_font_family(e);
                        }
                        b"w:color" if in_rpr => {
                            current_style.color = get_attr(e, b"w:val");
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(ref e)) => match e.name().as_ref() {
                    b"w:style" => {
                        if let Some(id) = current_style_id.take() {
                            styles.insert(id, current_style.clone());
                        }
                        current_style = Style::new();
                        in_rpr = false;
                        in_ppr = false;
                    }
                    b"w:rPr" => in_rpr = false,
                    b"w:pPr" => in_ppr = false,
                    b"w:numPr" => in_style_num_pr = false,
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        // Resolve style inheritance
        let style_ids: Vec<String> = styles.keys().cloned().collect();
        for id in &style_ids {
            let mut resolved = styles.get(id).cloned().unwrap_or_default();
            let mut visited = HashSet::from([id.clone()]);
            let mut base_id = resolved.based_on.clone();
            while let Some(ref bid) = base_id {
                if !visited.insert(bid.clone()) {
                    break; // cycle
                }
                if let Some(base) = styles.get(bid) {
                    if resolved.font_size.is_none() {
                        resolved.font_size = base.font_size;
                    }
                    if resolved.font_family.is_none() {
                        resolved.font_family = base.font_family.clone();
                    }
                    if resolved.bold.is_none() {
                        resolved.bold = base.bold;
                    }
                    if resolved.italic.is_none() {
                        resolved.italic = base.italic;
                    }
                    if resolved.color.is_none() {
                        resolved.color = base.color.clone();
                    }
                    if resolved.alignment.is_none() {
                        resolved.alignment = base.alignment.clone();
                    }
                    if resolved.heading_level.is_none() {
                        resolved.heading_level = base.heading_level;
                    }
                    if resolved.num_id.is_none() {
                        resolved.num_id = base.num_id;
                        if resolved.ilvl.is_none() {
                            resolved.ilvl = base.ilvl;
                        }
                    }
                    base_id = base.based_on.clone();
                } else {
                    break;
                }
            }
            styles.insert(id.clone(), resolved);
        }

        Ok(styles)
    }

    // --- Comments ---

    fn parse_comments(&mut self) -> Result<Vec<Comment>> {
        let content = match self.read_archive_file("word/comments.xml") {
            Ok(c) => c,
            Err(_) => return Ok(Vec::new()),
        };

        let mut comments = Vec::new();
        let mut reader = Reader::from_str(&content);
        let mut buf = Vec::new();

        let mut current: Option<Comment> = None;
        let mut text_buf = String::new();
        let mut in_text = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => match e.name().as_ref() {
                    b"w:comment" => {
                        let id = get_attr(e, b"w:id")
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                        let author = get_attr(e, b"w:author").unwrap_or_default();
                        let date = get_attr(e, b"w:date");
                        let para_id = extract_comment_para_id(e);
                        current = Some(Comment {
                            id,
                            author,
                            date,
                            text: String::new(),
                            para_id,
                            parent_id: None,
                            thread_id: None,
                        });
                        text_buf.clear();
                    }
                    b"w:t" if current.is_some() => {
                        in_text = true;
                    }
                    _ => {}
                },
                Ok(Event::Text(ref e)) if in_text => {
                    if let Ok(t) = e.unescape() {
                        text_buf.push_str(&t);
                    }
                }
                Ok(Event::End(ref e)) => match e.name().as_ref() {
                    b"w:t" => in_text = false,
                    b"w:p" if current.is_some() => {
                        if !text_buf.is_empty() {
                            if let Some(ref mut c) = current {
                                if !c.text.is_empty() {
                                    c.text.push('\n');
                                }
                                c.text.push_str(&text_buf);
                            }
                            text_buf.clear();
                        }
                    }
                    b"w:comment" => {
                        if let Some(comment) = current.take() {
                            comments.push(comment);
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

        Ok(comments)
    }

    fn parse_comment_threads(&mut self) -> Result<HashMap<String, String>> {
        let content = match self.read_archive_file("word/commentsExtended.xml") {
            Ok(c) => c,
            Err(_) => return Ok(HashMap::new()),
        };

        let mut reader = Reader::from_str(&content);
        let mut buf = Vec::new();
        let mut threads = HashMap::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    if e.name().as_ref() == b"w15:commentEx" {
                        if let Some(para_id) = extract_comment_para_id(e) {
                            if let Some(parent_para_id) = extract_comment_parent_para_id(e) {
                                threads.insert(para_id, parent_para_id);
                            } else {
                                threads.entry(para_id).or_insert_with(String::new);
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(threads)
    }

    fn attach_comment_threads(
        &self,
        comments: &mut [Comment],
        comment_threads: HashMap<String, String>,
    ) {
        let para_to_comment_id: HashMap<String, u32> = comments
            .iter()
            .filter_map(|comment| {
                comment
                    .para_id
                    .as_ref()
                    .map(|para_id| (para_id.clone(), comment.id))
            })
            .collect();

        let mut thread_cache: HashMap<String, Option<u32>> = HashMap::new();
        let mut thread_visited: HashSet<String> = HashSet::new();

        for comment in comments.iter_mut() {
            let Some(para_id) = comment.para_id.as_ref() else {
                comment.parent_id = None;
                comment.thread_id = Some(comment.id);
                continue;
            };

            let parent_para_id = comment_threads
                .get(para_id)
                .filter(|parent| !parent.is_empty())
                .cloned();

            comment.parent_id = parent_para_id
                .as_ref()
                .and_then(|parent_para_id| para_to_comment_id.get(parent_para_id).copied());

            comment.thread_id = Some(self.resolve_thread_id(
                para_id,
                &comment_threads,
                &para_to_comment_id,
                &mut thread_cache,
                &mut thread_visited,
                comment.id,
            ));
        }
    }

    fn resolve_thread_id(
        &self,
        para_id: &str,
        comment_threads: &HashMap<String, String>,
        para_to_comment_id: &HashMap<String, u32>,
        cache: &mut HashMap<String, Option<u32>>,
        visited: &mut HashSet<String>,
        fallback_comment_id: u32,
    ) -> u32 {
        if !visited.insert(para_id.to_string()) {
            return fallback_comment_id;
        }

        if let Some(cached) = cache.get(para_id) {
            visited.remove(para_id);
            return cached.unwrap_or(fallback_comment_id);
        }

        let resolved = match comment_threads.get(para_id) {
            Some(parent_para_id) if !parent_para_id.is_empty() => {
                if let Some(&parent_comment_id) = para_to_comment_id.get(parent_para_id) {
                    self.resolve_thread_id(
                        parent_para_id,
                        comment_threads,
                        para_to_comment_id,
                        cache,
                        visited,
                        parent_comment_id,
                    )
                } else {
                    fallback_comment_id
                }
            }
            _ => fallback_comment_id,
        };

        cache.insert(para_id.to_string(), Some(resolved));
        visited.remove(para_id);
        resolved
    }

    // --- Images ---

    fn parse_images(&mut self) -> Result<HashMap<String, String>> {
        let mut images = HashMap::new();

        let image_relationships: Vec<(String, String)> = self
            .relationships
            .iter()
            .filter(|rel| {
                self.referenced_image_ids.contains(&rel.id)
                    && (rel.rel_type.contains("image") || rel.target.starts_with("media/"))
            })
            .map(|rel| (rel.id.clone(), rel.target.clone()))
            .collect();

        for (rel_id, target) in image_relationships {
            let full_path = if target.starts_with("media/") {
                format!("word/{}", target)
            } else if target.starts_with("word/") {
                target
            } else {
                format!("word/{}", target)
            };

            let lower = full_path.to_lowercase();
            let is_emf = lower.ends_with(".emf");
            let is_wmf = lower.ends_with(".wmf");

            // Read the bytes in a scope so the archive borrow is released before
            // we call other &self methods (e.g. placeholder_data_uri).
            let bytes = match self.archive.by_name(&full_path) {
                Ok(mut file) => {
                    let mut buffer = Vec::with_capacity(file.size() as usize);
                    file.read_to_end(&mut buffer).ok().map(|_| buffer)
                }
                Err(_) => None,
            };

            let Some(buffer) = bytes else {
                if is_emf || is_wmf {
                    images.insert(rel_id, self.placeholder_data_uri(&full_path));
                }
                continue;
            };

            // EMF is a vector format browsers can't show; rasterize it to PNG.
            if is_emf {
                if let Some(png) = crate::emf::emf_to_png(&buffer) {
                    let b64 = base64::engine::general_purpose::STANDARD.encode(&png);
                    images.insert(rel_id, format!("data:image/png;base64,{}", b64));
                } else {
                    images.insert(rel_id, self.placeholder_data_uri(&full_path));
                }
            } else if is_wmf {
                // WMF is not yet rasterized; show a placeholder.
                images.insert(rel_id, self.placeholder_data_uri(&full_path));
            } else {
                let mime = mime_for_path(&full_path);
                let b64 = base64::engine::general_purpose::STANDARD.encode(&buffer);
                images.insert(rel_id, format!("data:{};base64,{}", mime, b64));
            }
        }

        Ok(images)
    }

    // --- Headers & Footers ---

    fn parse_headers_footers(&mut self) -> Result<(Vec<HeaderFooter>, Vec<HeaderFooter>)> {
        let mut headers = Vec::new();
        let mut footers = Vec::new();

        let header_footer_relationships: Vec<(bool, String)> = self
            .relationships
            .iter()
            .filter_map(|rel| {
                let is_header = rel.rel_type.contains("header");
                let is_footer = rel.rel_type.contains("footer");
                if !is_header && !is_footer {
                    return None;
                }

                let full_path = if rel.target.starts_with("word/") {
                    rel.target.clone()
                } else {
                    format!("word/{}", rel.target)
                };
                Some((is_header, full_path))
            })
            .collect();

        for (is_header, path) in header_footer_relationships {
            if let Ok(content) = self.read_archive_file(&path) {
                let blocks = self.parse_body_xml(&content);
                let section = extract_section_number(&path);
                let hf = HeaderFooter {
                    content: blocks,
                    section,
                };
                if is_header {
                    headers.push(hf);
                } else {
                    footers.push(hf);
                }
            }
        }

        Ok((headers, footers))
    }

    // --- Footnotes ---

    fn parse_footnotes(&mut self) -> Result<Vec<Footnote>> {
        let content = match self.read_archive_file("word/footnotes.xml") {
            Ok(c) => c,
            Err(_) => return Ok(Vec::new()),
        };

        let mut footnotes = Vec::new();
        let mut reader = Reader::from_str(&content);
        let mut buf = Vec::new();

        let mut current_id: Option<u32> = None;
        let mut current_blocks: Vec<BlockElement> = Vec::new();
        let mut para_runs: Vec<Run> = Vec::new();
        let mut run_text = String::new();
        let mut in_run = false;
        let mut in_text = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => match e.name().as_ref() {
                    b"w:footnote" => {
                        current_id = get_attr(e, b"w:id").and_then(|s| s.parse().ok());
                        current_blocks.clear();
                    }
                    b"w:r" if current_id.is_some() => {
                        in_run = true;
                        run_text.clear();
                    }
                    b"w:t" if in_run => {
                        in_text = true;
                    }
                    _ => {}
                },
                Ok(Event::Text(ref e)) if in_text => {
                    if let Ok(t) = e.unescape() {
                        run_text.push_str(&t);
                    }
                }
                Ok(Event::End(ref e)) => match e.name().as_ref() {
                    b"w:t" => in_text = false,
                    b"w:r" => {
                        in_run = false;
                        if !run_text.is_empty() {
                            para_runs.push(Run::new(run_text.clone()));
                        }
                        run_text.clear();
                    }
                    b"w:p" if current_id.is_some() => {
                        if !para_runs.is_empty() {
                            current_blocks.push(BlockElement::Paragraph {
                                runs: std::mem::take(&mut para_runs),
                                style: None,
                                alignment: None,
                                list_level: None,
                                list_format: None,
                            });
                        }
                    }
                    b"w:footnote" => {
                        if let Some(id) = current_id.take() {
                            // Skip special footnotes (separator, continuation)
                            if id > 0 {
                                footnotes.push(Footnote {
                                    id,
                                    content: std::mem::take(&mut current_blocks),
                                });
                            }
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

        Ok(footnotes)
    }

    // --- Numbering ---

    fn parse_numbering(&mut self) -> HashMap<(u32, u8), NumberingLevelInfo> {
        let content = match self.read_archive_file("word/numbering.xml") {
            Ok(c) => c,
            Err(_) => return HashMap::new(),
        };

        let mut reader = Reader::from_str(&content);
        let mut buf = Vec::new();

        let mut abstract_nums: HashMap<u32, Vec<(u8, NumberingLevelInfo)>> = HashMap::new();
        let mut num_to_abstract: HashMap<u32, u32> = HashMap::new();

        let mut cur_abstract_id: Option<u32> = None;
        let mut cur_lvl_ilvl: Option<u8> = None;
        let mut lvl_num_fmt = String::new();
        let mut lvl_text = String::new();
        let mut lvl_start: u32 = 1;
        let mut in_num = false;
        let mut cur_num_id: Option<u32> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => match e.name().as_ref() {
                    b"w:abstractNum" => {
                        cur_abstract_id =
                            get_attr(e, b"w:abstractNumId").and_then(|v| v.parse().ok());
                    }
                    b"w:lvl" if cur_abstract_id.is_some() => {
                        cur_lvl_ilvl = get_attr(e, b"w:ilvl").and_then(|v| v.parse().ok());
                        lvl_num_fmt.clear();
                        lvl_text.clear();
                        lvl_start = 1;
                    }
                    b"w:start" if cur_lvl_ilvl.is_some() => {
                        if let Some(v) = get_attr(e, b"w:val") {
                            lvl_start = v.parse().unwrap_or(1);
                        }
                    }
                    b"w:numFmt" if cur_lvl_ilvl.is_some() => {
                        if let Some(v) = get_attr(e, b"w:val") {
                            lvl_num_fmt = v;
                        }
                    }
                    b"w:lvlText" if cur_lvl_ilvl.is_some() => {
                        if let Some(v) = get_attr(e, b"w:val") {
                            lvl_text = v;
                        }
                    }
                    b"w:num" => {
                        in_num = true;
                        cur_num_id = get_attr(e, b"w:numId").and_then(|v| v.parse().ok());
                    }
                    b"w:abstractNumId" if in_num => {
                        if let (Some(num_id), Some(abs_id)) = (
                            cur_num_id,
                            get_attr(e, b"w:val").and_then(|v| v.parse().ok()),
                        ) {
                            num_to_abstract.insert(num_id, abs_id);
                        }
                    }
                    _ => {}
                },
                Ok(Event::End(ref e)) => match e.name().as_ref() {
                    b"w:lvl" => {
                        if let (Some(abs_id), Some(ilvl)) = (cur_abstract_id, cur_lvl_ilvl.take()) {
                            abstract_nums.entry(abs_id).or_default().push((
                                ilvl,
                                NumberingLevelInfo {
                                    num_fmt: lvl_num_fmt.clone(),
                                    lvl_text: lvl_text.clone(),
                                    start: lvl_start,
                                },
                            ));
                        }
                    }
                    b"w:abstractNum" => {
                        cur_abstract_id = None;
                    }
                    b"w:num" => {
                        in_num = false;
                        cur_num_id = None;
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        let mut result = HashMap::new();
        for (&num_id, &abs_id) in &num_to_abstract {
            if let Some(levels) = abstract_nums.get(&abs_id) {
                for (ilvl, info) in levels {
                    result.insert((num_id, *ilvl), info.clone());
                }
            }
        }
        result
    }

    // --- Relationships ---

    fn read_relationships(&mut self) -> Vec<Relationship> {
        let content = match self.read_archive_file("word/_rels/document.xml.rels") {
            Ok(c) => c,
            Err(_) => return Vec::new(),
        };

        let mut rels = Vec::new();
        let mut reader = Reader::from_str(&content);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    if e.name().as_ref() == b"Relationship" {
                        let id = get_attr(e, b"Id").unwrap_or_default();
                        let rel_type = get_attr(e, b"Type").unwrap_or_default();
                        let target = get_attr(e, b"Target").unwrap_or_default();
                        if !id.is_empty() {
                            rels.push(Relationship {
                                id,
                                rel_type,
                                target,
                            });
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        rels
    }

    // --- Helpers ---

    fn read_archive_file(&mut self, path: &str) -> Result<String> {
        match self.archive.by_name(path) {
            Ok(mut file) => {
                let mut content = String::with_capacity(file.size() as usize);
                file.read_to_string(&mut content)?;
                Ok(content)
            }
            Err(e) => Err(ParseError::Zip(e)),
        }
    }

    /// Parse a body-like XML fragment (for headers/footers) into BlockElements
    fn parse_body_xml(&mut self, content: &str) -> Vec<BlockElement> {
        let mut blocks = Vec::new();
        let mut reader = Reader::from_str(content);
        let mut buf = Vec::new();

        let mut runs: Vec<Run> = Vec::new();
        let mut run_ctx = RunCtx::new();
        let mut in_run = false;
        let mut in_text = false;
        let mut para_style: Option<String> = None;
        let mut para_alignment: Option<String> = None;
        let mut in_para = false;
        let mut hyperlink_url: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => match e.name().as_ref() {
                    b"w:p" => {
                        in_para = true;
                        para_style = None;
                        para_alignment = None;
                    }
                    b"w:hyperlink" if in_para => {
                        hyperlink_url = get_attr(e, b"r:id").and_then(|id| {
                            self.relationships
                                .iter()
                                .find(|r| r.id == id)
                                .map(|r| r.target.clone())
                        });
                    }
                    b"w:pStyle" if in_para => {
                        para_style = get_attr(e, b"w:val");
                    }
                    b"w:jc" if in_para => {
                        para_alignment = get_attr(e, b"w:val");
                    }
                    b"w:r" if in_para => {
                        in_run = true;
                        run_ctx = RunCtx::new();
                        run_ctx.link_url = hyperlink_url.clone();
                    }
                    b"w:t" if in_run => in_text = true,
                    b"w:b" if in_run => run_ctx.bold = !is_val_false(e),
                    b"w:i" if in_run => run_ctx.italic = !is_val_false(e),
                    b"w:u" if in_run => {
                        let val = get_attr(e, b"w:val");
                        run_ctx.underline = val.as_deref() != Some("none");
                    }
                    b"w:strike" if in_run => run_ctx.strikethrough = !is_val_false(e),
                    b"w:sz" if in_run => {
                        if let Some(val) = get_attr(e, b"w:val") {
                            if let Ok(half_pt) = val.parse::<f32>() {
                                run_ctx.font_size = Some(half_pt / 2.0);
                            }
                        }
                    }
                    b"w:rFonts" if in_run => {
                        run_ctx.font_family = extract_font_family(e);
                    }
                    b"w:color" if in_run => {
                        run_ctx.color = get_attr(e, b"w:val");
                    }
                    b"w:highlight" if in_run => {
                        run_ctx.highlight = get_attr(e, b"w:val");
                    }
                    b"a:blip" if in_run => {
                        if let Some(image_id) = get_attr(e, b"r:embed") {
                            self.referenced_image_ids.insert(image_id.clone());
                            run_ctx.image_id = Some(image_id);
                        }
                    }
                    b"v:imagedata" if in_run => {
                        if let Some(image_id) = get_attr(e, b"r:id") {
                            self.referenced_image_ids.insert(image_id.clone());
                            run_ctx.image_id = Some(image_id);
                        }
                    }
                    _ => {}
                },
                Ok(Event::Text(ref e)) if in_text => {
                    if let Ok(t) = e.unescape() {
                        run_ctx.text.push_str(&t);
                    }
                }
                Ok(Event::End(ref e)) => match e.name().as_ref() {
                    b"w:t" => in_text = false,
                    b"w:hyperlink" => {
                        hyperlink_url = None;
                    }
                    b"w:r" => {
                        in_run = false;
                        let finished_run = std::mem::replace(&mut run_ctx, RunCtx::new());
                        if !finished_run.text.is_empty() || finished_run.image_id.is_some() {
                            runs.push(finished_run.into_run(&[]));
                        }
                    }
                    b"w:p" => {
                        in_para = false;
                        blocks.push(BlockElement::Paragraph {
                            runs: std::mem::take(&mut runs),
                            style: para_style.take(),
                            alignment: para_alignment.take(),
                            list_level: None,
                            list_format: None,
                        });
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        blocks
    }

    fn placeholder_data_uri(&self, path: &str) -> String {
        let filename = path.split('/').last().unwrap_or(path);
        let svg = format!(
            "<svg xmlns='http://www.w3.org/2000/svg' width='200' height='80'>\
             <rect width='200' height='80' fill='#f0f0f0' stroke='#ccc'/>\
             <text x='100' y='35' text-anchor='middle' font-size='11' fill='#666'>Unsupported format</text>\
             <text x='100' y='55' text-anchor='middle' font-size='9' fill='#999'>{}</text>\
             </svg>",
            filename
        );
        let b64 = base64::engine::general_purpose::STANDARD.encode(svg.as_bytes());
        format!("data:image/svg+xml;base64,{}", b64)
    }

    pub fn get_image_mime_type(&self, filename: &str) -> &'static str {
        mime_for_path(filename)
    }

    pub fn is_supported_image(&self, filename: &str) -> bool {
        let lower = filename.to_lowercase();
        lower.ends_with(".png")
            || lower.ends_with(".jpg")
            || lower.ends_with(".jpeg")
            || lower.ends_with(".gif")
            || lower.ends_with(".bmp")
    }
}

// --- Internal context types ---

struct ParagraphCtx {
    runs: Vec<Run>,
    style: Option<String>,
    alignment: Option<String>,
    has_page_break_after: bool,
    num_id: Option<u32>,
    ilvl: Option<u8>,
}

impl ParagraphCtx {
    fn new() -> Self {
        Self {
            runs: Vec::new(),
            style: None,
            alignment: None,
            has_page_break_after: false,
            num_id: None,
            ilvl: None,
        }
    }

    fn into_block_element(
        self,
        list_level: Option<u8>,
        list_format: Option<String>,
    ) -> BlockElement {
        BlockElement::Paragraph {
            runs: self.runs,
            style: self.style,
            alignment: self.alignment,
            list_level,
            list_format,
        }
    }
}

struct RunCtx {
    text: String,
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    font_size: Option<f32>,
    font_family: Option<String>,
    color: Option<String>,
    highlight: Option<String>,
    footnote_ref: Option<u32>,
    image_id: Option<String>,
    link_url: Option<String>,
}

impl RunCtx {
    fn new() -> Self {
        Self {
            text: String::new(),
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            font_size: None,
            font_family: None,
            color: None,
            highlight: None,
            footnote_ref: None,
            image_id: None,
            link_url: None,
        }
    }

    fn into_run(self, comment_ranges: &[u32]) -> Run {
        Run {
            text: self.text,
            bold: self.bold,
            italic: self.italic,
            underline: self.underline,
            strikethrough: self.strikethrough,
            font_size: self.font_size,
            font_family: self.font_family,
            color: self.color,
            highlight: self.highlight,
            comment_ref: comment_ranges.last().copied(),
            footnote_ref: self.footnote_ref,
            image_id: self.image_id,
            link_url: self.link_url,
        }
    }
}

struct TableCtx {
    rows: Vec<TableRow>,
    current_row: Option<Vec<TableCell>>,
    cell_content: Vec<BlockElement>,
}

impl TableCtx {
    fn new() -> Self {
        Self {
            rows: Vec::new(),
            current_row: None,
            cell_content: Vec::new(),
        }
    }

    fn push_cell_content(&mut self, element: BlockElement) {
        self.cell_content.push(element);
    }

    fn finish_cell(&mut self) {
        if let Some(ref mut row) = self.current_row {
            row.push(TableCell::new(std::mem::take(&mut self.cell_content)));
        }
    }

    fn finish_row(&mut self) {
        if let Some(cells) = self.current_row.take() {
            self.rows.push(TableRow { cells });
        }
    }
}

// --- Free functions ---

fn get_attr(e: &BytesStart, key: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == key {
            return std::str::from_utf8(&attr.value).ok().map(String::from);
        }
    }
    None
}

fn extract_font_family(e: &BytesStart) -> Option<String> {
    get_attr(e, b"w:ascii")
        .or_else(|| get_attr(e, b"w:hAnsi"))
        .or_else(|| get_attr(e, b"w:eastAsia"))
        .or_else(|| get_attr(e, b"w:cs"))
}

fn extract_comment_para_id(e: &BytesStart) -> Option<String> {
    get_attr(e, b"w15:paraId").or_else(|| get_attr(e, b"w:paraId"))
}

fn extract_comment_parent_para_id(e: &BytesStart) -> Option<String> {
    get_attr(e, b"w15:paraIdParent").or_else(|| get_attr(e, b"w:paraIdParent"))
}

fn is_val_false(e: &BytesStart) -> bool {
    match get_attr(e, b"w:val") {
        Some(v) => v == "0" || v == "false",
        None => false, // absence means true for boolean properties
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
    } else {
        "application/octet-stream"
    }
}

fn extract_section_number(path: &str) -> u32 {
    // Extract number from "header1.xml", "footer2.xml", etc.
    let filename = path.split('/').last().unwrap_or(path);
    filename
        .chars()
        .filter(|c| c.is_ascii_digit())
        .collect::<String>()
        .parse()
        .unwrap_or(1)
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
    fn test_mime_types() {
        assert_eq!(mime_for_path("test.png"), "image/png");
        assert_eq!(mime_for_path("test.jpg"), "image/jpeg");
        assert_eq!(mime_for_path("test.jpeg"), "image/jpeg");
        assert_eq!(mime_for_path("test.gif"), "image/gif");
        assert_eq!(mime_for_path("test.bmp"), "image/bmp");
    }

    #[test]
    fn test_get_attr() {
        let xml = r#"<w:pStyle w:val="Heading1"/>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();
        if let Ok(Event::Empty(ref e)) = reader.read_event_into(&mut buf) {
            assert_eq!(get_attr(e, b"w:val"), Some("Heading1".to_string()));
            assert_eq!(get_attr(e, b"w:missing"), None);
        }
    }

    #[test]
    fn test_extract_section_number() {
        assert_eq!(extract_section_number("header1.xml"), 1);
        assert_eq!(extract_section_number("word/footer2.xml"), 2);
        assert_eq!(extract_section_number("header.xml"), 1); // fallback
    }

    #[test]
    fn test_extract_font_family_prefers_actual_font_slots() {
        let xml = r#"<w:rFonts w:ascii="Aptos" w:eastAsia="Garamond" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();

        if let Ok(Event::Empty(ref e)) = reader.read_event_into(&mut buf) {
            assert_eq!(extract_font_family(e), Some("Aptos".to_string()));
        } else {
            panic!("expected font element");
        }
    }

    #[test]
    fn test_parse_preserves_trailing_spaces_and_font_family() {
        let docx_path = write_test_docx(&[(
            "word/document.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:rFonts w:eastAsia="Garamond"/>
        </w:rPr>
        <w:t xml:space="preserve">traceable </w:t>
      </w:r>
      <w:r>
        <w:t>on</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#,
        )]);

        let mut parser = DocxParser::from_path(docx_path.to_str().unwrap()).expect("parser init");
        let document = parser.parse().expect("parse document");

        let paragraph_runs = match &document.body[0] {
            BlockElement::Paragraph { runs, .. } => runs,
            other => panic!("expected paragraph, got {:?}", other),
        };

        let text = paragraph_runs
            .iter()
            .map(|run| run.text.as_str())
            .collect::<String>();
        assert_eq!(text, "traceable on");
        assert_eq!(paragraph_runs[0].font_family, Some("Garamond".to_string()));

        let _ = fs::remove_file(docx_path);
    }

    #[test]
    fn test_parse_comment_threads_from_comments_extended() {
        let docx_path = write_test_docx(&[
            (
                "word/document.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Threaded comments</w:t></w:r></w:p>
  </w:body>
</w:document>"#,
            ),
            (
                "word/comments.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w:comment w:id="0" w:author="Alice" w15:paraId="11111111">
    <w:p><w:r><w:t>Root comment</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="1" w:author="Bob" w15:paraId="22222222">
    <w:p><w:r><w:t>First reply</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="2" w:author="Cara" w15:paraId="33333333">
    <w:p><w:r><w:t>Nested reply</w:t></w:r></w:p>
  </w:comment>
</w:comments>"#,
            ),
            (
                "word/commentsExtended.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w15:commentEx w15:paraId="11111111"/>
  <w15:commentEx w15:paraId="22222222" w15:paraIdParent="11111111"/>
  <w15:commentEx w15:paraId="33333333" w15:paraIdParent="22222222"/>
</w15:commentsEx>"#,
            ),
        ]);

        let mut parser = DocxParser::from_path(docx_path.to_str().unwrap()).expect("parser init");
        let document = parser.parse().expect("parse document");

        assert_eq!(document.comments.len(), 3);
        assert_eq!(document.comments[0].para_id.as_deref(), Some("11111111"));
        assert_eq!(document.comments[0].parent_id, None);
        assert_eq!(document.comments[0].thread_id, Some(0));

        assert_eq!(document.comments[1].para_id.as_deref(), Some("22222222"));
        assert_eq!(document.comments[1].parent_id, Some(0));
        assert_eq!(document.comments[1].thread_id, Some(0));

        assert_eq!(document.comments[2].para_id.as_deref(), Some("33333333"));
        assert_eq!(document.comments[2].parent_id, Some(1));
        assert_eq!(document.comments[2].thread_id, Some(0));

        let _ = fs::remove_file(docx_path);
    }

    fn write_test_docx(files: &[(&str, &str)]) -> PathBuf {
        let path = std::env::temp_dir().join(format!(
            "papyr-parser-test-{}-{}.docx",
            std::process::id(),
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .expect("clock drift")
                .as_nanos()
        ));

        let file = File::create(&path).expect("create docx file");
        let mut zip = zip::ZipWriter::new(file);
        let options: FileOptions<'_, ()> = FileOptions::default();

        for &(file_name, content) in files {
            zip.start_file(file_name, options).expect("start zip file");
            zip.write_all(content.as_bytes()).expect("write zip file");
        }
        zip.finish().expect("finish docx");

        path
    }
}
