# Papyr - Specification

> A lightweight, Rust-powered DOCX and XLSX viewer. Messenger of the gods -- delivers the document to you, fast and light.

## 1. Overview

Papyr is a read-only DOCX and XLSX viewer built with Rust and Tauri 2.0. It replaces the need to open Microsoft Word or Excel when you just want to read a document or quickly inspect a workbook. The app is fast, minimal, and cross-platform (Windows, macOS, Linux).

### Goals

- Open and render .docx files with good fidelity
- Open .xlsx files quickly as a spreadsheet grid preview
- Read-only -- no editing
- Lightweight binary (~3-5MB) with low memory footprint
- Fast startup and document load times
- Display comments, headers/footers, and footnotes

### Non-Goals

- Editing or saving DOCX or XLSX files
- 100% OOXML spec compliance (the spec is 6,000+ pages)
- Rendering macros, form fields, or embedded OLE objects
- Automatic line-overflow pagination (explicit page breaks only)
- Evaluating XLSX formulas or rendering charts, pivot tables, macros, embedded objects, or full Excel-compatible layout

## 2. Architecture

```
+-------------------+       IPC (JSON)       +--------------------+
|   Rust Backend    | <--------------------> |   Web Frontend     |
|   (src-tauri/)    |                        |   (src/)           |
|                   |                        |                    |
| - DOCX parser     |  invoke('open_file')   | - Document renderer|
| - XLSX parser     | ------------------>    | - Spreadsheet grid |
| - ZIP extraction   | ------------------>   | - Comment panel    |
| - XML parsing      |  <-- DocumentModel    | - Page layout CSS  |
| - Image extraction  |      as JSON          | - File open UI     |
| - Comment extraction|                       | - Dark/light theme |
+-------------------+                        +--------------------+
```

**Tauri 2.0** provides the app shell. The Rust backend parses DOCX and XLSX files and sends a tagged structured model as JSON to the web frontend via Tauri's IPC invoke mechanism. The frontend renders the model as styled HTML in the system webview.

## 3. Tech Stack

| Component | Technology | Rationale |
|-----------|-----------|-----------|
| App shell | Tauri 2.0 | Lightweight native app with system webview |
| DOCX parsing | `zip` + `quick-xml` | Direct control over OOXML parsing; `docx-rust` is generation-focused |
| XLSX parsing | `zip` + `quick-xml` | Direct control over a fast workbook preview path |
| Image encoding | `base64` | Embed images as data URIs in HTML |
| Serialization | `serde` + `serde_json` | Rust-to-JS document model transfer |
| Frontend | Vanilla HTML/CSS/JS | No framework, no build step, minimal footprint |
| Page layout | CSS | Print-like paginated view (white pages on gray background) |

### Why manual ZIP+XML over `docx-rust`?

`docx-rust` is primarily designed for generating DOCX files. Its parsing support is limited for edge cases. For a viewer, we need reliable read access to multiple OOXML parts:

- `word/document.xml` -- body content
- `word/styles.xml` -- formatting definitions
- `word/comments.xml` -- comment threads
- `word/_rels/document.xml.rels` -- relationships (image references)
- `word/header*.xml` / `word/footer*.xml` -- headers and footers
- `word/footnotes.xml` -- footnote definitions
- `word/media/*` -- embedded images

Direct ZIP+XML parsing avoids fighting a library's abstraction for features it does not fully support.

## 4. Document Model

The Rust backend parses DOCX into the following model, serialized as JSON for the frontend.

### Core Types

```rust
struct Document {
    body: Vec<BlockElement>,
    comments: Vec<Comment>,
    headers: Vec<HeaderFooter>,
    footers: Vec<HeaderFooter>,
    footnotes: Vec<Footnote>,
    styles: HashMap<String, Style>,
    images: HashMap<String, String>,  // rId -> base64 data URI
}

enum BlockElement {
    Paragraph {
        runs: Vec<Run>,
        style: Option<String>,
        alignment: Option<String>,
    },
    Table {
        rows: Vec<TableRow>,
    },
    PageBreak,
}

struct Run {
    text: String,
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    font_size: Option<f32>,       // in pt
    font_family: Option<String>,
    color: Option<String>,        // hex
    highlight: Option<String>,
    comment_ref: Option<u32>,     // links to Comment.id
    footnote_ref: Option<u32>,    // links to Footnote.id
    image_id: Option<String>,     // links to images map
}

struct TableRow {
    cells: Vec<TableCell>,
}

struct TableCell {
    content: Vec<BlockElement>,
    col_span: u32,
    row_span: u32,
    shading: Option<String>,      // background color hex
}

struct Comment {
    id: u32,
    author: String,
    date: Option<String>,
    text: String,
}

struct HeaderFooter {
    content: Vec<BlockElement>,
    section: u32,
}

struct Footnote {
    id: u32,
    content: Vec<BlockElement>,
}

struct Style {
    based_on: Option<String>,
    font_size: Option<f32>,
    font_family: Option<String>,
    bold: Option<bool>,
    italic: Option<bool>,
    color: Option<String>,
    alignment: Option<String>,
    heading_level: Option<u8>,    // 1-6 for heading styles
}
```

### XLSX Preview Types

The Rust backend parses XLSX into a sparse workbook preview model, serialized as tagged JSON for the frontend.

```rust
struct XlsxWorkbook {
    sheets: Vec<XlsxSheetSummary>,
    active_sheet_index: usize,
    active_sheet: Option<XlsxSheet>,
    limits: XlsxPreviewLimits,
}

struct XlsxSheet {
    index: usize,
    name: String,
    rows: Vec<XlsxRow>,
    row_count: u32,
    column_count: u32,
    max_row: u32,
    max_column: u32,
    truncated: bool,
    truncated_reasons: Vec<String>,
}

struct XlsxCell {
    reference: String,
    row: u32,
    column: u32,
    value: String,
    raw_value: Option<String>,
    value_type: XlsxCellValueType,
    formula: Option<String>,
    style_index: Option<u32>,
    number_format: Option<String>,
}
```

The preview model includes visible cell values, shared strings, inline strings, numbers, booleans, dates, formulas with cached values, and basic number/date/text formats.

## 5. DOCX Parsing Details

### 5.1 ZIP Extraction

A .docx file is a ZIP archive. The `zip` crate opens it and provides access to individual files by path.

### 5.2 Document Body (`word/document.xml`)

The main content. Key OOXML elements:

| OOXML Element | Maps To |
|--------------|---------|
| `w:p` | `BlockElement::Paragraph` |
| `w:r` | `Run` |
| `w:t` | `Run.text` |
| `w:rPr` | Run formatting (bold, italic, etc.) |
| `w:pPr` | Paragraph properties (style, alignment) |
| `w:tbl` | `BlockElement::Table` |
| `w:tr` | `TableRow` |
| `w:tc` | `TableCell` |
| `w:br` with `w:type="page"` | `BlockElement::PageBreak` |
| `w:commentRangeStart` | Start of commented text |
| `w:commentRangeEnd` | End of commented text |
| `w:footnoteReference` | `Run.footnote_ref` |
| `w:drawing` / `w:pict` | Image reference (resolved via relationships) |

### 5.3 Styles (`word/styles.xml`)

Defines named styles with inheritance (`w:basedOn`). The parser resolves the full inheritance chain and flattens each style into concrete CSS-friendly properties.

### 5.4 Comments (`word/comments.xml`)

Each `w:comment` element contains an `id`, `author`, `date`, and body paragraphs. Comments are linked to document text via `w:commentRangeStart`/`w:commentRangeEnd` markers in `document.xml` that reference the comment ID.

### 5.5 Headers and Footers

`word/header{N}.xml` and `word/footer{N}.xml` contain header/footer content. They are linked to document sections via `w:sectPr` in `document.xml`, which references them by relationship ID.

### 5.6 Footnotes (`word/footnotes.xml`)

Each `w:footnote` element contains an `id` and body paragraphs. Footnote references in the document body use `w:footnoteReference` elements with the footnote ID.

### 5.7 Images

Images are stored in `word/media/`. The relationship file (`word/_rels/document.xml.rels`) maps relationship IDs to image file paths. `w:drawing` elements in the document reference images by relationship ID. The parser reads the image bytes and encodes them as base64 data URIs.

## 6. Frontend Rendering

### 6.1 Document Layout

The frontend renders the document in a paginated, print-like layout:

- Gray background (the "desk")
- White page panels with A4-ish proportions (max-width ~816px, padding for margins)
- Page breaks create visual separations between pages
- Scrollable vertical layout

### 6.2 Element Rendering

| Model Element | HTML Output |
|--------------|-------------|
| Paragraph (heading) | `<h1>` through `<h6>` |
| Paragraph (normal) | `<p>` with inline styles |
| Run | `<span>` with computed styles |
| Table | `<table>` with `colspan`/`rowspan` |
| Image | `<img src="data:...">` |
| Page break | CSS `break-before: page` or visual separator `<div>` |

### 6.3 Comments Panel

- Collapsible right sidebar
- Each comment shows author, date, and text
- Commented text in the document is highlighted (yellow/orange background)
- Clicking highlighted text scrolls to the corresponding comment in the sidebar
- Clicking a comment scrolls to the highlighted text in the document

### 6.4 Headers, Footers, and Footnotes

- Headers render at the top of each page section
- Footers render at the bottom of each page section
- Footnote references are superscript numbers in the text
- Footnote content renders at the bottom of the page (separated by a horizontal rule)

### 6.5 Theming

- Light theme (default): white pages, dark text
- Dark theme: dark pages, light text, inverted background
- Toggle via button in toolbar or keyboard shortcut

## 7. App Features

### 7.1 File Opening

- Native file dialog (Tauri dialog plugin), filtered to `.docx` and `.xlsx`
- Drag-and-drop a `.docx` or `.xlsx` file onto the window
- Double-click `.docx` or `.xlsx` files (via OS file association, future enhancement)

### 7.2 Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Ctrl+O | Open file |
| Ctrl+F | Find in document |
| Ctrl++ / Ctrl+- | Zoom document in or out |
| Ctrl+Mouse Wheel | Zoom document in or out |
| Ctrl+Q | Quit |
| Ctrl+D | Toggle dark/light theme |
| Ctrl+] | Toggle comments panel |

### 7.3 Recent Files

Store a list of recently opened files in Tauri app data. Show in a dropdown or on the welcome screen.

### 7.4 Error Handling

- Malformed or unsupported DOCX/XLSX: show a user-friendly error message, not a crash
- Unsupported elements (macros, OLE objects): silently skip, render what we can
- Empty document: show a "This document is empty" message
- Empty worksheet: show a "This sheet is empty" message

## 8. Known Limitations

| Limitation | Detail |
|-----------|--------|
| EMF/WMF images | Vector formats common in older DOCX files. Will show a placeholder. Conversion can be added later. |
| Complex nested tables | Irregular cell merges and deeply nested tables may not render perfectly. Graceful degradation. |
| Automatic pagination | No line-overflow pagination engine. Only explicit `w:br type="page"` breaks are honored. |
| Form fields | Not supported. Skipped during rendering. |
| Macros | Not supported. Ignored. |
| Track changes | Not rendered. The document shows its current accepted state. |
| Large documents | 100+ page documents with many images may be slow. Lazy loading is a future optimization. |
| XLSX formula evaluation | Not supported. Cached formula values are displayed when present. |
| XLSX charts and pivots | Not supported. Workbook preview focuses on visible sheet cells. |
| XLSX layout fidelity | Merged cells, frozen panes, filters, drawings, embedded objects, and full Excel-compatible layout are not part of the fast preview path. |

## 9. Implementation Phases

### Phase 1: Project Setup
- Scaffold Tauri 2.0 project with Rust backend + vanilla JS frontend
- Add Rust dependencies (`zip`, `quick-xml`, `serde`, `serde_json`, `base64`)
- Configure Tauri file dialog plugin

### Phase 2: DOCX Parsing (parallelizable)
- Core parser: ZIP extraction, `document.xml` parsing, paragraphs, runs, basic formatting
- Table parser: `w:tbl` elements, rows, cells, merges
- Image extractor: `word/media/` extraction, base64 encoding
- Comment parser: `word/comments.xml`, range markers
- Header/footer/footnote parser
- Style parser and inheritance resolver

### Phase 2b: XLSX Parsing
- Workbook parser: ZIP validation, workbook sheets, relationships
- Shared string parser: stream only strings referenced by preview cells
- Style parser: basic number, date, percentage, and text formats
- Worksheet parser: sparse rows and cells with row, column, and cell limits
- Sheet loading command for loading non-active sheets on demand

### Phase 3: Frontend Rendering
- HTML document renderer (paragraphs, runs, tables, images, page breaks)
- Comment display panel with cross-linking
- Header/footer/footnote rendering
- Spreadsheet renderer (sheet tabs, sticky row/column headers, sparse grid preview, truncation notices)

### Phase 4: Polish
- Dark/light theme
- Drag-and-drop, recent files
- Keyboard shortcuts
- Error handling and edge cases
- App icon and branding
